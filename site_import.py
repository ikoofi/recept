"""Импорт рецепта по ссылке. Пока поддержан только RussianFood.

Разметки schema.org на сайте нет, поэтому читаем вёрстку: при её изменении разбор придётся править.
Результат всегда идёт в форму на проверку, сам по себе рецепт в книгу не попадает.
"""
import re
import urllib.error
import urllib.request
from urllib.parse import urljoin, urlsplit

from lxml import html

SUPPORTED_HOSTS = {"www.russianfood.com", "russianfood.com"}
USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
TIMEOUT_SECONDS = 20
MAX_PAGE_BYTES = 5 * 1024 * 1024
MAX_IMAGE_BYTES = 10 * 1024 * 1024


class SiteImportError(Exception):
    """Сообщение этой ошибки показывается пользователю."""


def _download(url, limit):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:
        # Сайт мог перенаправить на другой адрес: проверяем и его.
        if urlsplit(response.geturl()).hostname not in SUPPORTED_HOSTS:
            raise SiteImportError("Страница перенаправила на другой сайт.")
        data = response.read(limit + 1)
        charset = response.headers.get_content_charset()
    if len(data) > limit:
        raise SiteImportError("Страница или фото слишком большие.")
    return data, charset


def fetch_page(url):
    """Скачивает страницу рецепта и возвращает её текст."""
    parts = urlsplit(url.strip())
    if parts.scheme not in ("http", "https") or parts.hostname not in SUPPORTED_HOSTS:
        raise SiteImportError("Пока поддерживаются только ссылки на рецепты с russianfood.com.")
    try:
        data, charset = _download(url.strip(), MAX_PAGE_BYTES)
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise SiteImportError(f"Не удалось открыть страницу: {exc}") from exc
    return data.decode(charset or "windows-1251", "replace")


def fetch_image(url):
    """Содержимое фото, либо None: без него рецепт всё равно полезен."""
    try:
        return _download(url, MAX_IMAGE_BYTES)[0]
    except (SiteImportError, urllib.error.URLError, TimeoutError, OSError):
        return None


def _text(element):
    return " ".join(element.text_content().split())


def _steps(doc):
    steps = [
        _text(p)
        for step in doc.xpath('//div[contains(@class,"step_images_n")]//div[@class="step_n"]')
        for p in step.xpath(".//p")
    ]
    if steps:
        return [s for s in steps if s]
    # Старые рецепты: весь текст одним блоком, абзацы разделены тегами <p> или <br>.
    blocks = doc.xpath('//div[@id="how"]')
    if not blocks:
        return []
    block = blocks[0]
    for bad in block.xpath(".//script | .//style | .//noindex"):
        bad.drop_tree()
    for br in block.xpath(".//br"):
        br.tail = "\n" + (br.tail or "")
    paragraphs = [" ".join(line.split()) for line in block.text_content().split("\n")]
    return [p for p in paragraphs if p]


def _minutes(text):
    match = re.search(
        r"время приготовления\s*(?:(\d+)\s*ч\w*\.?\s*)?(?:(\d+)\s*мин)?", text, re.IGNORECASE
    )
    if not match or not (match.group(1) or match.group(2)):
        return None
    return int(match.group(1) or 0) * 60 + int(match.group(2) or 0)


def parse_page(page, base_url):
    """Достаёт из страницы поля рецепта. Строки ингредиентов остаются как на сайте."""
    doc = html.fromstring(page)
    titles = doc.xpath("//h1")
    lines = [
        _text(row)
        for row in doc.xpath('//table[contains(@class,"ingr")]//tr[starts-with(@class,"ingr_tr")]')
    ]
    if not titles or not lines:
        raise SiteImportError("На странице не нашёлся рецепт. Проверьте ссылку.")

    portion = doc.xpath('//table[contains(@class,"ingr")]//span[@class="portion"]')
    servings = re.search(r"\d+", _text(portion[0])) if portion else None
    description = " ".join(doc.xpath('//meta[@name="description"]/@content'))
    image = doc.xpath('//meta[@property="og:image"]/@content')
    return {
        "title": _text(titles[0]),
        "servings": int(servings.group()) if servings else None,
        "prep_minutes": _minutes(description) or _minutes(page),
        "lines": lines,
        "steps": _steps(doc),
        "image_url": urljoin(base_url, image[0]) if image else None,
    }


_NUMBER = r"\d+(?:[.,]\d+)?(?:/\d+)?|[½¼¾⅓⅔]"
_AMOUNT = re.compile(rf"^(?:({_NUMBER})\s*[-–—]\s*)?({_NUMBER})\s*(.*)$")
# Старые рецепты пишут без тире и даже без пробелов: «вода ледяная1ст. ложка», «картофель 700-900г».
_NO_SEPARATOR = re.compile(rf"^(.+?)\s*((?:(?:{_NUMBER})\s*[-–—]\s*)?(?:{_NUMBER})\s*\D*)$")
_FRACTIONS = {"½": "0.5", "¼": "0.25", "¾": "0.75", "⅓": "0.33", "⅔": "0.67"}


def _number(text):
    text = _FRACTIONS.get(text, text).replace(",", ".")
    if "/" in text:
        top, bottom = text.split("/")
        return str(round(float(top) / float(bottom), 2)) if float(bottom) else None
    return text


def parse_ingredient_line(line, normalize_unit):
    """Разбирает «Сливки жирностью 20 % - 150 мл». Возвращает словарь или None для заголовков.

    Поля: name, quantity (строка), quantity_max (верх диапазона или нет), unit,
    note — пояснение, если строку надо проверить.
    """
    line = " ".join(line.split())
    # Звёздочка отделяет группы, а строка с двоеточием на конце — их заголовок.
    if not re.search(r"\w", line) or line.endswith(":"):
        return None
    name, separator, amount = line, None, ""
    parts = re.split(r"(\s+[-–—]\s+)", line)
    if len(parts) >= 3:
        name, separator, amount = "".join(parts[:-2]), parts[-2], parts[-1]
    name = name.strip()
    needs_check = False

    if not separator:
        taste = re.search(r"\s*\bпо вкусу\b", name, re.IGNORECASE)
        if taste:
            name = name[: taste.start()].strip() or name
            return {"name": name, "quantity": "0", "unit": "to_taste", "note": None}
        joined = _NO_SEPARATOR.match(name)
        if not joined:
            return {"name": name, "quantity": "0", "unit": "to_taste",
                    "note": "количество на сайте не указано"}
        full_name = name
        name, amount = joined.group(1).strip(), joined.group(2)
        # Что-то вроде «Сливки 20 %» или «Бульон 2 кубика по 5 г»: разбор здесь только догадка.
        if normalize_unit(amount.lstrip("0123456789.,/-–— ½¼¾⅓⅔").rstrip(".") or "шт") is None:
            return {"name": full_name, "quantity": "0", "unit": "to_taste",
                    "note": "количество на сайте не указано"}
        needs_check = bool(re.search(r"\d", name))

    # Вариант в скобках — «180 г (1 шт.)»: берём основное значение.
    amount = re.sub(r"\([^)]*\)", " ", amount)
    amount = " ".join(amount.split())
    if re.fullmatch(r"по вкусу\.?", amount, re.IGNORECASE):
        return {"name": name, "quantity": "0", "unit": "to_taste", "note": None}

    match = _AMOUNT.match(amount)
    if not match:
        # Без числа: «щепотка» означает одну щепотку.
        unit = normalize_unit(amount.rstrip("."))
        if unit is None:
            return {"name": name, "quantity": "0", "unit": "to_taste",
                    "note": f"не удалось разобрать количество «{amount}»"}
        return {"name": name, "quantity": "1", "unit": unit, "note": None}

    low, high, unit_text = match.groups()
    quantity = _number(high)
    quantity_max = None
    if low and quantity is not None:
        bound = _number(low)
        if bound is not None and float(bound) != float(quantity):
            quantity, quantity_max = sorted((bound, quantity), key=float)
    unit_text = unit_text.strip().rstrip(".") or "шт"
    unit = normalize_unit(unit_text)
    note = None
    if needs_check:
        note = f"проверьте количество, на сайте написано «{line}»"
    if unit is None or quantity is None:
        return {"name": name, "quantity": "0", "unit": "to_taste",
                "note": f"не удалось разобрать количество «{amount}»"}
    return {"name": name, "quantity": quantity, "quantity_max": quantity_max, "unit": unit, "note": note}
