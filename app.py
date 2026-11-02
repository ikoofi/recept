"""Локальная книга рецептов: Flask + SQLite, экспорт рецепта в DOCX."""
import base64
import io
import json
import os
import re
import shutil
import sqlite3
import uuid
from contextlib import contextmanager
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urlsplit

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.exceptions import UnrecognizedImageError
from docx.shared import Cm, Pt
from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    url_for,
)
from werkzeug.datastructures import MultiDict

import site_import

BASE_DIR = Path(__file__).resolve().parent
# База лежит рядом с программой; RECEPT_DATA позволяет хранить её в другом месте.
DATA_DIR = Path(os.environ.get("RECEPT_DATA", BASE_DIR / "data"))
DB_PATH = DATA_DIR / "recipes.db"
PHOTOS_DIR = DATA_DIR / "photos"
IMPORTS_DIR = DATA_DIR / "imports"  # файлы импорта, ожидающие решения по дубликатам

# Только форматы, которые можно вставить в DOCX без дополнительных библиотек.
PHOTO_SIGNATURES = {b"\xff\xd8\xff": "jpg", b"\x89PNG\r\n\x1a\n": "png"}
PHOTO_MAX_BYTES = 10 * 1024 * 1024

WEEK_DAYS = ("Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье")
QUICK_MINUTES = 30
LOCAL_HOSTS = {"127.0.0.1", "localhost"}

# Отделы магазина. «Другое» — для того, что не относится к продуктам: например, салфетки.
CATEGORIES = (
    "Мясо и мясные продукты",
    "Рыба и морепродукты",
    "Молочные продукты",
    "Яйца и яичные продукты",
    "Плодоовощная группа",
    "Зерно-мучные и бакалейные продукты",
    "Жиры и масла",
    "Сахар, мёд, кондитерское сырьё",
    "Вкусовые и ароматические продукты",
    "Полуфабрикаты и готовые компоненты",
    "Другое",
)
# Частые продукты определяются сами, про остальные спрашиваем. Более точные шаблоны идут
# первыми: сливочное масло — молочный продукт, хотя в названии есть «масло». \\b — начало слова.
CATEGORY_GUESSES = (
    ("Молочные продукты", r"\bсливочн\w* масл|\bмасл[оа] сливочн"),
    ("Жиры и масла", r"\bмасл[оа]\b|\bмаргарин|\bсмалец|\bсало\b|\bспред"),
    ("Полуфабрикаты и готовые компоненты", r"\bфарш|\bтесто|\bбульон|\bконцентрат|овощн\w* смес"),
    ("Рыба и морепродукты", r"\bрыб|\bсель|\bсёмг|\bсемг|\bлосос|\bтреск|\bфорел|\bминта"
                            r"|\bкреветк|\bкальмар|\bмиди|\bтунец|\bскумбри|\bхек|\bикр"),
    ("Вкусовые и ароматические продукты",
     r"\bсоль\b|\bсоли\b|\bперец чёрн|\bчёрный перец|\bчерный перец|\bперец душист|\bперец молот"
     r"|\bперец красн|\bпаприк|\bзира|\bкорица|\bорегано|\bлавров|\bкарри|\bкуркум|\bхмели"
     r"|\bсунели|\bкориандр|\bшафран|\bваниль|\bгвоздик|\bприправ|\bтимьян|\bрозмарин|\bмускат"
     r"|\bмайонез|\bкетчуп|\bгорчиц|\bсоус|\bуксус|\bхрен|\bвин[оа]\b|\bконьяк|\bводк"),
    ("Плодоовощная группа",
     r"томатн\w* паст|консервирован|солён|солен|маринован|квашен|\bоливк|\bмаслин|\bкорнишон"
     r"|\bкаперс|\bкартоф|\bкартошк|\bморков|\bсвёкл|\bсвекл|\bкапуст|\bлук(?:а|ов\w*)?\b"
     r"|\bчеснок|\bпомидор|\bтомат|\bогур|\bперец|\bбаклажан|\bкабач|\bтыкв|\bредис|\bсельдере"
     r"|\bшпинат|\bукроп|\bпетрушк|\bкинз|\bбазилик|\bзелен|\bсалат|\bшампиньон|\bгриб|\bвешенк"
     r"|\bкукуруз|\bброкколи|\bимбир|\bяблок|\bгруш|\bбанан|\bапельсин|\bлимон|\bмандарин"
     r"|\bвиноград|\bперсик|\bабрикос|\bслив[аы]\b|\bвишн|\bчерешн|\bклубник|\bмалин|\bсмородин"
     r"|\bчерник|\bграна?т|\bананас|\bавокадо|\bкиви|\bгорош"),
    ("Мясо и мясные продукты", r"\bмяс|\bговя|\bсвин|\bбаран|\bтелят|\bкуриц|\bкурин|\bиндей|\bутк"
                               r"|\bгус|\bбекон|\bколбас|\bсосис|\bветчин|\bкопч|\bрёбр|\bребр"
                               r"|\bпечен[ьк]|\bвырезк|\bгрудинк|\bсердц|\bязык"),
    ("Яйца и яичные продукты", r"\bяйц|\bжелтк|\bбелк|\bмеланж|\bяичн\w* порош"),
    ("Зерно-мучные и бакалейные продукты",
     r"\bмук|\bрис\b|\bгреч|\bовсян|\bпшен|\bперлов|\bманн|\bкрупа|\bмакарон|\bспагетти|\bлапш"
     r"|\bпаста\b|\bвермишел|\bгорох|\bфасол|\bчечевиц|\bхлеб|\bбатон|\bбулк|\bбулоч|\bлаваш"
     r"|\bпанировоч|\bсухар|\bотруб|\bдрожж|\bразрыхлит|\bсод[аы]\b|\bкрахмал"),
    ("Молочные продукты", r"\bмолок|\bкефир|\bсметан|\bсливк|\bтворог|\bсыр\b|\bфета|\bпармезан"
                          r"|\bмоцарелл|\bйогурт|\bряженк"),
    ("Сахар, мёд, кондитерское сырьё", r"\bсахар|\bмёд|\bмед\b|\bшоколад|\bкакао|\bварень|\bорех"
                                       r"|\bминдал|\bфундук|\bизюм|\bкурага|\bчернослив|\bкокос|\bсироп"),
    ("Другое", r"\bвода\b|\bсок\b|\bкофе|\bчай"),
)

SORTS = {
    "title": "r.title_key",
    "new": "r.id DESC",
    "time": "r.prep_minutes IS NULL, r.prep_minutes, r.title_key",
}

MAX_MINUTES = 100000
MAX_SERVINGS = 1000

# Границы количества: мельче сотой значение округлится до нуля при показе.
QTY_MIN = Decimal("0.01")
QTY_MAX = Decimal("1000000000")

EXPORT_FORMAT = "recept"
EXPORT_VERSION = 2

# Порядок важен: так единицы показываются в выпадающем списке.
UNITS = {
    "g": "г",
    "kg": "кг",
    "ml": "мл",
    "l": "л",
    "tbsp": "ст. л.",
    "tsp": "ч. л.",
    "pcs": "шт.",
    "cloves": "зубч.",
    "to_taste": "по вкусу",
}

# Разные написания встроенных единиц: чтобы «гр.» не стало отдельной своей единицей.
UNIT_SYNONYMS = {
    "г": "g", "гр": "g", "грамм": "g", "граммов": "g", "граммы": "g",
    "кг": "kg", "килограмм": "kg", "килограммов": "kg",
    "мл": "ml", "миллилитр": "ml", "миллилитров": "ml",
    "л": "l", "литр": "l", "литра": "l", "литров": "l",
    "стл": "tbsp", "столоваяложка": "tbsp", "столовыхложек": "tbsp", "столовыеложки": "tbsp",
    "чл": "tsp", "чайнаяложка": "tsp", "чайныхложек": "tsp", "чайныеложки": "tsp",
    "стложка": "tbsp", "стложки": "tbsp", "стложек": "tbsp", "столоваяложки": "tbsp",
    "чложка": "tsp", "чложки": "tsp", "чложек": "tsp", "чайнаяложки": "tsp",
    "шт": "pcs", "штука": "pcs", "штук": "pcs", "штуки": "pcs",
    "зубч": "cloves", "зубчик": "cloves", "зубчика": "cloves", "зубчиков": "cloves",
    "повкусу": "to_taste",
}
UNIT_MAX_LENGTH = 20

# Как складывать единицы в списке покупок: (группа, множитель к базовой единице).
SHOPPING_UNITS = {
    "g": ("mass", 1),
    "kg": ("mass", 1000),
    "ml": ("volume", 1),
    "l": ("volume", 1000),
    "tbsp": ("tbsp", 1),
    "tsp": ("tsp", 1),
    "pcs": ("pcs", 1),
    "cloves": ("cloves", 1),
}


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "local-recipe-book")
# Файл импорта вместе с фотографиями может быть большим.
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024


def title_key(title):
    """Ключ для сравнения названий: SQLite сам не умеет без учёта регистра в кириллице."""
    return " ".join(title.split()).casefold()


def fold(value):
    return value.casefold() if isinstance(value, str) else value


def migrate_custom_units(conn):
    """Старые базы принимали только встроенные единицы: убираем это ограничение из таблицы."""
    sql = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'ingredients'"
    ).fetchone()[0]
    if "CHECK" not in sql:
        return
    conn.commit()
    shutil.copy2(DB_PATH, DATA_DIR / "recipes-before-custom-units.db")
    # SQLite не умеет снимать CHECK: пересоздаём таблицу, id ингредиентов сохраняются.
    conn.executescript(
        """
        BEGIN;
        CREATE TABLE ingredients_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            recipe_id INTEGER NOT NULL REFERENCES recipes(id) ON DELETE CASCADE,
            position INTEGER NOT NULL,
            name TEXT NOT NULL,
            quantity TEXT NOT NULL DEFAULT '0',
            unit TEXT NOT NULL,
            is_main INTEGER NOT NULL DEFAULT 0
        );
        INSERT INTO ingredients_new (id, recipe_id, position, name, quantity, unit, is_main)
            SELECT id, recipe_id, position, name, quantity, unit, is_main FROM ingredients;
        DROP TABLE ingredients;
        ALTER TABLE ingredients_new RENAME TO ingredients;
        CREATE INDEX idx_ingredients_recipe ON ingredients (recipe_id, position);
        COMMIT;
        """
    )


def init_db():
    for folder in (DATA_DIR, PHOTOS_DIR, IMPORTS_DIR):
        folder.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript((BASE_DIR / "schema.sql").read_text(encoding="utf-8"))
        # Базы старых версий: CREATE TABLE IF NOT EXISTS не добавляет новые столбцы.
        columns = {row[1] for row in conn.execute("PRAGMA table_info(recipes)")}
        for column, kind in (
            ("photo", "TEXT"),
            ("prep_minutes", "INTEGER"),
            ("servings", "INTEGER"),
            ("favorite", "INTEGER NOT NULL DEFAULT 0"),
            ("source_url", "TEXT"),
        ):
            if column not in columns:
                conn.execute(f"ALTER TABLE recipes ADD COLUMN {column} {kind}")
        migrate_custom_units(conn)
        columns = {row[1] for row in conn.execute("PRAGMA table_info(ingredients)")}
        if "quantity_max" not in columns:
            conn.execute("ALTER TABLE ingredients ADD COLUMN quantity_max TEXT")
        columns = {row[1] for row in conn.execute("PRAGMA table_info(shopping_recipes)")}
        for column in ("scale_by", "scale_qty"):
            if column not in columns:
                conn.execute(f"ALTER TABLE shopping_recipes ADD COLUMN {column} TEXT")


@contextmanager
def get_db():
    """Курсор в транзакции: коммит при успехе, откат при ошибке."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = lambda cur, row: {d[0]: v for d, v in zip(cur.description, row)}
    conn.execute("PRAGMA foreign_keys = ON")
    conn.create_function("fold", 1, fold, deterministic=True)
    try:
        yield conn.cursor()
        conn.commit()
    except BaseException:
        conn.rollback()
        raise
    finally:
        conn.close()


init_db()


def format_number(value):
    # Округление до сотых нужно для пересчитанных значений (например, 1/3 порции).
    q = Decimal(value).quantize(Decimal("0.01")).normalize()
    return format(q, "f").replace(".", ",")


def unit_label(unit):
    if unit == "servings":  # расчёт по порциям, см. scale_recipe
        return "порц."
    return UNITS.get(unit, unit)


def normalize_unit(value):
    """Код встроенной единицы или своя единица в нижнем регистре. None, если не подходит."""
    label = " ".join((value or "").split()).casefold()
    if value in UNITS:
        return value
    key = label.replace(" ", "").replace(".", "")
    if key in UNIT_SYNONYMS:
        return UNIT_SYNONYMS[key]
    # __new__ — служебный пункт «Своя единица», а servings — расчёт по порциям.
    if not key or len(label) > UNIT_MAX_LENGTH or label in ("servings", "__new__"):
        return None
    if not re.search(r"[^\W\d_]", label):  # хотя бы одна буква
        return None
    return label


def format_quantity(quantity, unit, quantity_max=None):
    if unit == "to_taste":
        return UNITS[unit]
    if quantity_max is not None:
        return f"{format_number(quantity)}–{format_number(quantity_max)} {unit_label(unit)}"
    return f"{format_number(quantity)} {unit_label(unit)}"


def ingredient_quantity(ing):
    """Одно число для расчёта. У диапазона это его середина: 160–200 считается как 180."""
    low = Decimal(ing["quantity"])
    high = ing.get("quantity_max")
    return (low + Decimal(high)) / 2 if high is not None else low


def plain_number(value):
    return format(Decimal(value).quantize(Decimal("0.01")).normalize(), "f")


def format_minutes(minutes):
    hours, rest = divmod(minutes, 60)
    parts = [f"{hours} ч" if hours else "", f"{rest} мин" if rest else ""]
    return " ".join(p for p in parts if p)


def photo_url(recipe):
    if recipe["photo"]:
        return url_for("photo", name=recipe["photo"])
    return url_for("static", filename="placeholder.svg")


app.jinja_env.filters["qty"] = lambda ing: format_quantity(
    ing["quantity"], ing["unit"], ing.get("quantity_max")
)


def qty_input(ing):
    """Значение поля количества в форме: «160-200» для диапазона."""
    if ing["unit"] == "to_taste":
        return ""
    if ing.get("quantity_max") is not None:
        return f"{plain_number(ing['quantity'])}-{plain_number(ing['quantity_max'])}"
    return plain_number(ing["quantity"])


app.jinja_env.filters["qty_input"] = qty_input
app.jinja_env.filters["scale_value"] = lambda ing: plain_number(ingredient_quantity(ing))
app.jinja_env.filters["number"] = format_number
app.jinja_env.filters["unit_label"] = unit_label
app.jinja_env.filters["minutes"] = format_minutes
app.jinja_env.filters["photo_url"] = photo_url
app.jinja_env.filters["name_key"] = title_key
app.jinja_env.filters["source_label"] = lambda url: urlsplit(url).hostname or url


def guess_category(name):
    key = title_key(name)
    return next((c for c, pattern in CATEGORY_GUESSES if re.search(pattern, key)), None)


def stored_categories(cur):
    cur.execute("SELECT name_key, category FROM ingredient_categories")
    # Отдел, которого больше нет в списке, считается невыбранным: книга спросит снова.
    return {row["name_key"]: row["category"] for row in cur.fetchall() if row["category"] in CATEGORIES}


def set_category(cur, name, category):
    if category in CATEGORIES:
        cur.execute(
            "INSERT OR REPLACE INTO ingredient_categories (name_key, category) VALUES (?, ?)",
            (title_key(name), category),
        )


def known_category(cur, name):
    """Отдел продукта: выбранный пользователем, иначе из встроенного словаря, иначе None."""
    cur.execute(
        "SELECT category FROM ingredient_categories WHERE name_key = ?", (title_key(name),)
    )
    row = cur.fetchone()
    return row["category"] if row and row["category"] in CATEGORIES else guess_category(name)


@app.before_request
def protect_book():
    """Книгу могут использовать только её собственные страницы.

    Host защищает от подмены адреса сайта, Origin и Sec-Fetch-Site — от команд с чужих страниц,
    открытых в том же браузере. На Android к серверу могут обратиться другие приложения
    телефона, поэтому там каждый запрос ещё обязан нести секретную cookie из MainActivity.
    """
    if request.host.rsplit(":", 1)[0] not in LOCAL_HOSTS:
        abort(403)
    token = app.config.get("ACCESS_TOKEN")
    if token and request.endpoint != "unlock" and request.cookies.get("access") != token:
        abort(403)
    if request.method in ("GET", "HEAD", "OPTIONS"):
        return
    origin = request.headers.get("Origin")
    if origin is not None and origin != request.host_url.rstrip("/"):
        abort(403)
    if request.headers.get("Sec-Fetch-Site") not in (None, "same-origin", "none"):
        abort(403)


def parse_quantity(value):
    """Количество из строки ввода, либо None, если это не число из допустимого диапазона."""
    try:
        q = Decimal((value or "").strip().replace(",", "."))
    except InvalidOperation:
        return None
    # is_finite отсекает NaN и Infinity, которые Decimal тоже считает числами.
    return q if q.is_finite() and QTY_MIN <= q <= QTY_MAX else None


def parse_quantity_range(value):
    """Количество или диапазон из поля формы: (от, до). None, если значение не подходит."""
    parts = re.split(r"\s*[-–—]\s*", (value or "").strip())
    if len(parts) == 1:
        low = parse_quantity(parts[0])
        return (low, None) if low is not None else None
    if len(parts) != 2:
        return None
    low, high = parse_quantity(parts[0]), parse_quantity(parts[1])
    if low is None or high is None:
        return None
    if low == high:
        return low, None
    return min(low, high), max(low, high)


def scale_recipe(recipe, args):
    """Пересчитывает рецепт из ?by=<id ингредиента или servings>&qty=<количество>.

    В recipe["scale_base"] кладёт то, по чему идёт расчёт: выбранный вариант, иначе
    главный ингредиент, иначе первый с количеством. recipe["scaled"] — был ли пересчёт.
    """
    options = [i for i in recipe["ingredients"] if i["unit"] != "to_taste"]
    servings = None
    if recipe["servings"]:
        servings = {
            "id": "servings",
            "name": "Порции",
            "quantity": recipe["servings"],
            "unit": "servings",
            "is_main": False,
        }
        options.insert(0, servings)
    by = args.get("by")
    base = next((i for i in options if str(i["id"]) == by), None)
    if base is None:
        base = next((i for i in options if i["is_main"]), options[0] if options else None)

    recipe["scale_options"] = options
    recipe["scale_base"] = base
    recipe["scaled"] = False
    qty = parse_quantity(args.get("qty"))
    if qty is None or base is None or not ingredient_quantity(base):
        return recipe
    factor = qty / ingredient_quantity(base)
    for ing in recipe["ingredients"]:
        ing["quantity"] = Decimal(ing["quantity"]) * factor
        if ing.get("quantity_max") is not None:
            ing["quantity_max"] = Decimal(ing["quantity_max"]) * factor
    if servings:
        servings["quantity"] = Decimal(servings["quantity"]) * factor
        recipe["servings"] = servings["quantity"]
    recipe["scaled"] = True
    return recipe


def fetch_recipe(recipe_id):
    with get_db() as cur:
        cur.execute("SELECT * FROM recipes WHERE id = ?", (recipe_id,))
        recipe = cur.fetchone()
        if recipe is None:
            abort(404)
        cur.execute(
            "SELECT * FROM ingredients WHERE recipe_id = ? ORDER BY position", (recipe_id,)
        )
        recipe["ingredients"] = cur.fetchall()
        cur.execute("SELECT * FROM steps WHERE recipe_id = ? ORDER BY position", (recipe_id,))
        recipe["steps"] = cur.fetchall()
        cur.execute(
            "SELECT t.name FROM tags t JOIN recipe_tags rt ON rt.tag_id = t.id"
            " WHERE rt.recipe_id = ? ORDER BY t.name_key",
            (recipe_id,),
        )
        recipe["tags"] = [row["name"] for row in cur.fetchall()]
    recipe["main"] = next((i for i in recipe["ingredients"] if i["is_main"]), None)
    return recipe


def parse_count(value, label, limit, errors):
    """Необязательное целое число из формы."""
    text = (value or "").strip()
    if not text:
        return None
    if not text.isdecimal() or len(text) > 6 or not 1 <= int(text) <= limit:
        errors.append(f"{label}: нужно целое число от 1 до {limit}.")
        return None
    return int(text)


def parse_tags(text):
    tags = {}
    for part in (text or "").split(","):
        name = " ".join(part.split())
        if name:
            tags.setdefault(title_key(name), name)
    return list(tags.values())


def parse_form(form):
    """Собирает рецепт из формы. Возвращает (recipe, errors)."""
    errors = []
    title = form.get("title", "").strip()
    if not title:
        errors.append("Укажите название рецепта.")
    author = form.get("author", "").strip() or None
    source_url = form.get("source_url", "").strip() or None
    # В ссылку на странице рецепта нельзя пускать ничего, кроме обычных адресов.
    if source_url and not re.match(r"https?://", source_url):
        errors.append("Ссылка на источник должна начинаться с http:// или https://.")
        source_url = None
    prep_minutes = parse_count(form.get("prep_minutes"), "Время приготовления", MAX_MINUTES, errors)
    servings = parse_count(form.get("servings"), "Порции", MAX_SERVINGS, errors)

    main_key = form.get("main")
    ingredients = []
    rows = zip(
        form.getlist("ing_key"),
        form.getlist("ing_name"),
        form.getlist("ing_qty"),
        form.getlist("ing_unit"),
    )
    categories = form.getlist("ing_category")
    for n, (key, name, qty, unit) in enumerate(rows):
        name = name.strip()
        if not name:
            continue
        unit = normalize_unit(unit)
        if unit is None:
            errors.append(
                f"Укажите единицу измерения для «{name}»: до {UNIT_MAX_LENGTH} символов, с буквами."
            )
            continue
        quantity_max = None
        if unit == "to_taste":
            quantity = Decimal(0)
        else:
            parsed = parse_quantity_range(qty)
            if parsed is None:
                errors.append(
                    f"Количество для «{name}»: число от 0,01 или диапазон, например 160-200."
                )
                continue
            quantity, quantity_max = parsed
        ingredients.append(
            {
                "name": name,
                "quantity": quantity,
                "quantity_max": quantity_max,
                "unit": unit,
                "is_main": key == main_key,
                "category": categories[n] if n < len(categories) else "",
            }
        )

    if not ingredients:
        errors.append("Добавьте хотя бы один ингредиент.")
    elif sum(i["is_main"] for i in ingredients) != 1:
        errors.append("Отметьте главный ингредиент.")
    elif next(i for i in ingredients if i["is_main"])["unit"] == "to_taste":
        errors.append("У главного ингредиента должно быть количество, а не «по вкусу».")

    steps = [{"text": s.strip()} for s in form.getlist("step") if s.strip()]
    if not steps:
        errors.append("Добавьте хотя бы один шаг приготовления.")

    recipe = {
        "title": title,
        "author": author,
        "photo": None,
        "source_url": source_url,
        "draft_photo": form.get("draft_photo") or None,
        "prep_minutes": prep_minutes,
        "servings": servings,
        "tags": parse_tags(form.get("tags")),
        "ingredients": ingredients,
        "steps": steps,
    }
    return recipe, errors


def require_categories(recipe, errors):
    """При вводе вручную у нового продукта должен быть выбран отдел магазина."""
    with get_db() as cur:
        for ing in recipe["ingredients"]:
            if ing["category"] not in CATEGORIES and not known_category(cur, ing["name"]):
                errors.append(f"Выберите категорию для «{ing['name']}».")


def photo_extension(data):
    return next((ext for sig, ext in PHOTO_SIGNATURES.items() if data.startswith(sig)), None)


def read_photo_upload(file, errors):
    """Содержимое загруженного фото, либо None, если его нет или оно не подходит."""
    if file is None or not file.filename:
        return None
    data = file.read()
    if not photo_extension(data):
        errors.append("Фото должно быть в формате JPG или PNG.")
        return None
    if len(data) > PHOTO_MAX_BYTES:
        errors.append("Фото должно быть не больше 10 МБ.")
        return None
    return data


def store_photo(data):
    name = f"{uuid.uuid4().hex}.{photo_extension(data)}"
    (PHOTOS_DIR / name).write_bytes(data)
    return name


def delete_photo(name):
    if name:
        (PHOTOS_DIR / name).unlink(missing_ok=True)


def title_exists(cur, title, exclude_id=None):
    cur.execute(
        "SELECT 1 FROM recipes WHERE title_key = ? AND id <> ? LIMIT 1",
        (title_key(title), exclude_id or 0),
    )
    return cur.fetchone() is not None


def insert_recipe(cur, recipe):
    cur.execute(
        "INSERT INTO recipes (title, title_key, author, photo, prep_minutes, servings, source_url)"
        " VALUES (?, ?, ?, ?, ?, ?, ?)",
        (
            recipe["title"],
            title_key(recipe["title"]),
            recipe["author"],
            recipe["photo"],
            recipe["prep_minutes"],
            recipe["servings"],
            recipe.get("source_url"),
        ),
    )
    recipe_id = cur.lastrowid
    save_children(cur, recipe_id, recipe)
    return recipe_id


def update_recipe(cur, recipe_id, recipe):
    cur.execute(
        "UPDATE recipes SET title = ?, title_key = ?, author = ?, photo = ?, prep_minutes = ?,"
        " servings = ?, source_url = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (
            recipe["title"],
            title_key(recipe["title"]),
            recipe["author"],
            recipe["photo"],
            recipe["prep_minutes"],
            recipe["servings"],
            recipe.get("source_url"),
            recipe_id,
        ),
    )
    save_children(cur, recipe_id, recipe)
    # Ингредиенты получили новые id, сохранённый расчёт в списке покупок на них не укажет.
    cur.execute(
        "UPDATE shopping_recipes SET scale_by = NULL, scale_qty = NULL WHERE recipe_id = ?",
        (recipe_id,),
    )


def save_children(cur, recipe_id, recipe):
    cur.execute("DELETE FROM ingredients WHERE recipe_id = ?", (recipe_id,))
    cur.execute("DELETE FROM steps WHERE recipe_id = ?", (recipe_id,))
    cur.executemany(
        "INSERT INTO ingredients (recipe_id, position, name, quantity, quantity_max, unit, is_main)"
        " VALUES (?, ?, ?, ?, ?, ?, ?)",
        [
            (
                recipe_id,
                pos,
                i["name"],
                format(Decimal(i["quantity"]), "f"),
                format(Decimal(i["quantity_max"]), "f") if i.get("quantity_max") is not None else None,
                i["unit"],
                i["is_main"],
            )
            for pos, i in enumerate(recipe["ingredients"], 1)
        ],
    )
    cur.executemany(
        "INSERT INTO steps (recipe_id, position, text) VALUES (?, ?, ?)",
        [(recipe_id, pos, s["text"]) for pos, s in enumerate(recipe["steps"], 1)],
    )
    for ing in recipe["ingredients"]:
        # Отдел из файла импорта не перезаписывает то, что пользователь уже выбрал сам.
        if ing.get("category") and known_category(cur, ing["name"]) is None:
            set_category(cur, ing["name"], ing["category"])
    cur.execute("DELETE FROM recipe_tags WHERE recipe_id = ?", (recipe_id,))
    for name in recipe["tags"]:
        key = title_key(name)
        cur.execute("INSERT OR IGNORE INTO tags (name, name_key) VALUES (?, ?)", (name, key))
        cur.execute(
            "INSERT INTO recipe_tags (recipe_id, tag_id)"
            " SELECT ?, id FROM tags WHERE name_key = ?",
            (recipe_id, key),
        )
    prune_tags(cur)


def prune_tags(cur):
    # Категория исчезает из списков вместе с последним рецептом.
    cur.execute("DELETE FROM tags WHERE id NOT IN (SELECT tag_id FROM recipe_tags)")


def all_tags(cur):
    cur.execute("SELECT name, name_key FROM tags ORDER BY name_key")
    return cur.fetchall()


def render_form(recipe, mode):
    with get_db() as cur:
        tags = all_tags(cur)
        cur.execute("SELECT name FROM ingredients ORDER BY id")
        names = {}
        for row in cur.fetchall():
            names.setdefault(title_key(row["name"]), row["name"])
        # Своя единица живёт, пока её используют рецепты: опечатка исчезнет вместе с ней.
        cur.execute("SELECT DISTINCT unit FROM ingredients ORDER BY unit")
        custom_units = [row["unit"] for row in cur.fetchall() if row["unit"] not in UNITS]
    return render_template(
        "form.html",
        recipe=recipe,
        units=UNITS,
        mode=mode,
        categories=CATEGORIES,
        custom_units=custom_units,
        all_tags=tags,
        ingredient_names=[names[k] for k in sorted(names)],
    )


def search_recipes(cur, q="", tag="", favorite=False, quick=False, sort="title"):
    sql = (
        "SELECT r.id, r.title, r.author, r.photo, r.prep_minutes, r.servings, r.favorite,"
        " i.name AS main_ingredient"
        " FROM recipes r LEFT JOIN ingredients i ON i.recipe_id = r.id AND i.is_main"
    )
    where, params = [], []
    if q:
        where.append("(fold(r.title) LIKE ? OR fold(r.author) LIKE ? OR fold(i.name) LIKE ?)")
        params += [f"%{q.casefold()}%"] * 3
    if tag:
        where.append(
            "EXISTS (SELECT 1 FROM recipe_tags rt JOIN tags t ON t.id = rt.tag_id"
            " WHERE rt.recipe_id = r.id AND t.name_key = ?)"
        )
        params.append(tag)
    if favorite:
        where.append("r.favorite = 1")
    if quick:
        where.append("r.prep_minutes <= ?")
        params.append(QUICK_MINUTES)
    if where:
        sql += " WHERE " + " AND ".join(where)
    # Порядок берётся только из SORTS, поэтому подстановка в SQL безопасна.
    sql += " ORDER BY " + SORTS.get(sort, SORTS["title"])
    cur.execute(sql, params)
    return cur.fetchall()


@app.context_processor
def page_context():
    with get_db() as cur:
        cur.execute("SELECT COUNT(*) AS n FROM shopping_recipes")
        count = cur.fetchone()["n"]
    # Тема хранится в cookie: они не зависят от порта, а на Android он при каждом запуске новый.
    theme = request.cookies.get("theme")
    return {"shopping_count": count, "theme": theme if theme in ("light", "dark") else None}


@app.route("/")
def index():
    filters = {
        "q": request.args.get("q", "").strip() or None,
        "tag": request.args.get("tag", "").strip() or None,
        "fav": request.args.get("fav") or None,
        "quick": request.args.get("quick") or None,
        "sort": request.args.get("sort") if request.args.get("sort") in SORTS else None,
    }
    with get_db() as cur:
        recipes = search_recipes(
            cur,
            filters["q"] or "",
            filters["tag"] or "",
            favorite=bool(filters["fav"]),
            quick=bool(filters["quick"]),
            sort=filters["sort"] or "title",
        )
        tags = all_tags(cur)
        cur.execute(
            "SELECT rt.recipe_id, t.name FROM recipe_tags rt JOIN tags t ON t.id = rt.tag_id"
            " ORDER BY t.name_key"
        )
        by_recipe = {}
        for row in cur.fetchall():
            by_recipe.setdefault(row["recipe_id"], []).append(row["name"])
    for recipe in recipes:
        recipe["tags"] = by_recipe.get(recipe["id"], [])
    return render_template(
        "index.html", recipes=recipes, filters=filters, all_tags=tags, quick_minutes=QUICK_MINUTES
    )


def local_redirect(target, fallback):
    """Возврат на страницу из формы: только в пределах самой книги."""
    if target and target.startswith("/") and not target.startswith("//"):
        return redirect(target)
    return redirect(fallback)


@app.route("/recipe/<int:recipe_id>/favorite", methods=["POST"])
def recipe_favorite(recipe_id):
    with get_db() as cur:
        cur.execute("UPDATE recipes SET favorite = 1 - favorite WHERE id = ?", (recipe_id,))
    return local_redirect(request.form.get("back"), url_for("recipe_view", recipe_id=recipe_id))


@app.route("/recipe/<int:recipe_id>/copy", methods=["POST"])
def recipe_copy(recipe_id):
    recipe = fetch_recipe(recipe_id)
    for ing in recipe["ingredients"]:
        ing["quantity"] = Decimal(ing["quantity"])
    photo_path = PHOTOS_DIR / recipe["photo"] if recipe["photo"] else None
    # У копии свой файл: иначе удаление одного рецепта оставит другой без фото.
    if photo_path and photo_path.exists():
        recipe["photo"] = store_photo(photo_path.read_bytes())
    else:
        recipe["photo"] = None
    with get_db() as cur:
        recipe["title"] = free_copy_title(cur, recipe["title"])
        new_id = insert_recipe(cur, recipe)
    flash(f"Создана копия «{recipe['title']}». Измените её и сохраните.")
    return redirect(url_for("recipe_edit", recipe_id=new_id))


DRAFT_PHOTO_NAME = re.compile(r"draft_[0-9a-f]{32}\.(?:jpg|png)")


def read_draft_photo(name):
    """Фото рецепта, импортированного по ссылке. Файл живёт только до сохранения."""
    if not name or not DRAFT_PHOTO_NAME.fullmatch(name):
        return None
    path = IMPORTS_DIR / name
    if not path.exists():
        return None
    data = path.read_bytes()
    path.unlink()
    return data


@app.route("/import/draft/<name>")
def draft_photo(name):
    if not DRAFT_PHOTO_NAME.fullmatch(name):
        abort(404)
    return send_from_directory(IMPORTS_DIR, name)


@app.route("/recipe/fetch", methods=["POST"])
def recipe_fetch():
    """Заполняет форму рецептом с сайта. В книгу он попадёт только после проверки и сохранения."""
    url = request.form.get("url", "").strip()
    try:
        page = site_import.parse_page(site_import.fetch_page(url), url)
    except site_import.SiteImportError as exc:
        flash(str(exc), "error")
        return render_form(None, "new")

    ingredients, notes = [], []
    for line in page["lines"]:
        ing = site_import.parse_ingredient_line(line, normalize_unit)
        if ing is None:
            continue
        ing["quantity"] = Decimal(ing["quantity"])
        ing["quantity_max"] = Decimal(ing["quantity_max"]) if ing.get("quantity_max") else None
        ing["is_main"] = False
        ing["category"] = ""
        ingredients.append(ing)
        if ing["note"]:
            notes.append(f"«{ing['name']}»: {ing['note']}.")
    # Сайт не указывает главный ингредиент. Предлагаем первый с количеством.
    first = next((i for i in ingredients if i["unit"] != "to_taste"), None)
    if first:
        first["is_main"] = True

    for old in IMPORTS_DIR.glob("draft_*"):
        old.unlink(missing_ok=True)
    draft = None
    image = site_import.fetch_image(page["image_url"]) if page["image_url"] else None
    if image and photo_extension(image) and len(image) <= PHOTO_MAX_BYTES:
        draft = f"draft_{uuid.uuid4().hex}.{photo_extension(image)}"
        (IMPORTS_DIR / draft).write_bytes(image)

    recipe = {
        "title": page["title"],
        "author": None,
        "photo": None,
        "draft_photo": draft,
        "source_url": url,
        "prep_minutes": page["prep_minutes"],
        "servings": page["servings"],
        "tags": [],
        "ingredients": ingredients,
        "steps": [{"text": s} for s in page["steps"]],
    }
    flash(
        "Рецепт заполнен с сайта, но ещё не сохранён. Проверьте количества, "
        "главный ингредиент и нажмите «Сохранить»."
    )
    for note in notes:
        flash(note, "error")
    if not page["steps"]:
        flash("Шаги приготовления на странице не нашлись, добавьте их вручную.", "error")
    return render_form(recipe, "new")


@app.route("/recipe/new", methods=["GET", "POST"])
def recipe_new():
    if request.method == "GET":
        return render_form(None, "new")
    recipe, errors = parse_form(request.form)
    require_categories(recipe, errors)
    upload = read_photo_upload(request.files.get("photo"), errors)
    if not errors:
        with get_db() as cur:
            for ing in recipe["ingredients"]:
                set_category(cur, ing["name"], ing["category"])
            if title_exists(cur, recipe["title"]):
                errors.append(f"Рецепт «{recipe['title']}» уже есть в книге.")
            else:
                upload = upload or read_draft_photo(recipe["draft_photo"])
                if upload:
                    recipe["photo"] = store_photo(upload)
                recipe_id = insert_recipe(cur, recipe)
    if not errors:
        flash("Рецепт сохранён.")
        return redirect(url_for("recipe_view", recipe_id=recipe_id))
    for e in errors:
        flash(e, "error")
    return render_form(recipe, "new")


@app.route("/recipe/<int:recipe_id>")
def recipe_view(recipe_id):
    recipe = scale_recipe(fetch_recipe(recipe_id), request.args)
    with get_db() as cur:
        cur.execute("SELECT 1 FROM shopping_recipes WHERE recipe_id = ?", (recipe_id,))
        in_shopping = cur.fetchone() is not None
    return render_template("view.html", recipe=recipe, units=UNITS, in_shopping=in_shopping)


@app.route("/recipe/<int:recipe_id>/cook")
def recipe_cook(recipe_id):
    recipe = scale_recipe(fetch_recipe(recipe_id), request.args)
    return render_template("cook.html", recipe=recipe)


@app.route("/photo/<name>")
def photo(name):
    return send_from_directory(PHOTOS_DIR, name)


@app.route("/recipe/<int:recipe_id>/edit", methods=["GET", "POST"])
def recipe_edit(recipe_id):
    existing = fetch_recipe(recipe_id)
    if request.method == "GET":
        return render_form(existing, "edit")
    recipe, errors = parse_form(request.form)
    require_categories(recipe, errors)
    upload = read_photo_upload(request.files.get("photo"), errors)
    recipe["photo"] = existing["photo"]
    if not errors:
        with get_db() as cur:
            for ing in recipe["ingredients"]:
                set_category(cur, ing["name"], ing["category"])
            if title_exists(cur, recipe["title"], exclude_id=recipe_id):
                errors.append(f"Рецепт «{recipe['title']}» уже есть в книге.")
            else:
                if upload:
                    recipe["photo"] = store_photo(upload)
                elif "photo_remove" in request.form:
                    recipe["photo"] = None
                update_recipe(cur, recipe_id, recipe)
    if not errors:
        if recipe["photo"] != existing["photo"]:
            delete_photo(existing["photo"])
        flash("Изменения сохранены.")
        return redirect(url_for("recipe_view", recipe_id=recipe_id))
    for e in errors:
        flash(e, "error")
    recipe["id"] = recipe_id
    return render_form(recipe, "edit")


@app.route("/recipe/<int:recipe_id>/delete", methods=["POST"])
def recipe_delete(recipe_id):
    with get_db() as cur:
        cur.execute("SELECT photo FROM recipes WHERE id = ?", (recipe_id,))
        row = cur.fetchone()
        cur.execute("DELETE FROM recipes WHERE id = ?", (recipe_id,))
        prune_tags(cur)
    if row:
        delete_photo(row["photo"])
    flash("Рецепт удалён.")
    return redirect(url_for("index"))


def shopping_recipes():
    """Выбранные рецепты, уже пересчитанные так, как задано в списке покупок."""
    with get_db() as cur:
        cur.execute(
            "SELECT recipe_id, scale_by, scale_qty FROM shopping_recipes"
            " ORDER BY added_at, recipe_id"
        )
        rows = cur.fetchall()
    return [
        scale_recipe(fetch_recipe(row["recipe_id"]), {"by": row["scale_by"], "qty": row["scale_qty"]})
        for row in rows
    ]


def build_shopping_list(recipes):
    """Складывает одинаковые ингредиенты рецептов. Названия сравниваются без учёта регистра."""
    items = {}
    for recipe in recipes:
        for ing in recipe["ingredients"]:
            item = items.setdefault(
                title_key(ing["name"]), {"name": ing["name"], "amounts": {}, "to_taste": False}
            )
            if ing["unit"] == "to_taste":
                item["to_taste"] = True
            else:
                # Своя единица складывается только сама с собой.
                group, multiplier = SHOPPING_UNITS.get(ing["unit"], (ing["unit"], 1))
                low = Decimal(ing["quantity"]) * multiplier
                high = Decimal(ing.get("quantity_max") or ing["quantity"]) * multiplier
                # Диапазоны складываются по границам: 160–200 г + 50 г = 210–250 г.
                total = item["amounts"].get(group, (Decimal(0), Decimal(0)))
                item["amounts"][group] = (total[0] + low, total[1] + high)
    return [
        {"key": key, "name": items[key]["name"], "amount": format_shopping_amount(items[key])}
        for key in sorted(items)
    ]


def format_shopping_amount(item):
    parts = []
    for group, (low, high) in item["amounts"].items():
        unit, divisor = group, 1
        if group == "mass":
            unit, divisor = ("kg", 1000) if high >= 1000 else ("g", 1)
        elif group == "volume":
            unit, divisor = ("l", 1000) if high >= 1000 else ("ml", 1)
        parts.append(
            format_quantity(low / divisor, unit, high / divisor if high != low else None)
        )
    if item["to_taste"]:
        parts.append(UNITS["to_taste"])
    return " + ".join(parts)


@app.route("/shopping")
def shopping():
    """Шаг 1: выбор рецептов и пересчёт каждого."""
    q = request.args.get("q", "").strip()
    selected = shopping_recipes()
    chosen_ids = {r["id"] for r in selected}
    with get_db() as cur:
        found = [r for r in search_recipes(cur, q) if r["id"] not in chosen_ids]
    return render_template("shopping.html", selected=selected, found=found, q=q, units=UNITS)


def shopping_groups():
    """Содержимое списка покупок: разбивка по отделам, продукты без отдела и то, что есть дома."""
    recipes = shopping_recipes()
    products = build_shopping_list(recipes)
    with get_db() as cur:
        stored = stored_categories(cur)
        cur.execute("SELECT name_key FROM pantry")
        pantry = {row["name_key"] for row in cur.fetchall()}
        cur.execute("SELECT id, name, amount FROM shopping_custom ORDER BY id")
        custom = cur.fetchall()
    groups = {category: [] for category in CATEGORIES}
    unsorted_items, at_home = [], []
    entries = [dict(p, custom_id=None) for p in products]
    entries += [
        {"key": f"custom:{c['id']}", "name": c["name"], "amount": c["amount"], "custom_id": c["id"]}
        for c in custom
    ]
    for item in entries:
        name_key = title_key(item["name"])
        if item["custom_id"] is None and name_key in pantry:
            at_home.append(item)
            continue
        category = stored.get(name_key) or guess_category(item["name"])
        if category:
            groups[category].append(item)
        else:
            unsorted_items.append(item)
    return {
        "recipes": recipes,
        "groups": [(c, groups[c]) for c in CATEGORIES if groups[c]],
        "unsorted": unsorted_items,
        "at_home": at_home,
        "keys": {item["key"] for item in entries},
    }


@app.route("/shopping/list")
def shopping_list():
    """Шаг 2: итог по отделам магазина."""
    data = shopping_groups()
    with get_db() as cur:
        cur.execute("SELECT item_key FROM shopping_checks")
        checks = {row["item_key"] for row in cur.fetchall()}
        # Продукт, которого больше нет в списке, не должен оказаться отмеченным при возвращении.
        for key in checks - data["keys"]:
            cur.execute("DELETE FROM shopping_checks WHERE item_key = ?", (key,))
    return render_template(
        "shopping_list.html", checks=checks & data["keys"], categories=CATEGORIES, **data
    )


@app.route("/shopping/custom/add", methods=["POST"])
def shopping_custom_add():
    name = " ".join(request.form.get("name", "").split())
    amount = " ".join(request.form.get("amount", "").split())
    if name:
        with get_db() as cur:
            cur.execute("INSERT INTO shopping_custom (name, amount) VALUES (?, ?)", (name, amount))
            set_category(cur, name, request.form.get("category"))
    return redirect(url_for("shopping_list"))


@app.route("/shopping/custom/remove", methods=["POST"])
def shopping_custom_remove():
    item_id = request.form.get("id", type=int)
    with get_db() as cur:
        cur.execute("DELETE FROM shopping_custom WHERE id = ?", (item_id,))
        cur.execute("DELETE FROM shopping_checks WHERE item_key = ?", (f"custom:{item_id}",))
    return redirect(url_for("shopping_list"))


@app.route("/ingredient-category")
def ingredient_category():
    """Форма спрашивает здесь, знает ли книга отдел продукта."""
    with get_db() as cur:
        return jsonify(category=known_category(cur, request.args.get("name", "")))


@app.route("/products/category", methods=["POST"])
def product_category():
    with get_db() as cur:
        set_category(cur, request.form.get("name", ""), request.form.get("category"))
    return local_redirect(request.form.get("back"), url_for("shopping_list"))


@app.route("/products/home", methods=["POST"])
def product_home():
    key = title_key(request.form.get("name", ""))
    with get_db() as cur:
        if request.form.get("at_home") == "1":
            cur.execute("INSERT OR IGNORE INTO pantry (name_key) VALUES (?)", (key,))
        else:
            cur.execute("DELETE FROM pantry WHERE name_key = ?", (key,))
    return local_redirect(request.form.get("back"), url_for("shopping_list"))


@app.route("/shopping/check", methods=["POST"])
def shopping_check():
    key = request.form.get("key", "")
    with get_db() as cur:
        if request.form.get("checked") == "1":
            cur.execute("INSERT OR IGNORE INTO shopping_checks (item_key) VALUES (?)", (key,))
        else:
            cur.execute("DELETE FROM shopping_checks WHERE item_key = ?", (key,))
    return "", 204


@app.route("/menu")
def menu():
    with get_db() as cur:
        cur.execute(
            "SELECT m.id, m.day, r.id AS recipe_id, r.title FROM menu_items m"
            " JOIN recipes r ON r.id = m.recipe_id ORDER BY m.day, m.id"
        )
        rows = cur.fetchall()
        recipes = search_recipes(cur)
    days = [{"index": i, "name": name, "items": []} for i, name in enumerate(WEEK_DAYS)]
    for row in rows:
        days[row["day"]]["items"].append(row)
    return render_template("menu.html", days=days, recipes=recipes, has_items=bool(rows))


@app.route("/menu/add", methods=["POST"])
def menu_add():
    day = request.form.get("day", type=int)
    recipe_id = request.form.get("recipe_id", type=int)
    if day in range(len(WEEK_DAYS)):
        with get_db() as cur:
            cur.execute(
                "INSERT OR IGNORE INTO menu_items (day, recipe_id)"
                " SELECT ?, id FROM recipes WHERE id = ?",
                (day, recipe_id),
            )
    return redirect(url_for("menu"))


@app.route("/menu/remove", methods=["POST"])
def menu_remove():
    with get_db() as cur:
        cur.execute("DELETE FROM menu_items WHERE id = ?", (request.form.get("id", type=int),))
    return redirect(url_for("menu"))


@app.route("/menu/clear", methods=["POST"])
def menu_clear():
    with get_db() as cur:
        cur.execute("DELETE FROM menu_items")
    return redirect(url_for("menu"))


@app.route("/menu/to-shopping", methods=["POST"])
def menu_to_shopping():
    """Добавляет блюда меню в покупки. Блюдо на несколько дней умножается на число дней."""
    with get_db() as cur:
        cur.execute("SELECT recipe_id, COUNT(*) AS days FROM menu_items GROUP BY recipe_id")
        counts = cur.fetchall()
    scales = {}
    for row in counts:
        if row["days"] > 1:
            base = scale_recipe(fetch_recipe(row["recipe_id"]), {})["scale_base"]
            if base:
                qty = ingredient_quantity(base) * row["days"]
                scales[row["recipe_id"]] = (str(base["id"]), format(qty, "f"))
    with get_db() as cur:
        for row in counts:
            cur.execute(
                "INSERT OR IGNORE INTO shopping_recipes (recipe_id) VALUES (?)", (row["recipe_id"],)
            )
            if row["recipe_id"] in scales:
                cur.execute(
                    "UPDATE shopping_recipes SET scale_by = ?, scale_qty = ? WHERE recipe_id = ?",
                    (*scales[row["recipe_id"]], row["recipe_id"]),
                )
    flash("Блюда из меню добавлены к выбранным рецептам.")
    return redirect(url_for("shopping"))


def shopping_back():
    """Куда вернуться после изменения списка: на рецепт или на страницу покупок с тем же поиском."""
    recipe_id = request.form.get("back_recipe", type=int)
    if recipe_id:
        return redirect(url_for("recipe_view", recipe_id=recipe_id))
    return redirect(url_for("shopping", q=request.form.get("q") or None))


@app.route("/shopping/add", methods=["POST"])
def shopping_add():
    with get_db() as cur:
        # SELECT отсекает рецепт, который уже успели удалить.
        cur.execute(
            "INSERT OR IGNORE INTO shopping_recipes (recipe_id) SELECT id FROM recipes WHERE id = ?",
            (request.form.get("recipe_id", type=int),),
        )
    return shopping_back()


@app.route("/shopping/remove", methods=["POST"])
def shopping_remove():
    with get_db() as cur:
        cur.execute(
            "DELETE FROM shopping_recipes WHERE recipe_id = ?",
            (request.form.get("recipe_id", type=int),),
        )
    return shopping_back()


@app.route("/shopping/scale", methods=["POST"])
def shopping_scale():
    recipe_id = request.form.get("recipe_id", type=int)
    if "reset" in request.form:
        by, qty = None, None
    else:
        by, qty = request.form.get("by"), request.form.get("qty", "").strip().replace(",", ".")
        if parse_quantity(qty) is None:
            flash("Количество для пересчёта должно быть числом от 0,01.", "error")
            return shopping_back()
    with get_db() as cur:
        cur.execute(
            "UPDATE shopping_recipes SET scale_by = ?, scale_qty = ? WHERE recipe_id = ?",
            (by, qty, recipe_id),
        )
    return shopping_back()


@app.route("/shopping/clear", methods=["POST"])
def shopping_clear():
    with get_db() as cur:
        cur.execute("DELETE FROM shopping_recipes")
        cur.execute("DELETE FROM shopping_checks")
        cur.execute("DELETE FROM shopping_custom")
    return shopping_back()


@app.route("/shopping/docx")
def shopping_docx():
    data = shopping_groups()
    if not data["groups"] and not data["unsorted"]:
        return redirect(url_for("shopping_list"))
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(12)
    doc.add_heading("Список покупок", level=0)
    if data["recipes"]:
        p = doc.add_paragraph()
        p.add_run("Для рецептов: ").italic = True
        p.add_run(", ".join(r["title"] for r in data["recipes"]))
    groups = list(data["groups"])
    if data["unsorted"]:
        groups.append(("Другое", data["unsorted"]))
    for category, items in groups:
        doc.add_heading(category, level=1)
        for item in items:
            amount = f" — {item['amount']}" if item["amount"] else ""
            doc.add_paragraph(f"☐  {item['name']}{amount}")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="Список покупок.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/recipe/<int:recipe_id>/docx")
def recipe_docx(recipe_id):
    recipe = scale_recipe(fetch_recipe(recipe_id), request.args)
    doc = build_docx(recipe)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{safe_filename(recipe['title'])}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


def safe_filename(title):
    return re.sub(r'[\\/:*?"<>|]+', "_", title).strip() or "recipe"


def export_photo(recipe):
    path = PHOTOS_DIR / recipe["photo"] if recipe["photo"] else None
    if path is None or not path.exists():
        return None
    return base64.b64encode(path.read_bytes()).decode("ascii")


def import_photo(data):
    """Фото из файла импорта. Повреждённое фото не мешает добавить сам рецепт."""
    try:
        photo_data = base64.b64decode(data.get("photo") or "", validate=True)
    except (TypeError, ValueError):
        return None
    if photo_extension(photo_data) and len(photo_data) <= PHOTO_MAX_BYTES:
        return photo_data
    return None


def recipe_to_export(recipe):
    with get_db() as cur:
        export_categories = stored_categories(cur)
    return {
        "title": recipe["title"],
        "author": recipe["author"],
        "prep_minutes": recipe["prep_minutes"],
        "servings": recipe["servings"],
        "source_url": recipe["source_url"],
        "tags": recipe["tags"],
        "photo": export_photo(recipe),
        "ingredients": [
            {
                "name": i["name"],
                "quantity": format(Decimal(i["quantity"]).normalize(), "f"),
                "quantity_max": (
                    format(Decimal(i["quantity_max"]).normalize(), "f")
                    if i.get("quantity_max") is not None
                    else None
                ),
                "unit": i["unit"],
                "is_main": bool(i["is_main"]),
                "category": export_categories.get(title_key(i["name"])) or guess_category(i["name"]),
            }
            for i in recipe["ingredients"]
        ],
        "steps": [s["text"] for s in recipe["steps"]],
    }


def export_response(recipes, filename):
    payload = {
        "format": EXPORT_FORMAT,
        "version": EXPORT_VERSION,
        "recipes": [recipe_to_export(r) for r in recipes],
    }
    buf = io.BytesIO(json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"))
    return send_file(
        buf, as_attachment=True, download_name=filename, mimetype="application/json"
    )


def import_to_form(data):
    """Переводит рецепт из файла в вид формы, чтобы проверить его тем же parse_form."""
    if not isinstance(data, dict):
        raise ValueError
    form = MultiDict()
    form.add("title", str(data.get("title") or ""))
    form.add("author", str(data.get("author") or ""))
    form.add("source_url", str(data.get("source_url") or ""))
    form.add("prep_minutes", str(data.get("prep_minutes") or ""))
    form.add("servings", str(data.get("servings") or ""))
    form.add("tags", ", ".join(str(t) for t in data.get("tags") or []))
    for n, ing in enumerate(data.get("ingredients") or []):
        form.add("ing_key", str(n))
        form.add("ing_name", str(ing.get("name") or ""))
        quantity = str(ing.get("quantity") or "")
        if ing.get("quantity_max"):
            quantity += f"-{ing['quantity_max']}"
        form.add("ing_qty", quantity)
        form.add("ing_unit", str(ing.get("unit") or ""))
        form.add("ing_category", str(ing.get("category") or ""))
        if ing.get("is_main"):
            form.add("main", str(n))
    for step in data.get("steps") or []:
        form.add("step", str(step))
    return form


@app.route("/recipe/<int:recipe_id>/export")
def recipe_export(recipe_id):
    recipe = fetch_recipe(recipe_id)
    return export_response([recipe], f"{safe_filename(recipe['title'])}.json")


@app.route("/export")
def export_recipes():
    """Экспорт рецептов из ?id=&id=… Без id выгружается вся книга."""
    wanted = set(request.args.getlist("id", type=int))
    with get_db() as cur:
        cur.execute("SELECT id FROM recipes ORDER BY title_key")
        ids = [row["id"] for row in cur.fetchall()]
    if wanted:
        ids = [i for i in ids if i in wanted]
    if not ids:
        flash("Нет рецептов для экспорта.", "error")
        return redirect(url_for("import_recipes"))
    recipes = [fetch_recipe(i) for i in ids]
    name = recipes[0]["title"] if len(recipes) == 1 else "рецепты"
    return export_response(recipes, f"{safe_filename(name)}.json")


def parse_import_entry(data):
    try:
        return parse_form(import_to_form(data))
    except (AttributeError, TypeError, ValueError):
        return {"title": ""}, ["неверный формат"]


def find_by_title(cur, title):
    cur.execute("SELECT id, title, author, photo FROM recipes WHERE title_key = ?", (title_key(title),))
    return cur.fetchone()


def find_conflicts(cur, entries):
    """Рецепты файла, название которых уже есть в книге или раньше в этом же файле."""
    conflicts = []
    seen = set()
    for n, data in enumerate(entries, 1):
        recipe, errors = parse_import_entry(data)
        if errors:
            continue
        key = title_key(recipe["title"])
        existing = find_by_title(cur, recipe["title"])
        if existing or key in seen:
            conflicts.append(
                {"index": n, "title": recipe["title"], "author": recipe["author"], "existing": existing}
            )
        seen.add(key)
    return conflicts


def free_copy_title(cur, title):
    n = 2
    while title_exists(cur, f"{title} ({n})"):
        n += 1
    return f"{title} ({n})"


def apply_import(entries, choices):
    """Добавляет рецепты файла. choices: номер рецепта → skip, replace или copy."""
    result = {"added": [], "replaced": [], "copied": [], "skipped": [], "invalid": []}
    old_photos = []
    with get_db() as cur:
        for n, data in enumerate(entries, 1):
            recipe, errors = parse_import_entry(data)
            name = recipe["title"] or f"рецепт №{n}"
            if errors:
                result["invalid"].append((name, errors))
                continue
            existing = find_by_title(cur, recipe["title"])
            choice = choices.get(n, "skip") if existing else "add"
            if choice not in ("add", "replace", "copy"):
                result["skipped"].append(name)
                continue
            upload = import_photo(data)
            if upload:
                recipe["photo"] = store_photo(upload)
            if choice == "replace":
                # Рецепт из файла без фото не должен стирать уже имеющееся.
                if upload:
                    old_photos.append(existing["photo"])
                else:
                    recipe["photo"] = existing["photo"]
                update_recipe(cur, existing["id"], recipe)
                result["replaced"].append(name)
            elif choice == "copy":
                recipe["title"] = free_copy_title(cur, recipe["title"])
                insert_recipe(cur, recipe)
                result["copied"].append(recipe["title"])
            else:
                insert_recipe(cur, recipe)
                result["added"].append(name)
    for photo_name in old_photos:
        delete_photo(photo_name)
    return result


def render_transfer(**context):
    """Общая страница импорта и экспорта."""
    with get_db() as cur:
        recipes = search_recipes(cur)
    context.setdefault("result", None)
    context.setdefault("conflicts", None)
    return render_template("import.html", recipes=recipes, **context)


def import_file_contents(raw):
    """Разбирает файл импорта: сразу добавляет рецепты или спрашивает про дубликаты."""
    try:
        payload = json.loads(raw)
        if payload.get("format") != EXPORT_FORMAT or not isinstance(payload.get("recipes"), list):
            raise ValueError
    except (AttributeError, ValueError, UnicodeDecodeError):
        flash("Это не файл рецептов. Выберите файл .json, выгруженный из книги рецептов.", "error")
        return render_transfer()

    with get_db() as cur:
        conflicts = find_conflicts(cur, payload["recipes"])
    if not conflicts:
        return render_transfer(result=apply_import(payload["recipes"], {}))

    # Есть дубликаты: ничего не добавляем, пока пользователь не решит, что с ними делать.
    for old in IMPORTS_DIR.glob("*.json"):
        old.unlink(missing_ok=True)
    token = uuid.uuid4().hex
    (IMPORTS_DIR / f"{token}.json").write_bytes(raw)
    return render_transfer(
        conflicts=conflicts,
        token=token,
        new_count=sum(not parse_import_entry(d)[1] for d in payload["recipes"]) - len(conflicts),
    )


@app.route("/import", methods=["GET", "POST"])
def import_recipes():
    if request.method == "GET":
        return render_transfer()
    file = request.files.get("file")
    return import_file_contents(file.read() if file else b"")


@app.route("/import/incoming")
def import_incoming():
    """Файл, открытый в книге из другого приложения Android. Его кладёт сюда MainActivity."""
    path = IMPORTS_DIR / "incoming.json"
    if not app.config.get("ANDROID") or not path.exists():
        return redirect(url_for("import_recipes"))
    raw = path.read_bytes()
    path.unlink()
    return import_file_contents(raw)


@app.route("/import/apply", methods=["POST"])
def import_apply():
    token = request.form.get("token", "")
    path = IMPORTS_DIR / f"{token}.json"
    if not re.fullmatch(r"[0-9a-f]{32}", token) or not path.exists():
        flash("Файл импорта устарел. Загрузите его ещё раз.", "error")
        return redirect(url_for("import_recipes"))
    entries = json.loads(path.read_bytes())["recipes"]
    choices = {n: request.form.get(f"choice_{n}", "skip") for n in range(1, len(entries) + 1)}
    result = apply_import(entries, choices)
    path.unlink(missing_ok=True)
    return render_transfer(result=result)


@app.route("/unlock")
def unlock():
    """Первый запрос приложения Android: обменивает токен из адреса на cookie.

    Cookie приходит в самом ответе сервера, поэтому следующая страница её точно получит.
    """
    token = app.config.get("ACCESS_TOKEN")
    if not token or request.args.get("token") != token:
        abort(403)
    response = local_redirect(request.args.get("next"), url_for("index"))
    response.set_cookie("access", token, httponly=True, samesite="Lax")
    return response


@app.errorhandler(403)
def forbidden(_error):
    return "Доступ закрыт: книгу можно открыть только из самого приложения.", 403


@app.route("/ping")
def ping():
    # Запускалка так узнаёт, что книга уже открыта.
    return "recept"


@app.route("/quit", methods=["POST"])
def quit_app():
    shutdown = app.config.get("SHUTDOWN")
    if shutdown is None:
        abort(404)
    shutdown()
    return render_template("bye.html")


@app.errorhandler(413)
def too_large(_error):
    flash("Файл слишком большой.", "error")
    return redirect(request.path)


def build_docx(recipe):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(12)

    title = doc.add_heading(recipe["title"], level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if recipe["photo"]:
        try:
            doc.add_picture(str(PHOTOS_DIR / recipe["photo"]), width=Cm(10))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except (OSError, UnrecognizedImageError):
            pass  # без фото документ всё равно полезен

    if recipe["author"]:
        p = doc.add_paragraph()
        p.add_run("Автор: ").italic = True
        p.add_run(recipe["author"])

    if recipe["main"]:
        p = doc.add_paragraph()
        p.add_run("Главный ингредиент: ").italic = True
        p.add_run(recipe["main"]["name"]).bold = True

    meta = []
    if recipe["servings"]:
        meta.append(f"Порций: {format_number(recipe['servings'])}")
    if recipe["prep_minutes"]:
        meta.append(f"Время: {format_minutes(recipe['prep_minutes'])}")
    if meta:
        doc.add_paragraph(" · ".join(meta))

    doc.add_heading("Ингредиенты", level=1)
    table = doc.add_table(rows=0, cols=2)
    table.style = "Light List Accent 1"
    for ing in recipe["ingredients"]:
        cells = table.add_row().cells
        name_run = cells[0].paragraphs[0].add_run(ing["name"])
        qty_run = cells[1].paragraphs[0].add_run(
            format_quantity(ing["quantity"], ing["unit"], ing.get("quantity_max"))
        )
        if ing["is_main"]:
            name_run.bold = qty_run.bold = True
        cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_heading("Приготовление", level=1)
    for step in recipe["steps"]:
        doc.add_paragraph(step["text"], style="List Number")

    return doc


if __name__ == "__main__":
    # Режим разработки. Обычный запуск — через launcher.pyw.
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5000)), debug=True)
