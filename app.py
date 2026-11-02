"""Локальная книга рецептов: Flask + SQLite, экспорт рецепта в DOCX."""
import base64
import io
import json
import os
import re
import sqlite3
import uuid
from contextlib import contextmanager
from decimal import Decimal, InvalidOperation
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.exceptions import UnrecognizedImageError
from docx.shared import Cm, Pt
from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    url_for,
)
from werkzeug.datastructures import MultiDict

BASE_DIR = Path(__file__).resolve().parent
# База лежит рядом с программой; RECEPT_DATA позволяет хранить её в другом месте.
DATA_DIR = Path(os.environ.get("RECEPT_DATA", BASE_DIR / "data"))
DB_PATH = DATA_DIR / "recipes.db"
PHOTOS_DIR = DATA_DIR / "photos"
IMPORTS_DIR = DATA_DIR / "imports"  # файлы импорта, ожидающие решения по дубликатам

# Только форматы, которые можно вставить в DOCX без дополнительных библиотек.
PHOTO_SIGNATURES = {b"\xff\xd8\xff": "jpg", b"\x89PNG\r\n\x1a\n": "png"}
PHOTO_MAX_BYTES = 10 * 1024 * 1024

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


def init_db():
    for folder in (DATA_DIR, PHOTOS_DIR, IMPORTS_DIR):
        folder.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript((BASE_DIR / "schema.sql").read_text(encoding="utf-8"))
        # Базы старых версий: CREATE TABLE IF NOT EXISTS не добавляет новые столбцы.
        columns = {row[1] for row in conn.execute("PRAGMA table_info(recipes)")}
        for column, kind in (("photo", "TEXT"), ("prep_minutes", "INTEGER"), ("servings", "INTEGER")):
            if column not in columns:
                conn.execute(f"ALTER TABLE recipes ADD COLUMN {column} {kind}")


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


def format_quantity(quantity, unit):
    if unit == "to_taste":
        return UNITS[unit]
    return f"{format_number(quantity)} {UNITS[unit]}"


def format_minutes(minutes):
    hours, rest = divmod(minutes, 60)
    parts = [f"{hours} ч" if hours else "", f"{rest} мин" if rest else ""]
    return " ".join(p for p in parts if p)


def photo_url(recipe):
    if recipe["photo"]:
        return url_for("photo", name=recipe["photo"])
    return url_for("static", filename="placeholder.svg")


app.jinja_env.filters["qty"] = lambda ing: format_quantity(ing["quantity"], ing["unit"])
app.jinja_env.filters["qty_input"] = lambda ing: (
    "" if ing["unit"] == "to_taste"
    else format(Decimal(ing["quantity"]).quantize(Decimal("0.01")).normalize(), "f")
)
app.jinja_env.filters["number"] = format_number
app.jinja_env.filters["minutes"] = format_minutes
app.jinja_env.filters["photo_url"] = photo_url


def parse_quantity(value):
    """Количество из строки ввода, либо None, если это не число из допустимого диапазона."""
    try:
        q = Decimal((value or "").strip().replace(",", "."))
    except InvalidOperation:
        return None
    # is_finite отсекает NaN и Infinity, которые Decimal тоже считает числами.
    return q if q.is_finite() and QTY_MIN <= q <= QTY_MAX else None


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
    if qty is None or base is None or not Decimal(base["quantity"]):
        return recipe
    factor = qty / Decimal(base["quantity"])
    for ing in recipe["ingredients"]:
        ing["quantity"] = Decimal(ing["quantity"]) * factor
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
    for key, name, qty, unit in rows:
        name = name.strip()
        if not name:
            continue
        if unit not in UNITS:
            errors.append(f"Неизвестная единица измерения у «{name}».")
            continue
        if unit == "to_taste":
            quantity = Decimal(0)
        else:
            quantity = parse_quantity(qty)
            if quantity is None:
                errors.append(f"Количество для «{name}» должно быть числом от 0,01.")
                continue
        ingredients.append(
            {"name": name, "quantity": quantity, "unit": unit, "is_main": key == main_key}
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
        "prep_minutes": prep_minutes,
        "servings": servings,
        "tags": parse_tags(form.get("tags")),
        "ingredients": ingredients,
        "steps": steps,
    }
    return recipe, errors


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
        "INSERT INTO recipes (title, title_key, author, photo, prep_minutes, servings)"
        " VALUES (?, ?, ?, ?, ?, ?)",
        (
            recipe["title"],
            title_key(recipe["title"]),
            recipe["author"],
            recipe["photo"],
            recipe["prep_minutes"],
            recipe["servings"],
        ),
    )
    recipe_id = cur.lastrowid
    save_children(cur, recipe_id, recipe)
    return recipe_id


def update_recipe(cur, recipe_id, recipe):
    cur.execute(
        "UPDATE recipes SET title = ?, title_key = ?, author = ?, photo = ?, prep_minutes = ?,"
        " servings = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (
            recipe["title"],
            title_key(recipe["title"]),
            recipe["author"],
            recipe["photo"],
            recipe["prep_minutes"],
            recipe["servings"],
            recipe_id,
        ),
    )
    save_children(cur, recipe_id, recipe)


def save_children(cur, recipe_id, recipe):
    cur.execute("DELETE FROM ingredients WHERE recipe_id = ?", (recipe_id,))
    cur.execute("DELETE FROM steps WHERE recipe_id = ?", (recipe_id,))
    cur.executemany(
        "INSERT INTO ingredients (recipe_id, position, name, quantity, unit, is_main)"
        " VALUES (?, ?, ?, ?, ?, ?)",
        [
            (recipe_id, pos, i["name"], format(i["quantity"], "f"), i["unit"], i["is_main"])
            for pos, i in enumerate(recipe["ingredients"], 1)
        ],
    )
    cur.executemany(
        "INSERT INTO steps (recipe_id, position, text) VALUES (?, ?, ?)",
        [(recipe_id, pos, s["text"]) for pos, s in enumerate(recipe["steps"], 1)],
    )
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
    return render_template(
        "form.html",
        recipe=recipe,
        units=UNITS,
        mode=mode,
        all_tags=tags,
        ingredient_names=[names[k] for k in sorted(names)],
    )


def search_recipes(cur, q="", tag=""):
    sql = (
        "SELECT r.id, r.title, r.author, r.photo, r.prep_minutes, r.servings,"
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
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY r.title_key"
    cur.execute(sql, params)
    return cur.fetchall()


@app.context_processor
def shopping_count():
    with get_db() as cur:
        cur.execute("SELECT COUNT(*) AS n FROM shopping_recipes")
        return {"shopping_count": cur.fetchone()["n"]}


@app.route("/")
def index():
    q = request.args.get("q", "").strip()
    tag = request.args.get("tag", "").strip()
    with get_db() as cur:
        recipes = search_recipes(cur, q, tag)
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
    return render_template("index.html", recipes=recipes, q=q, tag=tag, all_tags=tags)


@app.route("/recipe/new", methods=["GET", "POST"])
def recipe_new():
    if request.method == "GET":
        return render_form(None, "new")
    recipe, errors = parse_form(request.form)
    upload = read_photo_upload(request.files.get("photo"), errors)
    if not errors:
        with get_db() as cur:
            if title_exists(cur, recipe["title"]):
                errors.append(f"Рецепт «{recipe['title']}» уже есть в книге.")
            else:
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
    upload = read_photo_upload(request.files.get("photo"), errors)
    recipe["photo"] = existing["photo"]
    if not errors:
        with get_db() as cur:
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
    with get_db() as cur:
        cur.execute("SELECT recipe_id FROM shopping_recipes ORDER BY added_at, recipe_id")
        ids = [row["recipe_id"] for row in cur.fetchall()]
    return [fetch_recipe(i) for i in ids]


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
                group, multiplier = SHOPPING_UNITS[ing["unit"]]
                amount = Decimal(ing["quantity"]) * multiplier
                item["amounts"][group] = item["amounts"].get(group, Decimal(0)) + amount
    return [
        {"name": items[key]["name"], "amount": format_shopping_amount(items[key])}
        for key in sorted(items)
    ]


def format_shopping_amount(item):
    parts = []
    for group, amount in item["amounts"].items():
        if group == "mass":
            parts.append(
                format_quantity(amount / 1000, "kg") if amount >= 1000 else format_quantity(amount, "g")
            )
        elif group == "volume":
            parts.append(
                format_quantity(amount / 1000, "l") if amount >= 1000 else format_quantity(amount, "ml")
            )
        else:
            parts.append(format_quantity(amount, group))
    if item["to_taste"]:
        parts.append(UNITS["to_taste"])
    return " + ".join(parts)


@app.route("/shopping")
def shopping():
    q = request.args.get("q", "").strip()
    selected = shopping_recipes()
    chosen_ids = {r["id"] for r in selected}
    with get_db() as cur:
        found = [r for r in search_recipes(cur, q) if r["id"] not in chosen_ids]
    return render_template(
        "shopping.html", selected=selected, items=build_shopping_list(selected), found=found, q=q
    )


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


@app.route("/shopping/clear", methods=["POST"])
def shopping_clear():
    with get_db() as cur:
        cur.execute("DELETE FROM shopping_recipes")
    return shopping_back()


@app.route("/shopping/docx")
def shopping_docx():
    recipes = shopping_recipes()
    if not recipes:
        return redirect(url_for("shopping"))
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(12)
    doc.add_heading("Список покупок", level=0)
    p = doc.add_paragraph()
    p.add_run("Для рецептов: ").italic = True
    p.add_run(", ".join(r["title"] for r in recipes))
    for item in build_shopping_list(recipes):
        doc.add_paragraph(f"☐  {item['name']} — {item['amount']}")
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
    return {
        "title": recipe["title"],
        "author": recipe["author"],
        "prep_minutes": recipe["prep_minutes"],
        "servings": recipe["servings"],
        "tags": recipe["tags"],
        "photo": export_photo(recipe),
        "ingredients": [
            {
                "name": i["name"],
                "quantity": format(Decimal(i["quantity"]).normalize(), "f"),
                "unit": i["unit"],
                "is_main": bool(i["is_main"]),
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
    form.add("prep_minutes", str(data.get("prep_minutes") or ""))
    form.add("servings", str(data.get("servings") or ""))
    form.add("tags", ", ".join(str(t) for t in data.get("tags") or []))
    for n, ing in enumerate(data.get("ingredients") or []):
        form.add("ing_key", str(n))
        form.add("ing_name", str(ing.get("name") or ""))
        form.add("ing_qty", str(ing.get("quantity") or ""))
        form.add("ing_unit", str(ing.get("unit") or ""))
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
def export_all():
    with get_db() as cur:
        cur.execute("SELECT id FROM recipes ORDER BY title_key")
        ids = [row["id"] for row in cur.fetchall()]
    return export_response([fetch_recipe(i) for i in ids], "рецепты.json")


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


@app.route("/import", methods=["GET", "POST"])
def import_recipes():
    if request.method == "GET":
        return render_template("import.html", result=None, conflicts=None)

    file = request.files.get("file")
    try:
        raw = file.read()
        payload = json.loads(raw)
        if payload.get("format") != EXPORT_FORMAT or not isinstance(payload.get("recipes"), list):
            raise ValueError
    except (AttributeError, ValueError, UnicodeDecodeError):
        flash("Это не файл рецептов. Выберите файл .json, выгруженный из книги рецептов.", "error")
        return render_template("import.html", result=None, conflicts=None)

    with get_db() as cur:
        conflicts = find_conflicts(cur, payload["recipes"])
    if not conflicts:
        return render_template(
            "import.html", result=apply_import(payload["recipes"], {}), conflicts=None
        )

    # Есть дубликаты: ничего не добавляем, пока пользователь не решит, что с ними делать.
    for old in IMPORTS_DIR.glob("*.json"):
        old.unlink(missing_ok=True)
    token = uuid.uuid4().hex
    (IMPORTS_DIR / f"{token}.json").write_bytes(raw)
    return render_template(
        "import.html",
        result=None,
        conflicts=conflicts,
        token=token,
        new_count=sum(not parse_import_entry(d)[1] for d in payload["recipes"]) - len(conflicts),
    )


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
    return render_template("import.html", result=result, conflicts=None)


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
            format_quantity(ing["quantity"], ing["unit"])
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
