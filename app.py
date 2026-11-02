"""Локальная книга рецептов: Flask + SQLite, экспорт рецепта в DOCX."""
import io
import json
import os
import re
import sqlite3
from contextlib import contextmanager
from decimal import Decimal, InvalidOperation
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for
from werkzeug.datastructures import MultiDict

BASE_DIR = Path(__file__).resolve().parent
# База лежит рядом с программой; RECEPT_DATA позволяет хранить её в другом месте.
DATA_DIR = Path(os.environ.get("RECEPT_DATA", BASE_DIR / "data"))
DB_PATH = DATA_DIR / "recipes.db"

# Границы количества: мельче сотой значение округлится до нуля при показе.
QTY_MIN = Decimal("0.01")
QTY_MAX = Decimal("1000000000")

EXPORT_FORMAT = "recept"
EXPORT_VERSION = 1

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


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "local-recipe-book")


def title_key(title):
    """Ключ для сравнения названий: SQLite сам не умеет без учёта регистра в кириллице."""
    return " ".join(title.split()).casefold()


def fold(value):
    return value.casefold() if isinstance(value, str) else value


def init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript((BASE_DIR / "schema.sql").read_text(encoding="utf-8"))


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


def format_quantity(quantity, unit):
    if unit == "to_taste":
        return UNITS[unit]
    # Округление до сотых нужно для пересчитанных значений (например, 1/3 порции).
    q = Decimal(quantity).quantize(Decimal("0.01")).normalize()
    text = format(q, "f").replace(".", ",")
    return f"{text} {UNITS[unit]}"


app.jinja_env.filters["qty"] = lambda ing: format_quantity(ing["quantity"], ing["unit"])
app.jinja_env.filters["qty_input"] = lambda ing: (
    "" if ing["unit"] == "to_taste"
    else format(Decimal(ing["quantity"]).quantize(Decimal("0.01")).normalize(), "f")
)


def parse_quantity(value):
    """Количество из строки ввода, либо None, если это не число из допустимого диапазона."""
    try:
        q = Decimal((value or "").strip().replace(",", "."))
    except InvalidOperation:
        return None
    # is_finite отсекает NaN и Infinity, которые Decimal тоже считает числами.
    return q if q.is_finite() and QTY_MIN <= q <= QTY_MAX else None


def scale_recipe(recipe, args):
    """Пересчитывает рецепт из ?by=<id ингредиента>&qty=<количество>.

    В recipe["scale_base"] кладёт ингредиент, по которому идёт расчёт: выбранный,
    иначе главный, иначе первый с количеством. recipe["scaled"] — был ли пересчёт.
    """
    options = [i for i in recipe["ingredients"] if i["unit"] != "to_taste"]
    by = args.get("by", type=int)
    base = next((i for i in options if i["id"] == by), None)
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
    recipe["main"] = next((i for i in recipe["ingredients"] if i["is_main"]), None)
    return recipe


def parse_form(form):
    """Собирает рецепт из формы. Возвращает (recipe, errors)."""
    errors = []
    title = form.get("title", "").strip()
    if not title:
        errors.append("Укажите название рецепта.")
    author = form.get("author", "").strip() or None

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

    return {"title": title, "author": author, "ingredients": ingredients, "steps": steps}, errors


def title_exists(cur, title, exclude_id=None):
    cur.execute(
        "SELECT 1 FROM recipes WHERE title_key = ? AND id <> ? LIMIT 1",
        (title_key(title), exclude_id or 0),
    )
    return cur.fetchone() is not None


def insert_recipe(cur, recipe):
    cur.execute(
        "INSERT INTO recipes (title, title_key, author) VALUES (?, ?, ?)",
        (recipe["title"], title_key(recipe["title"]), recipe["author"]),
    )
    recipe_id = cur.lastrowid
    save_children(cur, recipe_id, recipe)
    return recipe_id


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


@app.route("/")
def index():
    q = request.args.get("q", "").strip()
    sql = (
        "SELECT r.id, r.title, r.author, i.name AS main_ingredient"
        " FROM recipes r LEFT JOIN ingredients i ON i.recipe_id = r.id AND i.is_main"
    )
    params = ()
    if q:
        sql += " WHERE fold(r.title) LIKE ? OR fold(r.author) LIKE ? OR fold(i.name) LIKE ?"
        params = (f"%{q.casefold()}%",) * 3
    sql += " ORDER BY r.title_key"
    with get_db() as cur:
        cur.execute(sql, params)
        recipes = cur.fetchall()
    return render_template("index.html", recipes=recipes, q=q)


@app.route("/recipe/new", methods=["GET", "POST"])
def recipe_new():
    if request.method == "POST":
        recipe, errors = parse_form(request.form)
        if not errors:
            with get_db() as cur:
                if title_exists(cur, recipe["title"]):
                    errors.append(f"Рецепт «{recipe['title']}» уже есть в книге.")
                else:
                    recipe_id = insert_recipe(cur, recipe)
        if not errors:
            flash("Рецепт сохранён.")
            return redirect(url_for("recipe_view", recipe_id=recipe_id))
        for e in errors:
            flash(e, "error")
        return render_template("form.html", recipe=recipe, units=UNITS, mode="new")
    return render_template("form.html", recipe=None, units=UNITS, mode="new")


@app.route("/recipe/<int:recipe_id>")
def recipe_view(recipe_id):
    recipe = scale_recipe(fetch_recipe(recipe_id), request.args)
    return render_template("view.html", recipe=recipe, units=UNITS)


@app.route("/recipe/<int:recipe_id>/edit", methods=["GET", "POST"])
def recipe_edit(recipe_id):
    existing = fetch_recipe(recipe_id)
    if request.method == "POST":
        recipe, errors = parse_form(request.form)
        if not errors:
            with get_db() as cur:
                if title_exists(cur, recipe["title"], exclude_id=recipe_id):
                    errors.append(f"Рецепт «{recipe['title']}» уже есть в книге.")
                else:
                    cur.execute(
                        "UPDATE recipes SET title = ?, title_key = ?, author = ?,"
                        " updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                        (recipe["title"], title_key(recipe["title"]), recipe["author"], recipe_id),
                    )
                    save_children(cur, recipe_id, recipe)
        if not errors:
            flash("Изменения сохранены.")
            return redirect(url_for("recipe_view", recipe_id=recipe_id))
        for e in errors:
            flash(e, "error")
        recipe["id"] = recipe_id
        return render_template("form.html", recipe=recipe, units=UNITS, mode="edit")
    return render_template("form.html", recipe=existing, units=UNITS, mode="edit")


@app.route("/recipe/<int:recipe_id>/delete", methods=["POST"])
def recipe_delete(recipe_id):
    with get_db() as cur:
        cur.execute("DELETE FROM recipes WHERE id = ?", (recipe_id,))
    flash("Рецепт удалён.")
    return redirect(url_for("index"))


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


def recipe_to_export(recipe):
    return {
        "title": recipe["title"],
        "author": recipe["author"],
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


@app.route("/import", methods=["GET", "POST"])
def import_recipes():
    if request.method == "GET":
        return render_template("import.html", result=None)

    file = request.files.get("file")
    try:
        payload = json.load(file.stream)
        if payload.get("format") != EXPORT_FORMAT or not isinstance(payload.get("recipes"), list):
            raise ValueError
    except (AttributeError, ValueError, UnicodeDecodeError):
        flash("Это не файл рецептов. Выберите файл .json, выгруженный из книги рецептов.", "error")
        return render_template("import.html", result=None)

    result = {"added": [], "skipped": [], "invalid": []}
    with get_db() as cur:
        for n, data in enumerate(payload["recipes"], 1):
            try:
                recipe, errors = parse_form(import_to_form(data))
            except (AttributeError, ValueError):
                recipe, errors = {"title": ""}, ["неверный формат"]
            name = recipe["title"] or f"рецепт №{n}"
            if errors:
                result["invalid"].append((name, errors))
            elif title_exists(cur, recipe["title"]):
                result["skipped"].append(name)
            else:
                insert_recipe(cur, recipe)
                result["added"].append(name)
    return render_template("import.html", result=result)


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


def build_docx(recipe):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(12)

    title = doc.add_heading(recipe["title"], level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if recipe["author"]:
        p = doc.add_paragraph()
        p.add_run("Автор: ").italic = True
        p.add_run(recipe["author"])

    if recipe["main"]:
        p = doc.add_paragraph()
        p.add_run("Главный ингредиент: ").italic = True
        p.add_run(recipe["main"]["name"]).bold = True

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
