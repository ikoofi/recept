-- Таблицы книги рецептов (SQLite). Выполняется приложением при каждом запуске.

CREATE TABLE IF NOT EXISTS recipes (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    title       TEXT NOT NULL,
    -- Название без учёта регистра и лишних пробелов: «Борщ» и «борщ » — один рецепт.
    title_key   TEXT NOT NULL UNIQUE,
    author      TEXT,
    photo       TEXT,            -- имя файла в data/photos
    prep_minutes INTEGER,
    servings    INTEGER,
    favorite    INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at  TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- quantity хранится строкой, чтобы не терять точность дробей. Для to_taste quantity = '0'.
CREATE TABLE IF NOT EXISTS ingredients (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    recipe_id   INTEGER NOT NULL REFERENCES recipes(id) ON DELETE CASCADE,
    position    INTEGER NOT NULL,
    name        TEXT NOT NULL,
    quantity    TEXT NOT NULL DEFAULT '0',
    unit        TEXT NOT NULL,   -- код из UNITS или своя единица текстом, например «стакан»
    is_main     INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_ingredients_recipe ON ingredients (recipe_id, position);

CREATE TABLE IF NOT EXISTS steps (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    recipe_id   INTEGER NOT NULL REFERENCES recipes(id) ON DELETE CASCADE,
    position    INTEGER NOT NULL,
    text        TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_steps_recipe ON steps (recipe_id, position);

-- Категории: свободные метки, у рецепта их может быть несколько.
CREATE TABLE IF NOT EXISTS tags (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    name_key    TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS recipe_tags (
    recipe_id   INTEGER NOT NULL REFERENCES recipes(id) ON DELETE CASCADE,
    tag_id      INTEGER NOT NULL REFERENCES tags(id) ON DELETE CASCADE,
    PRIMARY KEY (recipe_id, tag_id)
);

-- Рецепты, отобранные для списка покупок. Хранятся в базе, чтобы выбор не терялся.
CREATE TABLE IF NOT EXISTS shopping_recipes (
    recipe_id   INTEGER PRIMARY KEY REFERENCES recipes(id) ON DELETE CASCADE,
    scale_by    TEXT,            -- id ингредиента или servings, см. scale_recipe
    scale_qty   TEXT,
    added_at    TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- Купленные продукты в списке покупок. Ключ — название без учёта регистра:
-- один продукт отмечается сразу и в блюдах, и в итоге.
CREATE TABLE IF NOT EXISTS shopping_checks (
    item_key    TEXT PRIMARY KEY
);

-- Меню на неделю. day: 0 — понедельник, 6 — воскресенье.
CREATE TABLE IF NOT EXISTS menu_items (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    day         INTEGER NOT NULL CHECK (day BETWEEN 0 AND 6),
    recipe_id   INTEGER NOT NULL REFERENCES recipes(id) ON DELETE CASCADE,
    UNIQUE (day, recipe_id)
);

-- Отдел магазина для продукта, выбранный пользователем. Ключ — название без учёта регистра.
CREATE TABLE IF NOT EXISTS ingredient_categories (
    name_key    TEXT PRIMARY KEY,
    category    TEXT NOT NULL
);

-- Продукты, которые обычно есть дома: в список покупок не попадают.
CREATE TABLE IF NOT EXISTS pantry (
    name_key    TEXT PRIMARY KEY
);

-- Свои пункты списка покупок, не из рецептов.
CREATE TABLE IF NOT EXISTS shopping_custom (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    amount      TEXT NOT NULL DEFAULT ''
);
