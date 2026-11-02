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
    unit        TEXT NOT NULL
                CHECK (unit IN ('l','ml','tbsp','tsp','kg','g','pcs','cloves','to_taste')),
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
    added_at    TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
