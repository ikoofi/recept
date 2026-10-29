-- Таблицы книги рецептов (SQLite). Выполняется приложением при каждом запуске.

CREATE TABLE IF NOT EXISTS recipes (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    title       TEXT NOT NULL,
    -- Название без учёта регистра и лишних пробелов: «Борщ» и «борщ » — один рецепт.
    title_key   TEXT NOT NULL UNIQUE,
    author      TEXT,
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
