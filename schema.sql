-- Таблицы книги рецептов. Загружается скриптом setup_db.sh.

CREATE TABLE IF NOT EXISTS recipes (
    id          INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    title       VARCHAR(255) NOT NULL,
    author      VARCHAR(255) NULL,
    created_at  TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at  TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- unit: l, ml, tbsp, tsp, kg, g, pcs, cloves, to_taste. Для to_taste quantity = 0.
CREATE TABLE IF NOT EXISTS ingredients (
    id          INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    recipe_id   INT UNSIGNED NOT NULL,
    position    SMALLINT UNSIGNED NOT NULL,
    name        VARCHAR(255) NOT NULL,
    quantity    DECIMAL(10,3) NOT NULL DEFAULT 0,
    unit        ENUM('l','ml','tbsp','tsp','kg','g','pcs','cloves','to_taste') NOT NULL,
    is_main     BOOLEAN NOT NULL DEFAULT FALSE,
    CONSTRAINT fk_ingredients_recipe FOREIGN KEY (recipe_id)
        REFERENCES recipes(id) ON DELETE CASCADE,
    INDEX idx_ingredients_recipe (recipe_id, position)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS steps (
    id          INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    recipe_id   INT UNSIGNED NOT NULL,
    position    SMALLINT UNSIGNED NOT NULL,
    text        TEXT NOT NULL,
    CONSTRAINT fk_steps_recipe FOREIGN KEY (recipe_id)
        REFERENCES recipes(id) ON DELETE CASCADE,
    INDEX idx_steps_recipe (recipe_id, position)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
