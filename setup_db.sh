#!/usr/bin/env bash
# Создаёт базу и пользователя MySQL из .env, затем загружает schema.sql.
# Запуск: ./setup_db.sh  (попросит пароль sudo для доступа к root в MySQL)
set -euo pipefail
cd "$(dirname "$0")"
set -a; source .env; set +a

sudo mysql <<SQL
CREATE DATABASE IF NOT EXISTS \`${DB_NAME}\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
CREATE USER IF NOT EXISTS '${DB_USER}'@'localhost' IDENTIFIED BY '${DB_PASSWORD}';
ALTER USER '${DB_USER}'@'localhost' IDENTIFIED BY '${DB_PASSWORD}';
GRANT ALL PRIVILEGES ON \`${DB_NAME}\`.* TO '${DB_USER}'@'localhost';
FLUSH PRIVILEGES;
SQL

mysql -h "$DB_HOST" -P "$DB_PORT" -u "$DB_USER" -p"$DB_PASSWORD" "$DB_NAME" < schema.sql
echo "База '${DB_NAME}' готова."
