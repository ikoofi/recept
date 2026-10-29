#!/usr/bin/env bash
# Собирает готовую папку для Windows: встроенный Python + библиотеки + программа.
# Результат: dist/КнигаРецептов.zip — распаковать на Windows и запустить «Книга рецептов.bat».
# Собирать можно на Linux: библиотеки скачиваются сразу в виде пакетов для Windows.
set -euo pipefail
cd "$(dirname "$0")"

PY_VERSION=3.12.10
PY_SHORT=312
NAME="КнигаРецептов"
OUT="dist/$NAME"
CACHE="build"
EMBED_ZIP="$CACHE/python-$PY_VERSION-embed-amd64.zip"
PIP=".venv/bin/pip"

mkdir -p "$CACHE"
if [ ! -f "$EMBED_ZIP" ]; then
    echo "Скачиваю встроенный Python $PY_VERSION для Windows..."
    curl -fL -o "$EMBED_ZIP.part" \
        "https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
    mv "$EMBED_ZIP.part" "$EMBED_ZIP"
fi

rm -rf "$OUT" "dist/$NAME.zip"
mkdir -p "$OUT/python"
unzip -q "$EMBED_ZIP" -d "$OUT/python"

echo "Ставлю библиотеки для Windows..."
"$PIP" install -q --disable-pip-version-check \
    --target "$OUT/python/Lib/site-packages" \
    --platform win_amd64 --python-version "$PY_SHORT" --implementation cp \
    --only-binary=:all: -r requirements.txt

# Встроенный Python берёт пути только из этого файла: добавляем библиотеки и папку программы.
cat > "$OUT/python/python$PY_SHORT._pth" <<EOF
python$PY_SHORT.zip
.
Lib\\site-packages
..
EOF

cp app.py launcher.pyw schema.sql "$OUT/"
cp -r templates static "$OUT/"

# Первый запуск: открывает книгу и кладёт ярлык «Книга рецептов» на рабочий стол.
# Если .bat запустили прямо из архива (WinRAR/Проводник достают только его во временную
# папку), папки python рядом нет — показываем понятное окно. Текст окна передаётся
# PowerShell в base64 (UTF-16), чтобы кириллица не зависела от кодировки консоли.
NOT_UNPACKED=$(.venv/bin/python - <<'PY'
import base64
msg = ("Сначала распакуйте архив.\n\n"
       "Щёлкните по архиву КнигаРецептов правой кнопкой мыши, выберите «Извлечь всё…» "
       "(или «Извлечь в КнигаРецептов\\» в WinRAR), затем запустите "
       "«Книга рецептов.bat» из распакованной папки.")
ps = ("Add-Type -AssemblyName PresentationFramework;"
      "[System.Windows.MessageBox]::Show('%s', 'Книга рецептов', 'OK', 'Warning') | Out-Null"
      % msg.replace("'", "''"))
print(base64.b64encode(ps.encode("utf-16-le")).decode())
PY
)
printf '%s\r\n' \
    '@echo off' \
    'if not exist "%~dp0python\pythonw.exe" (' \
    "    powershell -NoProfile -EncodedCommand $NOT_UNPACKED" \
    '    exit /b 1' \
    ')' \
    'start "" "%~dp0python\pythonw.exe" "%~dp0launcher.pyw"' \
    > "$OUT/Книга рецептов.bat"

cat > "$OUT/Прочитай меня.txt" <<'EOF'
КНИГА РЕЦЕПТОВ

Запуск
  Дважды щёлкните «Книга рецептов.bat». Откроется браузер с книгой,
  а на рабочем столе появится ярлык «Книга рецептов» — дальше запускайте его.
  Устанавливать ничего не нужно.

Выход
  Кнопка «Выход» в правом верхнем углу книги.

Где хранятся рецепты
  В папке data (файл recipes.db). Не удаляйте её.
  Для резервной копии скопируйте папку data или нажмите «Экспорт всех».

Передать рецепты другому человеку
  «Экспорт» на странице рецепта или «Экспорт всех» — получится файл .json.
  Другой человек открывает свою книгу, нажимает «Импорт» и выбирает этот файл.
  Рецепты с уже существующим названием не добавляются.

Перенос папки
  Если папку переместили, удалите старый ярлык с рабочего стола,
  удалите файл data\.shortcut-created и снова запустите «Книга рецептов.bat».
EOF
sed -i 's/$/\r/' "$OUT/Прочитай меня.txt"

find "$OUT" -name "__pycache__" -type d -prune -exec rm -rf {} +
# zipfile из Python помечает имена как UTF-8, и Проводник Windows показывает кириллицу правильно.
.venv/bin/python -c 'import shutil, sys; shutil.make_archive(sys.argv[1], "zip", "dist", sys.argv[2])' \
    "dist/$NAME" "$NAME"
echo "Готово: dist/$NAME.zip ($(du -h "dist/$NAME.zip" | cut -f1))"
