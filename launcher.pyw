"""Запуск книги рецептов одним щелчком: сервер в фоне + браузер.

.pyw запускается через pythonw.exe без окна консоли. Если книга уже открыта,
повторный запуск просто открывает её в браузере.
"""
import os
import socket
import subprocess
import sys
import threading
import urllib.request
import webbrowser
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
HOST = "127.0.0.1"
PORT = int(os.environ.get("RECEPT_PORT", 8765))
URL = f"http://{HOST}:{PORT}/"
SHORTCUT_NAME = "Книга рецептов"


def already_running():
    try:
        with urllib.request.urlopen(URL + "ping", timeout=1) as resp:
            return resp.read() == b"recept"
    except OSError:
        return False


def port_busy():
    with socket.socket() as s:
        return s.connect_ex((HOST, PORT)) == 0


def redirect_output():
    # У pythonw нет консоли: sys.stdout/stderr равны None, и любая запись падает.
    log_dir = Path(os.environ.get("RECEPT_DATA", BASE_DIR / "data"))
    log_dir.mkdir(parents=True, exist_ok=True)
    log = open(log_dir / "log.txt", "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stderr = log


def show_error(text):
    if sys.platform == "win32":
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, text, SHORTCUT_NAME, 0x10)
    else:
        print(text, file=sys.stderr)


def create_desktop_shortcut():
    """Один раз кладёт ярлык на рабочий стол (только Windows)."""
    if sys.platform != "win32":
        return
    marker = Path(os.environ.get("RECEPT_DATA", BASE_DIR / "data")) / ".shortcut-created"
    if marker.exists():
        return
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    # Пути передаются через переменные окружения, чтобы не возиться с кавычками и кириллицей.
    env = dict(
        os.environ,
        RECEPT_NAME=SHORTCUT_NAME,
        RECEPT_TARGET=str(pythonw),
        RECEPT_ARGS=f'"{BASE_DIR / "launcher.pyw"}"',
        RECEPT_DIR=str(BASE_DIR),
        RECEPT_ICON=str(BASE_DIR / "static" / "icon.ico"),
    )
    script = (
        "$d = [Environment]::GetFolderPath('Desktop');"
        "$s = (New-Object -ComObject WScript.Shell).CreateShortcut("
        "(Join-Path $d ($env:RECEPT_NAME + '.lnk')));"
        "$s.TargetPath = $env:RECEPT_TARGET; $s.Arguments = $env:RECEPT_ARGS;"
        "$s.WorkingDirectory = $env:RECEPT_DIR; $s.IconLocation = $env:RECEPT_ICON;"
        "$s.Save()"
    )
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        env=env,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if result.returncode == 0:
        marker.touch()


def main():
    redirect_output()
    if already_running():
        webbrowser.open(URL)
        return
    if port_busy():
        show_error(f"Порт {PORT} занят другой программой. Книга рецептов не может запуститься.")
        return

    try:
        create_desktop_shortcut()
    except Exception as exc:  # ярлык — не главное, книга должна открыться
        print(f"Не удалось создать ярлык: {exc}")

    from werkzeug.serving import make_server

    from app import app

    app.config["LAUNCHER"] = True
    server = make_server(HOST, PORT, app, threaded=True)
    app.config["SHUTDOWN"] = lambda: threading.Thread(target=server.shutdown).start()
    threading.Timer(0.5, webbrowser.open, args=(URL,)).start()
    server.serve_forever()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        import traceback

        traceback.print_exc()
        show_error(f"Книга рецептов не запустилась:\n{exc}\n\nПодробности в data\\log.txt")
