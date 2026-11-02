"""Запуск сайта внутри Android-приложения. Вызывается из MainActivity."""
import os
import secrets
import sys
import threading


def start(webapp_dir, data_dir):
    """Запускает сервер в фоновом потоке и возвращает «порт токен»."""
    os.environ["RECEPT_DATA"] = data_dir
    sys.path.insert(0, webapp_dir)

    from werkzeug.serving import make_server

    from app import app

    app.config["ANDROID"] = True
    # Без этого токена сервер не отвечает, поэтому другие приложения телефона к нему не обратятся.
    token = secrets.token_urlsafe(32)
    app.config["ACCESS_TOKEN"] = token
    # Порт 0: система выберет свободный, занятый другим приложением не помешает.
    server = make_server("127.0.0.1", 0, app, threaded=True)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return f"{server.server_port} {token}"
