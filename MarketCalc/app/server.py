# -*- coding: utf-8 -*-
"""Локальный сервер приложения «МАРКЕТ»."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
import traceback
import urllib.parse
import urllib.request
import webbrowser
import winreg
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from dataio import (
    WEB_DIR,
    build_export,
    build_template,
    load_bundle,
    reset_bundle,
    save_bundle,
    user_data_dir,
)
from parsing import DataError, parse_upload

LOG_PATH = user_data_dir() / "app.log"
PORT_FILE = user_data_dir() / "port.txt"
LAST_SEEN = time.time()
MAX_UPLOAD = 25 * 1024 * 1024


def log(message: str) -> None:
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except OSError:
        pass


CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
    ".woff2": "font/woff2",
}


class Handler(BaseHTTPRequestHandler):
    server_version = "MarketCalc/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):  # noqa: A003
        pass

    def send_bytes(self, data: bytes, content_type: str, status: int = 200, headers=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def send_json(self, payload, status: int = 200):
        self.send_bytes(json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                        "application/json; charset=utf-8", status)

    def send_file(self, data: bytes, filename: str, content_type: str):
        quoted = urllib.parse.quote(filename)
        self.send_bytes(data, content_type, headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{quoted}",
        })

    def read_body(self) -> bytes:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD:
            raise ValueError("Файл слишком большой")
        return self.rfile.read(length) if length else b""

    def serve_static(self, path: str):
        if path in ("/", ""):
            path = "/index.html"
        target = (WEB_DIR / path.lstrip("/")).resolve()
        try:
            target.relative_to(WEB_DIR.resolve())
        except ValueError:
            return self.send_bytes(b"forbidden", "text/plain", 403)
        if not target.is_file():
            return self.send_bytes("Не найдено".encode("utf-8"), "text/plain; charset=utf-8", 404)
        self.send_bytes(target.read_bytes(),
                        CONTENT_TYPES.get(target.suffix.lower(), "application/octet-stream"))

    def do_GET(self):  # noqa: N802
        global LAST_SEEN
        path = urllib.parse.urlparse(self.path).path
        try:
            if path == "/api/ping":
                LAST_SEEN = time.time()
                return self.send_json({"ok": True})
            if path == "/api/data":
                return self.send_json(load_bundle())
            if path == "/api/template":
                return self.send_file(
                    build_template(load_bundle()),
                    "Шаблон_данных_Маркет.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            if path == "/favicon.ico":
                return self.send_bytes(b"", "image/x-icon", 204)
            if path.startswith("/api/"):
                return self.send_json({"error": "Не найдено"}, 404)
            return self.serve_static(path)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return self.send_json({"error": str(exc)}, 500)

    def do_POST(self):  # noqa: N802
        global LAST_SEEN
        LAST_SEEN = time.time()
        path = urllib.parse.urlparse(self.path).path
        try:
            if path == "/api/upload":
                return self.handle_upload()
            if path == "/api/export":
                payload = json.loads(self.read_body().decode("utf-8"))
                data = build_export(payload)
                return self.send_file(
                    data, f"Юнит-экономика_{time.strftime('%Y-%m-%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            if path == "/api/reset":
                return self.send_json(reset_bundle())
            return self.send_json({"error": "Не найдено"}, 404)
        except DataError as exc:
            return self.send_json({"error": str(exc)}, 400)
        except Exception as exc:  # noqa: BLE001
            log(traceback.format_exc())
            return self.send_json({"error": f"Не удалось обработать файл: {exc}"}, 500)

    def handle_upload(self):
        filename = urllib.parse.unquote(self.headers.get("X-Filename", "")) or "data.xlsx"
        data = self.read_body()
        if not data:
            raise DataError("Файл пустой. Выберите таблицу .xlsx или .csv.")
        bundle = parse_upload(data, filename, load_bundle())
        save_bundle(bundle)
        log(f"upload: {filename} -> {len(bundle['skus'])} sku")
        return self.send_json(bundle)


# ---------------------------------------------------------------- запуск
def find_browser():
    candidates = []
    for env in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if not base:
            continue
        candidates += [
            Path(base) / "Google/Chrome/Application/chrome.exe",
            Path(base) / "Microsoft/Edge/Application/msedge.exe",
        ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return None


def open_window(url: str) -> None:
    browser = find_browser()
    if browser and ("chrome" in browser.lower() or "edge" in browser.lower()):
        profile = user_data_dir() / "browser"
        profile.mkdir(parents=True, exist_ok=True)
        args = [
            browser, f"--app={url}", "--window-size=1480,960",
            f"--user-data-dir={profile}", "--no-first-run",
            "--no-default-browser-check", "--disable-features=Translate",
        ]
        try:
            subprocess.Popen(args, close_fds=True)
            return
        except OSError as exc:
            log(f"browser launch failed: {exc}")
    webbrowser.open(url)


def existing_instance():
    try:
        port = int(PORT_FILE.read_text().strip())
    except (OSError, ValueError):
        return None
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=0.6) as response:
            if response.status == 200:
                return f"http://127.0.0.1:{port}"
    except Exception:  # noqa: BLE001
        return None
    return None


def watchdog(server):
    while True:
        time.sleep(15)
        if time.time() - LAST_SEEN > 180:
            log("idle shutdown")
            server.shutdown()
            return


def uninstall() -> None:
    install_dir = Path(sys.executable).resolve().parent
    appdata = os.environ.get("APPDATA", "")
    shortcuts = [
        Path.home() / "Desktop" / "МАРКЕТ.lnk",
        Path(appdata) / "Microsoft/Windows/Start Menu/Programs" / "МАРКЕТ.lnk",
    ]
    for shortcut in shortcuts:
        try:
            shortcut.unlink()
        except OSError:
            pass
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                         r"Software\Microsoft\Windows\CurrentVersion\Uninstall\MarketCalc")
    except OSError:
        pass
    command = (
        'ping -n 3 127.0.0.1 >nul & '
        'taskkill /IM MarketCalc.exe /F >nul 2>&1 & '
        f'rmdir /s /q "{install_dir}"'
    )
    args = ["cmd.exe", "/c", command]
    kwargs = dict(cwd=str(Path(os.environ.get("TEMP", install_dir))), close_fds=True)
    try:
        subprocess.Popen(args, creationflags=0x01000000 | 0x00000008, **kwargs)
    except OSError:
        subprocess.Popen(args, creationflags=0x08000000, **kwargs)


def main() -> None:
    url = existing_instance()
    if url:
        open_window(url)
        return

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    port = server.server_address[1]
    url = f"http://127.0.0.1:{port}"
    PORT_FILE.write_text(str(port), encoding="utf-8")
    log(f"start {url}")

    threading.Thread(target=watchdog, args=(server,), daemon=True).start()
    if not os.environ.get("MARKETCALC_NO_BROWSER") and "--no-browser" not in sys.argv:
        threading.Timer(0.7, lambda: open_window(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        try:
            PORT_FILE.unlink()
        except OSError:
            pass


if __name__ == "__main__":
    if "--uninstall" in sys.argv:
        uninstall()
        sys.exit(0)
    main()
