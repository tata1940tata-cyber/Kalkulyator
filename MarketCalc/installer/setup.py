# -*- coding: utf-8 -*-
"""Установщик МАРКЕТ — установка в один клик."""

from __future__ import annotations

import os
import subprocess
import sys
import winreg
from pathlib import Path

APP_EXE = "MarketCalc.exe"
VERSION = "1.1.0"
PUBLISHER = "МАРКЕТ"

BG = "#0a0c10"
PANEL = "#121822"
LINE = "#243044"
TEXT = "#e9edf5"
MUTED = "#8b95a8"
ACCENT = "#ffd43b"
ACCENT_DARK = "#1b1505"

DEFAULT_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Programs" / "MarketCalc"
START_MENU = Path(os.environ.get("APPDATA", "")) / "Microsoft/Windows/Start Menu/Programs"
DESKTOP = Path.home() / "Desktop"


def payload_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS")) / "payload" / APP_EXE
    local = Path(__file__).resolve().parent.parent / "dist" / APP_EXE
    if local.is_file():
        return local
    raise FileNotFoundError("Не найден файл приложения MarketCalc.exe")


def create_shortcut(link: Path, target: Path) -> None:
    link.parent.mkdir(parents=True, exist_ok=True)
    ps = (
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{link}');"
        "$s.TargetPath='{target}';"
        "$s.WorkingDirectory='{work}';"
        "$s.IconLocation='{target},0';"
        "$s.Description='МАРКЕТ — калькулятор юнит-экономики';"
        "$s.Save()"
    ).format(link=str(link).replace("'", "''"),
             target=str(target).replace("'", "''"),
             work=str(target.parent).replace("'", "''"))
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        capture_output=True,
        creationflags=0x08000000,
    )


def write_registry(exe: Path) -> None:
    key = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\MarketCalc"
    size_kb = max(1, exe.stat().st_size // 1024)
    values = {
        "DisplayName": "МАРКЕТ — калькулятор юнит-экономики",
        "DisplayVersion": VERSION,
        "Publisher": PUBLISHER,
        "InstallLocation": str(exe.parent),
        "DisplayIcon": f"{exe},0",
        "UninstallString": f'"{exe}" --uninstall',
        "QuietUninstallString": f'"{exe}" --uninstall',
        "NoModify": 1,
        "NoRepair": 1,
        "EstimatedSize": size_kb,
    }
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key) as handle:
        for name, value in values.items():
            kind = winreg.REG_DWORD if isinstance(value, int) else winreg.REG_SZ
            winreg.SetValueEx(handle, name, 0, kind, value)


def install(target: Path, desktop_shortcut=True, start_menu=True,
            launch=False, on_progress=None, on_status=None) -> Path:
    source = payload_path()
    target.mkdir(parents=True, exist_ok=True)
    exe = target / APP_EXE

    subprocess.run(["taskkill", "/IM", APP_EXE, "/F"], capture_output=True,
                   creationflags=0x08000000)

    total = source.stat().st_size
    done = 0
    with open(source, "rb") as src, open(exe, "wb") as dst:
        while True:
            chunk = src.read(1024 * 512)
            if not chunk:
                break
            dst.write(chunk)
            done += len(chunk)
            if on_progress:
                on_progress(done / total)

    if on_status:
        on_status("Создаю ярлыки…")
    if desktop_shortcut:
        create_shortcut(DESKTOP / "МАРКЕТ.lnk", exe)
    if start_menu:
        create_shortcut(START_MENU / "МАРКЕТ.lnk", exe)

    if on_status:
        on_status("Регистрирую приложение…")
    write_registry(exe)

    if launch:
        subprocess.Popen([str(exe)], cwd=str(target), close_fds=True)
    return exe


# ------------------------------------------------------------------ GUI
def run_gui() -> None:
    import tkinter as tk
    from tkinter import messagebox, ttk

    root = tk.Tk()
    root.title("Установка МАРКЕТ")
    root.configure(bg=BG)
    root.geometry("540x330")
    root.resizable(False, False)

    try:
        root.iconbitmap(default=str(Path(getattr(sys, "_MEIPASS", "")) / "assets" / "icon.ico"))
    except Exception:  # noqa: BLE001
        pass

    def dark_titlebar():
        try:
            import ctypes
            root.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
            value = ctypes.c_int(1)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
        except Exception:  # noqa: BLE001
            pass

    dark_titlebar()

    wrap = tk.Frame(root, bg=BG, padx=28, pady=24)
    wrap.pack(fill="both", expand=True)

    tk.Label(wrap, text="МАРКЕТ", bg=BG, fg=ACCENT,
             font=("Segoe UI", 24, "bold")).pack(anchor="w")
    tk.Label(wrap, text="Калькулятор юнит-экономики Яндекс Маркета",
             bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 22))

    status_var = tk.StringVar(value="Подготавливаю установку…")
    status = tk.Label(wrap, textvariable=status_var, bg=BG, fg=TEXT,
                      font=("Segoe UI", 11), anchor="w")
    status.pack(fill="x", pady=(0, 10))

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Market.Horizontal.TProgressbar", troughcolor=PANEL, bordercolor=PANEL,
                    background=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT, thickness=9)
    progress = ttk.Progressbar(wrap, style="Market.Horizontal.TProgressbar",
                               orient="horizontal", mode="determinate", maximum=100)
    progress.pack(fill="x")

    tk.Label(
        wrap,
        text=f"Устанавливается для текущего пользователя:\n{DEFAULT_DIR}\n"
             "Ярлык появится на рабочем столе и в меню «Пуск». Прав администратора не требуется.",
        bg=BG, fg="#5f6b80", font=("Segoe UI", 8), justify="left", anchor="w",
    ).pack(anchor="w", pady=(16, 0))

    retry = tk.Button(wrap, text="Повторить", bg=ACCENT, fg=ACCENT_DARK, relief="flat",
                      activebackground="#ffe37a", activeforeground=ACCENT_DARK,
                      font=("Segoe UI", 10, "bold"), cursor="hand2", pady=8)
    retry.pack(fill="x", pady=(14, 0))
    retry.pack_forget()

    tk.Label(wrap, text=f"Версия {VERSION}", bg=BG, fg="#3f4a5c",
             font=("Segoe UI", 8)).pack(side="bottom", anchor="w")

    def set_progress(value):
        progress["value"] = value * 100
        root.update_idletasks()

    def set_status(text):
        status_var.set(text)
        root.update_idletasks()

    finished = {"value": False}

    def start():
        retry.pack_forget()
        progress["value"] = 0
        set_status("Копирую файлы…")
        try:
            exe = install(DEFAULT_DIR, desktop_shortcut=True, start_menu=True, launch=False,
                          on_progress=set_progress, on_status=set_status)
        except Exception as exc:  # noqa: BLE001
            set_status("Не удалось установить")
            messagebox.showerror("Ошибка установки", f"Установка не завершена:\n{exc}", parent=root)
            retry.pack(fill="x", pady=(14, 0))
            return

        progress["value"] = 100
        set_status("Готово. Запускаю МАРКЕТ…")
        finished["value"] = True
        subprocess.Popen([str(exe)], cwd=str(exe.parent), close_fds=True)
        root.after(1200, root.destroy)

    retry.configure(command=start)
    root.after(450, start)
    root.mainloop()


def main() -> None:
    args = sys.argv[1:]
    if "--silent" in args:
        target = DEFAULT_DIR
        if "--target" in args:
            target = Path(args[args.index("--target") + 1])
        shortcuts = "--no-shortcuts" not in args
        install(target, desktop_shortcut=shortcuts, start_menu=shortcuts, launch="--launch" in args)
        return
    run_gui()


if __name__ == "__main__":
    main()
