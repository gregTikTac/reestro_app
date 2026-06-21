# -*- coding: utf-8 -*-
"""
Подготовка Playwright/Chromium для автоматического сбора с Росреестра.

- Браузер хранится в писабельной папке рядом с приложением (ms-playwright),
  чтобы работать и в собранном .exe.
- Если Chromium ещё не скачан — ставится автоматически (нужен интернет один раз).
  Работает и в обычном Python, и во «замороженном» .exe (через driver Playwright).
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Callable

from .settings import app_dir


def _bundled_browsers_dir() -> Path | None:
    """Каталог с Chromium, вшитым в .exe (PyInstaller кладёт данные в _MEIPASS)."""
    mp = getattr(sys, "_MEIPASS", None)
    if not mp:
        return None
    p = Path(mp) / "ms-playwright"
    if p.is_dir() and any(p.glob("chromium-*")):
        return p
    return None


def browsers_dir() -> Path:
    """
    Каталог с браузерами Playwright.

    Приоритет: явный PLAYWRIGHT_BROWSERS_PATH → вшитый в .exe Chromium →
    писабельная папка рядом с приложением (для авто-скачивания).
    """
    env = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if env and env not in ("0", "1"):
        return Path(env)
    bundled = _bundled_browsers_dir()
    if bundled:
        return bundled
    return app_dir() / "ms-playwright"


def configure_env() -> None:
    """Прописывает PLAYWRIGHT_BROWSERS_PATH (вызывать до запуска Playwright)."""
    bd = browsers_dir()
    # Вшитый каталог трогать на запись не нужно; для скачиваемого — создаём.
    if _bundled_browsers_dir() is None:
        bd.mkdir(parents=True, exist_ok=True)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(bd)


def is_playwright_available() -> bool:
    try:
        import playwright  # noqa: F401
        return True
    except Exception:
        return False


def is_chromium_installed() -> bool:
    """True, если в каталоге браузеров есть распакованный Chromium."""
    bd = browsers_dir()
    if not bd.is_dir():
        return False
    for child in bd.glob("chromium*"):
        if child.is_dir():
            return True
    return False


def _install_command() -> tuple[list[str], dict]:
    """Команда установки Chromium (отдельно для .exe и обычного Python)."""
    env = dict(os.environ)
    env["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_dir())
    if getattr(sys, "frozen", False):
        # В .exe нет python -m playwright: ставим через node-драйвер Playwright.
        from playwright._impl._driver import (
            compute_driver_executable, get_driver_env,
        )
        node, cli = compute_driver_executable()
        env.update(get_driver_env())
        return [str(node), str(cli), "install", "chromium"], env
    return [sys.executable, "-m", "playwright", "install", "chromium"], env


def ensure_chromium(on_log: Callable[[str], None] | None = None) -> bool:
    """
    Гарантирует наличие Chromium. Возвращает True, если браузер готов.
    on_log — необязательный приёмник строк журнала.
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    configure_env()
    if not is_playwright_available():
        log("Playwright не установлен. Установите: pip install playwright")
        return False
    if is_chromium_installed():
        return True

    log("Скачиваю браузер Chromium (один раз, нужен интернет)…")
    cmd, env = _install_command()
    try:
        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.run(
            cmd, env=env, creationflags=creationflags,
            capture_output=True, text=True, timeout=900,
        )
    except Exception as exc:  # noqa: BLE001
        log(f"Не удалось установить Chromium: {exc}")
        return False
    if proc.returncode != 0:
        log("Ошибка установки Chromium:")
        log((proc.stderr or proc.stdout or "")[-800:])
        return False
    ok = is_chromium_installed()
    log("Chromium готов." if ok else "Chromium не появился после установки.")
    return ok
