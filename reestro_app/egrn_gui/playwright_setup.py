# -*- coding: utf-8 -*-
"""
Playwright для автоматического сбора с Росreestr.

Браузер — только системный Microsoft Edge (channel=msedge).
Chromium не скачивается и не вшивается в дистрибутив.
"""
from __future__ import annotations


def configure_env() -> None:
    """Заглушка для совместимости — Edge не требует PLAYWRIGHT_BROWSERS_PATH."""


def is_playwright_available() -> bool:
    try:
        import playwright  # noqa: F401
        return True
    except Exception:
        return False
