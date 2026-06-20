# -*- coding: utf-8 -*-
"""
Хранение параметров подключения (config.json) и настроек интерфейса (settings.json).

Файлы лежат рядом с приложением:
  - в разработке: reestro_app/config.json, reestro_app/settings.json
  - в собранном .exe: рядом с исполняемым файлом
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path


DEFAULT_CONFIG = {
    "baseUrl": "https://api.kontur.ru",
    "apiKey": "",
    "orgId": "",
    "proxy": "",
    "timeout": 120,
    "retries": 5,
    "pause": 2.0,
    "save_every": 5,
}

DEFAULT_SETTINGS = {
    "last_input": "",
    "last_output": "",
    "auto_name": True,
    "report_name": "",
    "range_from": "",
    "range_to": "",
    "limit": "",
}


def app_dir() -> Path:
    """Каталог, куда писать config/settings (рядом с .exe или с проектом)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _path(name: str) -> Path:
    override = os.environ.get("EGRN_CONFIG_DIR")
    base = Path(override) if override else app_dir()
    base.mkdir(parents=True, exist_ok=True)
    return base / name


def _load(name: str, defaults: dict) -> dict:
    p = _path(name)
    data = dict(defaults)
    if p.exists():
        try:
            with open(p, encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                data.update({k: v for k, v in loaded.items() if v != "" or k in loaded})
        except (OSError, ValueError):
            pass
    return data


def _save(name: str, data: dict):
    p = _path(name)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_config() -> dict:
    return _load("config.json", DEFAULT_CONFIG)


def save_config(cfg: dict):
    _save("config.json", cfg)


def config_path() -> Path:
    return _path("config.json")


def load_settings() -> dict:
    return _load("settings.json", DEFAULT_SETTINGS)


def save_settings(s: dict):
    _save("settings.json", s)
