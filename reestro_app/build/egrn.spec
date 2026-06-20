# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec для ЕГРН-Парсера.
Собирает GUI + вшивает движок (папку ../reestro) как данные.
Запуск: pyinstaller build/egrn.spec --noconfirm --clean
"""
import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules

# SPECPATH — каталог этого .spec (reestro_app/build); PROJECT — reestro_app
PROJECT = Path(SPECPATH).resolve().parent
ENGINE = PROJECT.parent / "reestro"            # ../reestro

# Включаем только нужные файлы движка (без output/dist/venv)
engine_datas = []
for name in ["reestro_parser.py", "cleanup_output.py", "rebuild_from_cache.py",
             "config.json"]:
    src = ENGINE / name
    if src.exists():
        engine_datas.append((str(src), "reestro"))
# справочники и шаблоны, если есть
for sub in ["input", "TZ"]:
    d = ENGINE / sub
    if d.exists():
        engine_datas.append((str(d), f"reestro/{sub}"))

# ресурсы приложения (иконка)
assets_dir = PROJECT / "assets"
if assets_dir.exists():
    engine_datas.append((str(assets_dir), "assets"))

hiddenimports = (
    collect_submodules("fpdf")
    + collect_submodules("egrn_gui")
    + ["openpyxl", "requests", "PySide6.QtCore", "PySide6.QtGui", "PySide6.QtWidgets"]
)

a = Analysis(
    [str(PROJECT / "app_entry.py")],
    pathex=[str(PROJECT)],
    binaries=[],
    datas=engine_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "playwright"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="ЕГРН-Парсер",
    console=False,
    icon=str(PROJECT / "assets" / "icon.ico") if (PROJECT / "assets" / "icon.ico").exists() else None,
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    name="ЕГРН-Парсер",
)
