# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec для ЕГРН-Парсера.
Собирает GUI + вшивает движок (папку ../reestro) как данные.
Запуск: pyinstaller build/egrn.spec --noconfirm --clean
"""
import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# SPECPATH — каталог этого .spec (reestro_app/build); PROJECT — reestro_app
PROJECT = Path(SPECPATH).resolve().parent
ENGINE = PROJECT.parent / "reestro"            # ../reestro

# Включаем только нужные файлы движка (без output/dist/venv)
engine_datas = []
for name in ["reestro_parser.py", "cleanup_output.py", "rebuild_from_cache.py",
             "fetch_rosreestr.py", "verify_ownership_preview.py", "config.json"]:
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
    + collect_submodules("playwright")
    + ["openpyxl", "requests", "greenlet", "pyee",
       "PySide6.QtCore", "PySide6.QtGui", "PySide6.QtWidgets"]
)

# Драйвер Playwright (node.exe + cli.js) — нужен для автоматического сбора.
engine_datas += collect_data_files("playwright")


def _find_ms_playwright() -> Path | None:
    """Каталог установленного Chromium для вшивания в .exe."""
    candidates = [
        PROJECT / "ms-playwright",                       # локально рядом с проектом
        Path(os.environ.get("LOCALAPPDATA", "")) / "ms-playwright",
    ]
    for c in candidates:
        if c.is_dir() and any(c.glob("chromium-*")):
            return c
    return None


# Вшиваем сам Chromium в дистрибутив (офлайн «из коробки»). headless_shell не
# нужен — мы запускаем браузер в видимом режиме, поэтому его пропускаем (экономия).
_msp = _find_ms_playwright()
if _msp:
    for f in _msp.rglob("*"):
        if not f.is_file():
            continue
        rel = f.relative_to(_msp)
        top = rel.parts[0] if rel.parts else ""
        if top.startswith("chromium_headless_shell"):
            continue
        dest = "ms-playwright"
        if len(rel.parts) > 1:
            dest = str(Path("ms-playwright") / rel.parent)
        engine_datas.append((str(f), dest))
    print(f"[spec] Chromium вшивается из: {_msp}")
else:
    print("[spec] ВНИМАНИЕ: ms-playwright не найден — Chromium НЕ вшит. "
          "Выполните: python -m playwright install chromium")

a = Analysis(
    [str(PROJECT / "app_entry.py")],
    pathex=[str(PROJECT)],
    binaries=[],
    datas=engine_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib"],
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
