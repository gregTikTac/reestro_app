# -*- coding: utf-8 -*-
"""Сборка ZIP для передачи заказчику (полный комплект «включил и работай»)."""
from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC_DIST = ROOT / "dist" / "ЕГРН-Парсер"
RELEASE_NAME = "ЕГРН-Парсер_0.1.0"
RELEASE_DIR = Path(__file__).resolve().parent / RELEASE_NAME
ZIP_PATH = Path(__file__).resolve().parent / f"{RELEASE_NAME}.zip"

ENGINE_CFG = """{
  "baseUrl": "https://api.kontur.ru",
  "apiKey": "",
  "orgId": "",
  "proxy": "",
  "timeout": 120,
  "retries": 5,
  "pause": 2.0,
  "save_every": 5
}
"""

PROCHITI = """ЕГРН-Парсер 0.1.0 — комплект для заказчика
============================================

БЫСТРЫЙ СТАРТ (включил и работай)
---------------------------------
1. Распакуйте ВЕСЬ архив в папку, напр. C:\\ЕГРН\\
   (нужны ЕГРН-Парсер.exe И папка _internal рядом!)
2. Запустите ЕГРН-Парсер.exe
3. Прочитайте docs\\ИНСТРУКЦИЯ_СОБСТВЕННИКИ.md — там пошагово:
   • настройка API Контура
   • пакетная обработка (Excel + PDF)
   • автоматический сбор формы собственности и обременений с Росреестра

ЧТО ВНУТРИ
----------
  ЕГРН-Парсер.exe           — программа (двойной щелчок)
  _internal\\                — библиотеки и встроенный браузер (НЕ УДАЛЯТЬ!)
  ПРОЧТИ_МЕНЯ.txt           — этот файл
  ИНСТРУКЦИЯ.md / .docx     — полная инструкция по программе
  docs\\
    ИНСТРУКЦИЯ_СОБСТВЕННИКИ.md / .docx  — НОВОЕ: форма собственности + обременения
    ПОШАГОВАЯ_ИНСТРУКЦИЯ.md / .docx      — разбор экрана со скриншотами
    screenshots\\                         — картинки для инструкций
  config.json.example       — образец полей API

НАСТРОЙКА (один раз)
--------------------
  Вкладка «Подключение»:
    • apiKey и orgId — из личного кабинета Контур.Реестро
    • логин Госуслуг кадастрового инженера (для подсказки; пароль — в браузере)
    • «Проверить подключение» → «Сохранить»

РАБОЧИЙ ЦИКЛ
------------
  1) «Пакетная обработка» — входной Excel → папка output → Старт
  2) «Собственники (Росреестр)» — та же папка output →
     «Запустить автоматический сбор» → один раз войти через Госуслуги →
     «Вход выполнен — продолжить» → дождаться конца
     (если автозаполнение не срабатывает — «Проверочный режим» +
     «Данные готовы — записать и далее»)
  3) Открыть report.xlsx и pdf\\ в папке output

  Браузер: Edge или Chrome (системный). Python не нужен.

ВАЖНО
-----
  • Российский IP для Росreestr.ru (без зарубежного VPN)
  • Не переносите один .exe без _internal
  • Папка результатов — корень output, не подпапка pdf\\
  • ФИО физлиц в открытых сведениях по закону не выдаются;
    в отчёте заполняется ФОРМА СОБСТВЕННОСТИ (как в образце)

by BeRealBear
"""


def _convert_docx():
    """Markdown → Word для всех инструкций (если установлен pandoc)."""
    script = Path(__file__).resolve().parent / "md_to_docx.py"
    if not script.is_file():
        return
    try:
        subprocess.run([sys.executable, str(script)], check=True, cwd=str(ROOT))
    except subprocess.CalledProcessError as exc:
        print(f"Предупреждение: не удалось создать .docx (код {exc.returncode})")


def _copy_doc(base: Path, rel: str, docs_dst: Path):
    """Копирует .md и .docx (если есть) в комплект."""
    src_md = base / rel
    if src_md.is_file():
        dst = docs_dst if rel.startswith("docs/") else RELEASE_DIR
        name = Path(rel).name
        if rel.startswith("docs/"):
            shutil.copy2(src_md, docs_dst / name)
        else:
            shutil.copy2(src_md, RELEASE_DIR / name)
        docx = src_md.with_suffix(".docx")
        if docx.is_file():
            if rel.startswith("docs/"):
                shutil.copy2(docx, docs_dst / docx.name)
            else:
                shutil.copy2(docx, RELEASE_DIR / docx.name)


def main():
    exe = SRC_DIST / "ЕГРН-Парсер.exe"
    if not exe.is_file():
        raise SystemExit(
            f"Не найден {exe}.\n"
            "Сначала соберите приложение:\n"
            "  cd reestro_app\n"
            "  .\\venv\\Scripts\\python.exe -m PyInstaller build\\egrn.spec --noconfirm --clean"
        )

    # Chromium в дистрибутиве
    msp = SRC_DIST / "_internal" / "ms-playwright"
    if not msp.is_dir() or not any(msp.glob("chromium-*")):
        print("ВНИМАНИЕ: в dist нет вшитого Chromium — заказчику понадобится интернет "
              "для авто-скачивания браузера при первом сборе.")

    print("Конвертация инструкций в Word…")
    _convert_docx()

    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()

    print("Копирование приложения (может занять минуту — ~650 МБ)…")
    shutil.copytree(SRC_DIST, RELEASE_DIR)

    print("Инструкции и скриншоты…")
    docs_dst = RELEASE_DIR / "docs"
    docs_dst.mkdir(parents=True, exist_ok=True)
    shots_dst = docs_dst / "screenshots"
    shots_dst.mkdir(parents=True, exist_ok=True)

    for rel in [
        "ИНСТРУКЦИЯ.md",
        "docs/ПОШАГОВАЯ_ИНСТРУКЦИЯ.md",
        "docs/ИНСТРУКЦИЯ_СОБСТВЕННИКИ.md",
    ]:
        _copy_doc(ROOT, rel, docs_dst)

    shots_src = ROOT / "docs" / "screenshots"
    if shots_src.is_dir():
        for p in shots_src.glob("*.png"):
            name = p.name
            if name.startswith("03_"):
                name = "03_odinochnyy.png"
            shutil.copy2(p, shots_dst / name)

    example = ROOT.parent / "reestro" / "config.json.example"
    if example.is_file():
        shutil.copy2(example, RELEASE_DIR / "config.json.example")

    # Убрать чужие ключи из комплекта
    key_file = RELEASE_DIR / "_internal" / "reestro" / "TZ" / "ключ API.txt"
    if key_file.is_file():
        key_file.unlink()

    cfg_path = RELEASE_DIR / "_internal" / "reestro" / "config.json"
    if cfg_path.parent.is_dir():
        cfg_path.write_text(ENGINE_CFG, encoding="utf-8")

    (RELEASE_DIR / "ПРОЧТИ_МЕНЯ.txt").write_text(PROCHITI, encoding="utf-8")

    # Чистые настройки у заказчика (без ключей и путей разработчика)
    (RELEASE_DIR / "config.json").write_text(ENGINE_CFG, encoding="utf-8")
    (RELEASE_DIR / "settings.json").write_text(
        '{\n  "last_input": "",\n  "last_output": "",\n  "auto_name": true,\n'
        '  "report_name": "",\n  "gosuslugi_login": "",\n  "range_from": "",\n'
        '  "range_to": "",\n  "limit": ""\n}\n',
        encoding="utf-8",
    )

    n_files = sum(1 for _ in RELEASE_DIR.rglob("*") if _.is_file())
    dir_mb = sum(f.stat().st_size for f in RELEASE_DIR.rglob("*") if f.is_file()) / (1024 ** 2)
    print(f"Файлов в комплекте: {n_files}, размер папки: {dir_mb:.0f} MB")

    print("Архивирование (может занять несколько минут)…")
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED, compresslevel=3) as zf:
        for path in RELEASE_DIR.rglob("*"):
            if path.is_file():
                arc = Path(RELEASE_NAME) / path.relative_to(RELEASE_DIR)
                zf.write(path, arc.as_posix())

    zip_mb = ZIP_PATH.stat().st_size / (1024 * 1024)
    print(f"\nГотово: {ZIP_PATH}")
    print(f"Размер ZIP: {zip_mb:.0f} MB")
    print(f"Распакованная папка: {RELEASE_DIR}")


if __name__ == "__main__":
    main()
