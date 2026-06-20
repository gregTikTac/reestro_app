# -*- coding: utf-8 -*-
"""Сборка ZIP для передачи заказчику."""
from __future__ import annotations

import shutil
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
  "orgId": ""
}
"""

PROCHITI = """ЕГРН-Парсер 0.1.0 — комплект для заказчика
==========================================

С чего начать
  1. Распакуйте архив в любую папку, напр. C:\\Реестро\\
  2. Запускайте ЕГРН-Парсер.exe из этой папки (нужна папка _internal рядом!)
  3. Прочитайте ИНСТРУКЦИЯ.md
  4. Пошаговый разбор экрана: docs\\ПОШАГОВАЯ_ИНСТРУКЦИЯ.md (со скриншотами)

Настройка API
  - На вкладке «Подключение» введите apiKey и orgId из Контур.Реестро
  - Нажмите «Проверить подключение» → «Сохранить»
  - Образец полей: config.json.example

Состав папки
  ЕГРН-Парсер.exe      — программа
  _internal\\           — библиотеки (не удалять!)
  ИНСТРУКЦИЯ.md        — полная инструкция
  docs\\                — пошаговая инструкция + screenshots\\
  config.json.example  — пример полей API

Важно
  - Не переносите один .exe без _internal
  - Папка результатов — корень output, не подпапка pdf\\
  - Python не требуется

by BeRealBear
"""


def main():
    exe = SRC_DIST / "ЕГРН-Парсер.exe"
    if not exe.is_file():
        raise SystemExit(f"Не найден {exe}. Сначала запустите build\\build_exe.bat")

    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()

    print("Копирование приложения...")
    shutil.copytree(SRC_DIST, RELEASE_DIR)

    print("Инструкции и скриншоты...")
    shutil.copy2(ROOT / "ИНСТРУКЦИЯ.md", RELEASE_DIR / "ИНСТРУКЦИЯ.md")
    docs = RELEASE_DIR / "docs"
    shots_dst = docs / "screenshots"
    shots_dst.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ROOT / "docs" / "ПОШАГОВАЯ_ИНСТРУКЦИЯ.md", docs / "ПОШАГОВАЯ_ИНСТРУКЦИЯ.md")

    shots_src = ROOT / "docs" / "screenshots"
    for p in shots_src.glob("*.png"):
        name = p.name
        if name.startswith("03_"):
            name = "03_odinochnyy.png"
        shutil.copy2(p, shots_dst / name)

    example = ROOT.parent / "reestro" / "config.json.example"
    if example.is_file():
        shutil.copy2(example, RELEASE_DIR / "config.json.example")

    key_file = RELEASE_DIR / "_internal" / "reestro" / "TZ" / "ключ API.txt"
    if key_file.is_file():
        key_file.unlink()

    (RELEASE_DIR / "_internal" / "reestro" / "config.json").write_text(
        ENGINE_CFG, encoding="utf-8")
    (RELEASE_DIR / "ПРОЧТИ_МЕНЯ.txt").write_text(PROCHITI, encoding="utf-8")

    print("Архивирование...")
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for path in RELEASE_DIR.rglob("*"):
            if path.is_file():
                arc = Path(RELEASE_NAME) / path.relative_to(RELEASE_DIR)
                zf.write(path, arc.as_posix())

    size_mb = ZIP_PATH.stat().st_size / (1024 * 1024)
    print(f"Готово: {ZIP_PATH}")
    print(f"Размер: {size_mb:.1f} MB")
    print(f"Файлов в комплекте: {sum(1 for _ in RELEASE_DIR.rglob('*') if _.is_file())}")


if __name__ == "__main__":
    main()
