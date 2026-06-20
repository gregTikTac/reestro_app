# -*- coding: utf-8 -*-
"""
Проверка, что PyInstaller собрал полную папку dist (с _internal и python311.dll).
Запуск: python build/verify_dist.py
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
DIST = PROJECT / "dist"


def find_app_dir() -> Path | None:
    if not DIST.is_dir():
        return None
    for d in DIST.iterdir():
        if d.is_dir() and list(d.glob("*.exe")):
            return d
    return None


def main() -> int:
    app = find_app_dir()
    if not app:
        print(f"ОШИБКА: в {DIST} нет папки с .exe")
        print("Выполните: build\\build_exe.bat")
        return 1

    internal = app / "_internal"
    dll = internal / "python311.dll"
    exe = next(app.glob("*.exe"))

    errors = []
    if not internal.is_dir():
        errors.append(f"Нет папки {internal.name} рядом с exe")
    if not dll.is_file():
        errors.append(f"Нет {dll.name} в _internal")

    if errors:
        print("ОШИБКА: сборка неполная:")
        for e in errors:
            print(" -", e)
        print()
        print("НЕ копируйте и НЕ запускайте exe из build\\egrn\\")
        print("Пересоберите: build\\build_exe.bat")
        return 1

    # README для пользователя в папке dist
    readme = app / "ПРОЧТИ_МЕНЯ.txt"
    readme.write_text(
        f"ЕГРН-Парсер — как запускать\n"
        f"{'=' * 40}\n\n"
        f"1. Запускайте ТОЛЬКО этот файл:\n"
        f"   {exe.name}\n\n"
        f"2. НЕ переносите один .exe без папки _internal!\n"
        f"   Нужна ВСЯ эта папка целиком ({app.name}).\n\n"
        f"3. НЕ запускайте exe из build\\egrn\\ — там ошибка python311.dll.\n\n"
        f"4. Инструкция: см. ИНСТРУКЦИЯ.md (если скопирована).\n\n"
        f"Папка _internal содержит {len(list(internal.iterdir()))} файлов библиотек.\n",
        encoding="utf-8",
    )

    instr_src = PROJECT / "ИНСТРУКЦИЯ.md"
    instr_dst = app / "ИНСТРУКЦИЯ.md"
    if instr_src.is_file():
        shutil.copy2(instr_src, instr_dst)

    print(f"OK: {app.name}")
    print(f"   exe: {exe.name}")
    print(f"   _internal: {len(list(internal.iterdir()))} файлов")
    print(f"   python311.dll: найден")
    print(f"   ПРОЧТИ_МЕНЯ.txt создан")
    return 0


if __name__ == "__main__":
    sys.exit(main())
