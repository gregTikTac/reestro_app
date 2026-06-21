# -*- coding: utf-8 -*-
"""Конвертация инструкций Markdown → Word (.docx) через pandoc."""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
PANDOC_CANDIDATES = [
    Path(r"C:\Users\79111\AppData\Local\Pandoc\pandoc.exe"),
    Path(r"C:\Program Files\Pandoc\pandoc.exe"),
    Path(shutil.which("pandoc") or ""),
]


def find_pandoc() -> Path:
    for p in PANDOC_CANDIDATES:
        if p and p.is_file():
            return p
    raise SystemExit(
        "pandoc не найден. Установите: winget install JohnMacFarlane.Pandoc")


def convert(md: Path, docx: Path, *, resource_path: Path | None = None):
    pandoc = find_pandoc()
    cmd = [
        str(pandoc),
        str(md),
        "-o", str(docx),
        "--from", "markdown",
        "--to", "docx",
        "--standalone",
    ]
    if resource_path:
        cmd.extend(["--resource-path", str(resource_path)])
    print(f"  {md.name} -> {docx.name}")
    subprocess.run(cmd, check=True)


def main():
    jobs = [
        (ROOT / "ИНСТРУКЦИЯ.md", ROOT / "ИНСТРУКЦИЯ.docx", None),
        (DOCS / "ПОШАГОВАЯ_ИНСТРУКЦИЯ.md", DOCS / "ПОШАГОВАЯ_ИНСТРУКЦИЯ.docx", DOCS),
        (DOCS / "ИНСТРУКЦИЯ_СОБСТВЕННИКИ.md", DOCS / "ИНСТРУКЦИЯ_СОБСТВЕННИКИ.docx", DOCS),
    ]
    print("Конвертация инструкций в .docx …")
    for md, docx, rp in jobs:
        if not md.is_file():
            raise SystemExit(f"Не найден: {md}")
        convert(md, docx, resource_path=rp)
        size_kb = docx.stat().st_size // 1024
        print(f"    OK ({size_kb} KB): {docx}")
    print("Готово.")


if __name__ == "__main__":
    main()
