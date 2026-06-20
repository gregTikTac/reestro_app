# -*- coding: utf-8 -*-
"""
Пути к папке результатов: корень output, cache/json, проверки.
"""
from __future__ import annotations

from pathlib import Path


def cache_json_dir(output_root: Path) -> Path:
    return Path(output_root).resolve() / "cache" / "json"


def count_cache_json(output_root: Path) -> int:
    d = cache_json_dir(output_root)
    if not d.is_dir():
        return 0
    return sum(1 for _ in d.glob("*.json"))


def resolve_output_root(path: str | Path) -> Path:
    """
    Находит корень папки результатов.

    Пользователь часто выбирает в обзоре подпапку pdf\\ или cache\\json\\
    вместо корня output. Поднимаемся вверх, пока не найдём cache/json или report.xlsx.
    """
    p = Path(path).resolve()
    if p.is_file():
        p = p.parent

    original = p
    for _ in range(6):
        cache = p / "cache" / "json"
        if cache.is_dir() and any(cache.glob("*.json")):
            return p
        if (p / "report.xlsx").is_file():
            return p
        if p.name.lower() == "json" and p.parent.name.lower() == "cache":
            return p.parent.parent
        if p.name.lower() == "cache" and (p / "json").is_dir():
            return p.parent
        if p.name.lower() == "pdf" and (p.parent / "cache" / "json").is_dir():
            return p.parent
        parent = p.parent
        if parent == p:
            break
        p = parent

    return original


def describe_output_folder(path: str | Path) -> str:
    """Краткая подсказка для UI: что найдено в папке."""
    root = resolve_output_root(path)
    n_json = count_cache_json(root)
    cache = cache_json_dir(root)
    report = root / "report.xlsx"
    pdf = root / "pdf"

    parts = [f"корень: {root}"]
    if n_json:
        parts.append(f"JSON-кэш: {n_json} файлов ({cache})")
    else:
        parts.append(f"JSON-кэш: не найден (ожидается {cache})")
    if report.is_file():
        parts.append("report.xlsx: есть")
    if pdf.is_dir():
        n_pdf = sum(1 for _ in pdf.glob("*.pdf"))
        parts.append(f"PDF: {n_pdf} файлов")

    chosen = Path(path).resolve()
    if chosen != root:
        parts.insert(0, f"Выбрано: {chosen} -> исправлено на корень output")

    return "   |   ".join(parts)
