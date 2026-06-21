# -*- coding: utf-8 -*-
"""
Предпросмотр: как форма собственности попадёт в report.xlsx и PDF.

Не ходит в Росреестр и не тратит API Контура — только читает локальные файлы:
  - output/cache/json/{kn}.json   — ответ API (права)
  - output/cache/rosreestr/{kn}.json — форма с lk.rosreestr.ru
  - input/ownership_forms.csv     — ручной справочник

Запуск:
  python verify_ownership_preview.py --kn 77:01:0001001:1037
  python verify_ownership_preview.py -o output_test_first
  python verify_ownership_preview.py -o output_test_first --kn 43:40:000233:782
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from reestro_parser import (
    BASE_DIR,
    InputRow,
    build_xlsx_rows,
    extract_rights,
    get_ownership_form,
    load_object_encumbrances,
    load_ownership_overrides,
    normalize_kn,
    _str,
)

COL_FIO = 28
COL_FORM = 29
COL_SHARE = 30
COL_RIGHT = 31


def preview_one(kn: str, out_dir: Path, overrides: dict) -> dict | None:
    kn = normalize_kn(kn)
    rr_cache = out_dir / "cache" / "rosreestr"
    json_cache = out_dir / "cache" / "json"
    jf = json_cache / f"{kn.replace(':', '_')}.json"
    # имя файла кэша может быть с подчёркиваниями
    if not jf.exists():
        for p in json_cache.glob("*.json"):
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
            if normalize_kn(d.get("cadastralNumber", "")) == kn:
                jf = p
                break
        else:
            return None

    with open(jf, encoding="utf-8") as f:
        data = json.load(f)
    info = data.get("info") or {}
    rights = extract_rights(info)
    own_form = get_ownership_form(kn, rr_cache, overrides, fetch=False)

    row = InputRow()
    row.cadastral = kn
    row.full_address = _str(info.get("address") or (data.get("input") or {}).get("full_address"))

    xrows = build_xlsx_rows(row, info, rights, "preview.pdf", "preview-id", "01.01.2026",
                            ownership_form=own_form)

    rr_file = rr_cache / f"{kn.replace(':', '_')}.json"
    if not rr_file.exists():
        for p in rr_cache.glob("*.json"):
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
            if normalize_kn(d.get("cadastralNumber", "")) == kn:
                rr_file = p
                break

    encs = load_object_encumbrances(rr_cache, kn)
    return {
        "kn": kn,
        "address": row.full_address,
        "json_cache": str(jf),
        "rosreestr_cache": str(rr_file) if rr_file.exists() else "(нет — запустите fetch_rosreestr.py)",
        "ownership_form": own_form or "(нет - будет «данные отсутствуют»)",
        "encumbrances": encs,
        "rights_count": len(rights),
        "rows": xrows,
    }


def print_preview(p: dict):
    print("=" * 60)
    print(f"КН:      {p['kn']}")
    print(f"Адрес:   {p['address'] or '—'}")
    print(f"JSON:    {p['json_cache']}")
    print(f"Росреestr: {p['rosreestr_cache']}")
    print(f"Форма (источник для PDF «Правообладатель» и кол. «Вид собственности»):")
    print(f"         {p['ownership_form']}")
    encs = p.get("encumbrances") or []
    if encs:
        print("Обременения (PDF поле 4 «Ограничение прав и обременение»):")
        for e in encs:
            print(f"         - {e}")
    else:
        print("Обременения: нет в кэше → PDF покажет «не зарегистрировано»")
    print(f"Прав в API: {p['rights_count']}")
    print("-" * 60)
    print("Строки в report.xlsx (как после «Пересобрать из кэша»):")
    print(f"  {'№':<3} {'ФИО (кол.29)':<28} {'Вид собств. (кол.30)':<28} {'Право (кол.32)'}")
    for i, r in enumerate(p["rows"], 1):
        fio = _str(r[COL_FIO])[:26]
        form = _str(r[COL_FORM])[:26]
        right = _str(r[COL_RIGHT])[:40]
        print(f"  {i:<3} {fio:<28} {form:<28} {right}")
    print()
    print("PDF раздел 2, поле «Правообладатель» = та же форма собственности.")
    print("Колонка «ФИО» в Excel = тип из API (rightOwnerType), часто пусто в открытых сведениях.")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--output", default=str(BASE_DIR / "output_test_first"))
    ap.add_argument("--kn", default=None, help="Один КН для проверки.")
    ap.add_argument("--forms", default=str(BASE_DIR / "input" / "ownership_forms.csv"))
    args = ap.parse_args()

    out = Path(args.output)
    overrides = load_ownership_overrides(args.forms)
    print(f"Папка: {out}")
    print(f"Справочник форм: {len(overrides)} записей\n")

    if args.kn:
        kns = [normalize_kn(args.kn)]
    else:
        jdir = out / "cache" / "json"
        if not jdir.is_dir():
            raise SystemExit(f"Нет кэша: {jdir}")
        kns = []
        for p in sorted(jdir.glob("*.json")):
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
            kn = normalize_kn(d.get("cadastralNumber", ""))
            if kn:
                kns.append(kn)

    if not kns:
        raise SystemExit("Нет КН для проверки.")

    for kn in kns:
        p = preview_one(kn, out, overrides)
        if not p:
            print(f"КН {kn}: нет JSON-кэша в {out / 'cache' / 'json'}")
            continue
        print_preview(p)


if __name__ == "__main__":
    main()
