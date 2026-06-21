# -*- coding: utf-8 -*-
"""
Очистка output после нескольких прогонов парсера:
  - убрать дубликаты КН в report.xlsx (оставить один блок строк на объект);
  - удалить PDF, на которые нет ссылок в report (лишние / от прошлых запусков).

По умолчанию — пробный прогон (--dry-run). Для записи: --apply

Пример:
    python cleanup_output.py -o output
    python cleanup_output.py -o output --apply
    python cleanup_output.py -o output --apply --rebuild-from-cache -i TZ/Запрос.xlsx
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from reestro_parser import (
    BASE_DIR,
    InputRow,
    apply_input_to_row,
    build_xlsx_rows,
    extract_rights,
    generate_pdf,
    get_ownership_form,
    load_existing_report,
    load_input_index,
    load_object_encumbrances,
    load_ownership_overrides,
    new_extract_id,
    normalize_kn,
    patch_report_rights_columns,
    pdf_name_for_extract,
    save_api_cache,
    write_report_xlsx,
    _str,
)


def _block_key(row: list) -> tuple[str, str, str]:
    return (normalize_kn(row[1]), _str(row[32]), _str(row[34]))


def _parse_report_date(value: str) -> datetime | None:
    s = _str(value)
    for fmt in ("%d.%m.%Y", "%d.%m.%Y %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:19], fmt)
        except ValueError:
            continue
    return None


def _score_block(key: tuple[str, str, str], rows: list[list], pdf_dir: Path) -> tuple:
    kn, _eid, pdf_name = key
    pdf_path = pdf_dir / pdf_name if pdf_name.lower().endswith(".pdf") else None
    pdf_exists = bool(pdf_path and pdf_path.exists())
    dates = [_parse_report_date(r[33]) for r in rows]
    dates = [d for d in dates if d]
    latest = max(dates) if dates else datetime.min
    # PDF на диске > без PDF; при равенстве — более поздняя дата выписки
    return (1 if pdf_exists else 0, latest)


def dedupe_report_rows(rows: list[list], pdf_dir: Path) -> tuple[list[list], dict]:
    """
    Один кадастровый номер → один блок строк (несколько строк = несколько прав, это норма).
    Если КН встречается с разными PDF/выпиской — оставляем лучший блок.
    """
    no_kn: list[list] = []
    kn_blocks: OrderedDict[str, OrderedDict[tuple, list[list]]] = OrderedDict()

    for row in rows:
        kn = normalize_kn(row[1])
        if not kn:
            no_kn.append(row)
            continue
        key = _block_key(row)
        blocks = kn_blocks.setdefault(kn, OrderedDict())
        blocks.setdefault(key, []).append(row)

    kept: list[list] = []
    stats = {
        "input_rows": len(rows),
        "unique_kn": len(kn_blocks),
        "duplicate_kn": 0,
        "removed_rows": 0,
        "no_kn_rows": len(no_kn),
    }

    for kn, blocks in kn_blocks.items():
        keys = list(blocks.keys())
        if len(keys) > 1:
            stats["duplicate_kn"] += 1
            best = max(keys, key=lambda k: _score_block(k, blocks[k], pdf_dir))
            kept.extend(blocks[best])
            for k in keys:
                if k != best:
                    stats["removed_rows"] += len(blocks[k])
        else:
            kept.extend(blocks[keys[0]])

    # Строки без КН: убрать полные дубликаты (EXT + адрес)
    seen_no_kn: set[tuple[str, str]] = set()
    deduped_no_kn: list[list] = []
    for row in no_kn:
        sig = (_str(row[0]), _str(row[2]))
        if sig in seen_no_kn:
            stats["removed_rows"] += 1
            continue
        seen_no_kn.add(sig)
        deduped_no_kn.append(row)

    result = patch_report_rights_columns(deduped_no_kn + kept)
    stats["output_rows"] = len(result)
    return result, stats


def referenced_pdfs(rows: list[list]) -> set[str]:
    out: set[str] = set()
    for row in rows:
        name = _str(row[34])
        if name.lower().endswith(".pdf"):
            out.add(name)
    return out


def remove_orphan_pdfs(pdf_dir: Path, keep: set[str], *, apply: bool,
                     trash_dir: Path | None) -> dict:
    removed: list[str] = []
    kept = 0
    for path in sorted(pdf_dir.glob("*.pdf")):
        if path.name in keep:
            kept += 1
            continue
        removed.append(path.name)
        if apply:
            if trash_dir:
                trash_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(path), str(trash_dir / path.name))
            else:
                path.unlink()
    return {"pdf_kept": kept, "pdf_removed": len(removed), "removed_names": removed}


def _meta_by_kn(existing_rows: list[list]) -> dict[str, dict]:
    """Первая строка каждого КН: ext_number, адрес, выписка, PDF."""
    meta: dict[str, dict] = {}
    for r in existing_rows:
        kn = normalize_kn(r[1])
        if not kn or kn in meta:
            continue
        meta[kn] = {
            "ext_number": _str(r[0]),
            "address": _str(r[2]),
            "extract_id": _str(r[32]),
            "extract_date": _str(r[33]),
            "pdf": _str(r[34]),
        }
    return meta


def _pick_pdf_for_kn(kn: str, meta: dict, pdf_dir: Path) -> tuple[str, str, str]:
    """extract_id, extract_date, pdf_name — из report или новый UUID."""
    m = meta.get(kn, {})
    pdf = _str(m.get("pdf"))
    eid = _str(m.get("extract_id"))
    edate = _str(m.get("extract_date"))
    if pdf.lower().endswith(".pdf") and (pdf_dir / pdf).exists():
        return eid or pdf[:-4], edate or datetime.now().strftime("%d.%m.%Y"), pdf
    eid = new_extract_id()
    return eid, datetime.now().strftime("%d.%m.%Y"), pdf_name_for_extract(eid)


def rebuild_report_from_cache(out_dir: Path, input_paths: list[Path],
                              ownership_forms: Path | None,
                              existing_rows: list[list],
                              *, regen_pdf: bool) -> list[list]:
    """Пересборка report только из JSON-кэша (без дублей КН)."""
    cache_dir = out_dir / "cache" / "json"
    rr_cache_dir = out_dir / "cache" / "rosreestr"
    pdf_dir = out_dir / "pdf"
    rr_cache_dir.mkdir(parents=True, exist_ok=True)

    overrides = load_ownership_overrides(
        ownership_forms or (BASE_DIR / "input" / "ownership_forms.csv"))
    input_idx = load_input_index(input_paths)
    meta = _meta_by_kn(existing_rows)
    today = datetime.now().strftime("%d.%m.%Y")

    rows: list[list] = []
    for jf in sorted(cache_dir.glob("*.json")):
        with open(jf, encoding="utf-8") as f:
            data = json.load(f)
        info = data.get("info") or {}
        kn = normalize_kn(
            data.get("cadastralNumber") or info.get("cadastralNumber") or jf.stem)
        if not kn:
            continue

        cached_input = data.get("input") or {}
        row = InputRow()
        row.cadastral = kn
        m = meta.get(kn, {})
        apply_input_to_row(
            row,
            {"ext_number": m.get("ext_number"), "full_address": m.get("address")},
            cached_input,
            input_idx.get(kn),
        )
        if not row.full_address:
            row.full_address = _str(info.get("address"))

        if row.ext_number or row.full_address:
            save_api_cache(cache_dir, kn, info, None, input_row=row)

        extract_id, extract_date, pdf_name = _pick_pdf_for_kn(kn, meta, pdf_dir)
        if not extract_date:
            extract_date = today

        rights = extract_rights(info)
        own_form = get_ownership_form(kn, rr_cache_dir, overrides, fetch=False)
        encs = load_object_encumbrances(rr_cache_dir, kn)

        if regen_pdf or not (pdf_dir / pdf_name).exists():
            generate_pdf(info, rights, row, pdf_dir / pdf_name,
                         ownership_form=own_form, encumbrances_override=encs)

        rows.extend(build_xlsx_rows(
            row, info, rights, pdf_name, extract_id, extract_date,
            ownership_form=own_form))

    return patch_report_rights_columns(rows)


def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("-o", "--output", default=str(BASE_DIR / "output"),
                   help="Папка output (report.xlsx, pdf/, cache/).")
    p.add_argument("--apply", action="store_true",
                   help="Записать report и удалить/переместить лишние PDF.")
    p.add_argument("--trash-dir", default=None,
                   help="Переместить лишние PDF сюда вместо удаления.")
    p.add_argument("--rebuild-from-cache", action="store_true",
                   help="Сначала собрать report только из cache/json (без дублей КН).")
    p.add_argument("-i", "--input", action="append", default=None,
                   help="Входной XLSX/CSV для EXT_NUMBER (можно несколько раз).")
    p.add_argument("--ownership-forms", default=None)
    p.add_argument("--no-pdf", action="store_true",
                   help="При --rebuild-from-cache не пересоздавать PDF.")
    args = p.parse_args()

    out_dir = Path(args.output)
    report_path = out_dir / "report.xlsx"
    pdf_dir = out_dir / "pdf"
    apply = args.apply
    dry = not apply

    if not report_path.exists() and not args.rebuild_from_cache:
        raise SystemExit(f"Нет {report_path}. Укажите --rebuild-from-cache, если есть cache/json.")

    pdf_dir.mkdir(parents=True, exist_ok=True)

    existing_rows: list[list] = []
    if report_path.exists():
        _, existing_rows = load_existing_report(report_path, pdf_dir)

    if args.rebuild_from_cache:
        cache_dir = out_dir / "cache" / "json"
        if not cache_dir.exists() or not any(cache_dir.glob("*.json")):
            raise SystemExit(f"Нет JSON-кэша: {cache_dir}")
        input_paths = [Path(x) for x in args.input] if args.input else []
        if not input_paths:
            tz = BASE_DIR / "TZ" / "Запрос.xlsx"
            if tz.exists():
                input_paths = [tz]
        print(f"Пересборка report из кэша ({len(list(cache_dir.glob('*.json')))} JSON)...")
        rows = rebuild_report_from_cache(
            out_dir, input_paths, Path(args.ownership_forms) if args.ownership_forms else None,
            existing_rows,
            regen_pdf=not args.no_pdf and apply,
        )
        unique_kn = len({normalize_kn(r[1]) for r in rows if normalize_kn(r[1])})
        print(f"  Из кэша: {unique_kn} уникальных КН, {len(rows)} строк в report")
    else:
        rows, dedupe_stats = dedupe_report_rows(existing_rows, pdf_dir)
        print(f"report.xlsx: было {dedupe_stats['input_rows']} строк, "
              f"уникальных КН: {dedupe_stats['unique_kn']}, "
              f"КН с дублями: {dedupe_stats['duplicate_kn']}, "
              f"удалено строк: {dedupe_stats['removed_rows']}, "
              f"останется: {dedupe_stats['output_rows']}")

    keep_pdfs = referenced_pdfs(rows)
    pdf_stats = remove_orphan_pdfs(
        pdf_dir, keep_pdfs, apply=apply,
        trash_dir=Path(args.trash_dir) if args.trash_dir else None,
    )
    print(f"PDF: оставить {pdf_stats['pdf_kept']}, "
          f"лишних {pdf_stats['pdf_removed']}")

    if dry:
        print("\n[ПРОБНЫЙ ПРОГОН] Ничего не изменено. Добавьте --apply для записи.")
        if pdf_stats["pdf_removed"] and pdf_stats["pdf_removed"] <= 20:
            for name in pdf_stats["removed_names"]:
                print(f"  - {name}")
        elif pdf_stats["pdf_removed"]:
            for name in pdf_stats["removed_names"][:10]:
                print(f"  - {name}")
            print(f"  ... и ещё {pdf_stats['pdf_removed'] - 10}")
        return

    unique_kn = len({normalize_kn(r[1]) for r in rows if normalize_kn(r[1])})
    stats = {
        "total": unique_kn + sum(1 for r in rows if not normalize_kn(r[1])),
        "ok": unique_kn,
        "failed": 0,
        "skipped_no_kn": sum(1 for r in rows if not normalize_kn(r[1])),
        "skipped_already": 0,
    }
    write_report_xlsx(rows, stats, report_path)
    print(f"\nГотово: {report_path} ({len(rows)} строк, {unique_kn} КН)")
    print(f"PDF: удалено/перемещено {pdf_stats['pdf_removed']}, осталось {pdf_stats['pdf_kept']}")


if __name__ == "__main__":
    main()
