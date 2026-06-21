# -*- coding: utf-8 -*-
"""
Пересборка report.xlsx и PDF из локального JSON-кэша БЕЗ запросов к API.

Назначение: применить новые правила заполнения колонок («Вид собственности» =
форма собственности из Росреестра; «Доля» — пустая при отсутствии данных) к уже
обработанным объектам, по которым ответ API сохранён в output/cache/json.

Платный API Контур.Реестро НЕ вызывается — используются только кэш и
справочник форм собственности (override-файл + output/cache/rosreestr).

Запуск:
    python rebuild_from_cache.py -o output_test_first
    python rebuild_from_cache.py -o output --ownership-forms input/ownership_forms.csv
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from reestro_parser import (
    BASE_DIR,
    InputRow,
    INPUT_SKIP_NAMES,
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


def _meta_by_kn(existing_rows: list[list]) -> dict[str, dict]:
    """Первая строка каждого КН: ext_number, адрес, номер/дата выписки, имя PDF."""
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


def _collect_input_paths(args) -> list[Path]:
    if args.input:
        return [Path(p) for p in args.input]
    paths = []
    inp_dir = BASE_DIR / "input"
    if inp_dir.is_dir():
        for p in sorted(inp_dir.iterdir()):
            if p.suffix.lower() in (".csv", ".xlsx", ".xlsm"):
                if p.name.lower() not in INPUT_SKIP_NAMES:
                    paths.append(p)
    return paths


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("-o", "--output", default=str(BASE_DIR / "output_test_first"),
                   help="Папка результатов (с cache/json и report.xlsx).")
    p.add_argument(
        "-i", "--input", action="append", default=None,
        help="Входной CSV/XLSX для EXT_NUMBER и адреса (можно указать несколько раз). "
             "По умолчанию: все CSV/XLSX из папки input/.",
    )
    p.add_argument("--ownership-forms", default=None,
                   help="Справочник форм собственности (по умолчанию input/ownership_forms.csv).")
    p.add_argument("--no-pdf", action="store_true",
                   help="Не пересобирать PDF, только report.xlsx.")
    args = p.parse_args()

    out_dir = Path(args.output)
    cache_dir = out_dir / "cache" / "json"
    rr_cache_dir = out_dir / "cache" / "rosreestr"
    rr_cache_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = out_dir / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    report = out_dir / "report.xlsx"

    if not cache_dir.exists():
        raise SystemExit(f"Нет кэша: {cache_dir}")

    overrides_path = args.ownership_forms or (BASE_DIR / "input" / "ownership_forms.csv")
    overrides = load_ownership_overrides(overrides_path)
    print(f"Форма собственности: {len(overrides)} записей из справочника")

    input_paths = _collect_input_paths(args)
    input_idx = load_input_index(input_paths)
    if input_idx:
        print(f"EXT_NUMBER: загружено {len(input_idx)} КН из {len(input_paths)} входных файлов")

    _, existing_rows = load_existing_report(report, pdf_dir)
    meta = _meta_by_kn(existing_rows)
    today = datetime.now().strftime("%d.%m.%Y")

    new_rows: list[list] = []
    cache_kns: set[str] = set()
    rebuilt = 0

    for jf in sorted(cache_dir.glob("*.json")):
        with open(jf, encoding="utf-8") as f:
            data = json.load(f)
        info = data.get("info") or {}
        kn = normalize_kn(data.get("cadastralNumber") or info.get("cadastralNumber") or jf.stem)
        if not kn:
            continue
        cache_kns.add(kn)
        m = meta.get(kn, {})
        cached_input = data.get("input") or {}

        row = InputRow()
        row.cadastral = kn
        apply_input_to_row(
            row,
            {"ext_number": m.get("ext_number"), "full_address": m.get("address")},
            cached_input,
            input_idx.get(kn),
        )
        if not row.full_address:
            row.full_address = _str(info.get("address"))

        # Сохраняем EXT_NUMBER в кэш, чтобы не терялся при следующих пересборках
        if row.ext_number or row.full_address:
            save_api_cache(cache_dir, kn, info, None, input_row=row)

        extract_id = m.get("extract_id") or new_extract_id()
        extract_date = m.get("extract_date") or today
        pdf_name = m.get("pdf") or pdf_name_for_extract(extract_id)

        rights = extract_rights(info)
        own_form = get_ownership_form(kn, rr_cache_dir, overrides, fetch=False)
        encs = load_object_encumbrances(rr_cache_dir, kn)

        if not args.no_pdf:
            generate_pdf(info, rights, row, pdf_dir / pdf_name,
                         ownership_form=own_form, encumbrances_override=encs)

        new_rows.extend(build_xlsx_rows(
            row, info, rights, pdf_name, extract_id, extract_date,
            ownership_form=own_form))
        rebuilt += 1

    # Переносим только осмысленные строки, которых нет в кэше:
    #   - валидный КН (не из кэша) — реальный объект из прошлых прогонов;
    #   - либо строка без КН, но с адресом — пропущенный объект (нет кадастрового).
    from reestro_parser import CADASTRAL_RE

    def _keep(r) -> bool:
        if CADASTRAL_RE.search(_str(r[1])):
            return normalize_kn(r[1]) not in cache_kns
        return not _str(r[1]) and bool(_str(r[2]))

    kept = [r for r in existing_rows if _keep(r)]
    merged = patch_report_rights_columns(kept + new_rows)

    stats = {
        "total": len(cache_kns) + len({normalize_kn(r[1]) for r in kept if not normalize_kn(r[1])}),
        "ok": rebuilt,
        "failed": 0,
        "skipped_no_kn": sum(1 for r in kept if not normalize_kn(r[1])),
        "skipped_already": 0,
    }
    write_report_xlsx(merged, stats, report)
    print(f"Пересобрано объектов из кэша: {rebuilt}")
    print(f"report.xlsx обновлён: {report} ({len(merged)} строк), запросов к API: 0")


if __name__ == "__main__":
    main()
