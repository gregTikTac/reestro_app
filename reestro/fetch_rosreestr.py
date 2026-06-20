# -*- coding: utf-8 -*-
"""
Полуавтоматический загрузчик «Формы собственности» с lk.rosreestr.ru.

Открывает реальный браузер (Playwright/Chromium). Капчу и навигацию по сайту
делает человек, скрипт автоматически считывает строку «Форма собственности» с
открытой карточки объекта и сохраняет её:
  - в кэш output/cache/rosreestr/{kn}.json (как и авто-фетч парсера);
  - в справочник input/ownership_forms.csv (КН;Форма) — его потом подхватывает
    reestro_parser.py / rebuild_from_cache.py без обращений к Росреестру.

Если автоматически распознать форму не удалось — можно ввести её вручную в
консоли (значение тоже сохранится). Уже известные КН (в кэше/справочнике)
пропускаются.

ВНИМАНИЕ: Росреестр доступен только с российского IP. При включённом VPN с
зарубежным выходом сайт не открывается (см. README).

Запуск (примеры):
    python fetch_rosreestr.py -o output_test_first          # КН с правами из report.xlsx
    python fetch_rosreestr.py -o output --all               # все КН из report.xlsx
    python fetch_rosreestr.py --kn 77:01:0001001:1037       # один номер
    python fetch_rosreestr.py --kn-file kn_list.txt         # список из файла
"""
import argparse
import csv
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from reestro_parser import (
    BASE_DIR,
    CADASTRAL_RE,
    ROSREESTR_ONLINE_URL,
    load_existing_report,
    load_ownership_overrides,
    normalize_kn,
    save_ownership_cache,
    _str,
)

FORMS_DEFAULT = BASE_DIR / "input" / "ownership_forms.csv"


def extract_form_from_text(text: str) -> str:
    """Ищет значение строки «Форма собственности» в тексте карточки объекта."""
    if not text:
        return ""
    lines = [l.strip() for l in re.split(r"[\r\n]+", text)]
    for i, line in enumerate(lines):
        if re.search(r"форма\s+собственности", line, re.IGNORECASE):
            # значение в той же строке после ':' …
            m = re.search(r"форма\s+собственности\s*[:\-]?\s*(.+)", line, re.IGNORECASE)
            if m and m.group(1).strip():
                return m.group(1).strip()
            # … либо на следующей непустой строке
            for nxt in lines[i + 1:]:
                if nxt:
                    return nxt
    return ""


def append_to_forms_csv(path: Path, kn: str, form: str):
    """Добавляет/обновляет запись КН;Форма в справочнике (без дублей)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    rows: dict[str, str] = {}
    if path.exists():
        with open(path, encoding="utf-8-sig", newline="") as f:
            for r in csv.reader(f, delimiter=";"):
                if len(r) >= 2 and CADASTRAL_RE.search(r[0] or ""):
                    rows[normalize_kn(r[0])] = r[1]
    rows[normalize_kn(kn)] = form
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Кадастровый номер", "Форма собственности"])
        for k, v in rows.items():
            w.writerow([k, v])


def collect_kns(args) -> list[str]:
    if args.kn:
        return [normalize_kn(args.kn)]
    if args.kn_file:
        out = []
        with open(args.kn_file, encoding="utf-8-sig") as f:
            for line in f:
                m = CADASTRAL_RE.search(line)
                if m:
                    out.append(m.group(0))
        return out

    out_dir = Path(args.output)
    report = out_dir / "report.xlsx"
    if not report.exists():
        raise SystemExit(f"Нет report.xlsx: {report}. Укажите --kn или --kn-file.")
    _, rows = load_existing_report(report, out_dir / "pdf")
    kns: list[str] = []
    seen = set()
    for r in rows:
        if not CADASTRAL_RE.search(_str(r[1])):
            continue
        kn = normalize_kn(r[1])
        if kn in seen:
            continue
        # по умолчанию только объекты с правами (без прав форма = «данные отсутствуют»)
        has_right = _str(r[31]) and _str(r[31]) != "данные отсутствуют"
        if args.all or has_right:
            seen.add(kn)
            kns.append(kn)
    return kns


def try_autofill(page, kn: str) -> bool:
    """Best-effort: подставить КН в поле поиска. Капчу/поиск делает человек."""
    selectors = [
        "#query",
        "input[placeholder*='адастров']",
        "input[name*='adastral']",
        "input[id*='adastral']",
        "input[type='search']",
    ]
    for sel in selectors:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.fill(kn)
                return True
        except Exception:
            continue
    return False


def print_captcha_hint(page):
    """Подсказка по капче на странице поиска."""
    print("  Капча: введите символы с картинки в поле «Введите символы» (#captcha).")
    try:
        cap = page.query_selector("#captcha, input[name='captcha']")
        if cap:
            cap.scroll_into_view_if_needed()
    except Exception:
        pass


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("-o", "--output", default=str(BASE_DIR / "output"),
                   help="Папка результатов (report.xlsx + cache/rosreestr).")
    p.add_argument("--forms", default=str(FORMS_DEFAULT),
                   help="Справочник КН;Форма (по умолчанию input/ownership_forms.csv).")
    p.add_argument("--kn", default=None, help="Обработать один КН.")
    p.add_argument("--kn-file", default=None, help="Файл со списком КН (по одному в строке).")
    p.add_argument("--all", action="store_true",
                   help="Все КН из report.xlsx (а не только с зарегистрированными правами).")
    p.add_argument("--redo", action="store_true",
                   help="Переспрашивать даже уже известные КН.")
    args = p.parse_args()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise SystemExit("Нет Playwright. Установите:\n"
                         "  pip install playwright\n"
                         "  python -m playwright install chromium")

    out_dir = Path(args.output)
    rr_cache = out_dir / "cache" / "rosreestr"
    rr_cache.mkdir(parents=True, exist_ok=True)
    forms_path = Path(args.forms)

    known = load_ownership_overrides(forms_path)
    for jf in rr_cache.glob("*.json"):
        known.setdefault(normalize_kn(jf.stem), "known")

    kns = collect_kns(args)
    todo = [k for k in kns if args.redo or k not in known]
    print(f"Кадастровых номеров к обработке: {len(todo)} (всего в списке: {len(kns)})")
    if not todo:
        print("Все номера уже известны. Используйте --redo, чтобы переспросить.")
        return

    profile_dir = out_dir / ".pw_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as pw:
        # ignore_https_errors: на части сетей lk.rosreestr.ru отдаёт ERR_CERT_AUTHORITY_INVALID
        ctx = pw.chromium.launch_persistent_context(
            str(profile_dir), headless=False,
            viewport={"width": 1280, "height": 900},
            ignore_https_errors=True,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        for n, kn in enumerate(todo, 1):
            print("\n" + "=" * 60)
            print(f"[{n}/{len(todo)}] КН: {kn}")
            try:
                page.goto(ROSREESTR_ONLINE_URL, wait_until="domcontentloaded", timeout=60000)
            except Exception as e:
                print(f"  Не удалось открыть страницу: {e}")
                print("  Проверьте VPN/сеть (нужен российский IP).")
            page.wait_for_timeout(1500)
            if try_autofill(page, kn):
                print("  КН подставлен в поле «Введите адрес или кадастровый номер».")
            else:
                print(f"  Поле не найдено — введите КН вручную: {kn}")

            print_captcha_hint(page)
            print("  ДЕЙСТВИЯ В БРАУЗЕРЕ:")
            print("    1) введите капчу; 2) нажмите «Найти»; 3) откройте карточку объекта;")
            print("    4) найдите строку «Форма собственности» (раздел о правах).")
            input("  Когда «Форма собственности» видна на странице — нажмите Enter… ")

            form = ""
            try:
                text = page.inner_text("body")
                form = extract_form_from_text(text)
            except Exception as e:
                print(f"  Ошибка чтения страницы: {e}")

            if form:
                print(f"  Распознано: «{form}»")
                ok = input("  Сохранить это значение? [Enter=да / n=ввести вручную] ").strip().lower()
                if ok == "n":
                    form = ""
            if not form:
                form = input("  Введите форму собственности вручную (Enter — пропустить): ").strip()

            if not form:
                print("  Пропущено (значение не сохранено).")
                continue

            save_ownership_cache(rr_cache, kn, form, "rosreestr-manual")
            append_to_forms_csv(forms_path, kn, form)
            print(f"  Сохранено: {kn} → «{form}»")

        print("\nГотово. Теперь пересоберите отчёт из кэша (без платных запросов):")
        print(f"  python rebuild_from_cache.py -o {args.output} --ownership-forms {forms_path}")
        ctx.close()


if __name__ == "__main__":
    main()
