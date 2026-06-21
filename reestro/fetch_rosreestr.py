# -*- coding: utf-8 -*-
"""
Загрузчик «Формы собственности» с lk.rosreestr.ru (режим кадастрового инженера).

Открывает реальный браузер (Playwright/Chromium) с постоянным профилем. Сценарий:
  1) ОДИН раз вы входите через Госуслуги (кадастровый инженер) — сессия
     сохраняется в профиле и переиспользуется при следующих запусках;
  2) дальше скрипт по каждому КН сам открывает «Справочную информацию online»,
     подставляет номер, запускает поиск и считывает строку «Форма собственности»;
  3) если на конкретном объекте всё же появляется капча — скрипт ждёт, пока вы
     её решите, и продолжает; форму всегда можно ввести вручную.

Найденная форма сохраняется:
  - в кэш <output>/cache/rosreestr/{kn}.json (его подхватывает GUI и
    rebuild_from_cache.py при «Пересобрать из кэша» — без запросов к API Контура);
  - в справочник input/ownership_forms.csv (КН;Форма).

ПРОВЕРКА КАПЧИ: запустите с одним КН (--kn ...). Если после входа кадастровым
инженером поиск проходит без капчи — значит, можно гнать пачкой в авто-режиме.

ВНИМАНИЕ: Росреестр доступен только с российского IP. При зарубежном VPN сайт
не открывается. Открытые сведения содержат ФОРМУ собственности, но НЕ ФИО физлиц.

Запуск (примеры):
    python fetch_rosreestr.py --kn 77:01:0001001:1037        # проверка одного КН
    python fetch_rosreestr.py -o output_test_first --auto     # пачкой, авто-режим
    python fetch_rosreestr.py -o output --all --auto          # все КН из report.xlsx
    python fetch_rosreestr.py --kn-file kn_list.txt --auto    # список из файла
"""
import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Callable

sys.path.insert(0, str(Path(__file__).resolve().parent))

from reestro_parser import (
    BASE_DIR,
    CADASTRAL_RE,
    ROSREESTR_ONLINE_URL,
    extract_encumbrances_from_text,
    load_existing_report,
    load_ownership_overrides,
    normalize_kn,
    save_ownership_cache,
    _str,
)

FORMS_DEFAULT = BASE_DIR / "input" / "ownership_forms.csv"

# Типовые формы собственности (для распознавания на странице Росreestr).
KNOWN_OWNERSHIP_FORMS = (
    "Частная собственность",
    "Муниципальная собственность",
    "Государственная собственность",
    "Государственная федеральная",
    "Государственная субъекта Российской Федерации",
    "Государственная субъекта РФ",
    "Общая долевая собственность",
    "Общая совместная собственность",
)

# Строки интерфейса — не путать со значением поля.
_UI_SKIP_RE = re.compile(
    r"^(сформировать|найти|поиск|главная|войти|госуслуги|капча|меню|"
    r"справочная информация|открытая служба|copyright|©|"
    r"введите символы|кадастровый номер\s*$|адрес\s*$)",
    re.IGNORECASE,
)


def _is_ui_noise(line: str) -> bool:
    s = (line or "").strip()
    if not s or len(s) < 3:
        return True
    if _UI_SKIP_RE.search(s):
        return True
    if s.lower() in ("да", "нет", "ok", "—", "-"):
        return True
    return False


def extract_form_from_text(text: str) -> str:
    """Ищет форму собственности в тексте карточки / справки Росreestr."""
    if not text:
        return ""
    lines = [l.strip() for l in re.split(r"[\r\n]+", text) if l.strip()]

    # 1) Метка «Форма собственности» → значение в той же или следующей строке.
    for i, line in enumerate(lines):
        if not re.search(r"форма\s+собственности", line, re.IGNORECASE):
            continue
        m = re.search(r"форма\s+собственности\s*[:\-]?\s*(.+)", line, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if val and not _is_ui_noise(val):
                return val
        for nxt in lines[i + 1: i + 4]:
            if _is_ui_noise(nxt):
                continue
            if re.search(r"форма\s+собственности", nxt, re.IGNORECASE):
                continue
            return nxt

    # 2) Метка «Правообладатель» (в справке иногда дублирует форму).
    for i, line in enumerate(lines):
        if not re.search(r"правообладател", line, re.IGNORECASE):
            continue
        m = re.search(r"правообладател\w*\s*[:\-]?\s*(.+)", line, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if val and not _is_ui_noise(val) and "данные отсутствуют" not in val.lower():
                for known in KNOWN_OWNERSHIP_FORMS:
                    if known.lower() in val.lower():
                        return known
                if "собственность" in val.lower():
                    return val
        for nxt in lines[i + 1: i + 3]:
            if _is_ui_noise(nxt):
                continue
            for known in KNOWN_OWNERSHIP_FORMS:
                if known.lower() in nxt.lower():
                    return known

    # 3) Известные формулировки где угодно в тексте (один однозначный матч).
    found: list[str] = []
    low = text.lower()
    for known in KNOWN_OWNERSHIP_FORMS:
        if known.lower() in low:
            found.append(known)
    if len(found) == 1:
        return found[0]
    if found:
        # предпочитаем «Частная/Муниципальная/… собственность» целиком
        for f in found:
            if f.endswith("собственность"):
                return f
        return found[0]
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


SEARCH_INPUT_SELECTORS = (
    "#query",
    "input[name='query']",
    "input[placeholder*='кадастровый номер объекта']",
    "input[placeholder*='Введите адрес']",
    "input[name='cadastralNumber']",
    "input[id*='cadastral']",
    "input[name*='cadastral']",
    "input[placeholder*='адастров']",
    "input[placeholder*='Кадастров']",
    "input[type='search']",
)

# Вид объекта в форме ЛК — в интерфейсе часто множественное число («Здания»).
LK_OBJECT_TYPES = (
    "Здание",
    "Помещение",
    "Земельный участок",
)
LK_TYPE_UI_LABELS: dict[str, tuple[str, ...]] = {
    "Здание": ("Здания", "Здание"),
    "Помещение": ("Помещения", "Помещение"),
    "Земельный участок": ("Земельные участки", "Земельный участок"),
}

# Пауза между объектами (Росreestr режет по лимиту обращений).
PAUSE_BETWEEN_OBJECTS_SEC = 5.0
RATE_LIMIT_WAIT_SEC = 90


class CollectionAborted(Exception):
    """Прерывание сбора: cancel — остановка, skip — следующий объект."""

    def __init__(self, reason: str = "cancel"):
        self.reason = reason
        super().__init__(reason)


def _check_abort(should_cancel=None, should_skip=None, *, allow_skip: bool = True) -> None:
    if should_cancel and should_cancel():
        raise CollectionAborted("cancel")
    if allow_skip and should_skip and should_skip():
        raise CollectionAborted("skip")


def _interruptible_wait(page, ms: int, should_cancel=None, should_skip=None,
                        *, step: int = 300, allow_skip: bool = True) -> None:
    remaining = int(ms)
    while remaining > 0:
        _check_abort(should_cancel, should_skip, allow_skip=allow_skip)
        chunk = min(step, remaining)
        page.wait_for_timeout(chunk)
        remaining -= chunk

# Публичная справочная — расширенный перебор.
ROSREESTR_OBJECT_TYPES = LK_OBJECT_TYPES + (
    "Сооружение",
    "Объект незавершенного строительства",
)


def rosreestr_object_type_from_info(info: dict | None) -> str:
    """Подбирает «Вид объекта» для формы ЛК по данным API Контура."""
    if not info:
        return ""
    ct = _str(info.get("constructionType") or info.get("objectType") or "")
    low = ct.lower()
    if "помещен" in low:
        return "Помещение"
    if "здан" in low or "дом" in low:
        return "Здание"
    if "земель" in low or "участ" in low:
        return "Земельный участок"
    if "сооруж" in low:
        return "Сооружение"
    if "незаверш" in low:
        return "Объект незавершенного строительства"
    rg = _str(info.get("realEstateGroup") or "").lower()
    if "земел" in rg:
        return "Земельный участок"
    if "капиталь" in rg or "строитель" in rg:
        return "Здание"
    return ""


def load_object_type_hint(out_dir: Path, kn: str) -> str:
    """Читает подсказку «Вид объекта» из cache/json/{kn}.json."""
    cache = out_dir / "cache" / "json" / f"{normalize_kn(kn).replace(':', '_')}.json"
    if not cache.is_file():
        return ""
    try:
        import json
        data = json.loads(cache.read_text(encoding="utf-8"))
        return rosreestr_object_type_from_info(data.get("info") or data)
    except Exception:
        return ""


def is_lk_personal_form(page) -> bool:
    """Форма ЛК после входа (кнопка «НАЙТИ», поле «Вид объекта»)."""
    try:
        text = _page_text(page)
        if re.search(r"вид\s+объекта", text, re.I) and re.search(
                r"найти|адрес\s+или\s+кадастров", text, re.I):
            return True
        if page.get_by_text("Вид объекта", exact=False).count() > 0:
            return page.get_by_role(
                "button", name=re.compile(r"^\s*найти\s*$", re.I)).count() > 0
    except Exception:
        pass
    return False


def pick_rosreestr_page(ctx):
    """Активная вкладка lk.rosreestr.ru (после Госуслуг вход часто в другой вкладке)."""
    pages = list(ctx.pages)
    for p in reversed(pages):
        url = (p.url or "").lower()
        if "rosreestr.ru" in url and "gosuslugi" not in url and "esia." not in url:
            try:
                p.bring_to_front()
            except Exception:
                pass
            return p
    return pages[-1] if pages else None


def is_auth_page(page) -> bool:
    """True, если открыта страница входа (Госуслуги / OAuth), а не форма поиска."""
    url = (page.url or "").lower()
    if any(x in url for x in ("esia.", "gosuslugi", "/idp/", "oauth", "login")):
        return True
    low = _page_text(page).lower()
    if "gosuslugi" in url or "esia." in url:
        return True
    if "войти через госуслуги" in low and not page.query_selector("#query"):
        return True
    return False


def wait_for_search_form(page, timeout_ms: int = 30000,
                         should_cancel=None, should_skip=None) -> bool:
    """
    Ждёт форму поиска: публичную (#query) или ЛК («Вид объекта» + поле КН).
    React-страница рисует поля через 2–5 с после domcontentloaded.
    """
    if timeout_ms <= 0:
        return False
    step = 500
    waited = 0
    check_js = """() => {
        const q = document.querySelector(
            '#query, input[name="query"], input[placeholder*="кадастров"]');
        if (q && q.offsetParent !== null) return true;
        const t = document.body && document.body.innerText || '';
        return t.includes('Вид объекта') &&
            (t.includes('кадастровый номер') || t.includes('НАЙТИ'));
    }"""
    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip, allow_skip=False)
        try:
            if page.evaluate(check_js):
                return True
        except Exception:
            pass
        try:
            el = page.query_selector(", ".join(SEARCH_INPUT_SELECTORS))
            if el and el.is_visible():
                return True
        except Exception:
            pass
        _interruptible_wait(page, step, should_cancel, should_skip, allow_skip=False)
        waited += step
    return False


def object_type_candidates(hint: str = "", *, lk: bool = False) -> list[str]:
    """Список типов объекта для перебора (сначала подсказка из кэша API)."""
    pool = LK_OBJECT_TYPES if lk else ROSREESTR_OBJECT_TYPES
    out: list[str] = []
    if hint:
        out.append(hint)
    for t in pool:
        if t not in out:
            out.append(t)
    return out


def _search_form_root(page):
    """Центральная форма поиска (не боковое меню и не чужие react-select)."""
    try:
        for getter in (
            lambda: page.locator("main, [role='main']").filter(
                has=page.get_by_text(re.compile(
                    r"Справочная информация.*online|Вид объекта", re.I))
            ).first,
            lambda: page.locator("div").filter(
                has=page.get_by_text(
                    re.compile(r"Справочная информация по объектам", re.I))
            ).filter(
                has=page.get_by_text("Вид объекта", exact=False)
            ).filter(
                has=page.get_by_role("button", name=re.compile(r"найти", re.I))
            ).first,
        ):
            root = getter()
            if root.count() and root.is_visible(timeout=1200):
                return root
    except Exception:
        pass
    try:
        kn = page.get_by_placeholder(re.compile(
            r"Введите адрес или кадастров", re.I)).first
        if kn.is_visible(timeout=1500):
            return kn.locator(
                "xpath=ancestor::form | ancestor::div[contains(@class,'content')][1]"
            ).first
    except Exception:
        pass
    return page.locator("body")


def _dismiss_dropdowns(page) -> None:
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(250)
    except Exception:
        pass


def has_rate_limit(page) -> bool:
    """Росreestr: «Превышен лимит обращений, попробуйте позже»."""
    low = _page_text(page).lower()
    return "превышен лимит" in low or "лимит обращений" in low


def _object_type_is_set(page) -> bool:
    """True, если в «Вид объекта» выбрано значение (не placeholder)."""
    root = _search_form_root(page)
    try:
        box = root.locator("div").filter(
            has=page.get_by_text("Вид объекта", exact=False)).first
        t = box.inner_text(timeout=2000)
        low = (t or "").lower()
        return (bool(t) and "выберите значение" not in low
                and "не найдено" not in low)
    except Exception:
        return False


def _open_object_type_dropdown(page) -> bool:
    """Открывает выпадающий список «Вид объекта» в центральной форме."""
    _dismiss_dropdowns(page)
    root = _search_form_root(page)
    row = root.locator("div").filter(
        has=page.get_by_text("Вид объекта", exact=False)).first
    open_attempts = [
        lambda: row.locator("[class*='control'], [class*='Control']").first,
        lambda: row.get_by_text("Выберите значение из справочника", exact=False).first,
        lambda: row.locator("[class*='select']").locator(
            "[class*='control'], [class*='Control']").first,
    ]
    for get_el in open_attempts:
        try:
            el = get_el()
            if el.is_visible(timeout=2000):
                el.scroll_into_view_if_needed()
                el.click(force=True, timeout=4000)
                page.wait_for_timeout(600)
                return True
        except Exception:
            continue
    return False


def _click_dropdown_option(page, label: str) -> bool:
    """Клик по пункту открытого выпадающего списка (только выбор, без ввода текста)."""
    menu_selectors = (
        "[class*='menu'][class*='select']",
        "[id*='listbox']",
        "[class*='MenuList']",
    )
    for msel in menu_selectors:
        try:
            menu = page.locator(msel).filter(
                has=page.locator("[class*='option']")).last
            if not menu.is_visible(timeout=1500):
                continue
            opt = menu.locator("[class*='option']").filter(
                has_text=re.compile(rf"^{re.escape(label)}$", re.I)).first
            if opt.is_visible(timeout=1500):
                opt.click(force=True, timeout=3000)
                page.wait_for_timeout(400)
                return True
        except Exception:
            continue
    try:
        opt = page.get_by_role("option", name=re.compile(
            rf"^{re.escape(label)}$", re.I)).first
        if opt.is_visible(timeout=1500):
            opt.click(force=True, timeout=3000)
            page.wait_for_timeout(400)
            return True
    except Exception:
        pass
    return False


def _select_lk_object_type(page, canonical: str) -> bool:
    """
    Поле 2: «Вид объекта». Пробуем подписи из ЛК (Здания / Здание и т.д.).
    Поле 1 «тип поиска» не меняем.
    """
    if not canonical:
        return False
    labels = LK_TYPE_UI_LABELS.get(canonical, (canonical,))
    for label in labels:
        if _object_type_is_set(page):
            _dismiss_dropdowns(page)
        if not _open_object_type_dropdown(page):
            continue
        if _click_dropdown_option(page, label) and _object_type_is_set(page):
            return True
    return False


def _fill_kn_input(page, kn: str) -> bool:
    """Поле 3: «Адрес или кадастровый номер» — ввод КН."""
    _dismiss_dropdowns(page)
    kn_norm = normalize_kn(kn)
    root = _search_form_root(page)

    attempts = [
        # ЛК: текстовое поле под меткой (не react-select).
        lambda: root.locator("div").filter(
            has=page.get_by_text(re.compile(r"Адрес или кадастровый номер", re.I))
        ).locator(
            "input[placeholder*='кадастров'], input[type='text']:not([id*='react-select'])"
        ).first,
        lambda: root.get_by_placeholder(re.compile(
            r"Введите адрес или кадастров", re.I)).first,
        lambda: root.locator(
            "input[placeholder*='кадастровый номер объекта']:not([id*='react-select'])"
        ).first,
        lambda: root.locator(
            "input[type='text']:not([id*='react-select'])"
        ).filter(has=page.locator("[placeholder*='кадастров']")).first,
    ]
    for sel in ("#query", "input[name='query']"):
        attempts.append(lambda s=sel: page.locator(s).first)

    for get_el in attempts:
        try:
            el = get_el()
            if el.count() == 0 or not el.is_visible(timeout=1200):
                continue
            el.scroll_into_view_if_needed()
            el.click(force=True, timeout=3000)
            el.press("Control+a")
            el.fill(kn)
            page.wait_for_timeout(300)
            val = (el.input_value() or "").strip()
            if normalize_kn(val) == kn_norm or kn_norm.replace(":", "") in val.replace(":", ""):
                return True
            el.fill("")
            el.type(kn, delay=20)
            page.wait_for_timeout(200)
            val = (el.input_value() or "").strip()
            return normalize_kn(val) == kn_norm or kn_norm.replace(":", "") in val.replace(":", "")
        except Exception:
            continue
    return False


def _fill_lk_fields(page, kn: str, object_type: str, on_log=None) -> bool:
    """Заполняет форму ЛК: поле 2 (вид) + поле 3 (КН). Поле 1 не меняем."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    log(f"  поле 2 — вид объекта: «{object_type}»")
    if not _select_lk_object_type(page, object_type):
        log(f"  не удалось выбрать «{object_type}» в выпадающем списке")
        return False

    log(f"  поле 3 — кадастровый номер: {kn}")
    if not _fill_kn_input(page, kn):
        log("  не удалось ввести КН в поле «Адрес или кадастровый номер»")
        return False

    # Кнопка «НАЙТИ» должна стать активной.
    try:
        root = _search_form_root(page)
        btn = root.get_by_role("button", name=re.compile(r"^\s*найти\s*$", re.I)).first
        for _ in range(15):
            if btn.is_visible() and btn.is_enabled():
                log("  форма заполнена, кнопка «Найти» активна")
                return True
            page.wait_for_timeout(200)
    except Exception:
        pass
    log("  поля заполнены (кнопка «Найти» пока не активна — проверьте в браузере)")
    return True


def fill_search_form(page, kn: str, *, object_type: str = "",
                     on_log=None, lk_type: str | None = None) -> bool:
    """
    Заполняет форму поиска (публичную или ЛК).
    lk_type — один конкретный вид объекта для ЛК; иначе перебор candidates.
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if not wait_for_search_form(page, timeout_ms=12000):
        return False

    if is_lk_personal_form(page):
        types = [lk_type] if lk_type else object_type_candidates(object_type, lk=True)
        for ot in types:
            if not ot:
                continue
            if _fill_lk_fields(page, kn, ot, on_log):
                return True
        return False

    # Публичная справочная (без входа в ЛК).
    for label in ("кадастровому номеру", "Кадастровый номер"):
        try:
            loc = page.get_by_text(label, exact=False).first
            if loc.is_visible(timeout=800):
                loc.click(timeout=2000)
                page.wait_for_timeout(400)
                break
        except Exception:
            continue
    return _fill_kn_input(page, kn)


def try_autofill(page, kn: str, *, object_type: str = "",
                 on_log=None) -> bool:
    """Подставить КН (и вид объекта в ЛК)."""
    return fill_search_form(page, kn, object_type=object_type, on_log=on_log)


def is_on_reference_page(page) -> bool:
    """True, если открыта «Справочная информация online» с формой поиска."""
    url = (page.url or "").lower()
    if "real-estate-objects-online" not in url:
        return False
    return wait_for_search_form(page, timeout_ms=2000)


def navigate_to_reference_online(page, *, on_log=None) -> bool:
    """
    Переходит на «Справочная информация online» из любой страницы ЛК
    (напр. «Мои заявки» после выбора пользователя).
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if is_on_reference_page(page):
        return True

    url = (page.url or "").lower()
    if is_auth_page(page):
        return False

    if "rosreestr.ru" in url and "real-estate-objects-online" not in url:
        log("Перехожу на «Справочную информацию online»…")

    try:
        goto_rosreestr_page(page, timeout=60000)
        page.wait_for_timeout(1200)
        if wait_for_search_form(page, timeout_ms=25000):
            log("Справочная информация online открыта.")
            return True
    except Exception as exc:  # noqa: BLE001
        log(f"  переход по адресу: {exc}")

    for pattern in (
        r"Справочная\s+информация\s+по\s+объектам",
        r"Справочная\s+информация",
    ):
        try:
            link = page.get_by_role("link", name=re.compile(pattern, re.I)).first
            if link.is_visible(timeout=1500):
                link.click(timeout=5000)
                page.wait_for_timeout(2000)
                if wait_for_search_form(page, timeout_ms=20000):
                    log("Справочная информация online открыта (меню).")
                    return True
        except Exception:
            continue

    try:
        page.locator("a[href*='real-estate-objects-online']").first.click(timeout=4000)
        page.wait_for_timeout(2000)
        if wait_for_search_form(page, timeout_ms=20000):
            log("Справочная информация online открыта (ссылка).")
            return True
    except Exception:
        pass

    return False


def ensure_search_page(page, *, on_log=None, timeout_ms: int = 45000) -> bool:
    """Открывает справочную и ждёт форму поиска; не продолжать без неё."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if is_on_reference_page(page):
        return True

    if is_auth_page(page):
        log("  сейчас открыта страница входа — завершите авторизацию в браузере")
        return False

    if navigate_to_reference_online(page, on_log=on_log):
        return True

    if wait_for_search_form(page, timeout_ms=timeout_ms):
        return True

    if is_auth_page(page):
        log("  открылась страница входа — дождитесь возврата на Росreestr")
    else:
        log("  форма поиска не появилась (нет поля «Адрес или кадастровый номер»)")
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


CAPTCHA_SELECTORS = [
    "#captcha",
    "input[name='captcha']",
    "img[src*='captcha']",
    "[class*='captcha']",
]

SEARCH_BUTTON_SELECTORS = [
    "button:has-text('НАЙТИ')",
    "button:has-text('Найти')",
    "button:has-text('Сформировать запрос')",
    "input[type='submit'][value*='Сформировать']",
    "input[type='submit'][value*='Найти']",
    "button:has-text('Сформировать')",
    "a:has-text('Сформировать запрос')",
    "button:has-text('Поиск')",
    "button[type='submit']",
    "[class*='search'] button",
]


def has_captcha(page) -> bool:
    """True, если на странице видна капча."""
    for sel in CAPTCHA_SELECTORS:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                return True
        except Exception:
            continue
    return False


def click_search(page, should_cancel=None, should_skip=None) -> bool:
    """Best-effort: «НАЙТИ» (ЛК) или «Сформировать запрос» (публичная)."""
    root = _search_form_root(page)
    try:
        btn = root.get_by_role("button", name=re.compile(r"^\s*найти\s*$", re.I)).first
        for _ in range(25):
            _check_abort(should_cancel, should_skip)
            if btn.is_visible() and btn.is_enabled():
                btn.scroll_into_view_if_needed()
                btn.click()
                return True
            _interruptible_wait(page, 200, should_cancel, should_skip)
    except CollectionAborted:
        raise
    except Exception:
        pass
    for sel in SEARCH_BUTTON_SELECTORS:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible() and el.is_enabled():
                el.scroll_into_view_if_needed()
                el.click()
                return True
        except Exception:
            continue
    try:
        btn = page.get_by_role("button", name=re.compile(r"сформировать|найти", re.I)).first
        if btn.is_visible(timeout=1000) and btn.is_enabled():
            btn.click()
            return True
    except Exception:
        pass
    try:
        page.keyboard.press("Enter")
        return True
    except Exception:
        return False


def _page_text(page) -> str:
    """Текст страницы + содержимое видимых фреймов (на случай вложенной вёрстки)."""
    parts: list[str] = []
    try:
        parts.append(page.inner_text("body"))
    except Exception:
        pass
    try:
        for frame in page.frames:
            if frame == page.main_frame:
                continue
            try:
                parts.append(frame.inner_text("body"))
            except Exception:
                continue
    except Exception:
        pass
    return "\n".join(parts)


def _kn_in_text(text: str, kn: str) -> bool:
    a = normalize_kn(kn).replace(":", "").replace(" ", "")
    b = (text or "").replace(":", "").replace(" ", "")
    return bool(a) and a in b


def wait_for_results(page, kn: str, timeout_ms: int = 35000,
                     should_cancel=None, should_skip=None) -> bool:
    """Ждёт таблицу результатов или карточку объекта после «Найти»."""
    step = 700
    waited = 0
    kn_norm = normalize_kn(kn)
    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip)
        if has_captcha(page):
            return False
        if has_rate_limit(page):
            return False
        text = _page_text(page)
        low = text.lower()
        if re.search(r"найдено\s+результатов:\s*0", low):
            return False
        m = re.search(r"найдено\s+результатов:\s*(\d+)", low)
        if m and int(m.group(1)) >= 1 and _kn_in_text(text, kn_norm):
            return True
        if extract_form_from_text(text):
            return True
        if _kn_in_text(text, kn_norm) and re.search(
                r"форма\s+собственности|кадастровая\s+стоимость|"
                r"сведения\s+о\s+правах|обременени|справка\s+об\s+объект",
                text, re.IGNORECASE):
            return True
        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step
    return False


def open_object_card_from_results(page, kn: str, *, on_log=None,
                                  should_cancel=None, should_skip=None) -> bool:
    """Клик по КН в таблице «Найдено результатов» → справка об объекте."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if extract_form_from_text(_page_text(page)):
        return True

    kn_norm = normalize_kn(kn)
    kn_parts = (kn_norm, kn_norm.replace(":", "-"))
    root = _search_form_root(page)

    for part in kn_parts:
        for get_link in (
            lambda p=part: root.get_by_role("link", name=re.compile(
                re.escape(p))).first,
            lambda p=part: root.locator("table a, tbody a, [class*='table'] a").filter(
                has_text=re.compile(re.escape(p))).first,
            lambda p=part: page.locator("table a, tbody a").filter(
                has_text=re.compile(re.escape(p))).first,
        ):
            try:
                _check_abort(should_cancel, should_skip)
                link = get_link()
                if link.is_visible(timeout=2500):
                    log(f"  открываю справку по ссылке {kn_norm}")
                    link.click(timeout=5000)
                    _interruptible_wait(page, 2000, should_cancel, should_skip)
                    return True
            except CollectionAborted:
                raise
            except Exception:
                continue

    click_result_if_needed(page, kn_norm)
    _interruptible_wait(page, 1500, should_cancel, should_skip)
    return _kn_in_text(_page_text(page), kn_norm)


def click_result_if_needed(page, kn: str) -> None:
    """Если после поиска остались только ссылки — клик по строке с КН."""
    if extract_form_from_text(_page_text(page)):
        return
    kn_esc = normalize_kn(kn)
    for part in (kn_esc, kn_esc.replace(":", "-")):
        try:
            loc = page.get_by_text(part, exact=False).first
            if loc.is_visible(timeout=1500):
                loc.click(timeout=3000)
                page.wait_for_timeout(1200)
                return
        except Exception:
            continue


def read_form_with_wait(page, timeout_ms: int = 35000,
                        should_cancel=None, should_skip=None) -> str:
    """Ждёт появления строки «Форма собственности» и возвращает значение."""
    step = 700
    waited = 0
    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip)
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight / 2)")
        except Exception:
            pass
        form = extract_form_from_text(_page_text(page))
        if form:
            return form
        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step
    return ""


def read_page_data(page) -> tuple[str, list[str]]:
    """Читает (форма собственности, обременения) с текущей страницы."""
    text = _page_text(page)
    return extract_form_from_text(text), extract_encumbrances_from_text(text)


def launch_rosreestr_context(pw, profile_dir: Path):
    """
    Запускает браузер для lk.rosreestr.ru.

    Сайт часто блокирует встроенный Chromium Playwright (ERR_CONNECTION_RESET,
    белый экран). Поэтому сначала пробуем системный Edge, затем Chrome,
    и только потом — вшитый/скачанный Chromium.
    """
    profile = str(profile_dir)
    base_kw = dict(
        headless=False,
        viewport={"width": 1280, "height": 900},
        ignore_https_errors=True,
    )
    last_err: Exception | None = None
    for channel in ("msedge", "chrome", None):
        kw = dict(base_kw)
        label = "Microsoft Edge" if channel == "msedge" else (
            "Google Chrome" if channel == "chrome" else "Chromium (Playwright)")
        if channel:
            kw["channel"] = channel
        try:
            ctx = pw.chromium.launch_persistent_context(profile, **kw)
            return ctx, label
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            continue
    raise RuntimeError(
        "Не удалось запустить браузер для Росreestr. "
        "Установите Microsoft Edge или Google Chrome, либо выполните: "
        "python -m playwright install chromium"
    ) from last_err


def goto_rosreestr_page(page, *, timeout: int = 60000, retries: int = 3) -> None:
    """Открывает страницу справочной информации; повтор при сбросе соединения."""
    last_err: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            page.goto(ROSREESTR_ONLINE_URL, wait_until="domcontentloaded",
                      timeout=timeout)
            return
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            msg = str(exc)
            if attempt < retries and any(x in msg for x in (
                    "ERR_CONNECTION_RESET", "ERR_ABORTED", "Timeout",
                    "interrupted by another navigation")):
                page.wait_for_timeout(1500 * attempt)
                continue
            raise
    if last_err:
        raise last_err


def wait_for_captcha(page, timeout_ms: int = 120000, on_log=None,
                     should_cancel=None, should_skip=None) -> bool:
    """Ждёт, пока оператор введёт капчу в браузере (если она есть)."""
    if not has_captcha(page):
        return True
    if on_log:
        on_log("  капча — введите символы с картинки в браузере")
    step = 800
    waited = 0
    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip, allow_skip=False)
        if not has_captcha(page):
            return True
        try:
            cap = page.query_selector("#captcha, input[name='captcha']")
            if cap and len((cap.input_value() or "").strip()) >= 4:
                return True
        except Exception:
            pass
        _interruptible_wait(page, step, should_cancel, should_skip, allow_skip=False)
        waited += step
    return False


def auto_fetch_one(page, kn: str, *, auto: bool,
                   on_log=None, wait_captcha: bool = True,
                   object_type: str = "",
                   should_cancel: Callable[[], bool] | None = None,
                   should_skip: Callable[[], bool] | None = None,
                   assisted: bool = False,
                   wait_assist: Callable[[str], bool] | None = None) -> tuple[str, list[str]]:
    """
    Открывает справочную, подставляет КН, запускает поиск и читает данные.
    Возвращает (форма, список обременений).
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if assisted:
        log("  проверочный режим: КН подставляется автоматически")
        if object_type:
            log(f"  подсказка — вид объекта: «{object_type}», КН: {kn}")
        else:
            log(f"  подсказка — КН: {kn}")
        if not ensure_search_page(page, on_log=log, timeout_ms=45000):
            if is_auth_page(page):
                log("  нужен вход в ЛК — завершите вход и откройте справочную online")
            else:
                log("  откройте «Справочная информация online» в браузере")
            return "", []
        if wait_for_search_form(page, timeout_ms=8000,
                                should_cancel=should_cancel,
                                should_skip=should_skip):
            if _fill_kn_input(page, kn):
                log(f"  поле 3 — кадастровый номер подставлен: {kn}")
            else:
                log(f"  не удалось автоматически ввести КН — введите вручную: {kn}")
        else:
            log("  форма поиска не готова — введите КН вручную")
        if wait_assist is not None:
            log("  выберите «Вид объекта», нажмите «Найти», откройте справку")
            if not wait_assist(kn):
                raise CollectionAborted("cancel")
        _check_abort(should_cancel, should_skip)
        form = read_form_with_wait(page, 15000, should_cancel, should_skip)
        if not form:
            form, encs = read_page_data(page)
        else:
            encs = extract_encumbrances_from_text(_page_text(page))
        if form:
            log(f"  прочитано: «{form}»")
        else:
            log("  форма не распознана на текущей странице")
        return form, encs

    if not ensure_search_page(page, on_log=log, timeout_ms=45000):
        if is_auth_page(page):
            log("  нужен вход в ЛК — завершите вход и откройте справочную online")
        else:
            log("  поле КН не найдено — откройте в браузере «Справочная информация online»")
        if not on_log:
            print(f"  Поле не найдено — откройте справочную и введите КН: {kn}")
        return "", []

    def _run_search_after_fill() -> tuple[str, list[str]]:
        _check_abort(should_cancel, should_skip)
        if has_captcha(page):
            if wait_captcha and wait_for_captcha(
                    page, on_log=on_log, should_cancel=should_cancel,
                    should_skip=should_skip):
                pass
            else:
                log("  на странице капча — решите её в браузере")
                return "", []
        if not click_search(page, should_cancel, should_skip):
            log("  кнопка «Найти» / «Сформировать запрос» недоступна")
            return "", []
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        _interruptible_wait(page, 1200, should_cancel, should_skip)
        if has_rate_limit(page):
            log(f"  превышен лимит обращений — пауза {int(RATE_LIMIT_WAIT_SEC)} с")
            _interruptible_wait(
                page, int(RATE_LIMIT_WAIT_SEC * 1000), should_cancel, should_skip)
            if has_rate_limit(page):
                log("  лимит не снялся — нажмите «Пропустить объект» или подождите")
                return "", []
        if not wait_for_results(page, kn, should_cancel=should_cancel,
                                should_skip=should_skip):
            text = _page_text(page)
            if re.search(r"найдено\s+результатов:\s*0", text, re.I):
                log("  объект не найден (0 результатов) — проверьте вид объекта и КН")
            elif has_rate_limit(page):
                log("  превышен лимит обращений Росreestr")
            else:
                log("  результаты поиска не появились")
            return "", []
        if not open_object_card_from_results(
                page, kn, on_log=log, should_cancel=should_cancel,
                should_skip=should_skip):
            log("  не удалось открыть справку по КН в таблице результатов")
            return "", []
        _interruptible_wait(page, 1500, should_cancel, should_skip)
        form = read_form_with_wait(page, 35000, should_cancel, should_skip)
        encs = extract_encumbrances_from_text(_page_text(page))
        if form:
            return form, encs
        return read_page_data(page)

    if auto and is_lk_personal_form(page):
        for ot in object_type_candidates(object_type, lk=True):
            _check_abort(should_cancel, should_skip)
            if not wait_for_search_form(page, timeout_ms=8000,
                                        should_cancel=should_cancel,
                                        should_skip=should_skip):
                log("  форма поиска не загрузилась")
                continue
            if not _fill_lk_fields(page, kn, ot, on_log=log):
                continue
            form, encs = _run_search_after_fill()
            if form:
                return form, encs
            log(f"  по виду «{ot}» справка не найдена — пробую другой")
            try:
                goto_rosreestr_page(page, timeout=45000)
                _interruptible_wait(page, 1000, should_cancel, should_skip)
            except CollectionAborted:
                raise
            except Exception:
                pass
        log("  не удалось получить справку ни по одному виду объекта")
        return "", []

    if not try_autofill(page, kn, object_type=object_type, on_log=log):
        log("  не удалось заполнить форму поиска (вид объекта + КН)")
        return "", []

    if has_captcha(page):
        if wait_captcha and wait_for_captcha(
                page, on_log=on_log, should_cancel=should_cancel,
                should_skip=should_skip):
            pass
        else:
            log("  на странице капча — решите её в браузере")
            return "", []

    if auto:
        form, encs = _run_search_after_fill()
        if not form and not encs:
            log("  справка по объекту не появилась (таймаут или капча)")
            return "", []
        if form:
            return form, encs
        log("  справка открылась, но форма не распознана в тексте страницы")
        return "", encs

    if has_captcha(page):
        print_captcha_hint(page)
        print("  Появилась капча — решите её и запустите поиск в браузере.")
    else:
        print("  Авто-чтение не дало результата — откройте карточку объекта вручную.")
    input("  Когда «Форма собственности» видна на странице — нажмите Enter… ")
    return read_page_data(page)


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
    p.add_argument("--auto", action="store_true",
                   help="Авто-режим: сам запускает поиск и читает форму "
                        "(ручное вмешательство только при капче).")
    p.add_argument("--no-login-wait", action="store_true",
                   help="Не ждать ручного входа в начале (если уже вошли ранее).")
    p.add_argument("--profile", default=None,
                   help="Папка профиля браузера (сессия входа). По умолчанию общий "
                        "профиль рядом со скриптом — вход сохраняется между запусками "
                        "и для разных папок результатов.")
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

    # Стабильный общий профиль: вход сохраняется между запусками и не зависит от
    # папки результатов — поэтому логиниться нужно один раз, а не каждый прогон.
    if args.profile:
        profile_dir = Path(args.profile)
    else:
        profile_dir = BASE_DIR / ".pw_profile_rosreestr"
    profile_dir.mkdir(parents=True, exist_ok=True)
    print(f"Профиль браузера (сессия входа): {profile_dir}")

    with sync_playwright() as pw:
        ctx, browser_label = launch_rosreestr_context(pw, profile_dir)
        print(f"Браузер: {browser_label}")
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        # Шаг входа: открываем сайт и ждём, пока человек авторизуется (один раз).
        if not args.no_login_wait:
            try:
                goto_rosreestr_page(page, timeout=60000)
            except Exception as e:
                print(f"Не удалось открыть {ROSREESTR_ONLINE_URL}: {e}")
                print("Проверьте VPN/сеть (нужен российский IP).")
                print("Если белый экран — закройте программу и удалите папку профиля "
                      f"браузера: {profile_dir}")
            print("=" * 60)
            print("ВХОД: войдите через Госуслуги как кадастровый инженер.")
            print("Дождитесь поля «Адрес или кадастровый номер» на странице справочной.")
            print("После входа проверьте — появляется ли капча при поиске.")
            input("Когда форма поиска видна — нажмите Enter для старта… ")
            page = pick_rosreestr_page(ctx) or page
            if not ensure_search_page(page, timeout_ms=60000):
                print("Форма поиска не найдена. Откройте справочную online в браузере.")
                ctx.close()
                return

        saved = 0
        for n, kn in enumerate(todo, 1):
            print("\n" + "=" * 60)
            print(f"[{n}/{len(todo)}] КН: {kn}")
            page = pick_rosreestr_page(ctx) or page

            form, encs = auto_fetch_one(page, kn, auto=args.auto)

            if form:
                print(f"  Форма собственности: «{form}»")
                if encs:
                    print(f"  Обременения: {'; '.join(encs)}")
                if not args.auto:
                    ok = input("  Сохранить? [Enter=да / n=ввести вручную] ").strip().lower()
                    if ok == "n":
                        form = ""
            if not form:
                form = input("  Введите форму собственности вручную "
                             "(Enter — пропустить): ").strip()
            if not encs and form and not args.auto:
                manual_enc = input("  Обременения (через ';', Enter — нет): ").strip()
                if manual_enc:
                    encs = [e.strip() for e in manual_enc.split(";") if e.strip()]

            if not form:
                print("  Пропущено (значение не сохранено).")
                continue

            save_ownership_cache(rr_cache, kn, form, "rosreestr-manual",
                                 encumbrances=encs)
            append_to_forms_csv(forms_path, kn, form)
            saved += 1
            enc_note = f", обременений: {len(encs)}" if encs else ""
            print(f"  Сохранено: {kn} → «{form}»{enc_note}")

        print("\n" + "=" * 60)
        print(f"Готово. Сохранено форм: {saved} из {len(todo)}.")
        print("Теперь примените форму к отчёту и PDF (без платных запросов):")
        print(f"  В GUI: вкладка «Обслуживание» → «Пересобрать из кэша» для {args.output}")
        print(f"  Или CLI: python rebuild_from_cache.py -o {args.output} "
              f"--ownership-forms {forms_path}")
        ctx.close()


if __name__ == "__main__":
    main()
