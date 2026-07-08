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


_SHORT_OWNERSHIP = {
    "частная": "Частная",
    "муниципальная": "Муниципальная",
    "государственная": "Государственная",
    "государственная федеральная": "Государственная федеральная",
}


def _normalize_short_ownership(val: str) -> str:
    """«Частная» в модалке ЛК → нормализованное значение."""
    s = (val or "").strip()
    if not s:
        return ""
    low = s.lower()
    if low in _SHORT_OWNERSHIP:
        return _SHORT_OWNERSHIP[low]
    if low.endswith("собственность") or "собственность" in low:
        return s
    return ""


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
    # Пропустить блок «Действия» в начале модалки.
    low = text.lower()
    for cut in ("общая информация", "вид объекта недвижимости", "форма собственности"):
        p = low.find(cut)
        if p > 0 and "действ" in low[:p]:
            text = text[p:]
            break
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
            short = _normalize_short_ownership(nxt)
            if short:
                return short
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


_PREV_NUM_LABELS = (
    ("кадастровый номер", "Кадастровый номер"),
    ("инвентарный номер", "Инвентарный номер"),
    ("условный номер", "Условный номер"),
    ("иной номер", "Иной номер"),
)

_EMPTY_FIELD_VALUES = frozenset({"данные отсутствуют", "—", "-", ""})


def _format_prev_line(name: str, value: str) -> str:
    """«Кадастровый номер» + значение → одна строка для PDF/кэша."""
    name = re.sub(r"\s+", " ", (name or "").strip())
    value = re.sub(r"\s+", " ", (value or "").strip())
    if not value or value.lower() in _EMPTY_FIELD_VALUES:
        return ""
    low = name.lower()
    for key, title in _PREV_NUM_LABELS:
        if low == key or low.startswith(key):
            return f"{title} {value}"
    return f"{name} {value}" if name else value


_ACTIONS_POISON = (
    "прошу предоставить",
    "вид документа",
    "дата выдачи",
    "орган, выдавший",
    "прикрепить файл",
    "прикрепить подпись",
    "вид выписки",
)

_OBJECT_CARD_SECTIONS = (
    "Общая информация",
    "Характеристики объекта",
    "Сведения о кадастровой",
    "Ранее присвоенные номера",
    "Сведения о правах",
)


def _is_actions_poison(text: str) -> bool:
    low = (text or "").lower()
    return any(p in low for p in _ACTIONS_POISON)


def _read_modal_subinfo(page) -> list[dict]:
    """
    Пары метка/значение только из карточки объекта (.build-card-wrapper),
    без блока «Действия» (data-id=actions).
    """
    modal = _object_modal_locator(page)
    if modal is None:
        return []
    try:
        rows = modal.evaluate(
            """(root) => {
                const out = [];
                root.querySelectorAll('.build-card-wrapper .build-card-wrapper__info')
                    .forEach(sec => {
                        const section = (sec.querySelector('h3')?.innerText || '').trim();
                        sec.querySelectorAll('li.build-card-wrapper__info__ul__subinfo')
                            .forEach(li => {
                                const nameEl = li.querySelector('[class*="subinfo__name"]');
                                const valEl = li.querySelector('[class*="__line"]');
                                const name = (nameEl?.innerText || '').trim();
                                const value = (valEl?.innerText || '').trim();
                                if (name && value) out.push({section, name, value});
                            });
                    });
                return out;
            }"""
        )
        return rows if isinstance(rows, list) else []
    except Exception:
        return []


def _object_card_text(page) -> str:
    """Текст только карточки объекта (без формы «Действия»)."""
    modal = _object_modal_locator(page)
    if modal is None:
        return ""
    parts: list[str] = []
    try:
        cards = modal.locator(".build-card-wrapper")
        for i in range(min(cards.count(), 4)):
            card = cards.nth(i)
            if card.is_visible(timeout=300):
                t = card.inner_text(timeout=2000)
                if t and not _is_actions_poison(t):
                    parts.append(t)
    except Exception:
        pass
    return "\n".join(parts)


def validate_object_card_data(form: str, rows: list[dict],
                              kn: str = "") -> tuple[bool, str]:
    """Проверка: прочитана карточка объекта, а не форма «Действия»."""
    if form and _is_actions_poison(form):
        return False, "распознан текст блока «Действия», а не карточки объекта"

    if not rows:
        return False, "карточка объекта не найдена — прокрутите модалку до «Общая информация»"

    names = [(r.get("name") or "").lower() for r in rows]
    if not any("кадастровый номер" in n for n in names):
        return False, "в карточке нет кадастрового номера"

    if kn:
        kn_norm = normalize_kn(kn).replace(":", "")
        kn_vals = [
            (r.get("value") or "").replace(":", "")
            for r in rows if "кадастровый номер" in (r.get("name") or "").lower()
        ]
        if kn_vals and not any(kn_norm in v or v in kn_norm for v in kn_vals if v):
            return False, f"КН в карточке не совпадает с ожидаемым {normalize_kn(kn)}"

    has_form_field = any("форма собственности" in n for n in names)
    if has_form_field and not form:
        return False, "поле «Форма собственности» видно, но значение не прочитано"

    if kn and not (form or "").strip():
        return False, "форма собственности не прочитана"

    return True, ""


def extract_encumbrances_from_modal(page) -> list[str]:
    """Обременения из секции «Сведения о правах…» в карточке объекта."""
    modal = _object_modal_locator(page)
    if modal is None:
        return []
    try:
        chunk = modal.evaluate(
            """(root) => {
                for (const sec of root.querySelectorAll(
                        '.build-card-wrapper .build-card-wrapper__info')) {
                    const h3 = (sec.querySelector('h3')?.innerText || '').toLowerCase();
                    if (!h3.includes('правах') && !h3.includes('ограничен')) continue;
                    return sec.innerText || '';
                }
                return '';
            }"""
        )
        if chunk:
            return extract_encumbrances_from_text(str(chunk))
    except Exception:
        pass
    return extract_encumbrances_from_text(_object_card_text(page))


def _prev_lines_from_subinfo_rows(rows: list[dict]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for row in rows:
        section = (row.get("section") or "").lower()
        if "ранее присвоенные номера" not in section:
            continue
        line = _format_prev_line(row.get("name") or "", row.get("value") or "")
        if line and line not in seen:
            seen.add(line)
            out.append(line)
    return out


def extract_previous_numbers_from_modal(page) -> list[str]:
    """Блок «Ранее присвоенные номера» из DOM модалки."""
    return _prev_lines_from_subinfo_rows(_read_modal_subinfo(page))


def extract_form_from_modal(page) -> str:
    """«Форма собственности» из DOM модалки."""
    for row in _read_modal_subinfo(page):
        name = (row.get("name") or "").lower()
        if "форма собственности" not in name:
            continue
        val = (row.get("value") or "").strip()
        if not val or val.lower() in _EMPTY_FIELD_VALUES:
            continue
        short = _normalize_short_ownership(val)
        return short or val
    return ""


def extract_previous_numbers_from_text(text: str) -> list[str]:
    """Строки блока «Ранее присвоенные номера» с lk.rosreestr.ru."""
    if not text:
        return []
    # Отсечь блок «Действия» в начале модалки.
    low = text.lower()
    for cut in ("общая информация", "вид объекта недвижимости", "характеристики объекта"):
        p = low.find(cut)
        if p > 0 and low[:p].find("действ") >= 0:
            text = text[p:]
            break
    m = re.search(r"Ранее\s+присвоенные\s+номера", text, re.I)
    if not m:
        return []
    chunk = text[m.end():]
    stop = re.search(
        r"Характеристики объекта|Сведения о кадастровой|Форма собственности|"
        r"Сведения о правах|Кадастровая стоимость|Статус объекта",
        chunk, re.I)
    if stop:
        chunk = chunk[:stop.start()]

    out: list[str] = []
    seen: set[str] = set()
    lines = [ln.strip() for ln in chunk.splitlines() if ln.strip()]
    i = 0
    while i < len(lines):
        low = lines[i].lower()
        matched = None
        title = ""
        for key, label in _PREV_NUM_LABELS:
            if low == key or low.startswith(key + ":"):
                matched = key
                title = label
                val = lines[i].split(":", 1)[-1].strip() if ":" in lines[i] else ""
                if not val or val.lower() == key:
                    i += 1
                    if i >= len(lines):
                        break
                    val = lines[i].strip()
                i += 1
                if val and val.lower() not in ("данные отсутствуют", "—", "-"):
                    line = f"{title} {val}"
                    if line not in seen:
                        seen.add(line)
                        out.append(line)
                break
        if matched is None:
            # Значение на той же строке после метки (табличная вёрстка модалки).
            for key, label in _PREV_NUM_LABELS:
                m = re.match(rf"^{re.escape(key)}\s*[:\-]?\s*(.+)$", low, re.I)
                if m:
                    val = m.group(1).strip()
                    if val and val.lower() not in ("данные отсутствуют", "—", "-"):
                        line = f"{label} {lines[i].split(':', 1)[-1].strip()}"
                        if ":" not in lines[i]:
                            line = f"{label} {val}"
                        if line not in seen:
                            seen.add(line)
                            out.append(line)
                    i += 1
                    matched = key
                    break
        if matched is None:
            i += 1
    return out


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
    "Здание": ("Здание", "Здания"),
    "Помещение": ("Помещение", "Помещения"),
    "Земельный участок": ("Земельный участок", "Земельные участки"),
}

# Точные подписи пунктов в выпадающем меню ЛK (см. lk.rosreestr.ru).
LK_MENU_OPTION_TEXT: dict[str, str] = {
    "Здание": "Здание",
    "Помещение": "Помещение",
    "Земельный участок": "Земельный участок",
}

# Пауза между объектами (Росreestr режет по лимиту обращений).
PAUSE_BETWEEN_OBJECTS_SEC = 8.0
RATE_LIMIT_WAIT_SEC = 5
RESULTS_WAIT_MS = 45000
RESULTS_POLL_MS = 400
MODAL_WAIT_MS = 20000
LK_FIND_MAX_ATTEMPTS = 3
LK_LINK_CLICK_INTERVAL_MS = 3000
LK_LINK_CLICK_MAX = 15
LK_SEARCH_WAIT_AFTER_CLICK_MS = 7000
LK_SEARCH_TO_LINK_PAUSE_MS = 1500
MODAL_READY_STABLE_MS = 1000
MODAL_READY_MIN_ROWS = 3
SEARCH_MIN_WAIT_BEFORE_OUTCOME_MS = 800

# Ручной режим: запись событий браузера (см. browser_recorder.py)
_ACTIVE_RECORDER = None
_ACTIVE_RECORDER_LOG = None


def _poll_assist_events(kn: str, on_log=None) -> None:
    rec = _ACTIVE_RECORDER
    if rec is None:
        return
    log_fn = on_log or _ACTIVE_RECORDER_LOG
    for ev in rec.poll(kn=kn):
        if not log_fn:
            continue
        kind = ev.get("kind", "")
        el_id = ev.get("id", "")
        cls = ev.get("cls", "")
        if kind == "click":
            if el_id == "realestateobjects-search":
                log_fn("  [клик] НАЙТИ")
            elif ev.get("tag") == "A":
                log_fn(f"  [клик] ссылка: {(ev.get('text') or '')[:40]}")
            elif "row-" in cls or "cadNumber" in cls:
                log_fn(f"  [клик] строка результата: {(ev.get('text') or '')[:40]}")
            elif "loading-container" in cls:
                continue
        elif kind == "navigate":
            log_fn(f"  [страница] {ev.get('url', '')[:80]}")


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
    """Форма ЛК после входа (2+ dropdown, поле «Вид объекта», без капчи)."""
    try:
        wrap = _realestate_wrapper(page)
        if wrap is not None:
            n_ctrl = wrap.locator(".rros-ui-lib-dropdown__control").count()
            has_vid = wrap.get_by_text("Вид объекта", exact=True).count() > 0
            if n_ctrl >= 2 and has_vid and not has_captcha(page):
                return True
        text = _page_text(page)
        if has_captcha(page):
            return False
        if re.search(r"вид\s+объекта", text, re.I) and re.search(
                r"найти|адрес\s+или\s+кадастров", text, re.I):
            return True
        if page.get_by_text("Вид объекта", exact=False).count() > 0:
            return page.get_by_role(
                "button", name=re.compile(r"^\s*найти\s*$", re.I)).count() > 0
    except Exception:
        pass
    return False


def is_guest_reference_form(page) -> bool:
    """Гостевая форма (капча, нет «Вид объекта»)."""
    try:
        wrap = _realestate_wrapper(page)
        if wrap is not None:
            n_ctrl = wrap.locator(".rros-ui-lib-dropdown__control").count()
            has_vid = wrap.get_by_text("Вид объекта", exact=True).count() > 0
            if n_ctrl < 2 or not has_vid:
                return True
    except Exception:
        pass
    return has_captcha(page) and not is_lk_personal_form(page)


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
        const wrap = document.querySelector('.realestateobjects-wrapper');
        if (wrap && wrap.querySelector('#query')) return true;
        const t = document.body && document.body.innerText || '';
        return (wrap !== null || t.includes('Вид объекта')) &&
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


def lk_auto_type_candidates(hint: str = "") -> list[str]:
    """Авто-режим ЛK: сначала Помещение, при 0 результатов — Здание."""
    _ = hint  # подсказка API не меняет порядок (операторский сценарий)
    return ["Помещение", "Здание"]


def _label_patterns(label: str) -> tuple[str, ...]:
    """Варианты подписи для react-select (Здание / Здания и т.д.)."""
    patterns = [label]
    if label.endswith("ие"):
        patterns.extend((label + "я", label.rstrip("ие") + "ия"))
    elif label.endswith("ок"):
        patterns.append(label + "и")
    elif label.endswith("ие"):
        patterns.append(label[:-1] + "ия")
    # короткий префикс для фильтрации (Здан, Помещ)
    if len(label) >= 4:
        patterns.append(label[:4])
    seen: list[str] = []
    for p in patterns:
        if p and p not in seen:
            seen.append(p)
    return tuple(seen)


def _objtypes_hidden_value(page) -> str:
    try:
        row = _object_type_row(page)
        el = row.locator('input[name="objTypes"]').first
        if el.count():
            return (el.get_attribute("value") or "").strip()
    except Exception:
        pass
    try:
        root = _search_form_root(page)
        el = root.locator('input[name="objTypes"]').first
        if el.count():
            return (el.get_attribute("value") or "").strip()
    except Exception:
        pass
    return ""


def _clear_lk_object_type(page) -> None:
    """Сброс «Вид объекта» перед новым выбором."""
    row = _object_type_row(page)
    try:
        clear_btn = row.locator(
            "[class*='clear-indicator'], [class*='ClearIndicator']").first
        if clear_btn.count() and clear_btn.is_visible(timeout=600):
            clear_btn.click(force=True, timeout=2000)
            page.wait_for_timeout(350)
            return
    except Exception:
        pass
    try:
        inp = _object_type_input(page)
        if inp.count() and inp.is_visible(timeout=600):
            inp.click(force=True, timeout=2000)
            inp.press("Control+a")
            inp.press("Backspace")
            page.wait_for_timeout(250)
    except Exception:
        pass
    _dismiss_dropdowns(page)


def _realestate_wrapper(page):
    """Блок формы «Справочная информация online» в ЛК."""
    try:
        wrap = page.locator(".realestateobjects-wrapper").first
        if wrap.count() and wrap.is_visible(timeout=1200):
            return wrap
    except Exception:
        pass
    return None


def _search_form_root(page):
    """Центральная форма поиска (не боковое меню и не чужие react-select)."""
    wrap = _realestate_wrapper(page)
    if wrap is not None:
        return wrap
    try:
        kn_ph = page.get_by_placeholder(re.compile(
            r"Введите адрес или кадастров", re.I)).first
        if kn_ph.is_visible(timeout=2000):
            for xpath in (
                "xpath=ancestor::form[1]",
                "xpath=ancestor::div[.//*[contains(text(),'Вид объекта')]][1]",
                "xpath=ancestor::div[contains(@class,'content') or "
                "contains(@class,'Content')][1]",
            ):
                form = kn_ph.locator(xpath)
                if form.count() and form.first.is_visible(timeout=800):
                    txt = form.first.inner_text(timeout=1500)
                    if "вид объекта" in txt.lower() and "найти" in txt.lower():
                        return form.first
    except Exception:
        pass
    try:
        for getter in (
            lambda: page.locator("main, [role='main']").filter(
                has=page.get_by_placeholder(re.compile(
                    r"Введите адрес или кадастров", re.I))
            ).filter(
                has=page.get_by_text("Вид объекта", exact=False)
            ).first,
            lambda: page.locator("div").filter(
                has=page.get_by_placeholder(re.compile(
                    r"Введите адрес или кадастров", re.I))
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
    return page.locator("body")


def _dismiss_dropdowns(page) -> None:
    if _object_modal_locator(page) is not None:
        return
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(250)
    except Exception:
        pass


def has_rate_limit(page) -> bool:
    """Росreestr: «Превышен лимит обращений, попробуйте позже»."""
    low = _page_text(page).lower()
    return "превышен лимит" in low or "лимит обращений" in low


def _object_type_row(page):
    """Строка формы «Вид объекта» — 2-й блок .realestateobjects-wrapper__option."""
    root = _search_form_root(page)
    try:
        row = root.locator(".realestateobjects-wrapper__option").nth(1)
        if row.count() and row.is_visible(timeout=800):
            return row
    except Exception:
        pass
    return root.locator(".realestateobjects-wrapper__option").filter(
        has=root.get_by_text("Вид объекта", exact=True)).first


def _object_type_control(page):
    """Control react-select «Вид объекта» (2-й dropdown в форме)."""
    root = _search_form_root(page)
    try:
        ctrl = root.locator(".rros-ui-lib-dropdown__control").nth(1)
        if ctrl.count() and ctrl.is_visible(timeout=800):
            return ctrl
    except Exception:
        pass
    return _object_type_row(page).locator(".rros-ui-lib-dropdown__control").first


def _object_type_input(page):
    """Input react-select «Вид объекта»."""
    row = _object_type_row(page)
    try:
        inp = row.locator("input[id*='react-select']").first
        if inp.count():
            return inp
    except Exception:
        pass
    root = _search_form_root(page)
    return root.locator(".realestateobjects-wrapper__option").nth(1).locator(
        "input[id*='react-select']").first


def _object_type_display(page) -> str:
    """Текущее значение «Вид объекта» (single-value, input или control)."""
    try:
        row = _object_type_row(page)
        sv = row.locator(".rros-ui-lib-dropdown__single-value").first
        if sv.count() and sv.is_visible(timeout=500):
            t = (sv.inner_text(timeout=800) or "").strip()
            if t and "выберите" not in t.lower():
                return t
    except Exception:
        pass
    try:
        inp = _object_type_input(page)
        if inp.count() and inp.is_visible(timeout=400):
            val = (inp.input_value() or "").strip()
            if val and "выберите" not in val.lower():
                return val
    except Exception:
        pass
    try:
        ctrl = _object_type_control(page)
        if ctrl.count() and ctrl.is_visible(timeout=400):
            raw = (ctrl.inner_text(timeout=800) or "").strip()
            for noise in ("Вид объекта", "Выберите значение из справочника"):
                raw = raw.replace(noise, "")
            raw = re.sub(r"\s+", " ", raw).strip()
            if raw and "выберите" not in raw.lower():
                return raw
    except Exception:
        pass
    return ""


def _object_type_committed(page) -> bool:
    """True, если react-select зафиксировал значение (hidden objTypes или single-value)."""
    hidden = _objtypes_hidden_value(page)
    if hidden:
        return True
    try:
        row = _object_type_row(page)
        sv = row.locator(".rros-ui-lib-dropdown__single-value").first
        if sv.count() and sv.is_visible(timeout=500):
            t = (sv.inner_text(timeout=600) or "").strip()
            return bool(t) and "выберите" not in t.lower()
        ph = row.locator(".rros-ui-lib-dropdown__placeholder").first
        if ph.count() and ph.is_visible(timeout=300):
            return False
    except Exception:
        pass
    return False


def _object_type_is_set(page) -> bool:
    """True, если в «Вид объекта» выбрано и зафиксировано значение."""
    return _object_type_committed(page)


def _object_type_current(page) -> str:
    """Текущее значение «Вид объекта» в форме ЛК."""
    return _object_type_display(page)


def _object_type_matches(page, canonical: str) -> bool:
    """True, если в форме уже выбран нужный вид объекта."""
    if not canonical:
        return False
    cur = _object_type_display(page).lower()
    if not cur:
        hidden = _objtypes_hidden_value(page)
        if not hidden:
            return False
        cur = hidden.lower()
    canon = canonical.lower()
    if canon in cur or cur in canon:
        return True
    if cur.startswith(canon[:4]) or canon.startswith(cur[:4]):
        return True
    if "помещ" in canon and "помещ" in cur:
        return True
    if "здан" in canon and "здан" in cur:
        return True
    if "земел" in canon and ("земел" in cur or "участ" in cur):
        return True
    for label in LK_TYPE_UI_LABELS.get(canonical, (canonical,)):
        low = label.lower()
        if low in cur or cur in low or cur.startswith(low[:4]):
            return True
    return False


def _open_object_type_dropdown(page) -> bool:
    """Открывает выпадающий список «Вид объекта» в центральной форме."""
    return _open_lk_object_type_menu(page)


def _wait_lk_option_menu(page, timeout_ms: int = 8000) -> bool:
    """Ждёт появления пунктов react-select (после открытия «Вид объекта»)."""
    step = 250
    waited = 0
    while waited < timeout_ms:
        try:
            opts = page.locator('[id*="react-select"][id*="-option-"]')
            for i in range(min(opts.count(), 40)):
                opt = opts.nth(i)
                if not opt.is_visible(timeout=150):
                    continue
                txt = (opt.inner_text(timeout=300) or "").strip()
                if txt and len(txt) < 120 and "тип поиска" not in txt.lower():
                    return True
        except Exception:
            pass
        try:
            opts = page.locator(
                ".rros-ui-lib-dropdown__menu .rros-ui-lib-dropdown__option")
            for i in range(min(opts.count(), 40)):
                opt = opts.nth(i)
                if not opt.is_visible(timeout=150):
                    continue
                txt = (opt.inner_text(timeout=300) or "").strip()
                if txt and len(txt) < 120 and "тип поиска" not in txt.lower():
                    return True
        except Exception:
            pass
        try:
            opt = page.get_by_role("option").first
            if opt.is_visible(timeout=150):
                return True
        except Exception:
            pass
        page.wait_for_timeout(step)
        waited += step
    return False


def _visible_lk_menus(page):
    """Видимые выпадающие меню rros-ui-lib (портал в body)."""
    out = []
    try:
        for menu in page.locator(".rros-ui-lib-dropdown__menu").all():
            if menu.is_visible():
                out.append(menu)
    except Exception:
        pass
    return out


def _open_lk_object_type_menu(page) -> bool:
    """Открывает меню «Вид объекта» (2-й dropdown)."""
    control = _object_type_control(page)
    inp = _object_type_input(page)
    indicator = _object_type_row(page).locator(
        ".rros-ui-lib-dropdown__dropdown-indicator").first

    for target in (control, inp, indicator):
        try:
            if not target.count() or not target.is_visible(timeout=1500):
                continue
            target.scroll_into_view_if_needed()
            target.click(force=True, timeout=5000)
            page.wait_for_timeout(500)
            if _wait_lk_option_menu(page, timeout_ms=4000):
                return True
            try:
                inp.press("ArrowDown", timeout=1500)
            except Exception:
                page.keyboard.press("ArrowDown")
            page.wait_for_timeout(450)
            if _wait_lk_option_menu(page, timeout_ms=3000):
                return True
        except Exception:
            continue
    return False


def _pick_lk_object_type_option(page, label: str) -> bool:
    """Клик по пункту открытого меню «Вид объекта» (Здание / Помещение / …)."""
    option_text = LK_MENU_OPTION_TEXT.get(label, label)
    patterns = (option_text,) + _label_patterns(label)

    for menu in reversed(_visible_lk_menus(page)):
        for pat in patterns:
            try:
                opt = menu.locator(".rros-ui-lib-dropdown__option").filter(
                    has_text=re.compile(rf"^\s*{re.escape(pat)}\s*$", re.I)).first
                if opt.count() and opt.is_visible(timeout=1200):
                    opt.scroll_into_view_if_needed()
                    opt.click(force=True, timeout=5000)
                    page.wait_for_timeout(500)
                    return True
            except Exception:
                continue

    for pat in patterns:
        try:
            opt = page.get_by_role(
                "option", name=re.compile(rf"^\s*{re.escape(pat)}\s*$", re.I)).first
            if opt.is_visible(timeout=1500):
                opt.click(force=True, timeout=5000)
                page.wait_for_timeout(500)
                return True
        except Exception:
            pass

    for pat in patterns:
        try:
            opt = page.locator(".rros-ui-lib-dropdown__option").filter(
                has_text=re.compile(rf"^\s*{re.escape(pat)}\s*$", re.I)).last
            if opt.count() and opt.is_visible(timeout=1500):
                opt.click(force=True, timeout=5000)
                page.wait_for_timeout(500)
                return True
        except Exception:
            pass
    return False


def _commit_lk_object_type(page, canonical: str, *, on_log=None) -> bool:
    """
    Выбор «Вид объекта»: клик в поле → ввод текста → Enter (зафиксировать).
    Как вручную: Пomещение → Enter, затем КН.
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if not canonical:
        return False
    if _object_type_matches(page, canonical) and _object_type_committed(page):
        return True

    option_text = LK_MENU_OPTION_TEXT.get(canonical, canonical)
    inp = _object_type_input(page)

    for attempt in range(1, 4):
        _clear_lk_object_type(page)
        log(f"  вид объекта «{option_text}» — попытка {attempt}…")
        try:
            if not inp.count() or not inp.is_visible(timeout=2000):
                log("  поле «Вид объекта» не найдено")
                continue
            inp.scroll_into_view_if_needed()
            inp.click(force=True, timeout=5000)
            page.wait_for_timeout(350)
            inp.press("Control+a")
            inp.press("Backspace")
            page.wait_for_timeout(200)
            inp.type(option_text, delay=85)
            page.wait_for_timeout(750)
            _wait_lk_option_menu(page, timeout_ms=2500)

            inp.press("Enter")
            page.wait_for_timeout(1000)
            if _object_type_committed(page) and _object_type_matches(page, canonical):
                log(f"  зафиксировано: «{option_text}»")
                return True

            inp.click(force=True, timeout=3000)
            page.wait_for_timeout(200)
            inp.press("ArrowDown")
            page.wait_for_timeout(250)
            inp.press("Enter")
            page.wait_for_timeout(1000)
            if _object_type_committed(page) and _object_type_matches(page, canonical):
                log(f"  зафиксировано (↓+Enter): «{option_text}»")
                return True

            if _open_lk_object_type_menu(page) and _pick_lk_object_type_option(page, canonical):
                page.wait_for_timeout(400)
                page.keyboard.press("Enter")
                page.wait_for_timeout(900)
                if _object_type_committed(page) and _object_type_matches(page, canonical):
                    log(f"  зафиксировано (клик+Enter): «{option_text}»")
                    return True
        except Exception as exc:
            log(f"  ошибка выбора вида: {exc}")

    if _object_type_committed(page):
        log(f"  в поле «{_object_type_display(page)}», нужен «{canonical}»")
    return False


def _select_lk_object_type(page, canonical: str, *, on_log=None) -> bool:
    return _commit_lk_object_type(page, canonical, on_log=on_log)


def _kn_query_value(page) -> str:
    try:
        el = _search_form_root(page).locator("#query").first
        if el.count():
            return (el.input_value() or "").strip()
    except Exception:
        pass
    return ""


def _kn_input_matches(page, kn: str) -> bool:
    kn_norm = normalize_kn(kn)
    val = normalize_kn(_kn_query_value(page))
    if not val:
        return False
    return val == kn_norm or kn_norm.replace(":", "") in val.replace(":", "")


def _is_search_form_filled(page, kn: str) -> bool:
    """True, если вид объекта зафиксирован и КН введён (как на скриншоте)."""
    return _object_type_committed(page) and _kn_input_matches(page, kn)


def _wait_find_button_enabled(page, timeout_ms: int = 12000) -> bool:
    """Ждёт активной кнопки «НАЙТИ» после заполнения формы."""
    root = _search_form_root(page)
    btn = root.locator("#realestateobjects-search").first
    if not btn.count():
        btn = root.get_by_role("button", name=re.compile(r"^\s*найти\s*$", re.I)).first
    step = 250
    waited = 0
    while waited < timeout_ms:
        try:
            if btn.is_visible():
                if btn.get_attribute("disabled") is None:
                    return True
                if btn.is_enabled():
                    return True
        except Exception:
            pass
        page.wait_for_timeout(step)
        waited += step
    return False


def _results_block(page):
    """Блок «Найдено результатов» + таблица (DOM ЛK)."""
    return page.locator(".realestateobjects-wrapper__results").first


def _results_area(page):
    try:
        loc = _results_block(page)
        if loc.count() and loc.is_visible(timeout=400):
            return loc
    except Exception:
        pass
    return page.locator(".realestateobjects-wrapper").first


def _results_area_text(page) -> str:
    try:
        return _results_area(page).inner_text(timeout=2500)
    except Exception:
        return ""


def _parse_results_count(text: str) -> int | None:
    m = re.search(r"найдено\s+результатов:\s*(\d+)", text or "", re.I)
    return int(m.group(1)) if m else None


def _results_total_count_el(page):
    return page.locator(".realestateobjects-wrapper__results__total-count").first


def _read_results_count_dom(page) -> int | None:
    """Читает число из span.realestateobjects-wrapper__results__total-count."""
    try:
        el = _results_total_count_el(page)
        if el.count() and el.is_visible(timeout=500):
            return _parse_results_count(el.inner_text(timeout=1500))
    except Exception:
        pass
    return _parse_results_count(_results_area_text(page))


def _is_results_loading(page) -> bool:
    try:
        loc = page.locator(
            ".realestateobjects-wrapper__results .rros-ui-lib-loading-container"
        ).first
        if not loc.count() or not loc.is_visible(timeout=200):
            return False
        busy = loc.get_attribute("aria-busy")
        if busy == "true":
            return True
        spinner = loc.locator(
            "[class*='loading--active'], [class*='spinner'], .rros-ui-lib-loading"
        ).first
        return spinner.count() > 0 and spinner.is_visible(timeout=150)
    except Exception:
        return False


def _wait_results_loading(page, timeout_ms: int = 35000,
                          should_cancel=None, should_skip=None) -> None:
    """Ждёт завершения индикатора загрузки таблицы."""
    step = 350
    waited = 0
    seen_loading = False
    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip)
        if _is_results_loading(page):
            seen_loading = True
        elif seen_loading or _read_results_count_dom(page) is not None:
            return
        elif waited > 1500 and _results_table_rows(page).count() > 0:
            return
        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step


def _results_table_rows(page):
    block = _results_block(page)
    return block.locator(".rros-ui-lib-table__row[data-test-id^='row-']")


def _results_table_signature(page) -> str:
    try:
        block = _results_block(page)
        if not block.count() or not block.is_visible(timeout=200):
            return ""
        return (block.inner_text(timeout=1000) or "")[:800]
    except Exception:
        return ""


def _results_row_has_kn(page, kn: str) -> bool:
    kn_norm = normalize_kn(kn)
    needle = kn_norm.replace(":", "")
    if not needle:
        return False
    try:
        rows = _results_table_rows(page)
        for i in range(min(rows.count(), 25)):
            txt = rows.nth(i).inner_text(timeout=800)
            if _kn_in_text(txt, kn_norm):
                return True
    except Exception:
        pass
    return False


def _find_results_row(page, kn: str):
    """Строка таблицы результатов с нужным КН."""
    kn_norm = normalize_kn(kn)
    rows = _results_table_rows(page)
    try:
        n = min(rows.count(), 25)
    except Exception:
        n = 0
    for i in range(n):
        row = rows.nth(i)
        try:
            if not row.is_visible(timeout=400):
                continue
            txt = row.inner_text(timeout=1000)
            if any(_kn_in_text(txt, p) for p in (kn_norm, kn_norm.replace(":", "-"))):
                return row
        except Exception:
            continue
    return None


def search_has_zero_results(page, *, fresh: bool = False,
                            before_count: int | None = None) -> bool:
    """True, если в блоке результатов «Найдено результатов: 0»."""
    try:
        block = _results_block(page)
        if not block.count() or not block.is_visible(timeout=500):
            return False
    except Exception:
        return False
    count = _read_results_count_dom(page)
    if count is None:
        return False
    if fresh and before_count is not None and count == before_count:
        return False
    return count == 0


def _lk_form_card(page):
    return page.locator(".realestateobjects-wrapper.card").first


def _is_lk_form_loading(page) -> bool:
    try:
        loc = page.locator(
            ".realestateobjects-wrapper.card .rros-ui-lib-loading-container"
        ).first
        if not loc.count() or not loc.is_visible(timeout=150):
            return False
        if loc.get_attribute("aria-busy") == "true":
            return True
        spinner = loc.locator(
            "[class*='loading--active'], [class*='spinner']"
        ).first
        return spinner.count() > 0 and spinner.is_visible(timeout=100)
    except Exception:
        return False


def _wait_lk_form_idle(page, timeout_ms: int = 10000) -> None:
    step = 200
    waited = 0
    while waited < timeout_ms and _is_lk_form_loading(page):
        page.wait_for_timeout(step)
        waited += step
    page.wait_for_timeout(200)


def _js_click_find_button(page) -> bool:
    return bool(page.evaluate("""() => {
        const b = document.querySelector(
            '.realestateobjects-wrapper.card #realestateobjects-search')
            || document.getElementById('realestateobjects-search');
        if (!b) return false;
        b.disabled = false;
        b.removeAttribute('disabled');
        b.scrollIntoView({block: 'center', inline: 'nearest'});
        b.focus({preventScroll: true});
        for (const type of ['pointerdown', 'mousedown', 'mouseup', 'click']) {
            b.dispatchEvent(new MouseEvent(type, {
                bubbles: true, cancelable: true, view: window
            }));
        }
        return true;
    }"""))


def _search_request_started(page, before_count: int | None, kn: str = "") -> bool:
    """True, если после «Найти» пошёл запрос (спиннер или новый счётчик)."""
    if _is_results_loading(page) or _is_lk_form_loading(page):
        return True
    count = _read_results_count_dom(page)
    if count is not None and count != before_count:
        return True
    if kn and _find_kn_result_link(page, kn):
        return True
    return False


def _click_find_button_once(page, kn: str = "") -> bool:
    """JS/locator-клик «Найти»; True, если запрос пошёл."""
    _wait_lk_form_idle(page)
    _dismiss_dropdowns(page)

    wrap = _lk_form_card(page)
    btn = wrap.locator("#realestateobjects-search").first
    if not btn.count():
        btn = page.locator("#realestateobjects-search").first

    btn.scroll_into_view_if_needed()
    page.wait_for_timeout(250)
    before_count = _read_results_count_dom(page)

    if _js_click_find_button(page):
        page.wait_for_timeout(500)
        if _search_request_started(page, before_count, kn):
            return True

    try:
        page.get_by_role(
            "button", name=re.compile(r"^\s*найти\s*$", re.I)
        ).first.click(force=True, timeout=5000)
        page.wait_for_timeout(500)
        if _search_request_started(page, before_count, kn):
            return True
    except Exception:
        pass

    try:
        btn.click(force=True, timeout=5000)
        page.wait_for_timeout(500)
        if _search_request_started(page, before_count, kn):
            return True
    except Exception:
        pass

    try:
        _submit_via_kn_enter(page)
        page.wait_for_timeout(500)
        return _search_request_started(page, before_count, kn)
    except Exception:
        return False


def _prepare_lk_search(page, kn: str = "") -> None:
    """Закрывает чужую карточку перед новым поиском."""
    modal = _object_modal_locator(page)
    if modal is None:
        return
    if kn and _modal_has_kn_strict(page, kn):
        return
    _close_object_modal(page)
    page.wait_for_timeout(600)


def _modal_shell_open(page) -> bool:
    modal = _object_modal_locator(page)
    if modal is None:
        return False
    try:
        card = modal.locator(".build-card-wrapper").first
        return card.count() > 0 and card.is_visible(timeout=300)
    except Exception:
        return False


def _is_modal_content_loading(page) -> bool:
    modal = _object_modal_locator(page)
    if modal is None:
        return True
    try:
        spin = modal.locator(".rros-ui-lib-loading-container").first
        if spin.count() and spin.is_visible(timeout=200):
            busy = spin.get_attribute("aria-busy")
            if busy == "true":
                return True
            inner = spin.locator(
                "[class*='loading--active'], .rros-ui-lib-loading"
            ).first
            if inner.count() and inner.is_visible(timeout=150):
                return True
    except Exception:
        pass
    try:
        card = modal.locator(".build-card-wrapper").first
        if not card.count() or not card.is_visible(timeout=250):
            return True
    except Exception:
        return True
    return False


def _modal_has_kn_strict(page, kn: str) -> bool:
    kn_norm = normalize_kn(kn)
    for row in _read_modal_subinfo(page):
        name = (row.get("name") or "").lower()
        if "кадастровый номер" not in name:
            continue
        if _kn_in_text(row.get("value") or "", kn_norm):
            return True
    return False


def _modal_data_ready(page, kn: str) -> tuple[bool, str]:
    kn_norm = normalize_kn(kn)
    if _object_modal_locator(page) is None:
        return False, "модальное окно не открыто"
    if _is_modal_content_loading(page):
        return False, "загрузка карточки"
    if not _modal_has_kn_strict(page, kn_norm):
        return False, f"КН {kn_norm} не найден в «Общая информация»"
    rows = _read_modal_subinfo(page)
    if len(rows) < MODAL_READY_MIN_ROWS:
        return False, f"мало полей в карточке ({len(rows)})"
    form = extract_form_from_modal(page)
    if not form:
        form = extract_form_from_text(_object_card_text(page))
    return validate_object_card_data(form, rows, kn_norm)


def _search_outcome_after_click(page, kn: str, before_count: int | None,
                                before_sig: str = "",
                                min_waited_ms: int = 0):
    """
    После нажатия «Найти»: 'link' | 'zero' | None (ещё ждём).
    Учитывает только блок .realestateobjects-wrapper__results.
    """
    kn_norm = normalize_kn(kn)
    if _results_has_kn_link(page, kn_norm):
        return "link"
    if (_object_modal_locator(page) is not None
            and min_waited_ms >= SEARCH_MIN_WAIT_BEFORE_OUTCOME_MS):
        ok, _ = _modal_data_ready(page, kn_norm)
        if ok:
            return "link"
    count = _read_results_count_dom(page)
    if count == 0:
        if _results_has_kn_link(page, kn_norm):
            return "link"
        if has_rate_limit(page):
            return None
        if min_waited_ms < 2000 and before_count == 0:
            return None
        if before_count is not None and count != before_count:
            if min_waited_ms < 1500:
                return None
            return "zero"
        if min_waited_ms >= 2500 and before_count == 0:
            return "zero"
        return None
    if count is not None and count >= 1:
        if _is_results_loading(page):
            return None
        if min_waited_ms < SEARCH_MIN_WAIT_BEFORE_OUTCOME_MS:
            return None
        sig = _results_table_signature(page)
        if before_sig and sig == before_sig and before_count == count:
            return None
        if _results_row_has_kn(page, kn_norm):
            if not _kn_in_text(sig, kn_norm):
                return None
            return "link"
        if before_count is not None and count != before_count:
            return None
    return None


def lk_search_until_outcome(page, kn: str, *, on_log=None,
                            should_cancel=None, should_skip=None) -> str:
    """
    «Найти» (до LK_FIND_MAX_ATTEMPTS) → одно ожидание ответа до 45 с.
    Повтор только если кнопка не сработала или сервер не ответил.
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    kn_norm = normalize_kn(kn)
    _prepare_lk_search(page, kn_norm)

    if _results_has_kn_link(page, kn_norm):
        cnt = _read_results_count_dom(page)
        log(f"  найдено результатов: {cnt if cnt is not None else '≥1'}")
        return "link"

    for attempt in range(1, LK_FIND_MAX_ATTEMPTS + 1):
        _check_abort(should_cancel, should_skip)
        if attempt > 1:
            _prepare_lk_search(page, kn_norm)
        if (_results_row_has_kn(page, kn_norm)
                or _results_has_kn_link(page, kn_norm)):
            cnt = _read_results_count_dom(page)
            log(f"  найдено результатов: {cnt if cnt is not None else '≥1'}")
            return "link"
        before_count = _read_results_count_dom(page)
        before_sig = _results_table_signature(page)
        log(f"  «Найти» — попытка {attempt}/{LK_FIND_MAX_ATTEMPTS}…")

        if not _click_find_button_once(page, kn_norm):
            log("  «Найти» не сработала — повтор JS-клика…")
            _js_click_find_button(page)
            page.wait_for_timeout(800)
            if not _search_request_started(page, before_count, kn_norm):
                log("  поиск не отправлен")
                continue

        log(f"  жду ответ сервера (до {LK_SEARCH_WAIT_AFTER_CLICK_MS // 1000} с)…")
        waited = 0
        stable_zero_ticks = 0
        last_progress = -10000

        while waited < LK_SEARCH_WAIT_AFTER_CLICK_MS:
            _check_abort(should_cancel, should_skip)
            _poll_assist_events(kn_norm, on_log=on_log)
            outcome = _search_outcome_after_click(
                page, kn_norm, before_count, before_sig, waited)
            if outcome == "link":
                cnt = _read_results_count_dom(page)
                log(f"  найдено результатов: {cnt if cnt is not None else '≥1'}")
                return "link"
            if outcome == "zero":
                if _results_has_kn_link(page, kn_norm):
                    cnt = _read_results_count_dom(page)
                    log(f"  найдено результатов: {cnt if cnt is not None else '≥1'}")
                    return "link"
                log("  найдено 0 результатов")
                return "zero"
            if (_is_results_loading(page) or _is_lk_form_loading(page)
                    ) and waited - last_progress >= 8000:
                log(f"  загрузка… ({waited // 1000} с)")
                last_progress = waited
            count = _read_results_count_dom(page)
            if (count == 0 and before_count == 0
                    and not _is_results_loading(page) and waited >= 2000):
                stable_zero_ticks += 1
                if stable_zero_ticks >= 4:
                    if _results_has_kn_link(page, kn_norm):
                        cnt = _read_results_count_dom(page)
                        log(f"  найдено результатов: {cnt if cnt is not None else '≥1'}")
                        return "link"
                    log("  найдено 0 результатов")
                    return "zero"
            else:
                stable_zero_ticks = 0
            _interruptible_wait(page, RESULTS_POLL_MS, should_cancel, should_skip)
            waited += RESULTS_POLL_MS

        if _results_row_has_kn(page, kn_norm) or _results_has_kn_link(page, kn_norm):
            cnt = _read_results_count_dom(page)
            log(f"  найдено результатов: {cnt if cnt is not None else '≥1'}")
            return "link"
        log("  ответ не получен — повтор «Найти»")
    return "fail"


def open_card_until_ready(page, kn: str, *, on_log=None,
                          should_cancel=None, should_skip=None) -> bool:
    """Левый клик по ссылке с КН каждые 3 с, пока не откроется модалка."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    kn_norm = normalize_kn(kn)
    if _modal_data_ready(page, kn_norm)[0]:
        log("  карточка объекта уже открыта")
        return True

    for attempt in range(1, LK_LINK_CLICK_MAX + 1):
        _check_abort(should_cancel, should_skip)
        if _modal_shell_open(page) and _modal_has_kn_strict(page, kn_norm):
            log("  карточка объекта загружена")
            return True
        target = _find_kn_result_link(page, kn_norm)
        if target is None:
            log("  ссылка с КН не видна — жду таблицу…")
            if not wait_for_results(page, kn_norm, timeout_ms=15000,
                                    should_cancel=should_cancel,
                                    should_skip=should_skip, on_log=log):
                return False
            target = _find_kn_result_link(page, kn_norm)
        if target is not None:
            log(f"  клик по ссылке ({attempt}/{LK_LINK_CLICK_MAX})…")
            try:
                _click_result_target(page, target)
            except Exception as exc:
                log(f"  клик: {exc}")
            shell_wait = 0
            while shell_wait < LK_LINK_CLICK_INTERVAL_MS:
                _check_abort(should_cancel, should_skip)
                if _modal_shell_open(page) and _modal_has_kn_strict(page, kn_norm):
                    log("  карточка объекта загружена")
                    return True
                _interruptible_wait(
                    page, 400, should_cancel, should_skip)
                shell_wait += 400
        else:
            _interruptible_wait(
                page, LK_LINK_CLICK_INTERVAL_MS, should_cancel, should_skip)
    if _modal_shell_open(page) and _modal_has_kn_strict(page, kn_norm):
        log("  карточка объекта загружена")
        return True
    log("  модальное окно карточки не открылось")
    return False


def _lk_open_and_read_card(page, kn: str, *, on_log=None,
                           should_cancel=None, should_skip=None,
                           skip_link_pause: bool = False
                           ) -> tuple[str, list[str], list[str]]:
    """Пауза → клик по ссылке → ожидание карточки → чтение данных."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    kn_norm = normalize_kn(kn)
    already_open = (_modal_shell_open(page)
                    and _modal_has_kn_strict(page, kn_norm))

    if not already_open:
        if not skip_link_pause:
            log(f"  пауза {LK_SEARCH_TO_LINK_PAUSE_MS // 1000} с перед переходом по ссылке…")
            _interruptible_wait(
                page, LK_SEARCH_TO_LINK_PAUSE_MS, should_cancel, should_skip)

        if not open_card_until_ready(
                page, kn, on_log=log, should_cancel=should_cancel,
                should_skip=should_skip):
            log("  не удалось открыть карточку по ссылке")
            return "", [], []
    else:
        log("  карточка объекта уже открыта — жду данные")

    if not wait_for_modal_card_ready(
            page, kn, timeout_ms=MODAL_WAIT_MS,
            should_cancel=should_cancel, should_skip=should_skip, on_log=log):
        ok, reason = _modal_data_ready(page, kn_norm)
        if reason:
            log(f"  {reason}")
        log("  данные карточки не загрузились — прокрутите модалку вручную")
        return "", [], []

    log("  читаю данные карточки…")
    form, encs, prev = read_page_data(page, kn)
    if not form:
        log("  повторное чтение карточки…")
        _scroll_object_modal(page)
        page.wait_for_timeout(800)
        form, encs, prev = read_page_data(page, kn)
    if form:
        log(f"  форма собственности: «{form}»")
        if encs:
            log(f"  обременений: {len(encs)}")
    else:
        rows = _read_modal_subinfo(page)
        _, reason = validate_object_card_data("", rows, kn_norm)
        if not reason:
            _, reason = _modal_data_ready(page, kn_norm)
        if reason:
            log(f"  {reason}")
    return form, encs, prev


def lk_fetch_with_type(page, kn: str, object_type: str, *, on_log=None,
                       should_cancel=None, should_skip=None) -> tuple[str, list[str], list[str]]:
    """Заполнить форму → цикл «Найти» → открыть карточку → прочитать данные."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    kn_norm = normalize_kn(kn)
    _prepare_lk_search(page, kn_norm)

    if _results_has_kn_link(page, kn_norm):
        log("  в таблице уже есть ссылка с КН — перехожу к карточке")
        return _lk_open_and_read_card(
            page, kn, on_log=log, should_cancel=should_cancel,
            should_skip=should_skip)
    if (_modal_shell_open(page) and _modal_has_kn_strict(page, kn_norm)):
        log("  карточка объекта уже открыта — читаю данные")
        return _lk_open_and_read_card(
            page, kn, on_log=log, should_cancel=should_cancel,
            should_skip=should_skip, skip_link_pause=True)

    if not _fill_lk_fields(page, kn, object_type, on_log=log):
        log(f"  не удалось зафиксировать «{object_type}» (Enter)")
        return "", [], []

    outcome = lk_search_until_outcome(
        page, kn, on_log=log, should_cancel=should_cancel, should_skip=should_skip)
    if outcome == "zero":
        if _results_has_kn_link(page, kn_norm):
            log("  ссылка с КН появилась — перехожу к карточке")
            return _lk_open_and_read_card(
                page, kn, on_log=log, should_cancel=should_cancel,
                should_skip=should_skip)
        return "", [], []
    if outcome != "link":
        if _results_has_kn_link(page, kn_norm):
            log("  ссылка с КН в таблице — перехожу к карточке")
            return _lk_open_and_read_card(
                page, kn, on_log=log, should_cancel=should_cancel,
                should_skip=should_skip)
        log("  поиск не дал ссылку на объект")
        return "", [], []

    return _lk_open_and_read_card(
        page, kn, on_log=log, should_cancel=should_cancel,
        should_skip=should_skip)


def _find_kn_result_link(page, kn: str):
    """Кликабельный элемент с КН: ссылка, ячейка или строка таблицы."""
    if not kn:
        return None
    kn_norm = normalize_kn(kn)
    parts = (kn_norm, kn_norm.replace(":", "-"))

    row = _find_results_row(page, kn_norm)
    if row is not None:
        for part in parts:
            try:
                link = row.locator("a").filter(
                    has_text=re.compile(re.escape(part))).first
                if link.count() and link.is_visible(timeout=500):
                    return link
            except Exception:
                pass
            try:
                cell = row.locator(
                    "[class*='cadNumber'], [class*='cell'], [role='cell']"
                ).filter(has_text=re.compile(re.escape(part))).first
                if cell.count() and cell.is_visible(timeout=500):
                    inner = cell.locator("a").first
                    if inner.count() and inner.is_visible(timeout=300):
                        return inner
                    return cell
            except Exception:
                pass
            try:
                txt_el = row.get_by_text(part, exact=True).first
                if txt_el.count() and txt_el.is_visible(timeout=500):
                    return txt_el
            except Exception:
                pass
        try:
            txt = row.inner_text(timeout=800)
            if any(_kn_in_text(txt, p) for p in parts):
                return row
        except Exception:
            pass
        return None

    areas = (
        _results_block(page),
        page.locator(".realestateobjects-wrapper").first,
        _search_form_root(page),
        page.locator("main").first,
    )
    for area in areas:
        try:
            if not area.count() or not area.is_visible(timeout=300):
                continue
        except Exception:
            continue
        for part in parts:
            for get_link in (
                lambda p=part, a=area: a.locator("[class*='cadNumber'] a").filter(
                    has_text=re.compile(re.escape(p))).first,
                lambda p=part, a=area: a.locator(
                    ".realestateobjects-wrapper__results a").filter(
                    has_text=re.compile(re.escape(p))).first,
                lambda p=part, a=area: a.locator(
                    "[class*='results'] a, tbody a, .rt-tbody a").filter(
                    has_text=re.compile(re.escape(p))).first,
                lambda p=part, a=area: a.get_by_role(
                    "link", name=re.compile(re.escape(p))).first,
                lambda p=part, a=area: a.locator("a").filter(
                    has_text=re.compile(rf"^\s*{re.escape(p)}\s*$")).first,
            ):
                try:
                    link = get_link()
                    if link.count() and link.is_visible(timeout=400):
                        return link
                except Exception:
                    continue
            for get_cell in (
                lambda p=part, a=area: a.get_by_text(p, exact=True).first,
                lambda p=part, a=area: a.locator("td, [class*='cell']").filter(
                    has_text=re.compile(re.escape(p))).first,
            ):
                try:
                    cell = get_cell()
                    if cell.count() and cell.is_visible(timeout=400):
                        link = cell.locator("a").first
                        if link.count() and link.is_visible(timeout=200):
                            return link
                        return cell
                except Exception:
                    continue
    return None


def _results_has_kn_link(page, kn: str) -> bool:
    """True, если в таблице результатов видна ссылка с КН."""
    return _find_kn_result_link(page, kn) is not None


def wait_for_results(page, kn: str, timeout_ms: int = RESULTS_WAIT_MS,
                     should_cancel=None, should_skip=None,
                     on_log=None) -> bool:
    """
    Ждёт после «Найти»: либо «Найдено результатов: 0», либо ссылку с КН.
    Не считает успехом устаревшую таблицу с другим КН.
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    step = RESULTS_POLL_MS
    waited = 0
    kn_norm = normalize_kn(kn)
    last_progress = -5000

    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip)
        _poll_assist_events(kn_norm, on_log=on_log)
        if has_captcha(page):
            return False
        if _object_modal_locator(page) is not None:
            return True

        if search_has_zero_results(page):
            log("  ответ: найдено 0 результатов")
            return False

        count = _read_results_count_dom(page)
        if count is not None and count >= 1:
            if _is_results_loading(page):
                if waited - last_progress >= 5000:
                    log(f"  загрузка таблицы ({waited // 1000} с)…")
                    last_progress = waited
            elif _results_rows_ready(page, kn_norm) or _find_kn_result_link(page, kn_norm):
                log(f"  найдено результатов: {count}")
                return True
            elif waited - last_progress >= 5000:
                log(f"  жду строку [data-test-id=row-0] с КН {kn_norm}…")
                last_progress = waited

        link = _find_kn_result_link(page, kn_norm)
        if link is not None:
            log(f"  ссылка с КН {kn_norm} появилась")
            return True

        if waited - last_progress >= 5000:
            if count is None:
                log(f"  жду «Найдено результатов» ({waited // 1000} с)…")
            elif count == 0:
                log("  жду обновления таблицы…")
            else:
                log(f"  жду таблицу результатов ({waited // 1000} с)…")
            last_progress = waited

        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step

    return False


def try_click_login(page, *, on_log=None) -> bool:
    """На странице без входа — клик «ВОЙТИ» (далее Госуслуги вручную)."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if is_lk_personal_form(page):
        return False
    try:
        for get_btn in (
            lambda: page.locator(".login-link").first,
            lambda: page.locator(".top-info__login").first,
            lambda: page.get_by_text("ВОЙТИ", exact=True).first,
            lambda: page.get_by_role("link", name=re.compile(r"^\s*войти\s*$", re.I)).first,
        ):
            btn = get_btn()
            if btn.count() and btn.is_visible(timeout=1500):
                log("  нажимаю «Войти» на странице Росreestr…")
                btn.click(force=True, timeout=5000)
                page.wait_for_timeout(2500)
                return True
    except Exception:
        pass
    return False


def _click_dropdown_option(page, label: str) -> bool:
    """Клик по пункту открытого выпадающего списка ЛK."""
    if not _wait_lk_option_menu(page, timeout_ms=1500):
        return False
    return _pick_lk_object_type_option(page, label)


def _fill_kn_input(page, kn: str) -> bool:
    """Поле 3: «Адрес или кадастровый номер» — ввод КН."""
    kn_norm = normalize_kn(kn)
    root = _search_form_root(page)

    attempts = [
        lambda: root.locator("#query").first,
        lambda: root.locator("input[name='query']").first,
        lambda: root.locator("input.realestateobjects-wrapper__option_input").first,
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
            el.press("Backspace")
            page.wait_for_timeout(150)
            el.type(kn, delay=35)
            page.wait_for_timeout(350)
            val = (el.input_value() or "").strip()
            if normalize_kn(val) == kn_norm or kn_norm.replace(":", "") in val.replace(":", ""):
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(200)
                except Exception:
                    pass
                return True
            el.click(force=True)
            el.fill(kn)
            page.wait_for_timeout(300)
            val = (el.input_value() or "").strip()
            if normalize_kn(val) == kn_norm or kn_norm.replace(":", "") in val.replace(":", ""):
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(200)
                except Exception:
                    pass
                return True
            el.press("Control+a")
            el.type(kn, delay=20)
            page.wait_for_timeout(200)
            val = (el.input_value() or "").strip()
            ok = normalize_kn(val) == kn_norm or kn_norm.replace(":", "") in val.replace(":", "")
            if ok:
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(200)
                except Exception:
                    pass
            return ok
        except Exception:
            continue
    return False


def _fill_lk_fields(page, kn: str, object_type: str, on_log=None) -> bool:
    """Заполняет форму ЛK: вид объекта (Enter) → КН → ждёт «Найти»."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if not _commit_lk_object_type(page, object_type, on_log=log):
        log(f"  не удалось зафиксировать «{object_type}» (Enter)")
        return False

    log(f"  поле 3 — кадастровый номер: {kn}")
    if not _fill_kn_input(page, kn):
        log("  не удалось ввести КН в поле «Адрес или кадастровый номер»")
        return False

    if _is_search_form_filled(page, kn):
        log("  форма заполнена — вид объекта и КН на месте")
        return True
    if _wait_find_button_enabled(page, timeout_ms=5000):
        log("  форма заполнена, кнопка «Найти» активна")
        return True
    if _object_type_committed(page) and _kn_input_matches(page, kn):
        log("  форма заполнена — можно нажимать «Найти»")
        return True
    log("  КН введён, но форма не готова — проверьте вид объекта")
    return False


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
]

SEARCH_BUTTON_SELECTORS = [
    "#realestateobjects-search",
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
    """True, если в форме поиска видна капча (гостевой режим)."""
    try:
        root = _search_form_root(page)
        for sel in CAPTCHA_SELECTORS:
            el = root.locator(sel).first
            if el.count() and el.is_visible(timeout=500):
                return True
    except Exception:
        pass
    return False


def _results_rows_ready(page, kn: str = "") -> bool:
    """Строки таблицы отрисованы и загрузка завершена."""
    if _is_results_loading(page):
        return False
    try:
        rows = _results_table_rows(page)
        if rows.count() == 0:
            return False
        if rows.first.is_visible(timeout=500):
            if kn:
                return _results_row_has_kn(page, kn)
            return True
    except Exception:
        pass
    return False


def _poll_search_response(page, *, before_text: str, kn: str,
                          timeout_ms: int = 12000,
                          should_cancel=None, should_skip=None) -> bool:
    step = 400
    waited = 0
    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip)
        if _search_response_started(page, before_text=before_text, kn=kn):
            return True
        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step
    return _search_response_started(page, before_text=before_text, kn=kn)


def _wait_find_button_clickable(page, timeout_ms: int = 15000) -> None:
    btn = page.locator("#realestateobjects-search").first
    step = 200
    waited = 0
    while waited < timeout_ms:
        try:
            if btn.count() and btn.is_visible(timeout=300):
                if btn.get_attribute("disabled") is None:
                    return
        except Exception:
            pass
        page.wait_for_timeout(step)
        waited += step


def _search_response_started(page, *, before_text: str = "", kn: str = "") -> bool:
    """True, если после отправки поиска появился ответ (0 или ≥1)."""
    if kn and _find_kn_result_link(page, kn):
        return True
    count = _read_results_count_dom(page)
    before_count = _parse_results_count(before_text) if before_text else None
    if count == 0:
        return True
    if count is not None and count >= 1:
        if kn and (_results_rows_ready(page, kn) or _results_row_has_kn(page, kn)):
            return True
        if count != before_count and kn and _kn_in_text(_results_area_text(page), kn):
            return True
    if search_has_zero_results(page):
        return True
    return False


def _submit_via_kn_enter(page) -> None:
    """Enter в поле КН — как при ручном вводе."""
    q = page.locator("#query").first
    if not q.count():
        q = page.locator("input[name='query']").first
    if not q.count():
        raise RuntimeError("поле КН не найдено")
    q.scroll_into_view_if_needed()
    q.click(force=True, timeout=3000)
    page.wait_for_timeout(150)
    q.press("Enter")


def _submit_via_find_button(page) -> None:
    """Клик по #realestateobjects-search в карточке формы."""
    _click_find_button_once(page)


def _click_result_target(page, target) -> None:
    """Левый клик по ссылке <a> с КН в таблице результатов."""
    target.scroll_into_view_if_needed()
    page.wait_for_timeout(200)

    link = target
    try:
        tag = (target.evaluate("el => (el.tagName || '').toUpperCase()") or "")
    except Exception:
        tag = ""
    if tag != "A":
        try:
            inner = target.locator("a").first
            if inner.count() and inner.is_visible(timeout=800):
                link = inner
        except Exception:
            pass

    handle = None
    try:
        handle = link.element_handle()
    except Exception:
        pass
    if handle:
        clicked = page.evaluate("""(el) => {
            const a = (el.tagName || '').toUpperCase() === 'A'
                ? el : (el.querySelector && el.querySelector('a')) || el;
            if (!a) return false;
            a.scrollIntoView({block: 'center', inline: 'nearest'});
            a.focus({preventScroll: true});
            for (const type of ['pointerdown', 'mousedown', 'mouseup', 'click']) {
                a.dispatchEvent(new MouseEvent(type, {
                    bubbles: true, cancelable: true, view: window,
                    button: 0, buttons: 1,
                }));
            }
            return true;
        }""", handle)
        if clicked:
            page.wait_for_timeout(300)
            return

    for click in (
        lambda: link.click(timeout=8000, button="left"),
        lambda: link.click(force=True, timeout=8000, button="left"),
    ):
        try:
            click()
            return
        except Exception:
            continue

    if handle:
        page.evaluate("""(el) => {
            if (!el) return;
            el.dispatchEvent(new MouseEvent('click', {
                bubbles: true, cancelable: true, view: window, button: 0,
            }));
        }""", handle)


def submit_lk_search(page, kn: str = "", *, on_log=None,
                     should_cancel=None, should_skip=None) -> bool:
    """
    Запускает поиск: клик «Найти» → Enter в КН.
    Ждёт span.realestateobjects-wrapper__results__total-count.
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    before = _results_area_text(page)
    if kn and not _is_search_form_filled(page, kn):
        log("  форма не готова к поиску")
        return False

    steps = (
        ("клик «Найти» (#realestateobjects-search)", _submit_via_find_button),
        ("Enter в поле КН", _submit_via_kn_enter),
    )
    for label, action in steps:
        _check_abort(should_cancel, should_skip)
        try:
            log(f"  поиск: {label}…")
            action(page)
            if _poll_search_response(
                    page, before_text=before, kn=kn, timeout_ms=12000,
                    should_cancel=should_cancel, should_skip=should_skip):
                _wait_results_loading(page, should_cancel=should_cancel,
                                      should_skip=should_skip)
                log(f"  поиск отправлен ({label})")
                _poll_assist_events(kn, on_log=on_log)
                return True
        except CollectionAborted:
            raise
        except Exception as exc:
            log(f"  {label} не сработал: {exc}")

    log("  не удалось запустить поиск")
    return False


def click_search(page, should_cancel=None, should_skip=None,
                 kn: str = "", on_log=None) -> bool:
    """Best-effort: «НАЙТИ» (ЛК) или «Сформировать запрос» (публичная)."""
    if kn and is_lk_personal_form(page):
        return submit_lk_search(
            page, kn, on_log=on_log,
            should_cancel=should_cancel, should_skip=should_skip)

    root = _search_form_root(page)
    force = bool(kn) and _is_search_form_filled(page, kn)
    try:
        btn = page.locator("#realestateobjects-search").first
        if btn.count() and btn.is_visible(timeout=1500):
            btn.scroll_into_view_if_needed()
            if force or btn.is_enabled():
                btn.click(force=True, timeout=3000)
                return True
    except CollectionAborted:
        raise
    except Exception:
        pass
    try:
        btn = root.get_by_role("button", name=re.compile(r"^\s*найти\s*$", re.I)).first
        if btn.is_visible(timeout=1000):
            btn.click(force=True, timeout=3000)
            return True
    except CollectionAborted:
        raise
    except Exception:
        pass
    for sel in SEARCH_BUTTON_SELECTORS:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click(timeout=3000)
                return True
        except Exception:
            continue
    try:
        _submit_via_kn_enter(page)
        return True
    except Exception:
        pass
    try:
        page.keyboard.press("Enter")
        return True
    except Exception:
        return False


def _object_modal_locator(page):
    """Модальное окно «Сведения об объекте» на lk.rosreestr.ru."""
    try:
        loc = page.locator(".realestate-object-modal").first
        if loc.is_visible(timeout=500):
            return loc
    except Exception:
        pass
    for sel in (
        "[role='dialog']",
        "[class*='modal'][class*='content']",
        "[class*='Modal']",
        "[class*='modal']",
    ):
        try:
            loc = page.locator(sel).filter(
                has_text=re.compile(
                    r"Сведения\s+об\s+объекте|Форма\s+собственности|"
                    r"Ранее\s+присвоенные\s+номера",
                    re.I))
            if loc.count() > 0:
                el = loc.last
                if el.is_visible(timeout=400):
                    return el
        except Exception:
            continue
    try:
        title = page.get_by_text("Сведения об объекте", exact=False).first
        if title.is_visible(timeout=400):
            modal = title.locator(
                "xpath=ancestor::*[contains(@class,'modal') or @role='dialog'][1]")
            if modal.count() and modal.first.is_visible(timeout=400):
                return modal.first
    except Exception:
        pass
    return None


def _collapse_actions_panel(page) -> None:
    """Сворачивает блок «Действия» (запрос выписки), мешающий чтению карточки."""
    modal = _object_modal_locator(page)
    if modal is None:
        return
    try:
        panel = modal.locator('[data-id="actions"]').first
        if panel.count() == 0:
            panel = modal.locator(".rros-ui-lib-panel").filter(
                has_text=re.compile(r"Действия", re.I)).first
        if panel.count() == 0:
            return
        poison = panel.get_by_text("Прошу предоставить", exact=False)
        if poison.count() and poison.first.is_visible(timeout=400):
            header = panel.locator(".rros-ui-lib-panel-header").first
            if header.is_visible(timeout=400):
                header.click(timeout=2000)
                page.wait_for_timeout(450)
    except Exception:
        pass


def _modal_text(page) -> str:
    """Текст карточки объекта (без блока «Действия»)."""
    card = _object_card_text(page)
    if card:
        return card
    parts: list[str] = []
    modal = _object_modal_locator(page)
    if modal is not None:
        try:
            t = modal.inner_text(timeout=3000)
            if t and not _is_actions_poison(t):
                parts.append(t)
        except Exception:
            pass
    return "\n".join(parts)


def _scroll_object_modal(page) -> None:
    """Прокрутка модалки к карточке объекта (минуя «Действия»)."""
    _collapse_actions_panel(page)
    modal = _object_modal_locator(page)
    if modal is None:
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        except Exception:
            pass
        return
    try:
        card = modal.locator(".build-card-wrapper").first
        if card.count() and card.is_visible(timeout=500):
            card.scroll_into_view_if_needed(timeout=3000)
            page.wait_for_timeout(300)
        for title in _OBJECT_CARD_SECTIONS:
            try:
                h = modal.get_by_role(
                    "heading",
                    name=re.compile(re.escape(title[:14]), re.I)).first
                if h.is_visible(timeout=400):
                    h.scroll_into_view_if_needed(timeout=2000)
                    page.wait_for_timeout(180)
            except Exception:
                continue
        modal.evaluate(
            "el => { const s = el.querySelector('.build-card-wrapper') || el;"
            " if (s) { s.scrollTop = s.scrollHeight; } }")
        page.wait_for_timeout(250)
    except Exception:
        pass


def _page_text(page) -> str:
    """Текст модалки (приоритет) + страницы + iframe."""
    parts: list[str] = []
    modal = _modal_text(page)
    if modal:
        parts.append(modal)
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
    kn_norm = normalize_kn(kn)
    needle = kn_norm.replace(":", "")
    if not needle:
        return False
    for m in CADASTRAL_RE.finditer(text or ""):
        if normalize_kn(m.group(0)).replace(":", "") == needle:
            return True
    compact = (text or "").replace(":", "").replace(" ", "")
    return compact == needle


def _modal_has_kn(page, kn: str) -> bool:
    return _modal_has_kn_strict(page, kn)


def wait_for_modal_card_ready(page, kn: str, timeout_ms: int = 35000,
                              should_cancel=None, should_skip=None,
                              on_log=None) -> bool:
    """Ждёт полной загрузки карточки: КН, поля и форма собственности."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    kn_norm = normalize_kn(kn)
    step = 500
    waited = 0
    last_scroll = -3000
    stable_since: int | None = None
    last_reason = ""
    last_reason_log = -5000

    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip)
        if not _object_modal_locator(page):
            stable_since = None
            _interruptible_wait(page, step, should_cancel, should_skip)
            waited += step
            continue
        if waited - last_scroll >= 1500:
            _scroll_object_modal(page)
            last_scroll = waited
        ok, reason = _modal_data_ready(page, kn_norm)
        last_reason = reason
        if ok:
            if stable_since is None:
                stable_since = waited
            elif waited - stable_since >= MODAL_READY_STABLE_MS:
                return True
        else:
            stable_since = None
            if reason and waited - last_reason_log >= 4000:
                log(f"  {reason}")
                last_reason_log = waited
        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step

    if last_reason:
        log(f"  {last_reason}")
    return _modal_data_ready(page, kn_norm)[0]


def wait_for_object_modal(page, timeout_ms: int = MODAL_WAIT_MS,
                          should_cancel=None, should_skip=None) -> bool:
    """Ждёт модальное окно «Сведения об объекте» после клика по КН."""
    step = 600
    waited = 0
    while waited < timeout_ms:
        _check_abort(should_cancel, should_skip)
        modal = _object_modal_locator(page)
        if modal is not None:
            try:
                card = modal.locator(".build-card-wrapper").first
                if card.count() and card.is_visible(timeout=800):
                    return True
                if modal.is_visible(timeout=400):
                    return True
            except Exception:
                return True
        if extract_form_from_text(_page_text(page)):
            return True
        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step
    return _object_modal_locator(page) is not None


def open_object_card_from_results(page, kn: str, *, on_log=None,
                                  should_cancel=None, should_skip=None) -> bool:
    """Клик по КН в таблице «Найдено результатов» → справка об объекте."""
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    if _object_modal_locator(page) is not None:
        log("  карточка объекта уже открыта")
        return True
    if extract_form_from_text(_page_text(page)):
        return True

    kn_norm = normalize_kn(kn)
    _wait_results_loading(page, should_cancel=should_cancel, should_skip=should_skip)
    link = _find_kn_result_link(page, kn_norm)
    if link is not None:
        try:
            _check_abort(should_cancel, should_skip)
            log(f"  открываю справку по строке/ссылке {kn_norm}")
            _click_result_target(page, link)
            if wait_for_object_modal(
                    page, should_cancel=should_cancel,
                    should_skip=should_skip):
                log("  модальное окно объекта загружено")
                return True
            _interruptible_wait(page, 2000, should_cancel, should_skip)
            if _object_modal_locator(page) is not None:
                return True
        except CollectionAborted:
            raise
        except Exception as exc:
            log(f"  клик по результату: {exc}")

    row = _find_results_row(page, kn_norm)
    if row is not None:
        try:
            log(f"  открываю строку таблицы {kn_norm}")
            _click_result_target(page, row)
            if wait_for_object_modal(page, should_cancel=should_cancel,
                                     should_skip=should_skip):
                return True
        except CollectionAborted:
            raise
        except Exception:
            pass

    kn_parts = (kn_norm, kn_norm.replace(":", "-"))
    root = _search_form_root(page)

    for part in kn_parts:
        for get_link in (
            lambda p=part: root.locator("[class*='cadNumber'] a").filter(
                has_text=re.compile(re.escape(p))).first,
            lambda p=part: root.locator(".realestateobjects-wrapper__results a").filter(
                has_text=re.compile(re.escape(p))).first,
            lambda p=part: root.locator("[class*='results'] tbody a, "
                                        "[class*='results'] a").filter(
                has_text=re.compile(re.escape(p))).first,
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
                if link.count() and link.is_visible(timeout=2500):
                    log(f"  открываю справку по ссылке {kn_norm}")
                    link.scroll_into_view_if_needed()
                    link.click(timeout=8000)
                    if wait_for_object_modal(
                            page, should_cancel=should_cancel,
                            should_skip=should_skip):
                        log("  модальное окно объекта загружено")
                        return True
                    _interruptible_wait(page, 2000, should_cancel, should_skip)
                    if _object_modal_locator(page) is not None:
                        return True
            except CollectionAborted:
                raise
            except Exception:
                continue

    click_result_if_needed(page, kn_norm)
    _interruptible_wait(page, 1500, should_cancel, should_skip)
    if wait_for_object_modal(page, should_cancel=should_cancel,
                             should_skip=should_skip):
        return True
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
        _scroll_object_modal(page)
        form = extract_form_from_modal(page)
        if not form:
            form = extract_form_from_text(_object_card_text(page) or _modal_text(page))
        if form:
            return form
        _interruptible_wait(page, step, should_cancel, should_skip)
        waited += step
    return ""


def read_page_data(page, kn: str = "") -> tuple[str, list[str], list[str]]:
    """Читает (форма собственности, обременения, ранее присвоенные номера)."""
    _scroll_object_modal(page)
    rows = _read_modal_subinfo(page)
    form = extract_form_from_modal(page)
    prev = extract_previous_numbers_from_modal(page)
    card_text = _object_card_text(page) or _modal_text(page)
    if not form:
        form = extract_form_from_text(card_text)
    if not prev:
        prev = extract_previous_numbers_from_text(card_text)
    encs = extract_encumbrances_from_modal(page)
    if not encs:
        encs = extract_encumbrances_from_text(card_text)

    ok, reason = validate_object_card_data(form, rows, kn)
    if (not ok or not form) and kn:
        waited_form = read_form_with_wait(page, timeout_ms=12000)
        if waited_form:
            form = waited_form
        rows = _read_modal_subinfo(page)
        if not prev:
            prev = extract_previous_numbers_from_modal(page)
        if not encs:
            encs = extract_encumbrances_from_modal(page)
        if not encs:
            encs = extract_encumbrances_from_text(
                _object_card_text(page) or _modal_text(page))
        ok, reason = validate_object_card_data(form, rows, kn)

    if not ok and rows:
        _scroll_object_modal(page)
        rows = _read_modal_subinfo(page)
        if not form:
            form = extract_form_from_modal(page) or extract_form_from_text(
                _object_card_text(page))
        if not prev:
            prev = extract_previous_numbers_from_modal(page)
        if not encs:
            encs = extract_encumbrances_from_modal(page)
        ok, reason = validate_object_card_data(form, rows, kn)

    if not ok and reason:
        return "", [], []

    if kn and not (form or "").strip():
        return "", [], []

    return form, encs, prev


def launch_rosreestr_context(pw, profile_dir: Path):
    """Запускает системный Microsoft Edge для lk.rosreestr.ru."""
    profile = str(profile_dir)
    base_kw = dict(
        headless=False,
        viewport={"width": 1280, "height": 900},
        ignore_https_errors=True,
        channel="msedge",
    )
    try:
        ctx = pw.chromium.launch_persistent_context(profile, **base_kw)
        return ctx, "Microsoft Edge"
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "Не удалось запустить Microsoft Edge для Росreestr. "
            "Установите Microsoft Edge (обычно уже есть в Windows 10/11)."
        ) from exc


def _close_object_modal(page) -> bool:
    """Закрывает модальное окно объекта (если открыто)."""
    modal = _object_modal_locator(page)
    if modal is None:
        return False
    for sel in (
        ".realestate-object-modal__btn",
        "button:has-text('Закрыть')",
        "[aria-label*='Закрыть']",
        "[class*='modal'] [class*='close']",
    ):
        try:
            btn = modal.locator(sel).first
            if btn.count() and btn.is_visible(timeout=800):
                btn.click(timeout=3000)
                page.wait_for_timeout(500)
                if _object_modal_locator(page) is None:
                    return True
        except Exception:
            continue
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(450)
    except Exception:
        pass
    return _object_modal_locator(page) is None


def prepare_for_next_object(page, *, on_log=None, timeout_ms: int = 30000) -> bool:
    """
    После сохранения данных: закрыть модалку и вернуть форму поиска
    (без полной перезагрузки, если возможно).
    """
    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    log("  подготовка к следующему объекту…")
    _close_object_modal(page)
    page.wait_for_timeout(400)
    _dismiss_dropdowns(page)
    try:
        goto_rosreestr_page(page, timeout=45000)
        page.wait_for_timeout(600)
    except Exception:
        pass
    ok = wait_for_search_form(page, timeout_ms=timeout_ms)
    if ok:
        log("  форма поиска восстановлена")
    else:
        log("  форма поиска не появилась — проверьте вкладку «Справочная информация online»")
    return ok


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
                   wait_assist: Callable[[str], bool] | None = None) -> tuple[str, list[str], list[str]]:
    """
    Открывает справочную, подставляет КН, запускает поиск и читает данные.
    Возвращает (форма, обременения, ранее присвоенные номера).
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
            return "", [], []
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
        form, encs, prev = read_page_data(page, kn)
        if form:
            log(f"  прочитано: «{form}»")
            if prev:
                log(f"  ранее присвоенные номера: {len(prev)} строк")
        else:
            rows = _read_modal_subinfo(page)
            _, reason = validate_object_card_data("", rows, kn)
            if reason:
                log(f"  {reason}")
            else:
                log("  форма не распознана — прокрутите модалку до «Общая информация»")
        return form, encs, prev

    if not ensure_search_page(page, on_log=log, timeout_ms=45000):
        if is_auth_page(page):
            log("  нужен вход в ЛК — завершите вход и откройте справочную online")
        else:
            log("  поле КН не найдено — откройте в браузере «Справочная информация online»")
        if not on_log:
            print(f"  Поле не найдено — откройте справочную и введите КН: {kn}")
        return "", [], []

    def _run_search_after_fill() -> tuple[str, list[str], list[str]]:
        _check_abort(should_cancel, should_skip)
        if has_captcha(page):
            if wait_captcha and wait_for_captcha(
                    page, on_log=on_log, should_cancel=should_cancel,
                    should_skip=should_skip):
                pass
            else:
                log("  на странице капча — решите её в браузере")
                return "", [], []
        log("  запускаю поиск…")
        if not click_search(page, should_cancel, should_skip, kn=kn, on_log=log):
            log("  поиск не запустился — «Найти» / Enter не сработали")
            return "", [], []
        _interruptible_wait(page, 600, should_cancel, should_skip)
        log("  жду «0 результатов» или ссылку с КН…")
        if not wait_for_results(page, kn, should_cancel=should_cancel,
                                should_skip=should_skip, on_log=log):
            if _results_has_kn_link(page, kn):
                log("  ссылка с КН в таблице — открываю карточку")
            elif search_has_zero_results(page):
                log("  найдено 0 результатов — попробую другой вид объекта")
                return "", [], []
            text = _page_text(page)
            if re.search(r"найдено\s+результатов:\s*0", text, re.I):
                log("  объект не найден (0 результатов) — проверьте вид объекта и КН")
            else:
                log("  результаты поиска не появились — проверьте интернет и вкладку ЛК")
            if not _results_has_kn_link(page, kn):
                return "", [], []
        log("  результаты найдены — открываю карточку объекта…")
        if not open_object_card_from_results(
                page, kn, on_log=log, should_cancel=should_cancel,
                should_skip=should_skip):
            log("  не удалось открыть справку по КН в таблице результатов")
            return "", [], []
        if not wait_for_object_modal(page, should_cancel=should_cancel,
                                     should_skip=should_skip):
            log("  модальное окно не загрузилось вовремя — пробую прочитать данные")
        _interruptible_wait(page, 1500, should_cancel, should_skip)
        log("  читаю данные карточки…")
        form, encs, prev = read_page_data(page, kn)
        if form:
            log(f"  форма собственности: «{form}»")
            if prev:
                log(f"  ранее присвоенные номера: {len(prev)} строк")
            if encs:
                log(f"  обременений: {len(encs)}")
        if not form:
            rows = _read_modal_subinfo(page)
            _, reason = validate_object_card_data("", rows, kn)
            if reason:
                log(f"  {reason}")
        return form, encs, prev

    if auto and is_lk_personal_form(page):
        if has_captcha(page):
            log("  на странице капча — дождитесь решения или введите символы")
            if wait_captcha and not wait_for_captcha(
                    page, timeout_ms=600000, on_log=on_log,
                    should_cancel=should_cancel,
                    should_skip=should_skip):
                log("  капча не решена — пропуск объекта")
                return "", [], []
        types = lk_auto_type_candidates(object_type)
        for ot in types:
            _check_abort(should_cancel, should_skip)
            if not wait_for_search_form(page, timeout_ms=8000,
                                        should_cancel=should_cancel,
                                        should_skip=should_skip):
                log("  форма поиска не загрузилась")
                continue
            log(f"  сценарий: «{ot}» → КН {kn} → «Найти» (до {LK_FIND_MAX_ATTEMPTS} раз)")
            form, encs, prev = lk_fetch_with_type(
                page, kn, ot, on_log=log,
                should_cancel=should_cancel, should_skip=should_skip)
            if form:
                return form, encs, prev
            if search_has_zero_results(page):
                log(f"  по «{ot}» объект не найден — другой вид объекта")
            else:
                log(f"  по «{ot}» справка не получена")
            if ot != types[-1]:
                _close_object_modal(page)
                page.wait_for_timeout(300)
                _clear_lk_object_type(page)
                _dismiss_dropdowns(page)
        log("  не удалось получить справку ни по одному виду объекта")
        return "", [], []

    if not try_autofill(page, kn, object_type=object_type, on_log=log):
        log("  не удалось заполнить форму поиска (вид объекта + КН)")
        return "", [], []

    if has_captcha(page):
        if wait_captcha and wait_for_captcha(
                page, on_log=on_log, should_cancel=should_cancel,
                should_skip=should_skip):
            pass
        else:
            log("  на странице капча — решите её в браузере")
            return "", [], []

    if auto:
        form, encs, prev = _run_search_after_fill()
        if not form and not encs:
            log("  справка по объекту не появилась (таймаут или капча)")
            return "", [], []
        if form:
            return form, encs, prev
        log("  справка открылась, но форма не распознана в тексте страницы")
        return "", encs, prev

    if has_captcha(page):
        print_captcha_hint(page)
        print("  Появилась капча — решите её и запустите поиск в браузере.")
    else:
        print("  Авто-чтение не дало результата — откройте карточку объекта вручную.")
    input("  Когда «Форма собственности» видна на странице — нажмите Enter… ")
    return read_page_data(page, kn)


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

        if args.no_login_wait:
            try:
                goto_rosreestr_page(page, timeout=60000)
                page = pick_rosreestr_page(ctx) or page
            except Exception as e:
                print(f"Не удалось открыть {ROSREESTR_ONLINE_URL}: {e}")
                print("Проверьте сеть и профиль браузера.")
            if is_auth_page(page) or not wait_for_search_form(page, timeout_ms=5000):
                print("=" * 60)
                print("Ожидание входа в ЛК (до 5 минут).")
                print("Войдите через Госуслуги как кадастровый инженер в окне Edge.")
                print("После появления формы поиска сбор начнётся автоматически.")
                for _ in range(60):
                    page = pick_rosreestr_page(ctx) or page
                    if navigate_to_reference_online(page):
                        break
                    if wait_for_search_form(page, timeout_ms=3000):
                        break
                    page.wait_for_timeout(5000)
                else:
                    print("Вход не выполнен — завершение.")
                    ctx.close()
                    return

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
            if not ensure_search_page(page, on_log=print if args.auto else None,
                                       timeout_ms=60000):
                print("  Форма поиска недоступна — пропуск объекта.")
                continue

            form, encs, prev_nums = auto_fetch_one(
                page, kn, auto=args.auto,
                object_type=load_object_type_hint(out_dir, kn),
                on_log=print if args.auto else None)

            if form:
                print(f"  Форма собственности: «{form}»")
                if encs:
                    print(f"  Обременения: {'; '.join(encs)}")
                if prev_nums:
                    print(f"  Ранее присвоенные номера: {len(prev_nums)} строк")
                if not args.auto:
                    ok = input("  Сохранить? [Enter=да / n=ввести вручную] ").strip().lower()
                    if ok == "n":
                        form = ""
            if not form and not args.auto:
                form = input("  Введите форму собственности вручную "
                             "(Enter — пропустить): ").strip()
            elif not form and args.auto:
                print("  Пропущено: форма не распознана (авто-режим).")
            if not encs and form and not args.auto:
                manual_enc = input("  Обременения (через ';', Enter — нет): ").strip()
                if manual_enc:
                    encs = [e.strip() for e in manual_enc.split(";") if e.strip()]

            if not form:
                print("  Пропущено (значение не сохранено).")
                continue

            save_ownership_cache(rr_cache, kn, form,
                                 "rosreestr-auto" if args.auto else "rosreestr-manual",
                                 encumbrances=encs, previous_numbers=prev_nums)
            append_to_forms_csv(forms_path, kn, form)
            saved += 1
            enc_note = f", обременений: {len(encs)}" if encs else ""
            print(f"  Сохранено: {kn} → «{form}»{enc_note}")

            if args.auto and n < len(todo):
                try:
                    prepare_for_next_object(page, on_log=print, timeout_ms=30000)
                except Exception:
                    pass
                print(f"  Пауза {int(PAUSE_BETWEEN_OBJECTS_SEC)} с перед следующим объектом…")
                try:
                    page.wait_for_timeout(int(PAUSE_BETWEEN_OBJECTS_SEC * 1000))
                except Exception:
                    pass

        print("\n" + "=" * 60)
        print(f"Готово. Сохранено форм: {saved} из {len(todo)}.")
        print("Теперь примените форму к отчёту и PDF (без платных запросов):")
        print(f"  В GUI: вкладка «Обслуживание» - «Пересобрать из кэша» для {args.output}")
        print(f"  Или CLI: python rebuild_from_cache.py -o {args.output} "
              f"--ownership-forms {forms_path}")
        ctx.close()


if __name__ == "__main__":
    main()
