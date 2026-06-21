# -*- coding: utf-8 -*-
"""
Автоматический сбор формы собственности и обременений с lk.rosreestr.ru.

Идея: оператор входит через Госуслуги ОДИН раз (код подтверждения 2FA вводится
в браузере — это невозможно автоматизировать). Сессия сохраняется в постоянном
профиле. Дальше по списку КН данные собираются БЕЗ участия оператора и
складываются в кэш <output>/cache/rosreestr/{kn}.json, который применяется к
report.xlsx и PDF.

Модуль не зависит от Qt: общается через колбэки (лог/прогресс/нужен вход/отмена).
"""
from __future__ import annotations

import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .engine_bridge import _engine_dir
from .playwright_setup import configure_env, ensure_chromium, browsers_dir
from .settings import app_dir


def _import_engine_modules():
    """Импортирует reestro_parser и fetch_rosreestr из каталога движка."""
    d = _engine_dir()
    if str(d) not in sys.path:
        sys.path.insert(0, str(d))
    import reestro_parser as rp
    import fetch_rosreestr as fr
    return rp, fr


def profile_dir() -> Path:
    """Постоянный профиль браузера (сессия входа). Один на приложение."""
    return app_dir() / "browser_profile_rosreestr"


def reset_browser_profile() -> Path:
    """Переименовывает повреждённый профиль браузера (сессию входа придётся повторить)."""
    import shutil
    from datetime import datetime

    prof = profile_dir()
    if not prof.exists():
        return prof
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = prof.parent / f"{prof.name}_backup_{stamp}"
    if backup.exists():
        shutil.rmtree(backup, ignore_errors=True)
    prof.rename(backup)
    prof.mkdir(parents=True, exist_ok=True)
    return backup


@dataclass
class CollectResult:
    total: int = 0
    saved: int = 0
    skipped_existing: int = 0
    failed: int = 0
    cancelled: bool = False
    error: str = ""


def collect_kns_from_output(out_root: Path, *, only_with_rights: bool = True,
                            redo: bool = False) -> list[str]:
    """Список КН для сбора из report.xlsx (минус уже собранные, если не redo)."""
    rp, _ = _import_engine_modules()
    report = out_root / "report.xlsx"
    rr_cache = out_root / "cache" / "rosreestr"
    if not report.exists():
        return []
    _, rows = rp.load_existing_report(report, out_root / "pdf")
    known = set()
    if not redo and rr_cache.is_dir():
        known = {rp.normalize_kn(p.stem) for p in rr_cache.glob("*.json")}
    out: list[str] = []
    seen = set()
    for r in rows:
        if not rp.CADASTRAL_RE.search(rp._str(r[1])):
            continue
        kn = rp.normalize_kn(r[1])
        if kn in seen:
            continue
        has_right = rp._str(r[31]) and rp._str(r[31]) != "данные отсутствуют"
        if only_with_rights and not has_right:
            continue
        if kn in known:
            seen.add(kn)
            continue
        seen.add(kn)
        out.append(kn)
    return out


def _apply_cache_to_report(out_root: Path, on_log: Callable[[str], None]) -> None:
    """Пересобирает report.xlsx и PDF из кэша (без API)."""
    rp, _ = _import_engine_modules()
    import cleanup_output as cu

    report = out_root / "report.xlsx"
    pdf_dir = out_root / "pdf"
    existing: list[list] = []
    if report.exists():
        _, existing = rp.load_existing_report(report, pdf_dir)
    tz = _engine_dir() / "TZ" / "Запрос.xlsx"
    inputs = [tz] if tz.exists() else []
    rows = cu.rebuild_report_from_cache(out_root, inputs, None, existing, regen_pdf=True)
    uniq = len({rp.normalize_kn(r[1]) for r in rows if rp.normalize_kn(r[1])})
    rp.write_report_xlsx(rows, {
        "total": uniq, "ok": uniq, "failed": 0,
        "skipped_no_kn": 0, "skipped_already": 0,
    }, report)
    on_log(f"  применено к report.xlsx и PDF ({uniq} КН)")


def run_collection(
    out_root: Path,
    *,
    only_with_rights: bool = True,
    redo: bool = False,
    assisted: bool = False,
    apply_after_each: bool = False,
    on_log: Callable[[str], None] = lambda s: None,
    on_progress: Callable[[int, int], None] = lambda i, n: None,
    on_need_login: Callable[[], None] = lambda: None,
    is_login_done: Callable[[], bool] = lambda: True,
    on_form_ready: Callable[[], None] = lambda: None,
    should_cancel: Callable[[], bool] = lambda: False,
    should_skip: Callable[[], bool] = lambda: False,
    consume_skip: Callable[[], None] = lambda: None,
    wait_assist: Callable[[str], bool] | None = None,
) -> CollectResult:
    """
    Полный цикл сбора.

    on_need_login() — показать UI входа (кнопка «Продолжить» изначально неактивна).
    is_login_done() — True после нажатия «Вход выполнен — продолжить».
    on_form_ready() — справочная online открыта, можно нажимать «Продолжить».
    should_cancel() — проверка флага отмены между объектами.
    """
    res = CollectResult()
    configure_env()
    if not ensure_chromium(on_log):
        on_log("Встроенный Chromium не найден — попробую системный Edge или Chrome.")

    rp, fr = _import_engine_modules()
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:  # noqa: BLE001
        res.error = f"Playwright недоступен: {exc}"
        return res

    rr_cache = out_root / "cache" / "rosreestr"
    rr_cache.mkdir(parents=True, exist_ok=True)
    forms_csv = _engine_dir() / "input" / "ownership_forms.csv"

    kns = collect_kns_from_output(out_root, only_with_rights=only_with_rights,
                                  redo=redo)
    res.total = len(kns)
    if not kns:
        on_log("Нет КН для сбора (всё уже собрано или нет report.xlsx с правами).")
        return res

    prof = profile_dir()
    prof.mkdir(parents=True, exist_ok=True)
    on_log(f"Профиль браузера (вход сохраняется): {prof}")
    on_log(f"Браузеры Playwright: {browsers_dir()}")
    on_log(f"К сбору: {len(kns)} объектов.")

    with sync_playwright() as pw:
        try:
            ctx, browser_label = fr.launch_rosreestr_context(pw, prof)
        except Exception as exc:  # noqa: BLE001
            res.error = str(exc)
            on_log(res.error)
            return res
        on_log(f"Браузер: {browser_label}")
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        # Разовый вход через Госуслуги (2FA в браузере).
        try:
            fr.goto_rosreestr_page(page, timeout=60000)
        except Exception as exc:  # noqa: BLE001
            on_log(f"Не удалось открыть Росреестр: {exc}")
            on_log("Проверьте интернет и российский IP.")
            on_log("Если белый экран — нажмите «Сбросить профиль браузера» "
                   "и запустите сбор снова (вход в Госуслуги повторится).")
        on_log("Если требуется — войдите через Госуслуги, выберите пользователя.")
        on_log("После входа программа сама откроет «Справочную информацию online».")
        on_log("Кнопка «Вход выполнен — продолжить» станет активной, когда форма "
               "поиска будет готова.")

        on_need_login()
        form_ready_notified = False
        login_deadline = time.time() + 20 * 60
        while time.time() < login_deadline:
            if should_cancel():
                ctx.close()
                res.cancelled = True
                return res
            if is_login_done():
                break
            page = fr.pick_rosreestr_page(ctx) or page
            try:
                url = (page.url or "").lower()
                if "rosreestr.ru" in url and not fr.is_on_reference_page(page):
                    fr.navigate_to_reference_online(page, on_log=on_log)
                if fr.is_on_reference_page(page):
                    if not form_ready_notified:
                        on_form_ready()
                        form_ready_notified = True
            except Exception:
                pass
            end = time.time() + 2
            while time.time() < end:
                if should_cancel():
                    ctx.close()
                    res.cancelled = True
                    return res
                time.sleep(0.25)
        else:
            ctx.close()
            res.error = "Таймаут ожидания входа (20 мин)."
            on_log(res.error)
            return res

        if should_cancel() or not is_login_done():
            ctx.close()
            res.cancelled = True
            return res

        page = fr.pick_rosreestr_page(ctx) or page
        on_log("Подтверждён вход — проверяю справочную…")
        fr.navigate_to_reference_online(page, on_log=on_log)
        on_log("Проверяю форму поиска после входа…")
        if not fr.ensure_search_page(page, on_log=on_log, timeout_ms=60000):
            ctx.close()
            res.error = (
                "Форма поиска Росreestr не найдена. Откройте в браузере "
                "«Справочная информация online», дождитесь поля для КН и "
                "запустите сбор снова.")
            on_log(res.error)
            return res
        on_log("Форма поиска готова — начинаю обход объектов.")
        if assisted:
            on_log("Режим: проверочный (с оператором). Автозаполнение отключено.")

        for i, kn in enumerate(kns, 1):
            if should_cancel():
                res.cancelled = True
                break
            on_progress(i, len(kns))
            on_log(f"[{i}/{len(kns)}] {kn}")
            page = fr.pick_rosreestr_page(ctx) or page
            try:
                if not assisted:
                    try:
                        fr.navigate_to_reference_online(page, on_log=on_log)
                    except Exception:
                        pass
                obj_type = fr.load_object_type_hint(out_root, kn)
                form, encs = fr.auto_fetch_one(
                    page, kn, auto=not assisted, wait_captcha=True,
                    object_type=obj_type,
                    should_cancel=should_cancel,
                    should_skip=should_skip,
                    assisted=assisted,
                    wait_assist=wait_assist,
                    on_log=lambda m: on_log(f"  {m}" if not m.startswith("  ") else m),
                )
            except fr.CollectionAborted as exc:
                if exc.reason == "cancel":
                    res.cancelled = True
                    break
                consume_skip()
                on_log("  объект пропущен — следующий")
                try:
                    fr.goto_rosreestr_page(page, timeout=30000)
                    fr._interruptible_wait(page, 800, should_cancel, should_skip)
                except Exception:
                    pass
                continue
            except Exception as exc:  # noqa: BLE001
                on_log(f"  ошибка: {exc}")
                res.failed += 1
                continue

            if not form:
                on_log("  форма не распознана — пропуск (см. сообщение выше)")
                res.failed += 1
                try:
                    fr.goto_rosreestr_page(page, timeout=30000)
                    fr._interruptible_wait(page, 800, should_cancel, should_skip)
                except fr.CollectionAborted:
                    res.cancelled = True
                    break
                except Exception:
                    pass
                try:
                    fr._interruptible_wait(
                        page, int(fr.PAUSE_BETWEEN_OBJECTS_SEC * 1000),
                        should_cancel, should_skip)
                except fr.CollectionAborted:
                    if should_cancel():
                        res.cancelled = True
                        break
                continue

            rp.save_ownership_cache(rr_cache, kn, form,
                                    "rosreestr-assist" if assisted else "rosreestr-auto",
                                    encumbrances=encs)
            try:
                fr.append_to_forms_csv(forms_csv, kn, form)
            except Exception:  # noqa: BLE001
                pass
            res.saved += 1
            note = f"; обременений: {len(encs)}" if encs else ""
            on_log(f"  сохранено: «{form}»{note}")
            if apply_after_each:
                try:
                    _apply_cache_to_report(out_root, on_log)
                except Exception as exc:  # noqa: BLE001
                    on_log(f"  не удалось применить к Excel/PDF: {exc}")
            if assisted:
                try:
                    fr.goto_rosreestr_page(page, timeout=30000)
                    fr._interruptible_wait(page, 800, should_cancel, should_skip)
                except fr.CollectionAborted:
                    if should_cancel():
                        res.cancelled = True
                        break
                except Exception:
                    pass
                continue
            try:
                fr.goto_rosreestr_page(page, timeout=30000)
                fr._interruptible_wait(page, 800, should_cancel, should_skip)
            except fr.CollectionAborted:
                if should_cancel():
                    res.cancelled = True
                    break
            except Exception:
                pass
            try:
                fr._interruptible_wait(
                    page, int(fr.PAUSE_BETWEEN_OBJECTS_SEC * 1000),
                    should_cancel, should_skip)
            except fr.CollectionAborted:
                if should_cancel():
                    res.cancelled = True
                    break

        ctx.close()

    return res
