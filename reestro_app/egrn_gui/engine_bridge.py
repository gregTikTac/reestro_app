# -*- coding: utf-8 -*-
"""
Мост между GUI и движком reestro_parser.

- Находит и импортирует движок (reestro_parser) из соседней папки `reestro`
  (как в исходниках, так и в собранном .exe).
- Предоставляет высокоуровневые операции для интерфейса:
    * check_balance()        — проверка подключения и остатка единиц;
    * BatchRunner.run(...)    — пакетная/диапазонная/одиночная обработка с событиями,
                               отменой и защитой от дублей PDF/строк Excel.
- НЕ зависит от Qt: можно тестировать отдельно и переиспользовать в CLI.
"""
from __future__ import annotations

import json
import os
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

# --------------------------------------------------------------------------- #
# Поиск и импорт движка reestro_parser
# --------------------------------------------------------------------------- #


def _engine_dir() -> Path:
    """Каталог с reestro_parser.py (исходники или содержимое собранного .exe)."""
    candidates: list[Path] = []
    env = os.environ.get("REESTRO_ENGINE_DIR")
    if env:
        candidates.append(Path(env))
    # Внутри собранного приложения (PyInstaller кладёт данные в _MEIPASS)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "reestro")
    here = Path(__file__).resolve()
    # reestro_app/egrn_gui/engine_bridge.py -> parser/reestro
    candidates.append(here.parents[2] / "reestro")
    candidates.append(here.parents[1] / "reestro")
    for c in candidates:
        if (c / "reestro_parser.py").exists():
            return c
    raise ImportError(
        "Не найден движок reestro_parser.py. Ожидался каталог 'reestro' рядом с "
        "приложением или путь в переменной REESTRO_ENGINE_DIR."
    )


_ENGINE = None


def engine():
    """Ленивый импорт модуля reestro_parser."""
    global _ENGINE
    if _ENGINE is None:
        d = _engine_dir()
        if str(d) not in sys.path:
            sys.path.insert(0, str(d))
        import reestro_parser as _mod  # type: ignore
        _ENGINE = _mod
    return _ENGINE


# --------------------------------------------------------------------------- #
# Проверка подключения / баланса
# --------------------------------------------------------------------------- #


def check_balance(cfg: dict, timeout: int = 30) -> dict:
    """
    Запрос баланса. Возвращает {ok, status, units, message, raw}.
    Не бросает исключений — любые ошибки упаковываются в результат.
    """
    eng = engine()
    base = (cfg.get("baseUrl") or "https://api.kontur.ru").rstrip("/")
    try:
        client = eng.ReestroClient(cfg, pause=0, timeout=timeout, retries=1)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "status": None, "units": None,
                "message": f"Ошибка инициализации клиента: {exc}", "raw": ""}

    url = f"{base}/realty/billing/v1/balance"
    try:
        r = client.session.get(url, timeout=timeout)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "status": None, "units": None,
                "message": f"Нет связи с сервером: {exc.__class__.__name__} ({exc})",
                "raw": ""}

    text = r.text or ""
    if r.status_code == 401:
        return {"ok": False, "status": 401, "units": None,
                "message": "401 — неверный apiKey или orgId.", "raw": text[:500]}
    if r.status_code == 402:
        return {"ok": False, "status": 402, "units": None,
                "message": "402 — недостаточно средств на балансе.", "raw": text[:500]}
    if r.status_code != 200:
        return {"ok": False, "status": r.status_code, "units": None,
                "message": f"HTTP {r.status_code}", "raw": text[:500]}

    units = None
    try:
        data = r.json()
        items = data.get("items") if isinstance(data, dict) else None
        if isinstance(items, list):
            for it in items:
                code = (it.get("type") or {}).get("code", "")
                if code == "address_api_open_data":
                    units = (it.get("balance") or {}).get("value")
                    break
    except Exception:  # noqa: BLE001
        pass

    msg = "Подключение успешно."
    if units is not None:
        msg += f" Доступно единиц (address_api_open_data): {units}."
    return {"ok": True, "status": 200, "units": units, "message": msg,
            "raw": text[:500]}


# --------------------------------------------------------------------------- #
# Параметры и события обработки
# --------------------------------------------------------------------------- #


@dataclass
class BatchParams:
    config: dict
    output_dir: Path
    report_name: str = "report.xlsx"
    input_path: Path | None = None
    single_kns: list[str] = field(default_factory=list)
    range_from: int | None = None          # 1-индекс по строкам входного файла
    range_to: int | None = None
    limit: int | None = None
    pause: float = 2.0
    timeout: int = 120
    retries: int = 5
    save_every: int = 5
    force: bool = False
    use_cache: bool = True


# Тип колбэка события: dict с ключом 'type'
EventCb = Callable[[dict], None]
StopFlag = Callable[[], bool]


class Cancelled(Exception):
    pass


class BatchRunner:
    """
    Оркестратор обработки поверх примитивов движка. Гарантирует:
      - 1 объект (КН) = 1 PDF = 1 блок строк в report.xlsx (антидубли через kn_index.json);
      - устойчивость к обрыву сети (объект пропускается без записи, повторяем позже);
      - автосохранение report каждые N объектов;
      - кооперативную отмену между объектами.
    """

    def __init__(self, params: BatchParams, on_event: EventCb | None = None,
                 should_stop: StopFlag | None = None):
        self.p = params
        self._emit = on_event or (lambda e: None)
        self._stop = should_stop or (lambda: False)
        self.eng = engine()

    # -- служебное ------------------------------------------------------- #
    def _ev(self, etype: str, **kw):
        e = {"type": etype, "ts": datetime.now().isoformat(timespec="seconds")}
        e.update(kw)
        self._emit(e)

    def _check_stop(self):
        if self._stop():
            raise Cancelled()

    def _load_index(self, path: Path) -> dict:
        if path.exists():
            try:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                return data if isinstance(data, dict) else {}
            except (OSError, ValueError):
                return {}
        return {}

    def _save_index(self, path: Path, index: dict):
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(index, f, ensure_ascii=False, indent=2)

    def _build_items(self):
        eng = self.eng
        if self.p.single_kns:
            items = []
            for kn in self.p.single_kns:
                kn = (kn or "").strip()
                if not kn:
                    continue
                row = eng.InputRow()
                row.cadastral = kn
                items.append(row)
            return items
        if not self.p.input_path:
            return []
        items = eng.read_input(Path(self.p.input_path))
        lo = (self.p.range_from or 1) - 1
        hi = self.p.range_to if self.p.range_to else len(items)
        lo = max(0, lo)
        hi = min(len(items), hi)
        return items[lo:hi]

    # -- основной прогон ------------------------------------------------- #
    def run(self) -> dict:
        eng = self.eng
        p = self.p
        out_dir = Path(p.output_dir)
        pdf_dir = out_dir / "pdf"
        cache_dir = out_dir / "cache" / "json"
        rr_cache_dir = out_dir / "cache" / "rosreestr"
        out_dir.mkdir(parents=True, exist_ok=True)
        pdf_dir.mkdir(parents=True, exist_ok=True)
        cache_dir.mkdir(parents=True, exist_ok=True)
        rr_cache_dir.mkdir(parents=True, exist_ok=True)
        index_path = out_dir / "cache" / "kn_index.json"
        report_path = out_dir / p.report_name

        # справочник форм собственности (если есть рядом с движком)
        try:
            ov_path = _engine_dir() / "input" / "ownership_forms.csv"
            overrides = eng.load_ownership_overrides(ov_path)
        except Exception:  # noqa: BLE001
            overrides = {}

        client = eng.ReestroClient(p.config, pause=p.pause,
                                   timeout=p.timeout, retries=p.retries)

        processed_kn, existing_rows = eng.load_existing_report(report_path, pdf_dir)
        kn_index = self._load_index(index_path)

        items = self._build_items()
        total = len(items)
        today = datetime.now().strftime("%d.%m.%Y")

        ok = failed = skipped_no_kn = skipped_already = network = 0
        objects_done = 0
        from requests.exceptions import RequestException  # локальный импорт

        def ownership_for(kn: str) -> str:
            return eng.get_ownership_form(kn, rr_cache_dir, overrides, fetch=False)

        def save_report(force: bool = False):
            if not force and objects_done == 0:
                return
            merged = eng.patch_report_rights_columns(existing_rows)
            try:
                eng.write_report_xlsx(merged, {
                    "total": total, "ok": ok, "failed": failed,
                    "skipped_no_kn": skipped_no_kn,
                    "skipped_already": skipped_already,
                    "network_skipped": network,
                }, report_path)
                self._ev("saved", path=str(report_path), rows=len(merged))
            except PermissionError:
                self._ev("log", level="ERROR",
                         message=f"Не удалось сохранить {report_path.name}: файл открыт "
                                 f"в Excel. Закройте его — сохраню при следующем шаге.")

        self._ev("start", total=total, report=str(report_path),
                 mode=("single" if p.single_kns else "batch"))

        try:
            for n, row in enumerate(items, start=1):
                self._check_stop()
                m = eng.CADASTRAL_RE.search(row.cadastral or "")
                cadastral = m.group(0) if m else ""
                label = row.cadastral or row.full_address or f"объект {n}"

                if not cadastral:
                    skipped_no_kn += 1
                    row.cadastral = ""
                    xrows = eng.build_xlsx_rows(row, None, [], "", "", today)
                    existing_rows.extend(xrows)
                    self._ev("object", index=n, total=total, kn="", status="no_kn",
                             message=f"{label}: нет кадастрового номера — без PDF")
                    self._ev("progress", index=n, total=total)
                    continue

                row.cadastral = cadastral
                kn_norm = eng.normalize_kn(cadastral)

                if kn_norm in processed_kn and not p.force:
                    skipped_already += 1
                    self._ev("object", index=n, total=total, kn=kn_norm,
                             status="skipped",
                             message=f"{kn_norm}: уже обработан — пропуск")
                    self._ev("progress", index=n, total=total)
                    continue

                # антидубли: переиспользуем extract_id (а значит имя PDF) для КН
                reuse_id = kn_index.get(kn_norm, {}).get("extract_id") if isinstance(
                    kn_index.get(kn_norm), dict) else None
                extract_id = reuse_id or eng.new_extract_id()
                extract_date = today
                pdf_name = eng.pdf_name_for_extract(extract_id)
                pdf_path = pdf_dir / pdf_name

                # данные: кэш или API
                info = None
                from_cache = False
                if p.use_cache and not p.force:
                    info = eng.load_api_cache(cache_dir, kn_norm)
                    from_cache = info is not None

                if not from_cache:
                    self._check_stop()
                    try:
                        resp = client.object_info(cadastral)
                    except RequestException as exc:
                        network += 1
                        self._ev("object", index=n, total=total, kn=kn_norm,
                                 status="network",
                                 message=f"{kn_norm}: обрыв сети "
                                         f"({exc.__class__.__name__}) — не записан, "
                                         f"повторите позже")
                        self._ev("progress", index=n, total=total)
                        continue

                    if resp.status_code == 200:
                        try:
                            info = resp.json()
                        except ValueError:
                            info = None
                        if info and info.get("cadastralNumber"):
                            if p.use_cache:
                                eng.save_api_cache(cache_dir, kn_norm, info, resp,
                                                   input_row=row)
                        else:
                            info = None
                            reason = "пустой ответ (объект без КН в ЕГРН)"
                    else:
                        info = None
                        reason = self._http_reason(resp)

                # удаляем старые строки этого КН (антидубль), затем пишем заново
                existing_rows[:] = eng._drop_kn_rows(existing_rows, kn_norm)

                if info and info.get("cadastralNumber"):
                    rights = eng.extract_rights(info)
                    own = ownership_for(kn_norm)
                    encs = eng.load_object_encumbrances(rr_cache_dir, kn_norm)
                    try:
                        eng.generate_pdf(info, rights, row, pdf_path,
                                         ownership_form=own,
                                         encumbrances_override=encs)
                    except Exception as exc:  # noqa: BLE001
                        failed += 1
                        self._ev("object", index=n, total=total, kn=kn_norm,
                                 status="error",
                                 message=f"{kn_norm}: ошибка PDF — {exc}")
                        self._ev("progress", index=n, total=total)
                        continue
                    xrows = eng.build_xlsx_rows(row, info, rights, pdf_name,
                                               extract_id, extract_date,
                                               ownership_form=own)
                    existing_rows.extend(xrows)
                    processed_kn.add(kn_norm)
                    kn_index[kn_norm] = {"extract_id": extract_id, "pdf": pdf_name}
                    ok += 1
                    objects_done += 1
                    self._ev("object", index=n, total=total, kn=kn_norm, status="ok",
                             message=f"{kn_norm}: готово (прав: {len(rights)})"
                                     + (" [из кэша]" if from_cache else ""))
                else:
                    failed += 1
                    objects_done += 1
                    full_reason = locals().get("reason", "нет сведений")
                    eng.generate_pdf_not_found(row, cadastral, full_reason, pdf_path)
                    xrows = eng.build_xlsx_rows(row, None, [], pdf_name,
                                               extract_id, extract_date)
                    existing_rows.extend(xrows)
                    processed_kn.add(kn_norm)
                    kn_index[kn_norm] = {"extract_id": extract_id, "pdf": pdf_name}
                    self._ev("object", index=n, total=total, kn=kn_norm,
                             status="failed", message=f"{kn_norm}: {full_reason}")

                self._ev("progress", index=n, total=total)

                if p.save_every and objects_done % p.save_every == 0:
                    save_report()
                    self._save_index(index_path, kn_index)

                if p.limit and (ok + failed) >= p.limit:
                    self._ev("log", level="INFO",
                             message=f"Достигнут лимит {p.limit} новых объектов.")
                    break

        except Cancelled:
            self._ev("log", level="WARN", message="Остановлено пользователем.")
        finally:
            save_report(force=True)
            self._save_index(index_path, kn_index)

        result = {
            "total": total, "ok": ok, "failed": failed,
            "skipped_no_kn": skipped_no_kn, "skipped_already": skipped_already,
            "network": network, "report": str(report_path),
            "pdf_dir": str(pdf_dir),
        }
        self._ev("done", **result)
        return result

    def _http_reason(self, resp) -> str:
        sc = resp.status_code
        if sc == 404:
            return "объект не найден в ЕГРН (HTTP 404)"
        if sc == 401:
            return "ошибка авторизации (HTTP 401) — проверьте apiKey/orgId"
        if sc == 402:
            return "недостаточно средств на балансе (HTTP 402)"
        if sc == 400:
            try:
                err = resp.json()
            except ValueError:
                err = {}
            msg = (err or {}).get("message", "")
            if "баланс" in msg.lower() or (err or {}).get("code") == "validation":
                return "нет доступных единиц address_api_open_data (HTTP 400)"
            return f"некорректный запрос (HTTP 400) — {msg}"
        return f"ошибка API: HTTP {sc}"
