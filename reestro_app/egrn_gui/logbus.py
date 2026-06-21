# -*- coding: utf-8 -*-
"""
Запись структурированного лога прогона в JSONL для последующего анализа.

Файл: <output>/logs/run_<YYYYMMDD_HHMMSS>.jsonl
Каждая строка — JSON-объект события: {ts, level, status, kn, message}.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


STATUS_LEVEL = {
    "ok": "INFO",
    "skipped": "INFO",
    "no_kn": "WARN",
    "network": "WARN",
    "failed": "WARN",
    "error": "ERROR",
}


class RunLogger:
    """Пишет события прогона в JSONL-файл (по событию на строку)."""

    def __init__(self, output_dir: Path, *, prefix: str = "run"):
        self.dir = Path(output_dir) / "logs"
        self.dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.path = self.dir / f"{prefix}_{stamp}.jsonl"
        self._fh = open(self.path, "a", encoding="utf-8")

    def write_event(self, event: dict):
        etype = event.get("type")
        if etype not in ("object", "log", "saved", "start", "done"):
            return
        status = event.get("status", "")
        level = event.get("level") or STATUS_LEVEL.get(status, "INFO")
        rec = {
            "ts": event.get("ts") or datetime.now().isoformat(timespec="seconds"),
            "level": level,
            "status": status or etype,
            "kn": event.get("kn", ""),
            "message": event.get("message", ""),
        }
        self._fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
        self._fh.flush()

    def write_line(self, message: str, *, level: str = "INFO", kn: str = "",
                   status: str = "log") -> None:
        """Произвольная строка журнала (для сбора Росreestr и др.)."""
        self.write_event({
            "type": "log",
            "level": level,
            "status": status,
            "kn": kn,
            "message": message,
        })

    def close(self):
        try:
            self._fh.close()
        except Exception:  # noqa: BLE001
            pass


def list_log_files(output_dir: Path) -> list[Path]:
    d = Path(output_dir) / "logs"
    if not d.exists():
        return []
    files = [p for p in d.glob("*.jsonl")
             if p.name.startswith(("run_", "rosreestr_"))]
    return sorted(files, key=lambda p: p.name, reverse=True)


def read_log(path: Path) -> list[dict]:
    out: list[dict] = []
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    out.append(json.loads(line))
                except ValueError:
                    continue
    except OSError:
        pass
    return out


def summarize(records: list[dict]) -> dict:
    summary = {"total": len(records), "by_status": {}, "by_level": {},
               "errors": 0, "first_ts": "", "last_ts": ""}
    for r in records:
        st = r.get("status", "")
        lv = r.get("level", "")
        summary["by_status"][st] = summary["by_status"].get(st, 0) + 1
        summary["by_level"][lv] = summary["by_level"].get(lv, 0) + 1
        if lv == "ERROR":
            summary["errors"] += 1
    if records:
        summary["first_ts"] = records[0].get("ts", "")
        summary["last_ts"] = records[-1].get("ts", "")
    return summary
