# -*- coding: utf-8 -*-
"""Запись событий браузера (клики, клавиши) для ручного режима сбора."""
from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Callable


_RECORDER_JS = """
(() => {
  if (window.__egrnRecorderInstalled) return;
  window.__egrnRecorderInstalled = true;
  window.__egrnEvents = window.__egrnEvents || [];
  const push = (payload) => {
    window.__egrnEvents.push({ ts: Date.now(), ...payload });
    if (window.__egrnEvents.length > 500) window.__egrnEvents.shift();
  };
  const targetInfo = (t) => {
    if (!t || !t.tagName) return {};
    return {
      tag: t.tagName,
      id: t.id || "",
      name: t.name || "",
      cls: (t.className && t.className.toString().slice(0, 160)) || "",
      text: ((t.innerText || t.value || t.placeholder || "") + "").slice(0, 120),
      href: t.href || "",
    };
  };
  document.addEventListener("click", (e) => {
    push({ kind: "click", ...targetInfo(e.target), x: e.clientX, y: e.clientY });
  }, true);
  document.addEventListener("keydown", (e) => {
    push({ kind: "keydown", key: e.key, code: e.code, ...targetInfo(e.target) });
  }, true);
  document.addEventListener("input", (e) => {
    push({ kind: "input", ...targetInfo(e.target) });
  }, true);
  document.addEventListener("change", (e) => {
    push({ kind: "change", ...targetInfo(e.target) });
  }, true);
})();
"""


class BrowserRecorder:
    """Снимает клики/клавиши со страницы Росreestr в jsonl."""

    def __init__(
        self,
        page,
        *,
        out_path: Path | None = None,
        on_event: Callable[[dict], None] | None = None,
    ):
        self.page = page
        self.out_path = out_path
        self.on_event = on_event
        self._installed = False
        self._last_url = ""

    def install(self) -> None:
        if self._installed:
            return
        try:
            self.page.evaluate(_RECORDER_JS)
        except Exception:
            pass
        try:
            self.page.add_init_script(_RECORDER_JS)
        except Exception:
            pass
        self._installed = True
        self._last_url = self.page.url or ""

    def poll(self, *, kn: str = "") -> list[dict]:
        """Забирает накопленные события и дописывает в лог."""
        events: list[dict] = []
        try:
            self.install()
            raw = self.page.evaluate(
                "() => { const e = window.__egrnEvents || []; "
                "window.__egrnEvents = []; return e; }")
            if raw:
                events.extend(raw)
        except Exception:
            pass
        url = ""
        try:
            url = self.page.url or ""
        except Exception:
            pass
        if url and url != self._last_url:
            events.append({
                "ts": int(time.time() * 1000),
                "kind": "navigate",
                "url": url,
            })
            self._last_url = url
        out: list[dict] = []
        for ev in events:
            row = {
                "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "kn": kn,
                **{k: v for k, v in ev.items() if k != "ts"},
            }
            out.append(row)
            if self.on_event:
                self.on_event(row)
            if self.out_path:
                with self.out_path.open("a", encoding="utf-8") as f:
                    f.write(json.dumps(row, ensure_ascii=False) + "\n")
        return out

    def snapshot_form(self, *, kn: str = "", label: str = "") -> None:
        """Снимок состояния формы поиска для анализа."""
        try:
            import fetch_rosreestr as fr  # noqa: WPS433

            row = {
                "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "kind": "snapshot",
                "label": label,
                "kn": kn,
                "url": self.page.url or "",
                "object_type": fr._object_type_display(self.page),
                "obj_types": fr._objtypes_hidden_value(self.page),
                "query": fr._kn_query_value(self.page),
                "results": fr._results_area_text(self.page)[:500],
            }
            if self.on_event:
                self.on_event(row)
            if self.out_path:
                with self.out_path.open("a", encoding="utf-8") as f:
                    f.write(json.dumps(row, ensure_ascii=False) + "\n")
        except Exception:
            pass
