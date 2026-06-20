# -*- coding: utf-8 -*-
"""
Рабочий поток (QThread) для пакетной обработки.

Запускает BatchRunner в фоне, транслирует его события в сигналы Qt и
поддерживает кооперативную отмену (Стоп). Параллельно пишет JSONL-лог прогона.
"""
from __future__ import annotations

import threading
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from .engine_bridge import BatchParams, BatchRunner
from .logbus import RunLogger


class BatchWorker(QThread):
    event = Signal(dict)        # любое событие движка
    finished_ok = Signal(dict)  # итоговая статистика
    failed = Signal(str)        # фатальная ошибка

    def __init__(self, params: BatchParams, parent=None):
        super().__init__(parent)
        self._params = params
        self._stop_flag = threading.Event()
        self._logger: RunLogger | None = None

    def request_stop(self):
        self._stop_flag.set()

    def _on_event(self, e: dict):
        if self._logger:
            self._logger.write_event(e)
        self.event.emit(e)

    def run(self):
        try:
            self._logger = RunLogger(Path(self._params.output_dir))
        except Exception:  # noqa: BLE001
            self._logger = None
        try:
            runner = BatchRunner(
                self._params,
                on_event=self._on_event,
                should_stop=self._stop_flag.is_set,
            )
            result = runner.run()
            self.finished_ok.emit(result)
        except Exception as exc:  # noqa: BLE001
            import traceback
            self.failed.emit(f"{exc}\n\n{traceback.format_exc()}")
        finally:
            if self._logger:
                self._logger.close()
