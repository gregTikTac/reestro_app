# -*- coding: utf-8 -*-
"""
Вкладка «Анализ логов»: просмотр, фильтрация и сводка по JSONL-логам прогонов.
"""
from __future__ import annotations

import csv
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QPushButton, QLabel,
    QLineEdit, QTableWidget, QTableWidgetItem, QFileDialog, QGroupBox,
    QHeaderView, QMessageBox, QApplication,
)

from ..logbus import list_log_files, read_log, summarize
from ..engine_bridge import BatchParams
from ..worker import BatchWorker


# статусы, которые имеет смысл прогонять заново
RETRYABLE = {"network", "error", "failed"}


LEVEL_COLOR = {
    "ERROR": QColor("#ffcdd2"),
    "WARN": QColor("#fff9c4"),
    "INFO": QColor("#ffffff"),
}

STATUS_TITLES = {
    "ok": "OK", "failed": "Без сведений", "network": "Обрыв сети",
    "skipped": "Пропущено", "no_kn": "Без КН", "error": "Ошибка",
    "saved": "Сохранение", "start": "Старт", "done": "Итог",
}


class LogsTab(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self._records: list[dict] = []
        self._current_dir: Path | None = None
        self._retry_worker: BatchWorker | None = None

        root = QVBoxLayout(self)

        # выбор файла
        top = QHBoxLayout()
        self.file_combo = QComboBox()
        self.file_combo.setMinimumWidth(360)
        self.file_combo.currentIndexChanged.connect(self._load_selected)
        self.refresh_btn = QPushButton("Обновить список")
        self.refresh_btn.clicked.connect(lambda: self.refresh_files(self._current_dir))
        self.open_btn = QPushButton("Открыть файл…")
        self.open_btn.clicked.connect(self._open_file)
        self.folder_btn = QPushButton("Папка результатов…")
        self.folder_btn.clicked.connect(self._pick_folder)
        top.addWidget(QLabel("Лог прогона:"))
        top.addWidget(self.file_combo, 1)
        top.addWidget(self.refresh_btn)
        top.addWidget(self.open_btn)
        top.addWidget(self.folder_btn)
        root.addLayout(top)

        # сводка
        self.summary = QLabel("Лог не загружен.")
        self.summary.setWordWrap(True)
        self.summary.setStyleSheet("padding: 6px; background: #f5f5f5;")
        root.addWidget(self.summary)

        # фильтры
        filt = QGroupBox("Фильтры")
        f = QHBoxLayout(filt)
        self.level_combo = QComboBox()
        self.level_combo.addItems(["все уровни", "INFO", "WARN", "ERROR"])
        self.level_combo.currentIndexChanged.connect(self._apply_filter)
        self.status_combo = QComboBox()
        self.status_combo.addItem("все статусы", "")
        for k, v in STATUS_TITLES.items():
            self.status_combo.addItem(v, k)
        self.status_combo.currentIndexChanged.connect(self._apply_filter)
        self.search = QLineEdit()
        self.search.setPlaceholderText("поиск по КН или тексту…")
        self.search.textChanged.connect(self._apply_filter)
        f.addWidget(QLabel("Уровень:"))
        f.addWidget(self.level_combo)
        f.addWidget(QLabel("Статус:"))
        f.addWidget(self.status_combo)
        f.addWidget(self.search, 1)
        self.export_btn = QPushButton("Экспорт в CSV")
        self.export_btn.clicked.connect(self._export_csv)
        f.addWidget(self.export_btn)
        root.addWidget(filt)

        # таблица
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Время", "Уровень", "КН", "Сообщение"])
        self.table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 70)
        self.table.setColumnWidth(2, 170)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.cellDoubleClicked.connect(self._copy_kn)
        root.addWidget(self.table, 1)

        bottom = QHBoxLayout()
        self.hint = QLabel("Двойной клик по строке — скопировать КН.")
        self.hint.setStyleSheet("color: #777;")
        bottom.addWidget(self.hint)
        bottom.addStretch(1)
        self.retry_btn = QPushButton("Повторить проблемные (сеть/ошибки)")
        self.retry_btn.setToolTip(
            "Заново обработать объекты со статусом «обрыв сети», «ошибка» или "
            "«без сведений» из текущего лога. Запросы идут в ту же папку результатов.")
        self.retry_btn.clicked.connect(self._retry_failed)
        bottom.addWidget(self.retry_btn)
        root.addLayout(bottom)

        # начальная загрузка
        out = win.settings.get("last_output", "")
        if out:
            self.refresh_files(Path(out))

    # -- загрузка файлов ------------------------------------------------ #
    def refresh_files(self, out_dir: Path | None):
        if out_dir:
            self._current_dir = Path(out_dir)
        self.file_combo.blockSignals(True)
        self.file_combo.clear()
        files = list_log_files(self._current_dir) if self._current_dir else []
        for p in files:
            self.file_combo.addItem(p.name, str(p))
        self.file_combo.blockSignals(False)
        if files:
            self.file_combo.setCurrentIndex(0)
            self._load_selected()
        else:
            self._records = []
            self.summary.setText("Логи не найдены в выбранной папке.")
            self.table.setRowCount(0)

    def _pick_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Папка результатов")
        if path:
            self.refresh_files(Path(path))

    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть лог", "", "Логи (*.jsonl);;Все файлы (*.*)")
        if path:
            self._current_dir = Path(path).parent.parent
            self._records = read_log(Path(path))
            self._refresh_summary(Path(path).name)
            self._apply_filter()

    def _load_selected(self):
        data = self.file_combo.currentData()
        if not data:
            return
        self._records = read_log(Path(data))
        self._refresh_summary(self.file_combo.currentText())
        self._apply_filter()

    def _refresh_summary(self, name: str):
        s = summarize(self._records)
        bs = s["by_status"]
        parts = [f"{STATUS_TITLES.get(k, k)}: {v}" for k, v in bs.items()
                 if k in STATUS_TITLES]
        dur = ""
        if s["first_ts"] and s["last_ts"]:
            dur = f"   с {s['first_ts'][11:19]} по {s['last_ts'][11:19]}"
        self.summary.setText(
            f"<b>{name}</b> — событий: {s['total']}, ошибок: {s['errors']}.   "
            + "   ".join(parts) + dur)

    # -- фильтрация и отображение --------------------------------------- #
    def _filtered(self) -> list[dict]:
        lvl = self.level_combo.currentText()
        status = self.status_combo.currentData()
        q = self.search.text().strip().lower()
        out = []
        for r in self._records:
            if lvl != "все уровни" and r.get("level") != lvl:
                continue
            if status and r.get("status") != status:
                continue
            if q and q not in (r.get("kn", "") + " " + r.get("message", "")).lower():
                continue
            out.append(r)
        return out

    def _apply_filter(self):
        rows = self._filtered()
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            ts = r.get("ts", "")
            ts_short = ts[11:19] if len(ts) >= 19 else ts
            cells = [ts_short, r.get("level", ""), r.get("kn", ""),
                     r.get("message", "")]
            color = LEVEL_COLOR.get(r.get("level", "INFO"), QColor("#ffffff"))
            for j, val in enumerate(cells):
                item = QTableWidgetItem(val)
                item.setBackground(color)
                self.table.setItem(i, j, item)

    def _problem_kns(self) -> list[str]:
        seen: set[str] = set()
        out: list[str] = []
        for r in self._records:
            if r.get("status") in RETRYABLE:
                kn = (r.get("kn") or "").strip()
                if kn and kn not in seen:
                    seen.add(kn)
                    out.append(kn)
        return out

    def _retry_failed(self):
        if not self._current_dir:
            QMessageBox.warning(self, "Повтор", "Сначала выберите папку с логами.")
            return
        cfg = self.win.current_config()
        if not cfg.get("apiKey") or not cfg.get("orgId"):
            QMessageBox.warning(self, "Повтор",
                                "Заполните apiKey/orgId на вкладке «Подключение».")
            return
        kns = self._problem_kns()
        if not kns:
            QMessageBox.information(self, "Повтор",
                                   "В текущем логе нет проблемных объектов "
                                   "(сеть/ошибки/без сведений).")
            return
        if QMessageBox.question(
                self, "Повтор",
                f"Повторно обработать {len(kns)} проблемн(ых) объект(ов) "
                f"в папку:\n{self._current_dir}?") != QMessageBox.Yes:
            return
        if not self.win.set_running(True):
            return

        params = BatchParams(
            config=cfg,
            output_dir=self._current_dir,
            report_name="report.xlsx",
            single_kns=kns,
            pause=float(cfg.get("pause", 2.0)),
            timeout=int(cfg.get("timeout", 120)),
            retries=int(cfg.get("retries", 5)),
            save_every=1,
            force=True,
        )
        self.retry_btn.setEnabled(False)
        self.win.statusBar().showMessage(
            f"Повтор {len(kns)} объект(ов)…")
        self._retry_worker = BatchWorker(params, self)
        self._retry_worker.finished_ok.connect(self._on_retry_done)
        self._retry_worker.failed.connect(self._on_retry_failed)
        self._retry_worker.start()

    def _on_retry_done(self, res: dict):
        self.win.set_running(False)
        self.retry_btn.setEnabled(True)
        self._retry_worker = None
        self.win.statusBar().showMessage(
            f"Повтор завершён: OK {res['ok']}, без сведений {res['failed']}, "
            f"обрыв сети {res['network']}.", 6000)
        self.win.results_tab.show_result(res)
        self.refresh_files(self._current_dir)

    def _on_retry_failed(self, msg: str):
        self.win.set_running(False)
        self.retry_btn.setEnabled(True)
        self._retry_worker = None
        QMessageBox.critical(self, "Повтор — ошибка", msg)

    def _copy_kn(self, row: int, _col: int):
        item = self.table.item(row, 2)
        if item and item.text():
            QApplication.clipboard().setText(item.text())
            self.win.statusBar().showMessage(f"Скопирован КН: {item.text()}", 3000)

    def _export_csv(self):
        rows = self._filtered()
        if not rows:
            QMessageBox.information(self, "Экспорт", "Нет данных для экспорта.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить CSV", "log_export.csv", "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Время", "Уровень", "Статус", "КН", "Сообщение"])
            for r in rows:
                w.writerow([r.get("ts", ""), r.get("level", ""),
                            r.get("status", ""), r.get("kn", ""),
                            r.get("message", "")])
        self.win.statusBar().showMessage(f"Экспортировано: {path}", 4000)
