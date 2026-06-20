# -*- coding: utf-8 -*-
"""Вкладка «Пакетная обработка»."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit, QPushButton,
    QFileDialog, QSpinBox, QCheckBox, QGroupBox, QLabel, QProgressBar,
    QPlainTextEdit, QMessageBox,
)

from ..engine_bridge import BatchParams, engine
from ..paths import resolve_output_root
from ..worker import BatchWorker


def build_report_name(
    *,
    auto: bool,
    manual: str = "",
    input_path: str = "",
    fallback_stem: str = "report",
) -> str:
    """Имя Excel-отчёта: авто Отчёт_<файл>_<дата>.xlsx или заданное вручную."""
    if auto:
        stem = Path(input_path).stem if input_path else fallback_stem
        date = datetime.now().strftime("%Y%m%d")
        name = f"Отчёт_{stem}_{date}.xlsx"
    else:
        name = manual.strip() or "report.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


STATUS_COLOR = {
    "ok": "#1b5e20",
    "failed": "#e65100",
    "network": "#bf360c",
    "no_kn": "#827717",
    "skipped": "#555555",
    "error": "#b71c1c",
}


class BatchTab(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.worker: BatchWorker | None = None
        s = win.settings

        root = QVBoxLayout(self)

        # --- источник и назначение ---
        io_box = QGroupBox("Файлы")
        form = QFormLayout(io_box)

        in_row = QHBoxLayout()
        self.input_edit = QLineEdit(s.get("last_input", ""))
        self.input_edit.setPlaceholderText(
            r"напр. C:\Реестро\Запрос.xlsx  (столбцы: ЕГРН / Адрес полностью)")
        self.input_edit.setToolTip(
            "Таблица .xlsx/.csv с кадастровыми номерами в столбце «ЕГРН» "
            "(или «Кадастровый номер»). Строки без КН попадут только в Excel.")
        in_btn = QPushButton("Обзор…")
        in_btn.clicked.connect(self.pick_input)
        in_row.addWidget(self.input_edit)
        in_row.addWidget(in_btn)
        form.addRow("Входная таблица (КН):", self._wrap(in_row))
        self.input_info = QLabel("")
        form.addRow("", self.input_info)

        out_row = QHBoxLayout()
        self.output_edit = QLineEdit(s.get("last_output", ""))
        self.output_edit.setPlaceholderText(
            r"напр. C:\Реестро\output  (здесь появятся pdf\, report.xlsx, logs\)")
        self.output_edit.setToolTip(
            "Папка для результатов. Запускайте всегда с одной и той же папкой — "
            "уже обработанные объекты не запрашиваются повторно.")
        out_btn = QPushButton("Обзор…")
        out_btn.clicked.connect(self.pick_output)
        out_row.addWidget(self.output_edit)
        out_row.addWidget(out_btn)
        form.addRow("Папка результатов:", self._wrap(out_row))
        out_hint = QLabel(
            "Корень output: здесь появятся Excel-отчёт, папки pdf\\, cache\\, logs\\. "
            "Не выбирайте подпапку pdf\\.")
        out_hint.setWordWrap(True)
        out_hint.setStyleSheet("color: #555; font-size: 11px;")
        form.addRow("", out_hint)

        name_row = QHBoxLayout()
        self.auto_name = QCheckBox("Авто-имя")
        self.auto_name.setChecked(bool(s.get("auto_name", True)))
        self.auto_name.toggled.connect(self._toggle_name)
        self.report_name = QLineEdit(s.get("report_name", ""))
        self.report_name.setPlaceholderText("напр. Отчёт_Запрос_20260619.xlsx")
        self.report_name.setToolTip(
            "Имя файла отчёта. При «Авто-имя» формируется как "
            "Отчёт_<входной файл>_<дата>.xlsx.")
        name_row.addWidget(self.auto_name)
        name_row.addWidget(self.report_name)
        form.addRow("Имя файла отчёта:", self._wrap(name_row))
        root.addWidget(io_box)
        self._toggle_name(self.auto_name.isChecked())

        # --- параметры прогона ---
        par_box = QGroupBox("Параметры прогона")
        pform = QFormLayout(par_box)
        rng = QHBoxLayout()
        self.range_from = QSpinBox()
        self.range_from.setRange(0, 1_000_000)
        self.range_from.setValue(int(s.get("range_from") or 0))
        self.range_from.setSpecialValueText("с начала")
        self.range_from.setToolTip("Номер первой строки входного файла (напр. 30). "
                                   "0 = с начала.")
        self.range_to = QSpinBox()
        self.range_to.setRange(0, 1_000_000)
        self.range_to.setValue(int(s.get("range_to") or 0))
        self.range_to.setSpecialValueText("до конца")
        self.range_to.setToolTip("Номер последней строки (напр. 75). 0 = до конца. "
                                 "Пример: с 30 по 75 — объекты 30…75.")
        rng.addWidget(QLabel("с №"))
        rng.addWidget(self.range_from)
        rng.addWidget(QLabel("по №"))
        rng.addWidget(self.range_to)
        rng.addStretch(1)
        pform.addRow("Диапазон объектов:", self._wrap(rng))

        self.limit = QSpinBox()
        self.limit.setRange(0, 1_000_000)
        self.limit.setSpecialValueText("без лимита")
        self.limit.setValue(int(s.get("limit") or 0))
        self.limit.setToolTip("Сколько новых объектов запросить через API за прогон "
                              "(уже готовые в счёт не идут). 0 = без лимита.")
        pform.addRow("Лимит новых объектов:", self.limit)

        self.force = QCheckBox("Перезаписать уже обработанные (--force)")
        pform.addRow("", self.force)
        root.addWidget(par_box)

        # --- управление ---
        ctl = QHBoxLayout()
        self.start_btn = QPushButton("Старт")
        self.start_btn.clicked.connect(self.on_start)
        self.stop_btn = QPushButton("Стоп")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.on_stop)
        ctl.addWidget(self.start_btn)
        ctl.addWidget(self.stop_btn)
        ctl.addStretch(1)
        root.addLayout(ctl)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        root.addWidget(self.progress)
        self.counters = QLabel("OK: 0   Без сведений: 0   Пропущено: 0   "
                               "Обрыв сети: 0   Без КН: 0")
        root.addWidget(self.counters)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(5000)
        root.addWidget(self.log, 1)

        self._reset_counts()
        self._refresh_input_info()

    # -- helpers -------------------------------------------------------- #
    def _wrap(self, layout):
        w = QWidget()
        w.setLayout(layout)
        return w

    def _toggle_name(self, auto: bool):
        self.report_name.setEnabled(not auto)

    def _reset_counts(self):
        self.c = {"ok": 0, "failed": 0, "skipped": 0, "network": 0, "no_kn": 0}

    def _update_counts(self):
        self.counters.setText(
            f"OK: {self.c['ok']}   Без сведений: {self.c['failed']}   "
            f"Пропущено: {self.c['skipped']}   Обрыв сети: {self.c['network']}   "
            f"Без КН: {self.c['no_kn']}")

    def pick_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите таблицу", self.input_edit.text(),
            "Таблицы (*.xlsx *.xlsm *.csv)")
        if path:
            self.input_edit.setText(path)
            self._refresh_input_info()

    def pick_output(self):
        path = QFileDialog.getExistingDirectory(
            self, "Папка результатов (корень output)", self.output_edit.text())
        if path:
            self.output_edit.setText(str(resolve_output_root(path)))

    def _refresh_input_info(self):
        p = self.input_edit.text().strip()
        if not p or not Path(p).exists():
            self.input_info.setText("")
            return
        try:
            eng = engine()
            items = eng.read_input(Path(p))
            with_kn = sum(1 for r in items
                          if eng.CADASTRAL_RE.search(r.cadastral or ""))
            self.input_info.setText(
                f"строк: {len(items)}   с кадастровым номером: {with_kn}")
        except Exception as exc:  # noqa: BLE001
            self.input_info.setText(f"не удалось прочитать: {exc}")


    # -- запуск --------------------------------------------------------- #
    def on_start(self):
        cfg = self.win.current_config()
        if not cfg.get("apiKey") or not cfg.get("orgId"):
            QMessageBox.warning(self, "Старт",
                                "Сначала заполните apiKey/orgId на вкладке «Подключение».")
            return
        inp = self.input_edit.text().strip()
        out = self.output_edit.text().strip()
        if not inp or not Path(inp).exists():
            QMessageBox.warning(self, "Старт", "Выберите существующую входную таблицу.")
            return
        if not out:
            QMessageBox.warning(self, "Старт", "Выберите папку результатов.")
            return
        out_root = resolve_output_root(out)
        if out_root != Path(out).resolve():
            self.output_edit.setText(str(out_root))
        if not self.win.set_running(True):
            return

        report_name = build_report_name(
            auto=self.auto_name.isChecked(),
            manual=self.report_name.text(),
            input_path=inp,
        )

        params = BatchParams(
            config=cfg,
            output_dir=out_root,
            report_name=report_name,
            input_path=Path(inp),
            range_from=self.range_from.value() or None,
            range_to=self.range_to.value() or None,
            limit=self.limit.value() or None,
            pause=float(cfg.get("pause", 2.0)),
            timeout=int(cfg.get("timeout", 120)),
            retries=int(cfg.get("retries", 5)),
            save_every=int(cfg.get("save_every", 5)),
            force=self.force.isChecked(),
        )

        # сохранить настройки
        s = self.win.settings
        s["last_input"] = inp
        s["last_output"] = str(out_root)
        s["auto_name"] = self.auto_name.isChecked()
        s["report_name"] = self.report_name.text().strip()
        s["range_from"] = self.range_from.value()
        s["range_to"] = self.range_to.value()
        s["limit"] = self.limit.value()
        self.win.save_all()

        self.log.clear()
        self._reset_counts()
        self._update_counts()
        self.progress.setValue(0)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self.worker = BatchWorker(params, self)
        self.worker.event.connect(self.on_event)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()
        self._append("Старт обработки…", "#000")

    def on_stop(self):
        if self.worker:
            self.worker.request_stop()
            self._append("Остановка после текущего объекта…", "#827717")
            self.stop_btn.setEnabled(False)

    def on_event(self, e: dict):
        t = e.get("type")
        if t == "start":
            self.progress.setMaximum(max(1, e.get("total", 1)))
            self._append(f"Объектов к обработке: {e.get('total', 0)} → "
                         f"{e.get('report', '')}", "#000")
        elif t == "object":
            st = e.get("status", "")
            if st in self.c:
                self.c[st] += 1
                self._update_counts()
            self._append(e.get("message", ""), STATUS_COLOR.get(st, "#000"))
        elif t == "progress":
            self.progress.setValue(e.get("index", 0))
        elif t == "saved":
            self._append(f"  report сохранён ({e.get('rows', 0)} строк)", "#1565c0")
        elif t == "log":
            lvl = e.get("level", "INFO")
            color = {"ERROR": "#b71c1c", "WARN": "#827717"}.get(lvl, "#000")
            self._append(e.get("message", ""), color)

    def on_finished(self, res: dict):
        self._finish()
        self._append(
            f"Готово. OK: {res['ok']}, без сведений: {res['failed']}, "
            f"пропущено: {res['skipped_already']}, обрыв сети: {res['network']}, "
            f"без КН: {res['skipped_no_kn']}.", "#1b5e20")
        self.win.results_tab.show_result(res)
        self.win.logs_tab.refresh_files(Path(res["report"]).parent)

    def on_failed(self, msg: str):
        self._finish()
        self._append("ОШИБКА: " + msg, "#b71c1c")
        QMessageBox.critical(self, "Ошибка обработки", msg)

    def _finish(self):
        self.win.set_running(False)
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.worker = None

    def _append(self, text: str, color: str = "#000"):
        if not text:
            return
        self.log.appendHtml(
            f'<span style="color:{color}">{_esc(text)}</span>')


def _esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
