# -*- coding: utf-8 -*-
"""Вкладка «Одиночный объект»: обработка по КН вручную."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPlainTextEdit, QPushButton, QLabel,
    QLineEdit, QFileDialog, QMessageBox, QGroupBox, QFormLayout,
)

from ..engine_bridge import BatchParams
from ..paths import resolve_output_root
from ..worker import BatchWorker
from .tab_batch import STATUS_COLOR, _esc, build_report_name


class SingleTab(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.worker: BatchWorker | None = None

        root = QVBoxLayout(self)

        box = QGroupBox("Кадастровые номера (по одному в строке)")
        v = QVBoxLayout(box)
        self.kn_edit = QPlainTextEdit()
        self.kn_edit.setPlaceholderText(
            "По одному КН в строке, формат ЧЧ:ЧЧ:ЧЧЧЧЧЧЧ:Ч, напр.:\n"
            "43:40:000864:351\n77:01:0001001:1037")
        self.kn_edit.setToolTip("Кадастровый номер: две цифры : две цифры : "
                                "квартал : номер объекта.")
        v.addWidget(self.kn_edit)
        root.addWidget(box)

        out_box = QGroupBox("Папка результатов")
        form = QFormLayout(out_box)
        row = QHBoxLayout()
        self.output_edit = QLineEdit(win.settings.get("last_output", ""))
        self.output_edit.setPlaceholderText(r"напр. C:\Реестро\output")
        btn = QPushButton("Обзор…")
        btn.clicked.connect(self.pick_output)
        row.addWidget(self.output_edit)
        row.addWidget(btn)
        w = QWidget()
        w.setLayout(row)
        form.addRow("Папка:", w)
        hint = QLabel(
            "Имя отчёта задаётся на вкладке «Пакетная обработка» "
            "(поле «Имя файла отчёта»). PDF всегда в подпапке pdf\\.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #555; font-size: 11px;")
        form.addRow("", hint)
        root.addWidget(out_box)

        ctl = QHBoxLayout()
        self.run_btn = QPushButton("Обработать")
        self.run_btn.clicked.connect(self.on_run)
        self.open_pdf_btn = QPushButton("Открыть папку PDF")
        self.open_pdf_btn.clicked.connect(self.open_pdf_dir)
        ctl.addWidget(self.run_btn)
        ctl.addWidget(self.open_pdf_btn)
        ctl.addStretch(1)
        root.addLayout(ctl)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        root.addWidget(self.log, 1)

    def pick_output(self):
        path = QFileDialog.getExistingDirectory(
            self, "Папка результатов (корень output)", self.output_edit.text())
        if path:
            self.output_edit.setText(str(resolve_output_root(path)))

    def open_pdf_dir(self):
        out = self.output_edit.text().strip()
        if out:
            pdf = Path(out) / "pdf"
            if pdf.exists():
                import os
                os.startfile(str(pdf))  # noqa: S606 (Windows)
            else:
                QMessageBox.information(self, "PDF", "Папка pdf ещё не создана.")

    def on_run(self):
        cfg = self.win.current_config()
        if not cfg.get("apiKey") or not cfg.get("orgId"):
            QMessageBox.warning(self, "Обработка",
                                "Заполните apiKey/orgId на вкладке «Подключение».")
            return
        kns = [ln.strip() for ln in self.kn_edit.toPlainText().splitlines()
               if ln.strip()]
        if not kns:
            QMessageBox.warning(self, "Обработка", "Введите хотя бы один КН.")
            return
        out = self.output_edit.text().strip()
        if not out:
            QMessageBox.warning(self, "Обработка", "Выберите папку результатов.")
            return
        out_root = resolve_output_root(out)
        if out_root != Path(out).resolve():
            self.output_edit.setText(str(out_root))
        if not self.win.set_running(True):
            return

        s = self.win.settings
        report_name = build_report_name(
            auto=bool(s.get("auto_name", True)),
            manual=s.get("report_name", ""),
            input_path=s.get("last_input", ""),
            fallback_stem="Одиночный",
        )

        params = BatchParams(
            config=cfg,
            output_dir=out_root,
            report_name=report_name,
            single_kns=kns,
            pause=float(cfg.get("pause", 2.0)),
            timeout=int(cfg.get("timeout", 120)),
            retries=int(cfg.get("retries", 5)),
            save_every=1,
            force=True,  # одиночная обработка всегда обновляет объект
        )
        self.win.settings["last_output"] = str(out_root)
        self.win.save_all()

        self.log.clear()
        self.run_btn.setEnabled(False)
        self._append(f"Обрабатываю {len(kns)} объект(ов)…", "#000")

        self.worker = BatchWorker(params, self)
        self.worker.event.connect(self.on_event)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def on_event(self, e: dict):
        if e.get("type") == "object":
            self._append(e.get("message", ""),
                         STATUS_COLOR.get(e.get("status", ""), "#000"))
        elif e.get("type") == "log":
            self._append(e.get("message", ""), "#827717")

    def on_finished(self, res: dict):
        self.win.set_running(False)
        self.run_btn.setEnabled(True)
        self.worker = None
        self._append(
            f"Готово. OK: {res['ok']}, без сведений: {res['failed']}, "
            f"обрыв сети: {res['network']}.", "#1b5e20")
        self.win.results_tab.show_result(res)
        self.win.logs_tab.refresh_files(Path(res["report"]).parent)

    def on_failed(self, msg: str):
        self.win.set_running(False)
        self.run_btn.setEnabled(True)
        self.worker = None
        self._append("ОШИБКА: " + msg, "#b71c1c")
        QMessageBox.critical(self, "Ошибка", msg)

    def _append(self, text: str, color: str = "#000"):
        if text:
            self.log.appendHtml(f'<span style="color:{color}">{_esc(text)}</span>')
