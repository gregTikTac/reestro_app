# -*- coding: utf-8 -*-
"""Вкладка «Результаты»: сводка последнего прогона и быстрые действия."""
from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QGroupBox,
    QGridLayout, QMessageBox,
)


class ResultsTab(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self._last: dict | None = None

        root = QVBoxLayout(self)
        box = QGroupBox("Последний прогон")
        grid = QGridLayout(box)
        self.cards: dict[str, QLabel] = {}
        labels = [
            ("total", "Всего объектов"),
            ("ok", "Сформировано отчётов"),
            ("failed", "Без сведений"),
            ("skipped_already", "Пропущено (готовые)"),
            ("skipped_no_kn", "Без КН"),
            ("network", "Обрыв сети"),
        ]
        for i, (key, title) in enumerate(labels):
            t = QLabel(title)
            v = QLabel("—")
            v.setStyleSheet("font-size: 20px; font-weight: bold;")
            grid.addWidget(t, (i // 3) * 2, i % 3)
            grid.addWidget(v, (i // 3) * 2 + 1, i % 3)
            self.cards[key] = v
        root.addWidget(box)

        self.path_label = QLabel("Отчёт: —")
        self.path_label.setWordWrap(True)
        root.addWidget(self.path_label)

        btns = QHBoxLayout()
        self.open_report = QPushButton("Открыть отчёт")
        self.open_report.clicked.connect(self._open_report)
        self.open_folder = QPushButton("Открыть папку результатов")
        self.open_folder.clicked.connect(self._open_folder)
        btns.addWidget(self.open_report)
        btns.addWidget(self.open_folder)
        btns.addStretch(1)
        root.addLayout(btns)
        root.addStretch(1)

    def show_result(self, res: dict):
        self._last = res
        for key, lbl in self.cards.items():
            lbl.setText(str(res.get(key, 0)))
        self.path_label.setText(f"Отчёт: {res.get('report', '—')}")
        name = Path(res.get("report", "")).name
        self.open_report.setText(f"Открыть {name}" if name else "Открыть отчёт")

    def _open_report(self):
        if self._last and Path(self._last["report"]).exists():
            os.startfile(self._last["report"])  # noqa: S606
        else:
            QMessageBox.information(self, "Результаты", "Отчёт ещё не сформирован.")

    def _open_folder(self):
        if self._last:
            folder = Path(self._last["report"]).parent
            if folder.exists():
                os.startfile(str(folder))  # noqa: S606
                return
        QMessageBox.information(self, "Результаты", "Папка ещё не создана.")
