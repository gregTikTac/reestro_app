# -*- coding: utf-8 -*-
"""Главное окно с вкладками."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QMainWindow, QTabWidget, QMessageBox, QLabel

from . import APP_NAME, __version__
from . import settings as cfgio


def _asset(name: str) -> Path:
    """Путь к ресурсу (в исходниках и в собранном .exe)."""
    import sys
    meipass = getattr(sys, "_MEIPASS", None)
    candidates = []
    if meipass:
        candidates.append(Path(meipass) / "assets" / name)
    candidates.append(Path(__file__).resolve().parents[1] / "assets" / name)
    for c in candidates:
        if c.exists():
            return c
    return candidates[-1]
from .tabs.tab_connection import ConnectionTab
from .tabs.tab_batch import BatchTab
from .tabs.tab_single import SingleTab
from .tabs.tab_owners import OwnersTab
from .tabs.tab_maintenance import MaintenanceTab
from .tabs.tab_logs import LogsTab
from .tabs.tab_results import ResultsTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {__version__}")
        self.resize(1040, 720)

        icon_path = _asset("icon.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.config = cfgio.load_config()
        self.settings = cfgio.load_settings()
        self._running = False

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.connection_tab = ConnectionTab(self)
        self.batch_tab = BatchTab(self)
        self.single_tab = SingleTab(self)
        self.maintenance_tab = MaintenanceTab(self)
        self.owners_tab = OwnersTab(self)
        self.logs_tab = LogsTab(self)
        self.results_tab = ResultsTab(self)

        self.tabs.addTab(self.connection_tab, "Подключение")
        self.tabs.addTab(self.batch_tab, "Пакетная обработка")
        self.tabs.addTab(self.single_tab, "Одиночный объект")
        self.tabs.addTab(self.owners_tab, "Собственники (Росреестр)")
        self.tabs.addTab(self.maintenance_tab, "Обслуживание")
        self.tabs.addTab(self.logs_tab, "Анализ логов")
        self.tabs.addTab(self.results_tab, "Результаты")

        self.statusBar().showMessage("Готово")

        # мелкая подпись в правом нижнем углу
        credit = QLabel("by BeRealBear")
        credit.setStyleSheet("color: #9e9e9e; font-size: 10px; padding: 0 8px;")
        self.statusBar().addPermanentWidget(credit)

    # -- общий доступ к параметрам -------------------------------------- #
    def current_config(self) -> dict:
        """config + сетевые параметры в формате, понятном движку."""
        c = dict(self.config)
        return c

    def set_running(self, value: bool) -> bool:
        """Глобальный флаг: одновременно допускается один прогон."""
        if value and self._running:
            QMessageBox.warning(self, APP_NAME,
                                "Уже выполняется обработка. Дождитесь завершения "
                                "или остановите её.")
            return False
        self._running = value
        return True

    def is_running(self) -> bool:
        return self._running

    def save_all(self):
        cfgio.save_config(self.config)
        cfgio.save_settings(self.settings)

    def closeEvent(self, e):
        self.save_all()
        super().closeEvent(e)
