# -*- coding: utf-8 -*-
"""Точка входа GUI ЕГРН-Парсера."""
from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from . import APP_NAME, __version__
from .main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(__version__)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
