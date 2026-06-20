# -*- coding: utf-8 -*-
"""
Точка входа для PyInstaller (.exe).

PyInstaller запускает entry-скрипт как __main__, поэтому нельзя использовать
egrn_gui/main.py с относительными импортами (from .main_window …).
Этот файл импортирует пакет egrn_gui абсолютным путём.
"""
from egrn_gui.main import main

if __name__ == "__main__":
    main()
