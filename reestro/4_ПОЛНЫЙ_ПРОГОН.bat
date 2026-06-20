@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Парсер ЕГРН — полный прогон
echo.
echo ============================================================
echo   ШАГ 4: Полное сканирование всех объектов
echo ============================================================
echo.
echo Будут обработаны все НОВЫЕ кадастровые номера из TZ\Запрос.xlsx
echo Уже готовые объекты пропускаются автоматически.
echo.
echo Нажмите любую клавишу для старта или закройте окно для отмены.
pause >nul
echo.
if not exist "venv\Scripts\activate.bat" (
    echo [ОШИБКА] Сначала запустите 1_УСТАНОВКА.bat
    pause
    exit /b 1
)
if not exist "TZ\Запрос.xlsx" (
    echo [ОШИБКА] Нет файла TZ\Запрос.xlsx
    pause
    exit /b 1
)
call venv\Scripts\activate.bat
python reestro_parser.py
echo.
echo ============================================================
echo   Готово. Результаты: output\report.xlsx и output\pdf\
echo ============================================================
echo.
pause
