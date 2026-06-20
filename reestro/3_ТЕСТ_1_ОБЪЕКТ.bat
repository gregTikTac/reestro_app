@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Парсер ЕГРН — тест 1 объект
echo.
echo ============================================================
echo   ШАГ 3: Тест — один объект
echo ============================================================
echo.
if not exist "venv\Scripts\activate.bat" (
    echo [ОШИБКА] Сначала запустите 1_УСТАНОВКА.bat
    pause
    exit /b 1
)
if not exist "TZ\Запрос.xlsx" (
    echo [ВНИМАНИЕ] Нет файла TZ\Запрос.xlsx
    echo Положите файл Запрос.xlsx в папку TZ\
    echo.
    pause
    exit /b 1
)
call venv\Scripts\activate.bat
python run_one_new.py
echo.
echo Проверьте папку output\ — report.xlsx и pdf\
echo.
pause
