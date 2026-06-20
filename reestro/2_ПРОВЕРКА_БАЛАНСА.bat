@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Парсер ЕГРН — проверка баланса
echo.
echo ============================================================
echo   ШАГ 2: Проверка API (баланс)
echo ============================================================
echo.
if not exist "venv\Scripts\activate.bat" (
    echo [ОШИБКА] Сначала запустите 1_УСТАНОВКА.bat
    echo.
    pause
    exit /b 1
)
if not exist "config.json" (
    echo [ОШИБКА] Нет config.json — заполните apiKey и orgId
    echo.
    pause
    exit /b 1
)
call venv\Scripts\activate.bat
python check_balance.py
echo.
echo Если видите "balance HTTP 200" — всё в порядке.
echo.
pause
