@echo off
chcp 65001 >nul
setlocal

echo ============================================================
echo  Парсер ЕГРН — установка зависимостей
echo ============================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден в PATH.
    echo Скачайте Python 3.11+ с https://python.org/downloads
    echo При установке включите "Add Python to PATH".
    pause
    exit /b 1
)

python --version
echo.

if not exist venv (
    echo Создание виртуального окружения venv...
    python -m venv venv
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось создать venv.
        pause
        exit /b 1
    )
)

echo Активация venv и установка пакетов...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo [ОШИБКА] pip install завершился с ошибкой.
    pause
    exit /b 1
)

echo.
echo Проверка импортов...
python -c "import requests, openpyxl, fpdf; print('OK: requests, openpyxl, fpdf2')"
if errorlevel 1 (
    echo [ОШИБКА] Проверка импортов не прошла.
    pause
    exit /b 1
)

if not exist config.json (
    if exist config.json.example (
        copy /Y config.json.example config.json >nul
        echo.
        echo Создан config.json из шаблона. Заполните apiKey и orgId!
    )
)

echo.
echo ============================================================
echo  Базовая установка завершена.
echo.
echo  Далее:
echo    1. Отредактируйте config.json (apiKey, orgId)
echo    2. Положите TZ\Запрос.xlsx во входную папку
echo    3. python check_balance.py
echo    4. python run_one_new.py
echo.
echo  Для fetch_rosreestr.py дополнительно:
echo    python -m playwright install chromium
echo.
echo  Подробная инструкция: ДОКУМЕНТАЦИЯ.md
echo ============================================================
pause
