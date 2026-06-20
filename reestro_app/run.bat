@echo off
chcp 65001 >nul
REM Запуск приложения из исходников (для разработки/проверки).
cd /d "%~dp0"
if not exist venv (
    echo Создаю окружение и ставлю зависимости...
    python -m venv venv
    call venv\Scripts\activate
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate
)
python app_entry.py
