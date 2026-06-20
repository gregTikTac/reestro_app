@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Парсер ЕГРН — форма собственности
echo.
echo ============================================================
echo   ШАГ 5: Форма собственности (Росреестр)
echo ============================================================
echo.
echo Откроется браузер. Решите капчу, откройте карточку объекта.
echo Когда видна строка "Форма собственности" — нажмите Enter в этом окне.
echo.
if not exist "venv\Scripts\activate.bat" (
    echo [ОШИБКА] Сначала запустите 1_УСТАНОВКА.bat
    pause
    exit /b 1
)
call venv\Scripts\activate.bat
python fetch_rosreestr.py -o output --all
echo.
echo Обновление отчёта из кэша...
python rebuild_from_cache.py -o output
echo.
echo Готово.
pause
