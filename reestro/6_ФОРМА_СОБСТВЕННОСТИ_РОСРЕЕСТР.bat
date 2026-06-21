@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo  Загрузка ФОРМЫ СОБСТВЕННОСТИ с lk.rosreestr.ru
echo  (вход кадастровым инженером, проверка капчи)
echo ============================================================
echo.
echo  1) Откроется браузер. Войдите через Госуслуги (кад. инженер).
echo  2) Проверьте: появляется ли капча при поиске объекта.
echo  3) Дальше скрипт сам считает форму по списку КН.
echo.

REM Папка результатов (где report.xlsx, cache\). По умолчанию output_test_first.
set "OUT=%~1"
if "%OUT%"=="" set "OUT=output_test_first"

REM Установка Playwright при первом запуске (один раз).
python -c "import playwright" 2>nul
if errorlevel 1 (
    echo Устанавливаю Playwright (один раз)...
    python -m pip install playwright
    python -m playwright install chromium
)

echo.
echo Папка результатов: %OUT%
echo Режим: авто (ручное вмешательство только при капче)
echo.
python fetch_rosreestr.py -o "%OUT%" --all --auto

echo.
echo ============================================================
echo  Готово. Теперь в приложении: «Обслуживание» -> «Пересобрать из кэша»
echo  для папки %OUT% — форма попадёт в report.xlsx и PDF.
echo ============================================================
pause
