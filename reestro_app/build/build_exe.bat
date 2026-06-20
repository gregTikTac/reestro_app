@echo off
chcp 65001 >nul
REM Сборка ЕГРН-Парсер — ГОТОВАЯ папка dist\ЕГРН-Парсер\
REM Запускать из reestro_app\ (двойной клик или из cmd)

cd /d "%~dp0\.."

if not exist venv (
    echo Создаю виртуальное окружение...
    python -m venv venv
)
call venv\Scripts\activate

echo Устанавливаю зависимости...
pip install -r requirements.txt -q
pip install pyinstaller -q

echo.
echo Собираю .exe (это 1-2 минуты)...
pyinstaller build\egrn.spec --noconfirm --clean --distpath dist --workpath build\_work
if errorlevel 1 (
    echo ОШИБКА сборки. Смотрите текст выше.
    pause
    exit /b 1
)

echo.
echo Проверка сборки...
python build\verify_dist.py
if errorlevel 1 (
    echo.
    echo Сборка неполная. НЕ запускайте exe из build\egrn\
    pause
    exit /b 1
)

echo.
echo ========================================
echo  ГОТОВО. Запускайте программу отсюда:
echo  dist\ЕГРН-Парсер\ЕГРН-Парсер.exe
echo.
echo  НЕ запускайте exe из build\egrn\ — там нет библиотек!
echo ========================================
echo.
echo Для установщика:
echo   1) python build\verify_dist.py  — должно быть OK
echo   2) Inno Setup: build\installer.iss -^> Compile
echo   3) Установщик: dist_installer\ЕГРН-Парсер_Setup_0.1.0.exe
echo      (НЕ из build\dist_installer — старая папка, если была)
pause
