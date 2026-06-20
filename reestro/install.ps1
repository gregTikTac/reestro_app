# Парсер ЕГРН — установка зависимостей (PowerShell)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "============================================================"
Write-Host " Парсер ЕГРН — установка зависимостей"
Write-Host "============================================================`n"

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "[ОШИБКА] Python не найден в PATH." -ForegroundColor Red
    Write-Host "Скачайте Python 3.11+ с https://python.org/downloads"
    Write-Host 'При установке включите "Add Python to PATH".'
    exit 1
}

python --version

if (-not (Test-Path "venv")) {
    Write-Host "`nСоздание виртуального окружения venv..."
    python -m venv venv
}

Write-Host "Активация venv и установка пакетов..."
& ".\venv\Scripts\Activate.ps1"
python -m pip install --upgrade pip
pip install -r requirements.txt

Write-Host "`nПроверка импортов..."
python -c "import requests, openpyxl, fpdf; print('OK: requests, openpyxl, fpdf2')"

if (-not (Test-Path "config.json") -and (Test-Path "config.json.example")) {
    Copy-Item "config.json.example" "config.json"
    Write-Host "`nСоздан config.json из шаблона. Заполните apiKey и orgId!" -ForegroundColor Yellow
}

Write-Host @"

============================================================
 Базовая установка завершена.

 Далее:
   1. Отредактируйте config.json (apiKey, orgId)
   2. Положите TZ\Запрос.xlsx во входную папку
   3. python check_balance.py
   4. python run_one_new.py

 Для fetch_rosreestr.py дополнительно:
   python -m playwright install chromium

 Подробная инструкция: ДОКУМЕНТАЦИЯ.md
============================================================
"@
