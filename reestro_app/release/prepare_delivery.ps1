# -*- coding: utf-8 -*-
# Сборка ZIP для передачи заказчику
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$srcDist = Join-Path $root "dist\ЕГРН-Парсер"
$releaseName = "ЕГРН-Парсер_0.1.0"
$releaseDir = Join-Path $PSScriptRoot $releaseName
$zipPath = Join-Path $PSScriptRoot "$releaseName.zip"

if (-not (Test-Path (Join-Path $srcDist "ЕГРН-Парсер.exe"))) {
    throw "Не найден dist\ЕГРН-Парсер\ЕГРН-Парсер.exe — сначала build_exe.bat"
}

if (Test-Path $releaseDir) { Remove-Item $releaseDir -Recurse -Force }
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

New-Item -ItemType Directory -Force -Path "$releaseDir\docs\screenshots" | Out-Null

Write-Host "Копирование приложения..."
Copy-Item "$srcDist\*" $releaseDir -Recurse -Force

Write-Host "Инструкции и скриншоты..."
Copy-Item (Join-Path $root "ИНСТРУКЦИЯ.md") $releaseDir -Force
Copy-Item (Join-Path $root "docs\ПОШАГОВАЯ_ИНСТРУКЦИЯ.md") "$releaseDir\docs\" -Force
$shots = Join-Path $root "docs\screenshots"
Get-ChildItem $shots -Filter "*.png" | ForEach-Object {
    $dest = $_.Name -replace "odin.*\.png", "03_odinochnyy.png"
    Copy-Item $_.FullName (Join-Path "$releaseDir\docs\screenshots" $dest) -Force
}

# ASCII-имя для скрина 3 (если в md кириллица)
$mdPath = Join-Path $releaseDir "docs\ПОШАГОВАЯ_ИНСТРУКЦИЯ.md"
(Get-Content $mdPath -Raw -Encoding UTF8) `
    -replace "03_odinочный\.png", "03_odinochnyy.png" |
    Set-Content $mdPath -Encoding UTF8 -NoNewline

Copy-Item (Join-Path $root "..\reestro\config.json.example") (Join-Path $releaseDir "config.json.example") -Force

# Убрать секреты из комплекта
$keyFile = Join-Path $releaseDir "_internal\reestro\TZ\ключ API.txt"
if (Test-Path $keyFile) { Remove-Item $keyFile -Force }
$engineCfg = Join-Path $releaseDir "_internal\reestro\config.json"
@'
{
  "baseUrl": "https://api.kontur.ru",
  "apiKey": "",
  "orgId": ""
}
'@ | Set-Content $engineCfg -Encoding UTF8

@'
ЕГРН-Парсер 0.1.0 — комплект для заказчика
==========================================

С чего начать
  1. Распакуйте архив в любую папку, напр. C:\Реестро\
  2. Запускайте ЕГРН-Парсер.exe из этой папки (нужна папка _internal рядом!)
  3. Прочитайте ИНСТРУКЦИЯ.md
  4. Пошаговый разбор экрана: docs\ПОШАГОВАЯ_ИНСТРУКЦИЯ.md (со скриншотами)

Настройка API
  - На вкладке «Подключение» введите apiKey и orgId из Контур.Реестро
  - Нажмите «Проверить подключение» → «Сохранить»
  - Образец полей: config.json.example (ключи вводятся в программе, не в файле сборки)

Состав папки
  ЕГРН-Парсер.exe      — программа
  _internal\           — библиотеки (не удалять!)
  ИНСТРУКЦИЯ.md        — полная инструкция
  docs\                — пошаговая инструкция + screenshots\
  config.json.example  — пример полей API

Важно
  - Не переносите один .exe без _internal
  - Папка результатов — корень output, не подпапка pdf\
  - Python не требуется

by BeRealBear
'@ | Set-Content (Join-Path $releaseDir "ПРОЧТИ_МЕНЯ.txt") -Encoding UTF8

Write-Host "Архивирование..."
Compress-Archive -Path $releaseDir -DestinationPath $zipPath -CompressionLevel Optimal -Force

$sizeMb = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host "Готово: $zipPath ($sizeMb MB)"
