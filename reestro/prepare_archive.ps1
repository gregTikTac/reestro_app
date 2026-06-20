# Build distribution archive (ASCII-only script for Windows encoding)
$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot
$Dist = Join-Path $Root "dist\reestro_parser_full"

$IncludeFiles = @(
    "reestro_parser.py",
    "rebuild_from_cache.py",
    "fetch_rosreestr.py",
    "run_one_new.py",
    "check_balance.py",
    "config.json.example",
    "requirements.txt",
    "install.bat",
    "install.ps1",
    "0_ОТКРЫТЬ_ПАПКУ_ПРОЕКТА.bat",
    "1_УСТАНОВКА.bat",
    "2_ПРОВЕРКА_БАЛАНСА.bat",
    "3_ТЕСТ_1_ОБЪЕКТ.bat",
    "4_ПОЛНЫЙ_ПРОГОН.bat",
    "5_ФОРМА_СОБСТВЕННОСТИ.bat",
    "README.md",
    "kns.csv"
)

if (Test-Path $Dist) {
    Remove-Item $Dist -Recurse -Force
}
New-Item -ItemType Directory -Path $Dist -Force | Out-Null

foreach ($f in $IncludeFiles) {
    $src = Join-Path $Root $f
    if (Test-Path $src) {
        Copy-Item $src (Join-Path $Dist $f)
        Write-Host "  + $f"
    }
}

# Docs with Cyrillic names - copy by pattern
Get-ChildItem $Root -File | Where-Object {
    $_.Extension -eq ".md" -and $_.Name -ne "README.md"
} | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $Dist $_.Name)
    Write-Host "  + $($_.Name)"
}
Get-ChildItem $Root -File -Filter "*.txt" | Where-Object {
    $_.Name -ne "requirements.txt"
} | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $Dist $_.Name)
    Write-Host "  + $($_.Name)"
}

foreach ($d in @("input", "postman", "TZ")) {
    $srcDir = Join-Path $Root $d
    if (-not (Test-Path $srcDir)) { continue }
    $destDir = Join-Path $Dist $d
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    Get-ChildItem $srcDir -File | ForEach-Object {
        if ($d -eq "TZ" -and $_.Name -like "*API*") {
            Write-Host "  - TZ\$($_.Name) (skipped: secrets)" -ForegroundColor DarkYellow
            return
        }
        Copy-Item $_.FullName (Join-Path $destDir $_.Name)
        Write-Host "  + $d\$($_.Name)"
    }
}

$ZipPath = Join-Path (Join-Path $Root "dist") "reestro_parser_full.zip"
if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
Compress-Archive -Path $Dist -DestinationPath $ZipPath -Force

Write-Host ""
Write-Host "Done:"
Write-Host "  Folder: $Dist"
Write-Host "  ZIP:    $ZipPath"
Write-Host ""
Write-Host "Add TZ\Zapros.xlsx manually before sending if needed."
