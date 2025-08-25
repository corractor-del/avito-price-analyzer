# Автоматическая сборка Avito Price Analyzer в .exe (PowerShell)
# Запуск: правый клик по build_windows.ps1 -> Run with PowerShell (или в консоли: powershell -ExecutionPolicy Bypass -File .\build_windows.ps1)

$ErrorActionPreference = "Stop"
$proj = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $proj

Write-Host "[1/4] Создаю виртуальное окружение..."
python -m venv venv

Write-Host "[2/4] Активирую окружение и ставлю зависимости..."
& "$proj\venv\Scripts\Activate.ps1"
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller

Write-Host "[3/4] Сборка через PyInstaller..."
pyinstaller AvitoPriceAnalyzer.spec

Write-Host "[4/4] Готово! Файл: dist\AvitoPriceAnalyzer\AvitoPriceAnalyzer.exe"
