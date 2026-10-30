# Bygger en frittstaaende versjon av Kunstlopassistenten for USB-pinne.
# Resultat: dist\Kunstlopassistenten\ (kopier hele mappen til USB og start Kunstlopassistenten.exe).
# Kjor: powershell -ExecutionPolicy Bypass -File build_usb.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$python = ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) { $python = "python" }

& $python -m pip install -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) { throw "pip install feilet" }

# Versjon fra git, siden git ikke finnes paa andre PCer
$version = (git rev-parse --short HEAD 2>$null)
if (-not $version) { $version = Get-Date -Format "yyyy-MM-dd" }
Set-Content -Path "version.txt" -Value $version -Encoding utf8

# Logoen er ikke i git (*.jpg ignoreres), saa den tas bare med hvis den finnes
$logoArgs = @()
if (Test-Path "Lil Logo.jpg") { $logoArgs = @("--add-data", "Lil Logo.jpg;.") }

& $python -m PyInstaller --noconfirm --clean --onedir --windowed `
    --name Kunstlopassistenten `
    @logoArgs `
    --add-data "version.txt;." `
    --hidden-import babel.numbers `
    kunstlopassistenten.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller feilet" }

Remove-Item "version.txt"
Write-Host ""
Write-Host "Ferdig: dist\Kunstlopassistenten\Kunstlopassistenten.exe (versjon $version)"
