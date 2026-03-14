# ╔══════════════════════════════════════════════════════════╗
# ║  BUILD GEO → Excel Converter (.exe)                     ║
# ║  Script PowerShell pour créer un exécutable standalone   ║
# ╚══════════════════════════════════════════════════════════╝

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  GEO to Excel Converter - Build Windows .exe" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Se placer à la racine du projet
Set-Location "$PSScriptRoot\.."

# Vérifier Python
try {
    $pyVersion = python --version 2>&1
    Write-Host "[OK] $pyVersion" -ForegroundColor Green
} catch {
    Write-Host "[ERREUR] Python n'est pas installe." -ForegroundColor Red
    Write-Host "Installez Python 3.10+ depuis https://python.org" -ForegroundColor Yellow
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}

# Vérifier que le script principal existe
if (-not (Test-Path "src\geo_to_excel\app.py")) {
    Write-Host "[ERREUR] src\geo_to_excel\app.py introuvable." -ForegroundColor Red
    Write-Host "Lancez ce script depuis la racine du projet." -ForegroundColor Yellow
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}

# Étape 1 : Installer les dépendances
Write-Host ""
Write-Host "[1/3] Installation des dependances..." -ForegroundColor Yellow
pip install openpyxl ezdxf pyproj pyinstaller --quiet --upgrade
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERREUR] Echec de l'installation." -ForegroundColor Red
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}
Write-Host "[OK] Dependances installees." -ForegroundColor Green

# Étape 2 : Trouver le dossier pyproj data
Write-Host ""
Write-Host "[2/3] Creation de l'executable (cela peut prendre 1-2 min)..." -ForegroundColor Yellow

$pyprojDataDir = python -c "import pyproj; import os; print(os.path.join(os.path.dirname(pyproj.__file__)))" 2>$null

$pyinstallerArgs = @(
    "--onefile",
    "--windowed",
    "--name", "GEO_to_Excel",
    "--hidden-import=openpyxl",
    "--hidden-import=ezdxf",
    "--hidden-import=pyproj",
    "--hidden-import=pyproj.database",
    "--hidden-import=pyproj._datadir",
    "--hidden-import=pyproj.enums",
    "--hidden-import=pyproj._crs",
    "--hidden-import=pyproj._transformer",
    "--hidden-import=pyproj._geod",
    "--hidden-import=certifi",
    "--collect-data=pyproj",
    "--collect-data=certifi",
    "--noconfirm",
    "--clean",
    "src\geo_to_excel\app.py"
)

& pyinstaller @pyinstallerArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERREUR] Echec du build PyInstaller." -ForegroundColor Red
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}

# Étape 3 : Nettoyage
Write-Host ""
Write-Host "[3/3] Nettoyage..." -ForegroundColor Yellow
Remove-Item -Recurse -Force "build" -ErrorAction SilentlyContinue
Remove-Item -Force "GEO_to_Excel.spec" -ErrorAction SilentlyContinue

# Taille de l'exe
$exePath = "dist\GEO_to_Excel.exe"
if (Test-Path $exePath) {
    $size = (Get-Item $exePath).Length / 1MB
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  BUILD TERMINE !" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Executable : dist\GEO_to_Excel.exe" -ForegroundColor White
    Write-Host "  Taille     : $([math]::Round($size, 1)) Mo" -ForegroundColor White
    Write-Host ""
    Write-Host "  Copiez GEO_to_Excel.exe ou vous voulez." -ForegroundColor Gray
    Write-Host "  Aucune installation de Python requise." -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Green
} else {
    Write-Host "[ERREUR] L'executable n'a pas ete cree." -ForegroundColor Red
}

Write-Host ""
Read-Host "Appuyez sur Entree pour quitter"
