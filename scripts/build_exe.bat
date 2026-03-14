@echo off
REM ╔══════════════════════════════════════════════════════════╗
REM ║  BUILD GEO → Excel Converter (.exe)                     ║
REM ║  Crée un exécutable standalone Windows                   ║
REM ╚══════════════════════════════════════════════════════════╝

echo.
echo ============================================================
echo   GEO to Excel Converter - Build Windows .exe
echo ============================================================
echo.

REM Se placer à la racine du projet (un niveau au-dessus de scripts/)
cd /d "%~dp0\.."

REM Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Installez Python 3.10+ depuis https://python.org
    pause
    exit /b 1
)

echo [1/3] Installation des dependances...
pip install openpyxl ezdxf pyproj pyinstaller --quiet
if errorlevel 1 (
    echo [ERREUR] Echec de l'installation des dependances.
    pause
    exit /b 1
)

echo [2/3] Creation de l'executable...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "GEO_to_Excel" ^
    --icon=NONE ^
    --hidden-import=openpyxl ^
    --hidden-import=ezdxf ^
    --hidden-import=pyproj ^
    --hidden-import=pyproj.database ^
    --hidden-import=pyproj._datadir ^
    --hidden-import=pyproj.enums ^
    --hidden-import=pyproj._crs ^
    --hidden-import=pyproj._transformer ^
    --hidden-import=pyproj._geod ^
    --hidden-import=certifi ^
    --collect-data=pyproj ^
    --collect-data=certifi ^
    --noconfirm ^
    --clean ^
    src\geo_to_excel\app.py

if errorlevel 1 (
    echo [ERREUR] Echec de la creation de l'executable.
    echo Verifiez les erreurs ci-dessus.
    pause
    exit /b 1
)

echo.
echo [3/3] Nettoyage...
rmdir /s /q build 2>nul
del GEO_to_Excel.spec 2>nul

echo.
echo ============================================================
echo   BUILD TERMINE !
echo.
echo   Executable: dist\GEO_to_Excel.exe
echo.
echo   Vous pouvez copier GEO_to_Excel.exe ou vous voulez.
echo   Aucune installation de Python requise pour l'utiliser.
echo ============================================================
echo.
pause
