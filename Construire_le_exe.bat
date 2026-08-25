@echo off
chcp 65001 >nul
echo ================================================
echo   Construction de Mon Livre de Recettes (.exe)
echo ================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou n'est pas dans le PATH.
    echo Installez Python depuis https://www.python.org/downloads/
    echo en cochant bien la case "Add Python to PATH", puis relancez ce script.
    pause
    exit /b 1
)

echo Etape 1/3 : Installation des dependances necessaires...
echo   (pillow, reportlab, openpyxl, qrcode, pytesseract, pyinstaller)
python -m pip install --upgrade pip >nul
python -m pip install pillow reportlab openpyxl qrcode pytesseract pyinstaller
if errorlevel 1 (
    echo [ERREUR] L'installation des dependances a echoue.
    pause
    exit /b 1
)

echo.
echo Etape 2/3 : Construction de l'executable (peut prendre 1 a 2 minutes)...
python -m PyInstaller --onefile --windowed --name "MesRecettes" main.pyw
if errorlevel 1 (
    echo [ERREUR] La construction de l'executable a echoue.
    pause
    exit /b 1
)

echo.
echo Etape 3/3 : Copie des fichiers necessaires a cote de l'executable...
copy /Y ingredients_par_defaut.json dist\ingredients_par_defaut.json >nul
copy /Y valeurs_nutritionnelles.json dist\valeurs_nutritionnelles.json >nul
copy /Y ingredient_allergenes.json dist\ingredient_allergenes.json >nul
copy /Y LISEZ-MOI.txt dist\LISEZ-MOI.txt >nul

echo.
echo ================================================
echo   Termine !
echo   Votre application se trouve dans :
echo   dist\MesRecettes.exe
echo ================================================
echo.
echo Vous pouvez deplacer le dossier "dist" entier ailleurs
echo (cle USB, autre PC, Bureau...) : il contient tout ce qu'il
echo faut pour fonctionner, sans avoir besoin d'installer Python.
echo Gardez simplement tous les fichiers de ce dossier ensemble.
echo.
pause
