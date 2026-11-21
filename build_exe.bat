@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ========================================================
echo   Stock Think Power - Creation de l'executable (.exe)
echo ========================================================
echo.

:: 1. Detection de l'interpreteur Python
set "PYTHON_EXE="

where py >nul 2>&1
if not errorlevel 1 (
    py -3.15 -c "import sys" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=py -3.15"
    ) else (
        py -3 -c "import sys" >nul 2>&1
        if not errorlevel 1 (
            set "PYTHON_EXE=py -3"
        ) else (
            set "PYTHON_EXE=py"
        )
    )
) else (
    where python >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=python"
    ) else (
        echo ERREUR : Python est introuvable sur cette machine.
        echo Veuillez installer Python ou l'ajouter au PATH.
        echo.
        pause
        exit /b 1
    )
)

echo Interpreteur Python detecte : %PYTHON_EXE%
%PYTHON_EXE% --version
echo.

:: 2. Verification des dependances
echo [1/3] Verification des dependances...
%PYTHON_EXE% -c "import customtkinter, openpyxl, fpdf, PyInstaller, darkdetect, PIL" >nul 2>&1
if errorlevel 1 (
    echo Des modules sont manquants, installation en cours...
    %PYTHON_EXE% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERREUR : Impossible d'installer les dependances pip.
        echo Verifiez votre connexion internet ou vos droits.
        echo.
        pause
        exit /b 1
    )
) else (
    echo Toutes les dependances Python sont deja installees.
)

:: 3. Generation de l'icone si absente
if not exist "app_icon.ico" (
    echo Generation de l'icone de l'application...
    %PYTHON_EXE% -c "from PIL import Image, ImageDraw; img = Image.new('RGBA', (256, 256), (0,0,0,0)); draw = ImageDraw.Draw(img); draw.rounded_rectangle([(12, 12), (244, 244)], radius=48, fill=(17, 24, 39, 255), outline=(59, 130, 246, 255), width=6); draw.rounded_rectangle([(36, 36), (220, 220)], radius=36, fill=(30, 41, 59, 255)); draw.polygon([(140, 48), (85, 135), (130, 135), (116, 208), (175, 120), (130, 120), (155, 48)], fill=(6, 182, 212, 255), outline=(59, 130, 246, 255)); img.save('app_icon.ico', format='ICO', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])"
)

:: 4. Compilation avec PyInstaller
echo.
echo [2/3] Compilation avec PyInstaller (mode portable tout-en-un)...
%PYTHON_EXE% -m PyInstaller --clean --noconfirm Stock_Think_Power.spec
if errorlevel 1 (
    echo.
    echo ERREUR : La compilation PyInstaller a echoue.
    echo Consultez les messages ci-dessus pour plus de details.
    echo.
    pause
    exit /b 1
)

:: 5. Verification du resultat
if not exist "dist\Stock_Think_Power.exe" (
    echo.
    echo ERREUR : Le fichier dist\Stock_Think_Power.exe n'a pas ete genere.
    pause
    exit /b 1
)

echo.
echo ========================================================
echo [3/3] TERMINE AVEC SUCCES !
echo ========================================================
echo Executable genere : dist\Stock_Think_Power.exe
echo.
dir "dist\Stock_Think_Power.exe" | findstr /i "Stock_Think_Power.exe"
echo.
echo --------------------------------------------------------
echo DEPLOIEMENT SUR UN AUTRE POSTE :
echo 1. Copiez le fichier dist\Stock_Think_Power.exe sur n'importe
echo    quel autre PC Windows (10 ou 11, 64-bit).
echo 2. Aucune installation de Python n'est requise sur l'autre PC.
echo 3. L'executable contient deja :
echo    - L'interpreteur et tous les modules requis (CustomTkinter, OpenPyXL, FPDF2)
echo    - Les DLLs Runtime C++ (MSVCP140, VCRUNTIME140)
echo    - Les themes sombres et polices vectorielles
echo    - L'icone integree
echo 4. Si vous souhaitez transferer vos donnees existantes, copiez
echo    aussi le fichier Stock_Think_Power_Base.xlsx a cote du .exe.
echo    Sinon, une base propre sera automatiquement creee au 1er lancement.
echo ========================================================
pause
