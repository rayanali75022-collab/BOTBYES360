@echo off
title Bot BYES 360 - Lancement
color 0A
echo.
echo  ==========================================
echo   Bot BYES 360 - Prise de Commande Auto
echo  ==========================================
echo.

REM Verifier que Python est installe
py --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python non trouve. Installe Python depuis python.org
    pause
    exit /b 1
)

echo [1/3] Python detecte OK
echo.

REM Installer les dependances
echo [2/3] Installation des dependances...
py -m pip install selenium pdfplumber --quiet
if errorlevel 1 (
    echo [ERREUR] Installation des dependances echouee
    pause
    exit /b 1
)
echo [2/3] Dependances OK
echo.

REM Lancer l'interface
echo [3/3] Lancement de l'interface...
echo.
py "%~dp0interface_bot.py"

if errorlevel 1 (
    echo.
    echo [ERREUR] Le bot a rencontre une erreur.
    pause
)
