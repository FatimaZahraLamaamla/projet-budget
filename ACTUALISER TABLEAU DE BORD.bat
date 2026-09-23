@echo off
chcp 65001 > nul
title Actualisation du Tableau de Bord

echo ============================================
echo   TABLEAU DE BORD - ENGAGEMENTS JUIN 2026
echo   Actualisation des donnees...
echo ============================================
echo.

REM Fermer Excel si ouvert (evite les conflits de fichier)
taskkill /F /IM EXCEL.EXE /T > nul 2>&1
timeout /t 2 /nobreak > nul

REM Lancer le script Python
"C:\Users\HP\AppData\Local\Programs\Python\Python312\python.exe" "%~dp0actualiser_tableau_de_bord.py"

echo.
if %ERRORLEVEL% EQU 0 (
    echo Actualisation terminee avec succes !
    echo Vous pouvez maintenant ouvrir le fichier Tableau de Bord.
) else (
    echo ERREUR lors de l'actualisation. Verifiez les fichiers.
)
echo.
pause
