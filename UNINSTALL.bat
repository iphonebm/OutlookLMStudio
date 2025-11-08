@echo off
echo ========================================
echo Desinstallation OutlookLMStudio
echo ========================================
echo.

:: Vérifier les droits administrateur
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERREUR: Ce script doit etre execute en tant qu'administrateur
    echo Clic droit sur le fichier ^> Executer en tant qu'administrateur
    echo.
    pause
    exit /b 1
)

echo [1/3] Fermeture d'Outlook...
taskkill /IM outlook.exe /F >nul 2>&1
timeout /t 2 >nul

echo [2/3] Suppression des cles de registre...

:: Supprimer les clés de registre pour l'utilisateur actuel
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1

:: Supprimer les clés de registre machine (si elles existent)
reg delete "HKLM\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1

echo [3/3] Nettoyage du cache ClickOnce...

:: Nettoyer le cache ClickOnce
set CLICKONCE_CACHE=%LOCALAPPDATA%\Apps\2.0
if exist "%CLICKONCE_CACHE%" (
    echo Suppression du cache ClickOnce...
    rd /s /q "%CLICKONCE_CACHE%" >nul 2>&1
)

echo.
echo ========================================
echo Desinstallation terminee !
echo ========================================
echo.
echo Le complement OutlookLMStudio a ete desinstalle.
echo.
echo Pour reinstaller :
echo 1. Executez INSTALL.bat en tant qu'administrateur
echo.
pause