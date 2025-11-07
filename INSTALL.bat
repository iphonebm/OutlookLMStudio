@echo off
echo ========================================
echo Installation OutlookLMStudio
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

echo [1/6] Fermeture d'Outlook...
taskkill /IM outlook.exe /F >nul 2>&1
timeout /t 2 >nul

echo [2/6] Desinstallation de l'ancienne version (si presente)...

:: Supprimer les clés de registre existantes
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1
reg delete "HKLM\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1

:: Nettoyer le cache ClickOnce
set CLICKONCE_CACHE=%LOCALAPPDATA%\Apps\2.0
if exist "%CLICKONCE_CACHE%" (
    echo Nettoyage du cache ClickOnce...
    rd /s /q "%CLICKONCE_CACHE%" >nul 2>&1
)

timeout /t 2 >nul

echo [3/6] Verification du fichier d'installation...
set VSTO_PATH=%~dp0bin\Debug\OutlookLMStudio.vsto
echo Chemin recherche: %VSTO_PATH%

if not exist "%VSTO_PATH%" (
    echo.
    echo ERREUR: Fichier OutlookLMStudio.vsto introuvable !
    echo.
    echo Le fichier devrait etre ici: %VSTO_PATH%
    echo.
    echo SOLUTION:
    echo 1. Ouvrez Visual Studio
    echo 2. Build ^> Clean Solution
    echo 3. Build ^> Rebuild Solution
    echo 4. Attendez la fin de la compilation
    echo 5. Relancez ce script
    echo.
    dir /b "%~dp0bin\Debug\*.vsto" 2>nul
    if %errorLevel% neq 0 (
        echo Aucun fichier .vsto trouve dans bin\Debug\
    ) else (
        echo Fichiers .vsto trouves dans bin\Debug\:
        dir /b "%~dp0bin\Debug\*.vsto"
    )
    echo.
    pause
    exit /b 1
)

echo Fichier trouve: %VSTO_PATH%
echo Taille: 
dir "%VSTO_PATH%" | find "OutlookLMStudio.vsto"
echo.

echo [4/6] Installation du complement VSTO...
echo.
echo Une fenetre d'installation va s'ouvrir...
echo Cliquez sur le bouton "Installer" pour continuer.
echo.
echo Chemin: %VSTO_PATH%
echo.

:: Lancer le fichier VSTO sans attendre (car start /wait peut bloquer)
start "" "%VSTO_PATH%"

echo.
echo Patientez pendant l'installation (environ 10 secondes)...
timeout /t 10 >nul

echo [5/6] Verification de l'installation...
timeout /t 2 >nul

:: Vérifier si le complément est enregistré
reg query "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" >nul 2>&1
if %errorLevel% equ 0 (
    echo ? Complement enregistre avec succes dans le registre (HKCU) !
) else (
    reg query "HKLM\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" >nul 2>&1
    if %errorLevel% equ 0 (
        echo ? Complement enregistre avec succes dans le registre (HKLM) !
    ) else (
        echo.
        echo ATTENTION: Le complement ne semble pas encore enregistre.
        echo.
        echo Cela peut signifier:
        echo 1. L'installation est encore en cours (patientez 30 secondes)
        echo 2. Vous avez annule l'installation
        echo 3. Il y a eu une erreur d'installation
        echo.
        echo Verifiez manuellement dans Outlook ^> Fichier ^> Options ^> Complements
        echo.
    )
)

echo [6/6] Lancement d'Outlook...
start outlook.exe

echo.
echo ========================================
echo Installation terminee !
echo ========================================
echo.
echo Le complement "LMStudio Assistant" devrait apparaitre dans Outlook
echo.
echo Si le volet n'apparait pas :
echo 1. Fichier ^> Options ^> Complements
echo 2. En bas : Gerer: Complements COM ^> Atteindre
echo 3. Cochez OutlookLMStudio
echo.
echo Si le complement apparait dans "Complements desactives" :
echo 1. Gerer: Complements desactives ^> Atteindre
echo 2. Selectionnez OutlookLMStudio
echo 3. Cliquez "Toujours activer ce complement"
echo 4. Redemarrez Outlook
echo.
echo Pour desinstaller : Executez UNINSTALL.bat
echo Pour les logs : %APPDATA%\OutlookLMStudio\logs.txt
echo.
pause