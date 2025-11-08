@echo off
setlocal enableextensions enabledelayedexpansion

:: ==========================================================
:: INSTALL.bat: Désinstallation complète puis réinstallation automatique
:: Ajout DEBUG pour diagnostiquer fermeture immédiate
:: ==========================================================

:: [0] Vérifier droits admin (fltmc plus fiable que 'net session')
fltmc >nul 2>&1
if errorlevel 1 (
    echo ERREUR: Script non exécuté en mode administrateur.
    echo Faites: Clic droit > "Exécuter en tant qu'administrateur".
    pause
    exit /b 1
)

echo DEBUG: Elevation OK
pause

:: [1] Fermer Outlook
echo [1/7] Fermeture d'Outlook...
taskkill /IM outlook.exe /F >nul 2>&1
timeout /t 2 >nul
echo DEBUG: Outlook fermé (ou pas lancé)

:: [2] Désinstallation via Programmes et fonctionnalités (registry) + fallback WMIC
echo [2/7] Désinstallation de l'ancienne version...
echo DEBUG: Scan des entrées de désinstallation (Applications installées)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$paths=@('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*');" ^
  "$apps=Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue ^| Where-Object { ($_.DisplayName -like '*OutlookLMStudio*') -or ($_.DisplayName -like '*OutookLMStudio*') -or ($_.DisplayName -like '*LMStudio*') };" ^
  "if($apps){ foreach($a in $apps){ if($a.DisplayName){ Write-Host ('   Trouvé: ' + $a.DisplayName) }; if($a.UninstallString){ Write-Host (' - Désinstallation: ' + $a.DisplayName); $cmd=$a.UninstallString.Trim(); if($cmd -match 'msiexec(\.exe)?'){ Start-Process cmd -ArgumentList '/c', ($cmd + ' /passive') -Wait } else { Start-Process cmd -ArgumentList '/c', $cmd -Wait } } } } else { Write-Host '   Aucune entrée trouvée (registry).' }" 2>nul

:: Tentative (optionnelle) de suppression Appx si jamais installé comme application UWP (peu probable)
echo DEBUG: Vérification Appx (UWP) pour 'OutookLMStudio'...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$pkgs = Get-AppxPackage -AllUsers ^| Where-Object { $_.Name -like '*OutookLMStudio*' -or $_.PackageFamilyName -like '*OutookLMStudio*' }; if($pkgs){ foreach($p in $pkgs){ Write-Host (' - Remove-AppxPackage: ' + $p.Name); try { Remove-AppxPackage -AllUsers -Package $p.PackageFullName -ErrorAction Stop } catch { Write-Host ('   [Info] Echec Remove-AppxPackage: ' + $_.Exception.Message) } } } else { Write-Host '   Aucun package Appx correspondant.' }" 2>nul

echo DEBUG: Fallback WMIC (peut échouer silencieusement sur versions récentes)
wmic product where "Name like '%%OutlookLMStudio%%' or Name like '%%OutookLMStudio%%'" call uninstall /nointeractive >nul 2>&1

:: [3] Nettoyage des clés de registre Add-in et VSTO
echo [3/7] Nettoyage du registre (Outlook/VSTO)...
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1
reg delete "HKLM\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1
echo DEBUG: Registre nettoyé

:: [4] Nettoyage du cache ClickOnce
echo [4/7] Nettoyage du cache ClickOnce...
set "CLICKONCE_CACHE=%LOCALAPPDATA%\Apps\2.0"
if exist "%CLICKONCE_CACHE%" (
    rd /s /q "%CLICKONCE_CACHE%" >nul 2>&1
    echo DEBUG: Cache ClickOnce supprimé
) else (
    echo DEBUG: Cache ClickOnce absent
)

:: [5] Détection du fichier VSTO
echo [5/7] Détection du fichier d'installation (.vsto)...
set "BASE=%~dp0"
if "%BASE:~-1%"=="\" set "BASE=%BASE:~0,-1%"
set "VSTO_PATH="
echo DEBUG: BASE=%BASE%
if not defined VSTO_PATH if exist "%BASE%\OutlookLMStudio.vsto" set "VSTO_PATH=%BASE%\OutlookLMStudio.vsto"
if not defined VSTO_PATH if exist "%BASE%\bin\Debug\OutlookLMStudio.vsto" set "VSTO_PATH=%BASE%\bin\Debug\OutlookLMStudio.vsto"
if not defined VSTO_PATH if exist "%BASE%\bin\Release\OutlookLMStudio.vsto" set "VSTO_PATH=%BASE%\bin\Release\OutlookLMStudio.vsto"
if not defined VSTO_PATH if exist "%BASE%\..\OutlookLMStudio.vsto" set "VSTO_PATH=%BASE%\..\OutlookLMStudio.vsto"
if not defined VSTO_PATH (
  for /r "%BASE%" %%f in (*.vsto) do ( set "VSTO_PATH=%%f" & goto :foundvsto )
)
if not defined VSTO_PATH (
  for /r "%BASE%\.." %%f in (*.vsto) do ( set "VSTO_PATH=%%f" & goto :foundvsto )
)
:foundvsto
echo DEBUG: VSTO_PATH=%VSTO_PATH%
if not defined VSTO_PATH (
    echo ERREUR: Fichier .vsto introuvable.
    dir /b "%BASE%\*.vsto" 2>nul
    dir /b "%BASE%\..\*.vsto" 2>nul
    echo.
    pause
    exit /b 1
)
echo Fichier VSTO: "%VSTO_PATH%"

:: [6] Installation VSTO
echo [6/7] Lancement de l'installation...
start "" "%VSTO_PATH%"
echo Patientez (15 s)...
timeout /t 15 >nul

echo Vérification enregistrement...
reg query "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" >nul 2>&1 && echo [OK] Addin HKCU || (
  reg query "HKLM\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" >nul 2>&1 && echo [OK] Addin HKLM || echo ATTENTION: Addin non trouvé dans le registre.
)

:: [7] Démarrer Outlook
echo [7/7] Démarrage d'Outlook...
start outlook.exe

echo.
echo ========================================
echo Réinstallation terminée.
echo ========================================

echo DEBUG: Script terminé - fenêtre maintenue ouverte.
pause
endlocal