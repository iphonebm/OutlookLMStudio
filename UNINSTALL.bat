@echo off
setlocal enableextensions enabledelayedexpansion

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

echo [1/5] Fermeture d'Outlook...
taskkill /IM outlook.exe /F >nul 2>&1
timeout /t 2 >nul

echo [2/5] Desinstallation via "Programmes et fonctionnalites"...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$paths=@('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*');" ^
  "$apps=Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*OutlookLMStudio*' -or $_.DisplayName -like '*LMStudio*' };" ^
  "if($apps){ foreach($a in $apps){ if($a.UninstallString){ Write-Host (' - ' + $a.DisplayName); $cmd=$a.UninstallString.Trim(); if($cmd -match 'msiexec(\.exe)?'){ Start-Process cmd -ArgumentList '/c', ($cmd + ' /passive') -Wait } else { Start-Process cmd -ArgumentList '/c', $cmd -Wait } } } } else { Write-Host 'Aucun element a desinstaller via la base de registre.' }" 2>nul

:: Fallback WMIC pour anciennes installations MSI (optionnel)
wmic product where "Name like '%%OutlookLMStudio%%'" call uninstall /nointeractive >nul 2>&1

echo [3/5] Suppression des cles de registre Outlook/VSTO...
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1
reg delete "HKLM\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\VSTO\Security\Inclusion\{75384258-9a61-432d-b12a-d48c8e01ce3a}" /f >nul 2>&1

echo [4/5] Nettoyage du cache ClickOnce...
set "CLICKONCE_CACHE=%LOCALAPPDATA%\Apps\2.0"
if exist "%CLICKONCE_CACHE%" rd /s /q "%CLICKONCE_CACHE%" >nul 2>&1

echo [5/5] Termine.
echo L'ancienne version devrait maintenant etre retiree de "Programmes et fonctionnalites".
echo Vous pouvez relancer INSTALL.bat pour reinstaller la derniere version.
echo.
pause
endlocal