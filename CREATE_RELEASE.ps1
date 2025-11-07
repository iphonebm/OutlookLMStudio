# CREATE_RELEASE.ps1 - Créer l'archive de release pour GitHub
# Exécutez ce script pour créer le fichier ZIP de release

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Création de la Release OutlookLMStudio" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$version = "v1.0.0"
$projectName = "OutlookLMStudio"
$releaseFolder = "Release"
$zipName = "${projectName}-${version}.zip"

# Vérifier que bin\Debug existe
if (-not (Test-Path "bin\Debug")) {
    Write-Host "? ERREUR: Le dossier bin\Debug n'existe pas!" -ForegroundColor Red
    Write-Host "   Compilez d'abord le projet (Build ? Rebuild Solution)" -ForegroundColor Yellow
    exit 1
}

# Vérifier que la DLL existe
if (-not (Test-Path "bin\Debug\OutlookLMStudio.dll")) {
    Write-Host "? ERREUR: OutlookLMStudio.dll introuvable!" -ForegroundColor Red
    Write-Host "   Compilez d'abord le projet (Build ? Rebuild Solution)" -ForegroundColor Yellow
    exit 1
}

# Vérifier que le fichier .vsto existe
if (-not (Test-Path "bin\Debug\OutlookLMStudio.vsto")) {
    Write-Host "? ERREUR: OutlookLMStudio.vsto introuvable!" -ForegroundColor Red
    Write-Host "   Compilez d'abord le projet (Build ? Rebuild Solution)" -ForegroundColor Yellow
    exit 1
}

Write-Host "? Fichiers de compilation trouvés" -ForegroundColor Green
Write-Host ""

# Créer le dossier Release s'il n'existe pas
if (Test-Path $releaseFolder) {
    Write-Host "???  Suppression de l'ancien dossier Release..." -ForegroundColor Yellow
    Remove-Item -Path $releaseFolder -Recurse -Force
}

Write-Host "?? Création du dossier Release..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $releaseFolder | Out-Null

# Créer la structure dans le dossier Release
Write-Host "?? Copie des fichiers..." -ForegroundColor Cyan

# Copier tout le contenu de bin\Debug
Write-Host "   - Fichiers binaires (bin\Debug\)" -ForegroundColor Gray
Copy-Item -Path "bin\Debug\*" -Destination $releaseFolder -Recurse -Force

# Copier les scripts d'installation
Write-Host "   - Scripts d'installation" -ForegroundColor Gray
Copy-Item -Path "INSTALL.bat" -Destination $releaseFolder -Force
Copy-Item -Path "UNINSTALL.bat" -Destination $releaseFolder -Force
Copy-Item -Path "DIAGNOSTIC.ps1" -Destination $releaseFolder -Force

# Copier la documentation
Write-Host "   - Documentation" -ForegroundColor Gray
Copy-Item -Path "README.md" -Destination $releaseFolder -Force
Copy-Item -Path "LICENSE" -Destination $releaseFolder -Force
Copy-Item -Path "CHANGELOG.md" -Destination $releaseFolder -Force
Copy-Item -Path "RELEASE_NOTES_${version}.md" -Destination "$releaseFolder\RELEASE_NOTES.md" -Force

# Créer un fichier d'instructions rapides
Write-Host "   - Guide d'installation rapide" -ForegroundColor Gray
$quickStart = @"
========================================
  OutlookLMStudio ${version}
========================================

INSTALLATION RAPIDE :

1. Fermez Outlook complètement
2. Clic droit sur INSTALL.bat
3. "Exécuter en tant qu'administrateur"
4. Lancez Outlook
5. Le complément apparaît automatiquement !

CONFIGURATION LMSTUDIO :

1. Lancez LMStudio
2. Chargez un modèle (ex: Mistral 7B)
3. Onglet "Local Server"
4. Cliquez "Start Server"

UTILISATION :

- Sélectionnez un email
- Clic droit ? "Générer Réponse(s) avec LMStudio"
- Vérifiez et envoyez !

SUPPORT :

- Logs : %APPDATA%\OutlookLMStudio\logs.txt
- Diagnostic : Exécutez DIAGNOSTIC.ps1
- Bugs : https://github.com/iphonebm/OutlookLMStudio/issues

Documentation complète : README.md
"@

$quickStart | Out-File -FilePath "$releaseFolder\QUICKSTART.txt" -Encoding UTF8

# Supprimer les fichiers inutiles du dossier Release
Write-Host "?? Nettoyage des fichiers inutiles..." -ForegroundColor Cyan
$filesToRemove = @(
    "*.pdb",
    "*.xml",
    "*.vshost.*",
    "*.application",
    "*.manifest",
    "OutlookLMStudio.dll.config"
)

foreach ($pattern in $filesToRemove) {
    Get-ChildItem -Path $releaseFolder -Filter $pattern -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
}

# Créer l'archive ZIP
Write-Host ""
Write-Host "?? Création de l'archive ZIP..." -ForegroundColor Cyan

if (Test-Path $zipName) {
    Write-Host "   Suppression de l'ancienne archive..." -ForegroundColor Yellow
    Remove-Item -Path $zipName -Force
}

# Utiliser Compress-Archive (PowerShell 5.0+)
Compress-Archive -Path "$releaseFolder\*" -DestinationPath $zipName -CompressionLevel Optimal

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  ? Release créée avec succès !" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Afficher les informations de la release
$zipInfo = Get-Item $zipName
Write-Host "?? Archive : $zipName" -ForegroundColor Cyan
Write-Host "?? Taille  : $([math]::Round($zipInfo.Length / 1MB, 2)) MB" -ForegroundColor Cyan
Write-Host "?? Dossier : $(Get-Location)" -ForegroundColor Cyan
Write-Host ""

# Afficher le contenu
Write-Host "?? Contenu de l'archive :" -ForegroundColor Cyan
Write-Host ""
Write-Host "   OutlookLMStudio.dll          - Complément compilé" -ForegroundColor Gray
Write-Host "   OutlookLMStudio.vsto         - Installateur" -ForegroundColor Gray
Write-Host "   INSTALL.bat                  - Script d'installation" -ForegroundColor Gray
Write-Host "   UNINSTALL.bat                - Script de désinstallation" -ForegroundColor Gray
Write-Host "   DIAGNOSTIC.ps1               - Script de diagnostic" -ForegroundColor Gray
Write-Host "   README.md                    - Documentation complète" -ForegroundColor Gray
Write-Host "   QUICKSTART.txt               - Guide rapide" -ForegroundColor Gray
Write-Host "   RELEASE_NOTES.md             - Notes de version" -ForegroundColor Gray
Write-Host "   CHANGELOG.md                 - Journal des modifications" -ForegroundColor Gray
Write-Host "   LICENSE                      - Licence MIT" -ForegroundColor Gray
Write-Host "   + Dépendances (DLLs)         - Bibliothèques requises" -ForegroundColor Gray
Write-Host ""

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PROCHAINES ÉTAPES" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Allez sur GitHub : https://github.com/iphonebm/OutlookLMStudio" -ForegroundColor Yellow
Write-Host "2. Cliquez sur 'Releases' ? 'Create a new release'" -ForegroundColor Yellow
Write-Host "3. Tag version : ${version}" -ForegroundColor Yellow
Write-Host "4. Titre : OutlookLMStudio ${version} - Première Release" -ForegroundColor Yellow
Write-Host "5. Description : Copiez le contenu de RELEASE_NOTES_${version}.md" -ForegroundColor Yellow
Write-Host "6. Uploadez le fichier : ${zipName}" -ForegroundColor Yellow
Write-Host "7. Cliquez 'Publish release'" -ForegroundColor Yellow
Write-Host ""

Write-Host "? Votre release est prête à être publiée sur GitHub !" -ForegroundColor Green
Write-Host ""

# Proposer d'ouvrir le dossier
$response = Read-Host "Voulez-vous ouvrir le dossier de la release ? (O/N)"
if ($response -eq "O" -or $response -eq "o") {
    Start-Process explorer.exe -ArgumentList (Get-Location)
}