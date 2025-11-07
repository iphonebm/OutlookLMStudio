# ?? DIAGNOSTIC - Sélection d'Email ne Fonctionne Pas

## ? Problème

Quand vous cliquez sur un email dans votre boîte de réception, **LMStudio Assistant ne détecte pas la sélection**.

## ? Modifications Apportées (Logs de Diagnostic)

J'ai ajouté des **logs de diagnostic complets** dans le code pour identifier exactement où le problème se situe.

### Fichiers Modifiés

1. **TaskPaneControl.cs**
   - Logs dans `InitializeEvents()` - vérifie si `Globals.ThisAddIn` est null
   - Logs dans `ThisAddIn_EmailSelected()` - trace la réception de l'événement
   - Logs dans `UpdateSelectedEmailInfo()` - vérifie l'affichage

2. **ThisAddIn.cs**
   - Logs détaillés dans `Explorer_SelectionChange()` - trace chaque étape
   - Logs dans `OnEmailSelected()` - vérifie le nombre d'abonnés à l'événement

## ?? Instructions pour Diagnostiquer

### Étape 1 : Recompiler (Avec des Erreurs de Test)

Le projet a des erreurs liées aux fichiers de **test unitaires** (OutlookLMStudio.Tests). Ces fichiers ne devraient PAS être compilés avec le projet principal.

**Solution temporaire** :
1. Déchargez le projet de test :
   ```
   Solution Explorer ? Clic droit sur "OutlookLMStudio.Tests" ? Décharger le projet
   ```

2. Recompilez :
   ```
   Build ? Rebuild Solution
   ```

### Étape 2 : Installer le Complément

```cmd
# Fermez Outlook complètement
# Puis en tant qu'administrateur :
INSTALL_MANUAL.bat
```

### Étape 3 : Activer les Logs de Debug

1. **Visual Studio** :
   - Ouvrez : `Vue ? Sortie` (ou `Ctrl + Alt + O`)
   - Dans le menu déroulant, sélectionnez "Debug"

2. **Lancez Outlook en mode Debug** :
   - Dans Visual Studio : `Déboguer ? Démarrer le débogage` (F5)
   - OU lancez Outlook normalement

### Étape 4 : Tester la Sélection

1. **Dans Outlook**, cliquez sur un email dans votre boîte de réception
2. **Observez** :
   - La fenêtre "Sortie" de Visual Studio (si en mode debug)
   - LE fichier de logs : `%APPDATA%\OutlookLMStudio\logs.txt`

### Étape 5 : Analyser les Logs

Ouvrez le fichier de logs :
```
%APPDATA%\OutlookLMStudio\logs.txt
```

Cherchez les messages suivants :

#### ? **SI TOUT FONCTIONNE**, vous devriez voir :

```
[timestamp] - Explorer_SelectionChange: Événement déclenché
[timestamp] - Explorer_SelectionChange: Selection.Count = 1
[timestamp] - Explorer_SelectionChange: Type d'élément = MailItemClass
[timestamp] - Explorer_SelectionChange: MailItem trouvé - [Sujet de l'email]
[timestamp] - OnEmailSelected: Début - [Sujet de l'email]
[timestamp] - OnEmailSelected: EmailSelected a 1 abonnés
[timestamp] - OnEmailSelected: Événement déclenché avec succès
[timestamp] - Email sélectionné: [Sujet de l'email]
```

#### ? **SI CELA NE FONCTIONNE PAS**, vous verrez l'un de ces messages :

**Cas 1 : L'événement ne se déclenche jamais**
```
PAS de message "Explorer_SelectionChange: Événement déclenché"
```
? **Problème** : L'événement `SelectionChange` n'est pas enregistré
? **Solution** : Vérifier que `SetupOutlookHandlers()` est bien appelé

**Cas 2 : Explorer est null**
```
Explorer_SelectionChange: Explorer est NULL
```
? **Problème** : Outlook n'a pas d'explorateur actif
? **Solution** : Redémarrer Outlook

**Cas 3 : Aucune sélection**
```
Explorer_SelectionChange: Selection.Count = 0
```
? **Problème** : Outlook ne détecte pas la sélection
? **Solution** : Cliquer directement sur un email (pas juste passer la souris)

**Cas 4 : Ce n'est pas un MailItem**
```
Explorer_SelectionChange: L'élément n'est PAS un MailItem
```
? **Problème** : Vous avez cliqué sur un rendez-vous, contact, etc.
? **Solution** : Cliquer sur un EMAIL

**Cas 5 : Aucun abonné à l'événement**
```
OnEmailSelected: EmailSelected a 0 abonnés
```
? **Problème** : `TaskPaneControl` ne s'est pas enregistré à l'événement
? **Solution** : Vérifier l'erreur au démarrage (message "ERREUR: Globals.ThisAddIn est null")

**Cas 6 : Globals.ThisAddIn est null**
```
TaskPaneControl: ERREUR - Globals.ThisAddIn est NULL!
```
? **Problème** : L'initialisation du complément a échoué
? **Solution** : Vérifier les logs de démarrage

## ?? Scénarios Courants

### Scénario A : "Globals.ThisAddIn est null"

**Cause** : Le `TaskPaneControl` est créé AVANT que `ThisAddIn` soit complètement initialisé.

**Solution** :
1. Modifiez `InitializeAddInComponents()` dans `ThisAddIn.cs`
2. Enregistrez l'événement APRÈS la création du TaskPane :

```csharp
_taskPaneControl = new TaskPaneControl();
_customTaskPane = this.CustomTaskPanes.Add(_taskPaneControl, "LMStudio Assistant");

// Enregistrer l'événement MAINTENANT (pas dans InitializeEvents)
this.EmailSelected += _taskPaneControl.ThisAddIn_EmailSelected;
```

### Scénario B : SelectionChange ne se déclenche jamais

**Cause** : L'événement n'est pas enregistré ou Outlook bloque l'événement.

**Solution** :
1. Vérifiez les logs de démarrage
2. Ajoutez un log dans `SetupOutlookHandlers()` :

```csharp
Logger.Log($"Explorer actif : {explorer != null}");
Logger.Log($"Selection actuelle : {explorer?.Selection?.Count ?? 0}");
```

### Scénario C : L'événement se déclenche mais rien ne s'affiche

**Cause** : Le problème est dans `UpdateSelectedEmailInfo()`.

**Solution** :
1. Vérifiez le log `UpdateSelectedEmailInfo - _currentMailItem est NULL`
2. Si NULL ? Le MailItem n'est pas assigné correctement
3. Si non-NULL ? Problème d'affichage UI

## ?? Fichier de Diagnostic Complet

Créez ce fichier `DIAGNOSTIC.ps1` :

```powershell
# DIAGNOSTIC.ps1
Write-Host "=== DIAGNOSTIC LMSTUDIO ASSISTANT ===" -ForegroundColor Cyan

# 1. Vérifier si le complément est installé
Write-Host "`n1. Vérification du registre..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio"
if (Test-Path $regPath) {
    Write-Host "   ? Complément trouvé dans le registre" -ForegroundColor Green
    Get-ItemProperty $regPath | Format-List
} else {
    Write-Host "   ? Complément NON trouvé dans le registre" -ForegroundColor Red
}

# 2. Vérifier les logs
Write-Host "`n2. Vérification des logs..." -ForegroundColor Yellow
$logPath = "$env:APPDATA\OutlookLMStudio\logs.txt"
if (Test-Path $logPath) {
    Write-Host "   ? Fichier de logs trouvé" -ForegroundColor Green
    Write-Host "   ?? Chemin : $logPath" -ForegroundColor Cyan
    
    # Afficher les dernières lignes
    Write-Host "`n   ?? Dernières 20 lignes:" -ForegroundColor Cyan
    Get-Content $logPath -Tail 20
} else {
    Write-Host "   ? Fichier de logs NON trouvé" -ForegroundColor Red
}

# 3. Vérifier si Outlook est en cours
Write-Host "`n3. Vérification d'Outlook..." -ForegroundColor Yellow
$outlook = Get-Process outlook -ErrorAction SilentlyContinue
if ($outlook) {
    Write-Host "   ??  Outlook est en cours d'exécution" -ForegroundColor Yellow
    Write-Host "   ?? Fermez Outlook avant de réinstaller" -ForegroundColor Cyan
} else {
    Write-Host "   ? Outlook n'est pas en cours d'exécution" -ForegroundColor Green
}

# 4. Vérifier la DLL
Write-Host "`n4. Vérification de la DLL..." -ForegroundColor Yellow
$dllPath = "bin\Debug\OutlookLMStudio.dll"
if (Test-Path $dllPath) {
    $dll = Get-Item $dllPath
    Write-Host "   ? DLL trouvée" -ForegroundColor Green
    Write-Host "   ?? Chemin : $($dll.FullName)" -ForegroundColor Cyan
    Write-Host "   ?? Date : $($dll.LastWriteTime)" -ForegroundColor Cyan
    Write-Host "   ?? Taille : $([math]::Round($dll.Length / 1KB, 2)) KB" -ForegroundColor Cyan
} else {
    Write-Host "   ? DLL NON trouvée - Recompilez le projet!" -ForegroundColor Red
}

Write-Host "`n=== FIN DU DIAGNOSTIC ===" -ForegroundColor Cyan
Write-Host "`nConsultez les logs pour plus de détails:" -ForegroundColor Yellow
Write-Host "$logPath`n" -ForegroundColor Cyan
```

Exécutez :
```powershell
.\DIAGNOSTIC.ps1
```

## ?? Prochaines Étapes

1. **Recompilez** (déchargez OutlookLMStudio.Tests si nécessaire)
2. **Réinstallez** avec INSTALL_MANUAL.bat
3. **Testez** en cliquant sur un email
4. **Consultez les logs** : `%APPDATA%\OutlookLMStudio\logs.txt`
5. **Partagez** les logs pour diagnostic approfondi

---

**Objectif** : Identifier EXACTEMENT où le processus de sélection échoue  
**Méthode** : Logs de diagnostic à chaque étape  
**Résultat attendu** : Logs clairs indiquant le problème précis