# ?? Résolution Erreur : "Application déjà installée"

## ? L'Erreur Complète

```
System.Deployment.Application.DeploymentException: 
Impossible d'installer cette application car une application disposant 
de la même identité est déjà installée. Pour installer cette application, 
modifiez la version du manifeste correspondante ou désinstallez 
l'application préexistante.
```

## ? Solution Automatique (RECOMMANDÉ)

### Étape 1 : Exécutez UNINSTALL.bat

```cmd
1. Clic droit sur UNINSTALL.bat
2. "Exécuter en tant qu'administrateur"
3. Attendez que ça finisse
```

Le script va :
- ? Fermer Outlook
- ? Supprimer les clés de registre
- ? Nettoyer le cache ClickOnce
- ? Tout préparer pour une réinstallation propre

### Étape 2 : Exécutez INSTALL.bat

```cmd
1. Clic droit sur INSTALL.bat
2. "Exécuter en tant qu'administrateur"
3. Le script fait tout automatiquement
```

Le nouveau script INSTALL.bat :
- ? Désinstalle automatiquement l'ancienne version
- ? Nettoie le cache
- ? Installe la nouvelle version
- ? Lance Outlook

### Étape 3 : Vérifiez

Dans Outlook :
1. Le volet "LMStudio Assistant" devrait apparaître
2. Si ce n'est pas le cas, consultez la section "Dépannage" ci-dessous

## ??? Solution Manuelle (Si le script ne fonctionne pas)

### Méthode 1 : Via le Panneau de Configuration

1. **Ouvrez** : Panneau de configuration ? Programmes ? Désinstaller un programme
2. **Recherchez** : "OutlookLMStudio" dans la liste
3. **Désinstallez**
4. **Redémarrez** votre ordinateur
5. **Réinstallez** : Double-clic sur `bin\Debug\OutlookLMStudio.vsto`

### Méthode 2 : Nettoyage Manuel Complet

#### A. Fermer Outlook

```
Gestionnaire des tâches (Ctrl+Shift+Esc)
Onglet "Processus"
Terminez "Microsoft Outlook" si présent
```

#### B. Supprimer les Clés de Registre

Ouvrez **Registry Editor** (Win + R ? `regedit`) :

```
1. Naviguez vers :
   HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookLMStudio
   
2. Clic droit ? Supprimer

3. Naviguez vers :
   HKEY_CURRENT_USER\Software\Microsoft\VSTO\Security\Inclusion\
   
4. Cherchez une clé avec "{75384258-9a61-432d-b12a-d48c8e01ce3a}"
   
5. Clic droit ? Supprimer
```

#### C. Nettoyer le Cache ClickOnce

```
1. Ouvrez l'Explorateur de fichiers
2. Naviguez vers : %LOCALAPPDATA%\Apps\2.0
3. Supprimez tout le dossier "2.0"
   (Peut demander les droits administrateur)
```

#### D. Redémarrez

```
Redémarrez votre ordinateur
Cela garantit que toutes les ressources sont libérées
```

#### E. Réinstallez

```
1. Allez dans : bin\Debug\
2. Double-cliquez sur : OutlookLMStudio.vsto
3. Suivez l'assistant d'installation
```

## ?? Vérification Après Installation

### 1. Dans Outlook

**Fichier** ? **Options** ? **Compléments** :

- Vous devriez voir "OutlookLMStudio" dans "Compléments actifs"
- LoadBehavior devrait être "3" (chargé au démarrage)

### 2. Le Volet Apparaît

Le volet "LMStudio Assistant" devrait être visible à droite.

Si ce n'est pas le cas :
- **Affichage** ? **Volets des tâches**
- Cherchez "LMStudio Assistant"

### 3. Les Logs Sont Créés

Vérifiez que le fichier existe :
```
%APPDATA%\OutlookLMStudio\logs.txt
```

Ouvrez-le et vérifiez :
```
=== Démarrage du complément OutlookLMStudio ===
Initialisation des composants du complément
Composants initialisés avec succès
```

## ?? Problèmes Persistants

### Le complément apparaît dans "Compléments désactivés"

**Cause** : Outlook a désactivé le complément après une erreur

**Solution** :
1. Fichier ? Options ? Compléments
2. En bas : "Gérer: **Compléments désactivés**" ? Atteindre
3. Sélectionnez OutlookLMStudio
4. Cliquez "Toujours activer ce complément"
5. Redémarrez Outlook

### Erreur "Trust not granted"

**Cause** : Problème de sécurité ClickOnce

**Solution** :
```powershell
# Exécutez en PowerShell en Administrateur :
mage -cc
```

Puis réinstallez.

### L'erreur revient toujours

**Cause** : Cache ClickOnce corrompu ou version .NET problématique

**Solution Radicale** :
```
1. UNINSTALL.bat en Administrateur
2. Supprimez : %LOCALAPPDATA%\Apps (tout le dossier)
3. Supprimez : C:\Users\Ash\source\repos\OutlookLMStudio\bin
4. Supprimez : C:\Users\Ash\source\repos\OutlookLMStudio\obj
5. Redémarrez l'ordinateur
6. Ouvrez Visual Studio en Administrateur
7. Build ? Rebuild Solution
8. INSTALL.bat en Administrateur
```

## ?? Prévention

### Pour éviter ce problème à l'avenir :

1. **Toujours utiliser UNINSTALL.bat** avant de réinstaller
2. **Fermer Outlook** avant toute modification
3. **Utiliser INSTALL.bat** qui gère automatiquement les conflits
4. **Ne pas modifier** manuellement les clés de registre (utiliser les scripts)

### Lors du développement :

```
1. Fermez Outlook
2. Build ? Rebuild Solution (en Administrateur)
3. INSTALL.bat (désinstalle et réinstalle automatiquement)
4. Lancez Outlook
```

## ?? Checklist de Dépannage

Avant de demander de l'aide, vérifiez :

- [ ] UNINSTALL.bat exécuté en Administrateur
- [ ] Outlook complètement fermé (vérifier Gestionnaire des tâches)
- [ ] Cache ClickOnce supprimé (%LOCALAPPDATA%\Apps\2.0)
- [ ] Ordinateur redémarré
- [ ] Visual Studio exécuté en Administrateur
- [ ] Rebuild Solution effectué
- [ ] INSTALL.bat exécuté en Administrateur
- [ ] Aucune autre instance du complément dans Programmes et Fonctionnalités
- [ ] Pas d'antivirus bloquant l'installation

## ? Scripts Fournis

Utilisez ces scripts pour éviter les manipulations manuelles :

| Script | Utilisation | Besoin Admin |
|--------|-------------|--------------|
| `INSTALL.bat` | Installation (désinstalle automatiquement l'ancienne version) | ? Oui |
| `UNINSTALL.bat` | Désinstallation complète | ? Oui |
| `Diagnostic.ps1` | Diagnostic des problèmes | ? Non |

## ?? Comprendre le Problème

### Pourquoi ça arrive ?

ClickOnce (le système d'installation VSTO) utilise :
- Une **identité unique** (GUID) : `{75384258-9a61-432d-b12a-d48c8e01ce3a}`
- Un **cache** dans `%LOCALAPPDATA%\Apps\2.0`
- Des **clés de registre** pour l'enregistrement

Si vous rebuilder sans désinstaller, ClickOnce voit :
- ? Même GUID (identité)
- ? Fichiers différents (nouveau build)
- ? Conflit ? Erreur "application déjà installée"

### La Solution

Les scripts `INSTALL.bat` et `UNINSTALL.bat` :
- ? Nettoient tout avant installation
- ? Évitent les conflits
- ? Garantissent une installation propre

## ?? Résumé Rapide

**3 étapes pour résoudre :**

```
1. Clic droit UNINSTALL.bat ? Admin ? Exécuter
2. Clic droit INSTALL.bat ? Admin ? Exécuter  
3. Lancez Outlook ? Profitez !
```

C'est tout ! ??

---

**Besoin d'aide ?** Consultez `TROUBLESHOOTING.md` ou les logs dans `%APPDATA%\OutlookLMStudio\logs.txt`