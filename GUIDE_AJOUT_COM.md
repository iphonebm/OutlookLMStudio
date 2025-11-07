# ?? Guide Visuel : Ajouter la Référence COM

## ?? IMPORTANT : Code Déjà Activé !

Le code du menu contextuel a été **automatiquement décommenté**. 
Il ne vous reste plus qu'à **ajouter la référence COM** pour que ça compile.

## ?? Étapes à Suivre (5 minutes)

### Étape 1 : Ouvrir le Gestionnaire de Références

```
Explorateur de Solutions
    ??? OutlookLMStudio (clic droit)
        ??? Ajouter
            ??? Référence...
```

**OU** utilisez le raccourci clavier : `Alt + P, R`

### Étape 2 : Naviguer vers les Références COM

Dans la fenêtre "Gestionnaire de références" :

```
[Assemblys] [Projets] [COM] [Parcourir]
                       ^^^
                    CLIQUEZ ICI
```

Puis dans le panneau de gauche :

```
COM
??? Toutes les composantes
??? Type Libraries  ? CLIQUEZ ICI
```

### Étape 3 : Rechercher et Ajouter la Référence

Dans la barre de recherche en haut à droite, tapez :
```
Office
```

Vous devriez voir une liste apparaître. Recherchez :

**Pour Office 2016/2019/365/2021** :
```
? Microsoft Office 16.0 Object Library
   Version: 16.0.0.0
```

**Pour Office 2013** :
```
? Microsoft Office 15.0 Object Library
   Version: 15.0.0.0
```

**Pour Office 2010** :
```
? Microsoft Office 14.0 Object Library
   Version: 14.0.0.0
```

### Étape 4 : Cocher et Confirmer

1. ?? **Cochez** la case de la bibliothèque correspondant à votre version
2. Cliquez sur le bouton **"OK"** en bas à droite

### Étape 5 : Vérifier l'Ajout

Dans l'Explorateur de Solutions :

```
OutlookLMStudio
??? Propriétés
??? Références
?   ??? Microsoft.Office.Core  ? DOIT APPARAÎTRE
?   ??? Microsoft.Office.Interop.Outlook
?   ??? Newtonsoft.Json
?   ??? ... autres références
```

## ?? Étape 6 : Rebuild

1. **Clean Solution** : `Build` ? `Clean Solution`
2. **Rebuild Solution** : `Build` ? `Rebuild Solution` (ou `Ctrl + Shift + B`)

## ? Vérification

Si tout est bon, vous devriez voir :
```
========== Build: 1 succeeded, 0 failed, 0 up-to-date, 0 skipped ==========
OutlookLMStudio -> C:\Users\Ash\source\repos\OutlookLMStudio\bin\Debug\OutlookLMStudio.dll
```

## ?? Test du Menu Contextuel

1. Lancez Outlook
2. Ouvrez ou sélectionnez un email
3. **Clic droit** sur l'email
4. Vous devriez voir : **"Générer Réponse avec LMStudio"**
5. Cliquez dessus ? La réponse se génère !

## ?? Si la Référence N'Apparaît Pas

### Solution A : Vérifier l'Installation d'Office
La bibliothèque COM n'apparaît que si Office est installé correctement.

Vérifiez dans :
```
C:\Program Files\Microsoft Office\
OU
C:\Program Files (x86)\Microsoft Office\
```

### Solution B : Réenregistrer Office

Ouvrez PowerShell en tant qu'Administrateur :
```powershell
# Pour Office 64-bit
cd "C:\Program Files\Microsoft Office\Office16"
./EXCEL.EXE /regserver

# Pour Office 32-bit
cd "C:\Program Files (x86)\Microsoft Office\Office16"
./EXCEL.EXE /regserver
```

### Solution C : Utiliser le Volet des Tâches

Si l'ajout de la référence pose problème, le **volet des tâches fonctionne parfaitement** sans cette référence !

## ?? Comparaison des Méthodes

| Fonctionnalité | Volet des Tâches | Menu Contextuel |
|----------------|------------------|-----------------|
| Sélectionner email | ? Automatique | ? Clic droit |
| Générer réponse | ? 1 clic | ? 1 clic |
| Personnaliser prompt | ? Oui | ? Non |
| Voir statut connexion | ? Oui | ? Non |
| Configuration | ? Oui | ? Non |
| Référence COM requise | ? Non | ? Oui |

## ?? Questions Fréquentes

### Q : La référence est grisée et je ne peux pas la cocher
**R** : Cela signifie qu'elle est déjà ajoutée. Vérifiez dans les Références du projet.

### Q : J'ai plusieurs versions d'Office, laquelle choisir ?
**R** : Choisissez la version la plus récente installée sur votre machine.

### Q : Le build échoue avec "CS0234"
**R** : Redémarrez Visual Studio et réessayez. Parfois VS a besoin d'un redémarrage.

### Q : Le menu n'apparaît pas après le build
**R** : 
1. Fermez complètement Outlook
2. Supprimez les dossiers `bin` et `obj`
3. Rebuild en tant qu'Administrateur
4. Relancez Outlook

## ?? Besoin d'Aide ?

Si vous rencontrez des problèmes :

1. **Consultez les logs** : `%APPDATA%\OutlookLMStudio\logs.txt`
2. **Vérifiez les erreurs de build** dans Visual Studio
3. **Utilisez le volet des tâches** en attendant (fonctionne sans référence COM)

## ? Une Fois Configuré

Le menu contextuel sera disponible :
- ? Sur tous les emails (boîte de réception, envoyés, brouillons...)
- ? Dans toutes les fenêtres Outlook
- ? Pour tous les comptes configurés
- ? Même après redémarrage d'Outlook

---

**Temps estimé** : 5 minutes  
**Difficulté** : Facile  
**Prérequis** : Microsoft Office installé

Bonne chance ! ??