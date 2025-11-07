# ?? Guide de Publication sur GitHub

## Étape 1 : Préparer la Release Localement

### 1.1 Compiler le Projet

```bash
# Dans Visual Studio
Build ? Clean Solution
Build ? Rebuild Solution
```

Vérifiez qu'il n'y a **aucune erreur** de compilation.

### 1.2 Créer l'Archive de Release

```powershell
# Dans PowerShell (dans le dossier du projet)
.\CREATE_RELEASE.ps1
```

Ce script va :
- ? Vérifier que la compilation est réussie
- ? Créer le dossier `Release/`
- ? Copier tous les fichiers nécessaires
- ? Créer l'archive `OutlookLMStudio-v1.0.0.zip`

**Résultat** : Un fichier `OutlookLMStudio-v1.0.0.zip` (~2-3 MB)

---

## Étape 2 : Commit et Push sur GitHub

### 2.1 Vérifier les Changements

```bash
git status
```

Vous devriez voir les nouveaux fichiers :
- `LICENSE`
- `CHANGELOG.md`
- `RELEASE_NOTES_v1.0.0.md`
- `CREATE_RELEASE.ps1`
- `GITHUB_PUBLICATION_GUIDE.md`

### 2.2 Commit des Changements

```bash
# Ajouter tous les nouveaux fichiers
git add LICENSE CHANGELOG.md RELEASE_NOTES_v1.0.0.md CREATE_RELEASE.ps1 GITHUB_PUBLICATION_GUIDE.md

# Commit
git commit -m "chore: Préparation de la release v1.0.0

- Ajout de la licence MIT
- Ajout du CHANGELOG
- Ajout des notes de release
- Ajout du script de création de release
- Documentation de publication"

# Push vers GitHub
git push origin master
```

---

## Étape 3 : Créer la Release sur GitHub

### 3.1 Accéder à la Page des Releases

1. Allez sur : https://github.com/iphonebm/OutlookLMStudio
2. Cliquez sur **"Releases"** (dans la barre latérale droite)
3. Cliquez sur **"Create a new release"** (ou "Draft a new release")

### 3.2 Configurer la Release

#### Tag Version
```
v1.0.0
```
- Cliquez sur "Choose a tag"
- Tapez `v1.0.0`
- Cliquez "Create new tag: v1.0.0 on publish"

#### Target
```
master
```
(Laissez par défaut)

#### Release Title
```
?? OutlookLMStudio v1.0.0 - Première Release Stable
```

#### Description

Copiez-collez le contenu de `RELEASE_NOTES_v1.0.0.md` dans la zone de description.

**OU** utilisez ce template court :

```markdown
# ?? Première Version Stable !

**OutlookLMStudio** est un complément Microsoft Outlook qui utilise LMStudio pour générer automatiquement des réponses d'emails professionnelles via IA locale.

## ? Fonctionnalités

- ? Génération simple et multiple (batch) de réponses
- ? Menu contextuel + Volet des tâches + Bouton Ribbon
- ? Sélection de modèle LMStudio depuis l'interface
- ? 100% local - Aucune donnée envoyée sur Internet
- ? Barre de progression et gestion d'erreurs
- ? Logs détaillés pour diagnostic

## ?? Performances

- 1 email : ~3-5 secondes
- 10 emails : ~30-50 secondes
- 20 emails : ~1-2 minutes

**Gagnez 90% de temps** sur vos réponses emails ! ??

## ?? Installation

1. Téléchargez `OutlookLMStudio-v1.0.0.zip` ci-dessous
2. Décompressez
3. Exécutez `INSTALL.bat` en admin
4. Lancez Outlook

## ?? Documentation

- [README complet](https://github.com/iphonebm/OutlookLMStudio/blob/master/README.md)
- [Changelog](https://github.com/iphonebm/OutlookLMStudio/blob/master/CHANGELOG.md)

## ?? Prérequis

- Windows 10/11
- Microsoft Outlook (2013+)
- .NET Framework 4.7.2+
- LMStudio ([télécharger](https://lmstudio.ai/))

---

**Première version publique** - Si vous rencontrez un problème, ouvrez une [Issue](https://github.com/iphonebm/OutlookLMStudio/issues) !
```

### 3.3 Uploader le Fichier ZIP

1. Faites défiler vers le bas jusqu'à **"Attach binaries"**
2. Cliquez sur la zone ou glissez-déposez `OutlookLMStudio-v1.0.0.zip`
3. Attendez la fin de l'upload

### 3.4 Options Supplémentaires

- ? Cochez **"Set as the latest release"**
- ? **NE PAS** cocher "This is a pre-release" (c'est une version stable)

### 3.5 Publier

Cliquez sur le bouton vert **"Publish release"**

---

## Étape 4 : Vérifier la Release

### 4.1 Vérifier la Page

1. Revenez sur https://github.com/iphonebm/OutlookLMStudio
2. Vous devriez voir un badge **"Latest"** à côté de v1.0.0
3. Le badge de release dans le README devrait pointer vers v1.0.0

### 4.2 Tester le Téléchargement

1. Cliquez sur le fichier ZIP
2. Il devrait se télécharger correctement
3. Testez l'installation en suivant les instructions

---

## Étape 5 : Communiquer la Release

### 5.1 Mettre à Jour le README

Le README contient déjà le lien vers les releases :
```markdown
[Releases](https://github.com/iphonebm/OutlookLMStudio/releases)
```

### 5.2 Partager (Optionnel)

Vous pouvez partager votre release sur :
- Reddit (r/selfhosted, r/LocalLLaMA, r/Outlook)
- Twitter/X
- LinkedIn
- Forums techniques

Template de message :
```
?? OutlookLMStudio v1.0.0 est disponible !

Automatisez vos réponses emails avec LMStudio directement dans Outlook.
? 100% local & privé
? Génération unique ou batch
? Gagnez 90% de temps

https://github.com/iphonebm/OutlookLMStudio
```

---

## ? Checklist Finale

Avant de publier, vérifiez :

- [ ] Le projet compile sans erreur
- [ ] Le fichier ZIP est créé (`OutlookLMStudio-v1.0.0.zip`)
- [ ] LICENSE est ajouté au repository
- [ ] CHANGELOG.md est à jour
- [ ] README.md contient les bons liens
- [ ] Les fichiers sont commit et push sur master
- [ ] Le tag v1.0.0 est créé sur GitHub
- [ ] Le ZIP est uploadé dans la release
- [ ] La release est marquée comme "Latest"
- [ ] La description de la release est complète

---

## ?? Commandes Rapides

```bash
# Tout en une fois
Build ? Rebuild Solution
.\CREATE_RELEASE.ps1

git add LICENSE CHANGELOG.md RELEASE_NOTES_v1.0.0.md CREATE_RELEASE.ps1
git commit -m "chore: Release v1.0.0"
git push origin master

# Puis créez la release manuellement sur GitHub
```

---

## ?? En Cas de Problème

### Le script CREATE_RELEASE.ps1 échoue

Vérifiez que :
- Le projet a été compilé (bin\Debug\ existe)
- Vous êtes dans le bon dossier
- PowerShell peut exécuter des scripts

### Le fichier ZIP est trop gros

Vérifiez que vous n'incluez pas :
- Les fichiers `.pdb` (debug)
- Les fichiers `.xml` (documentation)
- Les manifestes inutiles

### La release n'apparaît pas

Attendez quelques minutes, GitHub peut avoir un délai.
Rafraîchissez la page avec `Ctrl + F5`.

---

## ?? Support

Si vous avez besoin d'aide pour publier la release :
- Ouvrez une [Discussion](https://github.com/iphonebm/OutlookLMStudio/discussions)
- Consultez la [documentation GitHub](https://docs.github.com/en/repositories/releasing-projects-on-github)

---

**Bonne publication ! ??**