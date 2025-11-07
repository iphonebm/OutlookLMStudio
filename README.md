# ?? OutlookLMStudio

> Générez automatiquement des réponses d'emails professionnelles avec LMStudio directement depuis Outlook

[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.7.2-blue.svg)](https://dotnet.microsoft.com/)
[![Office](https://img.shields.io/badge/Microsoft%20Outlook-2013%2B-orange.svg)](https://www.microsoft.com/en-us/microsoft-365/outlook)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

**OutlookLMStudio** est un complément Microsoft Outlook qui exploite la puissance de LMStudio pour générer des réponses d'emails professionnelles en quelques secondes, tout en gardant vos données **100% locales**.

---

## ? Fonctionnalités

### ?? Génération Intelligente de Réponses
- **Génération unique** : Sélectionnez un email et générez une réponse personnalisée
- **Génération multiple** : Sélectionnez plusieurs emails et traitez-les en batch
- **Templates personnalisables** : Adaptez le style des réponses selon vos besoins

### ??? Interfaces Multiples
- **Volet des tâches** : Interface simple et intuitive
- **Menu contextuel** : Clic droit sur n'importe quel email
- **Bouton Ribbon** : Accès rapide depuis la barre d'outils Outlook

### ?? Confidentialité & Sécurité
- **100% local** : Aucune donnée envoyée sur Internet
- **Modèles LLM locaux** : Utilisez vos propres modèles via LMStudio
- **Contrôle total** : Vous gardez le contrôle de vos données

### ? Performance
- **Traitement par lots** : Générez des réponses pour 20+ emails en quelques minutes
- **Barre de progression** : Visualisez l'avancement en temps réel
- **Gestion des erreurs** : Rapport détaillé des succès/échecs

---

## ?? Installation

### Prérequis

- **Windows 10/11**
- **Microsoft Outlook** (2013, 2016, 2019, 2021, 365)
- **.NET Framework 4.7.2** ou supérieur
- **VSTO Runtime** (généralement déjà installé avec Office)
- **LMStudio** ([télécharger ici](https://lmstudio.ai/))

### Installation Rapide

1. **Téléchargez** la dernière version depuis [Releases](https://github.com/iphonebm/OutlookLMStudio/releases)

2. **Décompressez** l'archive

3. **Exécutez** `INSTALL.bat` en tant qu'administrateur :
   ```cmd
   Clic droit sur INSTALL.bat ? Exécuter en tant qu'administrateur
   ```

4. **Lancez Outlook** - Le complément apparaîtra automatiquement

### Installation depuis les sources

```bash
# Cloner le repository
git clone https://github.com/iphonebm/OutlookLMStudio.git
cd OutlookLMStudio

# Ouvrir dans Visual Studio
start OutlookLMStudio.sln

# Compiler le projet
Build ? Rebuild Solution

# Installer
INSTALL.bat (en tant qu'administrateur)
```

---

## ?? Utilisation

### Configuration de LMStudio

1. **Lancez LMStudio**
2. **Chargez un modèle** (ex: Mistral 7B, Llama 2, etc.)
3. **Démarrez le serveur local** :
   - Onglet "Local Server"
   - Cliquez "Start Server"
   - Par défaut sur `http://localhost:1234`

### Générer une Réponse Simple

1. **Sélectionnez** un email dans Outlook
2. **Méthode 1** - Volet des tâches :
   - Cliquez sur "Générer une réponse"
3. **Méthode 2** - Menu contextuel :
   - Clic droit ? "Générer Réponse(s) avec LMStudio"
4. **Vérifiez** le brouillon créé et modifiez si nécessaire
5. **Envoyez** !

### Génération Multiple (Batch)

1. **Sélectionnez plusieurs emails** :
   - `Ctrl + Clic` : Sélection multiple
   - `Shift + Clic` : Plage d'emails
2. **Clic droit** ? "Générer Réponse(s) avec LMStudio"
3. **Confirmez** dans la boîte de dialogue
4. **Attendez** la fin du traitement (barre de progression)
5. **Consultez** les brouillons créés

**Exemple** : Traitez 20 emails en ~2 minutes au lieu de 1h40 manuellement ! ??

---

## ?? Configuration

### Paramètres du Complément

Accédez aux paramètres via le bouton "Paramètres" dans le volet :

| Paramètre | Description | Défaut |
|-----------|-------------|--------|
| **URL API** | Adresse du serveur LMStudio | `http://localhost:1234` |
| **Modèle** | Nom du modèle à utiliser | Sélection depuis LMStudio |
| **Température** | Créativité des réponses (0.0-1.0) | `0.7` |
| **Max Tokens** | Longueur maximale de la réponse | `2000` |
| **Timeout** | Délai d'attente par email (secondes) | `30` |

---

## ?? Performances

### Temps de Traitement

| Nombre d'Emails | Temps Moyen | Gain vs Manuel |
|-----------------|-------------|----------------|
| 1 email | ~3-5 secondes | ~90% |
| 10 emails | ~30-50 secondes | ~95% |
| 20 emails | ~1-2 minutes | ~98% |

*Basé sur un modèle Mistral 7B sur un PC moyen*

### Modèles Recommandés

| Modèle | Taille | Vitesse | Qualité | Usage |
|--------|--------|---------|---------|-------|
| **Mistral 7B** | 7B | ??? | ???? | Recommandé |
| **Llama 2 7B** | 7B | ??? | ??? | Bon |
| **Phi-2** | 2.7B | ???? | ??? | Rapide |
| **Llama 2 13B** | 13B | ?? | ????? | Haute qualité |

---

## ??? Développement

### Structure du Projet

```
OutlookLMStudio/
??? ThisAddIn.cs              # Point d'entrée principal
??? TaskPaneControl.cs        # Interface du volet
??? SettingsForm.cs           # Fenêtre de paramètres
??? LMStudioSettings.cs       # Gestion de la configuration
??? ContextMenuHandler.cs     # Gestion du menu contextuel
??? LMStudioRibbon.cs         # Interface Ribbon
??? Models/
?   ??? LMStudioResponse.cs   # Modèles de données API
??? INSTALL.bat               # Script d'installation
??? UNINSTALL.bat             # Script de désinstallation
??? README.md                 # Ce fichier
```

### Technologies Utilisées

- **.NET Framework 4.7.2**
- **VSTO (Visual Studio Tools for Office)**
- **Microsoft Office Interop**
- **Newtonsoft.Json** - Sérialisation JSON
- **LMStudio API** - Interface avec les modèles LLM

### Compilation

```bash
# Prérequis
- Visual Studio 2019/2022
- Workload "Office/SharePoint development"

# Compiler
Build ? Rebuild Solution
```

---

## ?? Dépannage

### Le complément n'apparaît pas dans Outlook

1. **Vérifiez l'installation** :
   - Fichier ? Options ? Compléments
   - Cherchez "OutlookLMStudio" dans la liste

2. **S'il est désactivé** :
   - Gérer : Compléments désactivés ? Atteindre
   - Sélectionnez OutlookLMStudio
   - Cliquez "Toujours activer"

3. **Réinstallez** :
   ```cmd
   UNINSTALL.bat (en admin)
   INSTALL.bat (en admin)
   ```

### Erreur "LMStudio n'est pas accessible"

1. **Vérifiez que LMStudio est démarré**
2. **Vérifiez qu'un modèle est chargé**
3. **Testez l'API** :
   ```
   Ouvrez http://localhost:1234/v1/models dans un navigateur
   ```

### Les réponses sont vides

1. **Augmentez le timeout** dans les paramètres (60 secondes)
2. **Vérifiez les logs** : `%APPDATA%\OutlookLMStudio\logs.txt`
3. **Testez avec un modèle plus petit** (ex: Phi-2)

---

## ?? Logs & Diagnostic

Les logs sont enregistrés automatiquement :

```
%APPDATA%\OutlookLMStudio\logs.txt
```

Exécutez le script de diagnostic :

```powershell
.\DIAGNOSTIC.ps1
```

---

## ?? Contribution

Les contributions sont les bienvenues ! Voici comment participer :

1. **Fork** le projet
2. **Créez** une branche (`git checkout -b feature/AmazingFeature`)
3. **Committez** vos changements (`git commit -m 'Add AmazingFeature'`)
4. **Pushez** vers la branche (`git push origin feature/AmazingFeature`)
5. **Ouvrez** une Pull Request

### Idées de Contributions

- ?? Support de nouvelles langues
- ?? Amélioration de l'interface
- ?? Support d'autres backends LLM (Ollama, GPT4All, etc.)
- ?? Templates de réponses prédéfinis
- ?? Analyse de sentiment des emails
- ? Optimisations de performance

---

## ??? Roadmap

### Version 1.4.0 (Prochaine)
- [ ] Support d'Ollama en plus de LMStudio
- [ ] Templates de réponses multiples
- [ ] Détection automatique de la langue
- [ ] Interface de gestion des modèles

### Version 2.0.0
- [ ] Support de Microsoft Teams
- [ ] Analyse de sentiment
- [ ] Suggestions de follow-up
- [ ] Mode hors-ligne avec cache

---

## ?? Licence

Ce projet est sous licence **MIT** - voir le fichier [LICENSE](LICENSE) pour plus de détails.

---

## ?? Remerciements

- **[LMStudio](https://lmstudio.ai/)** - Pour leur excellent logiciel de gestion de modèles LLM
- **Microsoft** - Pour VSTO et les outils de développement Office
- **La communauté Open Source** - Pour les nombreuses bibliothèques utilisées

---

## ?? Contact & Support

- ?? **Bugs** : [Ouvrir une Issue](https://github.com/iphonebm/OutlookLMStudio/issues)
- ?? **Discussions** : [GitHub Discussions](https://github.com/iphonebm/OutlookLMStudio/discussions)

---

<div align="center">

**Fait avec ?? pour automatiser vos emails**

Si ce projet vous est utile, n'hésitez pas à lui donner une ? !

</div>