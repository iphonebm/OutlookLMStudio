# Changelog

Toutes les modifications notables de ce projet seront documentées dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/),
et ce projet adhère au [Semantic Versioning](https://semver.org/lang/fr/).

## [1.0.0] - 2025-01-XX

### ?? Première Release Publique

#### Ajouté
- Interface volet des tâches avec statut visuel (icône colorée)
- Menu contextuel pour génération rapide via clic droit
- Bouton Ribbon dans la barre d'outils Outlook
- Génération de réponse pour un email unique
- Génération multiple (batch) avec barre de progression
- Sélection de modèle depuis l'interface des paramètres
- Configuration complète via interface graphique :
  - URL API LMStudio
  - Sélection du modèle
  - Température (0.0-1.0)
  - Max Tokens
  - Timeout
  - Template de prompt personnalisable
- Logs détaillés automatiques (`%APPDATA%\OutlookLMStudio\logs.txt`)
- Scripts d'installation/désinstallation automatiques (`INSTALL.bat`, `UNINSTALL.bat`)
- Script de diagnostic (`DIAGNOSTIC.ps1`)
- Support d'Outlook 2013, 2016, 2019, 2021, 365
- Icône de statut colorée :
  - ?? Vert : LMStudio connecté
  - ?? Rouge : LMStudio déconnecté
  - ?? Orange : Génération en cours (animé)
- Gestion d'erreurs détaillée avec rapport succès/échecs
- Documentation complète (README.md, guides d'installation)

#### Sécurité
- 100% local : Aucune donnée envoyée sur Internet
- Utilisation exclusive de modèles LLM locaux via LMStudio
- Aucune télémétrie ou tracking

#### Performances
- Génération rapide : ~3-5 secondes par email (Mistral 7B)
- Batch processing : 20 emails en ~1-2 minutes
- Pause configurable entre les emails (500ms par défaut)

#### Technique
- .NET Framework 4.7.2
- VSTO (Visual Studio Tools for Office)
- Microsoft Office Interop
- Newtonsoft.Json pour la sérialisation
- API LMStudio compatible OpenAI

---

## [Non publié]

### Prévu pour v1.1.0
- Support d'Ollama en plus de LMStudio
- Templates de réponses multiples prédéfinis
- Détection automatique de la langue des emails
- Interface de gestion des modèles chargés
- Amélioration de l'animation de l'icône de statut

### Prévu pour v1.2.0
- Cache des réponses générées
- Mode brouillon automatique (sans confirmation)
- Raccourcis clavier personnalisables
- Thèmes clairs/sombres pour l'interface

### Prévu pour v2.0.0
- Support de Microsoft Teams
- Analyse de sentiment des emails
- Suggestions de follow-up automatiques
- Mode hors-ligne avec cache
- Support de GPT4All et autres backends LLM

---

## Format des Versions

Le format de version suit le Semantic Versioning (MAJOR.MINOR.PATCH) :
- **MAJOR** : Changements incompatibles avec les versions précédentes
- **MINOR** : Nouvelles fonctionnalités rétrocompatibles
- **PATCH** : Corrections de bugs rétrocompatibles

---

## Types de Changements

- `Ajouté` : Nouvelles fonctionnalités
- `Modifié` : Changements dans les fonctionnalités existantes
- `Déprécié` : Fonctionnalités bientôt supprimées
- `Supprimé` : Fonctionnalités retirées
- `Corrigé` : Corrections de bugs
- `Sécurité` : Corrections de vulnérabilités

---

[1.0.0]: https://github.com/iphonebm/OutlookLMStudio/releases/tag/v1.0.0