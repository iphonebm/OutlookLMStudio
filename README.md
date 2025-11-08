# 📧 OutlookLMStudio

> FR: Générez automatiquement des réponses d'emails professionnelles dans Microsoft Outlook via des modèles IA locaux (LMStudio). 100% privé.
> EN: Generate professional email replies in Microsoft Outlook using local AI language models through LMStudio. 100% private.

[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.7.2-blue.svg)](https://dotnet.microsoft.com/)
[![Outlook](https://img.shields.io/badge/Outlook-2013%2B-orange.svg)](https://www.microsoft.com/microsoft-365/outlook)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## 🌍 Langues / Languages
FR par défaut. English section below.

---

## 📑 Sommaire
1. 📌 Présentation  
2. 🚀 Fonctionnalités  
3. 🛠️ Installation  
4. ⚙️ Configuration LMStudio  
5. ✉️ Utilisation  
6. 🧬 Paramètres & Template  
7. ⚡ Performances & Modèles  
8. 🧱 Structure & Développement  
9. 🔒 Sécurité & Confidentialité  
10. 🩺 Dépannage / FAQ  
11. 🧾 Logs & Diagnostic  
12. 🤝 Contribution  
13. 🗺️ Roadmap  
14. 🧪 Idées Futures  
15. 📄 Licence  
16. 🙏 Remerciements  
17. 🆘 Support  
18. 🇬🇧 English Section  

---

## 📌 Présentation
OutlookLMStudio est un complément Outlook (VSTO) qui accélère la rédaction de vos emails en générant automatiquement des réponses contextualisées. Tout fonctionne localement via LMStudio : aucune donnée ne quitte votre machine.

Cas d’usage :
- Support client / helpdesk  
- Préqualification commerciale  
- Réponses internes récurrentes  
- Traitement batch d’arriérés (backlog)

---

## 🚀 Fonctionnalités
| Catégorie        | Détails |
|------------------|---------|
| Génération       | Réponses uniques ou par lot (20+ emails) |
| Intégration      | Volet, menu contextuel, bouton Ribbon |
| Modèles          | Tout modèle chargé dans LMStudio local |
| Personnalisation | Template de prompt modifiable |
| Paramètres       | Température, tokens max, timeout |
| Performance      | Barre de progression batch |
| Robustesse       | Gestion fine des erreurs, logs |
| Confidentialité  | 100% local, aucune API externe |

### ✅ Points forts
- Gain de temps massif  
- Qualité adaptable selon le modèle choisi  
- Aucune fuite de données sensibles  

---

## 🛠️ Installation

### Prérequis
| Composant | Version / Remarque |
|-----------|--------------------|
| Windows   | 10 / 11 |
| Outlook   | 2013, 2016, 2019, 2021, 365 |
| .NET      | .NET Framework 4.7.2+ |
| VSTO      | Runtime (souvent déjà présent) |
| LMStudio  | [Télécharger](https://lmstudio.ai/) |
| Visual Studio (optionnel) | 2019/2022 pour compiler |

### Installation rapide
1. Télécharger la release : [Releases](https://github.com/iphonebm/OutlookLMStudio/releases)  
2. Extraire l’archive  
3. Clic droit `INSTALL.bat` → Exécuter en tant qu’administrateur  
4. Ouvrir Outlook  

### Depuis les sources
```bash
git clone https://github.com/iphonebm/OutlookLMStudio.git
cd OutlookLMStudio
start OutlookLMStudio.sln
# Build → Rebuild Solution
INSTALL.bat  # admin
```
Désinstallation : `UNINSTALL.bat`.

---

## ⚙️ Configuration LMStudio
1. Ouvrir LMStudio  
2. Charger un modèle (ex : Mistral 7B, Llama 2, Phi-2)  
3. Onglet “Local Server” → Start Server  
4. URL par défaut : `http://localhost:1234`  
5. Vérifier : `http://localhost:1234/v1/models`  

---

## ✉️ Utilisation

### Réponse simple
1. Sélectionner un email  
2. Volet → “Générer une réponse” OU clic droit → “Générer Réponse(s)”  
3. Revoir le brouillon  
4. Envoyer  

### Batch
1. Sélection multiple (`Ctrl + clic` / `Shift + clic`)  
2. Clic droit → Générer  
3. Suivre la barre de progression  
4. Ouvrir brouillons générés  

> Astuce : ~20 emails ≈ 2 minutes (vs >1h manuel).

---

## 🧬 Paramètres & Template
| Paramètre       | Description                        | Défaut |
|-----------------|------------------------------------|--------|
| URL API         | Endpoint LMStudio                  | http://localhost:1234 |
| Modèle          | Nom du modèle chargé               | (dynamique) |
| Température     | Créativité (0–1)                   | 0.7 |
| Max Tokens      | Longueur maximale de la réponse    | 2000 |
| Timeout (s)     | Délai max par email                | 30 |
| Prompt Template | Style rédaction                    | Personnalisable |

### Exemple de template
```
Système : Vous êtes un assistant professionnel.
Générez une réponse courtoise et adaptée au ton de l'email original.

Email source :
{emailContent}

Rédigez une réponse pertinente, concise et professionnelle.
```
`{emailContent}` sera remplacé automatiquement.

---

## ⚡ Performances & Modèles

| Emails | Temps (approx) | Gain vs manuel |
|--------|----------------|----------------|
| 1      | 3–5 s          | ~90%           |
| 10     | 30–50 s        | ~95%           |
| 20     | 60–120 s       | ~98%           |

| Modèle       | Taille | Vitesse | Qualité | Usage |
|--------------|--------|---------|---------|-------|
| Mistral 7B   | 7B     | Bonne   | Bonne   | Général |
| Llama 2 7B   | 7B     | Moyenne | Bonne   | Ton neutre |
| Phi-2        | 2.7B   | Très rapide | Moyenne | Volume |
| Llama 2 13B  | 13B    | Plus lente | Haute | Nuance |

---

## 🧱 Structure & Développement
```
OutlookLMStudio/
  ThisAddIn.cs
  TaskPaneControl.cs
  LMStudioSettings.cs
  ContextMenuHandler.cs
  LMStudioRibbon.cs
  Models/
    LMStudioResponse.cs
  INSTALL.bat
  UNINSTALL.bat
  README.md
```

### Stack Technique
- C# (.NET Framework 4.7.2)  
- VSTO / Office Interop  
- PowerShell (scripts de diagnostic)  
- Batch (scripts d’installation)  
- Newtonsoft.Json  
- LMStudio Local REST API  

### Compilation
```bash
dotnet build OutlookLMStudio.csproj
```

---

## 🔒 Sécurité & Confidentialité
- Aucune requête vers Internet  
- Données uniquement locales  
- Dépend des politiques internes Outlook / Windows  
- Pour données sensibles : privilégier des modèles stables (Mistral, Llama 2)  

---

## 🩺 Dépannage / FAQ

| Problème               | Cause probable        | Solution |
|------------------------|-----------------------|----------|
| Add-in absent          | Désactivé par Outlook | Activer dans Options → Compléments |
| “LMStudio inaccessible” | Serveur non lancé     | Démarrer Local Server |
| Réponses vides         | Timeout trop court     | Passer à 60 s |
| Batch lent             | Modèle trop lourd      | Essayer Phi-2 ou Mistral 7B |
| Modèle introuvable     | Nom incorrect          | Vérifier `/v1/models` |

Logs : `%APPDATA%\OutlookLMStudio\logs.txt`  
Diagnostic : `DIAGNOSTIC.ps1` (si présent)

---

## 🧾 Logs & Diagnostic
```
%APPDATA%\OutlookLMStudio\logs.txt
```
Script (optionnel) :
```powershell
./DIAGNOSTIC.ps1
```

---

## 🤝 Contribution
1. Fork  
2. Branche : `feature/NouvelleFonction`  
3. Commit : `git commit -m "Ajoute NouvelleFonction"`  
4. Push  
5. Pull Request  

### Idées
- Support Ollama / GPT4All  
- Templates prêts à l’emploi  
- Analyse de sentiment  
- Auto-détection de langue  
- Suggestions de suivi intelligentes  
- UI améliorée  

---

## 📄 Licence
MIT — voir [LICENSE](LICENSE).

---

## 🙏 Remerciements
- [LMStudio](https://lmstudio.ai/)  
- Microsoft  
- Communauté open source  

---

## 🆘 Support
- Bugs : Issues GitHub  
- Améliorations : Pull Requests  
- ⭐ Si utile, laissez une étoile  

---

## 🇬🇧 English Section

### Overview
OutlookLMStudio is an Outlook VSTO add-in that drafts professional, context-aware email replies using local LLMs served by LMStudio. No data leaves your machine.

### Key Features
- Single or batch reply generation  
- Task pane, ribbon button, context menu integration  
- Editable prompt template & runtime parameters  
- Local-only execution (privacy first)  
- Progress + granular error reporting  

### Quick Start
1. Start LMStudio Local Server (`http://localhost:1234`)  
2. Load a model (Mistral / Llama / Phi-2)  
3. Select email → Generate reply  
4. Review draft → Send  

### Parameters
Same as French table.

### Troubleshooting
- Missing add-in → Enable in Outlook Add-ins  
- Empty reply → Increase timeout  
- Slow batch → Use a lighter model  

### Contributing
Fork → Branch → Commit → PR.

---

<div align="center">
Fait avec 💡 et ⚙️ — Made with focus on privacy & productivity.<br/>
Si ce projet vous aide, laissez une ⭐ / If useful, leave a ⭐
</div>
