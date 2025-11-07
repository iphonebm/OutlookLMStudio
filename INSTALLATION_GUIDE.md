# ?? INSTALLATION RAPIDE - OutlookLMStudio

## ? VOTRE COMPLÉMENT EST PRÊT !

La DLL compile parfaitement. Il suffit maintenant de l'installer dans Outlook.

## ?? Installation Automatique (RECOMMANDÉ)

### Étape 1 : Clic Droit sur `INSTALL.bat`
```
1. Clic droit sur INSTALL.bat
2. "Exécuter en tant qu'administrateur"
3. Suivez les instructions
```

### Étape 2 : C'est Tout !
Le script va :
- ? Fermer Outlook
- ? Installer le complément
- ? Relancer Outlook
- ? Le volet "LMStudio Assistant" apparaît !

## ?? Installation Manuelle (Alternative)

Si le script ne fonctionne pas :

### Option A : Double-clic sur le .vsto
```
1. Allez dans: bin\Debug\
2. Double-cliquez sur: OutlookLMStudio.vsto
3. Cliquez sur "Installer"
4. Lancez Outlook
```

### Option B : Depuis Outlook
```
1. Ouvrez Outlook
2. Fichier ? Options ? Compléments
3. En bas : "Gérer: Compléments COM" ? Atteindre
4. Cliquez sur "Ajouter..."
5. Naviguez vers: bin\Debug\OutlookLMStudio.dll
6. Cliquez OK
7. Cochez OutlookLMStudio
8. Redémarrez Outlook
```

## ? Vérification

### Dans Outlook, vérifiez :

1. **Le volet apparaît** : "LMStudio Assistant" à droite

2. **Le complément est actif** :
   - Fichier ? Options ? Compléments
   - "Compléments actifs" devrait lister "OutlookLMStudio"

3. **Les logs sont créés** :
   - Ouvrez : `%APPDATA%\OutlookLMStudio\logs.txt`
   - Vous devriez voir : "=== Démarrage du complément OutlookLMStudio ==="

## ?? Première Utilisation

### 1. Lancez LMStudio
```
- Démarrez LMStudio
- Chargez un modèle
- Allez dans "Local Server"
- Cliquez "Start Server"
- Vérifiez que l'URL est: http://localhost:1234
```

### 2. Testez dans Outlook
```
- Sélectionnez un email
- Le volet affiche "Email sélectionné: [Sujet]"
- Cliquez "Générer une réponse"
- Un brouillon de réponse s'ouvre !
```

### 3. Configurez (Optionnel)
```
- Cliquez "Paramètres" dans le volet
- Ajustez :
  * URL de l'API
  * Timeout
  * Température du modèle
  * Nombre de tokens
- Sauvegardez
```

## ?? Dépannage

### "n'est pas un complément Office valide"

**Cause** : Le complément n'est pas enregistré

**Solution** :
1. Utilisez `INSTALL.bat` en tant qu'administrateur
2. OU double-cliquez sur le `.vsto`

### Le volet n'apparaît pas

**Cause** : Complément désactivé

**Solution** :
1. Fichier ? Options ? Compléments
2. En bas : "Gérer: **Compléments désactivés**" ? Atteindre
3. Si OutlookLMStudio est listé, sélectionnez-le et cliquez "Toujours activer"
4. Redémarrez Outlook

### "LMStudio n'est pas accessible"

**Cause** : LMStudio pas lancé ou mauvaise URL

**Solution** :
1. Lancez LMStudio
2. Démarrez le serveur local
3. Testez dans le navigateur : `http://localhost:1234/v1/models`
4. Si ça marche, c'est bon !

### Erreur au démarrage d'Outlook

**Cause** : Erreur dans l'initialisation

**Solution** :
1. Consultez les logs : `%APPDATA%\OutlookLMStudio\logs.txt`
2. La cause de l'erreur sera indiquée
3. Vérifiez que toutes les DLL sont présentes dans `bin\Debug\`

## ?? Fichiers Nécessaires

Dans `bin\Debug\`, vérifiez que ces fichiers existent :

```
? OutlookLMStudio.dll (complément principal)
? OutlookLMStudio.dll.manifest
? OutlookLMStudio.vsto (installateur)
? OutlookLMStudio.dll.config
? Newtonsoft.Json.dll
? Toutes les DLL Microsoft.Office.*
```

Si un fichier manque, faites un **Rebuild** dans Visual Studio.

## ?? Commandes Utiles

### Voir si le complément est chargé
```
Fichier ? Options ? Compléments ? Compléments actifs
```

### Voir les logs
```cmd
notepad %APPDATA%\OutlookLMStudio\logs.txt
```

### Désinstaller
```
Fichier ? Options ? Compléments ? Compléments COM ? Atteindre
Décochez OutlookLMStudio
```

### Réinstaller
```
1. Désinstallez d'abord
2. Fermez Outlook
3. Exécutez INSTALL.bat en administrateur
```

## ?? Astuces

### Logging Détaillé
Les logs enregistrent TOUT :
- Chaque démarrage
- Chaque email sélectionné
- Chaque requête à LMStudio
- Toutes les erreurs

Consultez-les en cas de problème !

### Performance
- **Première génération** : Peut prendre 10-30 secondes (LMStudio charge le modèle)
- **Générations suivantes** : Beaucoup plus rapides

### Personnalisation
Éditez le template de prompt dans le volet pour changer le style des réponses.

## ? C'est Tout !

**Votre complément est compilé et prêt.**

Il suffit de l'installer avec `INSTALL.bat` et de profiter !

---

**Problème ?** Consultez les logs : `%APPDATA%\OutlookLMStudio\logs.txt`  
**Question ?** Regardez `TROUBLESHOOTING.md`