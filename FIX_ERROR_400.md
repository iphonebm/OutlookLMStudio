# ?? Résolution Erreur 400 - LMStudio

## ? Erreur Rencontrée

```
Le serveur distant a retourné une erreur : (400) Demande incorrecte.
```

## ? Problème Résolu !

### Ce qui a été corrigé :

1. **Format de la requête JSON** - Maintenant compatible avec l'API OpenAI utilisée par LMStudio
2. **Headers HTTP** - Ajout de `Accept: application/json`
3. **Logging détaillé** - Pour voir exactement ce qui est envoyé et reçu
4. **Gestion d'erreurs améliorée** - Messages d'erreur plus clairs avec détails du serveur

### Nouveau format de requête :

```json
{
  "model": "nom-du-modele",
  "prompt": "votre prompt",
  "max_tokens": 2000,
  "temperature": 0.7,
  "top_p": 1.0,
  "n": 1,
  "stream": false,
  "stop": ["</response>"]
}
```

## ?? Que Faire Maintenant

### 1. Rebuild le Projet

```
Dans Visual Studio :
- Build ? Rebuild Solution
- Attendez que ça finisse
```

### 2. Réinstallez le Complément

```
- Clic droit sur INSTALL.bat
- Exécuter en tant qu'administrateur
```

### 3. Configurez LMStudio Correctement

**IMPORTANT** : Vérifiez ces points dans LMStudio :

#### A. Démarrer le Serveur Local

1. Ouvrez LMStudio
2. Allez dans l'onglet **"Local Server"** (icône serveur)
3. Sélectionnez un modèle dans la liste déroulante
4. Cliquez sur **"Start Server"**
5. Vérifiez que l'URL affichée est : `http://localhost:1234`

#### B. Tester le Serveur

Dans votre navigateur, allez à :
```
http://localhost:1234/v1/models
```

Vous devriez voir quelque chose comme :
```json
{
  "object": "list",
  "data": [
    {
      "id": "nom-de-votre-modele",
      "object": "model",
      ...
    }
  ]
}
```

Si vous voyez ça, le serveur fonctionne ! ?

#### C. Notez le Nom du Modèle

Dans la réponse JSON ci-dessus, notez l'ID du modèle (ex: `llama-2-7b`, `mistral-7b`, etc.)

### 4. Configurez le Complément

Dans Outlook :

1. Ouvrez le volet "LMStudio Assistant"
2. Cliquez sur **"Paramètres"**
3. Vérifiez/Modifiez :
   - **URL** : `http://localhost:1234` (par défaut)
   - **Nom du modèle** : Le nom que vous avez noté (ex: `llama-2-7b`)
   - **Max Tokens** : `2000` (ou plus selon votre RAM)
   - **Température** : `0.7` (entre 0 et 1)
4. Cliquez **"Enregistrer"**

### 5. Testez !

1. Sélectionnez un email dans Outlook
2. Cliquez sur **"Générer une réponse"**
3. Attendez (la première fois peut prendre 30 secondes)
4. Une réponse devrait s'afficher ! ??

## ?? Vérification des Logs

Les nouveaux logs sont BEAUCOUP plus détaillés :

```
%APPDATA%\OutlookLMStudio\logs.txt
```

Vous verrez maintenant :
- ? Le JSON exact envoyé à LMStudio
- ? L'URL de l'endpoint utilisé
- ? La réponse brute reçue
- ? Les détails de toute erreur serveur

## ?? Si Ça Ne Marche Toujours Pas

### Erreur Possible 1 : "Connection refused"

**Cause** : Le serveur LMStudio n'est pas démarré

**Solution** :
1. Lancez LMStudio
2. Onglet "Local Server"
3. **START SERVER** (le bouton doit être en vert)

### Erreur Possible 2 : "404 Not Found"

**Cause** : Mauvaise URL ou endpoint

**Solution** :
1. Vérifiez que l'URL est : `http://localhost:1234`
2. PAS `https://`
3. PAS de `/` à la fin
4. Le port par défaut est `1234`

### Erreur Possible 3 : "No model loaded"

**Cause** : Aucun modèle chargé dans LMStudio

**Solution** :
1. Dans LMStudio, onglet "Local Server"
2. Sélectionnez un modèle dans la liste déroulante
3. Le modèle se charge (peut prendre 1-2 minutes)
4. Une fois chargé, START SERVER

### Erreur Possible 4 : "Timeout"

**Cause** : Le modèle met trop de temps à répondre

**Solution** :
1. Augmentez le timeout dans les paramètres (60 secondes)
2. Réduisez max_tokens (500 au lieu de 2000)
3. Utilisez un modèle plus petit (7B au lieu de 13B)

## ?? Configuration Recommandée

Pour un bon équilibre performance/qualité :

```
URL API          : http://localhost:1234
Timeout          : 60 secondes
Température      : 0.7 (créatif mais cohérent)
Max Tokens       : 1000 (réponses moyennes)
Stop Sequences   : </response>
Nom du Modèle    : [Le nom exact de votre modèle]
```

## ?? Astuces

### Modèles Recommandés pour Emails

- **Mistral 7B** : Rapide, excellente qualité française
- **Llama 2 7B** : Bon compromis vitesse/qualité
- **Phi-2** : Très rapide, modèle compact

### Ajuster la Température

- **0.1-0.3** : Réponses très formelles et prévisibles
- **0.5-0.7** : Équilibre (RECOMMANDÉ)
- **0.8-1.0** : Réponses créatives et variées

### Optimiser les Performances

1. **Première génération lente ?** Normal ! Le modèle se charge en mémoire
2. **Générations suivantes** : Beaucoup plus rapides
3. **Gardez LMStudio ouvert** : Pas besoin de recharger le modèle
4. **Fermez les autres applications** : Libère de la RAM pour le modèle

## ? Checklist Finale

Avant de tester, vérifiez :

- [ ] LMStudio est **lancé**
- [ ] Un modèle est **sélectionné**
- [ ] Le serveur est **démarré** (bouton vert)
- [ ] `http://localhost:1234/v1/models` fonctionne dans le navigateur
- [ ] Le complément est **réinstallé** avec la nouvelle version
- [ ] Le **nom du modèle** est correct dans les paramètres
- [ ] Les **logs** sont vides d'erreurs

Si tout est ?, ça devrait marcher !

## ?? Ça Marche ?

Félicitations ! Vous pouvez maintenant :
- Générer des réponses automatiquement
- Personnaliser le prompt dans le volet
- Ajuster les paramètres selon vos besoins
- Profiter de l'IA locale pour vos emails !

---

**Besoin d'aide ?** Consultez les logs détaillés dans `%APPDATA%\OutlookLMStudio\logs.txt`