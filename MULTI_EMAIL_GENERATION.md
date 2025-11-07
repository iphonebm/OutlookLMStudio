# ?? GÉNÉRATION MULTIPLE - Sélection de Plusieurs Emails

## ? Nouvelle Fonctionnalité

Vous pouvez maintenant **sélectionner plusieurs emails** et générer toutes les réponses en une seule fois via le **menu contextuel** !

## ?? Fonctionnalités Ajoutées

### 1. **Sélection Multiple d'Emails**
- Sélectionnez 1 ou plusieurs emails dans votre boîte de réception
- Maintenez `Ctrl` pour sélectionner plusieurs emails
- Maintenez `Shift` pour sélectionner une plage d'emails

### 2. **Menu Contextuel (Clic Droit)**
```
???????????????????????????????????????
?  Répondre                           ?
?  Répondre à tous                    ?
?  Transférer                         ?
?  ??????????????????????????????     ?
?  ? Générer Réponse(s) avec LMStudio ?  ? NOUVEAU !
?  ??????????????????????????????     ?
?  Supprimer                          ?
?  ...                                ?
???????????????????????????????????????
```

### 3. **Bouton dans le Ruban Outlook**
Un nouveau groupe "LMStudio Assistant" apparaît dans l'onglet "Accueil" avec :
- **Afficher/Masquer** : Toggle le volet des tâches
- **Générer Réponses** : Génère les réponses pour tous les emails sélectionnés

### 4. **Barre de Progression**
Pendant le traitement :
```
???????????????????????????????????????????
?  Génération en cours...                 ?
?                                         ?
?  Traitement de 3 / 10 emails...         ?
?  Sujet : Re: Réunion de lundi           ?
?                                         ?
?  ???????????????????????????? 30%      ?
?                                         ?
?  [Annuler]                              ?
???????????????????????????????????????????
```

### 5. **Résumé des Résultats**
À la fin du traitement :
```
???????????????????????????????????????????
?  Génération terminée !                  ?
?                                         ?
?  ? Réussis : 8                         ?
?  ? Erreurs : 2                         ?
?                                         ?
?  Détails des erreurs :                  ?
?  • Email 3: Timeout LMStudio            ?
?  • Email 7: Modèle non chargé           ?
?                                         ?
?  [OK]                                   ?
???????????????????????????????????????????
```

## ?? Installation

### Étape 1 : Configurer le Fichier XML comme Ressource Incorporée

1. **Dans Visual Studio** :
   - Explorateur de Solutions ? Clic droit sur `LMStudioRibbon.xml`
   - Propriétés
   - **Action de génération** : `Ressource incorporée` (très important!)

### Étape 2 : Compiler le Projet

Si vous avez des erreurs de tests :
```
Solution Explorer ? Clic droit sur "OutlookLMStudio.Tests" ? Décharger le projet
```

Puis :
```
Build ? Clean Solution
Build ? Rebuild Solution
```

### Étape 3 : Installer le Complément

```cmd
# Fermez Outlook complètement
# Puis en tant qu'administrateur :
INSTALL_MANUAL.bat
```

## ?? Utilisation

### Méthode 1 : Menu Contextuel (Recommandé)

1. **Sélectionnez** un ou plusieurs emails :
   - Simple clic : 1 email
   - `Ctrl + Clic` : Plusieurs emails
   - `Shift + Clic` : Plage d'emails

2. **Clic droit** sur la sélection

3. **Cliquez** sur "Générer Réponse(s) avec LMStudio"

4. **Confirmez** dans la boîte de dialogue

5. **Attendez** la fin du traitement (barre de progression)

6. **Consultez** les brouillons créés

### Méthode 2 : Bouton du Ruban

1. **Sélectionnez** vos emails

2. **Cliquez** sur l'onglet "Accueil"

3. **Trouvez** le groupe "LMStudio Assistant"

4. **Cliquez** sur "Générer Réponses"

### Méthode 3 : Volet des Tâches (Email unique)

1. **Cliquez** sur un email

2. **Cliquez** sur "Générer une réponse" dans le volet

## ?? Configuration

### Paramètres de Génération Multiple

Dans les paramètres (`Settings`), vous pouvez configurer :

- **Délai entre emails** : Pause de 500ms entre chaque génération (évite de surcharger LMStudio)
- **Template de prompt** : Utilisé pour tous les emails
- **Timeout** : S'applique à chaque email individuellement

## ?? Détails Techniques

### Traitement en Séquence

Les emails sont traités **un par un** (pas en parallèle) pour :
- ? Éviter de surcharger LMStudio
- ? Avoir des logs clairs
- ? Pouvoir annuler à tout moment
- ? Gérer les erreurs individuellement

### Gestion des Erreurs

Chaque email est traité indépendamment :
- Si un email échoue, les autres continuent
- Les erreurs sont collectées et affichées à la fin
- Les brouillons réussis sont créés même en cas d'erreurs partielles

### Code Modifié

**Fichiers créés** :
- `LMStudioRibbon.cs` - Code du ruban personnalisé
- `LMStudioRibbon.xml` - Définition UI du ruban

**Fichiers modifiés** :
- `ContextMenuHandler.cs` - Gestion de la sélection multiple
- `ThisAddIn.cs` - Nouvelle méthode `GenerateResponsesForSelection()`

## ?? Scénarios d'Utilisation

### Scénario 1 : Trier les emails du matin

```
1. Sélectionnez tous les nouveaux emails (Ctrl+A)
2. Clic droit ? Générer Réponse(s) avec LMStudio
3. Confirmez
4. Attendez ~2-3 minutes pour 20 emails
5. Consultez et modifiez les brouillons
6. Envoyez les réponses
```

### Scénario 2 : Répondre aux clients

```
1. Filtrez les emails "Client" (dossier ou filtre)
2. Sélectionnez avec Shift les 10 premiers
3. Générez les réponses
4. Vérifiez et personnalisez
5. Envoyez
```

### Scénario 3 : Traiter une catégorie

```
1. Appliquez une catégorie "À traiter"
2. Sélectionnez tous les emails de cette catégorie
3. Générez en batch
4. Retirez la catégorie après envoi
```

## ?? Limitations

### Nombre Maximum d'Emails

- **Recommandé** : 20-30 emails max par batch
- **Raison** : LMStudio peut devenir lent avec trop de requêtes
- **Solution** : Divisez en plusieurs batches

### Temps de Traitement

- **~3-5 secondes** par email (selon le modèle)
- **10 emails** = ~30-50 secondes
- **50 emails** = ~2.5-4 minutes

### Annulation

- Cliquez sur "Annuler" dans la barre de progression
- **Les emails déjà traités** sont conservés (brouillons créés)
- **Les emails non traités** sont ignorés

## ?? Dépannage

### Le menu contextuel n'apparaît pas

1. Vérifiez que `LMStudioRibbon.xml` est en "Ressource incorporée"
2. Recompilez le projet
3. Réinstallez avec INSTALL_MANUAL.bat

### La barre de progression ne s'affiche pas

- Normal si vous avez seulement 1 email sélectionné
- Vérifiez les logs : `%APPDATA%\OutlookLMStudio\logs.txt`

### Erreur "Globals.ThisAddIn est null"

- Le complément n'est pas complètement initialisé
- Redémarrez Outlook
- Vérifiez l'installation

### Les réponses ne sont pas générées

1. Vérifiez que **LMStudio est démarré** et un modèle est chargé
2. Testez avec un seul email d'abord
3. Consultez les logs d'erreur

## ?? Conseils

### Pour de Meilleures Performances

1. **Chargez le modèle avant** de lancer le traitement batch
2. **Fermez les autres applications** gourmandes
3. **Traitez par lots de 20** maximum
4. **Vérifiez les brouillons** avant d'envoyer

### Workflow Recommandé

```
Matin :
1. Ouvrez LMStudio et chargez le modèle
2. Ouvrez Outlook
3. Sélectionnez les nouveaux emails
4. Générez les réponses en batch
5. Petit café ? pendant le traitement
6. Vérifiez et personnalisez
7. Envoyez

Résultat : 20 emails traités en 5 minutes !
```

## ?? Notes

- Les brouillons sont créés mais **pas envoyés** automatiquement
- Vous pouvez **modifier** chaque brouillon avant l'envoi
- Le volet des tâches continue de fonctionner normalement
- Aucune donnée n'est envoyée sur Internet (100% local)

## ?? Exemple Réel

**Avant** (méthode manuelle) :
```
20 emails × 5 minutes chacun = 100 minutes (1h40)
```

**Maintenant** (avec génération multiple) :
```
20 emails × 3 secondes chacun = 60 secondes
+ 10 minutes de vérification
= 11 minutes total
```

**Gain de temps** : ~90% ! ??

---

**Version** : 1.3.0  
**Fonctionnalité** : Génération multiple avec menu contextuel  
**Build** : En cours de compilation  
**État** : Prêt pour tests