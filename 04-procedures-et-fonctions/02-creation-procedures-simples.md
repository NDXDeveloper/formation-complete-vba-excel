🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.2 Création de procédures simples

## Introduction

Maintenant que vous comprenez la différence entre Sub et Function, il est temps d'apprendre à créer vos propres procédures. Cette section vous guidera pas à pas dans la création de procédures simples mais utiles.

## Anatomie d'une procédure

### Structure de base

Chaque procédure VBA suit une structure précise et obligatoire :

```vba
Sub NomDeLaProcedure()
    ' Vos instructions ici
    ' Commentaires pour expliquer le code
    ' Autres instructions...
End Sub
```

### Les éléments essentiels

1. **Le mot-clé `Sub`** : Indique le début d'une procédure
2. **Le nom de la procédure** : Identifiant unique que vous choisissez
3. **Les parenthèses `()`** : Même vides, elles sont obligatoires
4. **Le corps de la procédure** : Vos instructions entre Sub et End Sub
5. **`End Sub`** : Marque la fin de la procédure

## Règles de nommage des procédures

### Règles obligatoires
- Le nom doit commencer par une **lettre** (pas de chiffre ou symbole)
- Pas d'**espaces** dans le nom (utilisez des underscores ou la notation CamelCase)
- Pas de **caractères spéciaux** (@, #, $, %, etc.)
- Maximum **255 caractères** (mais soyez raisonnable !)
- Ne pas utiliser de **mots réservés** VBA (Sub, Function, If, etc.)

### Bonnes pratiques de nommage
```vba
' ✅ Bons exemples
Sub AfficherMessage()
Sub Calculer_Total()
Sub FormaterTableau()
Sub SauvegarderDonnees()

' ❌ Mauvais exemples
Sub 123Test()        ' Commence par un chiffre
Sub Mon Calcul()     ' Contient des espaces
Sub Sub()           ' Mot réservé VBA
Sub @Fonction()     ' Caractère spécial
```

### Conventions recommandées
- **Noms descriptifs** : `FormaterCellule` plutôt que `FC`
- **Commencer par un verbe** : `Afficher`, `Calculer`, `Supprimer`
- **Utiliser la notation CamelCase** : `AfficherMessageBienvenue`
- **Éviter les abréviations obscures** : `Supprimer` plutôt que `Suppr`

## Création de votre première procédure

### Étape 1 : Ouvrir l'éditeur VBA
1. Dans Excel, appuyez sur **Alt + F11**
2. L'éditeur VBA s'ouvre

### Étape 2 : Insérer un module
1. Clic droit sur votre projet dans l'explorateur
2. **Insertion** > **Module**
3. Un nouveau module apparaît

### Étape 3 : Écrire votre première procédure

```vba
Sub MonPremiereProcedure()
    MsgBox "Félicitations ! Vous venez de créer votre première procédure !"
End Sub
```

### Étape 4 : Exécuter la procédure
1. Placez le curseur à l'intérieur de la procédure
2. Appuyez sur **F5** ou cliquez sur le bouton **Exécuter**
3. Votre message apparaît !

## Exemples de procédures simples

### Exemple 1 : Procédure d'affichage simple
```vba
Sub DireBonjour()
    MsgBox "Bonjour ! J'espère que vous passez une excellente journée."
End Sub
```
**Ce que fait cette procédure :** Affiche un message de bienvenue à l'utilisateur.

### Exemple 2 : Procédure de formatage
```vba
Sub FormaterTitre()
    Range("A1").Value = "RAPPORT MENSUEL"
    Range("A1").Font.Bold = True
    Range("A1").Font.Size = 16
    Range("A1").Font.Color = RGB(0, 0, 255)  ' Bleu
    Range("A1").HorizontalAlignment = xlCenter
End Sub
```
**Ce que fait cette procédure :** Crée et formate un titre en cellule A1.

### Exemple 3 : Procédure de nettoyage
```vba
Sub EffacerZoneDetravail()
    Range("A1:J20").ClearContents
    Range("A1:J20").ClearFormats
    MsgBox "Zone de travail nettoyée !"
End Sub
```
**Ce que fait cette procédure :** Efface le contenu et le formatage d'une zone spécifique.

### Exemple 4 : Procédure d'information
```vba
Sub AfficherInformationsFichier()
    Dim nomFichier As String
    Dim cheminFichier As String

    nomFichier = ActiveWorkbook.Name
    cheminFichier = ActiveWorkbook.Path

    MsgBox "Fichier actuel : " & nomFichier & vbNewLine & _
           "Emplacement : " & cheminFichier
End Sub
```
**Ce que fait cette procédure :** Affiche des informations sur le fichier Excel actuel.

## Procédures avec plusieurs instructions

Les procédures peuvent contenir plusieurs instructions qui s'exécutent dans l'ordre :

```vba
Sub CreerRapportSimple()
    ' Étape 1 : Effacer la feuille
    Cells.ClearContents

    ' Étape 2 : Créer l'en-tête
    Range("A1").Value = "RAPPORT DE VENTES"
    Range("A1").Font.Bold = True
    Range("A1").Font.Size = 14

    ' Étape 3 : Créer les colonnes
    Range("A3").Value = "Produit"
    Range("B3").Value = "Quantité"
    Range("C3").Value = "Prix unitaire"
    Range("D3").Value = "Total"

    ' Étape 4 : Formater les en-têtes de colonnes
    Range("A3:D3").Font.Bold = True
    Range("A3:D3").Interior.Color = RGB(200, 200, 200)  ' Gris clair

    ' Étape 5 : Ajuster la largeur des colonnes
    Range("A:D").AutoFit

    ' Étape 6 : Confirmer la création
    MsgBox "Rapport créé avec succès !"
End Sub
```

## Utilisation des commentaires

Les commentaires sont essentiels pour expliquer ce que fait votre code :

```vba
Sub ExempleAvecCommentaires()
    ' Ceci est un commentaire sur une ligne complète

    MsgBox "Hello World"  ' Commentaire en fin de ligne

    ' Les commentaires n'affectent pas l'exécution
    ' Ils sont uniquement pour la documentation

    Range("A1").Value = "Test"  ' Écrit "Test" en A1
End Sub
```

### Types de commentaires
- **Ligne complète** : `' Ceci est un commentaire`
- **Fin de ligne** : `Code instruction  ' Commentaire`
- **Bloc de commentaires** : Plusieurs lignes consécutives avec `'`

## Erreurs courantes à éviter

### 1. Oublier End Sub
```vba
' ❌ Incorrect
Sub MaProcedure()
    MsgBox "Test"
' Manque End Sub !

' ✅ Correct
Sub MaProcedure()
    MsgBox "Test"
End Sub
```

### 2. Erreur de syntaxe dans le nom
```vba
' ❌ Incorrect
Sub Ma Procédure()  ' Espace et accent
    MsgBox "Test"
End Sub

' ✅ Correct
Sub MaProcedure()
    MsgBox "Test"
End Sub
```

### 3. Oublier les parenthèses
```vba
' ❌ Incorrect
Sub MaProcedure
    MsgBox "Test"
End Sub

' ✅ Correct
Sub MaProcedure()
    MsgBox "Test"
End Sub
```

## Comment organiser vos procédures

### Dans un même module
```vba
' Procédures liées au formatage
Sub FormaterTitre()
    ' Code pour formater le titre
End Sub

Sub FormaterTableau()
    ' Code pour formater le tableau
End Sub

' Procédures liées aux données
Sub ImporterDonnees()
    ' Code pour importer
End Sub

Sub ExporterDonnees()
    ' Code pour exporter
End Sub
```

### Utilisation de séparateurs visuels
```vba
'=====================================
' PROCÉDURES DE FORMATAGE
'=====================================

Sub FormaterTitre()
    ' Code ici
End Sub

'=====================================
' PROCÉDURES DE DONNÉES
'=====================================

Sub ImporterDonnees()
    ' Code ici
End Sub
```

## Tester vos procédures

### Méthodes d'exécution
1. **F5** : Avec le curseur dans la procédure
2. **Menu Exécution > Exécuter Sub/UserForm**
3. **Bouton Exécuter** dans la barre d'outils
4. **Depuis Excel** : Alt+F8, sélectionner la macro

### Vérification du bon fonctionnement
- Testez chaque procédure individuellement
- Vérifiez que les résultats correspondent à vos attentes
- Observez s'il y a des messages d'erreur
- Assurez-vous que les modifications attendues sont appliquées

## Points clés à retenir

1. **Structure obligatoire** : `Sub NomProcedure()` ... `End Sub`
2. **Noms descriptifs** : Choisissez des noms qui expliquent ce que fait la procédure
3. **Une tâche par procédure** : Chaque procédure doit avoir un objectif clair
4. **Commentaires** : Documentez votre code pour vous et les autres
5. **Test systématique** : Vérifiez toujours que vos procédures fonctionnent
6. **Organisation** : Groupez les procédures similaires ensemble

## Prochaines étapes

Maintenant que vous savez créer des procédures simples, vous êtes prêt à apprendre à les rendre plus flexibles avec des paramètres. Cette compétence vous permettra de créer des procédures réutilisables qui s'adaptent à différentes situations.

⏭️
