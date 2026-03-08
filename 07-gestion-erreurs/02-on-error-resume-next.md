🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.2. On Error Resume Next

## Introduction à On Error Resume Next

L'instruction `On Error Resume Next` est l'une des méthodes les plus simples pour gérer les erreurs en VBA. Son principe est direct : **quand une erreur survient, ignore-la et continue avec la ligne suivante**. C'est comme dire à VBA "si tu rencontres un problème, ne t'arrête pas, passe à l'instruction d'après".

**Analogie simple :**
Imaginez que vous lisez une liste de courses à voix haute. Avec `On Error Resume Next`, si vous tombez sur un mot que vous ne savez pas prononcer, vous le sautez simplement et continuez avec le mot suivant, au lieu de vous arrêter complètement.

---

## Syntaxe de base

### Comment utiliser On Error Resume Next

```vba
Sub ExempleBase()
    On Error Resume Next        ' Active la gestion d'erreur

    ' Code qui peut générer des erreurs
    Worksheets("FeuilleInexistante").Range("A1").Value = "Test"  ' Cette ligne causera une erreur
    MsgBox "Cette ligne s'exécute quand même !"  ' Mais celle-ci s'exécute

    On Error GoTo 0            ' Désactive la gestion d'erreur
End Sub
```

### Placement de l'instruction

```vba
Sub PlacementCorrect()
    ' Déclarations de variables d'abord
    Dim resultat As Double

    ' Puis activation de la gestion d'erreur
    On Error Resume Next

    ' Code susceptible de générer des erreurs
    resultat = Range("A1").Value / Range("B1").Value
    Range("C1").Value = resultat

    ' Désactivation de la gestion d'erreur à la fin
    On Error GoTo 0
End Sub
```

---

## Comment fonctionne On Error Resume Next

### Comportement normal vs avec gestion d'erreur

#### Sans gestion d'erreur (comportement par défaut)

```vba
Sub SansGestionErreur()
    Range("A1").Value = 10
    Range("B1").Value = 0
    Range("C1").Value = Range("A1").Value / Range("B1").Value  ' Erreur 11: Division par zéro
    Range("D1").Value = "Cette ligne ne s'exécute jamais"      ' Code jamais atteint
    MsgBox "Fin du programme"  ' Message jamais affiché
End Sub
```

**Résultat :** Le programme s'arrête avec un message d'erreur, les lignes suivantes ne s'exécutent pas.

#### Avec On Error Resume Next

```vba
Sub AvecGestionErreur()
    On Error Resume Next

    Range("A1").Value = 10
    Range("B1").Value = 0
    Range("C1").Value = Range("A1").Value / Range("B1").Value  ' Erreur ignorée
    Range("D1").Value = "Cette ligne s'exécute !"             ' Code exécuté
    MsgBox "Fin du programme"  ' Message affiché

    On Error GoTo 0
End Sub
```

**Résultat :** Le programme continue, ignore l'erreur de division par zéro, et exécute toutes les lignes suivantes.

---

## L'objet Err : détecter et analyser les erreurs

### Propriétés importantes de l'objet Err

Même si `On Error Resume Next` ignore les erreurs, VBA les enregistre dans l'objet `Err` que vous pouvez consulter.

#### Err.Number - Le numéro de l'erreur

```vba
Sub ExempleErrNumber()
    On Error Resume Next

    ' Tentative d'accès à une feuille inexistante
    Worksheets("FeuilleInexistante").Range("A1").Value = "Test"

    ' Vérifier s'il y a eu une erreur
    If Err.Number <> 0 Then
        MsgBox "Erreur numéro : " & Err.Number  ' Affichera probablement 9
    Else
        MsgBox "Aucune erreur"
    End If

    On Error GoTo 0
End Sub
```

#### Err.Description - La description de l'erreur

```vba
Sub ExempleErrDescription()
    On Error Resume Next

    ' Division par zéro
    Dim resultat As Double
    resultat = 10 / 0

    ' Afficher les détails de l'erreur
    If Err.Number <> 0 Then
        MsgBox "Erreur " & Err.Number & ": " & Err.Description
        ' Affichera : "Erreur 11: Division par zéro"
    End If

    On Error GoTo 0
End Sub
```

#### Err.Clear - Effacer les informations d'erreur

```vba
Sub ExempleErrClear()
    On Error Resume Next

    ' Première erreur
    Worksheets("FeuilleInexistante").Range("A1").Value = "Test"
    MsgBox "Première erreur : " & Err.Number  ' 9

    ' Effacer l'erreur
    Err.Clear
    MsgBox "Après Clear : " & Err.Number      ' 0

    ' Nouvelle erreur
    Dim resultat As Double
    resultat = 10 / 0
    MsgBox "Nouvelle erreur : " & Err.Number  ' 11

    On Error GoTo 0
End Sub
```

---

## Techniques pratiques avec On Error Resume Next

### 1. Vérifier l'existence d'objets

#### Vérifier si une feuille existe

```vba
Function FeuilleExiste(nomFeuille As String) As Boolean
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = Worksheets(nomFeuille)

    ' Si aucune erreur, la feuille existe
    If Err.Number = 0 Then
        FeuilleExiste = True
    Else
        FeuilleExiste = False
    End If

    Err.Clear
    On Error GoTo 0
End Function

Sub UtiliserVerificationFeuille()
    If FeuilleExiste("Données") Then
        MsgBox "La feuille Données existe"
        Worksheets("Données").Range("A1").Value = "OK"
    Else
        MsgBox "La feuille Données n'existe pas"
    End If
End Sub
```

#### Vérifier si un classeur est ouvert

```vba
Function ClasseurOuvert(nomClasseur As String) As Boolean
    On Error Resume Next

    Dim wb As Workbook
    Set wb = Workbooks(nomClasseur)

    ClasseurOuvert = (Err.Number = 0)

    Err.Clear
    On Error GoTo 0
End Function

Sub UtiliserVerificationClasseur()
    If ClasseurOuvert("Données.xlsx") Then
        MsgBox "Le classeur est déjà ouvert"
    Else
        MsgBox "Le classeur n'est pas ouvert"
        ' Tentative d'ouverture
        Workbooks.Open "C:\Données.xlsx"
    End If
End Sub
```

### 2. Tentatives avec alternatives

#### Ouvrir un fichier avec chemin de secours

```vba
Sub OuvrirFichierAvecAlternative()
    On Error Resume Next

    ' Tentative d'ouverture du fichier principal
    Workbooks.Open "C:\Données\Principal.xlsx"

    If Err.Number <> 0 Then
        Err.Clear
        ' Tentative avec le fichier de sauvegarde
        Workbooks.Open "C:\Sauvegarde\Principal.xlsx"

        If Err.Number <> 0 Then
            MsgBox "Impossible d'ouvrir le fichier principal ou de sauvegarde"
            On Error GoTo 0
            Exit Sub
        Else
            MsgBox "Fichier de sauvegarde ouvert"
        End If
    Else
        MsgBox "Fichier principal ouvert"
    End If

    On Error GoTo 0
End Sub
```

### 3. Calculs avec gestion d'erreurs

#### Division sécurisée

```vba
Function DivisionSecurisee(dividende As Double, diviseur As Double) As Variant
    On Error Resume Next

    Dim resultat As Double
    resultat = dividende / diviseur

    If Err.Number = 0 Then
        DivisionSecurisee = resultat
    Else
        DivisionSecurisee = "Erreur: " & Err.Description
    End If

    Err.Clear
    On Error GoTo 0
End Function

Sub UtiliserDivisionSecurisee()
    ' Tests avec différentes valeurs
    Range("A1").Value = DivisionSecurisee(10, 2)    ' 5
    Range("A2").Value = DivisionSecurisee(10, 0)    ' "Erreur: Division par zéro"
    Range("A3").Value = DivisionSecurisee(15, 3)    ' 5
End Sub
```

---

## Gestion avancée avec On Error Resume Next

### 1. Combinaison avec des boucles

#### Traitement de données avec erreurs possibles

```vba
Sub TraiterDonneesAvecErreurs()
    On Error Resume Next

    Dim i As Integer
    Dim valeur As Variant
    Dim resultat As Double

    ' Traiter les données de A1 à A10
    For i = 1 To 10
        valeur = Cells(i, 1).Value

        ' Tentative de conversion en nombre et calcul
        resultat = valeur * 2

        If Err.Number = 0 Then
            ' Succès : écrire le résultat
            Cells(i, 2).Value = resultat
            Cells(i, 3).Value = "OK"
        Else
            ' Erreur : marquer comme problématique
            Cells(i, 2).Value = "N/A"
            Cells(i, 3).Value = "Erreur: " & Err.Description
            Err.Clear  ' Important : effacer l'erreur pour la prochaine itération
        End If
    Next i

    On Error GoTo 0
    MsgBox "Traitement terminé"
End Sub
```

### 2. Création robuste d'objets

#### Créer des feuilles avec noms uniques

```vba
Sub CreerFeuilleUnique(nomBase As String)
    On Error Resume Next

    Dim nomFeuille As String
    Dim compteur As Integer

    nomFeuille = nomBase
    compteur = 1

    ' Tenter de créer la feuille
    Dim nouvelleFeuille As Worksheet
    Do
        Set nouvelleFeuille = Worksheets.Add
        nouvelleFeuille.Name = nomFeuille

        If Err.Number = 0 Then
            ' Succès
            MsgBox "Feuille créée : " & nomFeuille
            Exit Do
        Else
            ' Le nom existe déjà : supprimer la feuille orpheline
            Application.DisplayAlerts = False
            nouvelleFeuille.Delete
            Application.DisplayAlerts = True
            Err.Clear
            compteur = compteur + 1
            nomFeuille = nomBase & "_" & compteur
        End If
    Loop While compteur < 100  ' Sécurité pour éviter une boucle infinie

    On Error GoTo 0
End Sub

Sub UtiliserCreationFeuille()
    CreerFeuilleUnique "Données"  ' Crée "Données", puis "Données_2", etc.
End Sub
```

---

## Bonnes pratiques avec On Error Resume Next

### 1. Toujours désactiver à la fin

```vba
Sub BonnePratiqueDesactivation()
    On Error Resume Next

    ' Votre code avec gestion d'erreur
    Range("Test").Value = "Valeur"

    ' IMPORTANT : Toujours désactiver
    On Error GoTo 0

    ' Le reste du code fonctionne normalement
    MsgBox "Fin"
End Sub
```

### 2. Vérifier les erreurs régulièrement

```vba
Sub BonnePratiqueVerification()
    On Error Resume Next

    ' Opération 1
    Workbooks.Open "Fichier1.xlsx"
    If Err.Number <> 0 Then
        MsgBox "Impossible d'ouvrir Fichier1: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    ' Opération 2
    Worksheets("Données").Range("A1").Value = "Test"
    If Err.Number <> 0 Then
        MsgBox "Problème avec la feuille Données: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo 0
    MsgBox "Toutes les opérations réussies"
End Sub
```

### 3. Utiliser des zones limitées

```vba
Sub BonnePratiqueZonesLimitees()
    ' Code normal
    Range("A1").Value = "Début"

    ' Zone avec gestion d'erreur limitée
    On Error Resume Next
    Worksheets("PeutPasExister").Range("A1").Value = "Test"
    If Err.Number <> 0 Then
        MsgBox "Feuille introuvable"
        Err.Clear
    End If
    On Error GoTo 0

    ' Retour au code normal
    Range("A2").Value = "Fin"
End Sub
```

---

## Avantages et inconvénients

### Avantages de On Error Resume Next

1. **Simplicité** : Très facile à comprendre et utiliser
2. **Continuité** : Le programme ne s'arrête pas brutalement
3. **Flexibilité** : Permet de tester l'existence d'objets facilement
4. **Contrôle** : Vous décidez comment réagir à chaque erreur

### Inconvénients et pièges

1. **Masque les erreurs** : Peut cacher des problèmes importants
2. **Difficile à déboguer** : Les erreurs passent inaperçues
3. **Performance** : Chaque erreur prend du temps même si ignorée
4. **Risque d'effet domino** : Une erreur peut en causer d'autres

### Quand utiliser On Error Resume Next

#### ✅ Utilisez-le pour :
- Vérifier l'existence d'objets (feuilles, classeurs, fichiers)
- Tentatives d'opérations avec alternatives
- Nettoyage de code (supprimer des objets qui peuvent ne pas exister)
- Tests de fonctionnalités optionnelles

#### ❌ Évitez-le pour :
- Ignorer toutes les erreurs sans distinction
- Code en production sans vérification des erreurs
- Calculs critiques où l'exactitude est importante
- Apprentissage (masque les erreurs que vous devriez voir)

---

## Exemples pratiques d'utilisation

### 1. Nettoyage d'une feuille

```vba
Sub NettoyerFeuille()
    On Error Resume Next

    ' Supprimer différents éléments qui peuvent ne pas exister
    Range("A1:Z100").ClearContents
    ActiveSheet.Shapes.SelectAll     ' Sélectionner toutes les formes
    Selection.Delete                 ' Supprimer la sélection
    ActiveSheet.ChartObjects.Delete  ' Supprimer tous les graphiques

    ' Aucune erreur même si ces éléments n'existent pas
    On Error GoTo 0

    MsgBox "Nettoyage terminé"
End Sub
```

### 2. Sauvegarde avec nom automatique

```vba
Sub SauvegardeAutomatique()
    On Error Resume Next

    Dim nomFichier As String
    Dim compteur As Integer

    compteur = 1

    Do
        nomFichier = "Sauvegarde_" & Format(Date, "yyyy-mm-dd") & "_" & compteur & ".xlsx"

        ' Tentative de sauvegarde
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & nomFichier

        If Err.Number = 0 Then
            MsgBox "Sauvegardé sous : " & nomFichier
            Exit Do
        Else
            Err.Clear
            compteur = compteur + 1
        End If
    Loop While compteur <= 100

    On Error GoTo 0
End Sub
```

---

## Récapitulatif

### Points clés à retenir

1. **On Error Resume Next** ignore les erreurs et continue l'exécution
2. **L'objet Err** contient les informations sur les erreurs (Number, Description)
3. **Err.Clear** efface les informations d'erreur
4. **On Error GoTo 0** désactive la gestion d'erreur
5. **Toujours vérifier Err.Number** après les opérations critiques
6. **Désactiver la gestion** dès que la zone risquée est passée

### Modèle type d'utilisation

```vba
Sub ModeleType()
    ' Code normal

    On Error Resume Next
    ' Code avec risque d'erreur
    If Err.Number <> 0 Then
        ' Traitement de l'erreur
        Err.Clear
    End If
    On Error GoTo 0

    ' Retour au code normal
End Sub
```

### Conseil pour débuter

Commencez par utiliser `On Error Resume Next` pour des cas simples comme vérifier l'existence d'objets. Une fois à l'aise, vous pourrez passer à des techniques plus avancées comme `On Error GoTo` pour une gestion plus sophistiquée.

Dans la section suivante, nous découvrirons `On Error GoTo`, qui permet une gestion d'erreurs plus structurée et puissante.

⏭️
