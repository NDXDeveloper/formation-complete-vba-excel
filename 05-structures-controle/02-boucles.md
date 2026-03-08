🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.2 Boucles

## Introduction

Les **boucles** permettent de répéter automatiquement un bloc d'instructions plusieurs fois. Sans les boucles, vous devriez écrire le même code des dizaines ou des centaines de fois ! Elles sont essentielles pour traiter de grandes quantités de données ou automatiser des tâches répétitives.

### Analogie de la machine à laver

Une machine à laver effectue plusieurs cycles :
- **Remplissage** : répète "ajouter de l'eau" jusqu'à ce que le niveau soit atteint
- **Lavage** : répète "tourner à droite, tourner à gauche" pendant 30 minutes
- **Rinçage** : répète le cycle "remplir d'eau, vider" 3 fois
- **Essorage** : répète "tourner très vite" pendant 10 minutes

Chaque étape est une boucle qui répète une action selon différents critères !

### Pourquoi utiliser des boucles ?

**Sans boucle** (répétitif et inefficace) :
```vba
Range("A1").Value = "Ligne 1"  
Range("A2").Value = "Ligne 2"  
Range("A3").Value = "Ligne 3"  
Range("A4").Value = "Ligne 4"  
Range("A5").Value = "Ligne 5"  
' ... Et si on voulait 1000 lignes ?
```

**Avec une boucle** (élégant et puissant) :
```vba
For i = 1 To 5
    Range("A" & i).Value = "Ligne " & i
Next i
```

## Types de boucles en VBA

1. **For...Next** : Répète un nombre défini de fois
2. **For Each...Next** : Parcourt tous les éléments d'une collection
3. **Do...Loop** : Répète tant qu'une condition est vraie/fausse
4. **While...Wend** : Version simplifiée de Do...Loop

## 5.2.1 For...Next

### Structure de base

```vba
For compteur = début To fin
    ' Instructions à répéter
Next compteur
```

Le **compteur** est une variable qui prend successivement toutes les valeurs de **début** à **fin**.

### Exemple simple

```vba
Sub CompterJusqueDix()
    Dim i As Integer

    For i = 1 To 10
        MsgBox "Nous en sommes à : " & i
    Next i

    MsgBox "Fini !"
End Sub
```

### Remplir des cellules avec For...Next

```vba
Sub RemplirColonne()
    Dim ligne As Integer

    For ligne = 1 To 20
        Range("A" & ligne).Value = "Produit " & ligne
        Range("B" & ligne).Value = ligne * 10  ' Prix
    Next ligne
End Sub
```

### Utiliser Step pour modifier l'incrément

```vba
Sub ExemplesAvecStep()
    Dim i As Integer

    ' Compter de 2 en 2
    For i = 2 To 20 Step 2
        Range("A" & i / 2).Value = i  ' 2, 4, 6, 8, 10...
    Next i

    ' Compter à rebours
    For i = 10 To 1 Step -1
        Range("B" & (11 - i)).Value = i  ' 10, 9, 8, 7...
    Next i

    ' Compter de 5 en 5
    For i = 5 To 50 Step 5
        Range("C" & (i / 5)).Value = i  ' 5, 10, 15, 20...
    Next i
End Sub
```

### Boucles imbriquées (une boucle dans une autre)

```vba
Sub CreerTableauMultiplication()
    Dim ligne As Integer
    Dim colonne As Integer

    ' Créer un tableau de multiplication 10x10
    For ligne = 1 To 10
        For colonne = 1 To 10
            Cells(ligne, colonne).Value = ligne * colonne
        Next colonne
    Next ligne

    ' Formater le tableau
    Range("A1:J10").Font.Size = 8
    Range("A1:J10").AutoFit
End Sub
```

### Utilisation pratique : Traiter des données

```vba
Sub CalculerTotaux()
    Dim i As Integer
    Dim total As Double

    ' Supposons que nous avons des prix en colonne B, lignes 2 à 11
    For i = 2 To 11
        If IsNumeric(Range("B" & i).Value) Then
            total = total + Range("B" & i).Value
        End If
    Next i

    ' Afficher le total en B12
    Range("B12").Value = total
    Range("B12").Font.Bold = True
    Range("A12").Value = "TOTAL :"
End Sub
```

## 5.2.2 For Each...Next

### Concept et utilité

`For Each...Next` parcourt automatiquement **tous les éléments** d'une collection (feuilles, cellules, fichiers...) sans que vous ayez besoin de connaître le nombre d'éléments.

### Structure de base

```vba
For Each élément In collection
    ' Instructions pour chaque élément
Next élément
```

### Parcourir une plage de cellules

```vba
Sub FormaterCellulesNegatives()
    Dim cellule As Range

    ' Parcourir toutes les cellules de A1 à D10
    For Each cellule In Range("A1:D10")
        If IsNumeric(cellule.Value) Then
            If cellule.Value < 0 Then
                cellule.Font.Color = RGB(255, 0, 0)  ' Rouge
                cellule.Font.Bold = True
            End If
        End If
    Next cellule
End Sub
```

### Parcourir toutes les feuilles d'un classeur

```vba
Sub ListerToutesLesFeuilles()
    Dim feuille As Worksheet
    Dim i As Integer

    i = 1
    For Each feuille In ThisWorkbook.Worksheets
        Range("A" & i).Value = feuille.Name
        i = i + 1
    Next feuille
End Sub
```

### Traiter la sélection actuelle

```vba
Sub TraiterSelection()
    Dim cellule As Range

    ' Vérifier que la sélection est bien une plage de cellules
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner des cellules d'abord"
        Exit Sub
    End If

    ' Traiter chaque cellule sélectionnée
    For Each cellule In Selection
        If cellule.Value = "" Then
            cellule.Value = "VIDE"
            cellule.Font.Color = RGB(128, 128, 128)  ' Gris
        End If
    Next cellule
End Sub
```

### Parcourir des objets graphiques

```vba
Sub SupprimerTousLesGraphiques()
    Dim forme As Shape

    ' Parcourir toutes les formes de la feuille active
    For Each forme In ActiveSheet.Shapes
        If forme.Type = msoChart Then  ' Si c'est un graphique
            forme.Delete
        End If
    Next forme

    MsgBox "Tous les graphiques ont été supprimés"
End Sub
```

### Avantages de For Each

```vba
Sub ComparaisonForEtForEach()
    Dim i As Integer
    Dim cellule As Range

    ' ❌ Avec For classique (plus complexe)
    For i = 1 To Selection.Cells.Count
        Selection.Cells(i).Font.Bold = True
    Next i

    ' ✅ Avec For Each (plus simple et lisible)
    For Each cellule In Selection
        cellule.Font.Bold = True
    Next cellule
End Sub
```

## 5.2.3 Do...Loop (While/Until)

### Concept

`Do...Loop` répète un bloc d'instructions **tant qu'une condition est vraie** ou **jusqu'à ce qu'une condition devienne vraie**. Contrairement à `For...Next`, vous ne savez pas à l'avance combien de fois la boucle va s'exécuter.

### Do While (tant que)

```vba
Do While condition
    ' Instructions à répéter
Loop
```

### Exemple : Demander une saisie valide

```vba
Sub DemanderAgeValide()
    Dim age As String
    Dim ageNumerique As Integer

    Do While True  ' Boucle infinie contrôlée
        age = InputBox("Entrez votre âge (entre 0 et 120) :")

        ' Vérifier si c'est un nombre
        If IsNumeric(age) Then
            ageNumerique = CInt(age)

            ' Vérifier si l'âge est dans la plage valide
            If ageNumerique >= 0 And ageNumerique <= 120 Then
                MsgBox "Âge valide : " & ageNumerique & " ans"
                Exit Do  ' Sortir de la boucle
            End If
        End If

        MsgBox "Âge invalide ! Veuillez recommencer."
    Loop
End Sub
```

### Do Until (jusqu'à ce que)

```vba
Sub ChercherCelluleVide()
    Dim ligne As Integer
    ligne = 1

    ' Chercher la première cellule vide en colonne A
    Do Until Range("A" & ligne).Value = ""
        ligne = ligne + 1
    Loop

    MsgBox "Première cellule vide trouvée en A" & ligne
    Range("A" & ligne).Select
End Sub
```

### Condition à la fin (Do...Loop While/Until)

```vba
Sub ExempleConditionFin()
    Dim nombre As Integer
    Dim tentatives As Integer

    Do
        tentatives = tentatives + 1
        nombre = Int(Rnd() * 10) + 1  ' Nombre aléatoire 1-10
        MsgBox "Tentative " & tentatives & " : " & nombre
    Loop Until nombre = 7  ' Répéter jusqu'à obtenir 7

    MsgBox "7 trouvé après " & tentatives & " tentatives !"
End Sub
```

### Exemple pratique : Nettoyer des données

```vba
Sub SupprimerLignesVides()
    Dim ligne As Integer
    ligne = 1

    Do While ligne <= ActiveSheet.UsedRange.Rows.Count
        ' Si toute la ligne est vide
        If Application.WorksheetFunction.CountA(Rows(ligne)) = 0 Then
            Rows(ligne).Delete
            ' Ne pas incrémenter ligne car les lignes suivantes remontent
        Else
            ligne = ligne + 1
        End If
    Loop

    MsgBox "Lignes vides supprimées"
End Sub
```

### Attention aux boucles infinies !

```vba
' ❌ DANGER - Boucle infinie !
Sub BoucleInfinie()
    Dim x As Integer
    x = 1

    Do While x > 0
        x = x + 1  ' x ne sera jamais <= 0 !
        ' Cette boucle ne s'arrêtera jamais !
    Loop
End Sub

' ✅ Solution - Ajouter une condition de sortie
Sub BoucleSecurisee()
    Dim x As Integer
    Dim compteur As Integer
    x = 1

    Do While x > 0 And compteur < 1000  ' Sécurité
        x = x + 1
        compteur = compteur + 1
    Loop

    If compteur >= 1000 Then
        MsgBox "Boucle interrompue pour éviter l'infini"
    End If
End Sub
```

## 5.2.4 While...Wend

### Structure simple

`While...Wend` est une version simplifiée de `Do While...Loop`. Elle est moins flexible mais plus concise pour des cas simples.

```vba
While condition
    ' Instructions à répéter
Wend
```

### Exemple basique

```vba
Sub ExempleWhileWend()
    Dim i As Integer
    i = 1

    While i <= 10
        Range("A" & i).Value = "Ligne " & i
        i = i + 1
    Wend
End Sub
```

### Comparaison Do While vs While Wend

```vba
Sub ComparaisonBoucles()
    Dim i As Integer

    ' ✅ Avec Do While (recommandé - plus flexible)
    i = 1
    Do While i <= 5
        Range("A" & i).Value = "Do While " & i
        i = i + 1
    Loop

    ' ✅ Avec While Wend (plus simple, moins flexible)
    i = 1
    While i <= 5
        Range("B" & i).Value = "While Wend " & i
        i = i + 1
    Wend
End Sub
```

### Limitations de While...Wend

```vba
' ❌ Impossible avec While...Wend
While condition1
    ' Pas d'Exit While possible !
    ' Pas de Until possible !
Wend

' ✅ Possible avec Do...Loop
Do While condition1
    If condition2 Then Exit Do  ' Sortie possible
    ' Plus de flexibilité
Loop
```

## Contrôle de flux dans les boucles

### Exit (sortir de la boucle)

```vba
Sub ExempleExit()
    Dim i As Integer

    For i = 1 To 100
        If Range("A" & i).Value = "STOP" Then
            MsgBox "Mot STOP trouvé à la ligne " & i
            Exit For  ' Sortir de la boucle For
        End If

        Range("A" & i).Value = "Ligne " & i
    Next i
End Sub

Sub ExempleExitDo()
    Dim compteur As Integer

    Do
        compteur = compteur + 1

        If compteur > 10 Then
            Exit Do  ' Sortir de la boucle Do
        End If

        MsgBox "Compteur : " & compteur
    Loop
End Sub
```

### Continue (passer à l'itération suivante)

VBA n'a pas de mot-clé "Continue" comme d'autres langages, mais on peut simuler ce comportement :

```vba
Sub SimulerContinue()
    Dim i As Integer

    For i = 1 To 10
        ' Sauter les nombres pairs
        If i Mod 2 = 0 Then
            GoTo SuiteLoop  ' Équivalent de "Continue"
        End If

        MsgBox "Nombre impair : " & i

SuiteLoop:
    Next i
End Sub

' ✅ Méthode plus élégante avec If
Sub AlternativeElégante()
    Dim i As Integer

    For i = 1 To 10
        ' Traiter seulement les nombres impairs
        If i Mod 2 <> 0 Then
            MsgBox "Nombre impair : " & i
        End If
    Next i
End Sub
```

## Exemples pratiques avancés

### Copier des données entre feuilles

```vba
Sub CopierDonneesVentes()
    Dim ligneSource As Integer
    Dim ligneDestination As Integer
    Dim ws1 As Worksheet, ws2 As Worksheet

    Set ws1 = Worksheets("Données")
    Set ws2 = Worksheets("Résumé")

    ligneDestination = 1

    ' Parcourir toutes les lignes de données
    For ligneSource = 2 To ws1.UsedRange.Rows.Count
        ' Copier seulement si le montant > 1000
        If ws1.Cells(ligneSource, 3).Value > 1000 Then
            ws2.Cells(ligneDestination, 1).Value = ws1.Cells(ligneSource, 1).Value  ' Nom
            ws2.Cells(ligneDestination, 2).Value = ws1.Cells(ligneSource, 3).Value  ' Montant
            ligneDestination = ligneDestination + 1
        End If
    Next ligneSource

    MsgBox "Données copiées : " & (ligneDestination - 1) & " enregistrements"
End Sub
```

### Rechercher et remplacer dans plusieurs feuilles

```vba
Sub RechercherRemplacerPartout()
    Dim feuille As Worksheet
    Dim cellule As Range
    Dim recherche As String
    Dim remplacement As String
    Dim compteur As Integer

    recherche = InputBox("Texte à rechercher :")
    remplacement = InputBox("Texte de remplacement :")

    ' Parcourir toutes les feuilles
    For Each feuille In ThisWorkbook.Worksheets
        ' Parcourir toutes les cellules utilisées
        For Each cellule In feuille.UsedRange
            If InStr(cellule.Value, recherche) > 0 Then
                cellule.Value = Replace(cellule.Value, recherche, remplacement)
                compteur = compteur + 1
            End If
        Next cellule
    Next feuille

    MsgBox compteur & " remplacements effectués"
End Sub
```

### Créer un rapport automatique

```vba
Sub CreerRapportMensuel()
    Dim mois As Integer
    Dim nomMois As String
    Dim ligne As Integer

    ligne = 1

    ' En-tête du rapport
    Range("A1").Value = "RAPPORT ANNUEL"
    Range("A1").Font.Bold = True
    ligne = 3

    ' Créer une ligne pour chaque mois
    For mois = 1 To 12
        nomMois = MonthName(mois)

        Range("A" & ligne).Value = nomMois
        Range("B" & ligne).Value = "=SOMME(" & nomMois & "!B:B)"  ' Formule dynamique
        Range("C" & ligne).Value = "=MOYENNE(" & nomMois & "!C:C)"

        ligne = ligne + 1
    Next mois

    ' Total
    Range("A" & ligne).Value = "TOTAL ANNÉE"
    Range("A" & ligne).Font.Bold = True
    Range("B" & ligne).Value = "=SOMME(B3:B14)"
    Range("B" & ligne).Font.Bold = True
End Sub
```

## Choix de la bonne boucle

### Guide de décision

**Utilisez For...Next quand :**
- Vous connaissez le nombre d'itérations
- Vous voulez un compteur précis
- Vous travaillez avec des indices numériques

```vba
' ✅ Bon usage de For...Next
For i = 1 To 100
    Cells(i, 1).Value = i * i
Next i
```

**Utilisez For Each...Next quand :**
- Vous parcourez une collection
- Vous ne connaissez pas le nombre d'éléments
- Vous travaillez avec des objets Excel

```vba
' ✅ Bon usage de For Each
For Each cellule In Selection
    cellule.Font.Bold = True
Next cellule
```

**Utilisez Do...Loop quand :**
- La condition de sortie est complexe
- Vous ne connaissez pas le nombre d'itérations
- Vous voulez une sortie flexible (Exit Do)

```vba
' ✅ Bon usage de Do...Loop
Do While Range("A" & ligne).Value <> ""
    ' Traiter la ligne
    ligne = ligne + 1
Loop
```

**Utilisez While...Wend quand :**
- Vous avez une condition simple
- Vous ne nécessitez pas de sortie anticipée
- Vous préférez une syntaxe concise

## Erreurs courantes et solutions

### 1. Modification de la collection pendant le parcours

```vba
' ❌ Problématique
For Each cellule In Range("A1:A10")
    If cellule.Value = "Supprimer" Then
        cellule.EntireRow.Delete  ' Modifie la collection !
    End If
Next cellule

' ✅ Solution - Parcourir à l'envers
For i = 10 To 1 Step -1
    If Range("A" & i).Value = "Supprimer" Then
        Rows(i).Delete
    End If
Next i
```

### 2. Oublier d'incrémenter le compteur

```vba
' ❌ Boucle infinie
Dim i As Integer  
i = 1  
Do While i <= 10  
    MsgBox i
    ' Oublié : i = i + 1
Loop

' ✅ Solution
Dim i As Integer  
i = 1  
Do While i <= 10  
    MsgBox i
    i = i + 1  ' Important !
Loop
```

### 3. Mauvaise gestion des indices

```vba
' ❌ Erreur si la plage change
For i = 1 To Range("A:A").Cells.Count
    ' Parcourt toute la colonne (plus d'un million de cellules !)
Next i

' ✅ Solution : remonter depuis la dernière ligne
For i = 1 To Range("A" & Rows.Count).End(xlUp).Row
    ' S'arrête à la dernière cellule non vide
Next i
```

## Optimisation des boucles

### Désactiver les calculs et l'affichage

```vba
Sub BoucleOptimisee()
    Dim i As Integer

    ' Optimisations pour les grandes boucles
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For i = 1 To 10000
        Cells(i, 1).Value = "Ligne " & i
    Next i

    ' Réactiver
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

### Utiliser des tableaux pour les gros volumes

```vba
Sub BoucleAvecTableau()
    Dim donnees(1 To 1000, 1 To 1) As String
    Dim i As Integer

    ' Remplir le tableau en mémoire (rapide)
    For i = 1 To 1000
        donnees(i, 1) = "Ligne " & i
    Next i

    ' Écrire tout d'un coup dans Excel (très rapide)
    Range("A1:A1000").Value = donnees
End Sub
```

## Récapitulatif des concepts clés

1. **For...Next** : Nombre d'itérations connu, avec compteur
2. **For Each...Next** : Parcourir des collections, plus simple
3. **Do...Loop** : Condition complexe, sortie flexible
4. **While...Wend** : Condition simple, syntaxe concise
5. **Exit** : Sortir prématurément d'une boucle
6. **Optimisation** : Désactiver calculs/affichage pour gros volumes
7. **Sécurité** : Éviter les boucles infinies avec des conditions de sortie

Les boucles sont l'outil le plus puissant pour automatiser les tâches répétitives. Maîtrisez-les et vous pourrez traiter des milliers de données en quelques secondes !

⏭️ [Instructions de contrôle (Exit, GoTo)](/05-structures-controle/03-instructions-controle.md)
