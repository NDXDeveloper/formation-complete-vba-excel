🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.6. Copier, coller, supprimer des données

## Introduction aux opérations de base

Les opérations de **copier**, **coller** et **supprimer** sont parmi les actions les plus courantes dans Excel. En VBA, ces opérations deviennent des outils puissants pour automatiser la manipulation de données. Vous pouvez reproduire et amplifier toutes les actions que vous faites manuellement avec Ctrl+C, Ctrl+V, et la touche Suppr.

**Analogie simple :**
- **Copier** = Photocopier un document (l'original reste intact)
- **Couper** = Découper un article de journal (l'original disparaît)
- **Coller** = Placer la copie ou l'original à un nouvel endroit
- **Supprimer** = Effacer avec une gomme

Ces opérations sont essentielles pour organiser, restructurer et nettoyer vos données automatiquement.

---

## Méthodes de copie

### 1. Copy - Copie simple

#### Copie de base

```vba
' Copier une cellule dans le presse-papier
Range("A1").Copy

' Copier une plage dans le presse-papier
Range("A1:C3").Copy

' Copier une ligne entière
Rows("5").Copy

' Copier une colonne entière
Columns("B").Copy

' Copier plusieurs plages
Range("A1:A3,C1:C3").Copy
```

#### Copie directe (sans passer par le presse-papier)

```vba
' Copier directement vers une destination (plus efficace)
Range("A1:A3").Copy Range("D1")        ' Copie A1:A3 vers D1:D3

' Note : la copie directe ne supporte qu'une seule destination contiguë

' Copier d'une feuille à une autre
Worksheets("Source").Range("A1:C3").Copy Worksheets("Destination").Range("A1")
```

### 2. Cut - Couper (déplacer)

#### Coupe de base

```vba
' Couper une cellule (déplacer)
Range("A1").Cut

' Couper une plage
Range("A1:C3").Cut

' Couper et coller directement
Range("A1:C3").Cut Destination:=Range("D1")   ' Les données sont maintenant en D1, plus en A1:C3
```

#### Coupe directe

```vba
' Couper directement vers une destination
Range("A1:A3").Cut Range("D1")     ' Déplace A1:A3 vers D1:D3
```

### 3. Copie de valeurs uniquement

#### Assigner des valeurs directement

```vba
' Copier seulement les valeurs (pas les formules ni la mise en forme)
Range("D1:D3").Value = Range("A1:A3").Value

' Copier une valeur unique vers plusieurs cellules
Range("A1:A10").Value = Range("B1").Value

' Copier des valeurs entre feuilles
Worksheets("Destination").Range("A1:A3").Value = Worksheets("Source").Range("A1:A3").Value
```

---

## Méthodes de collage

### 1. Paste - Collage standard

#### Collage simple

```vba
' Copier puis coller (méthode en 2 étapes)
Range("A1:A3").Copy                          ' Étape 1 : Copier  
ActiveSheet.Paste Destination:=Range("D1")   ' Étape 2 : Coller  

' Copie directe en une seule ligne (recommandé)
Range("A1:A3").Copy Range("D1")
```

#### Gestion du presse-papier

```vba
' Vider le presse-papier après collage
Range("A1:A3").Copy  
ActiveSheet.Paste Destination:=Range("D1")  
Application.CutCopyMode = False    ' Supprime les "fourmis" de sélection  
```

### 2. PasteSpecial - Collage spécialisé

#### Types de collage spécialisé

```vba
' Coller seulement les valeurs (pas les formules)
Range("A1:A3").Copy  
Range("D1").PasteSpecial xlPasteValues  

' Coller seulement les formules
Range("A1:A3").Copy  
Range("D1").PasteSpecial xlPasteFormulas  

' Coller seulement la mise en forme
Range("A1:A3").Copy  
Range("D1").PasteSpecial xlPasteFormats  

' Coller tout (équivalent au collage normal)
Range("A1:A3").Copy  
Range("D1").PasteSpecial xlPasteAll  

' Coller les commentaires
Range("A1:A3").Copy  
Range("D1").PasteSpecial xlPasteComments  

' Coller les largeurs de colonnes
Range("A1:C1").Copy  
Range("D1").PasteSpecial xlPasteColumnWidths  
```

#### Opérations mathématiques lors du collage

```vba
' Additionner les valeurs copiées aux valeurs existantes
Range("A1:A3").Copy  
Range("D1:D3").PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd  

' Soustraire
Range("D1:D3").PasteSpecial Paste:=xlPasteValues, Operation:=xlSubtract

' Multiplier
Range("D1:D3").PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply

' Diviser
Range("D1:D3").PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide
```

#### Transposer lors du collage

```vba
' Transposer (lignes → colonnes, colonnes → lignes)
Range("A1:A5").Copy                    ' 5 cellules en colonne  
Range("C1").PasteSpecial Transpose:=True   ' Colle en ligne (C1:G1)  

' Exemple : transformer une liste verticale en horizontale
Range("A1:A10").Copy  
Range("D1").PasteSpecial xlPasteValues, , , Transpose:=True  
```

### 3. Collage conditionnel

#### Coller en ignorant les cellules vides

```vba
' Coller en sautant les cellules vides
Range("A1:A5").Copy  
Range("D1").PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True  
```

---

## Méthodes de suppression

### 1. Clear - Effacement complet

#### Effacer tout (contenu + mise en forme)

```vba
' Effacer complètement des cellules
Range("A1:C3").Clear

' Effacer une ligne entière
Rows("5").Clear

' Effacer une colonne entière
Columns("B").Clear

' Effacer toute la feuille
Cells.Clear
```

### 2. ClearContents - Effacer le contenu uniquement

#### Garder la mise en forme

```vba
' Effacer seulement le contenu (garde couleurs, bordures, etc.)
Range("A1:C3").ClearContents

' Équivalent à appuyer sur la touche Suppr
Range("A1:C3") = ""                 ' Méthode alternative
```

### 3. ClearFormats - Effacer la mise en forme uniquement

#### Garder le contenu

```vba
' Effacer seulement la mise en forme (garde le contenu)
Range("A1:C3").ClearFormats

' Revenir au formatage par défaut
Range("A1:C3").ClearFormats
```

### 4. Effacements spécialisés

#### Effacer des éléments spécifiques

```vba
' Effacer les commentaires
Range("A1:C3").ClearComments

' Effacer les hyperliens
Range("A1:C3").Hyperlinks.Delete

' Effacer la validation de données
Range("A1:C3").Validation.Delete

' Effacer les notes (Excel 365/2019+ uniquement)
' Range("A1:C3").ClearNotes
```

### 5. Delete - Suppression avec déplacement

#### Supprimer en décalant les cellules

```vba
' Supprimer et décaler vers le haut
Range("A1:A3").Delete Shift:=xlShiftUp

' Supprimer et décaler vers la gauche
Range("A1:C1").Delete Shift:=xlShiftLeft

' Supprimer des lignes entières
Rows("5:7").Delete              ' Supprime les lignes 5, 6, 7

' Supprimer des colonnes entières
Columns("B:D").Delete           ' Supprime les colonnes B, C, D

' Suppression sans paramètre (Excel choisit automatiquement)
Range("A1:A3").Delete          ' Excel décide du décalage
```

---

## Techniques avancées de copie

### 1. Copie avec critères

#### Copier seulement certaines cellules

```vba
' Copier seulement les cellules non vides
Dim cellule As Range  
Dim plageDestination As Range  
Set plageDestination = Range("D1")  

For Each cellule In Range("A1:A10")
    If cellule.Value <> "" Then
        cellule.Copy plageDestination
        Set plageDestination = plageDestination.Offset(1, 0)  ' Ligne suivante
    End If
Next cellule
```

#### Copier avec conditions

```vba
' Copier seulement les nombres positifs
Dim i As Integer  
Dim j As Integer  
j = 1  

For i = 1 To 10
    If IsNumeric(Cells(i, 1).Value) And Cells(i, 1).Value > 0 Then
        Cells(j, 3).Value = Cells(i, 1).Value   ' Copier en colonne C
        j = j + 1
    End If
Next i
```

### 2. Copie entre classeurs

#### Copier vers un autre classeur

```vba
' Copier vers un autre classeur ouvert
Workbooks("Source.xlsx").Worksheets("Feuil1").Range("A1:C3").Copy _
    Workbooks("Destination.xlsx").Worksheets("Feuil1").Range("A1")

' Avec variables pour plus de clarté
Dim classeurSource As Workbook  
Dim classeurDest As Workbook  

Set classeurSource = Workbooks("Source.xlsx")  
Set classeurDest = Workbooks("Destination.xlsx")  

classeurSource.Worksheets("Données").Range("A1:Z100").Copy _
    classeurDest.Worksheets("Import").Range("A1")
```

### 3. Copie de mise en forme uniquement

#### Reproduire la mise en forme

```vba
' Copier la mise en forme d'une cellule modèle
Range("A1").Copy                        ' Cellule avec la mise en forme désirée  
Range("B1:B10").PasteSpecial xlPasteFormats    ' Appliquer cette mise en forme  

' Copier les largeurs de colonnes
Range("A:C").Copy  
Range("E:G").PasteSpecial xlPasteColumnWidths  

Application.CutCopyMode = False         ' Nettoyer le presse-papier
```

---

## Techniques avancées de suppression

### 1. Suppression conditionnelle

#### Supprimer selon des critères

```vba
' Supprimer les lignes contenant une valeur spécifique
Dim i As Long  
For i = 100 To 1 Step -1               ' Parcourir de bas en haut (important!)  
    If Cells(i, 1).Value = "À supprimer" Then
        Rows(i).Delete
    End If
Next i
```

#### Supprimer les lignes vides

```vba
' Supprimer toutes les lignes vides dans une plage
Dim i As Long  
Dim derniereLigne As Long  

derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

For i = derniereLigne To 1 Step -1
    If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then
        Rows(i).Delete
    End If
Next i
```

### 2. Suppression en bloc

#### Utiliser SpecialCells pour supprimer

```vba
' Supprimer toutes les cellules vides d'un coup
On Error Resume Next  
Range("A1:A100").SpecialCells(xlCellTypeBlanks).Delete Shift:=xlShiftUp  
On Error GoTo 0  

' Supprimer toutes les cellules avec erreurs
On Error Resume Next  
Range("A1:Z100").SpecialCells(xlCellTypeFormulas, xlErrors).Clear  
On Error GoTo 0  
```

---

## Gestion des erreurs et bonnes pratiques

### 1. Vérifications avant opérations

#### Vérifier l'existence des données

```vba
' Vérifier qu'il y a quelque chose à copier
If Range("A1").Value <> "" Then
    Range("A1").Copy Range("D1")
Else
    MsgBox "Aucune donnée à copier en A1"
End If

' Vérifier la taille des plages
Dim source As Range  
Dim destination As Range  

Set source = Range("A1:A10")  
Set destination = Range("D1:D5")  

If source.Cells.Count = destination.Cells.Count Then
    destination.Value = source.Value
Else
    MsgBox "Les plages n'ont pas la même taille"
End If
```

### 2. Optimisation des performances

#### Désactiver l'affichage et les calculs

```vba
Sub CopieOptimisee()
    ' Sauvegarder les états
    Dim ancienAffichage As Boolean
    Dim ancienCalcul As XlCalculation

    ancienAffichage = Application.ScreenUpdating
    ancienCalcul = Application.Calculation

    ' Optimiser
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Opérations de copie/suppression
    Range("A1:A1000").Copy Range("D1")
    Range("G1:G1000").Clear

    ' Restaurer les états
    Application.ScreenUpdating = ancienAffichage
    Application.Calculation = ancienCalcul

    ' Nettoyer
    Application.CutCopyMode = False
End Sub
```

### 3. Gestion des erreurs

#### Protéger contre les erreurs courantes

```vba
Sub CopieSecurisee()
    On Error GoTo GestionErreur

    ' Vérifier que la feuille source existe
    Dim feuilleSource As Worksheet
    Set feuilleSource = Worksheets("Source")

    ' Vérifier que la plage n'est pas vide
    If feuilleSource.Range("A1").Value = "" Then
        MsgBox "La cellule source est vide"
        Exit Sub
    End If

    ' Effectuer la copie
    feuilleSource.Range("A1:C10").Copy Worksheets("Destination").Range("A1")

    MsgBox "Copie réussie"
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de la copie : " & Err.Description
    Application.CutCopyMode = False
End Sub
```

---

## Exemples pratiques complets

### 1. Consolidation de données

#### Copier depuis plusieurs feuilles

```vba
Sub ConsoliderDonnees()
    Dim feuille As Worksheet
    Dim ligneDestination As Long

    ' Commencer en ligne 2 (ligne 1 pour les en-têtes)
    ligneDestination = 2

    ' Feuille de destination
    Worksheets("Consolidation").Activate

    ' Parcourir toutes les feuilles sauf "Consolidation"
    For Each feuille In Worksheets
        If feuille.Name <> "Consolidation" Then
            ' Trouver la dernière ligne avec des données
            Dim derniereLigne As Long
            derniereLigne = feuille.Cells(Rows.Count, 1).End(xlUp).Row

            ' Copier les données (sans les en-têtes)
            If derniereLigne > 1 Then
                feuille.Range("A2:C" & derniereLigne).Copy _
                    Worksheets("Consolidation").Cells(ligneDestination, 1)

                ' Mettre à jour la ligne de destination
                ligneDestination = ligneDestination + (derniereLigne - 1)
            End If
        End If
    Next feuille

    Application.CutCopyMode = False
    MsgBox "Consolidation terminée"
End Sub
```

### 2. Nettoyage et réorganisation

#### Supprimer les doublons et réorganiser

```vba
Sub NettoyerDonnees()
    Dim i As Long
    Dim j As Long
    Dim derniereLigne As Long

    ' Trouver la dernière ligne
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    ' Supprimer les lignes vides (de bas en haut)
    For i = derniereLigne To 2 Step -1
        If Cells(i, 1).Value = "" Then
            Rows(i).Delete
        End If
    Next i

    ' Recalculer la dernière ligne
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    ' Supprimer les doublons simples
    For i = derniereLigne To 2 Step -1
        For j = i - 1 To 1 Step -1
            If Cells(i, 1).Value = Cells(j, 1).Value And Cells(i, 1).Value <> "" Then
                Rows(i).Delete
                Exit For
            End If
        Next j
    Next i

    MsgBox "Nettoyage terminé"
End Sub
```

### 3. Sauvegarde et archivage

#### Copier vers un fichier d'archive

```vba
Sub ArchiverDonnees()
    Dim classeurArchive As Workbook
    Dim nomFichier As String

    ' Créer le nom du fichier avec la date
    nomFichier = "Archive_" & Format(Date, "yyyy-mm-dd") & ".xlsx"

    ' Créer un nouveau classeur pour l'archive
    Set classeurArchive = Workbooks.Add

    ' Copier toutes les données
    ThisWorkbook.Worksheets("Données").UsedRange.Copy _
        classeurArchive.Worksheets(1).Range("A1")

    ' Sauvegarder l'archive
    classeurArchive.SaveAs ThisWorkbook.Path & "\" & nomFichier
    classeurArchive.Close

    ' Nettoyer les données originales
    ThisWorkbook.Worksheets("Données").UsedRange.ClearContents

    Application.CutCopyMode = False
    MsgBox "Données archivées dans " & nomFichier
End Sub
```

---

## Récapitulatif et conseils

### Méthodes principales :

#### Copie :
- **Copy** : Copie standard vers le presse-papier ou directement
- **Value = Value** : Copie de valeurs uniquement (plus rapide)

#### Collage :
- **ActiveSheet.Paste** : Collage standard (méthode de Worksheet, pas de Range)
- **PasteSpecial** : Collage avec options (valeurs, formats, opérations)

#### Suppression :
- **Clear** : Efface tout (contenu + formats)
- **ClearContents** : Efface seulement le contenu
- **Delete** : Supprime avec décalage des cellules

### Bonnes pratiques :

1. **Préférez les assignations directes** pour les valeurs : `Range("D1").Value = Range("A1").Value`
2. **Utilisez Application.CutCopyMode = False** pour nettoyer le presse-papier
3. **Désactivez l'affichage** pour les opérations massives
4. **Parcourez de bas en haut** lors de suppressions de lignes
5. **Vérifiez l'existence** des données avant opérations
6. **Gérez les erreurs** pour éviter les plantages

### Optimisations :

- **Opérations par blocs** plutôt que cellule par cellule
- **Références directes** plutôt que copier-coller quand possible
- **Variables Range** pour éviter les accès répétés
- **Désactivation des calculs** pendant les opérations massives

Ces techniques de copie, collage et suppression forment la base de la plupart des automatisations Excel. Maîtriser ces opérations vous permettra de créer des macros robustes pour organiser, nettoyer et restructurer vos données efficacement.

⏭️
