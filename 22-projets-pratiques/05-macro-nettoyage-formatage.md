🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.5. Macro de nettoyage et formatage

## Introduction

Une macro de nettoyage et formatage est un outil essentiel pour automatiser la préparation et la présentation des données dans Excel. Elle permet de transformer des données brutes, souvent désorganisées ou mal formatées, en tableaux propres et professionnels. Dans ce chapitre, nous allons créer une macro complète qui nettoie et formate automatiquement vos données.

## Objectifs du projet

Notre macro de nettoyage et formatage permettra de :
- Supprimer les espaces indésirables et les caractères spéciaux
- Standardiser la casse du texte (majuscules, minuscules, première lettre)
- Nettoyer les données numériques et les dates
- Supprimer les lignes et colonnes vides
- Appliquer un formatage professionnel automatique
- Créer des en-têtes et des bordures
- Ajuster automatiquement la largeur des colonnes

## Comprendre les problèmes de données courantes

### Types de problèmes fréquents

Avant de créer notre macro, il est important de comprendre les problèmes les plus fréquents dans les données :

```vba
' Exemples de données problématiques :
' - Espaces en début/fin : "  Jean Dupont  "
' - Casse incohérente : "jEaN dUpOnT", "JEAN DUPONT"
' - Caractères indésirables : "Jean@Dupont#", "123.45€"
' - Dates mal formatées : "01-12-2023", "1/12/23"
' - Nombres avec du texte : "1,234.56 €", "123 unités"
' - Lignes/colonnes vides
' - Formatage incohérent
```

## Étape 1 : Structure générale de la macro

### Macro principale

```vba
Sub NettoyageFormatageComplet()
    ' Macro principale qui orchestre tout le processus de nettoyage

    ' Désactiver les mises à jour d'écran pour accélérer le processus
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Afficher un message d'information
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Cette macro va nettoyer et formater la feuille active." & vbCrLf & _
                    "Voulez-vous continuer ?", vbQuestion + vbYesNo, "Nettoyage et formatage")

    If reponse = vbNo Then
        ReactiverParametres
        Exit Sub
    End If

    ' Sauvegarder l'état actuel (possibilité d'annulation)
    Dim etatInitial As Worksheet
    Set etatInitial = ActiveSheet

    ' Démarrer le processus de nettoyage
    On Error GoTo GestionErreur

    ' Étapes du nettoyage dans l'ordre logique
    SupprimerLignesColonnesVides
    NettoyerTexte
    NettoyerNombres
    NettoyerDates
    FormaterTableau
    AjusterColonnes

    ' Afficher un message de succès
    MsgBox "Nettoyage et formatage terminés avec succès !", vbInformation, "Terminé"

    ' Réactiver les paramètres
    ReactiverParametres
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors du nettoyage : " & Err.Description, vbCritical, "Erreur"
    ReactiverParametres
End Sub

Private Sub ReactiverParametres()
    ' Réactiver tous les paramètres Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

## Étape 2 : Suppression des lignes et colonnes vides

### Supprimer les éléments vides

```vba
Private Sub SupprimerLignesColonnesVides()
    ' Supprimer toutes les lignes et colonnes complètement vides

    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim derniereColonne As Long
    Dim i As Long, j As Long
    Dim ligneVide As Boolean
    Dim colonneVide As Boolean

    Set ws = ActiveSheet

    ' Déterminer la zone de données réelle
    derniereLigne = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    derniereColonne = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Supprimer les lignes vides (de bas en haut pour éviter les problèmes d'index)
    For i = derniereLigne To 1 Step -1
        ligneVide = True

        ' Vérifier chaque cellule de la ligne
        For j = 1 To derniereColonne
            If ws.Cells(i, j).Value <> "" Then
                ligneVide = False
                Exit For
            End If
        Next j

        ' Supprimer la ligne si elle est vide
        If ligneVide Then
            ws.Rows(i).Delete
        End If
    Next i

    ' Recalculer après suppression des lignes
    derniereLigne = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    derniereColonne = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Supprimer les colonnes vides (de droite à gauche)
    For j = derniereColonne To 1 Step -1
        colonneVide = True

        ' Vérifier chaque cellule de la colonne
        For i = 1 To derniereLigne
            If ws.Cells(i, j).Value <> "" Then
                colonneVide = False
                Exit For
            End If
        Next i

        ' Supprimer la colonne si elle est vide
        If colonneVide Then
            ws.Columns(j).Delete
        End If
    Next j
End Sub
```

## Étape 3 : Nettoyage du texte

### Standardiser le texte

```vba
Private Sub NettoyerTexte()
    ' Nettoyer et standardiser toutes les données texte

    Dim ws As Worksheet
    Dim plageData As Range
    Dim cellule As Range
    Dim texteNettoye As String

    Set ws = ActiveSheet

    ' Définir la plage de données
    Set plageData = ws.UsedRange

    ' Parcourir chaque cellule
    For Each cellule In plageData
        If cellule.Value <> "" And IsText(cellule.Value) Then
            ' Nettoyer le texte de la cellule
            texteNettoye = NettoyerCelluleTexte(CStr(cellule.Value))
            cellule.Value = texteNettoye
        End If
    Next cellule
End Sub

Private Function NettoyerCelluleTexte(texte As String) As String
    Dim resultat As String
    Dim i As Integer
    Dim caractere As String

    ' Étape 1 : Supprimer les espaces en début et fin
    resultat = Trim(texte)

    ' Étape 2 : Supprimer les espaces multiples (remplacer par un seul espace)
    Do While InStr(resultat, "  ") > 0
        resultat = Replace(resultat, "  ", " ")
    Loop

    ' Étape 3 : Supprimer les caractères spéciaux indésirables
    Dim caracteresIndesirables As String
    caracteresIndesirables = "@#$%^&*()+={}[]|\\:;""'<>?/~`"

    For i = 1 To Len(caracteresIndesirables)
        caractere = Mid(caracteresIndesirables, i, 1)
        resultat = Replace(resultat, caractere, "")
    Next i

    ' Étape 4 : Standardiser la casse (première lettre de chaque mot en majuscule)
    resultat = FormatagePropreCasse(resultat)

    NettoyerCelluleTexte = resultat
End Function

Private Function FormatagePropreCasse(texte As String) As String
    ' Convertir en "Propre" (première lettre de chaque mot en majuscule)
    Dim mots() As String
    Dim i As Integer
    Dim resultat As String

    ' Diviser le texte en mots
    mots = Split(LCase(texte), " ")

    ' Traiter chaque mot
    For i = 0 To UBound(mots)
        If Len(mots(i)) > 0 Then
            ' Première lettre en majuscule, le reste en minuscule
            mots(i) = UCase(Left(mots(i), 1)) & Mid(mots(i), 2)
        End If
    Next i

    ' Reconstituer le texte
    FormatagePropreCasse = Join(mots, " ")
End Function

Private Function IsText(valeur As Variant) As Boolean
    ' Vérifier si une valeur est du texte (et non un nombre ou une date)
    IsText = (VarType(valeur) = vbString) And Not IsNumeric(valeur) And Not IsDate(valeur)
End Function
```

## Étape 4 : Nettoyage des nombres

### Standardiser les données numériques

```vba
Private Sub NettoyerNombres()
    ' Nettoyer et standardiser les données numériques

    Dim ws As Worksheet
    Dim plageData As Range
    Dim cellule As Range
    Dim valeurNettoyee As String
    Dim nombreFinal As Double

    Set ws = ActiveSheet
    Set plageData = ws.UsedRange

    For Each cellule In plageData
        If cellule.Value <> "" Then
            valeurNettoyee = NettoyerNombre(CStr(cellule.Value))

            ' Si c'est maintenant un nombre valide, le convertir
            If IsNumeric(valeurNettoyee) And valeurNettoyee <> "" Then
                nombreFinal = CDbl(valeurNettoyee)
                cellule.Value = nombreFinal

                ' Appliquer un formatage numérique approprié
                If nombreFinal = Int(nombreFinal) Then
                    ' Nombre entier
                    cellule.NumberFormat = "#,##0"
                Else
                    ' Nombre décimal
                    cellule.NumberFormat = "#,##0.00"
                End If
            End If
        End If
    Next cellule
End Sub

Private Function NettoyerNombre(texte As String) As String
    Dim resultat As String
    Dim i As Integer
    Dim caractere As String

    resultat = Trim(texte)

    ' Supprimer les symboles monétaires et unités courantes
    Dim symbolesASupprimer As Variant
    symbolesASupprimer = Array("€", "$", "£", "¥", "%", "°", "kg", "g", "m", "cm", "mm", _
                              "l", "ml", "h", "min", "s", "unités", "pcs", "pièces")

    Dim symbole As Variant
    For Each symbole In symbolesASupprimer
        resultat = Replace(resultat, symbole, "", , , vbTextCompare)
    Next symbole

    ' Supprimer les lettres (sauf E pour notation scientifique)
    Dim nouveauResultat As String
    For i = 1 To Len(resultat)
        caractere = Mid(resultat, i, 1)
        If IsNumeric(caractere) Or caractere = "." Or caractere = "," Or _
           caractere = "-" Or caractere = "+" Or UCase(caractere) = "E" Then
            nouveauResultat = nouveauResultat & caractere
        End If
    Next i

    resultat = nouveauResultat

    ' Standardiser les séparateurs décimaux (virgule vers point)
    ' Gérer le cas des milliers (ex: 1,234.56 ou 1.234,56)
    If InStr(resultat, ",") > 0 And InStr(resultat, ".") > 0 Then
        ' Les deux sont présents - déterminer le rôle de chacun
        If InStr(resultat, ",") < InStr(resultat, ".") Then
            ' Virgule avant point : virgule = milliers, point = décimal (ex: 1,234.56)
            resultat = Replace(resultat, ",", "")
        Else
            ' Point avant virgule : point = milliers, virgule = décimal (ex: 1.234,56)
            resultat = Replace(resultat, ".", "")
            resultat = Replace(resultat, ",", ".")
        End If
    ElseIf InStr(resultat, ",") > 0 Then
        ' Seulement des virgules - probablement décimales en français
        resultat = Replace(resultat, ",", ".")
    End If

    ' Nettoyer les signes multiples
    Do While InStr(resultat, "..") > 0
        resultat = Replace(resultat, "..", ".")
    Loop

    NettoyerNombre = Trim(resultat)
End Function
```

## Étape 5 : Nettoyage des dates

### Standardiser les formats de date

```vba
Private Sub NettoyerDates()
    ' Nettoyer et standardiser les formats de date

    Dim ws As Worksheet
    Dim plageData As Range
    Dim cellule As Range
    Dim dateNettoyee As Date

    Set ws = ActiveSheet
    Set plageData = ws.UsedRange

    For Each cellule In plageData
        If cellule.Value <> "" And Not IsNumeric(cellule.Value) Then
            If TenterConversionDate(CStr(cellule.Value), dateNettoyee) Then
                cellule.Value = dateNettoyee
                cellule.NumberFormat = "dd/mm/yyyy"
            End If
        End If
    Next cellule
End Sub

Private Function TenterConversionDate(texte As String, ByRef dateResultat As Date) As Boolean
    On Error GoTo ErreurConversion

    Dim texteNettoye As String
    texteNettoye = Trim(texte)

    ' Remplacer différents séparateurs par des slash
    texteNettoye = Replace(texteNettoye, "-", "/")
    texteNettoye = Replace(texteNettoye, ".", "/")
    texteNettoye = Replace(texteNettoye, " ", "/")

    ' Essayer différents formats de date
    Dim formatsAEssayer As Variant
    formatsAEssayer = Array("dd/mm/yyyy", "dd/mm/yy", "mm/dd/yyyy", "mm/dd/yy", _
                           "yyyy/mm/dd", "yy/mm/dd")

    Dim format As Variant
    For Each format In formatsAEssayer
        If EstDateValide(texteNettoye, CStr(format)) Then
            dateResultat = CDate(texteNettoye)
            TenterConversionDate = True
            Exit Function
        End If
    Next format

    ' Si les formats standards échouent, essayer la conversion directe
    If IsDate(texteNettoye) Then
        dateResultat = CDate(texteNettoye)

        ' Vérifier que la date est raisonnable (entre 1900 et 2100)
        If Year(dateResultat) >= 1900 And Year(dateResultat) <= 2100 Then
            TenterConversionDate = True
            Exit Function
        End If
    End If

ErreurConversion:
    TenterConversionDate = False
End Function

Private Function EstDateValide(texte As String, format As String) As Boolean
    On Error GoTo ErreurValidation

    ' Cette fonction vérifie si le texte correspond au format spécifié
    Dim parties() As String
    parties = Split(texte, "/")

    If UBound(parties) <> 2 Then
        EstDateValide = False
        Exit Function
    End If

    ' Vérifier que toutes les parties sont numériques
    Dim i As Integer
    For i = 0 To 2
        If Not IsNumeric(parties(i)) Then
            EstDateValide = False
            Exit Function
        End If
    Next i

    ' Essayer de créer la date
    Dim dateTest As Date
    dateTest = CDate(texte)

    ' Vérifications de cohérence
    If Year(dateTest) >= 1900 And Year(dateTest) <= 2100 And _
       Month(dateTest) >= 1 And Month(dateTest) <= 12 And _
       Day(dateTest) >= 1 And Day(dateTest) <= 31 Then
        EstDateValide = True
    Else
        EstDateValide = False
    End If

    Exit Function

ErreurValidation:
    EstDateValide = False
End Function
```

## Étape 6 : Formatage du tableau

### Appliquer un formatage professionnel

```vba
Private Sub FormaterTableau()
    ' Appliquer un formatage professionnel au tableau

    Dim ws As Worksheet
    Dim plageData As Range
    Dim ligneEnTete As Range

    Set ws = ActiveSheet
    Set plageData = ws.UsedRange

    ' Vérifier qu'il y a des données
    If plageData.Rows.Count < 1 Then Exit Sub

    ' Formatage général du tableau
    With plageData
        .Font.Name = "Calibri"
        .Font.Size = 11
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
    End With

    ' Formatage spécial pour la première ligne (en-têtes)
    Set ligneEnTete = ws.Range(ws.Cells(1, 1), ws.Cells(1, plageData.Columns.Count))

    With ligneEnTete
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(68, 114, 196)  ' Bleu professionnel
        .Font.Color = RGB(255, 255, 255)    ' Texte blanc
        .HorizontalAlignment = xlHAlignCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
        .Borders.Weight = xlMedium
    End With

    ' Formatage alternant pour les lignes de données
    If plageData.Rows.Count > 1 Then
        FormaterLignesAlternees plageData
    End If

    ' Ajouter des bordures
    AjouterBordures plageData

    ' Centrer les en-têtes et aligner les données
    AlignementDonnees plageData
End Sub

Private Sub FormaterLignesAlternees(plageData As Range)
    ' Créer un effet de lignes alternées pour améliorer la lisibilité

    Dim i As Long
    Dim ligneRange As Range

    For i = 2 To plageData.Rows.Count Step 2
        Set ligneRange = plageData.Rows(i)
        ligneRange.Interior.Color = RGB(242, 242, 242)  ' Gris très clair
    Next i
End Sub

Private Sub AjouterBordures(plageData As Range)
    ' Ajouter des bordures propres au tableau

    With plageData.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)  ' Gris moyen
        .Weight = xlThin
    End With

    ' Bordure extérieure plus épaisse
    With plageData.BorderAround
        .Color = RGB(68, 114, 196)
        .Weight = xlMedium
    End With
End Sub

Private Sub AlignementDonnees(plageData As Range)
    ' Appliquer l'alignement approprié selon le type de données

    Dim colonne As Range
    Dim cellule As Range
    Dim typeColonne As String

    For Each colonne In plageData.Columns
        typeColonne = DeterminerTypeColonne(colonne)

        Select Case typeColonne
            Case "Numérique"
                colonne.HorizontalAlignment = xlHAlignRight
            Case "Date"
                colonne.HorizontalAlignment = xlHAlignCenter
            Case "Texte"
                colonne.HorizontalAlignment = xlHAlignLeft
            Case Else
                colonne.HorizontalAlignment = xlHAlignLeft
        End Select
    Next colonne
End Sub

Private Function DeterminerTypeColonne(colonne As Range) As String
    ' Déterminer le type de données prédominant dans une colonne

    Dim cellule As Range
    Dim compteurNombre As Long, compteurDate As Long, compteurTexte As Long
    Dim total As Long

    ' Ignorer la première ligne (en-têtes)
    For Each cellule In colonne.Offset(1, 0).Resize(colonne.Rows.Count - 1, 1)
        If cellule.Value <> "" Then
            total = total + 1

            If IsNumeric(cellule.Value) Then
                compteurNombre = compteurNombre + 1
            ElseIf IsDate(cellule.Value) Then
                compteurDate = compteurDate + 1
            Else
                compteurTexte = compteurTexte + 1
            End If
        End If
    Next cellule

    ' Déterminer le type majoritaire
    If compteurNombre >= compteurDate And compteurNombre >= compteurTexte Then
        DeterminerTypeColonne = "Numérique"
    ElseIf compteurDate >= compteurTexte Then
        DeterminerTypeColonne = "Date"
    Else
        DeterminerTypeColonne = "Texte"
    End If
End Function
```

## Étape 7 : Ajustement automatique des colonnes

### Optimiser la largeur des colonnes

```vba
Private Sub AjusterColonnes()
    ' Ajuster automatiquement la largeur des colonnes pour un affichage optimal

    Dim ws As Worksheet
    Dim plageData As Range
    Dim i As Long
    Dim largeurOptimale As Double
    Dim largeurMaximale As Double
    Dim largeurMinimale As Double

    Set ws = ActiveSheet
    Set plageData = ws.UsedRange

    ' Définir les limites de largeur
    largeurMinimale = 8    ' Largeur minimale en caractères
    largeurMaximale = 50   ' Largeur maximale en caractères

    ' Ajuster chaque colonne individuellement
    For i = 1 To plageData.Columns.Count
        ' AutoFit initial
        ws.Columns(i).AutoFit

        ' Récupérer la largeur calculée
        largeurOptimale = ws.Columns(i).ColumnWidth

        ' Appliquer les limites
        If largeurOptimale < largeurMinimale Then
            ws.Columns(i).ColumnWidth = largeurMinimale
        ElseIf largeurOptimale > largeurMaximale Then
            ws.Columns(i).ColumnWidth = largeurMaximale
            ' Si la colonne est trop large, activer le retour à la ligne
            ws.Columns(i).WrapText = True
        End If
    Next i

    ' Ajuster la hauteur des lignes si nécessaire
    plageData.Rows.AutoFit
End Sub
```

## Étape 8 : Fonctions utilitaires supplémentaires

### Fonctions de nettoyage spécialisées

```vba
Sub NettoyageRapide()
    ' Version simplifiée pour un nettoyage rapide

    Application.ScreenUpdating = False

    Dim plageSelection As Range
    Set plageSelection = Selection

    ' Nettoyer seulement la sélection actuelle
    NettoyerSelection plageSelection

    Application.ScreenUpdating = True
    MsgBox "Nettoyage rapide terminé !", vbInformation
End Sub

Private Sub NettoyerSelection(plage As Range)
    ' Nettoyer uniquement la plage sélectionnée

    Dim cellule As Range

    For Each cellule In plage
        If cellule.Value <> "" Then
            ' Supprimer les espaces en excès
            If IsText(cellule.Value) Then
                cellule.Value = Trim(cellule.Value)
                ' Supprimer les espaces multiples
                Do While InStr(cellule.Value, "  ") > 0
                    cellule.Value = Replace(cellule.Value, "  ", " ")
                Loop
            End If
        End If
    Next cellule
End Sub

Sub SupprimerFormatage()
    ' Supprimer tout le formatage et garder seulement les données

    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Voulez-vous supprimer tout le formatage de la feuille ?", _
                    vbQuestion + vbYesNo, "Suppression du formatage")

    If reponse = vbYes Then
        With ActiveSheet.UsedRange
            .ClearFormats
            .Font.Name = "Calibri"
            .Font.Size = 11
        End With
        MsgBox "Formatage supprimé !", vbInformation
    End If
End Sub

Sub CreerTableauExcel()
    ' Convertir la plage de données en tableau Excel

    Dim ws As Worksheet
    Dim plageData As Range
    Dim tableau As ListObject

    Set ws = ActiveSheet
    Set plageData = ws.UsedRange

    ' Vérifier qu'il y a des données
    If plageData.Rows.Count < 2 Then
        MsgBox "Il faut au moins 2 lignes (en-têtes + données) pour créer un tableau.", _
               vbInformation
        Exit Sub
    End If

    ' Supprimer les tableaux existants
    Dim i As Integer
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i

    ' Créer le nouveau tableau
    Set tableau = ws.ListObjects.Add(xlSrcRange, plageData, , xlYes)
    tableau.TableStyle = "TableStyleMedium2"

    MsgBox "Tableau Excel créé avec succès !", vbInformation
End Sub
```

## Étape 9 : Interface utilisateur pour la macro

### Menu personnalisé

```vba
Sub CreerMenuNettoyage()
    ' Créer un menu personnalisé dans le ruban (optionnel)

    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Quelle action souhaitez-vous effectuer ?" & vbCrLf & vbCrLf & _
                    "Oui = Nettoyage complet" & vbCrLf & _
                    "Non = Nettoyage rapide" & vbCrLf & _
                    "Annuler = Formatage seulement", _
                    vbYesNoCancel + vbQuestion, "Options de nettoyage")

    Select Case reponse
        Case vbYes
            NettoyageFormatageComplet
        Case vbNo
            NettoyageRapide
        Case vbCancel
            FormaterTableau
            AjusterColonnes
            MsgBox "Formatage appliqué !", vbInformation
    End Select
End Sub
```

## Conseils d'utilisation et bonnes pratiques

### Avant d'utiliser la macro

1. **Sauvegardez toujours** votre fichier avant d'exécuter la macro
2. **Testez sur un échantillon** de données d'abord
3. **Vérifiez les types de données** présents dans votre tableau
4. **Identifiez les colonnes sensibles** (identifiants, codes spéciaux)

### Personnalisation possible

La macro peut être facilement adaptée pour :
- **Types de données spécifiques** : ajouter des règles de nettoyage pour vos données
- **Formatage d'entreprise** : utiliser les couleurs et polices de votre organisation
- **Règles métier** : intégrer des validations spécifiques à votre secteur
- **Langues différentes** : adapter les formats de date et nombres à votre région

### Limites et précautions

- La macro modifie les données de manière **irréversible**
- Certains **caractères spéciaux** peuvent être importants dans votre contexte
- Les **formats numériques régionaux** peuvent nécessiter des ajustements
- Les **très gros fichiers** peuvent nécessiter une optimisation supplémentaire

## Conclusion

Cette macro de nettoyage et formatage constitue un outil puissant pour automatiser la préparation de vos données. Elle combine des techniques de nettoyage sophistiquées avec un formatage professionnel, permettant de transformer rapidement des données brutes en tableaux présentables.

L'approche modulaire du code permet de facilement adapter et étendre les fonctionnalités selon vos besoins spécifiques. Avec ces bases, vous pouvez créer des outils de nettoyage sur mesure pour votre organisation.

⏭️
