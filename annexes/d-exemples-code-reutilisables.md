🔝 Retour au [Sommaire](/SOMMAIRE.md)

# D. Exemples de code réutilisables

## Introduction

Cette annexe contient des exemples de code VBA prêts à utiliser, testés et commentés. Ces "briques" de code peuvent être copiées directement dans vos projets et adaptées selon vos besoins. Chaque exemple est accompagné d'explications simples pour comprendre son fonctionnement.

**Comment utiliser ces exemples :**
- **Copiez le code** dans un module VBA
- **Lisez les commentaires** pour comprendre chaque étape
- **Modifiez les variables** selon vos besoins
- **Testez toujours** avant d'utiliser sur des données importantes
- **Adaptez** les exemples à votre contexte spécifique

---

## 1. Manipulation de base des cellules

### Écrire et lire des valeurs

```vba
Sub ExempleEcritureLecture()
    ' Écrire une valeur dans une cellule
    Range("A1").Value = "Bonjour"
    Range("B1").Value = 123
    Range("C1").Value = Date() ' Date du jour

    ' Lire une valeur depuis une cellule
    Dim monTexte As String
    Dim monNombre As Integer
    Dim maDate As Date

    monTexte = Range("A1").Value
    monNombre = Range("B1").Value
    maDate = Range("C1").Value

    ' Afficher les valeurs lues
    MsgBox "Texte: " & monTexte & ", Nombre: " & monNombre
End Sub
```

### Copier des données entre cellules

```vba
Sub CopierDonnees()
    ' Copier une cellule vers une autre
    Range("A1").Copy Range("D1")

    ' Copier une plage de cellules
    Range("A1:C3").Copy Range("E1")

    ' Copier avec valeurs seulement (sans formules)
    Range("A1:A10").Copy
    Range("B1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False ' Enlève les pointillés
End Sub
```

### Effacer le contenu

```vba
Sub EffacerContenu()
    ' Effacer le contenu d'une cellule
    Range("A1").ClearContents

    ' Effacer une plage entière
    Range("A1:C10").ClearContents

    ' Effacer tout (contenu + format)
    Range("A1:C10").Clear

    ' Effacer seulement le format
    Range("A1:C10").ClearFormats
End Sub
```

---

## 2. Formatage des cellules

### Formatage de base

```vba
Sub FormaterCellules()
    With Range("A1:C1")
        .Font.Bold = True               ' Gras
        .Font.Size = 14                 ' Taille de police
        .Font.Color = RGB(255, 0, 0)    ' Couleur rouge
        .Interior.Color = RGB(255, 255, 0) ' Fond jaune
        .HorizontalAlignment = xlCenter  ' Centré
        .Borders.LineStyle = xlContinuous ' Bordures
    End With
End Sub
```

### Formatage conditionnel simple

```vba
Sub FormaterSelonValeur()
    Dim cellule As Range

    ' Parcourir une plage et formater selon la valeur
    For Each cellule In Range("A1:A10")
        If IsNumeric(cellule.Value) Then
            If cellule.Value > 100 Then
                cellule.Interior.Color = RGB(0, 255, 0) ' Vert si > 100
            ElseIf cellule.Value < 50 Then
                cellule.Interior.Color = RGB(255, 0, 0) ' Rouge si < 50
            Else
                cellule.Interior.Color = RGB(255, 255, 0) ' Jaune entre 50 et 100
            End If
        End If
    Next cellule
End Sub
```

---

## 3. Gestion des feuilles de calcul

### Créer, renommer, supprimer des feuilles

```vba
Sub GererFeuilles()
    ' Créer une nouvelle feuille
    Dim nouvelleFeuille As Worksheet
    Set nouvelleFeuille = Worksheets.Add
    nouvelleFeuille.Name = "MaNouvelleFeuille"

    ' Vérifier si une feuille existe avant de la créer
    Dim nomFeuille As String
    nomFeuille = "FeuilleTest"

    If Not FeuilleExiste(nomFeuille) Then
        Worksheets.Add.Name = nomFeuille
    End If

    ' Supprimer une feuille (avec confirmation)
    Application.DisplayAlerts = False
    Worksheets("FeuilleTest").Delete
    Application.DisplayAlerts = True
End Sub

Function FeuilleExiste(nomFeuille As String) As Boolean
    Dim feuille As Worksheet
    For Each feuille In Worksheets
        If feuille.Name = nomFeuille Then
            FeuilleExiste = True
            Exit Function
        End If
    Next feuille
    FeuilleExiste = False
End Function
```

### Protéger et déprotéger des feuilles

```vba
Sub ProtegerFeuille()
    ' Protéger la feuille active
    ActiveSheet.Protect Password:="motdepasse", _
                      DrawingObjects:=True, _
                      Contents:=True, _
                      Scenarios:=True

    ' Déprotéger la feuille
    ActiveSheet.Unprotect Password:="motdepasse"
End Sub
```

---

## 4. Boucles et parcours de données

### Parcourir des lignes avec des données

```vba
Sub ParcourirDonnees()
    Dim derniereLigne As Long
    Dim i As Long

    ' Trouver la dernière ligne avec des données
    derniereLigne = Range("A" & Rows.Count).End(xlUp).Row

    ' Parcourir toutes les lignes
    For i = 1 To derniereLigne
        ' Traiter chaque ligne
        If Range("A" & i).Value <> "" Then
            Range("B" & i).Value = "Traité le " & Date()
        End If
    Next i
End Sub
```

### Rechercher une valeur

```vba
Function TrouverValeur(valeurCherchee As String, colonneRecherche As String) As Long
    Dim derniereLigne As Long
    Dim i As Long

    derniereLigne = Range(colonneRecherche & Rows.Count).End(xlUp).Row

    For i = 1 To derniereLigne
        If Range(colonneRecherche & i).Value = valeurCherchee Then
            TrouverValeur = i ' Retourne le numéro de ligne
            Exit Function
        End If
    Next i

    TrouverValeur = 0 ' Retourne 0 si non trouvé
End Function
```

---

## 5. Gestion des fichiers

### Ouvrir et fermer des classeurs

```vba
Sub OuvrirClasseur()
    Dim cheminFichier As String
    Dim monClasseur As Workbook

    cheminFichier = "C:\MonDossier\MonFichier.xlsx"

    ' Vérifier si le fichier existe
    If Dir(cheminFichier) <> "" Then
        Set monClasseur = Workbooks.Open(cheminFichier)
        MsgBox "Fichier ouvert avec succès"

        ' Faire quelque chose avec le classeur

        ' Fermer et sauvegarder
        monClasseur.Save
        monClasseur.Close
    Else
        MsgBox "Fichier introuvable: " & cheminFichier
    End If
End Sub
```

### Sauvegarder avec un nom spécifique

```vba
Sub SauvegarderSous()
    Dim nouveauNom As String
    nouveauNom = "MonFichier_" & Format(Date(), "yyyy-mm-dd") & ".xlsx"

    ' Sauvegarder dans le même dossier
    ActiveWorkbook.SaveAs Filename:=nouveauNom

    ' Ou spécifier un dossier complet
    ActiveWorkbook.SaveAs Filename:="C:\MonDossier\" & nouveauNom
End Sub
```

---

## 6. Interface utilisateur

### Boîtes de dialogue simples

```vba
Sub BoitesDialogue()
    Dim reponse As String
    Dim confirmation As VbMsgBoxResult

    ' Demander une saisie
    reponse = InputBox("Entrez votre nom:", "Saisie", "Nom par défaut")

    If reponse <> "" Then
        ' Demander confirmation
        confirmation = MsgBox("Voulez-vous continuer avec " & reponse & "?", _
                             vbYesNo + vbQuestion, "Confirmation")

        If confirmation = vbYes Then
            Range("A1").Value = reponse
            MsgBox "Nom sauvegardé!", vbInformation
        End If
    End If
End Sub
```

### Choisir un fichier

```vba
Sub ChoisirFichier()
    Dim cheminFichier As String

    ' Ouvrir la boîte de dialogue de sélection
    cheminFichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")

    If cheminFichier <> "False" Then
        MsgBox "Fichier sélectionné: " & cheminFichier
        ' Ici vous pouvez ouvrir le fichier
        Workbooks.Open cheminFichier
    Else
        MsgBox "Aucun fichier sélectionné"
    End If
End Sub
```

---

## 7. Calculs et formules

### Insérer des formules

```vba
Sub InsererFormules()
    ' Formule simple
    Range("C1").Formula = "=A1+B1"

    ' Formule avec références absolues
    Range("C2").Formula = "=A2*$B$1"

    ' Utiliser WorksheetFunction pour calculer en VBA
    Dim somme As Double
    somme = Application.WorksheetFunction.Sum(Range("A1:A10"))
    Range("A11").Value = somme
End Sub
```

### Fonctions personnalisées simples

```vba
Function CalculerTVA(montantHT As Double, Optional tauxTVA As Double = 0.2) As Double
    ' Fonction personnalisée pour calculer la TVA
    ' Utilisation dans Excel: =CalculerTVA(100; 0.196)
    CalculerTVA = montantHT * tauxTVA
End Function

Function ConcatenerAvecSeparateur(texte1 As String, texte2 As String, Optional separateur As String = " ") As String
    ' Concatène deux textes avec un séparateur
    ConcatenerAvecSeparateur = texte1 & separateur & texte2
End Function
```

---

## 8. Gestion d'erreurs

### Gestion d'erreur de base

```vba
Sub AvecGestionErreur()
    On Error GoTo GestionErreur

    ' Code qui peut générer une erreur
    Dim resultat As Double
    resultat = 10 / 0 ' Division par zéro

    MsgBox "Résultat: " & resultat
    Exit Sub

GestionErreur:
    MsgBox "Erreur " & Err.Number & ": " & Err.Description
    Resume Next ' Continue à la ligne suivante
End Sub
```

### Vérifications avant traitement

```vba
Sub TraitementSecurise()
    Dim valeur1 As String, valeur2 As String
    Dim nombre1 As Double, nombre2 As Double

    valeur1 = Range("A1").Value
    valeur2 = Range("B1").Value

    ' Vérifier que ce sont des nombres
    If IsNumeric(valeur1) And IsNumeric(valeur2) Then
        nombre1 = CDbl(valeur1)
        nombre2 = CDbl(valeur2)

        ' Vérifier la division par zéro
        If nombre2 <> 0 Then
            Range("C1").Value = nombre1 / nombre2
        Else
            Range("C1").Value = "Division par zéro impossible"
        End If
    Else
        MsgBox "Veuillez entrer des nombres valides en A1 et B1"
    End If
End Sub
```

---

## 9. Utilitaires pratiques

### Nettoyer des données

```vba
Sub NettoyerDonnees()
    Dim cellule As Range

    ' Nettoyer une plage de cellules
    For Each cellule In Range("A1:A100")
        If cellule.Value <> "" Then
            ' Supprimer les espaces en trop
            cellule.Value = Trim(cellule.Value)

            ' Mettre en forme correcte (première lettre majuscule)
            cellule.Value = StrConv(cellule.Value, vbProperCase)
        End If
    Next cellule
End Sub
```

### Créer une liste déroulante

```vba
Sub CreerListeDeroulante()
    Dim plageValidation As Range
    Set plageValidation = Range("D1:D10")

    With plageValidation.Validation
        .Delete ' Supprimer validation existante
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Oui,Non,En cours"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Erreur de saisie"
        .ErrorMessage = "Veuillez choisir une valeur dans la liste"
    End With
End Sub
```

---

## 10. Conseils d'utilisation

### Comment adapter ces exemples :

1. **Changez les références** : Remplacez "A1", "B1" par vos cellules
2. **Modifiez les plages** : Adaptez "A1:C10" à vos données
3. **Personnalisez les messages** : Changez les textes des MsgBox
4. **Ajustez les conditions** : Modifiez les If selon vos besoins

### Bonnes pratiques :

```vba
Sub ExempleBonnesPratiques()
    ' Toujours déclarer les variables
    Dim i As Long
    Dim derniereLigne As Long
    Dim feuille As Worksheet

    ' Désactiver les calculs pour la rapidité
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ' Définir la feuille de travail
    Set feuille = Worksheets("Feuil1")

    ' Votre code ici
    derniereLigne = feuille.Range("A" & Rows.Count).End(xlUp).Row

    For i = 1 To derniereLigne
        ' Traitement...
    Next i

    ' Réactiver à la fin
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Libérer les objets
    Set feuille = Nothing
End Sub
```

### Structure de base pour vos macros :

```vba
Sub ModeleBase()
    ' 1. Déclaration des variables
    Dim i As Long

    ' 2. Gestion d'erreur
    On Error GoTo GestionErreur

    ' 3. Désactivation des alertes si nécessaire
    Application.DisplayAlerts = False

    ' 4. Votre code principal ici

    ' 5. Nettoyage et réactivation
    Application.DisplayAlerts = True
    Exit Sub

GestionErreur:
    Application.DisplayAlerts = True
    MsgBox "Erreur: " & Err.Description
End Sub
```

**Conseil final :** Ces exemples sont des points de départ. N'hésitez pas à les combiner, les modifier et les adapter selon vos besoins spécifiques. Plus vous pratiquerez avec ces "briques" de code, plus vous deviendrez autonome en VBA !

⏭️
