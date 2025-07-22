🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.5 FileDialog pour sélection de fichiers

## Introduction

FileDialog est un objet VBA moderne qui permet d'afficher des boîtes de dialogue familières pour que les utilisateurs puissent sélectionner des fichiers ou des dossiers. Cette interface conviviale remplace les anciennes méthodes et offre une expérience utilisateur professionnelle, similaire à celle que l'on trouve dans toutes les applications Windows.

## Comprendre l'objet FileDialog

### Avantages de FileDialog
- **Interface familière** : Utilise les boîtes de dialogue standard de Windows
- **Filtres automatiques** : Peut limiter l'affichage à certains types de fichiers
- **Sélection multiple** : Permet de choisir plusieurs fichiers à la fois
- **Navigation intuitive** : L'utilisateur peut naviguer dans toute l'arborescence
- **Pas de saisie manuelle** : Évite les erreurs de frappe dans les chemins

### Types de FileDialog disponibles
- **msoFileDialogOpen** : Ouvrir un ou plusieurs fichiers
- **msoFileDialogSaveAs** : Sauvegarder un fichier
- **msoFileDialogFilePicker** : Sélectionner des fichiers (sans les ouvrir)
- **msoFileDialogFolderPicker** : Sélectionner un dossier

## Syntaxe de base

### Créer et utiliser un FileDialog

```vba
Sub ExempleBase()
    Dim DialogueFichier As FileDialog

    ' Créer le dialogue
    Set DialogueFichier = Application.FileDialog(msoFileDialogOpen)

    ' Afficher le dialogue
    If DialogueFichier.Show = -1 Then
        ' L'utilisateur a cliqué sur OK
        MsgBox "Fichier sélectionné : " & DialogueFichier.SelectedItems(1)
    Else
        ' L'utilisateur a cliqué sur Annuler
        MsgBox "Aucun fichier sélectionné"
    End If

    ' Nettoyer l'objet
    Set DialogueFichier = Nothing
End Sub
```

## Sélectionner un seul fichier

### Ouvrir un fichier Excel

```vba
Sub OuvrirFichierExcel()
    Dim DialogueFichier As FileDialog
    Dim CheminFichier As String

    ' Créer le dialogue d'ouverture
    Set DialogueFichier = Application.FileDialog(msoFileDialogOpen)

    With DialogueFichier
        ' Configurer le dialogue
        .Title = "Sélectionner un fichier Excel"
        .InitialFileName = "C:\Mes Documents\"

        ' Ajouter des filtres
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx;*.xls"
        .Filters.Add "Tous les fichiers", "*.*"

        ' Interdire la sélection multiple
        .AllowMultiSelect = False

        ' Afficher le dialogue
        If .Show = -1 Then
            CheminFichier = .SelectedItems(1)
            MsgBox "Vous avez sélectionné : " & CheminFichier

            ' Ici, vous pourriez ouvrir le fichier
            ' Workbooks.Open CheminFichier
        Else
            MsgBox "Aucun fichier sélectionné"
        End If
    End With

    Set DialogueFichier = Nothing
End Sub
```

### Sélectionner un fichier texte pour traitement

```vba
Sub SelectionnerFichierTexte()
    Dim DialogueFichier As FileDialog
    Dim CheminFichier As String
    Dim NumFichier As Integer
    Dim ContenuFichier As String

    Set DialogueFichier = Application.FileDialog(msoFileDialogFilePicker)

    With DialogueFichier
        .Title = "Choisir un fichier texte à analyser"
        .InitialFileName = "C:\Temp\"

        ' Filtres pour fichiers texte
        .Filters.Clear
        .Filters.Add "Fichiers texte", "*.txt;*.csv;*.log"
        .Filters.Add "Fichiers CSV", "*.csv"
        .Filters.Add "Tous les fichiers", "*.*"

        .AllowMultiSelect = False

        If .Show = -1 Then
            CheminFichier = .SelectedItems(1)

            ' Lire et afficher le contenu
            NumFichier = FreeFile
            Open CheminFichier For Input As #NumFichier
            ContenuFichier = Input(LOF(NumFichier), NumFichier)
            Close #NumFichier

            MsgBox "Contenu du fichier :" & vbCrLf & Left(ContenuFichier, 200) & "..."
        End If
    End With

    Set DialogueFichier = Nothing
End Sub
```

## Sélectionner plusieurs fichiers

### Traitement par lots de fichiers

```vba
Sub TraiterPlusieursFichiers()
    Dim DialogueFichier As FileDialog
    Dim i As Integer
    Dim NombreFichiers As Integer

    Set DialogueFichier = Application.FileDialog(msoFileDialogFilePicker)

    With DialogueFichier
        .Title = "Sélectionner plusieurs fichiers à traiter"
        .InitialFileName = "C:\Temp\"

        ' Configurer pour sélection multiple
        .AllowMultiSelect = True

        ' Filtres pour documents Office
        .Filters.Clear
        .Filters.Add "Documents Office", "*.xlsx;*.docx;*.pptx"
        .Filters.Add "Fichiers Excel", "*.xlsx;*.xls"
        .Filters.Add "Tous les fichiers", "*.*"

        If .Show = -1 Then
            NombreFichiers = .SelectedItems.Count
            MsgBox "Vous avez sélectionné " & NombreFichiers & " fichier(s)"

            ' Parcourir tous les fichiers sélectionnés
            For i = 1 To NombreFichiers
                Debug.Print "Fichier " & i & " : " & .SelectedItems(i)

                ' Ici vous pourriez traiter chaque fichier
                ' TraiterFichier .SelectedItems(i)
            Next i

            MsgBox "Traitement terminé pour " & NombreFichiers & " fichier(s)"
        End If
    End With

    Set DialogueFichier = Nothing
End Sub
```

### Copier plusieurs fichiers vers un dossier

```vba
Sub CopierFichiersSelectionnes()
    Dim DialogueFichier As FileDialog
    Dim DialogueDossier As FileDialog
    Dim fso As Object
    Dim i As Integer
    Dim DossierDestination As String

    ' D'abord sélectionner les fichiers
    Set DialogueFichier = Application.FileDialog(msoFileDialogFilePicker)

    With DialogueFichier
        .Title = "Sélectionner les fichiers à copier"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Tous les fichiers", "*.*"

        If .Show = -1 Then
            ' Ensuite sélectionner le dossier de destination
            Set DialogueDossier = Application.FileDialog(msoFileDialogFolderPicker)

            With DialogueDossier
                .Title = "Choisir le dossier de destination"

                If .Show = -1 Then
                    DossierDestination = .SelectedItems(1)

                    ' Créer l'objet FileSystemObject pour la copie
                    Set fso = CreateObject("Scripting.FileSystemObject")

                    ' Copier chaque fichier
                    For i = 1 To DialogueFichier.SelectedItems.Count
                        On Error Resume Next
                        fso.CopyFile DialogueFichier.SelectedItems(i), DossierDestination & "\"
                        If Err.Number = 0 Then
                            Debug.Print "Copié : " & fso.GetFileName(DialogueFichier.SelectedItems(i))
                        Else
                            Debug.Print "Erreur : " & fso.GetFileName(DialogueFichier.SelectedItems(i))
                        End If
                        On Error GoTo 0
                    Next i

                    MsgBox "Copie terminée vers : " & DossierDestination
                End If
            End With
        End If
    End With

    Set DialogueFichier = Nothing
    Set DialogueDossier = Nothing
    Set fso = Nothing
End Sub
```

## Sélectionner des dossiers

### Choisir un dossier de travail

```vba
Sub ChoisirDossierTravail()
    Dim DialogueDossier As FileDialog
    Dim DossierChoisi As String

    Set DialogueDossier = Application.FileDialog(msoFileDialogFolderPicker)

    With DialogueDossier
        .Title = "Sélectionner le dossier de travail"
        .InitialFileName = "C:\"

        If .Show = -1 Then
            DossierChoisi = .SelectedItems(1)
            MsgBox "Dossier sélectionné : " & DossierChoisi

            ' Vous pourriez sauvegarder ce chemin pour l'utiliser plus tard
            ' ThisWorkbook.Worksheets("Config").Range("B1") = DossierChoisi
        Else
            MsgBox "Aucun dossier sélectionné"
        End If
    End With

    Set DialogueDossier = Nothing
End Sub
```

### Créer une sauvegarde dans un dossier choisi

```vba
Sub SauvegarderDansDossierChoisi()
    Dim DialogueDossier As FileDialog
    Dim DossierSauvegarde As String
    Dim NomSauvegarde As String
    Dim CheminComplet As String

    Set DialogueDossier = Application.FileDialog(msoFileDialogFolderPicker)

    With DialogueDossier
        .Title = "Choisir l'emplacement de sauvegarde"
        .InitialFileName = "C:\Mes Documents\"

        If .Show = -1 Then
            DossierSauvegarde = .SelectedItems(1)

            ' Créer un nom de fichier avec la date
            NomSauvegarde = "Sauvegarde_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".xlsx"
            CheminComplet = DossierSauvegarde & "\" & NomSauvegarde

            ' Sauvegarder le classeur actuel
            On Error GoTo ErreurSauvegarde
            ThisWorkbook.SaveCopyAs CheminComplet
            MsgBox "Sauvegarde créée : " & CheminComplet

            Exit Sub

ErreurSauvegarde:
            MsgBox "Erreur lors de la sauvegarde : " & Err.Description
        End If
    End With

    Set DialogueDossier = Nothing
End Sub
```

## Dialogue de sauvegarde

### Sauvegarder avec un nom personnalisé

```vba
Sub SauvegarderSous()
    Dim DialogueSauvegarde As FileDialog
    Dim CheminSauvegarde As String

    Set DialogueSauvegarde = Application.FileDialog(msoFileDialogSaveAs)

    With DialogueSauvegarde
        .Title = "Enregistrer sous..."
        .InitialFileName = "C:\Mes Documents\MonRapport"

        ' Filtres pour la sauvegarde
        .Filters.Clear
        .Filters.Add "Classeur Excel", "*.xlsx"
        .Filters.Add "Classeur Excel 97-2003", "*.xls"
        .Filters.Add "Fichier CSV", "*.csv"

        If .Show = -1 Then
            CheminSauvegarde = .SelectedItems(1)

            ' Sauvegarder le classeur
            On Error GoTo ErreurSauvegarde
            ThisWorkbook.SaveAs CheminSauvegarde
            MsgBox "Fichier sauvegardé : " & CheminSauvegarde

            Exit Sub

ErreurSauvegarde:
            MsgBox "Impossible de sauvegarder : " & Err.Description
        End If
    End With

    Set DialogueSauvegarde = Nothing
End Sub
```

### Exporter des données vers un fichier

```vba
Sub ExporterDonnees()
    Dim DialogueSauvegarde As FileDialog
    Dim CheminExport As String
    Dim ws As Worksheet

    Set DialogueSauvegarde = Application.FileDialog(msoFileDialogSaveAs)
    Set ws = ActiveSheet

    With DialogueSauvegarde
        .Title = "Exporter les données vers..."
        .InitialFileName = "C:\Exports\Donnees_" & Format(Date, "yyyy-mm-dd")

        .Filters.Clear
        .Filters.Add "Fichier CSV", "*.csv"
        .Filters.Add "Fichier texte", "*.txt"
        .Filters.Add "Classeur Excel", "*.xlsx"

        If .Show = -1 Then
            CheminExport = .SelectedItems(1)

            ' Déterminer le format selon l'extension
            If LCase(Right(CheminExport, 4)) = ".csv" Then
                ' Exporter en CSV
                ws.Copy
                ActiveWorkbook.SaveAs CheminExport, xlCSV
                ActiveWorkbook.Close False
                MsgBox "Données exportées en CSV : " & CheminExport
            ElseIf LCase(Right(CheminExport, 5)) = ".xlsx" Then
                ' Exporter en Excel
                ws.Copy
                ActiveWorkbook.SaveAs CheminExport, xlOpenXMLWorkbook
                ActiveWorkbook.Close False
                MsgBox "Données exportées en Excel : " & CheminExport
            End If
        End If
    End With

    Set DialogueSauvegarde = Nothing
    Set ws = Nothing
End Sub
```

## Configuration avancée des filtres

### Créer des filtres personnalisés

```vba
Sub FiltresPersonnalises()
    Dim DialogueFichier As FileDialog

    Set DialogueFichier = Application.FileDialog(msoFileDialogOpen)

    With DialogueFichier
        .Title = "Ouvrir un document"

        ' Vider les filtres existants
        .Filters.Clear

        ' Ajouter des filtres spécifiques
        .Filters.Add "Images", "*.jpg;*.jpeg;*.png;*.gif;*.bmp"
        .Filters.Add "Documents Word", "*.docx;*.doc"
        .Filters.Add "Feuilles de calcul", "*.xlsx;*.xls;*.csv"
        .Filters.Add "Présentations", "*.pptx;*.ppt"
        .Filters.Add "Archives", "*.zip;*.rar;*.7z"
        .Filters.Add "Tous les fichiers", "*.*"

        ' Définir le filtre par défaut (le premier = 1)
        .FilterIndex = 3  ' Feuilles de calcul par défaut

        .AllowMultiSelect = False

        If .Show = -1 Then
            MsgBox "Fichier sélectionné : " & .SelectedItems(1)
        End If
    End With

    Set DialogueFichier = Nothing
End Sub
```

### Filtres dynamiques basés sur la date

```vba
Sub FiltresDynamiques()
    Dim DialogueFichier As FileDialog
    Dim FiltreDuJour As String
    Dim FiltreDuMois As String

    ' Créer des filtres basés sur la date actuelle
    FiltreDuJour = "*" & Format(Date, "yyyy-mm-dd") & "*.*"
    FiltreDuMois = "*" & Format(Date, "yyyy-mm") & "*.*"

    Set DialogueFichier = Application.FileDialog(msoFileDialogFilePicker)

    With DialogueFichier
        .Title = "Sélectionner des fichiers récents"

        .Filters.Clear
        .Filters.Add "Fichiers d'aujourd'hui", FiltreDuJour
        .Filters.Add "Fichiers de ce mois", FiltreDuMois
        .Filters.Add "Rapports", "*rapport*.*"
        .Filters.Add "Tous les fichiers", "*.*"

        .AllowMultiSelect = True

        If .Show = -1 Then
            MsgBox "Nombre de fichiers sélectionnés : " & .SelectedItems.Count
        End If
    End With

    Set DialogueFichier = Nothing
End Sub
```

## Fonctions utilitaires avec FileDialog

### Fonction pour sélectionner un fichier Excel

```vba
Function SelectionnerFichierExcel() As String
    Dim DialogueFichier As FileDialog
    Dim CheminFichier As String

    Set DialogueFichier = Application.FileDialog(msoFileDialogOpen)

    With DialogueFichier
        .Title = "Sélectionner un fichier Excel"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx;*.xls;*.xlsm"
        .AllowMultiSelect = False

        If .Show = -1 Then
            CheminFichier = .SelectedItems(1)
        Else
            CheminFichier = ""  ' Aucun fichier sélectionné
        End If
    End With

    Set DialogueFichier = Nothing
    SelectionnerFichierExcel = CheminFichier
End Function

Sub UtiliserFonctionSelection()
    Dim MonFichier As String

    MonFichier = SelectionnerFichierExcel()

    If MonFichier <> "" Then
        MsgBox "Vous avez choisi : " & MonFichier
        ' Traiter le fichier ici
    Else
        MsgBox "Aucun fichier sélectionné"
    End If
End Sub
```

### Fonction pour sélectionner un dossier

```vba
Function SelectionnerDossier(Titre As String) As String
    Dim DialogueDossier As FileDialog
    Dim CheminDossier As String

    Set DialogueDossier = Application.FileDialog(msoFileDialogFolderPicker)

    With DialogueDossier
        .Title = Titre

        If .Show = -1 Then
            CheminDossier = .SelectedItems(1)
        Else
            CheminDossier = ""
        End If
    End With

    Set DialogueDossier = Nothing
    SelectionnerDossier = CheminDossier
End Function

Sub ExempleUtilisationDossier()
    Dim DossierSource As String
    Dim DossierDestination As String

    DossierSource = SelectionnerDossier("Choisir le dossier source")
    If DossierSource = "" Then Exit Sub

    DossierDestination = SelectionnerDossier("Choisir le dossier de destination")
    If DossierDestination = "" Then Exit Sub

    MsgBox "Source : " & DossierSource & vbCrLf & "Destination : " & DossierDestination
End Sub
```

## Mémoriser les derniers emplacements

### Sauvegarder le dernier dossier utilisé

```vba
Sub SauvegarderDernierEmplacement()
    Dim DialogueFichier As FileDialog
    Dim CheminFichier As String
    Dim DernierDossier As String

    ' Récupérer le dernier dossier utilisé (sauvegardé dans une cellule)
    On Error Resume Next
    DernierDossier = ThisWorkbook.Worksheets("Config").Range("A1").Value
    If DernierDossier = "" Then DernierDossier = "C:\"
    On Error GoTo 0

    Set DialogueFichier = Application.FileDialog(msoFileDialogOpen)

    With DialogueFichier
        .Title = "Ouvrir un fichier"
        .InitialFileName = DernierDossier
        .Filters.Clear
        .Filters.Add "Tous les fichiers", "*.*"
        .AllowMultiSelect = False

        If .Show = -1 Then
            CheminFichier = .SelectedItems(1)

            ' Extraire et sauvegarder le dossier pour la prochaine fois
            DernierDossier = Left(CheminFichier, InStrRev(CheminFichier, "\"))

            On Error Resume Next
            ThisWorkbook.Worksheets("Config").Range("A1").Value = DernierDossier
            On Error GoTo 0

            MsgBox "Fichier sélectionné : " & CheminFichier
        End If
    End With

    Set DialogueFichier = Nothing
End Sub
```

## Bonnes pratiques avec FileDialog

### 1. Toujours nettoyer les objets

```vba
Sub BonnePratique()
    Dim DialogueFichier As FileDialog

    Set DialogueFichier = Application.FileDialog(msoFileDialogOpen)

    ' Utiliser le dialogue
    ' ...

    ' Toujours nettoyer à la fin
    Set DialogueFichier = Nothing
End Sub
```

### 2. Gérer le cas d'annulation

```vba
If DialogueFichier.Show = -1 Then
    ' L'utilisateur a validé
    ' Traitement normal
Else
    ' L'utilisateur a annulé
    MsgBox "Opération annulée"
    Exit Sub
End If
```

### 3. Vérifier les sélections multiples

```vba
If .AllowMultiSelect = True Then
    If .SelectedItems.Count > 0 Then
        For i = 1 To .SelectedItems.Count
            ' Traiter chaque fichier
        Next i
    End If
End If
```

### 4. Utiliser des titres explicites

```vba
.Title = "Sélectionner le fichier de données à importer"
' Au lieu de simplement "Ouvrir"
```

## Limitations et alternatives

### Limitations de FileDialog
- Nécessite Excel (pas disponible dans VBA autonome)
- Interface parfois moins flexible que les API Windows
- Filtres limités par rapport aux possibilités du système

### Alternative simple : InputBox

```vba
Sub AlternativeInputBox()
    Dim CheminFichier As String

    CheminFichier = InputBox("Entrez le chemin complet du fichier :", _
                            "Sélection de fichier", _
                            "C:\Temp\MonFichier.txt")

    If CheminFichier <> "" Then
        If Dir(CheminFichier) <> "" Then
            MsgBox "Fichier trouvé : " & CheminFichier
        Else
            MsgBox "Fichier introuvable !"
        End If
    End If
End Sub
```

---

*FileDialog offre une interface moderne et conviviale pour la sélection de fichiers et dossiers. Maîtriser cet outil vous permettra de créer des applications VBA avec une expérience utilisateur professionnelle et intuitive.*

⏭️
