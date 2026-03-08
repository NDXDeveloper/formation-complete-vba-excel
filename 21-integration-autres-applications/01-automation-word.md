🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.1 Automation avec Word

## Introduction à Word Automation

L'automation avec Word permet de créer, modifier et contrôler des documents Word directement depuis Excel avec VBA. C'est particulièrement utile pour générer des rapports, créer des lettres automatiques ou effectuer du publipostage.

## Première étape : Créer une connexion avec Word

### Méthode simple pour débuter

```vba
Sub PremierTestWord()
    ' Créer une connexion avec Word
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")

    ' Rendre Word visible
    wordApp.Visible = True

    ' Créer un nouveau document
    wordApp.Documents.Add

    ' Écrire du texte
    wordApp.Selection.TypeText "Bonjour depuis Excel !"

    ' Important : Fermer proprement
    wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub
```

**Explication ligne par ligne :**
- `Dim wordApp As Object` : Déclare une variable pour contenir Word
- `Set wordApp = CreateObject("Word.Application")` : Lance Word
- `wordApp.Visible = True` : Rend Word visible à l'écran
- `wordApp.Documents.Add` : Crée un nouveau document vide
- `wordApp.Selection.TypeText` : Écrit du texte dans le document
- `wordApp.Quit` : Ferme Word complètement
- `Set wordApp = Nothing` : Libère la mémoire

## Comprendre la hiérarchie des objets Word

Word est organisé comme une hiérarchie d'objets :

```
Application (Word lui-même)
├── Documents (Collection de tous les documents)
│   └── Document (Un document spécifique)
│       └── Range (Une portion de texte)
│           └── Characters, Words, Sentences, Paragraphs
```

### Exemple pratique de cette hiérarchie

```vba
Sub ExempleHierarchie()
    Dim wordApp As Object
    Dim monDoc As Object

    ' 1. Créer l'application Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True

    ' 2. Créer un document
    Set monDoc = wordApp.Documents.Add

    ' 3. Travailler avec le contenu
    monDoc.Range.Text = "Ceci est mon premier document automatisé !"

    ' Nettoyer
    wordApp.Quit SaveChanges:=False
    Set monDoc = Nothing
    Set wordApp = Nothing
End Sub
```

## Écrire et formater du texte

### Ajout de texte simple

```vba
Sub AjouterTexte()
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True

    Dim doc As Object
    Set doc = wordApp.Documents.Add

    ' Différentes façons d'ajouter du texte

    ' Méthode 1 : Avec Selection
    wordApp.Selection.TypeText "Première ligne" & vbCrLf

    ' Méthode 2 : Avec Range
    doc.Range.InsertAfter "Deuxième ligne" & vbCrLf

    ' Méthode 3 : Définir tout le contenu d'un coup
    doc.Range.Text = doc.Range.Text & "Troisième ligne"

    wordApp.Quit SaveChanges:=False
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

### Formatage basique du texte

```vba
Sub FormaterTexte()
    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Ajouter du texte
    wordApp.Selection.TypeText "Texte à formater"

    ' Sélectionner tout le texte
    wordApp.Selection.WholeStory

    ' Appliquer du formatage
    With wordApp.Selection.Font
        .Name = "Arial"
        .Size = 14
        .Bold = True
        .Color = RGB(255, 0, 0)  ' Rouge
    End With

    wordApp.Quit SaveChanges:=False
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Travailler avec des paragraphes

```vba
Sub GererParagraphes()
    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Ajouter plusieurs paragraphes
    doc.Range.Text = "Premier paragraphe" & vbCrLf & _
                     "Deuxième paragraphe" & vbCrLf & _
                     "Troisième paragraphe"

    ' Formater le premier paragraphe
    With doc.Paragraphs(1).Range.Font
        .Size = 16
        .Bold = True
    End With

    ' Centrer le deuxième paragraphe
    doc.Paragraphs(2).Alignment = 1  ' 1 = Centré

    wordApp.Quit SaveChanges:=False
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Insérer des données Excel dans Word

### Exemple simple : Transférer une valeur

```vba
Sub TransfererDonneeSimple()
    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Récupérer une valeur depuis Excel
    Dim valeurExcel As String
    valeurExcel = Range("A1").Value

    ' L'insérer dans Word
    doc.Range.Text = "Valeur depuis Excel : " & valeurExcel

    wordApp.Quit SaveChanges:=False
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

### Exemple plus complexe : Créer un tableau

```vba
Sub CreerTableauWord()
    Dim wordApp As Object
    Dim doc As Object
    Dim tableau As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Créer un tableau 3x3
    Set tableau = doc.Tables.Add(doc.Range, 3, 3)

    ' Remplir les en-têtes
    tableau.Cell(1, 1).Range.Text = "Nom"
    tableau.Cell(1, 2).Range.Text = "Âge"
    tableau.Cell(1, 3).Range.Text = "Ville"

    ' Remplir avec des données Excel (supposons qu'elles sont en A2:C3)
    Dim i As Integer
    For i = 1 To 2
        tableau.Cell(i + 1, 1).Range.Text = Cells(i + 1, 1).Value  ' Colonne A
        tableau.Cell(i + 1, 2).Range.Text = Cells(i + 1, 2).Value  ' Colonne B
        tableau.Cell(i + 1, 3).Range.Text = Cells(i + 1, 3).Value  ' Colonne C
    Next i

    ' Formater les en-têtes
    tableau.Rows(1).Range.Font.Bold = True

    wordApp.Quit SaveChanges:=False
    Set tableau = Nothing
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Ouvrir et modifier un document existant

```vba
Sub ModifierDocumentExistant()
    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True

    ' Ouvrir un document existant
    ' Changez le chemin vers votre document
    Set doc = wordApp.Documents.Open("C:\MonDossier\MonDocument.docx")

    ' Aller à la fin du document
    wordApp.Selection.EndKey 6  ' 6 = fin du document

    ' Ajouter du nouveau contenu
    wordApp.Selection.TypeText vbCrLf & "Ajouté par Excel le : " & Now()

    ' Sauvegarder
    doc.Save

    wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Rechercher et remplacer du texte

```vba
Sub RechercherRemplacer()
    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Ajouter du texte avec des marqueurs
    doc.Range.Text = "Bonjour [NOM], votre commande [NUMERO] est prête."

    ' Remplacer les marqueurs par des vraies valeurs
    With wordApp.Selection.Find
        .ClearFormatting
        .Text = "[NOM]"
        .Replacement.Text = "Jean Dupont"
        .Execute Replace:=2  ' 2 = Remplacer tout
    End With

    With wordApp.Selection.Find
        .ClearFormatting
        .Text = "[NUMERO]"
        .Replacement.Text = "12345"
        .Execute Replace:=2
    End With

    wordApp.Quit SaveChanges:=False
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Sauvegarder un document

```vba
Sub SauvegarderDocument()
    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Ajouter du contenu
    doc.Range.Text = "Document créé automatiquement le " & Date

    ' Sauvegarder avec un nom spécifique
    doc.SaveAs2 "C:\MonDossier\Rapport_" & Format(Date, "yyyy-mm-dd") & ".docx"

    ' Alternative : Sauvegarder en PDF
    ' doc.SaveAs2 "C:\MonDossier\Rapport.pdf", 17  ' 17 = Format PDF

    wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Fermer Word proprement

Il est très important de fermer Word correctement pour éviter que des instances restent ouvertes en arrière-plan.

```vba
Sub FermetureCorrecte()
    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Votre travail ici...
    doc.Range.Text = "Travail terminé"

    ' Fermeture correcte
    doc.Close SaveChanges:=True  ' Fermer le document en sauvegardant
    wordApp.Quit                 ' Fermer Word complètement

    ' Libérer la mémoire
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Gestion d'erreurs pour Word Automation

```vba
Sub AvecGestionErreurs()
    Dim wordApp As Object
    Dim doc As Object

    On Error GoTo GestionErreur

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Votre code ici...
    doc.Range.Text = "Test avec gestion d'erreurs"

    ' Nettoyage normal
    doc.Close SaveChanges:=False
    wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing

    Exit Sub

GestionErreur:
    MsgBox "Erreur : " & Err.Description

    ' Nettoyage en cas d'erreur
    If Not doc Is Nothing Then
        doc.Close SaveChanges:=False
        Set doc = Nothing
    End If

    If Not wordApp Is Nothing Then
        wordApp.Quit
        Set wordApp = Nothing
    End If
End Sub
```

## Exemple complet : Générer un rapport automatique

```vba
Sub GenererRapportComplet()
    Dim wordApp As Object
    Dim doc As Object

    ' Initialisation
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' Titre du document
    wordApp.Selection.Font.Size = 18
    wordApp.Selection.Font.Bold = True
    wordApp.Selection.TypeText "RAPPORT MENSUEL" & vbCrLf & vbCrLf

    ' Date de création
    wordApp.Selection.Font.Size = 12
    wordApp.Selection.Font.Bold = False
    wordApp.Selection.TypeText "Créé le : " & Date & vbCrLf & vbCrLf

    ' Section des données
    wordApp.Selection.Font.Bold = True
    wordApp.Selection.TypeText "Résumé des ventes :" & vbCrLf
    wordApp.Selection.Font.Bold = False

    ' Supposons des données en Excel A1:B5
    Dim i As Integer
    For i = 1 To 5
        wordApp.Selection.TypeText "• " & Cells(i, 1).Value & " : " & Cells(i, 2).Value & vbCrLf
    Next i

    ' Pied de page
    wordApp.Selection.TypeText vbCrLf & "Rapport généré automatiquement depuis Excel"

    ' Sauvegarder
    doc.SaveAs2 Environ("USERPROFILE") & "\Desktop\Rapport_" & Format(Date, "yyyy-mm-dd") & ".docx"

    MsgBox "Rapport créé sur le Bureau !"

    ' Nettoyage
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
```

## Points importants à retenir

### ✅ Bonnes pratiques
- Toujours utiliser `Set variable = Nothing` pour libérer la mémoire
- Fermer Word avec `wordApp.Quit` quand c'est terminé
- Utiliser la gestion d'erreurs pour éviter les plantages
- Tester le code sur de petits exemples avant de l'appliquer à de gros documents

### ⚠️ Erreurs courantes à éviter
- Oublier de fermer Word (instances multiples en arrière-plan)
- Ne pas gérer les erreurs
- Modifier directement `Selection` sans vérifier qu'elle existe
- Utiliser des chemins de fichiers en dur (préférer des variables)

### 💡 Conseils pour débuter
- Commencez par des exemples simples
- Utilisez `wordApp.Visible = True` pour voir ce qui se passe
- Testez chaque étape individuellement
- Gardez la documentation Word VBA à portée de main

L'automation avec Word offre des possibilités infinies pour automatiser la création de documents. Une fois ces bases maîtrisées, vous pourrez créer des solutions sophistiquées de génération de rapports et de publipostage.

⏭️
