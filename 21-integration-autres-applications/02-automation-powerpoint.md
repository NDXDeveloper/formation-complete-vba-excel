🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.2 Automation avec PowerPoint

## Introduction à PowerPoint Automation

L'automation avec PowerPoint permet de créer et modifier des présentations directement depuis Excel avec VBA. C'est idéal pour générer automatiquement des présentations avec des graphiques Excel, créer des rapports visuels ou mettre à jour des slides avec de nouvelles données.

## Première étape : Créer une connexion avec PowerPoint

### Méthode simple pour débuter

```vba
Sub PremierTestPowerPoint()
    ' Créer une connexion avec PowerPoint
    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")

    ' Rendre PowerPoint visible
    pptApp.Visible = True

    ' Créer une nouvelle présentation
    Dim pres As Object
    Set pres = pptApp.Presentations.Add

    ' Ajouter une diapositive
    Dim slide As Object
    Set slide = pres.Slides.Add(1, 1)  ' 1 = position, 1 = mise en page titre

    ' Ajouter du texte au titre
    slide.Shapes(1).TextFrame.TextRange.Text = "Ma première diapositive automatisée !"

    ' Important : Fermer proprement
    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

**Explication ligne par ligne :**
- `CreateObject("PowerPoint.Application")` : Lance PowerPoint
- `pptApp.Visible = True` : Rend PowerPoint visible à l'écran
- `pptApp.Presentations.Add` : Crée une nouvelle présentation vide
- `pres.Slides.Add(1, 1)` : Ajoute une diapositive à la position 1 avec une mise en page titre
- `slide.Shapes(1).TextFrame.TextRange.Text` : Modifie le texte du premier élément (titre)

## Comprendre la hiérarchie des objets PowerPoint

PowerPoint est organisé comme une hiérarchie d'objets :

```
Application (PowerPoint lui-même)
├── Presentations (Collection de toutes les présentations)
│   └── Presentation (Une présentation spécifique)
│       └── Slides (Collection des diapositives)
│           └── Slide (Une diapositive spécifique)
│               └── Shapes (Tous les objets sur la diapositive)
│                   └── Shape (Texte, image, graphique, etc.)
```

### Exemple pratique de cette hiérarchie

```vba
Sub ExempleHierarchiePowerPoint()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object

    ' 1. Créer l'application PowerPoint
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' 2. Créer une présentation
    Set pres = pptApp.Presentations.Add

    ' 3. Ajouter une diapositive
    Set slide = pres.Slides.Add(1, 1)

    ' 4. Modifier les éléments de la diapositive
    slide.Shapes(1).TextFrame.TextRange.Text = "Titre principal"
    slide.Shapes(2).TextFrame.TextRange.Text = "Sous-titre de ma présentation"

    ' Nettoyer
    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Types de mises en page (Layouts) courantes

PowerPoint propose différentes mises en page prédéfinies :

```vba
Sub DifferentesMisesEnPage()
    Dim pptApp As Object
    Dim pres As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' Différents types de mises en page
    ' 1 = Titre seul
    pres.Slides.Add(1, 1).Shapes(1).TextFrame.TextRange.Text = "Diapositive de titre"

    ' 2 = Titre + contenu
    Dim slide2 As Object
    Set slide2 = pres.Slides.Add(2, 2)
    slide2.Shapes(1).TextFrame.TextRange.Text = "Titre avec contenu"
    slide2.Shapes(2).TextFrame.TextRange.Text = "Voici le contenu de ma diapositive"

    ' 3 = Titre + deux contenus
    Dim slide3 As Object
    Set slide3 = pres.Slides.Add(3, 3)
    slide3.Shapes(1).TextFrame.TextRange.Text = "Titre avec deux colonnes"

    ' 11 = Vide (pour créer entièrement sa mise en page)
    pres.Slides.Add(4, 11)

    pptApp.Quit
    Set slide3 = Nothing
    Set slide2 = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Ajouter et formater du texte

### Texte simple dans des zones prédéfinies

```vba
Sub AjouterTexteSimple()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' Créer une diapositive titre + contenu
    Set slide = pres.Slides.Add(1, 2)

    ' Modifier le titre
    slide.Shapes(1).TextFrame.TextRange.Text = "Rapport mensuel"

    ' Modifier le contenu avec plusieurs lignes
    slide.Shapes(2).TextFrame.TextRange.Text = "• Ventes en hausse de 15%" & vbCrLf & _
                                               "• Nouveaux clients : 25" & vbCrLf & _
                                               "• Objectifs dépassés"

    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

### Formatage avancé du texte

```vba
Sub FormaterTexte()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add
    Set slide = pres.Slides.Add(1, 2)

    ' Formater le titre
    With slide.Shapes(1).TextFrame.TextRange.Font
        .Name = "Arial"
        .Size = 32
        .Bold = True
        .Color.RGB = RGB(0, 0, 255)  ' Bleu
    End With
    slide.Shapes(1).TextFrame.TextRange.Text = "Titre formaté"

    ' Formater le contenu
    slide.Shapes(2).TextFrame.TextRange.Text = "Texte du contenu"
    With slide.Shapes(2).TextFrame.TextRange.Font
        .Name = "Calibri"
        .Size = 18
        .Italic = True
        .Color.RGB = RGB(128, 128, 128)  ' Gris
    End With

    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Ajouter des zones de texte personnalisées

```vba
Sub AjouterZoneTextePersonnalisee()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object
    Dim textBox As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add
    Set slide = pres.Slides.Add(1, 11)  ' Mise en page vide

    ' Ajouter une zone de texte personnalisée
    ' AddTextbox(Orientation, Left, Top, Width, Height)
    Set textBox = slide.Shapes.AddTextbox(1, 100, 100, 400, 100)

    ' Ajouter du texte
    textBox.TextFrame.TextRange.Text = "Ceci est une zone de texte personnalisée"

    ' Formater la zone de texte
    With textBox.TextFrame.TextRange.Font
        .Size = 20
        .Bold = True
        .Color.RGB = RGB(255, 0, 0)  ' Rouge
    End With

    pptApp.Quit
    Set textBox = Nothing
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Insérer des données Excel dans PowerPoint

### Transférer des valeurs simples

```vba
Sub TransfererDonneesExcel()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add
    Set slide = pres.Slides.Add(1, 2)

    ' Titre avec données Excel
    slide.Shapes(1).TextFrame.TextRange.Text = "Rapport du " & Range("A1").Value

    ' Contenu avec plusieurs données Excel
    Dim contenu As String
    contenu = "• Total des ventes : " & Range("B1").Value & " €" & vbCrLf
    contenu = contenu & "• Nombre de commandes : " & Range("B2").Value & vbCrLf
    contenu = contenu & "• Moyenne par commande : " & Range("B3").Value & " €"

    slide.Shapes(2).TextFrame.TextRange.Text = contenu

    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

### Créer un tableau avec des données Excel

```vba
Sub CreerTableauPowerPoint()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object
    Dim tableau As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add
    Set slide = pres.Slides.Add(1, 11)  ' Mise en page vide

    ' Créer un tableau 4 lignes x 3 colonnes
    Set tableau = slide.Shapes.AddTable(4, 3, 50, 100, 600, 300)

    ' Remplir les en-têtes (première ligne)
    tableau.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Produit"
    tableau.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Quantité"
    tableau.Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = "Prix"

    ' Remplir avec des données Excel (supposons A2:C4)
    Dim i As Integer
    For i = 1 To 3
        tableau.Table.Cell(i + 1, 1).Shape.TextFrame.TextRange.Text = Cells(i + 1, 1).Value
        tableau.Table.Cell(i + 1, 2).Shape.TextFrame.TextRange.Text = Cells(i + 1, 2).Value
        tableau.Table.Cell(i + 1, 3).Shape.TextFrame.TextRange.Text = Cells(i + 1, 3).Value & " €"
    Next i

    ' Formater les en-têtes
    Dim j As Integer
    For j = 1 To 3
        tableau.Table.Cell(1, j).Shape.TextFrame.TextRange.Font.Bold = True
        tableau.Table.Cell(1, j).Shape.Fill.ForeColor.RGB = RGB(200, 200, 200)
    Next j

    pptApp.Quit
    Set tableau = Nothing
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Copier des graphiques Excel vers PowerPoint

```vba
Sub CopierGraphiqueExcel()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object

    ' Sélectionner et copier le graphique Excel
    ' (Supposons qu'il y a un graphique nommé "Graphique 1" dans la feuille active)
    ActiveSheet.ChartObjects("Graphique 1").Select
    ActiveChart.ChartArea.Copy

    ' Créer PowerPoint
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add
    Set slide = pres.Slides.Add(1, 11)  ' Mise en page vide

    ' Coller le graphique
    slide.Shapes.Paste

    ' Redimensionner et positionner le graphique
    With slide.Shapes(slide.Shapes.Count)  ' Le dernier objet ajouté
        .Left = 100
        .Top = 100
        .Width = 500
        .Height = 400
    End With

    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing

    ' Vider le presse-papiers
    Application.CutCopyMode = False
End Sub
```

## Ajouter des images

```vba
Sub AjouterImage()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object
    Dim image As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add
    Set slide = pres.Slides.Add(1, 11)

    ' Ajouter une image (changez le chemin)
    Set image = slide.Shapes.AddPicture("C:\MonDossier\MonImage.jpg", False, True, 100, 100, 300, 200)

    ' Alternative : utiliser une image depuis les ressources Windows
    ' Set image = slide.Shapes.AddPicture(Environ("WINDIR") & "\Web\Wallpaper\Windows\img0.jpg", False, True, 100, 100, 300, 200)

    pptApp.Quit
    Set image = Nothing
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Travailler avec plusieurs diapositives

```vba
Sub CreerPresentationComplete()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object
    Dim i As Integer

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' Diapositive de titre
    Set slide = pres.Slides.Add(1, 1)
    slide.Shapes(1).TextFrame.TextRange.Text = "Présentation automatisée"
    slide.Shapes(2).TextFrame.TextRange.Text = "Générée depuis Excel le " & Date

    ' Plusieurs diapositives de contenu
    For i = 1 To 3
        Set slide = pres.Slides.Add(i + 1, 2)  ' Position i+1, mise en page titre + contenu
        slide.Shapes(1).TextFrame.TextRange.Text = "Section " & i
        slide.Shapes(2).TextFrame.TextRange.Text = "Contenu de la section " & i & vbCrLf & _
                                                   "• Point 1" & vbCrLf & _
                                                   "• Point 2" & vbCrLf & _
                                                   "• Point 3"
    Next i

    ' Diapositive de conclusion
    Set slide = pres.Slides.Add(5, 1)
    slide.Shapes(1).TextFrame.TextRange.Text = "Merci pour votre attention"
    slide.Shapes(2).TextFrame.TextRange.Text = "Questions ?"

    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Ouvrir et modifier une présentation existante

```vba
Sub ModifierPresentationExistante()
    Dim pptApp As Object
    Dim pres As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' Ouvrir une présentation existante (changez le chemin)
    Set pres = pptApp.Presentations.Open("C:\MonDossier\MaPresentation.pptx")

    ' Modifier la première diapositive
    pres.Slides(1).Shapes(1).TextFrame.TextRange.Text = "Titre modifié le " & Date

    ' Ajouter une nouvelle diapositive à la fin
    Dim nouvelleSlide As Object
    Set nouvelleSlide = pres.Slides.Add(pres.Slides.Count + 1, 2)
    nouvelleSlide.Shapes(1).TextFrame.TextRange.Text = "Nouvelle diapositive"
    nouvelleSlide.Shapes(2).TextFrame.TextRange.Text = "Ajoutée automatiquement"

    ' Sauvegarder
    pres.Save

    pptApp.Quit
    Set nouvelleSlide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Appliquer des thèmes et des styles

```vba
Sub AppliquerTheme()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' Ajouter du contenu
    Set slide = pres.Slides.Add(1, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "Présentation avec thème"
    slide.Shapes(2).TextFrame.TextRange.Text = "Contenu stylisé"

    ' Appliquer un thème (si disponible)
    ' pres.ApplyTemplate "C:\Program Files\Microsoft Office\Templates\Themes\Ion.thmx"

    ' Alternative : définir manuellement les couleurs de thème
    slide.FollowMasterBackground = False
    slide.Background.Fill.Solid
    slide.Background.Fill.ForeColor.RGB = RGB(240, 248, 255)  ' Bleu très clair

    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Sauvegarder et exporter

```vba
Sub SauvegarderPresentation()
    Dim pptApp As Object
    Dim pres As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' Créer du contenu
    Dim slide As Object
    Set slide = pres.Slides.Add(1, 1)
    slide.Shapes(1).TextFrame.TextRange.Text = "Présentation à sauvegarder"

    ' Sauvegarder au format PowerPoint
    pres.SaveAs Environ("USERPROFILE") & "\Desktop\MaPresentation_" & Format(Date, "yyyy-mm-dd") & ".pptx"

    ' Alternative : Exporter en PDF
    pres.SaveAs Environ("USERPROFILE") & "\Desktop\MaPresentation.pdf", 32  ' 32 = Format PDF

    pptApp.Quit
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Fermer PowerPoint proprement

```vba
Sub FermetureCorrectePowerPoint()
    Dim pptApp As Object
    Dim pres As Object

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' Votre travail ici...
    pres.Slides.Add(1, 1).Shapes(1).TextFrame.TextRange.Text = "Travail terminé"

    ' Fermeture correcte
    pres.Close             ' Fermer la présentation
    pptApp.Quit           ' Fermer PowerPoint complètement

    ' Libérer la mémoire
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Exemple complet : Rapport de ventes automatisé

```vba
Sub GenererRapportVentesPowerPoint()
    Dim pptApp As Object
    Dim pres As Object
    Dim slide As Object

    ' Initialisation
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' === DIAPOSITIVE 1 : TITRE ===
    Set slide = pres.Slides.Add(1, 1)
    slide.Shapes(1).TextFrame.TextRange.Text = "RAPPORT DE VENTES"
    slide.Shapes(2).TextFrame.TextRange.Text = "Mois de " & MonthName(Month(Date)) & " " & Year(Date)

    ' Formater le titre
    With slide.Shapes(1).TextFrame.TextRange.Font
        .Size = 36
        .Bold = True
        .Color.RGB = RGB(0, 50, 100)
    End With

    ' === DIAPOSITIVE 2 : RÉSUMÉ EXÉCUTIF ===
    Set slide = pres.Slides.Add(2, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "Résumé exécutif"

    ' Récupérer des données Excel (supposons qu'elles sont dans des cellules spécifiques)
    Dim totalVentes As String
    Dim nbCommandes As String
    Dim moyenneCommande As String

    totalVentes = Format(Range("B1").Value, "#,##0") & " €"
    nbCommandes = Range("B2").Value
    moyenneCommande = Format(Range("B3").Value, "#,##0") & " €"

    Dim resumeTexte As String
    resumeTexte = "• Chiffre d'affaires total : " & totalVentes & vbCrLf
    resumeTexte = resumeTexte & "• Nombre de commandes : " & nbCommandes & vbCrLf
    resumeTexte = resumeTexte & "• Panier moyen : " & moyenneCommande & vbCrLf
    resumeTexte = resumeTexte & "• Objectif mensuel : ATTEINT ✓"

    slide.Shapes(2).TextFrame.TextRange.Text = resumeTexte

    ' === DIAPOSITIVE 3 : TABLEAU DÉTAILLÉ ===
    Set slide = pres.Slides.Add(3, 11)  ' Mise en page vide

    ' Ajouter un titre
    Dim titreTableau As Object
    Set titreTableau = slide.Shapes.AddTextbox(1, 50, 50, 600, 50)
    titreTableau.TextFrame.TextRange.Text = "Détail par produit"
    titreTableau.TextFrame.TextRange.Font.Size = 24
    titreTableau.TextFrame.TextRange.Font.Bold = True

    ' Créer un tableau avec les données Excel (supposons A5:C8)
    Dim tableau As Object
    Set tableau = slide.Shapes.AddTable(4, 3, 50, 120, 600, 250)

    ' En-têtes
    tableau.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Produit"
    tableau.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Quantité vendue"
    tableau.Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = "CA généré"

    ' Données (lignes 5 à 7 d'Excel)
    Dim i As Integer
    For i = 1 To 3
        tableau.Table.Cell(i + 1, 1).Shape.TextFrame.TextRange.Text = Cells(i + 4, 1).Value
        tableau.Table.Cell(i + 1, 2).Shape.TextFrame.TextRange.Text = Cells(i + 4, 2).Value
        tableau.Table.Cell(i + 1, 3).Shape.TextFrame.TextRange.Text = Format(Cells(i + 4, 3).Value, "#,##0") & " €"
    Next i

    ' Formater le tableau
    For i = 1 To 3
        tableau.Table.Cell(1, i).Shape.TextFrame.TextRange.Font.Bold = True
        tableau.Table.Cell(1, i).Shape.Fill.ForeColor.RGB = RGB(0, 50, 100)
        tableau.Table.Cell(1, i).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Next i

    ' === DIAPOSITIVE 4 : CONCLUSION ===
    Set slide = pres.Slides.Add(4, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "Conclusions et prochaines étapes"
    slide.Shapes(2).TextFrame.TextRange.Text = "• Excellente performance ce mois-ci" & vbCrLf & _
                                               "• Focus sur les produits les plus rentables" & vbCrLf & _
                                               "• Préparation du mois prochain" & vbCrLf & _
                                               "• Rapport généré automatiquement le " & Date

    ' Sauvegarder
    Dim cheminFichier As String
    cheminFichier = Environ("USERPROFILE") & "\Desktop\Rapport_Ventes_" & Format(Date, "yyyy-mm-dd") & ".pptx"
    pres.SaveAs cheminFichier

    MsgBox "Présentation créée avec succès !" & vbCrLf & cheminFichier

    ' Nettoyage
    pptApp.Quit
    Set titreTableau = Nothing
    Set tableau = Nothing
    Set slide = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
End Sub
```

## Points importants à retenir

### ✅ Bonnes pratiques
- Toujours libérer les objets avec `Set variable = Nothing`
- Fermer PowerPoint avec `pptApp.Quit` quand c'est terminé
- Utiliser `pptApp.Visible = True` pendant le développement pour voir le résultat
- Tester avec des présentations simples avant de créer des solutions complexes

### ⚠️ Erreurs courantes à éviter
- Oublier de fermer PowerPoint (instances multiples en arrière-plan)
- Essayer d'accéder à des formes qui n'existent pas
- Ne pas vérifier l'existence des fichiers images avant de les insérer
- Utiliser des chemins de fichiers en dur

### 💡 Conseils pour débuter
- Commencez par créer une seule diapositive simple
- Utilisez les mises en page prédéfinies (1, 2, 3, etc.) au début
- Testez le formatage sur de petits exemples
- Gardez une sauvegarde de vos données Excel avant d'automatiser

### 🎯 Utilisations typiques
- Rapports de ventes automatisés
- Présentations de tableaux de bord
- Mise à jour régulière de présentations avec nouvelles données
- Génération de supports de formation
- Création de présentations personnalisées en masse

L'automation avec PowerPoint ouvre de nombreuses possibilités pour créer des présentations dynamiques et toujours à jour avec vos données Excel. Une fois maîtrisée, cette technique vous fera gagner énormément de temps !

⏭️
