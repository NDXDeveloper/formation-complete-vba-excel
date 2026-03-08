🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.3 Graphiques et objets Shape

## Introduction : Pourquoi automatiser les graphiques ?

Imaginez que vous devez créer 50 graphiques identiques chaque mois pour vos rapports de vente. Plutôt que de les faire manuellement, VBA peut créer tous ces graphiques en quelques secondes ! De même, vous pouvez ajouter des formes, des flèches, des zones de texte pour rendre vos feuilles plus professionnelles et interactives.

Cette section vous apprendra à :
- 📊 Créer des graphiques par code
- 🎨 Personnaliser l'apparence des graphiques
- 🔷 Ajouter des formes (rectangles, cercles, flèches...)
- 📝 Insérer des zones de texte automatiquement
- 🖼️ Manipuler les images et objets graphiques

## Partie 1 : Les graphiques Excel en VBA

### Qu'est-ce qu'un graphique en VBA ?

En VBA, un graphique est un **objet Chart** qui peut être :
- **Incorporé dans une feuille** (ChartObject) - le plus courant
- **Sur une feuille séparée** (Chart) - moins fréquent

### Créer votre premier graphique

#### Exemple simple : Graphique en secteurs
```vba
Sub CreerGraphiqueSecteurs()
    ' Préparons d'abord quelques données
    Range("A1").Value = "Produit"
    Range("B1").Value = "Ventes"
    Range("A2:A5").Value = Application.Transpose(Array("Pommes", "Oranges", "Bananes", "Raisins"))
    Range("B2:B5").Value = Application.Transpose(Array(150, 200, 100, 80))

    ' Créer le graphique
    Dim monGraphique As ChartObject
    Set monGraphique = ActiveSheet.ChartObjects.Add(Left:=250, Top:=50, Width:=400, Height:=300)

    With monGraphique.Chart
        .SetSourceData Source:=Range("A1:B5")  ' Données à utiliser
        .ChartType = xlPie                     ' Type : secteurs
        .HasTitle = True                       ' Ajouter un titre
        .ChartTitle.Text = "Ventes par produit" ' Texte du titre
    End With

    MsgBox "Graphique créé avec succès !"
End Sub
```

**Explication du code :**
- `ChartObjects.Add` : Crée un nouvel emplacement pour le graphique
- `Left`, `Top` : Position du graphique (en points depuis le coin supérieur gauche)
- `Width`, `Height` : Taille du graphique
- `SetSourceData` : Définit les données à représenter
- `ChartType` : Type de graphique (secteurs, colonnes, courbes...)

#### Exemple : Graphique en colonnes avec personnalisation
```vba
Sub CreerGraphiqueColonnes()
    ' Préparons des données de ventes mensuelles
    Range("A1").Value = "Mois"
    Range("B1").Value = "Ventes"
    Range("A2:A7").Value = Application.Transpose(Array("Jan", "Fév", "Mar", "Avr", "Mai", "Jun"))
    Range("B2:B7").Value = Application.Transpose(Array(1200, 1500, 1800, 1600, 2100, 1900))

    ' Créer et personnaliser le graphique
    Dim graphique As ChartObject
    Set graphique = ActiveSheet.ChartObjects.Add(Left:=300, Top:=100, Width:=500, Height:=350)

    With graphique.Chart
        .SetSourceData Source:=Range("A1:B7")
        .ChartType = xlColumnClustered        ' Colonnes groupées

        ' Personnaliser le titre
        .HasTitle = True
        .ChartTitle.Text = "Évolution des ventes 2024"
        .ChartTitle.Font.Size = 16
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(0, 0, 150)  ' Bleu foncé

        ' Personnaliser les axes
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Mois"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Ventes (€)"

        ' Changer la couleur des colonnes
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(100, 150, 200)
    End With
End Sub
```

### Types de graphiques courants

| Type VBA | Nom Excel | Description |
|----------|-----------|-------------|
| `xlColumnClustered` | Histogramme groupé | Colonnes côte à côte |
| `xlLine` | Courbes | Lignes reliant les points |
| `xlPie` | Secteurs | Camembert |
| `xlArea` | Aires | Zones remplies |
| `xlXYScatter` | Nuage de points | Points dispersés |
| `xlBarClustered` | Barres groupées | Barres horizontales |

### Modifier un graphique existant

```vba
Sub ModifierGraphiqueExistant()
    ' Supposons qu'il y a déjà un graphique sur la feuille
    Dim graphique As ChartObject

    ' Vérifier s'il y a au moins un graphique
    If ActiveSheet.ChartObjects.Count > 0 Then
        Set graphique = ActiveSheet.ChartObjects(1)  ' Premier graphique

        With graphique.Chart
            ' Changer le type
            .ChartType = xlLine

            ' Modifier le titre
            .ChartTitle.Text = "Nouveau titre"

            ' Ajouter une légende
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
        End With

        MsgBox "Graphique modifié !"
    Else
        MsgBox "Aucun graphique trouvé sur cette feuille."
    End If
End Sub
```

## Partie 2 : Les objets Shape (Formes)

### Qu'est-ce qu'un objet Shape ?

Les **Shapes** sont tous les objets graphiques qu'on peut dessiner sur une feuille Excel :
- 🔷 Formes géométriques (rectangles, cercles, triangles...)
- ➡️ Flèches et connecteurs
- 📝 Zones de texte
- 🖼️ Images
- ⭐ Formes personnalisées

### Ajouter des formes de base

#### Exemple : Rectangle avec texte
```vba
Sub AjouterRectangle()
    Dim maForme As Shape

    ' Créer un rectangle
    Set maForme = ActiveSheet.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=100, _
        Top:=50, _
        Width:=200, _
        Height:=80)

    With maForme
        ' Personnaliser l'apparence
        .Fill.ForeColor.RGB = RGB(200, 220, 255)  ' Bleu clair
        .Line.ForeColor.RGB = RGB(0, 0, 200)      ' Bordure bleue
        .Line.Weight = 2                          ' Épaisseur bordure

        ' Ajouter du texte
        .TextFrame.Characters.Text = "Attention !"
        .TextFrame.Characters.Font.Size = 14
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
    End With
End Sub
```

#### Exemple : Cercle coloré
```vba
Sub AjouterCercle()
    Dim cercle As Shape

    Set cercle = ActiveSheet.Shapes.AddShape( _
        Type:=msoShapeOval, _
        Left:=300, _
        Top:=100, _
        Width:=100, _
        Height:=100)  ' Width = Height pour un cercle parfait

    With cercle
        .Fill.ForeColor.RGB = RGB(255, 200, 200)  ' Rose
        .Line.Visible = False                     ' Pas de bordure
        .Name = "MonCercle"                       ' Donner un nom
    End With
End Sub
```

### Types de formes courantes

| Constante VBA | Forme |
|---------------|-------|
| `msoShapeRectangle` | Rectangle |
| `msoShapeOval` | Cercle/Ellipse |
| `msoShapeTriangle` | Triangle |
| `msoShapeDiamond` | Losange |
| `msoShapeHeart` | Cœur |
| `msoShape5pointStar` | Étoile à 5 branches |

### Ajouter des flèches et connecteurs

```vba
Sub AjouterFleche()
    Dim fleche As Shape

    ' Créer une flèche
    Set fleche = ActiveSheet.Shapes.AddConnector( _
        Type:=msoConnectorStraight, _
        BeginX:=100, BeginY:=100, _
        EndX:=300, EndY:=100)

    With fleche.Line
        .ForeColor.RGB = RGB(255, 0, 0)      ' Rouge
        .Weight = 3                          ' Épaisseur
        .EndArrowheadStyle = msoArrowheadTriangle  ' Pointe de flèche
        .EndArrowheadWidth = msoArrowheadWidthMedium
    End With
End Sub
```

### Zones de texte

```vba
Sub AjouterZoneTexte()
    Dim zoneTexte As Shape

    Set zoneTexte = ActiveSheet.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=50, Top:=200, Width:=300, Height:=100)

    With zoneTexte
        ' Contenu du texte
        .TextFrame.Characters.Text = "Ceci est une zone de texte créée automatiquement par VBA !" & vbCrLf & "On peut mettre plusieurs lignes."

        ' Formatage du texte
        With .TextFrame.Characters.Font
            .Name = "Arial"
            .Size = 12
            .Color = RGB(0, 100, 0)  ' Vert
            .Bold = True
        End With

        ' Apparence de la zone
        .Fill.ForeColor.RGB = RGB(255, 255, 200)  ' Jaune clair
        .Line.ForeColor.RGB = RGB(200, 200, 0)    ' Bordure jaune foncé
    End With
End Sub
```

## Partie 3 : Manipulation avancée des objets

### Sélectionner et modifier des objets existants

```vba
Sub ModifierFormeParNom()
    Dim maForme As Shape

    ' Vérifier si la forme existe
    On Error GoTo FormeIntrouvable
    Set maForme = ActiveSheet.Shapes("MonCercle")  ' Nom donné précédemment

    ' Modifier la forme
    With maForme
        .Fill.ForeColor.RGB = RGB(100, 255, 100)  ' Vert clair
        .Left = 400  ' Déplacer
        .Top = 50
    End With

    MsgBox "Forme modifiée avec succès !"
    Exit Sub

FormeIntrouvable:
    MsgBox "La forme 'MonCercle' n'existe pas sur cette feuille."
End Sub
```

### Parcourir toutes les formes

```vba
Sub ListerToutesLesFormes()
    Dim forme As Shape
    Dim liste As String

    liste = "Formes présentes sur cette feuille :" & vbCrLf & vbCrLf

    ' Parcourir toutes les formes
    For Each forme In ActiveSheet.Shapes
        liste = liste & "- " & forme.Name & " (Type: " & forme.Type & ")" & vbCrLf
    Next forme

    If ActiveSheet.Shapes.Count = 0 Then
        liste = "Aucune forme trouvée sur cette feuille."
    End If

    MsgBox liste
End Sub
```

### Supprimer des objets

```vba
Sub SupprimerTousLesGraphiques()
    Dim reponse As VbMsgBoxResult

    reponse = MsgBox("Êtes-vous sûr de vouloir supprimer tous les graphiques ?", _
                     vbYesNo + vbQuestion, "Confirmation")

    If reponse = vbYes Then
        ' Supprimer tous les graphiques
        Dim i As Integer
        For i = ActiveSheet.ChartObjects.Count To 1 Step -1
            ActiveSheet.ChartObjects(i).Delete
        Next i

        MsgBox "Tous les graphiques ont été supprimés."
    End If
End Sub
```

## Exemple pratique complet : Tableau de bord automatisé

```vba
Sub CreerTableauDeBord()
    ' Nettoyer la feuille d'abord
    ActiveSheet.Cells.Clear
    ActiveSheet.ChartObjects.Delete

    ' Supprimer les formes une par une (la collection Shapes n'a pas de méthode Delete)
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Delete
    Next shp

    ' 1. Créer des données d'exemple
    Range("A1").Value = "Mois"
    Range("B1").Value = "Ventes"
    Range("A2:A7").Value = Application.Transpose(Array("Jan", "Fév", "Mar", "Avr", "Mai", "Jun"))
    Range("B2:B7").Value = Application.Transpose(Array(1200, 1500, 1800, 1600, 2100, 1900))

    ' 2. Créer un titre principal
    Dim titreShape As Shape
    Set titreShape = ActiveSheet.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=50, Top:=10, Width:=600, Height:=40)

    With titreShape
        .TextFrame.Characters.Text = "TABLEAU DE BORD - VENTES 2024"
        With .TextFrame.Characters.Font
            .Size = 18
            .Bold = True
            .Color = RGB(255, 255, 255)  ' Blanc
        End With
        .Fill.ForeColor.RGB = RGB(50, 50, 150)  ' Bleu foncé
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
    End With

    ' 3. Créer le graphique principal
    Dim graphique As ChartObject
    Set graphique = ActiveSheet.ChartObjects.Add(Left:=50, Top:=70, Width:=400, Height:=250)

    With graphique.Chart
        .SetSourceData Source:=Range("A1:B7")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Évolution mensuelle"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(100, 150, 200)
    End With

    ' 4. Ajouter des indicateurs
    Dim totalVentes As Double
    Dim moyenneVentes As Double

    totalVentes = Application.WorksheetFunction.Sum(Range("B2:B7"))
    moyenneVentes = Application.WorksheetFunction.Average(Range("B2:B7"))

    ' Indicateur Total
    Dim indicTotal As Shape
    Set indicTotal = ActiveSheet.Shapes.AddShape( _
        Type:=msoShapeRoundedRectangle, _
        Left:=480, Top:=70, Width:=150, Height:=60)

    With indicTotal
        .Fill.ForeColor.RGB = RGB(200, 255, 200)  ' Vert clair
        .TextFrame.Characters.Text = "TOTAL" & vbCrLf & Format(totalVentes, "# ##0 €")
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
    End With

    ' Indicateur Moyenne
    Dim indicMoyenne As Shape
    Set indicMoyenne = ActiveSheet.Shapes.AddShape( _
        Type:=msoShapeRoundedRectangle, _
        Left:=480, Top:=140, Width:=150, Height:=60)

    With indicMoyenne
        .Fill.ForeColor.RGB = RGB(255, 255, 200)  ' Jaune clair
        .TextFrame.Characters.Text = "MOYENNE" & vbCrLf & Format(moyenneVentes, "# ##0 €")
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
    End With

    ' 5. Ajouter une flèche d'évolution
    Dim fleche As Shape
    Set fleche = ActiveSheet.Shapes.AddConnector( _
        Type:=msoConnectorStraight, _
        BeginX:=480, BeginY:=260, EndX:=580, EndY:=220)

    With fleche.Line
        .ForeColor.RGB = RGB(0, 150, 0)  ' Vert
        .Weight = 4
        .EndArrowheadStyle = msoArrowheadTriangle
    End With

    ' 6. Note de tendance
    Dim noteTexte As Shape
    Set noteTexte = ActiveSheet.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=590, Top:=240, Width:=100, Height:=30)

    noteTexte.TextFrame.Characters.Text = "Tendance positive !"
    noteTexte.TextFrame.Characters.Font.Size = 10
    noteTexte.TextFrame.Characters.Font.Color = RGB(0, 150, 0)
    noteTexte.Line.Visible = False
    noteTexte.Fill.Visible = False

    MsgBox "Tableau de bord créé avec succès !"
End Sub
```

## Conseils pour débuter avec les graphiques et formes

### ✅ Bonnes pratiques

1. **Commencez simple** : Créez d'abord un graphique basique, puis ajoutez les personnalisations
2. **Nommez vos objets** : Donnez des noms explicites avec la propriété `.Name`
3. **Gérez les erreurs** : Vérifiez si les objets existent avant de les modifier
4. **Positionnement logique** : Planifiez l'emplacement de vos objets sur la feuille

### ⚠️ Pièges à éviter

1. **Données manquantes** : Vérifiez que vos plages de données contiennent bien des valeurs
2. **Objets en double** : Supprimez les anciens objets avant d'en créer de nouveaux
3. **Tailles fixes** : Évitez les positions en dur, adaptez-vous à la taille des données
4. **Surcharge visuelle** : Ne créez pas trop d'objets, gardez un design lisible

### 🔧 Outils de débogage

```vba
Sub DeboguerObjets()
    ' Compter les objets sur la feuille
    Debug.Print "Graphiques : " & ActiveSheet.ChartObjects.Count
    Debug.Print "Formes : " & ActiveSheet.Shapes.Count

    ' Lister les noms des formes
    Dim forme As Shape
    For Each forme In ActiveSheet.Shapes
        Debug.Print "Forme : " & forme.Name & " - Type : " & forme.Type
    Next forme

    ' Voir les résultats avec Ctrl+G dans l'éditeur VBA
End Sub
```

## Récapitulatif

Les graphiques et objets Shape en VBA vous permettent de :

- 📊 **Créer des graphiques automatiquement** à partir de vos données
- 🎨 **Personnaliser l'apparence** (couleurs, titres, légendes...)
- 🔷 **Ajouter des formes** pour enrichir vos feuilles
- 📝 **Insérer du texte mis en forme** avec les zones de texte
- 🎯 **Créer des tableaux de bord** visuels et interactifs

**Prochaine étape :** Nous verrons comment automatiser les tableaux croisés dynamiques pour des analyses de données encore plus poussées !

⏭️ [Tableaux croisés dynamiques](/14-fonctions-avancees-excel/04-tableaux-croises-dynamiques.md)
