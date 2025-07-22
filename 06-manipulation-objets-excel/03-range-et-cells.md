🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.3. Range et Cells

## Introduction aux objets Range et Cells

Après avoir maîtrisé les objets Application, Workbook et Worksheet, nous arrivons maintenant au cœur de la manipulation des données : les objets **Range** et **Cells**. Ces objets vous permettent d'interagir directement avec les cellules de vos feuilles Excel.

**Analogie simple :**
- **Range** = Une zone rectangulaire de cellules (comme sélectionner plusieurs cellules avec la souris)
- **Cells** = Une façon de désigner des cellules individuelles par leurs coordonnées (ligne, colonne)

Ces deux objets sont complémentaires et vous utiliserez constamment l'un ou l'autre selon vos besoins.

---

## L'objet Range

### Qu'est-ce qu'un Range ?

Un objet **Range** représente une ou plusieurs cellules sur une feuille Excel. C'est l'équivalent VBA de ce que vous sélectionnez quand vous cliquez et faites glisser votre souris sur des cellules dans Excel.

### Façons de désigner un Range

#### 1. Range par adresse de cellule (notation A1)

```vba
' Une seule cellule
Range("A1")                    ' La cellule A1
Range("B5")                    ' La cellule B5

' Une plage de cellules rectangulaire
Range("A1:C3")                 ' De A1 à C3 (9 cellules)
Range("B2:E10")                ' De B2 à E10

' Plusieurs plages séparées
Range("A1:A3,C1:C3")          ' Les cellules A1-A3 ET C1-C3
Range("A1,C3,E5")             ' Trois cellules spécifiques
```

#### 2. Range par nom défini

```vba
' Si vous avez défini un nom dans Excel (Formules > Gestionnaire de noms)
Range("MesVentes")             ' Utilise le nom défini "MesVentes"
Range("Zone_Calcul")           ' Utilise le nom défini "Zone_Calcul"
```

#### 3. Range combiné

```vba
' Combiner deux adresses pour former une plage
Range("A1", "C3")              ' Équivalent à Range("A1:C3")
Range("B2", "E10")             ' Équivalent à Range("B2:E10")
```

### Propriétés importantes de Range

#### 1. Valeur et contenu

```vba
' Lire la valeur d'une cellule
Dim contenu As Variant
contenu = Range("A1").Value
' Ou plus simplement :
contenu = Range("A1")          ' .Value est la propriété par défaut

' Écrire dans une cellule
Range("A1").Value = "Bonjour"
Range("A1") = "Bonjour"        ' Équivalent

' Écrire dans plusieurs cellules d'un coup
Range("A1:A3").Value = "Même texte partout"

' Effacer le contenu
Range("A1:C3").ClearContents   ' Efface seulement le contenu
Range("A1:C3").Clear           ' Efface contenu ET mise en forme
```

#### 2. Adresse et position

```vba
' Obtenir l'adresse d'un Range
Debug.Print Range("A1:C3").Address        ' Affiche "$A$1:$C$3"
Debug.Print Range("A1:C3").Address(False, False)  ' Affiche "A1:C3" (sans $)

' Nombre de lignes et colonnes
Debug.Print Range("A1:C3").Rows.Count     ' Affiche 3 (lignes)
Debug.Print Range("A1:C3").Columns.Count  ' Affiche 3 (colonnes)
Debug.Print Range("A1:C3").Cells.Count    ' Affiche 9 (total cellules)
```

#### 3. Formatage et apparence

```vba
' Police et taille
Range("A1").Font.Name = "Arial"
Range("A1").Font.Size = 14
Range("A1").Font.Bold = True
Range("A1").Font.Color = RGB(255, 0, 0)    ' Rouge

' Couleur de fond
Range("A1").Interior.Color = RGB(255, 255, 0)  ' Jaune
Range("A1").Interior.ColorIndex = 6             ' Jaune aussi (méthode ancienne)

' Bordures
Range("A1:C3").Borders.LineStyle = xlContinuous
Range("A1:C3").Borders.Weight = xlThick

' Alignement
Range("A1").HorizontalAlignment = xlCenter
Range("A1").VerticalAlignment = xlCenter
```

### Méthodes importantes de Range

#### 1. Sélection et activation

```vba
' Sélectionner une plage (équivalent au clic souris)
Range("A1:C3").Select

' Activer une cellule spécifique dans une sélection
Range("A1:C3").Select
Range("B2").Activate           ' B2 devient la cellule active
```

#### 2. Copier et coller

```vba
' Copier une plage
Range("A1:A3").Copy

' Coller dans une autre plage
Range("D1").Paste
' Ou
Range("A1:A3").Copy Range("D1")  ' Copie directe

' Collage spécial
Range("A1:A3").Copy
Range("D1").PasteSpecial xlPasteValues     ' Colle seulement les valeurs
Range("D1").PasteSpecial xlPasteFormats    ' Colle seulement la mise en forme
```

#### 3. Insertion et suppression

```vba
' Insérer des cellules
Range("A1:A3").Insert Shift:=xlShiftDown   ' Insère et pousse vers le bas
Range("A1:A3").Insert Shift:=xlShiftRight  ' Insère et pousse vers la droite

' Supprimer des cellules
Range("A1:A3").Delete Shift:=xlShiftUp     ' Supprime et remonte
Range("A1:A3").Delete Shift:=xlShiftLeft   ' Supprime et décale à gauche
```

---

## L'objet Cells

### Qu'est-ce que Cells ?

**Cells** est une façon alternative de désigner des cellules en utilisant des coordonnées numériques (ligne, colonne) au lieu de la notation alphabétique. C'est particulièrement utile dans les boucles et quand vous travaillez avec des positions calculées.

### Syntaxe de base de Cells

```vba
' Cells(ligne, colonne)
Cells(1, 1)                    ' Cellule A1 (ligne 1, colonne 1)
Cells(5, 3)                    ' Cellule C5 (ligne 5, colonne 3)
Cells(10, 1)                   ' Cellule A10 (ligne 10, colonne 1)
```

### Équivalences Range vs Cells

```vba
' Ces instructions sont équivalentes :
Range("A1") = "Bonjour"
Cells(1, 1) = "Bonjour"

Range("C5") = 100
Cells(5, 3) = 100

Range("Z26") = "Fin"
Cells(26, 26) = "Fin"          ' Z est la 26ème lettre
```

### Avantages de Cells

#### 1. Facilité dans les boucles

```vba
' Remplir une colonne avec des nombres
Dim i As Integer
For i = 1 To 10
    Cells(i, 1) = i            ' Plus simple que Range("A" & i)
Next i

' Remplir une ligne avec des nombres
For i = 1 To 5
    Cells(1, i) = i * 10       ' Cellules A1, B1, C1, D1, E1
Next i
```

#### 2. Calculs de position

```vba
' Variables pour position
Dim ligne As Integer
Dim colonne As Integer

ligne = 5
colonne = 3
Cells(ligne, colonne) = "Position calculée"   ' Cellule C5

' Déplacement relatif
Cells(ligne + 1, colonne) = "Ligne suivante"  ' Cellule C6
Cells(ligne, colonne + 1) = "Colonne suivante" ' Cellule D5
```

### Créer des Ranges avec Cells

```vba
' Range en utilisant Cells comme points de départ et fin
Range(Cells(1, 1), Cells(3, 3))           ' Équivalent à Range("A1:C3")
Range(Cells(2, 2), Cells(10, 5))          ' Équivalent à Range("B2:E10")

' Très utile avec des variables
Dim ligneDébut As Integer, ligneF fin As Integer
Dim colonneDébut As Integer, colonneFin As Integer

ligneDébut = 2
ligneFin = 10
colonneDébut = 1
colonneFin = 4

Range(Cells(ligneDébut, colonneDébut), Cells(ligneFin, colonneFin)).Select
' Sélectionne la zone A2:D10
```

---

## Navigation et déplacement

### Propriétés de navigation de Range

#### 1. Déplacement relatif

```vba
' À partir de A1
Dim celluleBase As Range
Set celluleBase = Range("A1")

' Cellules adjacentes
Set celluleDroite = celluleBase.Offset(0, 1)      ' B1 (même ligne, colonne +1)
Set celluleBas = celluleBase.Offset(1, 0)         ' A2 (ligne +1, même colonne)
Set celluleDiagonale = celluleBase.Offset(1, 1)   ' B2 (ligne +1, colonne +1)

' Utilisation directe
Range("A1").Offset(2, 3) = "Cellule D3"           ' 2 lignes plus bas, 3 colonnes à droite
```

#### 2. Redimensionnement

```vba
' Agrandir ou rétrécir une plage
Dim maPlage As Range
Set maPlage = Range("A1:B2")                      ' Plage 2x2

Set plagePlus = maPlage.Resize(4, 4)              ' Devient A1:D4 (4x4)
Set plageMoins = maPlage.Resize(1, 1)             ' Devient A1:A1 (1x1)

' Combinaison Offset + Resize
Range("A1").Offset(1, 1).Resize(3, 2) = "Test"   ' Rempli B2:C4
```

#### 3. Navigation jusqu'aux limites

```vba
' Trouver la dernière cellule utilisée dans une direction
Dim derniereCellule As Range

' Dernière cellule vers la droite (équivalent Ctrl+Flèche droite)
Set derniereCellule = Range("A1").End(xlToRight)

' Dernière cellule vers le bas (équivalent Ctrl+Flèche bas)
Set derniereCellule = Range("A1").End(xlDown)

' Dernière cellule vers la gauche
Set derniereCellule = Range("Z1").End(xlToLeft)

' Dernière cellule vers le haut
Set derniereCellule = Range("A100").End(xlUp)
```

### Exemples pratiques de navigation

#### 1. Trouver la dernière ligne avec des données

```vba
' Dernière ligne de la colonne A contenant des données
Dim derniereLigne As Long
derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row
Debug.Print "Dernière ligne avec données : " & derniereLigne

' Ou pour n'importe quelle colonne
derniereLigne = Cells(Rows.Count, "C").End(xlUp).Row  ' Colonne C
```

#### 2. Trouver la dernière colonne avec des données

```vba
' Dernière colonne de la ligne 1 contenant des données
Dim derniereColonne As Long
derniereColonne = Cells(1, Columns.Count).End(xlToLeft).Column
Debug.Print "Dernière colonne avec données : " & derniereColonne
```

#### 3. Sélectionner une zone de données complète

```vba
' Sélectionner toute la zone de données à partir d'A1
Range("A1").CurrentRegion.Select

' Ou définir cette zone dans une variable
Dim zoneDonnees As Range
Set zoneDonnees = Range("A1").CurrentRegion
Debug.Print "Zone de données : " & zoneDonnees.Address
```

---

## Manipulation avancée de Range et Cells

### Parcourir une plage de cellules

#### 1. Avec For Each (recommandé pour les valeurs)

```vba
' Parcourir chaque cellule d'une plage
Dim cellule As Range
For Each cellule In Range("A1:A10")
    Debug.Print cellule.Address & " : " & cellule.Value
Next cellule

' Traitement conditionnel
For Each cellule In Range("A1:A10")
    If IsNumeric(cellule.Value) Then
        cellule.Value = cellule.Value * 2  ' Doubler les nombres
    End If
Next cellule
```

#### 2. Avec des boucles numériques (plus flexible)

```vba
' Parcourir avec des indices
Dim i As Integer
For i = 1 To 10
    Debug.Print "A" & i & " : " & Cells(i, 1).Value
    Cells(i, 2) = Cells(i, 1).Value * 3    ' Copier en triplant en colonne B
Next i

' Parcourir une zone rectangulaire
Dim ligne As Integer, colonne As Integer
For ligne = 1 To 5
    For colonne = 1 To 3
        Cells(ligne, colonne) = "L" & ligne & "C" & colonne
    Next colonne
Next ligne
```

### Recherche dans les cellules

#### 1. Méthode Find

```vba
' Rechercher une valeur dans une plage
Dim celluleTrouvee As Range
Set celluleTrouvee = Range("A1:A100").Find("Recherché")

If Not celluleTrouvee Is Nothing Then
    Debug.Print "Trouvé en : " & celluleTrouvee.Address
    celluleTrouvee.Select
Else
    Debug.Print "Valeur non trouvée"
End If
```

#### 2. Recherche avec critères

```vba
' Recherche plus précise
Set celluleTrouvee = Range("A1:Z100").Find( _
    What:="MonTexte", _
    LookIn:=xlValues, _
    LookAt:=xlWhole, _
    MatchCase:=False)
```

### Tri et filtrage

#### 1. Tri simple

```vba
' Trier une plage par la première colonne (croissant)
Range("A1:C10").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

' Trier par plusieurs colonnes
Range("A1:C10").Sort _
    Key1:=Range("A1"), Order1:=xlAscending, _
    Key2:=Range("B1"), Order2:=xlDescending, _
    Header:=xlYes
```

## Conversion entre notations

### De coordonnées numériques vers notation A1

```vba
' Convertir ligne/colonne vers adresse A1
Dim adresse As String
adresse = Cells(5, 3).Address        ' Retourne "$C$5"
adresse = Cells(5, 3).Address(False, False)  ' Retourne "C5"
```

### De notation A1 vers coordonnées

```vba
' Obtenir ligne et colonne d'une adresse
Dim maPlage As Range
Set maPlage = Range("C5")

Debug.Print maPlage.Row             ' Affiche 5
Debug.Print maPlage.Column          ' Affiche 3
```

## Bonnes pratiques et conseils

### 1. Quand utiliser Range vs Cells

**Utilisez Range quand :**
- Vous connaissez l'adresse exacte ("A1", "B2:D10")
- Vous travaillez avec des plages nommées
- Le code est plus lisible avec la notation A1

**Utilisez Cells quand :**
- Vous travaillez dans des boucles
- Les positions sont calculées
- Vous manipulez des coordonnées variables

### 2. Performance et optimisation

```vba
' ÉVITEZ : Accès cellule par cellule dans une boucle
For i = 1 To 1000
    Cells(i, 1) = i                 ' Lent pour de gros volumes
Next i

' PRÉFÉREZ : Manipulation par blocs
Dim valeurs(1 To 1000, 1 To 1) As Variant
For i = 1 To 1000
    valeurs(i, 1) = i
Next i
Range("A1:A1000").Value = valeurs   ' Beaucoup plus rapide
```

### 3. Gestion des erreurs

```vba
' Vérifier l'existence d'une plage nommée
On Error Resume Next
Dim plageNommee As Range
Set plageNommee = Range("MonNom")
If plageNommee Is Nothing Then
    Debug.Print "La plage nommée n'existe pas"
End If
On Error GoTo 0
```

## Exemples pratiques courants

### 1. Remplir une série de nombres

```vba
' Remplir A1:A10 avec les nombres 1 à 10
Dim i As Integer
For i = 1 To 10
    Cells(i, 1) = i
Next i

' Ou avec Range
Range("A1:A10").Formula = "=ROW()"
Range("A1:A10").Value = Range("A1:A10").Value  ' Convertir formules en valeurs
```

### 2. Copier des données en filtrant

```vba
' Copier seulement les cellules non vides
Dim i As Integer, j As Integer
j = 1
For i = 1 To 20
    If Cells(i, 1) <> "" Then
        Cells(j, 3) = Cells(i, 1)   ' Copier en colonne C
        j = j + 1
    End If
Next i
```

### 3. Formatage conditionnel simple

```vba
' Colorier les cellules selon leur valeur
Dim cellule As Range
For Each cellule In Range("A1:A10")
    If IsNumeric(cellule.Value) Then
        If cellule.Value > 50 Then
            cellule.Interior.Color = RGB(0, 255, 0)    ' Vert si > 50
        Else
            cellule.Interior.Color = RGB(255, 0, 0)    ' Rouge si ≤ 50
        End If
    End If
Next cellule
```

## Points clés à retenir

- **Range** utilise la notation familière d'Excel (A1, B2:D10)
- **Cells** utilise des coordonnées numériques (ligne, colonne) - idéal pour les boucles
- Les deux objets ont les mêmes propriétés et méthodes principales
- **Offset** et **Resize** permettent de naviguer et redimensionner facilement
- **End(direction)** permet de trouver les limites des données
- **CurrentRegion** sélectionne automatiquement une zone de données
- Pour la performance, préférez les opérations par blocs aux boucles cellule par cellule
- Toujours vérifier l'existence des objets pour éviter les erreurs

Ces objets Range et Cells sont les outils fondamentaux pour toute manipulation de données dans Excel via VBA. Maîtriser leur utilisation vous permettra de créer des macros puissantes et efficaces pour automatiser vos tâches quotidiennes.

⏭️
