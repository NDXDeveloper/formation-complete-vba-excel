🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.4. Tableaux multidimensionnels

## Introduction aux tableaux multidimensionnels

Les **tableaux multidimensionnels** sont comme des **classeurs multi-tiroirs** : au lieu d'avoir une simple ligne d'éléments, vous avez des structures organisées en **lignes et colonnes** (2D), ou même en **couches supplémentaires** (3D et plus). Ces tableaux sont parfaits pour représenter des données complexes comme des grilles, des matrices, ou des structures hiérarchiques.

**Analogies simples :**
- **Tableau 1D** = Une liste de courses (une dimension : position dans la liste)
- **Tableau 2D** = Une feuille Excel (deux dimensions : ligne et colonne)
- **Tableau 3D** = Un classeur Excel (trois dimensions : ligne, colonne, feuille)
- **Tableau 4D+** = Une armoire avec plusieurs classeurs (et ainsi de suite...)

Ces structures permettent de modéliser naturellement des données du monde réel qui ont plusieurs caractéristiques organisées.

---

## Tableaux à deux dimensions (2D)

### Concept et représentation

Un tableau 2D est comme une **grille** ou une **matrice** avec des lignes et des colonnes. Chaque élément est identifié par deux indices : sa position de ligne et sa position de colonne.

```
Visualisation d'un tableau 2D (3x4) :
        Col1  Col2  Col3  Col4
Ligne1:  A     B     C     D  
Ligne2:  E     F     G     H  
Ligne3:  I     J     K     L  
```

### Déclaration et initialisation

```vba
Sub TableauDeuxDimensions()
    ' Déclaration d'un tableau 2D (3 lignes x 4 colonnes)
    Dim grille(1 To 3, 1 To 4) As String

    ' Initialisation élément par élément
    grille(1, 1) = "A": grille(1, 2) = "B": grille(1, 3) = "C": grille(1, 4) = "D"
    grille(2, 1) = "E": grille(2, 2) = "F": grille(2, 3) = "G": grille(2, 4) = "H"
    grille(3, 1) = "I": grille(3, 2) = "J": grille(3, 3) = "K": grille(3, 4) = "L"

    ' Affichage de la grille
    Dim ligne As Integer, colonne As Integer
    For ligne = 1 To 3
        Dim ligneTexte As String
        ligneTexte = ""
        For colonne = 1 To 4
            ligneTexte = ligneTexte & grille(ligne, colonne) & vbTab
        Next colonne
        Debug.Print ligneTexte
    Next ligne

    ' Résultat affiché :
    ' A    B    C    D
    ' E    F    G    H
    ' I    J    K    L
End Sub
```

### Exemples pratiques 2D

#### 1. **Matrice de nombres**

```vba
Sub MatriceNombres()
    ' Tableau pour stocker une table de multiplication
    Dim tableMultiplication(1 To 10, 1 To 10) As Integer

    ' Remplissage de la table
    Dim i As Integer, j As Integer
    For i = 1 To 10
        For j = 1 To 10
            tableMultiplication(i, j) = i * j
        Next j
    Next i

    ' Affichage de quelques valeurs
    Debug.Print "Table de multiplication 5x5 :"
    For i = 1 To 5
        Dim ligne As String
        ligne = ""
        For j = 1 To 5
            ligne = ligne & Format(tableMultiplication(i, j), "000") & " "
        Next j
        Debug.Print ligne
    Next i

    ' Accès direct à des valeurs spécifiques
    Debug.Print "7 x 8 = " & tableMultiplication(7, 8)  ' 56
End Sub
```

#### 2. **Grille de données Excel**

```vba
Sub GrilleDonneesExcel()
    ' Simuler une grille de données comme dans Excel
    Dim donneesVentes(1 To 12, 1 To 4) As Variant  ' 12 mois, 4 trimestres

    ' En-têtes (conceptuels)
    ' Colonne 1: Trimestre 1, Colonne 2: Trimestre 2, etc.
    ' Ligne 1: Janvier, Ligne 2: Février, etc.

    ' Remplissage avec des données de vente aléatoires
    Dim mois As Integer, trimestre As Integer
    For mois = 1 To 12
        For trimestre = 1 To 4
            donneesVentes(mois, trimestre) = Int(Rnd() * 1000) + 100  ' Entre 100 et 1099
        Next trimestre
    Next mois

    ' Affichage des données du premier trimestre
    Debug.Print "Ventes du premier trimestre :"
    Dim nomsMois As Variant
    nomsMois = Array("Jan", "Fév", "Mar", "Avr", "Mai", "Jun", "Jul", "Aoû", "Sep", "Oct", "Nov", "Déc")

    For mois = 1 To 12
        Debug.Print nomsMois(mois - 1) & ": " & donneesVentes(mois, 1) & "€"
    Next mois
End Sub
```

#### 3. **Jeu de plateau (échiquier)**

```vba
Sub Echiquier()
    ' Plateau d'échecs 8x8
    Dim plateau(1 To 8, 1 To 8) As String

    ' Initialiser avec des espaces vides
    Dim ligne As Integer, colonne As Integer
    For ligne = 1 To 8
        For colonne = 1 To 8
            plateau(ligne, colonne) = "."  ' Case vide
        Next colonne
    Next ligne

    ' Placer quelques pièces
    plateau(1, 1) = "T": plateau(1, 8) = "T"  ' Tours blanches
    plateau(1, 4) = "R": plateau(1, 5) = "D"  ' Roi et Dame blancs
    plateau(8, 1) = "t": plateau(8, 8) = "t"  ' Tours noires
    plateau(8, 4) = "r": plateau(8, 5) = "d"  ' Roi et Dame noirs

    ' Affichage du plateau
    Debug.Print "Plateau d'échecs :"
    For ligne = 8 To 1 Step -1  ' Afficher de haut en bas
        Dim ligneTexte As String
        ligneTexte = ligne & " "
        For colonne = 1 To 8
            ligneTexte = ligneTexte & plateau(ligne, colonne) & " "
        Next colonne
        Debug.Print ligneTexte
    Next ligne
    Debug.Print "  a b c d e f g h"
End Sub
```

---

## Tableaux à trois dimensions (3D)

### Concept et visualisation

Un tableau 3D ajoute une **troisième dimension** comme des "couches" ou des "niveaux". C'est comme avoir plusieurs grilles 2D empilées les unes sur les autres.

```
Visualisation d'un tableau 3D (2x3x2) :  
Couche 1:          Couche 2:  
  C1  C2  C3         C1  C2  C3
L1 A   B   C      L1 G   H   I  
L2 D   E   F      L2 J   K   L  
```

### Déclaration et utilisation

```vba
Sub TableauTroisDimensions()
    ' Déclaration : (lignes, colonnes, couches)
    Dim cube(1 To 2, 1 To 3, 1 To 2) As String

    ' Remplissage couche par couche
    ' Couche 1
    cube(1, 1, 1) = "A": cube(1, 2, 1) = "B": cube(1, 3, 1) = "C"
    cube(2, 1, 1) = "D": cube(2, 2, 1) = "E": cube(2, 3, 1) = "F"

    ' Couche 2
    cube(1, 1, 2) = "G": cube(1, 2, 2) = "H": cube(1, 3, 2) = "I"
    cube(2, 1, 2) = "J": cube(2, 2, 2) = "K": cube(2, 3, 2) = "L"

    ' Affichage couche par couche
    Dim couche As Integer, ligne As Integer, colonne As Integer
    For couche = 1 To 2
        Debug.Print "=== Couche " & couche & " ==="
        For ligne = 1 To 2
            Dim ligneTexte As String
            ligneTexte = ""
            For colonne = 1 To 3
                ligneTexte = ligneTexte & cube(ligne, colonne, couche) & " "
            Next colonne
            Debug.Print ligneTexte
        Next ligne
        Debug.Print ""
    Next couche
End Sub
```

### Exemples pratiques 3D

#### 1. **Données de ventes par région, mois et produit**

```vba
Sub VentesMultidimensionnelles()
    ' 3 régions x 12 mois x 5 produits
    Dim ventes(1 To 3, 1 To 12, 1 To 5) As Double

    ' Simuler des données de ventes
    Dim region As Integer, mois As Integer, produit As Integer
    For region = 1 To 3
        For mois = 1 To 12
            For produit = 1 To 5
                ventes(region, mois, produit) = Rnd() * 1000 + 500  ' Entre 500 et 1500
            Next produit
        Next mois
    Next region

    ' Calculer les totaux par région pour janvier (mois 1)
    Debug.Print "Ventes totales par région en janvier :"
    For region = 1 To 3
        Dim totalRegion As Double
        totalRegion = 0
        For produit = 1 To 5
            totalRegion = totalRegion + ventes(region, 1, produit)
        Next produit
        Debug.Print "Région " & region & ": " & Format(totalRegion, "0.00") & "€"
    Next region

    ' Meilleur produit en région 1
    Dim meilleureVente As Double
    Dim meilleurProduit As Integer
    meilleureVente = 0

    For produit = 1 To 5
        If ventes(1, 1, produit) > meilleureVente Then
            meilleureVente = ventes(1, 1, produit)
            meilleurProduit = produit
        End If
    Next produit

    Debug.Print "Meilleur produit en région 1 : Produit " & meilleurProduit & _
                " (" & Format(meilleureVente, "0.00") & "€)"
End Sub
```

#### 2. **Inventaire d'entrepôt (étage, allée, étagère)**

```vba
Sub InventaireEntrepot()
    ' 3 étages x 10 allées x 20 étagères
    Dim inventaire(1 To 3, 1 To 10, 1 To 20) As Integer

    ' Remplissage aléatoire de l'inventaire
    Dim etage As Integer, allee As Integer, etagere As Integer
    For etage = 1 To 3
        For allee = 1 To 10
            For etagere = 1 To 20
                inventaire(etage, allee, etagere) = Int(Rnd() * 100)  ' 0 à 99 articles
            Next etagere
        Next allee
    Next etage

    ' Fonction pour trouver un emplacement spécifique
    Debug.Print "Stock à l'étage 2, allée 5, étagère 10 : " & _
                inventaire(2, 5, 10) & " articles"

    ' Calculer le stock total d'une allée
    Dim stockAllee As Long
    stockAllee = 0
    etage = 1
    allee = 3

    For etagere = 1 To 20
        stockAllee = stockAllee + inventaire(etage, allee, etagere)
    Next etagere

    Debug.Print "Stock total étage " & etage & ", allée " & allee & " : " & stockAllee

    ' Trouver l'étagère la plus remplie
    Dim maxStock As Integer
    Dim etageMax As Integer, alleeMax As Integer, etagereMax As Integer
    maxStock = 0

    For etage = 1 To 3
        For allee = 1 To 10
            For etagere = 1 To 20
                If inventaire(etage, allee, etagere) > maxStock Then
                    maxStock = inventaire(etage, allee, etagere)
                    etageMax = etage
                    alleeMax = allee
                    etagereMax = etagere
                End If
            Next etagere
        Next allee
    Next etage

    Debug.Print "Étagère la plus remplie : Étage " & etageMax & _
                ", Allée " & alleeMax & ", Étagère " & etagereMax & _
                " (" & maxStock & " articles)"
End Sub
```

---

## Navigation et manipulation

### Parcours efficace des tableaux multidimensionnels

#### 1. **Parcours ordonné (ligne par ligne)**

```vba
Sub ParcoursOrdonne()
    Dim donnees(1 To 4, 1 To 3) As Integer

    ' Remplissage ordonné
    Dim valeur As Integer
    valeur = 1

    Dim i As Integer, j As Integer
    For i = 1 To 4
        For j = 1 To 3
            donnees(i, j) = valeur
            valeur = valeur + 1
        Next j
    Next i

    ' Affichage avec parcours ordonné
    Debug.Print "Parcours ligne par ligne :"
    For i = 1 To 4
        Dim ligne As String
        ligne = ""
        For j = 1 To 3
            ligne = ligne & Format(donnees(i, j), "00") & " "
        Next j
        Debug.Print ligne
    Next i
End Sub
```

#### 2. **Parcours en colonne**

```vba
Sub ParcoursColonne()
    Dim matrice(1 To 3, 1 To 4) As Integer

    ' Remplissage
    Dim i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 4
            matrice(i, j) = i * 10 + j
        Next j
    Next i

    ' Parcours colonne par colonne
    Debug.Print "Parcours colonne par colonne :"
    For j = 1 To 4  ' Colonnes en premier
        Debug.Print "Colonne " & j & ":"
        For i = 1 To 3  ' Puis lignes
            Debug.Print "  " & matrice(i, j)
        Next i
    Next j
End Sub
```

#### 3. **Recherche dans un tableau multidimensionnel**

```vba
Sub RechercheMultidimensionnelle()
    Dim grille(1 To 5, 1 To 5) As Integer

    ' Remplissage avec des valeurs aléatoires
    Dim i As Integer, j As Integer
    For i = 1 To 5
        For j = 1 To 5
            grille(i, j) = Int(Rnd() * 100) + 1  ' 1 à 100
        Next j
    Next i

    ' Rechercher une valeur spécifique
    Dim valeurCherchee As Integer
    valeurCherchee = 50

    Dim trouve As Boolean
    Dim ligneT As Integer, colonneT As Integer
    trouve = False

    For i = 1 To 5
        For j = 1 To 5
            If grille(i, j) = valeurCherchee Then
                trouve = True
                ligneT = i
                colonneT = j
                Exit For  ' Sortir de la boucle interne
            End If
        Next j
        If trouve Then Exit For  ' Sortir de la boucle externe
    Next i

    If trouve Then
        Debug.Print "Valeur " & valeurCherchee & " trouvée en (" & ligneT & ", " & colonneT & ")"
    Else
        Debug.Print "Valeur " & valeurCherchee & " non trouvée"
    End If

    ' Afficher la grille pour vérification
    Debug.Print "Grille complète :"
    For i = 1 To 5
        Dim ligne As String
        ligne = ""
        For j = 1 To 5
            ligne = ligne & Format(grille(i, j), "000") & " "
        Next j
        Debug.Print ligne
    Next i
End Sub
```

---

## Tableaux multidimensionnels dynamiques

### Redimensionnement avec ReDim

**Limitation importante :** Avec `ReDim Preserve`, seule la **dernière dimension** peut être modifiée.

```vba
Sub TableauxMultidimensionnelsDynamiques()
    ' Déclaration d'un tableau 2D dynamique
    Dim donnees() As String

    ' Première allocation
    ReDim donnees(1 To 3, 1 To 2)
    donnees(1, 1) = "A": donnees(1, 2) = "B"
    donnees(2, 1) = "C": donnees(2, 2) = "D"
    donnees(3, 1) = "E": donnees(3, 2) = "F"

    Debug.Print "Tableau initial (3x2) :"
    Dim i As Integer, j As Integer
    For i = 1 To 3
        Dim ligne As String
        ligne = ""
        For j = 1 To 2
            ligne = ligne & donnees(i, j) & " "
        Next j
        Debug.Print ligne
    Next i

    ' ✅ AUTORISÉ : Modifier la dernière dimension (colonnes)
    ReDim Preserve donnees(1 To 3, 1 To 4)
    donnees(1, 3) = "G": donnees(1, 4) = "H"
    donnees(2, 3) = "I": donnees(2, 4) = "J"
    donnees(3, 3) = "K": donnees(3, 4) = "L"

    Debug.Print "Après extension des colonnes (3x4) :"
    For i = 1 To 3
        ligne = ""
        For j = 1 To 4
            ligne = ligne & donnees(i, j) & " "
        Next j
        Debug.Print ligne
    Next i

    ' ❌ INTERDIT : Modifier la première dimension avec Preserve
    ' ReDim Preserve donnees(1 To 5, 1 To 4)  ' ERREUR !

    ' Solution : Recréer sans Preserve (perte des données)
    ReDim donnees(1 To 5, 1 To 4)
    Debug.Print "Tableau recrée (5x4) - données perdues"
End Sub
```

### Solution de contournement pour modifier toutes les dimensions

```vba
Sub ModifierToutesDimensions()
    ' Tableau original
    Dim original(1 To 2, 1 To 3) As Integer
    original(1, 1) = 10: original(1, 2) = 20: original(1, 3) = 30
    original(2, 1) = 40: original(2, 2) = 50: original(2, 3) = 60

    ' Nouveau tableau avec dimensions différentes
    Dim nouveau(1 To 4, 1 To 5) As Integer

    ' Copier les données existantes
    Dim i As Integer, j As Integer
    For i = 1 To 2  ' Limites de l'original
        For j = 1 To 3
            nouveau(i, j) = original(i, j)
        Next j
    Next i

    ' Ajouter de nouvelles données
    nouveau(3, 1) = 70: nouveau(3, 2) = 80
    nouveau(4, 4) = 90: nouveau(4, 5) = 100

    Debug.Print "Nouveau tableau (4x5) :"
    For i = 1 To 4
        Dim ligne As String
        ligne = ""
        For j = 1 To 5
            ligne = ligne & Format(nouveau(i, j), "000") & " "
        Next j
        Debug.Print ligne
    Next i
End Sub
```

---

## Applications pratiques avancées

### 1. Matrice mathématique

```vba
Sub OperationsMatrices()
    ' Multiplication de matrices 2x3 et 3x2
    Dim matA(1 To 2, 1 To 3) As Double
    Dim matB(1 To 3, 1 To 2) As Double
    Dim resultat(1 To 2, 1 To 2) As Double

    ' Initialisation de la matrice A
    matA(1, 1) = 1: matA(1, 2) = 2: matA(1, 3) = 3
    matA(2, 1) = 4: matA(2, 2) = 5: matA(2, 3) = 6

    ' Initialisation de la matrice B
    matB(1, 1) = 7: matB(1, 2) = 8
    matB(2, 1) = 9: matB(2, 2) = 10
    matB(3, 1) = 11: matB(3, 2) = 12

    ' Multiplication matricielle : C = A × B
    Dim i As Integer, j As Integer, k As Integer
    For i = 1 To 2
        For j = 1 To 2
            resultat(i, j) = 0
            For k = 1 To 3
                resultat(i, j) = resultat(i, j) + matA(i, k) * matB(k, j)
            Next k
        Next j
    Next i

    ' Affichage des résultats
    Debug.Print "Matrice A (2x3):"
    For i = 1 To 2
        Debug.Print matA(i, 1) & " " & matA(i, 2) & " " & matA(i, 3)
    Next i

    Debug.Print "Matrice B (3x2):"
    For i = 1 To 3
        Debug.Print matB(i, 1) & " " & matB(i, 2)
    Next i

    Debug.Print "Résultat A×B (2x2):"
    For i = 1 To 2
        Debug.Print resultat(i, 1) & " " & resultat(i, 2)
    Next i
End Sub
```

### 2. Analyse de données complexes

```vba
Sub AnalyseDonneesComplexes()
    ' Données de ventes : 4 trimestres x 5 régions x 3 produits
    Dim ventes(1 To 4, 1 To 5, 1 To 3) As Double

    ' Remplissage avec des données simulées
    Dim trimestre As Integer, region As Integer, produit As Integer
    For trimestre = 1 To 4
        For region = 1 To 5
            For produit = 1 To 3
                ventes(trimestre, region, produit) = Rnd() * 1000 + 200
            Next produit
        Next region
    Next trimestre

    ' Analyse 1 : Meilleur trimestre global
    Dim totalTrimestre(1 To 4) As Double
    For trimestre = 1 To 4
        For region = 1 To 5
            For produit = 1 To 3
                totalTrimestre(trimestre) = totalTrimestre(trimestre) + ventes(trimestre, region, produit)
            Next produit
        Next region
    Next trimestre

    Dim meilleurTrimestre As Integer
    Dim maxVente As Double
    maxVente = 0
    For trimestre = 1 To 4
        If totalTrimestre(trimestre) > maxVente Then
            maxVente = totalTrimestre(trimestre)
            meilleurTrimestre = trimestre
        End If
        Debug.Print "Trimestre " & trimestre & ": " & Format(totalTrimestre(trimestre), "0.00") & "€"
    Next trimestre

    Debug.Print "Meilleur trimestre : " & meilleurTrimestre & " (" & Format(maxVente, "0.00") & "€)"

    ' Analyse 2 : Performance par produit
    Dim totalProduit(1 To 3) As Double
    For produit = 1 To 3
        For trimestre = 1 To 4
            For region = 1 To 5
                totalProduit(produit) = totalProduit(produit) + ventes(trimestre, region, produit)
            Next region
        Next trimestre
        Debug.Print "Produit " & produit & ": " & Format(totalProduit(produit), "0.00") & "€"
    Next produit
End Sub
```

---

## Optimisation et bonnes pratiques

### 1. Ordre des boucles pour la performance

```vba
Sub OptimisationOrdreBoucles()
    Dim donnees(1 To 1000, 1 To 1000) As Double
    Dim debut As Double

    ' Test 1 : Ordre naturel (ligne puis colonne)
    debut = Timer
    Dim i As Long, j As Long
    For i = 1 To 1000
        For j = 1 To 1000
            donnees(i, j) = i + j
        Next j
    Next i
    Debug.Print "Ordre naturel : " & Format(Timer - debut, "0.000") & " secondes"

    ' Test 2 : Ordre inversé (colonne puis ligne) - souvent plus lent
    debut = Timer
    For j = 1 To 1000
        For i = 1 To 1000
            donnees(i, j) = i * j
        Next i
    Next j
    Debug.Print "Ordre inversé : " & Format(Timer - debut, "0.000") & " secondes"

    ' En général, l'ordre "ligne puis colonne" est plus efficace
    ' car il suit l'organisation mémoire du tableau
End Sub
```

### 2. Limitation de la profondeur

```vba
Sub LimitationProfondeur()
    ' Éviter trop de dimensions - difficile à gérer et peu performant

    ' ❌ Trop complexe (5 dimensions)
    ' Dim tableau5D(1 To 10, 1 To 10, 1 To 10, 1 To 10, 1 To 10) As Integer

    ' ✅ Alternative : structure avec tableaux 2D multiples
    Dim departement1(1 To 10, 1 To 10) As Integer
    Dim departement2(1 To 10, 1 To 10) As Integer
    Dim departement3(1 To 10, 1 To 10) As Integer

    ' Ou utiliser des tableaux de tableaux (plus avancé)
    Debug.Print "Préférer les structures simples et claires"
End Sub
```

### 3. Fonction d'aide pour l'affichage

```vba
Sub AfficherTableau2D(arr As Variant, titre As String)
    Debug.Print "=== " & titre & " ==="

    Dim minLigne As Integer, maxLigne As Integer
    Dim minCol As Integer, maxCol As Integer

    minLigne = LBound(arr, 1): maxLigne = UBound(arr, 1)
    minCol = LBound(arr, 2): maxCol = UBound(arr, 2)

    Dim i As Integer, j As Integer
    For i = minLigne To maxLigne
        Dim ligne As String
        ligne = ""
        For j = minCol To maxCol
            ligne = ligne & Format(arr(i, j), "000") & " "
        Next j
        Debug.Print ligne
    Next i
    Debug.Print ""
End Sub

Sub UtiliserAffichage()
    Dim test(1 To 3, 1 To 4) As Integer

    ' Remplissage
    Dim i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 4
            test(i, j) = i * 10 + j
        Next j
    Next i

    Call AfficherTableau2D(test, "Tableau de test")
End Sub
```

---

## Récapitulatif

### Points clés à retenir

1. **Tableau 2D** : Grille avec lignes et colonnes `(ligne, colonne)`
2. **Tableau 3D+** : Ajout de dimensions supplémentaires `(ligne, colonne, couche)`
3. **Navigation** : Boucles imbriquées dans l'ordre des dimensions
4. **ReDim Preserve** : Seule la dernière dimension peut être modifiée
5. **Performance** : Ordre des boucles important, limiter les dimensions

### Syntaxes essentielles

```vba
' Déclaration 2D
Dim arr2D(1 To lignes, 1 To colonnes) As Type

' Déclaration 3D
Dim arr3D(1 To x, 1 To y, 1 To z) As Type

' Parcours 2D
For i = 1 To UBound(arr2D, 1)
    For j = 1 To UBound(arr2D, 2)
        ' Traitement arr2D(i, j)
    Next j
Next i

' Redimensionnement dynamique (dernière dimension seulement)
ReDim Preserve arr2D(1 To lignes, 1 To nouvellesColonnes)
```

### Cas d'usage recommandés

| Dimensions | Utilisation typique | Exemple |
|------------|-------------------|---------|
| **2D** | Grilles, matrices, tableaux Excel | `donnees(ligne, colonne)` |
| **3D** | Données par période, région, catégorie | `ventes(mois, region, produit)` |
| **4D+** | Éviter si possible, préférer structures alternatives | - |

### Modèles de code recommandés

#### Modèle 2D standard
```vba
Sub Modele2D()
    Dim donnees(1 To nbLignes, 1 To nbColonnes) As Type

    ' Remplissage
    Dim i As Integer, j As Integer
    For i = 1 To nbLignes
        For j = 1 To nbColonnes
            donnees(i, j) = valeur
        Next j
    Next i

    ' Utilisation
    valeur = donnees(ligne, colonne)
End Sub
```

#### Modèle 3D pour analyses
```vba
Sub Modele3D()
    Dim analyses(1 To periodes, 1 To categories, 1 To metriques) As Double

    ' Remplissage
    Dim p As Integer, c As Integer, m As Integer
    For p = 1 To periodes
        For c = 1 To categories
            For m = 1 To metriques
                analyses(p, c, m) = calculer(p, c, m)
            Next m
        Next c
    Next p
End Sub
```

---

## Intégration avec Excel

### Lecture depuis Excel vers tableau multidimensionnel

```vba
Sub LireDepuisExcel()
    ' Lire une plage Excel dans un tableau 2D
    Dim plageSource As Range
    Set plageSource = Range("A1:D10")  ' 10 lignes x 4 colonnes

    ' Méthode 1 : Lecture directe avec Variant
    Dim donneesVariant As Variant
    donneesVariant = plageSource.Value  ' Automatiquement 2D

    ' Les indices commencent à 1 pour les tableaux issus d'Excel
    Debug.Print "Cellule A1 : " & donneesVariant(1, 1)
    Debug.Print "Cellule D10 : " & donneesVariant(10, 4)

    ' Méthode 2 : Lecture cellule par cellule vers tableau typé
    Dim donneesTypees(1 To 10, 1 To 4) As String
    Dim i As Integer, j As Integer

    For i = 1 To 10
        For j = 1 To 4
            donneesTypees(i, j) = plageSource.Cells(i, j).Value
        Next j
    Next i

    Debug.Print "Données lues depuis Excel"
End Sub
```

### Écriture depuis tableau vers Excel

```vba
Sub EcrireVersExcel()
    ' Créer un tableau 2D
    Dim resultats(1 To 5, 1 To 3) As Variant

    ' Remplir avec des données
    Dim i As Integer, j As Integer
    For i = 1 To 5
        For j = 1 To 3
            resultats(i, j) = "L" & i & "C" & j
        Next j
    Next i

    ' Méthode 1 : Écriture directe (très rapide)
    Range("F1:H5").Value = resultats

    ' Méthode 2 : Écriture cellule par cellule (plus lent)
    For i = 1 To 5
        For j = 1 To 3
            Cells(i + 10, j + 5).Value = resultats(i, j)
        Next j
    Next i

    Debug.Print "Données écrites vers Excel"
End Sub
```

### Traitement de grandes plages Excel

```vba
Sub TraiterGrandePlage()
    ' Pour de très grandes plages, utiliser un tableau est beaucoup plus rapide
    Dim debut As Double
    debut = Timer

    ' Lire toute la plage d'un coup
    Dim donneesExcel As Variant
    donneesExcel = Range("A1:Z1000").Value  ' 1000 lignes x 26 colonnes

    Debug.Print "Lecture terminée en " & Format(Timer - debut, "0.000") & " secondes"

    ' Traitement en mémoire (très rapide)
    debut = Timer
    Dim i As Long, j As Integer
    For i = 1 To 1000
        For j = 1 To 26
            If IsNumeric(donneesExcel(i, j)) Then
                donneesExcel(i, j) = donneesExcel(i, j) * 1.1  ' Augmentation de 10%
            End If
        Next j
    Next i

    Debug.Print "Traitement terminé en " & Format(Timer - debut, "0.000") & " secondes"

    ' Réécriture d'un coup
    debut = Timer
    Range("A1:Z1000").Value = donneesExcel
    Debug.Print "Écriture terminée en " & Format(Timer - debut, "0.000") & " secondes"
End Sub
```

---

## Techniques spécialisées

### 1. Transposition de tableaux 2D

```vba
Function TransposerTableau(original As Variant) As Variant
    ' Transposer un tableau 2D (lignes ↔ colonnes)
    Dim lignesOrig As Integer, colonnesOrig As Integer
    lignesOrig = UBound(original, 1) - LBound(original, 1) + 1
    colonnesOrig = UBound(original, 2) - LBound(original, 2) + 1

    ' Créer le tableau transposé
    Dim transpose() As Variant
    ReDim transpose(1 To colonnesOrig, 1 To lignesOrig)

    ' Copier en inversant les indices
    Dim i As Integer, j As Integer
    For i = 1 To lignesOrig
        For j = 1 To colonnesOrig
            transpose(j, i) = original(LBound(original, 1) + i - 1, LBound(original, 2) + j - 1)
        Next j
    Next i

    TransposerTableau = transpose
End Function

Sub UtiliserTransposition()
    ' Tableau original 3x2
    Dim original(1 To 3, 1 To 2) As String
    original(1, 1) = "A": original(1, 2) = "B"
    original(2, 1) = "C": original(2, 2) = "D"
    original(3, 1) = "E": original(3, 2) = "F"

    Debug.Print "Original (3x2):"
    Call AfficherTableau2D(original, "Original")

    ' Transposer
    Dim transpose As Variant
    transpose = TransposerTableau(original)

    Debug.Print "Transposé (2x3):"
    Call AfficherTableau2D(transpose, "Transposé")
End Sub
```

### 2. Recherche avancée dans tableaux multidimensionnels

```vba
Function RechercherDansTableau2D(tableau As Variant, valeurCherchee As Variant, _
                                 ByRef ligneResultat As Integer, ByRef colonneResultat As Integer) As Boolean
    ' Recherche une valeur dans un tableau 2D et retourne sa position

    Dim i As Integer, j As Integer
    For i = LBound(tableau, 1) To UBound(tableau, 1)
        For j = LBound(tableau, 2) To UBound(tableau, 2)
            If tableau(i, j) = valeurCherchee Then
                ligneResultat = i
                colonneResultat = j
                RechercherDansTableau2D = True
                Exit Function
            End If
        Next j
    Next i

    RechercherDansTableau2D = False
End Function

Sub UtiliserRechercheAvancee()
    Dim donnees(1 To 4, 1 To 3) As Integer

    ' Remplissage
    Dim i As Integer, j As Integer
    For i = 1 To 4
        For j = 1 To 3
            donnees(i, j) = i * 10 + j
        Next j
    Next i

    ' Recherche
    Dim ligne As Integer, colonne As Integer
    If RechercherDansTableau2D(donnees, 23, ligne, colonne) Then
        Debug.Print "Valeur 23 trouvée en position (" & ligne & ", " & colonne & ")"
    Else
        Debug.Print "Valeur 23 non trouvée"
    End If
End Sub
```

### 3. Agrégation de données multidimensionnelles

```vba
Sub AgregationDonnees()
    ' Simulation de données de ventes 3D : mois x région x produit
    Dim ventes(1 To 12, 1 To 5, 1 To 3) As Double

    ' Remplissage avec données aléatoires
    Dim mois As Integer, region As Integer, produit As Integer
    For mois = 1 To 12
        For region = 1 To 5
            For produit = 1 To 3
                ventes(mois, region, produit) = Rnd() * 1000 + 100
            Next produit
        Next region
    Next mois

    ' Agrégation 1 : Total par mois (toutes régions, tous produits)
    Dim totalMois(1 To 12) As Double
    For mois = 1 To 12
        For region = 1 To 5
            For produit = 1 To 3
                totalMois(mois) = totalMois(mois) + ventes(mois, region, produit)
            Next produit
        Next region
        Debug.Print "Mois " & mois & ": " & Format(totalMois(mois), "0.00") & "€"
    Next mois

    ' Agrégation 2 : Total par région (tous mois, tous produits)
    Dim totalRegion(1 To 5) As Double
    For region = 1 To 5
        For mois = 1 To 12
            For produit = 1 To 3
                totalRegion(region) = totalRegion(region) + ventes(mois, region, produit)
            Next produit
        Next mois
        Debug.Print "Région " & region & ": " & Format(totalRegion(region), "0.00") & "€"
    Next region

    ' Agrégation 3 : Moyenne par produit
    Dim moyenneProduit(1 To 3) As Double
    For produit = 1 To 3
        Dim somme As Double, compte As Integer
        somme = 0: compte = 0
        For mois = 1 To 12
            For region = 1 To 5
                somme = somme + ventes(mois, region, produit)
                compte = compte + 1
            Next region
        Next mois
        moyenneProduit(produit) = somme / compte
        Debug.Print "Produit " & produit & " - Moyenne: " & Format(moyenneProduit(produit), "0.00") & "€"
    Next produit
End Sub
```

---

## Debugging et visualisation

### Techniques de débogage pour tableaux multidimensionnels

```vba
Sub DebugTableauxMultidimensionnels()
    Dim donnees(1 To 3, 1 To 4, 1 To 2) As Integer

    ' Remplissage pour test
    Dim i As Integer, j As Integer, k As Integer
    For i = 1 To 3
        For j = 1 To 4
            For k = 1 To 2
                donnees(i, j, k) = i * 100 + j * 10 + k
            Next k
        Next j
    Next i

    ' Technique 1 : Affichage avec Debug.Print
    Debug.Print "=== DÉBOGAGE TABLEAU 3D ==="
    For k = 1 To 2
        Debug.Print "--- Couche " & k & " ---"
        For i = 1 To 3
            Dim ligne As String
            ligne = ""
            For j = 1 To 4
                ligne = ligne & Format(donnees(i, j, k), "000") & " "
            Next j
            Debug.Print ligne
        Next i
        Debug.Print ""
    Next k

    ' Technique 2 : Vérification des limites
    Debug.Print "Limites du tableau :"
    Debug.Print "Dimension 1 : " & LBound(donnees, 1) & " à " & UBound(donnees, 1)
    Debug.Print "Dimension 2 : " & LBound(donnees, 2) & " à " & UBound(donnees, 2)
    Debug.Print "Dimension 3 : " & LBound(donnees, 3) & " à " & UBound(donnees, 3)

    ' Technique 3 : Points de contrôle spécifiques
    Debug.Print "Valeurs de contrôle :"
    Debug.Print "donnees(1,1,1) = " & donnees(1, 1, 1)  ' Devrait être 111
    Debug.Print "donnees(3,4,2) = " & donnees(3, 4, 2)  ' Devrait être 342
End Sub
```

### Fonction générique d'affichage 3D

```vba
Sub AfficherTableau3D(arr As Variant, titre As String)
    Debug.Print "=== " & titre & " ==="

    Dim k As Integer, i As Integer, j As Integer
    For k = LBound(arr, 3) To UBound(arr, 3)
        Debug.Print "--- Couche " & k & " ---"
        For i = LBound(arr, 1) To UBound(arr, 1)
            Dim ligne As String
            ligne = ""
            For j = LBound(arr, 2) To UBound(arr, 2)
                ligne = ligne & Format(arr(i, j, k), "000") & " "
            Next j
            Debug.Print ligne
        Next i
        Debug.Print ""
    Next k
End Sub
```

---

## Erreurs courantes et solutions

### 1. Erreur d'indices

```vba
Sub GestionErreursIndices()
    Dim tableau(1 To 5, 1 To 3) As Integer

    ' ❌ ERREUR : Indice hors limites
    ' tableau(6, 1) = 100  ' Erreur : subscript out of range

    ' ✅ SOLUTION : Vérification des limites
    Dim ligne As Integer, colonne As Integer
    ligne = 6: colonne = 1

    If ligne >= LBound(tableau, 1) And ligne <= UBound(tableau, 1) And _
       colonne >= LBound(tableau, 2) And colonne <= UBound(tableau, 2) Then
        tableau(ligne, colonne) = 100
        Debug.Print "Valeur assignée"
    Else
        Debug.Print "Indices (" & ligne & ", " & colonne & ") hors limites"
        Debug.Print "Limites valides : (" & LBound(tableau, 1) & "-" & UBound(tableau, 1) & _
                    ", " & LBound(tableau, 2) & "-" & UBound(tableau, 2) & ")"
    End If
End Sub
```

### 2. Confusion dans l'ordre des dimensions

```vba
Sub ConfusionDimensions()
    ' Tableau représentant des données Excel : lignes x colonnes
    Dim donneesExcel(1 To 10, 1 To 5) As String  ' 10 lignes, 5 colonnes

    ' ❌ ERREUR COURANTE : Confondre ligne et colonne
    ' Pour accéder à la ligne 3, colonne 2 :
    ' Incorrect : donneesExcel(2, 3)

    ' ✅ CORRECT : ligne en premier, colonne en second
    donneesExcel(3, 2) = "Ligne 3, Colonne 2"

    ' Aide-mémoire : pensez "adresse Excel"
    ' B3 = ligne 3, colonne 2 = donneesExcel(3, 2)

    Debug.Print "Convention : tableau(ligne, colonne) comme Excel"
End Sub
```

### 3. Problèmes avec ReDim Preserve

```vba
Sub ProblemeRedimPreserve()
    Dim tableau() As Integer
    ReDim tableau(1 To 3, 1 To 2)
    tableau(1, 1) = 10: tableau(3, 2) = 32

    ' ❌ ERREUR : Essayer de modifier la première dimension
    On Error GoTo ErreurRedim
    ReDim Preserve tableau(1 To 5, 1 To 2)  ' ERREUR !
    Debug.Print "Redimensionnement réussi"
    Exit Sub

ErreurRedim:
    Debug.Print "Erreur ReDim Preserve : " & Err.Description
    Debug.Print "Solution : Seule la dernière dimension peut être modifiée"

    ' ✅ SOLUTION : Modifier seulement la dernière dimension
    On Error GoTo 0
    ReDim Preserve tableau(1 To 3, 1 To 4)  ' OK
    Debug.Print "Redimensionnement alternatif réussi"
    Debug.Print "Valeurs préservées - tableau(1,1) = " & tableau(1, 1)
    Debug.Print "Valeurs préservées - tableau(3,2) = " & tableau(3, 2)
End Sub
```

---

## Récapitulatif final

### Règles d'or des tableaux multidimensionnels

1. **Simplicité** : Préférer 2D ou 3D, éviter 4D+
2. **Convention** : Premier indice = lignes, second = colonnes (comme Excel)
3. **Performance** : Ordre des boucles important (ligne puis colonne)
4. **ReDim Preserve** : Seule la dernière dimension peut changer
5. **Vérification** : Toujours contrôler les limites avant accès

### Checklist de bonnes pratiques

- [ ] Dimensions clairement documentées dans les commentaires
- [ ] Noms de variables explicites pour les indices (`ligne`, `colonne` vs `i`, `j`)
- [ ] Vérification des limites avant accès aux éléments
- [ ] Ordre logique des boucles pour la performance
- [ ] Fonctions d'affichage pour le débogage
- [ ] Gestion d'erreurs pour les accès hors limites

### Quand utiliser les tableaux multidimensionnels

✅ **Utilisez les tableaux multidimensionnels pour :**
- Représenter des grilles de données (2D)
- Analyser des données par multiple critères (3D)
- Manipuler des matrices mathématiques
- Traiter de gros volumes de données Excel

❌ **Évitez les tableaux multidimensionnels pour :**
- Des structures de données complexes (préférer les classes)
- Plus de 3-4 dimensions (difficile à maintenir)
- Des données de types très différents (préférer des structures)

Les tableaux multidimensionnels sont des outils puissants pour organiser et traiter des données complexes. Maîtrisés correctement, ils vous permettront de créer des solutions VBA élégantes et performantes pour vos projets professionnels. Dans la section suivante, nous explorerons les techniques de parcours et manipulation qui vous donneront encore plus de contrôle sur vos données.

⏭️
