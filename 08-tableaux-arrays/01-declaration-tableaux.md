🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.1. Déclaration de tableaux

## Introduction à la déclaration de tableaux

Déclarer un tableau en VBA, c'est comme **réserver des casiers dans un vestiaire** : vous devez dire combien de casiers vous voulez, quel type d'objets vous allez y stocker, et comment ils seront numérotés. Cette déclaration est le fondement de tout travail avec les tableaux.

**Analogie simple :**
Imaginez que vous organisez un classement de films. Au lieu d'avoir 50 variables séparées (`film1`, `film2`, `film3`...), vous créez un **tableau** `films()` avec 50 emplacements numérotés. Chaque emplacement peut contenir le nom d'un film, et vous pouvez y accéder en précisant le numéro : `films(1)`, `films(2)`, etc.

---

## Syntaxe de base de déclaration

### Structure générale

```vba
Dim nomTableau(limites) As TypeDeDonnées
```

**Éléments de la syntaxe :**
- **Dim** : Mot-clé de déclaration (comme pour les variables normales)
- **nomTableau** : Le nom que vous donnez à votre tableau
- **(limites)** : Les indices minimum et maximum du tableau
- **As TypeDeDonnées** : Le type de données que contiendra le tableau

### Exemples de déclarations simples

```vba
Sub ExemplesDeclarationSimple()
    ' Tableau de 10 nombres entiers (indices 0 à 9)
    Dim nombres(9) As Integer

    ' Tableau de 5 noms (indices 0 à 4)
    Dim noms(4) As String

    ' Tableau de 12 valeurs décimales (indices 0 à 11)
    Dim moyennes(11) As Double

    ' Tableau de 7 dates (indices 0 à 6)
    Dim semaine(6) As Date

    ' Tableau de 100 valeurs booléennes (indices 0 à 99)
    Dim validations(99) As Boolean

    MsgBox "Tableaux déclarés avec succès !"
End Sub
```

---

## Comprendre les indices de tableaux

### Indices par défaut (0-based)

Par défaut, VBA commence la numérotation des tableaux à **0** :

```vba
Sub IndicesParDefaut()
    ' Déclaration d'un tableau de 5 éléments
    Dim couleurs(4) As String

    ' Les indices disponibles sont : 0, 1, 2, 3, 4
    couleurs(0) = "Rouge"
    couleurs(1) = "Vert"
    couleurs(2) = "Bleu"
    couleurs(3) = "Jaune"
    couleurs(4) = "Orange"

    ' Afficher le contenu
    Dim i As Integer
    For i = 0 To 4
        Debug.Print "couleurs(" & i & ") = " & couleurs(i)
    Next i
End Sub
```

### Indices personnalisés (1-based ou autres)

Vous pouvez spécifier vos propres limites d'indices :

```vba
Sub IndicesPersonnalises()
    ' Tableau avec indices de 1 à 10 (plus naturel)
    Dim notes(1 To 10) As Integer

    ' Tableau avec indices de 5 à 15
    Dim temperatures(5 To 15) As Double

    ' Tableau avec indices négatifs
    Dim variations(-5 To 5) As Integer

    ' Remplir le tableau des notes
    Dim i As Integer
    For i = 1 To 10
        notes(i) = i * 10  ' 10, 20, 30, ..., 100
        Debug.Print "Note " & i & " : " & notes(i)
    Next i
End Sub
```

### Option Base - Modifier la base par défaut

```vba
' À placer en début de module, avant toute procédure
Option Base 1

Sub AvecOptionBase()
    ' Maintenant les tableaux commencent à 1 par défaut
    Dim jours(7) As String    ' Indices de 1 à 7

    jours(1) = "Lundi"
    jours(2) = "Mardi"
    jours(3) = "Mercredi"
    jours(4) = "Jeudi"
    jours(5) = "Vendredi"
    jours(6) = "Samedi"
    jours(7) = "Dimanche"

    ' Affichage
    Dim i As Integer
    For i = 1 To 7
        Debug.Print jours(i)
    Next i
End Sub
```

---

## Types de données pour les tableaux

### Tableaux de types simples

```vba
Sub TypesSimples()
    ' Différents types de données
    Dim ages(1 To 5) As Integer
    Dim salaires(1 To 5) As Double
    Dim employes(1 To 5) As String
    Dim embauches(1 To 5) As Date
    Dim actifs(1 To 5) As Boolean

    ' Remplissage des données
    ages(1) = 25: employes(1) = "Alice": salaires(1) = 2500.5
    ages(2) = 30: employes(2) = "Bob": salaires(2) = 3200.75
    ages(3) = 35: employes(3) = "Claire": salaires(3) = 4100.25

    ' Affichage
    Dim i As Integer
    For i = 1 To 3
        Debug.Print employes(i) & " - " & ages(i) & " ans - " & salaires(i) & "€"
    Next i
End Sub
```

### Tableaux Variant (type flexible)

```vba
Sub TableauxVariant()
    ' Tableau Variant peut contenir n'importe quel type
    Dim donneesMixtes(1 To 6) As Variant

    donneesMixtes(1) = "Texte"
    donneesMixtes(2) = 123
    donneesMixtes(3) = 45.67
    donneesMixtes(4) = #1/1/2024#  ' Date
    donneesMixtes(5) = True
    donneesMixtes(6) = Array("sous", "tableau")  ' Même un autre tableau !

    ' Affichage avec vérification de type
    Dim i As Integer
    For i = 1 To 6
        Debug.Print "Element " & i & " : " & donneesMixtes(i) & " (Type: " & TypeName(donneesMixtes(i)) & ")"
    Next i
End Sub
```

### Tableaux d'objets

```vba
Sub TableauxObjets()
    ' Tableau contenant des objets Excel
    Dim feuilles(1 To 3) As Worksheet
    Dim plages(1 To 5) As Range

    ' Attention : il faut utiliser Set pour les objets
    Set feuilles(1) = ActiveSheet
    Set plages(1) = Range("A1")
    Set plages(2) = Range("B1:B5")

    ' Utilisation
    feuilles(1).Name = "Feuille modifiée"
    plages(1).Value = "Valeur dans A1"

    MsgBox "Objets manipulés via le tableau"
End Sub
```

---

## Déclarations avancées

### Tableaux multidimensionnels

```vba
Sub TableauxMultidimensionnels()
    ' Tableau à 2 dimensions (comme une grille Excel)
    Dim grille(1 To 3, 1 To 4) As Integer

    ' Tableau à 3 dimensions
    Dim cube(1 To 2, 1 To 3, 1 To 4) As String

    ' Remplissage de la grille 2D
    Dim ligne As Integer, colonne As Integer
    For ligne = 1 To 3
        For colonne = 1 To 4
            grille(ligne, colonne) = ligne * 10 + colonne
        Next colonne
    Next ligne

    ' Affichage de la grille
    For ligne = 1 To 3
        Dim ligneTexte As String
        ligneTexte = ""
        For colonne = 1 To 4
            ligneTexte = ligneTexte & grille(ligne, colonne) & vbTab
        Next colonne
        Debug.Print ligneTexte
    Next ligne

    ' Résultat affiché :
    ' 11    12    13    14
    ' 21    22    23    24
    ' 31    32    33    34
End Sub
```

### Tableaux dynamiques (déclaration sans taille)

```vba
Sub TableauxDynamiques()
    ' Déclaration sans spécifier la taille
    Dim donneesDynamiques() As Double
    Dim tableauTexte() As String

    ' La taille sera définie plus tard avec ReDim
    ReDim donneesDynamiques(1 To 10)
    ReDim tableauTexte(0 To 5)

    ' Maintenant on peut les utiliser
    donneesDynamiques(1) = 3.14
    donneesDynamiques(2) = 2.71
    tableauTexte(0) = "Premier élément"

    MsgBox "Tableaux dynamiques redimensionnés et utilisés"
End Sub
```

---

## Conventions de nommage

### Bonnes pratiques pour nommer les tableaux

```vba
Sub ConventionsNommage()
    ' ✅ BONS noms - descriptifs et clairs
    Dim notesEtudiants(1 To 30) As Integer
    Dim nomsClients(1 To 100) As String
    Dim ventesParMois(1 To 12) As Double
    Dim temperaturesSemaine(1 To 7) As Single

    ' ✅ Utilisation de préfixes pour les tableaux
    Dim arrNotes(1 To 30) As Integer        ' arr = array
    Dim tabVentes(1 To 12) As Double        ' tab = tableau

    ' ❌ MAUVAIS noms - peu descriptifs
    Dim a(10) As Integer      ' Que contient-il ?
    Dim donnees(5) As String  ' Trop vague
    Dim x(1 To 3) As Double   ' Incompréhensible

    MsgBox "Exemples de conventions de nommage"
End Sub
```

### Nommage selon le contenu

```vba
Sub NommageSelenContenu()
    ' Selon le type de données
    Dim listeEmail(1 To 50) As String
    Dim compteurErreurs(1 To 10) As Integer
    Dim pourcentageReussite(1 To 20) As Double

    ' Selon l'utilisation
    Dim donneesEntree(1 To 100) As Variant
    Dim resultatsCalcul(1 To 100) As Double
    Dim parametresConfig(1 To 5) As String

    ' Selon la source
    Dim valeursExcel(1 To 1000) As Variant
    Dim donneesUtilisateur(1 To 10) As String
    Dim parametresFichier(1 To 20) As String

    MsgBox "Tableaux nommés selon leur utilisation"
End Sub
```

---

## Initialisation lors de la déclaration

### Valeurs par défaut

Quand vous déclarez un tableau, VBA l'initialise automatiquement avec des valeurs par défaut :

```vba
Sub ValeursParDefaut()
    Dim nombres(1 To 5) As Integer      ' Initialisé avec des 0
    Dim textes(1 To 5) As String        ' Initialisé avec des chaînes vides ""
    Dim decimales(1 To 5) As Double     ' Initialisé avec des 0.0
    Dim booleens(1 To 5) As Boolean     ' Initialisé avec False
    Dim dates(1 To 5) As Date           ' Initialisé avec 30/12/1899

    ' Vérification des valeurs par défaut
    Debug.Print "Integer par défaut : " & nombres(1)     ' 0
    Debug.Print "String par défaut : '" & textes(1) & "'" ' ""
    Debug.Print "Double par défaut : " & decimales(1)    ' 0
    Debug.Print "Boolean par défaut : " & booleens(1)    ' False
    Debug.Print "Date par défaut : " & dates(1)          ' 30/12/1899
End Sub
```

### Initialisation immédiate

```vba
Sub InitialisationImmediate()
    ' Déclaration et initialisation en une fois avec Array()
    Dim jours As Variant
    jours = Array("Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche")

    ' Attention : Array() crée un tableau base 0
    Debug.Print jours(0)  ' "Lundi"
    Debug.Print jours(6)  ' "Dimanche"

    ' Pour un tableau typé, déclaration puis initialisation
    Dim notes(1 To 5) As Integer
    notes(1) = 15: notes(2) = 12: notes(3) = 18: notes(4) = 14: notes(5) = 16

    ' Ou avec une boucle d'initialisation
    Dim carres(1 To 10) As Integer
    Dim i As Integer
    For i = 1 To 10
        carres(i) = i * i  ' 1, 4, 9, 16, 25...
    Next i
End Sub
```

---

## Vérification et informations sur les tableaux

### Fonctions utiles pour les tableaux

```vba
Sub InformationsTableaux()
    Dim nombres(5 To 15) As Integer
    Dim textes(1 To 20) As String

    ' UBound : indice maximum
    Debug.Print "Indice max de nombres : " & UBound(nombres)  ' 15
    Debug.Print "Indice max de textes : " & UBound(textes)    ' 20

    ' LBound : indice minimum
    Debug.Print "Indice min de nombres : " & LBound(nombres)  ' 5
    Debug.Print "Indice min de textes : " & LBound(textes)    ' 1

    ' Calcul du nombre d'éléments
    Dim nbElements As Integer
    nbElements = UBound(nombres) - LBound(nombres) + 1
    Debug.Print "Nombre d'éléments dans nombres : " & nbElements  ' 11

    ' Vérifier si un tableau est initialisé
    Dim tableauVide() As String
    On Error Resume Next
    Dim test As Integer
    test = UBound(tableauVide)
    If Err.Number <> 0 Then
        Debug.Print "Le tableau tableauVide n'est pas initialisé"
        Err.Clear
    End If
    On Error GoTo 0
End Sub
```

### Fonction pour afficher un tableau

```vba
Sub AfficherTableau(arr As Variant, nomTableau As String)
    ' Fonction utilitaire pour afficher le contenu d'un tableau
    Debug.Print "=== Contenu du tableau " & nomTableau & " ==="

    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Debug.Print nomTableau & "(" & i & ") = " & arr(i)
    Next i

    Debug.Print "=== Fin du tableau ==="
End Sub

Sub ExempleAffichage()
    Dim fruits(1 To 4) As String
    fruits(1) = "Pomme"
    fruits(2) = "Banane"
    fruits(3) = "Orange"
    fruits(4) = "Kiwi"

    Call AfficherTableau(fruits, "fruits")
End Sub
```

---

## Erreurs courantes lors de la déclaration

### Erreur 1 : Accès hors limites

```vba
Sub ErreurHorsLimites()
    Dim nombres(1 To 5) As Integer

    ' ❌ ERREUR : L'indice 6 n'existe pas
    ' nombres(6) = 100  ' Provoque "Subscript out of range"

    ' ✅ CORRECT : Vérifier les limites avant l'accès
    Dim indice As Integer
    indice = 6

    If indice >= LBound(nombres) And indice <= UBound(nombres) Then
        nombres(indice) = 100
    Else
        Debug.Print "Indice " & indice & " hors limites pour ce tableau"
    End If
End Sub
```

### Erreur 2 : Oublier de redimensionner un tableau dynamique

```vba
Sub ErreurTableauDynamique()
    Dim donnees() As Integer

    ' ❌ ERREUR : Tableau pas encore dimensionné
    ' donnees(1) = 100  ' Provoque une erreur

    ' ✅ CORRECT : Redimensionner d'abord
    ReDim donnees(1 To 10)
    donnees(1) = 100  ' Maintenant ça fonctionne

    Debug.Print "Tableau redimensionné et utilisé correctement"
End Sub
```

### Erreur 3 : Confusion entre types

```vba
Sub ErreurTypes()
    Dim nombres(1 To 5) As Integer

    ' ❌ PROBLÈME : Perte de précision
    nombres(1) = 3.14  ' Devient 3 (arrondi, pas troncature)

    ' ❌ ERREUR : Type incompatible
    ' nombres(2) = "Texte"  ' Provoque "Type mismatch"

    ' ✅ CORRECT : Utiliser le bon type ou Variant
    Dim donneesMixtes(1 To 5) As Variant
    donneesMixtes(1) = 3.14
    donneesMixtes(2) = "Texte"
    donneesMixtes(3) = True

    Debug.Print "Types gérés correctement"
End Sub
```

---

## Conseils et bonnes pratiques

### 1. Choisir le bon type de données

```vba
Sub ChoisirBonType()
    ' Pour des entiers : Integer (-32768 à 32767) ou Long (-2 milliards à +2 milliards)
    Dim petitsNombres(1 To 100) As Integer
    Dim grandsNombres(1 To 100) As Long

    ' Pour des décimaux : Single (précision simple) ou Double (précision double)
    Dim coordonnees(1 To 50) As Single     ' Suffisant pour la plupart des cas
    Dim calculsPrecis(1 To 50) As Double   ' Pour les calculs financiers

    ' Pour du texte : String
    Dim descriptions(1 To 25) As String

    ' Quand le type varie : Variant (mais plus lent)
    Dim donneesMixtes(1 To 10) As Variant
End Sub
```

### 2. Commentaires et documentation

```vba
Sub BienDocumenter()
    ' Tableau des notes d'étudiants (indices 1 à 30 pour 30 étudiants)
    Dim notesEtudiants(1 To 30) As Integer

    ' Tableau des ventes mensuelles (indices 1 à 12 pour les 12 mois)
    Dim ventesParMois(1 To 12) As Double

    ' Tableau dynamique pour stocker les résultats de calcul
    ' Sera redimensionné selon le nombre de lignes de données
    Dim resultatsCalcul() As Double

    ' Matrice 2D pour représenter une grille de jeu (10x10)
    Dim grilleJeu(1 To 10, 1 To 10) As String
End Sub
```

### 3. Initialisation systématique

```vba
Sub InitialisationSystematique()
    Dim scores(1 To 10) As Integer
    Dim noms(1 To 10) As String

    ' Initialiser avec des valeurs par défaut explicites
    Dim i As Integer
    For i = 1 To 10
        scores(i) = 0           ' Score initial
        noms(i) = "Inconnu"     ' Nom par défaut
    Next i

    Debug.Print "Tableaux initialisés avec des valeurs par défaut"
End Sub
```

---

## Récapitulatif

### Points clés à retenir

1. **Syntaxe de base** : `Dim nomTableau(limites) As Type`
2. **Indices par défaut** : Commencent à 0, sauf avec `Option Base 1`
3. **Indices personnalisés** : `(1 To 10)`, `(5 To 15)`, `(-5 To 5)`
4. **Types de données** : Tous les types VBA sont supportés
5. **Tableaux dynamiques** : Déclarés avec `()`, dimensionnés avec `ReDim`
6. **Fonctions utiles** : `LBound()`, `UBound()` pour connaître les limites

### Modèle de déclaration recommandé

```vba
Sub ModeleDeclaration()
    ' 1. Commentaire explicatif
    ' Tableau des températures quotidiennes (1 à 31 pour les jours du mois)

    ' 2. Déclaration avec limites explicites
    Dim temperaturesJour(1 To 31) As Single

    ' 3. Initialisation si nécessaire
    Dim i As Integer
    For i = 1 To 31
        temperaturesJour(i) = 0.0  ' Température par défaut
    Next i

    ' 4. Utilisation avec vérification des limites
    Dim jour As Integer
    jour = 15

    If jour >= LBound(temperaturesJour) And jour <= UBound(temperaturesJour) Then
        temperaturesJour(jour) = 23.5
        Debug.Print "Température du jour " & jour & " : " & temperaturesJour(jour) & "°C"
    End If
End Sub
```

### Erreurs à éviter

- ❌ Accéder à des indices hors limites
- ❌ Oublier de redimensionner les tableaux dynamiques
- ❌ Utiliser des noms de tableaux peu descriptifs
- ❌ Ne pas initialiser les tableaux quand nécessaire
- ❌ Mélanger les types de données sans utiliser Variant

### Prochaine étape

Maintenant que vous savez déclarer des tableaux, la section suivante vous apprendra la différence entre tableaux statiques et dynamiques, et quand utiliser chaque type pour optimiser vos programmes.

⏭️ [Tableaux statiques et dynamiques](/08-tableaux-arrays/02-tableaux-statiques-dynamiques.md)
