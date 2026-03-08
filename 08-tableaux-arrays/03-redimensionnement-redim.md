🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.3. Redimensionnement (ReDim)

## Introduction à ReDim

L'instruction **ReDim** (redimensionnement) est l'outil magique qui permet de changer la taille des tableaux dynamiques pendant l'exécution de votre programme. C'est comme avoir un **accordéon numérique** : vous pouvez l'étendre ou le rétrécir selon vos besoins du moment.

**Analogie simple :**
Imaginez un sac de voyage extensible :
- **ReDim normal** = Vous videz complètement le sac puis le redimensionnez (perdez le contenu)
- **ReDim Preserve** = Vous agrandissez le sac en gardant tout ce qui s'y trouve déjà
- L'objectif est d'adapter la taille exactement à ce que vous voulez transporter

Cette instruction est fondamentale pour créer des programmes VBA qui s'adaptent dynamiquement aux données qu'ils traitent.

---

## Syntaxe de base de ReDim

### Structure générale

```vba
ReDim [Preserve] nomTableau(nouvelleDimension) [As TypeDonnées]
```

**Éléments de la syntaxe :**
- **ReDim** : Mot-clé de redimensionnement
- **Preserve** (optionnel) : Garde les données existantes
- **nomTableau** : Le tableau à redimensionner
- **nouvelleDimension** : Les nouvelles limites du tableau
- **As TypeDonnées** (optionnel) : Le type peut être spécifié lors du premier ReDim

### Exemples de base

```vba
Sub ExemplesReDimBase()
    ' Déclaration d'un tableau dynamique
    Dim nombres() As Integer

    ' Premier redimensionnement (obligatoire pour un tableau dynamique)
    ReDim nombres(1 To 5)

    ' Remplissage
    nombres(1) = 10
    nombres(2) = 20
    nombres(3) = 30
    nombres(4) = 40
    nombres(5) = 50

    Debug.Print "Première taille : " & UBound(nombres)  ' 5

    ' Redimensionnement plus grand (perte des données)
    ReDim nombres(1 To 10)
    Debug.Print "Après ReDim normal - nombres(1) = " & nombres(1)  ' 0 (données perdues)

    ' Remplissage à nouveau
    nombres(1) = 100
    nombres(10) = 1000

    ' Redimensionnement avec préservation
    ReDim Preserve nombres(1 To 15)
    Debug.Print "Après ReDim Preserve - nombres(1) = " & nombres(1)  ' 100 (données gardées)
    Debug.Print "Nouvelle taille : " & UBound(nombres)  ' 15
End Sub
```

---

## ReDim sans Preserve (redimensionnement destructif)

### Comportement et utilisation

Quand vous utilisez **ReDim** sans le mot-clé **Preserve**, VBA :
1. Détruit complètement le tableau existant
2. Crée un nouveau tableau avec les nouvelles dimensions
3. Initialise tous les éléments avec des valeurs par défaut

```vba
Sub ReDimDestructif()
    Dim scores() As Integer

    ' Première allocation
    ReDim scores(1 To 3)
    scores(1) = 85
    scores(2) = 92
    scores(3) = 78

    Debug.Print "=== AVANT REDIM ==="
    Debug.Print "scores(1) = " & scores(1)  ' 85
    Debug.Print "scores(2) = " & scores(2)  ' 92
    Debug.Print "scores(3) = " & scores(3)  ' 78
    Debug.Print "Taille : " & UBound(scores)  ' 3

    ' Redimensionnement destructif
    ReDim scores(1 To 6)

    Debug.Print "=== APRÈS REDIM ==="
    Debug.Print "scores(1) = " & scores(1)  ' 0 (donnée perdue !)
    Debug.Print "scores(2) = " & scores(2)  ' 0 (donnée perdue !)
    Debug.Print "scores(6) = " & scores(6)  ' 0 (nouveau élément)
    Debug.Print "Nouvelle taille : " & UBound(scores)  ' 6
End Sub
```

### Quand utiliser ReDim destructif

#### 1. **Réinitialisation complète**

```vba
Sub ReinitialisationComplete()
    Dim donnees() As Double

    ' Première utilisation
    ReDim donnees(1 To 5)
    donnees(1) = 1.5: donnees(2) = 2.7: donnees(3) = 3.9

    Debug.Print "Données initiales : " & donnees(1) & ", " & donnees(2) & ", " & donnees(3)

    ' Besoin de repartir à zéro avec une nouvelle taille
    ReDim donnees(1 To 10)  ' Efface tout et repart proprement

    Debug.Print "Après réinitialisation : " & donnees(1)  ' 0 (propre)
End Sub
```

#### 2. **Changement radical de structure**

```vba
Sub ChangementStructure()
    Dim tableau() As String

    ' Utilisation temporaire
    ReDim tableau(0 To 4)  ' Base 0
    tableau(0) = "Zéro": tableau(4) = "Quatre"

    ' Changement vers base 1 (plus logique)
    ReDim tableau(1 To 5)  ' Complètement différent
    tableau(1) = "Un": tableau(5) = "Cinq"

    Debug.Print "Nouvelle structure : " & tableau(1) & " à " & tableau(5)
End Sub
```

#### 3. **Performance : éviter la fragmentation**

```vba
Sub OptimisationPerformance()
    Dim buffer() As Integer

    ' Utilisation intensive avec beaucoup de Preserve (fragmentation)
    ReDim buffer(1 To 1)
    ' ... beaucoup de ReDim Preserve successifs ...

    ' À un moment, mieux vaut tout recréer proprement
    Dim tailleFinale As Integer
    tailleFinale = 1000

    ReDim buffer(1 To tailleFinale)  ' Recrée proprement

    Debug.Print "Buffer recréé avec " & UBound(buffer) & " éléments"
End Sub
```

---

## ReDim Preserve (redimensionnement conservateur)

### Comportement et règles

**ReDim Preserve** garde les données existantes mais a des **limitations importantes** :

1. **Tableaux 1D** : Peut changer seulement la limite supérieure
2. **Tableaux multidimensionnels** : Peut changer seulement la dernière dimension
3. **Nouvelle taille plus petite** : Les éléments "en trop" sont perdus

```vba
Sub ReDimPreserveRegles()
    ' === TABLEAU 1D ===
    Dim liste() As String
    ReDim liste(1 To 3)
    liste(1) = "Premier"
    liste(2) = "Deuxième"
    liste(3) = "Troisième"

    ' ✅ AUTORISÉ : Changer la limite supérieure
    ReDim Preserve liste(1 To 5)
    liste(4) = "Quatrième"
    liste(5) = "Cinquième"
    Debug.Print "Agrandissement OK : " & liste(1) & ", " & liste(4)

    ' ❌ INTERDIT : Changer la limite inférieure
    ' ReDim Preserve liste(0 To 5)  ' ERREUR !

    ' === TABLEAU 2D ===
    Dim grille() As Integer
    ReDim grille(1 To 2, 1 To 3)
    grille(1, 1) = 11: grille(2, 3) = 23

    ' ✅ AUTORISÉ : Changer seulement la dernière dimension
    ReDim Preserve grille(1 To 2, 1 To 5)
    grille(1, 4) = 14
    Debug.Print "2D - Dernière dimension changée : " & grille(1, 1) & ", " & grille(1, 4)

    ' ❌ INTERDIT : Changer la première dimension
    ' ReDim Preserve grille(1 To 4, 1 To 5)  ' ERREUR !
End Sub
```

### Techniques d'agrandissement

#### 1. **Ajouter des éléments un par un**

```vba
Sub AjouterElementsUnParUn()
    Dim noms() As String
    Dim taille As Integer
    taille = 0

    ' Ajouter des noms un par un
    taille = taille + 1
    ReDim Preserve noms(1 To taille)
    noms(taille) = "Alice"

    taille = taille + 1
    ReDim Preserve noms(1 To taille)
    noms(taille) = "Bob"

    taille = taille + 1
    ReDim Preserve noms(1 To taille)
    noms(taille) = "Claire"

    Debug.Print "Nombre de noms : " & taille
    Dim i As Integer
    For i = 1 To taille
        Debug.Print "Nom " & i & " : " & noms(i)
    Next i
End Sub
```

> **Note :** Pour éviter la répétition, créez une procédure séparée (voir section 8.2 pour un exemple avec `AjouterElement`).

#### 2. **Agrandissement par blocs (plus efficace)**

```vba
Sub AgrandissementParBlocs()
    Dim donnees() As Double
    Dim tailleUtilisee As Integer
    Dim tailleAllouee As Integer
    Dim tailleBloc As Integer
    Dim valeur As Double

    tailleUtilisee = 0
    tailleAllouee = 0
    tailleBloc = 5  ' Grandir par blocs de 5

    ' Simuler l'ajout de 6 valeurs
    Dim valeursAjout As Variant
    valeursAjout = Array(1.5, 2.7, 3.9, 4.2, 5.8, 6.1)

    Dim i As Integer
    For i = LBound(valeursAjout) To UBound(valeursAjout)
        tailleUtilisee = tailleUtilisee + 1

        ' Agrandir seulement si nécessaire
        If tailleUtilisee > tailleAllouee Then
            tailleAllouee = tailleAllouee + tailleBloc
            ReDim Preserve donnees(1 To tailleAllouee)
            Debug.Print "Agrandi à " & tailleAllouee & " éléments"
        End If

        donnees(tailleUtilisee) = valeursAjout(i)
    Next i

    Debug.Print "Utilisé : " & tailleUtilisee & " / Alloué : " & tailleAllouee
End Sub
```

### Techniques de rétrécissement

```vba
Sub RetrecissementTableau()
    Dim valeurs() As Integer

    ' Créer un tableau avec des données
    ReDim valeurs(1 To 10)
    Dim i As Integer
    For i = 1 To 10
        valeurs(i) = i * 10
    Next i

    Debug.Print "Taille originale : " & UBound(valeurs)  ' 10
    Debug.Print "valeurs(8) = " & valeurs(8)  ' 80

    ' Rétrécir le tableau (perd les éléments en trop)
    ReDim Preserve valeurs(1 To 6)

    Debug.Print "Nouvelle taille : " & UBound(valeurs)  ' 6
    Debug.Print "valeurs(6) = " & valeurs(6)  ' 60
    ' valeurs(8) n'existe plus !

    ' Vérification sécurisée
    If 8 <= UBound(valeurs) Then
        Debug.Print "valeurs(8) = " & valeurs(8)
    Else
        Debug.Print "valeurs(8) n'existe plus après rétrécissement"
    End If
End Sub
```

---

## Gestion des erreurs avec ReDim

### Erreurs courantes et solutions

#### 1. **Tableau non initialisé**

```vba
Sub GestionErreurNonInitialise()
    Dim tableau() As String

    ' ❌ ERREUR : Tableau pas encore initialisé
    ' tableau(1) = "Test"  ' Provoque une erreur

    ' ✅ SOLUTION : Vérifier et initialiser
    On Error Resume Next
    Dim test As Integer
    test = UBound(tableau)
    If Err.Number <> 0 Then
        ' Tableau non initialisé
        ReDim tableau(1 To 5)
        Debug.Print "Tableau initialisé"
        Err.Clear
    End If
    On Error GoTo 0

    ' Maintenant on peut l'utiliser
    tableau(1) = "Premier élément"
    Debug.Print tableau(1)
End Sub
```

#### 2. **ReDim Preserve avec changements interdits**

```vba
Sub GestionErreurPreserve()
    Dim grille() As Integer
    ReDim grille(1 To 3, 1 To 4)
    grille(1, 1) = 100

    ' Tentative de changement de première dimension
    On Error GoTo ErreurRedim
    ReDim Preserve grille(1 To 5, 1 To 4)  ' ERREUR !
    Debug.Print "Redimensionnement réussi"
    Exit Sub

ErreurRedim:
    Debug.Print "Erreur ReDim Preserve : " & Err.Description
    Debug.Print "Solution : Changer seulement la dernière dimension"

    ' Solution de contournement
    ReDim Preserve grille(1 To 3, 1 To 6)  ' OK
    Debug.Print "Redimensionnement alternatif réussi"
End Sub
```

#### 3. **Fonction sécurisée de redimensionnement**

```vba
Function RedimensionnerSecurise(ByRef arr() As Variant, nouvelleTaille As Integer) As Boolean
    On Error GoTo ErreurRedim

    ' Vérifier si le tableau est initialisé
    Dim ancienneTaille As Integer
    ancienneTaille = 0

    On Error Resume Next
    ancienneTaille = UBound(arr)
    If Err.Number <> 0 Then
        ' Tableau non initialisé
        ReDim arr(1 To nouvelleTaille)
        Err.Clear
    Else
        ' Tableau déjà initialisé
        ReDim Preserve arr(1 To nouvelleTaille)
    End If
    On Error GoTo 0

    RedimensionnerSecurise = True
    Exit Function

ErreurRedim:
    Debug.Print "Erreur de redimensionnement : " & Err.Description
    RedimensionnerSecurise = False
End Function

Sub UtiliserRedimensionnementSecurise()
    Dim monTableau() As Variant

    If RedimensionnerSecurise(monTableau, 5) Then
        monTableau(1) = "Premier"
        monTableau(5) = "Cinquième"
        Debug.Print "Redimensionnement réussi : " & UBound(monTableau)
    Else
        Debug.Print "Échec du redimensionnement"
    End If
End Sub
```

---

## Optimisation et performance

### Impact sur les performances

```vba
Sub ComparaisonPerformances()
    Dim tableau() As Integer
    Dim debut As Double
    Dim i As Long

    ' === TEST 1 : ReDim répétitifs (LENT) ===
    debut = Timer
    For i = 1 To 1000
        ReDim Preserve tableau(1 To i)
        tableau(i) = i
    Next i
    Debug.Print "ReDim répétitifs : " & Format(Timer - debut, "0.000") & " secondes"

    ' === TEST 2 : ReDim par blocs (RAPIDE) ===
    Erase tableau  ' Réinitialiser
    debut = Timer

    Dim tailleBloc As Integer
    Dim tailleAllouee As Integer
    tailleBloc = 100
    tailleAllouee = 0

    For i = 1 To 1000
        If i > tailleAllouee Then
            tailleAllouee = tailleAllouee + tailleBloc
            ReDim Preserve tableau(1 To tailleAllouee)
        End If
        tableau(i) = i
    Next i

    ' Ajuster à la taille finale
    ReDim Preserve tableau(1 To 1000)

    Debug.Print "ReDim par blocs : " & Format(Timer - debut, "0.000") & " secondes"

    ' === TEST 3 : ReDim unique (TRÈS RAPIDE) ===
    Erase tableau
    debut = Timer

    ReDim tableau(1 To 1000)
    For i = 1 To 1000
        tableau(i) = i
    Next i

    Debug.Print "ReDim unique : " & Format(Timer - debut, "0.000") & " secondes"
End Sub
```

### Stratégies d'optimisation

#### 1. **Estimation de taille**

```vba
Sub EstimationTaille()
    ' Estimer la taille nécessaire avant de commencer
    Dim nbLignes As Long
    nbLignes = Cells(Rows.Count, 1).End(xlUp).Row

    ' Allouer directement la bonne taille (ou légèrement plus)
    Dim donnees() As Variant
    ReDim donnees(1 To CLng(nbLignes * 1.1))  ' 10% de marge

    ' Remplir sans redimensionnement
    Dim i As Long
    For i = 1 To nbLignes
        donnees(i) = Cells(i, 1).Value
    Next i

    ' Ajuster à la taille exacte à la fin
    ReDim Preserve donnees(1 To nbLignes)

    Debug.Print "Optimisation par estimation : " & UBound(donnees) & " éléments"
End Sub
```

#### 2. **Pooling de tableaux**

```vba
' Variables globales pour le pool
Dim poolTableaux(1 To 10) As Variant  
Dim poolUtilise(1 To 10) As Boolean  

Function ObtenirTableauDuPool(taille As Integer) As Integer
    ' Chercher un tableau libre dans le pool
    Dim i As Integer
    For i = 1 To 10
        If Not poolUtilise(i) Then
            ReDim poolTableaux(i)(1 To taille)
            poolUtilise(i) = True
            ObtenirTableauDuPool = i
            Exit Function
        End If
    Next i

    ' Aucun tableau libre
    ObtenirTableauDuPool = -1
End Function

Sub LibererTableauDuPool(index As Integer)
    If index >= 1 And index <= 10 Then
        poolUtilise(index) = False
        Erase poolTableaux(index)
    End If
End Sub

Sub UtiliserPool()
    Dim indexTableau As Integer
    indexTableau = ObtenirTableauDuPool(100)

    If indexTableau <> -1 Then
        ' Utiliser poolTableaux(indexTableau)
        poolTableaux(indexTableau)(1) = "Données"
        Debug.Print "Tableau " & indexTableau & " utilisé"

        ' Libérer quand terminé
        Call LibererTableauDuPool(indexTableau)
    End If
End Sub
```

---

## Techniques avancées

### 1. Redimensionnement multidimensionnel

```vba
Sub RedimensionnementMultidimensionnel()
    ' Pour les tableaux 2D et plus, seule la dernière dimension peut changer avec Preserve
    Dim matrice() As Double

    ' Initialisation
    ReDim matrice(1 To 3, 1 To 4)
    matrice(1, 1) = 1.1
    matrice(3, 4) = 3.4

    ' ✅ OK : Changer la dernière dimension
    ReDim Preserve matrice(1 To 3, 1 To 6)
    matrice(1, 5) = 1.5
    Debug.Print "Matrice étendue : " & matrice(1, 1) & ", " & matrice(1, 5)

    ' ❌ Impossible : Changer la première dimension avec Preserve
    ' Solution : Créer une nouvelle matrice et copier
    Dim nouvelleMatrice() As Double
    ReDim nouvelleMatrice(1 To 5, 1 To 6)

    ' Copier les données existantes
    Dim i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 6
            nouvelleMatrice(i, j) = matrice(i, j)
        Next j
    Next i

    ' Remplacer l'ancienne matrice
    Erase matrice
    ReDim matrice(1 To 5, 1 To 6)
    For i = 1 To 5
        For j = 1 To 6
            If i <= 3 Then matrice(i, j) = nouvelleMatrice(i, j)
        Next j
    Next i

    Debug.Print "Première dimension étendue manuellement"
End Sub
```

### 2. Classe wrapper pour tableaux dynamiques

```vba
' Dans un module de classe nommé "TableauDynamique"
Private mDonnees() As Variant  
Private mTaille As Long  
Private mCapacite As Long  

Public Sub Initialiser(Optional tailleInitiale As Long = 10)
    mTaille = 0
    mCapacite = tailleInitiale
    ReDim mDonnees(1 To mCapacite)
End Sub

Public Sub Ajouter(valeur As Variant)
    mTaille = mTaille + 1

    ' Agrandir si nécessaire
    If mTaille > mCapacite Then
        mCapacite = mCapacite * 2  ' Doubler la capacité
        ReDim Preserve mDonnees(1 To mCapacite)
    End If

    mDonnees(mTaille) = valeur
End Sub

Public Function Obtenir(index As Long) As Variant
    If index >= 1 And index <= mTaille Then
        Obtenir = mDonnees(index)
    Else
        Err.Raise 9, , "Index hors limites"
    End If
End Function

Public Property Get Taille() As Long
    Taille = mTaille
End Property

' Utilisation de la classe (dans un module standard)
Sub UtiliserClasseTableau()
    Dim monTableau As New TableauDynamique

    monTableau.Initialiser 5
    monTableau.Ajouter "Premier"
    monTableau.Ajouter "Deuxième"
    monTableau.Ajouter "Troisième"

    Debug.Print "Élément 2 : " & monTableau.Obtenir(2)
    Debug.Print "Taille : " & monTableau.Taille
End Sub
```

---

## Conseils et bonnes pratiques

### 1. **Planification de croissance**

```vba
Sub PlanificationCroissance()
    ' Estimer la croissance probable
    Dim donneesUtilisateur() As String
    Dim tailleEstimee As Integer

    ' Demander ou estimer la taille
    tailleEstimee = InputBox("Combien d'éléments prévoyez-vous ?", "Estimation", "10")

    ' Allouer avec une marge
    ReDim donneesUtilisateur(1 To CLng(tailleEstimee * 1.2))  ' 20% de marge

    Debug.Print "Tableau préparé pour " & UBound(donneesUtilisateur) & " éléments"
End Sub
```

### 2. **Logging des redimensionnements**

```vba
Sub LoggingRedimensionnements()
    Dim tableau() As Integer
    Dim compteurRedim As Integer
    compteurRedim = 0

    ' Initialisation
    ReDim tableau(1 To 5)

    ' Redimensionnement avec log
    compteurRedim = compteurRedim + 1
    ReDim Preserve tableau(1 To 10)
    Debug.Print "ReDim #" & compteurRedim & " : nouvelle taille = 10"

    compteurRedim = compteurRedim + 1
    ReDim Preserve tableau(1 To 15)
    Debug.Print "ReDim #" & compteurRedim & " : nouvelle taille = 15"

    compteurRedim = compteurRedim + 1
    ReDim Preserve tableau(1 To 20)
    Debug.Print "ReDim #" & compteurRedim & " : nouvelle taille = 20"

    Debug.Print "Total de redimensionnements : " & compteurRedim
End Sub
```

### 3. **Nettoyage de mémoire**

```vba
Sub NettoyageMemoire()
    Dim grosTableau() As Double

    ' Utilisation intensive
    ReDim grosTableau(1 To 100000)
    ' ... traitement ...

    ' Libérer explicitement la mémoire quand terminé
    Erase grosTableau

    Debug.Print "Mémoire libérée"
End Sub
```

---

## Récapitulatif

### Points clés à retenir

1. **ReDim normal** : Recrée le tableau, perd les données
2. **ReDim Preserve** : Garde les données, limitations sur les dimensions
3. **Performance** : Éviter les ReDim répétitifs, préférer par blocs
4. **Erreurs** : Toujours vérifier l'initialisation avant utilisation
5. **Optimisation** : Estimer la taille, utiliser des marges, nettoyer la mémoire

### Modèles de code recommandés

#### Ajout d'éléments efficace
```vba
Sub AjoutEfficace()
    Dim arr() As Variant
    Dim taille As Integer, capacite As Integer
    Dim bloc As Integer: bloc = 10

    taille = 0: capacite = 0

    ' Pour chaque nouvel élément
    taille = taille + 1
    If taille > capacite Then
        capacite = capacite + bloc
        ReDim Preserve arr(1 To capacite)
    End If
    arr(taille) = "Nouvelle valeur"
End Sub
```

#### Vérification sécurisée
```vba
Function TableauInitialise(arr As Variant) As Boolean
    On Error Resume Next
    Dim test As Integer: test = UBound(arr)
    TableauInitialise = (Err.Number = 0)
    On Error GoTo 0
End Function
```

### Erreurs à éviter

- ❌ ReDim répétitifs sans stratégie de croissance
- ❌ Oublier de vérifier l'initialisation
- ❌ Modifier des dimensions interdites avec Preserve
- ❌ Ne pas libérer la mémoire des gros tableaux
- ❌ Redimensionner sans estimer les besoins

**ReDim** est un outil puissant mais qui doit être utilisé intelligemment. Dans la section suivante, nous explorerons les tableaux multidimensionnels qui ajoutent encore plus de possibilités à vos programmes VBA.

⏭️ [Tableaux multidimensionnels](/08-tableaux-arrays/04-tableaux-multidimensionnels.md)
