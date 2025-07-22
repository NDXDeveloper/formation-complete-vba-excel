🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.5. Parcours et manipulation des tableaux

## Introduction au parcours et à la manipulation

Le **parcours** d'un tableau consiste à visiter chaque élément de manière systématique, tandis que la **manipulation** englobe toutes les opérations que vous pouvez effectuer sur ces éléments : les modifier, les rechercher, les trier, les filtrer, ou les transformer. C'est comme **explorer méthodiquement une bibliothèque** pour inventorier, réorganiser ou retrouver des livres.

**Analogie simple :**
- **Parcours** = Visiter chaque rayon de la bibliothèque dans l'ordre
- **Manipulation** = Ranger les livres par ordre alphabétique, chercher un titre spécifique, ou déplacer des ouvrages
- **Algorithmes** = Les méthodes efficaces pour accomplir ces tâches

Maîtriser ces techniques vous permettra de traiter efficacement n'importe quel volume de données.

---

## Techniques de parcours fondamentales

### 1. Parcours séquentiel simple (For...Next)

Le parcours avec `For...Next` est la méthode la plus directe et la plus contrôlée pour visiter chaque élément d'un tableau.

```vba
Sub ParcoursSequentiel()
    Dim nombres(1 To 10) As Integer

    ' Remplissage initial
    Dim i As Integer
    For i = 1 To 10
        nombres(i) = i * i  ' Carrés parfaits : 1, 4, 9, 16, ...
    Next i

    ' Parcours et affichage
    Debug.Print "=== Parcours séquentiel ==="
    For i = LBound(nombres) To UBound(nombres)
        Debug.Print "nombres(" & i & ") = " & nombres(i)
    Next i

    ' Parcours avec transformation
    Debug.Print "=== Après transformation (x2) ==="
    For i = LBound(nombres) To UBound(nombres)
        nombres(i) = nombres(i) * 2
        Debug.Print "nombres(" & i & ") = " & nombres(i)
    Next i
End Sub
```

### 2. Parcours avec For Each (plus simple)

`For Each` simplifie le parcours en éliminant la gestion manuelle des indices.

```vba
Sub ParcoursForEach()
    Dim couleurs(1 To 5) As String
    couleurs(1) = "Rouge": couleurs(2) = "Vert": couleurs(3) = "Bleu"
    couleurs(4) = "Jaune": couleurs(5) = "Orange"

    ' Parcours simple avec For Each
    Debug.Print "=== Parcours avec For Each ==="
    Dim couleur As Variant
    For Each couleur In couleurs
        Debug.Print "Couleur : " & couleur
    Next couleur

    ' Limitation : For Each ne permet pas de modifier les éléments
    ' Pour modifier, il faut utiliser For...Next avec indices
End Sub
```

### 3. Parcours conditionnel

Parcourir en appliquant des conditions pour traiter seulement certains éléments.

```vba
Sub ParcoursConditionnel()
    Dim notes(1 To 15) As Integer

    ' Remplissage avec notes aléatoires
    Dim i As Integer
    For i = 1 To 15
        notes(i) = Int(Rnd() * 20) + 1  ' Notes de 1 à 20
    Next i

    ' Parcours conditionnel : traiter seulement les bonnes notes
    Debug.Print "=== Notes excellentes (≥ 16) ==="
    For i = LBound(notes) To UBound(notes)
        If notes(i) >= 16 Then
            Debug.Print "Élève " & i & " : " & notes(i) & "/20 - EXCELLENT"
        End If
    Next i

    ' Compter les éléments répondant à un critère
    Dim compteurReussite As Integer
    compteurReussite = 0
    For i = LBound(notes) To UBound(notes)
        If notes(i) >= 10 Then
            compteurReussite = compteurReussite + 1
        End If
    Next i

    Debug.Print "Nombre d'élèves en réussite : " & compteurReussite & "/" & UBound(notes)
End Sub
```

### 4. Parcours à rebours

Parcourir un tableau de la fin vers le début, utile pour les suppressions ou traitements spéciaux.

```vba
Sub ParcoursRebours()
    Dim liste(1 To 8) As String
    liste(1) = "Janvier": liste(2) = "Février": liste(3) = "Mars": liste(4) = "Avril"
    liste(5) = "Mai": liste(6) = "Juin": liste(7) = "Juillet": liste(8) = "Août"

    ' Parcours normal (début vers fin)
    Debug.Print "=== Parcours normal ==="
    Dim i As Integer
    For i = LBound(liste) To UBound(liste)
        Debug.Print i & ": " & liste(i)
    Next i

    ' Parcours à rebours (fin vers début)
    Debug.Print "=== Parcours à rebours ==="
    For i = UBound(liste) To LBound(liste) Step -1
        Debug.Print i & ": " & liste(i)
    Next i

    ' Utilité : supprimer des éléments sans affecter les indices suivants
    ' (Nous verrons cela dans les algorithmes de suppression)
End Sub
```

---

## Algorithmes de recherche

### 1. Recherche linéaire (séquentielle)

La recherche la plus simple : examiner chaque élément jusqu'à trouver celui recherché.

```vba
Function RechercheLineaire(tableau() As Variant, valeurCherchee As Variant) As Integer
    ' Retourne l'indice de la première occurrence, ou -1 si non trouvé

    Dim i As Integer
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i) = valeurCherchee Then
            RechercheLineaire = i
            Exit Function
        End If
    Next i

    RechercheLineaire = -1  ' Non trouvé
End Function

Sub UtiliserRechercheLineaire()
    Dim fruits(1 To 6) As String
    fruits(1) = "Pomme": fruits(2) = "Banane": fruits(3) = "Orange"
    fruits(4) = "Kiwi": fruits(5) = "Mangue": fruits(6) = "Ananas"

    ' Rechercher "Orange"
    Dim position As Integer
    position = RechercheLineaire(fruits, "Orange")

    If position <> -1 Then
        Debug.Print "Orange trouvée à la position " & position
    Else
        Debug.Print "Orange non trouvée"
    End If

    ' Rechercher un fruit inexistant
    position = RechercheLineaire(fruits, "Cerise")
    If position = -1 Then
        Debug.Print "Cerise non trouvée dans la liste"
    End If
End Sub
```

### 2. Recherche avec critères multiples

Rechercher selon plusieurs conditions simultanées.

```vba
Function RechercherParCriteres(ages() As Integer, salaires() As Double, _
                              ageMin As Integer, salaireMin As Double) As Integer
    ' Chercher le premier employé qui satisfait les deux critères

    If UBound(ages) <> UBound(salaires) Then
        Debug.Print "Erreur : les tableaux n'ont pas la même taille"
        RechercherParCriteres = -1
        Exit Function
    End If

    Dim i As Integer
    For i = LBound(ages) To UBound(ages)
        If ages(i) >= ageMin And salaires(i) >= salaireMin Then
            RechercherParCriteres = i
            Exit Function
        End If
    Next i

    RechercherParCriteres = -1
End Function

Sub UtiliserRechercheCriteres()
    Dim ages(1 To 5) As Integer
    Dim salaires(1 To 5) As Double

    ' Données d'employés
    ages(1) = 25: salaires(1) = 2800
    ages(2) = 35: salaires(2) = 3500
    ages(3) = 28: salaires(3) = 3200
    ages(4) = 45: salaires(4) = 4500
    ages(5) = 30: salaires(5) = 3000

    ' Chercher un employé de 30 ans ou plus avec un salaire d'au moins 3500€
    Dim position As Integer
    position = RechercherParCriteres(ages, salaires, 30, 3500)

    If position <> -1 Then
        Debug.Print "Employé trouvé à la position " & position & _
                    " : " & ages(position) & " ans, " & salaires(position) & "€"
    Else
        Debug.Print "Aucun employé ne correspond aux critères"
    End If
End Sub
```

### 3. Recherche de toutes les occurrences

Trouver toutes les positions où apparaît une valeur.

```vba
Function RechercherToutesOccurrences(tableau() As Variant, valeur As Variant) As Variant
    ' Retourne un tableau avec tous les indices où la valeur apparaît

    Dim resultats() As Integer
    Dim compteur As Integer
    compteur = 0

    ' Premier passage : compter les occurrences
    Dim i As Integer
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i) = valeur Then
            compteur = compteur + 1
        End If
    Next i

    If compteur = 0 Then
        RechercherToutesOccurrences = Array()  ' Tableau vide
        Exit Function
    End If

    ' Redimensionner le tableau de résultats
    ReDim resultats(1 To compteur)

    ' Deuxième passage : collecter les indices
    Dim indexResultat As Integer
    indexResultat = 0
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i) = valeur Then
            indexResultat = indexResultat + 1
            resultats(indexResultat) = i
        End If
    Next i

    RechercherToutesOccurrences = resultats
End Function

Sub UtiliserRechercheComplete()
    Dim lettres(1 To 10) As String
    lettres(1) = "A": lettres(2) = "B": lettres(3) = "A"
    lettres(4) = "C": lettres(5) = "A": lettres(6) = "D"
    lettres(7) = "B": lettres(8) = "A": lettres(9) = "E": lettres(10) = "A"

    ' Chercher toutes les occurrences de "A"
    Dim positions As Variant
    positions = RechercherToutesOccurrences(lettres, "A")

    If IsArray(positions) And UBound(positions) >= LBound(positions) Then
        Debug.Print "La lettre A trouvée aux positions :"
        Dim j As Integer
        For j = LBound(positions) To UBound(positions)
            Debug.Print "  Position " & positions(j)
        Next j
    Else
        Debug.Print "Aucune occurrence trouvée"
    End If
End Sub
```

---

## Algorithmes de tri

### 1. Tri à bulles (Bubble Sort) - Simple à comprendre

L'algorithme de tri le plus simple : comparer des éléments adjacents et les échanger si nécessaire.

```vba
Sub TriBulles(tableau() As Variant)
    ' Tri croissant par bulles
    Dim n As Integer
    n = UBound(tableau) - LBound(tableau) + 1

    Dim i As Integer, j As Integer
    Dim temp As Variant

    For i = 1 To n - 1
        For j = LBound(tableau) To UBound(tableau) - i
            ' Si l'élément actuel est plus grand que le suivant, échanger
            If tableau(j) > tableau(j + 1) Then
                temp = tableau(j)
                tableau(j) = tableau(j + 1)
                tableau(j + 1) = temp
            End If
        Next j
    Next i
End Sub

Sub UtiliserTriBulles()
    Dim nombres(1 To 8) As Integer
    nombres(1) = 64: nombres(2) = 34: nombres(3) = 25: nombres(4) = 12
    nombres(5) = 22: nombres(6) = 11: nombres(7) = 90: nombres(8) = 5

    Debug.Print "=== Avant tri ==="
    Dim i As Integer
    For i = LBound(nombres) To UBound(nombres)
        Debug.Print nombres(i);  ' Point-virgule pour affichage horizontal
    Next i
    Debug.Print ""  ' Retour à la ligne

    ' Trier
    Call TriBulles(nombres)

    Debug.Print "=== Après tri ==="
    For i = LBound(nombres) To UBound(nombres)
        Debug.Print nombres(i);
    Next i
    Debug.Print ""
End Sub
```

### 2. Tri par sélection (Selection Sort) - Plus efficace

Trouve répétitivement le plus petit élément et le place à sa position finale.

```vba
Sub TriSelection(tableau() As Variant)
    ' Tri croissant par sélection
    Dim i As Integer, j As Integer
    Dim indexMin As Integer
    Dim temp As Variant

    For i = LBound(tableau) To UBound(tableau) - 1
        ' Trouver l'indice du minimum dans la partie non triée
        indexMin = i
        For j = i + 1 To UBound(tableau)
            If tableau(j) < tableau(indexMin) Then
                indexMin = j
            End If
        Next j

        ' Échanger l'élément minimum avec l'élément à la position i
        If indexMin <> i Then
            temp = tableau(i)
            tableau(i) = tableau(indexMin)
            tableau(indexMin) = temp
        End If
    Next i
End Sub

Sub UtiliserTriSelection()
    Dim mots(1 To 6) As String
    mots(1) = "Zèbre": mots(2) = "Abeille": mots(3) = "Éléphant"
    mots(4) = "Chat": mots(5) = "Baleine": mots(6) = "Dauphin"

    Debug.Print "=== Mots avant tri ==="
    Dim i As Integer
    For i = LBound(mots) To UBound(mots)
        Debug.Print mots(i)
    Next i

    Call TriSelection(mots)

    Debug.Print "=== Mots après tri ==="
    For i = LBound(mots) To UBound(mots)
        Debug.Print mots(i)
    Next i
End Sub
```

### 3. Tri de tableaux avec indices associés

Trier un tableau tout en gardant la correspondance avec un autre tableau.

```vba
Sub TriAvecIndicesAssocies()
    ' Trier les salaires en gardant la correspondance avec les noms
    Dim noms(1 To 5) As String
    Dim salaires(1 To 5) As Double

    ' Données initiales
    noms(1) = "Alice": salaires(1) = 3200
    noms(2) = "Bob": salaires(2) = 2800
    noms(3) = "Claire": salaires(3) = 3800
    noms(4) = "David": salaires(4) = 2500
    noms(5) = "Eve": salaires(5) = 4200

    Debug.Print "=== Avant tri par salaire ==="
    Dim i As Integer, j As Integer
    For i = 1 To 5
        Debug.Print noms(i) & " : " & salaires(i) & "€"
    Next i

    ' Tri par sélection avec échange simultané des deux tableaux
    Dim indexMin As Integer
    Dim tempSalaire As Double, tempNom As String

    For i = LBound(salaires) To UBound(salaires) - 1
        indexMin = i
        For j = i + 1 To UBound(salaires)
            If salaires(j) > salaires(indexMin) Then  ' Tri décroissant
                indexMin = j
            End If
        Next j

        If indexMin <> i Then
            ' Échanger les salaires
            tempSalaire = salaires(i)
            salaires(i) = salaires(indexMin)
            salaires(indexMin) = tempSalaire

            ' Échanger les noms correspondants
            tempNom = noms(i)
            noms(i) = noms(indexMin)
            noms(indexMin) = tempNom
        End If
    Next i

    Debug.Print "=== Après tri par salaire (décroissant) ==="
    For i = 1 To 5
        Debug.Print noms(i) & " : " & salaires(i) & "€"
    Next i
End Sub
```

---

## Filtrage et transformation

### 1. Filtrage par critères

Extraire les éléments qui répondent à certains critères.

```vba
Function FiltrerTableau(tableau() As Variant, critere As String, valeurSeuil As Variant) As Variant
    ' Filtrer selon différents critères : ">", "<", "=", "<>", ">=", "<="

    Dim resultatsTemp() As Variant
    Dim compteur As Integer
    compteur = 0

    ' Premier passage : compter les éléments qui correspondent
    Dim i As Integer
    For i = LBound(tableau) To UBound(tableau)
        Dim correspond As Boolean
        correspond = False

        Select Case critere
            Case ">"
                correspond = (tableau(i) > valeurSeuil)
            Case "<"
                correspond = (tableau(i) < valeurSeuil)
            Case "="
                correspond = (tableau(i) = valeurSeuil)
            Case "<>"
                correspond = (tableau(i) <> valeurSeuil)
            Case ">="
                correspond = (tableau(i) >= valeurSeuil)
            Case "<="
                correspond = (tableau(i) <= valeurSeuil)
        End Select

        If correspond Then compteur = compteur + 1
    Next i

    If compteur = 0 Then
        FiltrerTableau = Array()  ' Tableau vide
        Exit Function
    End If

    ' Deuxième passage : collecter les éléments
    ReDim resultatsTemp(1 To compteur)
    Dim indexResultat As Integer
    indexResultat = 0

    For i = LBound(tableau) To UBound(tableau)
        correspond = False

        Select Case critere
            Case ">"
                correspond = (tableau(i) > valeurSeuil)
            Case "<"
                correspond = (tableau(i) < valeurSeuil)
            Case "="
                correspond = (tableau(i) = valeurSeuil)
            Case "<>"
                correspond = (tableau(i) <> valeurSeuil)
            Case ">="
                correspond = (tableau(i) >= valeurSeuil)
            Case "<="
                correspond = (tableau(i) <= valeurSeuil)
        End Select

        If correspond Then
            indexResultat = indexResultat + 1
            resultatsTemp(indexResultat) = tableau(i)
        End If
    Next i

    FiltrerTableau = resultatsTemp
End Function

Sub UtiliserFiltrage()
    Dim temperatures(1 To 10) As Double
    temperatures(1) = 18.5: temperatures(2) = 22.3: temperatures(3) = 15.8
    temperatures(4) = 25.1: temperatures(5) = 19.7: temperatures(6) = 28.4
    temperatures(7) = 16.2: temperatures(8) = 24.8: temperatures(9) = 21.5: temperatures(10) = 30.2

    ' Filtrer les températures supérieures à 20°
    Dim tempElevees As Variant
    tempElevees = FiltrerTableau(temperatures, ">", 20)

    If IsArray(tempElevees) Then
        Debug.Print "Températures > 20° :"
        Dim i As Integer
        For i = LBound(tempElevees) To UBound(tempElevees)
            Debug.Print "  " & tempElevees(i) & "°C"
        Next i
    End If

    ' Filtrer les températures exactement égales à une valeur
    Dim tempPrecises As Variant
    tempPrecises = FiltrerTableau(temperatures, "=", 22.3)

    If IsArray(tempPrecises) And UBound(tempPrecises) >= LBound(tempPrecises) Then
        Debug.Print "Températures exactement à 22.3° : " & UBound(tempPrecises) - LBound(tempPrecises) + 1
    End If
End Sub
```

### 2. Transformation (mapping)

Appliquer une fonction à chaque élément d'un tableau.

```vba
Function TransformerTableau(tableau() As Variant, operation As String, parametre As Variant) As Variant
    ' Appliquer une transformation à chaque élément

    Dim resultat() As Variant
    ReDim resultat(LBound(tableau) To UBound(tableau))

    Dim i As Integer
    For i = LBound(tableau) To UBound(tableau)
        Select Case LCase(operation)
            Case "multiplier"
                resultat(i) = tableau(i) * parametre
            Case "ajouter"
                resultat(i) = tableau(i) + parametre
            Case "puissance"
                resultat(i) = tableau(i) ^ parametre
            Case "concatener"
                resultat(i) = tableau(i) & parametre
            Case "majuscule"
                resultat(i) = UCase(tableau(i))
            Case "minuscule"
                resultat(i) = LCase(tableau(i))
            Case Else
                resultat(i) = tableau(i)  ' Pas de transformation
        End Select
    Next i

    TransformerTableau = resultat
End Function

Sub UtiliserTransformation()
    ' Transformation numérique
    Dim nombres(1 To 5) As Double
    nombres(1) = 10: nombres(2) = 20: nombres(3) = 30: nombres(4) = 40: nombres(5) = 50

    Dim doubles As Variant
    doubles = TransformerTableau(nombres, "multiplier", 2)

    Debug.Print "=== Transformation numérique (x2) ==="
    Dim i As Integer
    For i = LBound(doubles) To UBound(doubles)
        Debug.Print nombres(i) & " → " & doubles(i)
    Next i

    ' Transformation textuelle
    Dim prenoms(1 To 4) As String
    prenoms(1) = "alice": prenoms(2) = "bob": prenoms(3) = "claire": prenoms(4) = "david"

    Dim prenomsFormats As Variant
    prenomsFormats = TransformerTableau(prenoms, "majuscule", "")

    Debug.Print "=== Transformation textuelle (majuscules) ==="
    For i = LBound(prenomsFormats) To UBound(prenomsFormats)
        Debug.Print prenoms(i) & " → " & prenomsFormats(i)
    Next i
End Sub
```

### 3. Agrégation et statistiques

Calculer des valeurs statistiques sur les tableaux.

```vba
Function CalculerStatistiques(tableau() As Variant) As Variant
    ' Retourne un tableau avec : minimum, maximum, somme, moyenne, médiane

    If UBound(tableau) < LBound(tableau) Then
        CalculerStatistiques = Array("Erreur : tableau vide")
        Exit Function
    End If

    Dim min As Variant, max As Variant, somme As Double
    Dim i As Integer

    ' Initialiser avec le premier élément
    min = tableau(LBound(tableau))
    max = tableau(LBound(tableau))
    somme = 0

    ' Parcourir pour trouver min, max et calculer la somme
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i) < min Then min = tableau(i)
        If tableau(i) > max Then max = tableau(i)
        somme = somme + tableau(i)
    Next i

    Dim nbElements As Integer
    nbElements = UBound(tableau) - LBound(tableau) + 1
    Dim moyenne As Double
    moyenne = somme / nbElements

    ' Calculer la médiane (nécessite un tri temporaire)
    Dim tableauTrie() As Variant
    ReDim tableauTrie(LBound(tableau) To UBound(tableau))
    For i = LBound(tableau) To UBound(tableau)
        tableauTrie(i) = tableau(i)
    Next i

    Call TriSelection(tableauTrie)  ' Trier pour la médiane

    Dim mediane As Double
    If nbElements Mod 2 = 1 Then
        ' Nombre impair d'éléments
        mediane = tableauTrie(LBound(tableauTrie) + nbElements \ 2)
    Else
        ' Nombre pair d'éléments
        Dim milieu1 As Integer, milieu2 As Integer
        milieu1 = LBound(tableauTrie) + nbElements \ 2 - 1
        milieu2 = milieu1 + 1
        mediane = (tableauTrie(milieu1) + tableauTrie(milieu2)) / 2
    End If

    CalculerStatistiques = Array(min, max, somme, moyenne, mediane)
End Function

Sub UtiliserStatistiques()
    Dim notes(1 To 9) As Double
    notes(1) = 12: notes(2) = 15: notes(3) = 8: notes(4) = 18: notes(5) = 14
    notes(6) = 11: notes(7) = 16: notes(8) = 9: notes(9) = 13

    Dim stats As Variant
    stats = CalculerStatistiques(notes)

    Debug.Print "=== Statistiques des notes ==="
    Debug.Print "Minimum : " & stats(0)
    Debug.Print "Maximum : " & stats(1)
    Debug.Print "Somme : " & stats(2)
    Debug.Print "Moyenne : " & Format(stats(3), "0.00")
    Debug.Print "Médiane : " & stats(4)

    ' Afficher toutes les notes pour vérification
    Debug.Print "Notes : ";
    Dim i As Integer
    For i = LBound(notes) To UBound(notes)
        Debug.Print notes(i);
        If i < UBound(notes) Then Debug.Print ", ";
    Next i
    Debug.Print ""
End Sub
```

---

## Manipulation avancée

### 1. Fusion de tableaux

Combiner plusieurs tableaux en un seul.

```vba
Function FusionnerTableaux(tableau1() As Variant, tableau2() As Variant) As Variant
    ' Fusionner deux tableaux en un seul

    Dim taille1 As Integer, taille2 As Integer
    taille1 = UBound(tableau1) - LBound(tableau1) + 1
    taille2 = UBound(tableau2) - LBound(tableau2) + 1

    Dim resultat() As Variant
    ReDim resultat(1 To taille1 + taille2)

    ' Copier le premier tableau
    Dim i As Integer, indexResultat As Integer
    indexResultat = 1
    For i = LBound(tableau1) To UBound(tableau1)
        resultat(indexResultat) = tableau1(i)
        indexResultat = indexResultat + 1
    Next i

    ' Copier le deuxième tableau
    For i = LBound(tableau2) To UBound(tableau2)
        resultat(indexResultat) = tableau2(i)
        indexResultat = indexResultat + 1
    Next i

    FusionnerTableaux = resultat
End Function

Sub UtiliserFusion()
    Dim fruits(1 To 3) As String
    Dim legumes(1 To 4) As String

    fruits(1) = "Pomme": fruits(2) = "Banane": fruits(3) = "Orange"
    legumes(1) = "Carotte": legumes(2) = "Brocoli": legumes(3) = "Épinard": legumes(4) = "Tomate"

    Dim alimentation As Variant
    alimentation = FusionnerTableaux(fruits, legumes)

    Debug.Print "=== Tableau fusionné ==="
    Dim i As Integer
    For i = LBound(alimentation) To UBound(alimentation)
        Debug.Print i & ": " & alimentation(i)
    Next i
End Sub
```

### 2. Suppression d'éléments

Supprimer des éléments selon différents critères.

```vba
Function SupprimerElements(tableau() As Variant, valeurASupprimer As Variant) As Variant
    ' Supprimer toutes les occurrences d'une valeur

    Dim compteurNouveaux As Integer
    compteurNouveaux = 0

    ' Compter les éléments à garder
    Dim i As Integer
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i) <> valeurASupprimer Then
            compteurNouveaux = compteurNouveaux + 1
        End If
    Next i

    If compteurNouveaux = 0 Then
        SupprimerElements = Array()  ' Tableau vide
        Exit Function
    End If

    ' Créer le nouveau tableau
    Dim resultat() As Variant
    ReDim resultat(1 To compteurNouveaux)

    Dim indexResultat As Integer
    indexResultat = 1
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i) <> valeurASupprimer Then
            resultat(indexResultat) = tableau(i)
            indexResultat = indexResultat + 1
        End If
    Next i

    SupprimerElements = resultat
End Function

Function SupprimerDoublons(tableau() As Variant) As Variant
    ' Supprimer les doublons en gardant la première occurrence

    Dim resultatsTemp() As Variant
    Dim compteur As Integer
    compteur = 0

    Dim i As Integer, j As Integer
    For i = LBound(tableau) To UBound(tableau)
        Dim estDouble As Boolean
        estDouble = False

        ' Vérifier si cet élément existe déjà dans les résultats
        For j = 1 To compteur
            If resultatsTemp(j) = tableau(i) Then
                estDouble = True
                Exit For
            End If
        Next j

        If Not estDouble Then
            compteur = compteur + 1
            ReDim Preserve resultatsTemp(1 To compteur)
            resultatsTemp(compteur) = tableau(i)
        End If
    Next i

    SupprimerDoublons = resultatsTemp
End Function

Sub UtiliserSuppression()
    ' Test suppression d'éléments
    Dim lettres(1 To 8) As String
    lettres(1) = "A": lettres(2) = "B": lettres(3) = "A"
    lettres(4) = "C": lettres(5) = "B": lettres(6) = "D"
    lettres(7) = "A": lettres(8) = "E"

    Debug.Print "=== Tableau original ==="
    Dim i As Integer
    For i = LBound(lettres) To UBound(lettres)
        Debug.Print lettres(i);
    Next i
    Debug.Print ""

    ' Supprimer tous les "A"
    Dim sansA As Variant
    sansA = SupprimerElements(lettres, "A")

    Debug.Print "=== Après suppression des A ==="
    If IsArray(sansA) Then
        For i = LBound(sansA) To UBound(sansA)
            Debug.Print sansA(i);
        Next i
        Debug.Print ""
    End If

    ' Supprimer les doublons
    Dim sansDoublons As Variant
    sansDoublons = SupprimerDoublons(lettres)

    Debug.Print "=== Après suppression des doublons ==="
    For i = LBound(sansDoublons) To UBound(sansDoublons)
        Debug.Print sansDoublons(i);
    Next i
    Debug.Print ""
End Sub
```

### 3. Inversion et rotation

Manipuler l'ordre des éléments dans un tableau.

```vba
Sub InverserTableau(tableau() As Variant)
    ' Inverser l'ordre des éléments dans le tableau

    Dim debut As Integer, fin As Integer
    Dim temp As Variant

    debut = LBound(tableau)
    fin = UBound(tableau)

    Do While debut < fin
        ' Échanger les éléments aux extrémités
        temp = tableau(debut)
        tableau(debut) = tableau(fin)
        tableau(fin) = temp

        ' Avancer vers le centre
        debut = debut + 1
        fin = fin - 1
    Loop
End Sub

Function RoterTableau(tableau() As Variant, positions As Integer) As Variant
    ' Faire une rotation des éléments vers la droite

    Dim taille As Integer
    taille = UBound(tableau) - LBound(tableau) + 1

    ' Normaliser le nombre de positions (éviter les rotations inutiles)
    positions = positions Mod taille
    If positions < 0 Then positions = positions + taille

    Dim resultat() As Variant
    ReDim resultat(LBound(tableau) To UBound(tableau))

    Dim i As Integer
    For i = LBound(tableau) To UBound(tableau)
        Dim nouvelIndex As Integer
        nouvelIndex = LBound(tableau) + (i - LBound(tableau) + positions) Mod taille
        resultat(nouvelIndex) = tableau(i)
    Next i

    RoterTableau = resultat
End Function

Sub UtiliserInversionRotation()
    Dim nombres(1 To 6) As Integer
    nombres(1) = 10: nombres(2) = 20: nombres(3) = 30
    nombres(4) = 40: nombres(5) = 50: nombres(6) = 60

    Debug.Print "=== Tableau original ==="
    Dim i As Integer
    For i = LBound(nombres) To UBound(nombres)
        Debug.Print nombres(i);
    Next i
    Debug.Print ""

    ' Inversion
    Call InverserTableau(nombres)
    Debug.Print "=== Après inversion ==="
    For i = LBound(nombres) To UBound(nombres)
        Debug.Print nombres(i);
    Next i
    Debug.Print ""

    ' Remettre dans l'ordre original
    Call InverserTableau(nombres)

    ' Rotation
    Dim tournes As Variant
    tournes = RoterTableau(nombres, 2)  ' Rotation de 2 positions vers la droite

    Debug.Print "=== Après rotation de 2 positions ==="
    For i = LBound(tournes) To UBound(tournes)
        Debug.Print tournes(i);
    Next i
    Debug.Print ""
End Sub
```

---

## Optimisation des performances

### 1. Comparaison d'algorithmes

```vba
Sub ComparaisonPerformances()
    ' Comparer les performances de différents algorithmes

    ' Créer un gros tableau pour les tests
    Dim tailleTest As Long
    tailleTest = 1000

    Dim tableauTest() As Integer
    ReDim tableauTest(1 To tailleTest)

    ' Remplir avec des valeurs aléatoires
    Dim i As Long
    For i = 1 To tailleTest
        tableauTest(i) = Int(Rnd() * 1000)
    Next i

    ' Test 1 : Tri à bulles
    Dim tableauBulles() As Integer
    ReDim tableauBulles(1 To tailleTest)
    For i = 1 To tailleTest
        tableauBulles(i) = tableauTest(i)
    Next i

    Dim debut As Double
    debut = Timer
    Call TriBulles(tableauBulles)
    Debug.Print "Tri à bulles (" & tailleTest & " éléments) : " & _
                Format(Timer - debut, "0.000") & " secondes"

    ' Test 2 : Tri par sélection
    Dim tableauSelection() As Integer
    ReDim tableauSelection(1 To tailleTest)
    For i = 1 To tailleTest
        tableauSelection(i) = tableauTest(i)
    Next i

    debut = Timer
    Call TriSelection(tableauSelection)
    Debug.Print "Tri par sélection (" & tailleTest & " éléments) : " & _
                Format(Timer - debut, "0.000") & " secondes"

    ' Test 3 : Recherche linéaire vs recherche dans tableau trié
    Dim valeurCherchee As Integer
    valeurCherchee = tableauTest(tailleTest \ 2)  ' Prendre une valeur du milieu

    debut = Timer
    Dim position As Integer
    position = RechercheLineaire(tableauTest, valeurCherchee)
    Debug.Print "Recherche linéaire : " & Format(Timer - debut, "0.000") & " secondes"

    ' Dans un tableau trié, on pourrait utiliser une recherche binaire (plus avancé)
    Debug.Print "Position trouvée : " & position
End Sub
```

### 2. Optimisations spécifiques

```vba
Sub OptimisationsSpecifiques()
    ' Techniques pour améliorer les performances

    ' 1. Éviter les redimensionnements répétitifs
    Debug.Print "=== Test : Agrandissement de tableau ==="

    ' LENT : Redimensionnement à chaque ajout
    Dim tableauLent() As Integer
    Dim debut As Double
    debut = Timer

    Dim i As Integer
    For i = 1 To 1000
        ReDim Preserve tableauLent(1 To i)
        tableauLent(i) = i
    Next i

    Debug.Print "Méthode lente (ReDim à chaque fois) : " & _
                Format(Timer - debut, "0.000") & " secondes"

    ' RAPIDE : Redimensionnement par blocs
    Dim tableauRapide() As Integer
    Dim tailleActuelle As Integer, tailleAllouee As Integer
    tailleActuelle = 0: tailleAllouee = 0

    debut = Timer
    For i = 1 To 1000
        tailleActuelle = tailleActuelle + 1
        If tailleActuelle > tailleAllouee Then
            tailleAllouee = tailleAllouee + 100  ' Agrandir par blocs de 100
            ReDim Preserve tableauRapide(1 To tailleAllouee)
        End If
        tableauRapide(tailleActuelle) = i
    Next i

    ' Ajuster à la taille finale
    ReDim Preserve tableauRapide(1 To tailleActuelle)

    Debug.Print "Méthode rapide (ReDim par blocs) : " & _
                Format(Timer - debut, "0.000") & " secondes"

    ' 2. Utiliser les bons types de données
    Debug.Print "=== Test : Types de données ==="

    ' Variant vs type spécifique
    Dim tableauVariant() As Variant
    Dim tableauInteger() As Integer
    ReDim tableauVariant(1 To 10000)
    ReDim tableauInteger(1 To 10000)

    ' Test avec Variant
    debut = Timer
    For i = 1 To 10000
        tableauVariant(i) = i * 2
    Next i
    Debug.Print "Traitement avec Variant : " & Format(Timer - debut, "0.000") & " secondes"

    ' Test avec Integer
    debut = Timer
    For i = 1 To 10000
        tableauInteger(i) = i * 2
    Next i
    Debug.Print "Traitement avec Integer : " & Format(Timer - debut, "0.000") & " secondes"
End Sub
```

### 3. Traitement de gros volumes Excel

```vba
Sub TraiterGrosVolumeExcel()
    ' Technique optimisée pour traiter de gros volumes de données Excel

    Debug.Print "=== Traitement optimisé de données Excel ==="

    ' Supposons une plage A1:D10000 (40 000 cellules)
    Dim plageSource As String
    plageSource = "A1:D10000"

    ' LENT : Accès cellule par cellule
    Dim debut As Double
    debut = Timer

    ' Simulation d'accès lent (on ne fait que quelques lignes pour l'exemple)
    Dim i As Long, j As Integer
    For i = 1 To 100  ' Seulement 100 lignes pour l'exemple
        For j = 1 To 4
            Dim valeur As Variant
            valeur = Cells(i, j).Value
            ' Traitement...
        Next j
    Next i

    Debug.Print "Méthode lente (cellule par cellule, 100 lignes) : " & _
                Format(Timer - debut, "0.000") & " secondes"

    ' RAPIDE : Chargement en bloc dans un tableau
    debut = Timer

    ' Charger toute la plage d'un coup
    Dim donneesTableau As Variant
    donneesTableau = Range("A1:D100").Value  ' 100 lignes pour l'exemple

    ' Traitement en mémoire
    For i = 1 To 100
        For j = 1 To 4
            valeur = donneesTableau(i, j)
            ' Traitement... (beaucoup plus rapide)
            donneesTableau(i, j) = valeur * 1.1  ' Exemple : augmentation de 10%
        Next j
    Next i

    ' Réécrire d'un coup
    Range("A1:D100").Value = donneesTableau

    Debug.Print "Méthode rapide (tableau en mémoire, 100 lignes) : " & _
                Format(Timer - debut, "0.000") & " secondes"

    Debug.Print "La différence est encore plus flagrante avec de gros volumes !"
End Sub
```

---

## Patterns et techniques avancées

### 1. Tableau comme cache

```vba
' Variables globales pour le cache
Dim cacheCalculs() As Double
Dim cacheIndices() As Integer
Dim cacheInitialise As Boolean

Function CalculAvecCache(valeur As Integer) As Double
    ' Fonction coûteuse qui bénéficie d'un cache

    If Not cacheInitialise Then
        ReDim cacheCalculs(1 To 100)
        ReDim cacheIndices(1 To 100)
        cacheInitialise = True
    End If

    ' Vérifier si la valeur est déjà en cache
    Dim i As Integer
    For i = 1 To UBound(cacheIndices)
        If cacheIndices(i) = valeur Then
            CalculAvecCache = cacheCalculs(i)
            Debug.Print "Cache hit pour " & valeur
            Exit Function
        End If
    Next i

    ' Calcul coûteux (simulation)
    Dim resultat As Double
    resultat = valeur ^ 3 + Sqr(valeur) * 17.5  ' Calcul arbitraire

    ' Stocker en cache
    For i = 1 To UBound(cacheIndices)
        If cacheIndices(i) = 0 Then  ' Emplacement libre
            cacheIndices(i) = valeur
            cacheCalculs(i) = resultat
            Debug.Print "Cache miss pour " & valeur & " - stocké"
            Exit For
        End If
    Next i

    CalculAvecCache = resultat
End Function

Sub UtiliserCache()
    Debug.Print "=== Test du cache ==="

    ' Premier accès (cache miss)
    Debug.Print "Résultat 1 : " & CalculAvecCache(10)
    Debug.Print "Résultat 2 : " & CalculAvecCache(20)

    ' Deuxième accès (cache hit)
    Debug.Print "Résultat 1 bis : " & CalculAvecCache(10)
    Debug.Print "Résultat 2 bis : " & CalculAvecCache(20)
End Sub
```

### 2. Pipeline de transformation

```vba
Function Pipeline(donnees() As Variant, operations() As String, parametres() As Variant) As Variant
    ' Appliquer une série de transformations en séquence

    Dim resultat As Variant
    resultat = donnees

    Dim i As Integer
    For i = LBound(operations) To UBound(operations)
        resultat = TransformerTableau(resultat, operations(i), parametres(i))
    Next i

    Pipeline = resultat
End Function

Sub UtiliserPipeline()
    ' Données initiales
    Dim nombres(1 To 5) As Variant
    nombres(1) = 10: nombres(2) = 20: nombres(3) = 30: nombres(4) = 40: nombres(5) = 50

    ' Définir les opérations du pipeline
    Dim operations(1 To 3) As String
    Dim parametres(1 To 3) As Variant

    operations(1) = "multiplier": parametres(1) = 2      ' x2
    operations(2) = "ajouter": parametres(2) = 5         ' +5
    operations(3) = "puissance": parametres(3) = 1.5     ' ^1.5

    ' Appliquer le pipeline
    Dim resultat As Variant
    resultat = Pipeline(nombres, operations, parametres)

    Debug.Print "=== Pipeline de transformations ==="
    Dim i As Integer
    For i = LBound(nombres) To UBound(nombres)
        Debug.Print nombres(i) & " → " & Format(resultat(i), "0.00")
    Next i
End Sub
```

### 3. Tableaux en tant qu'objets de travail

```vba
Sub TableauCommeObjetTravail()
    ' Utiliser un tableau comme structure de travail temporaire

    ' Simulation : analyser des ventes par mois et par région
    Dim ventesBrutes(1 To 12, 1 To 5) As Double  ' 12 mois x 5 régions

    ' Remplir avec des données simulées
    Dim mois As Integer, region As Integer
    For mois = 1 To 12
        For region = 1 To 5
            ventesBrutes(mois, region) = Rnd() * 1000 + 500
        Next region
    Next mois

    ' Tableaux de travail pour les analyses
    Dim totalParMois(1 To 12) As Double
    Dim totalParRegion(1 To 5) As Double
    Dim moyenneParMois(1 To 12) As Double

    ' Calculs avec les tableaux de travail
    For mois = 1 To 12
        For region = 1 To 5
            totalParMois(mois) = totalParMois(mois) + ventesBrutes(mois, region)
            totalParRegion(region) = totalParRegion(region) + ventesBrutes(mois, region)
        Next region
        moyenneParMois(mois) = totalParMois(mois) / 5
    Next mois

    ' Trouver le meilleur mois
    Dim meilleurMois As Integer, maxVente As Double
    maxVente = 0
    For mois = 1 To 12
        If totalParMois(mois) > maxVente Then
            maxVente = totalParMois(mois)
            meilleurMois = mois
        End If
    Next mois

    Debug.Print "=== Analyse des ventes ==="
    Debug.Print "Meilleur mois : " & meilleurMois & " (" & Format(maxVente, "0.00") & "€)"

    ' Trouver la meilleure région
    Dim meilleureRegion As Integer, maxRegion As Double
    maxRegion = 0
    For region = 1 To 5
        If totalParRegion(region) > maxRegion Then
            maxRegion = totalParRegion(region)
            meilleureRegion = region
        End If
    Next region

    Debug.Print "Meilleure région : " & meilleureRegion & " (" & Format(maxRegion, "0.00") & "€)"
End Sub
```

---

## Récapitulatif et bonnes pratiques

### Principes fondamentaux du parcours

1. **For...Next** : Contrôle total, permet la modification
2. **For Each** : Simple mais lecture seule
3. **Parcours conditionnel** : Traiter seulement certains éléments
4. **Parcours à rebours** : Utile pour les suppressions

### Algorithmes essentiels à maîtriser

1. **Recherche linéaire** : Simple, fonctionne sur tout tableau
2. **Tri par sélection** : Efficace pour petites données
3. **Filtrage** : Extraire selon des critères
4. **Transformation** : Appliquer des fonctions à tous les éléments

### Optimisations de performance

1. **Éviter les redimensionnements répétitifs**
2. **Utiliser les bons types de données**
3. **Charger les données Excel en bloc**
4. **Implémenter des caches pour les calculs coûteux**

### Patterns recommandés

#### Parcours simple
```vba
For i = LBound(tableau) To UBound(tableau)
    ' Traitement de tableau(i)
Next i
```

#### Recherche avec résultat
```vba
Function Rechercher(arr() As Variant, val As Variant) As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            Rechercher = i
            Exit Function
        End If
    Next i
    Rechercher = -1
End Function
```

#### Transformation en place
```vba
For i = LBound(tableau) To UBound(tableau)
    tableau(i) = fonction(tableau(i))
Next i
```

### Erreurs courantes à éviter

❌ **Modifier un tableau pendant un For Each**
❌ **Oublier de vérifier les limites de tableau**
❌ **Redimensionner répétitivement sans stratégie**
❌ **Accéder à Excel cellule par cellule pour de gros volumes**
❌ **Ne pas gérer les tableaux vides**

### Conseils pour débuter

1. **Commencez simple** : Maîtrisez le parcours basique avant les algorithmes complexes
2. **Testez avec de petites données** : Vérifiez votre logique avant de passer aux gros volumes
3. **Utilisez Debug.Print** : Affichez les étapes intermédiaires pour comprendre
4. **Mesurez les performances** : Utilisez Timer pour comparer vos optimisations
5. **Réutilisez les patterns** : Créez des fonctions génériques pour les opérations courantes

### Progression recommandée

#### **Niveau débutant**
- Parcours simple avec For...Next
- Recherche linéaire de base
- Transformation simple (multiplication, addition)

#### **Niveau intermédiaire**
- Tri par sélection
- Filtrage avec critères
- Gestion des tableaux dynamiques

#### **Niveau avancé**
- Optimisations de performance
- Algorithmes spécialisés
- Patterns de cache et pipeline

La maîtrise du parcours et de la manipulation des tableaux vous donne un contrôle total sur vos données et vous permet de créer des solutions VBA sophistiquées et performantes. Ces techniques sont la base de tout traitement de données professionnel en VBA.

⏭️
