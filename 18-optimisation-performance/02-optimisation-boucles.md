🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.2 Optimisation des boucles

## Introduction

Les boucles sont l'une des structures les plus utilisées en VBA, mais aussi l'une des principales sources de ralentissement. Une boucle mal optimisée peut transformer une macro de quelques secondes en un processus de plusieurs minutes. Dans cette section, nous allons apprendre à identifier et corriger les problèmes de performance liés aux boucles.

## Comprendre le problème des boucles lentes

### Pourquoi les boucles peuvent être lentes

Le problème principal vient des interactions répétées entre VBA et Excel. Chaque fois que vous lisez ou écrivez dans une cellule depuis VBA, il y a un "coût" en temps. Multiplié par des milliers d'itérations, ce coût devient énorme.

### Exemple de boucle inefficace

```vba
Sub BoucleInefficace()
    Dim i As Long

    ' Cette boucle est TRÈS lente
    For i = 1 To 10000
        Cells(i, 1).Value = i                    ' Écriture cellule par cellule
        Cells(i, 2).Value = "Ligne " & i         ' Une autre écriture
        Cells(i, 3).Value = Cells(i, 1).Value * 2  ' Lecture puis écriture
    Next i
End Sub
```

Cette boucle fait 30 000 interactions avec Excel (10 000 × 3 opérations). C'est énormément !

## Technique #1 : Utiliser des tableaux en mémoire

### Principe de base

Au lieu de lire/écrire cellule par cellule, nous allons :
1. Charger les données d'Excel vers un tableau en mémoire
2. Traiter le tableau en mémoire (très rapide)
3. Réécrire tout le tableau vers Excel en une seule fois

### Exemple optimisé avec tableaux

```vba
Sub BoucleOptimiseeAvecTableau()
    Dim i As Long
    Dim donnees As Variant

    ' Créer un tableau pour 10000 lignes et 3 colonnes
    ReDim donnees(1 To 10000, 1 To 3)

    ' Traitement en mémoire (très rapide)
    For i = 1 To 10000
        donnees(i, 1) = i
        donnees(i, 2) = "Ligne " & i
        donnees(i, 3) = i * 2
    Next i

    ' Une seule écriture vers Excel pour toutes les données
    Range("A1:C10000").Value = donnees
End Sub
```

Cette version fait seulement 1 interaction avec Excel au lieu de 30 000 !

### Comparaison de performance

```vba
Sub ComparaisonPerformance()
    Dim tempsDebut As Double

    ' Test méthode lente
    tempsDebut = Timer
    BoucleInefficace
    Debug.Print "Méthode lente : " & Format(Timer - tempsDebut, "0.00") & " secondes"

    ' Test méthode optimisée
    tempsDebut = Timer
    BoucleOptimiseeAvecTableau
    Debug.Print "Méthode optimisée : " & Format(Timer - tempsDebut, "0.00") & " secondes"
End Sub
```

## Technique #2 : Traitement par blocs

### Quand utiliser cette technique

Parfois, vous ne pouvez pas traiter toutes les données en une fois (par exemple, si vous avez des millions de lignes). Dans ce cas, traitez par blocs de taille raisonnable.

### Exemple de traitement par blocs

```vba
Sub TraitementParBlocs()
    Dim i As Long, j As Long
    Dim tailleBloc As Long
    Dim donnees As Variant

    tailleBloc = 1000  ' Traiter 1000 lignes à la fois

    For i = 1 To 10000 Step tailleBloc
        ' Calculer la fin du bloc
        Dim finBloc As Long
        finBloc = i + tailleBloc - 1
        If finBloc > 10000 Then finBloc = 10000

        ' Créer le tableau pour ce bloc
        ReDim donnees(1 To finBloc - i + 1, 1 To 2)

        ' Traitement du bloc en mémoire
        For j = 1 To finBloc - i + 1
            donnees(j, 1) = i + j - 1
            donnees(j, 2) = "Bloc " & (i + j - 1)
        Next j

        ' Écriture du bloc vers Excel
        Range("A" & i & ":B" & finBloc).Value = donnees
    Next i
End Sub
```

## Technique #3 : Éviter les boucles imbriquées inutiles

### Le problème des boucles imbriquées

Les boucles imbriquées multiplient le nombre d'opérations. Une boucle dans une boucle peut rapidement devenir problématique.

### Exemple de boucle imbriquée inefficace

```vba
Sub BouclesImbriquees_Inefficace()
    Dim i As Long, j As Long

    ' Cette boucle fait 10000 × 10000 = 100 millions d'itérations !
    For i = 1 To 10000
        For j = 1 To 10000
            If Cells(i, 1).Value = Cells(j, 2).Value Then
                Cells(i, 3).Value = "Trouvé"
            End If
        Next j
    Next i
End Sub
```

### Version optimisée avec Collection

```vba
Sub BouclesOptimisees_AvecCollection()
    Dim i As Long
    Dim colRecherche As Collection
    Dim valeur As Variant

    ' Créer une collection pour la recherche rapide
    Set colRecherche = New Collection

    ' Première boucle : remplir la collection
    For i = 1 To 10000
        valeur = Cells(i, 2).Value
        If valeur <> "" Then
            On Error Resume Next
            colRecherche.Add valeur, CStr(valeur)
            On Error GoTo 0
        End If
    Next i

    ' Deuxième boucle : recherche rapide
    For i = 1 To 10000
        valeur = Cells(i, 1).Value
        On Error Resume Next
        colRecherche.Item(CStr(valeur))
        If Err.Number = 0 Then
            Cells(i, 3).Value = "Trouvé"
        End If
        On Error GoTo 0
    Next i
End Sub
```

## Technique #4 : Optimiser les conditions dans les boucles

### Éviter les calculs répétés

```vba
Sub ConditionsInefficaces()
    Dim i As Long

    For i = 1 To 10000
        ' INEFFICACE : UCase est appelé 10000 fois
        If UCase(Range("Z1").Value) = "ACTIF" Then
            Cells(i, 1).Value = i
        End If
    Next i
End Sub
```

### Version optimisée

```vba
Sub ConditionsOptimisees()
    Dim i As Long
    Dim estActif As Boolean

    ' Calculer une seule fois avant la boucle
    estActif = (UCase(Range("Z1").Value) = "ACTIF")

    For i = 1 To 10000
        If estActif Then
            Cells(i, 1).Value = i
        End If
    Next i
End Sub
```

## Technique #5 : Choisir le bon type de boucle

### For Next vs For Each

**For Next** est généralement plus rapide pour les plages de cellules :

```vba
Sub ForNextRapide()
    Dim i As Long
    Dim plage As Range
    Set plage = Range("A1:A10000")

    ' Rapide pour les indices numériques
    For i = 1 To plage.Rows.Count
        plage.Cells(i, 1).Value = i
    Next i
End Sub
```

**For Each** est pratique mais peut être plus lent :

```vba
Sub ForEachPlusLent()
    Dim cellule As Range
    Dim compteur As Long

    compteur = 1
    ' Plus lent pour de gros volumes
    For Each cellule In Range("A1:A10000")
        cellule.Value = compteur
        compteur = compteur + 1
    Next cellule
End Sub
```

## Technique #6 : Utiliser les fonctions Excel intégrées

### Quand Excel peut faire mieux que VBA

Parfois, Excel a des fonctions intégrées plus rapides que vos boucles VBA.

### Exemple : Remplissage de séries

```vba
Sub RemplissageLent_AvecBoucle()
    Dim i As Long

    ' Lent : boucle VBA
    For i = 1 To 10000
        Cells(i, 1).Value = i
    Next i
End Sub

Sub RemplissageRapide_SansVBA()
    ' Rapide : fonction Excel native
    Range("A1").Value = 1
    Range("A2").Value = 2
    Range("A1:A2").AutoFill Range("A1:A10000"), xlFillSeries
End Sub
```

### Exemple : Formules en lot

```vba
Sub FormulesLentes_CelluleParCellule()
    Dim i As Long

    ' Lent : formule cellule par cellule
    For i = 1 To 10000
        Cells(i, 2).Formula = "=A" & i & "*2"
    Next i
End Sub

Sub FormulesRapides_EnBloc()
    ' Rapide : formule appliquée à toute la plage
    ' Excel ajuste automatiquement les références relatives pour chaque cellule
    Range("B1:B10000").Formula = "=A1*2"
End Sub
```

## Technique #7 : Réduire les accès aux propriétés d'objets

### Problème des accès répétés

```vba
Sub AccesRepetes_Inefficace()
    Dim i As Long

    For i = 1 To 1000
        ' Chaque ligne accède 3 fois à ActiveSheet
        ActiveSheet.Cells(i, 1).Value = i
        ActiveSheet.Cells(i, 2).Value = ActiveSheet.Cells(i, 1).Value * 2
        ActiveSheet.Cells(i, 3).Value = "Ligne " & i
    Next i
End Sub
```

### Version optimisée avec variable d'objet

```vba
Sub AccesOptimise_AvecVariable()
    Dim i As Long
    Dim ws As Worksheet

    ' Une seule référence à la feuille
    Set ws = ActiveSheet

    For i = 1 To 1000
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = ws.Cells(i, 1).Value * 2
        ws.Cells(i, 3).Value = "Ligne " & i
    Next i
End Sub
```

## Technique #8 : Gérer les boucles conditionnelles

### Sortir des boucles dès que possible

```vba
Sub BoucleAvecSortiePrecoce()
    Dim i As Long
    Dim valeurTrouvee As Boolean

    valeurTrouvee = False

    For i = 1 To 10000
        If Cells(i, 1).Value = "STOP" Then
            valeurTrouvee = True
            Exit For  ' Sortir immédiatement, ne pas continuer inutilement
        End If

        ' Traitement uniquement si nécessaire
        If Not valeurTrouvee Then
            Cells(i, 2).Value = "Traité"
        End If
    Next i
End Sub
```

### Utiliser les instructions de contrôle

```vba
Sub BoucleAvecContinue()
    Dim i As Long

    For i = 1 To 1000
        ' Ignorer les lignes vides
        If Cells(i, 1).Value = "" Then
            GoTo SuiteBoucle  ' Équivalent de "Continue" dans d'autres langages
        End If

        ' Traitement seulement pour les cellules non vides
        Cells(i, 2).Value = UCase(Cells(i, 1).Value)

SuiteBoucle:
    Next i
End Sub
```

## Mesure de l'efficacité des optimisations

### Fonction de chronométrage

```vba
Function ChronometrerBoucle(nbIterations As Long) As Double
    Dim tempsDebut As Double
    Dim i As Long

    tempsDebut = Timer

    ' Votre boucle à tester ici
    For i = 1 To nbIterations
        ' Code de test
        Cells(i, 1).Value = i
    Next i

    ChronometrerBoucle = Timer - tempsDebut
End Function
```

### Comparaison systématique

```vba
Sub TesterToutesLesMethodes()
    Dim temps1 As Double, temps2 As Double, temps3 As Double

    ' Test 1 : Méthode cellule par cellule
    Application.ScreenUpdating = False
    temps1 = ChronometrerBoucle(5000)

    ' Test 2 : Méthode avec tableaux
    ' ... votre code optimisé
    temps2 = ChronometrerBoucle(5000)

    ' Test 3 : Méthode avec fonctions Excel
    ' ... votre code avec fonctions natives
    temps3 = ChronometrerBoucle(5000)

    Application.ScreenUpdating = True

    ' Afficher les résultats
    Debug.Print "Cellule par cellule : " & Format(temps1, "0.00") & "s"
    Debug.Print "Avec tableaux : " & Format(temps2, "0.00") & "s"
    Debug.Print "Fonctions natives : " & Format(temps3, "0.00") & "s"
    Debug.Print "Gain avec tableaux : " & Format(temps1 / temps2, "0.0") & "x plus rapide"
End Sub
```

## Bonnes pratiques pour les boucles

### Règles d'or

1. **Minimiser les interactions Excel** : Utilisez des tableaux en mémoire
2. **Déclarer les variables avec le bon type** : `Long` au lieu de `Integer` pour les compteurs
3. **Calculer une seule fois** : Sortez les calculs constants de la boucle
4. **Utiliser des variables d'objets** : Évitez `ActiveSheet.Cells(i,1)` répétés
5. **Préférer les fonctions natives Excel** quand c'est possible

### Variables et déclarations optimisées

```vba
Sub BonnesDeclarations()
    ' BON : Long pour les gros nombres
    Dim i As Long, j As Long

    ' BON : Variables d'objets pour éviter les accès répétés
    Dim ws As Worksheet
    Dim plage As Range

    ' BON : Tableaux typés pour les performances
    Dim donnees() As Variant

    Set ws = ActiveSheet
    Set plage = ws.Range("A1:C1000")

    ' Votre boucle optimisée ici
End Sub
```

### Gestion mémoire dans les boucles

```vba
Sub GestionMemoireBoucle()
    Dim donnees As Variant
    Dim resultat As Variant

    ' Dimensionner correctement dès le début
    ReDim donnees(1 To 10000, 1 To 5)
    ReDim resultat(1 To 10000, 1 To 3)

    ' Traitement en mémoire
    Dim i As Long
    For i = 1 To 10000
        ' Traitement direct dans le tableau
        resultat(i, 1) = donnees(i, 1) * 2
        resultat(i, 2) = donnees(i, 2) & " traité"
        resultat(i, 3) = i
    Next i

    ' Une seule écriture finale
    Range("D1:F10000").Value = resultat

    ' Libérer la mémoire si nécessaire pour de très gros tableaux
    Erase donnees
    Erase resultat
End Sub
```

## Cas particuliers et pièges à éviter

### Piège 1 : Les formules en boucle

```vba
' ÉVITER : Formule recalculée à chaque itération
For i = 1 To 1000
    Cells(i, 1).Formula = "=SUM(B" & i & ":E" & i & ")"
Next i

' PRÉFÉRER : Formule appliquée en une fois
Range("A1:A1000").Formula = "=SUM(B1:E1)"
```

### Piège 2 : Les recherches répétées

```vba
' ÉVITER : VLOOKUP dans chaque itération
For i = 1 To 1000
    Cells(i, 2).Value = Application.VLookup(Cells(i, 1).Value, Range("F:G"), 2, False)
Next i

' PRÉFÉRER : Formule VLOOKUP appliquée à toute la plage
' Excel ajuste automatiquement A1 → A2, A3... pour chaque ligne
Range("B1:B1000").Formula = "=VLOOKUP(A1,F:G,2,FALSE)"
```

## Résumé des techniques d'optimisation des boucles

1. **Tableaux en mémoire** : La technique la plus efficace pour la plupart des cas
2. **Traitement par blocs** : Pour les très gros volumes de données
3. **Éviter les boucles imbriquées** : Utiliser des collections ou dictionnaires
4. **Optimiser les conditions** : Calculer une seule fois les valeurs constantes
5. **Choisir le bon type de boucle** : For Next généralement plus rapide que For Each
6. **Fonctions Excel natives** : Souvent plus rapides que VBA pour certaines opérations
7. **Variables d'objets** : Réduire les accès répétés aux propriétés
8. **Sorties précoces** : Exit For dès que possible

En appliquant ces techniques, vous pouvez transformer des boucles qui prennent des minutes en processus de quelques secondes. L'investissement en temps pour optimiser est généralement largement rentabilisé par les gains de performance obtenus.

⏭️ [Gestion de la mémoire](/18-optimisation-performance/03-gestion-memoire.md)
