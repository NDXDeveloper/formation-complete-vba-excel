🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.5. Tests unitaires simples

## Qu'est-ce qu'un test unitaire ?

Un test unitaire est un petit programme qui **vérifie automatiquement** qu'une partie spécifique de votre code (généralement une fonction ou une procédure) fonctionne correctement. C'est comme avoir un assistant qui teste systématiquement chaque pièce de votre programme pour s'assurer qu'elle fait bien ce qu'elle est censée faire.

Imaginez que vous construisez une voiture : avant d'assembler tous les éléments, vous testez chaque pièce individuellement - les freins, le moteur, les phares, etc. Les tests unitaires fonctionnent de la même manière avec votre code : ils testent chaque "pièce" (fonction) séparément.

## Pourquoi faire des tests unitaires ?

**Détection précoce d'erreurs** : Les tests vous avertissent immédiatement si vous cassez quelque chose en modifiant votre code.

**Confiance dans le code** : Quand tous vos tests passent, vous savez que votre code fonctionne comme prévu.

**Documentation vivante** : Les tests montrent comment utiliser vos fonctions et quels résultats attendre.

**Facilitation des modifications** : Vous pouvez modifier votre code en toute sécurité, sachant que les tests détecteront les problèmes.

**Gain de temps à long terme** : Bien que l'écriture de tests prenne du temps au début, elle vous en fait gagner énormément par la suite.

## Les principes de base des tests unitaires

### Une fonction = Un test (ou plusieurs)
Chaque fonction importante de votre code devrait avoir au moins un test qui vérifie son comportement.

### Tests indépendants
Chaque test doit pouvoir s'exécuter seul, sans dépendre d'autres tests.

### Tests répétables
Un test doit donner le même résultat à chaque exécution (avec les mêmes données d'entrée).

### Tests lisibles
Un test doit être facile à comprendre - son nom et son contenu doivent expliquer clairement ce qui est testé.

## Structure d'un test unitaire simple

Un test unitaire suit généralement cette structure :

1. **Arrange** (Préparer) : Préparer les données d'entrée
2. **Act** (Agir) : Exécuter la fonction à tester
3. **Assert** (Vérifier) : Vérifier que le résultat est correct

```vba
Sub TestNomDeLaFonction()
    ' 1. Arrange - Préparer les données
    Dim entree As [Type]
    Dim resultatAttendu As [Type]

    entree = [valeur]
    resultatAttendu = [valeur attendue]

    ' 2. Act - Exécuter la fonction
    Dim resultatObtenu As [Type]
    resultatObtenu = NomDeLaFonction(entree)

    ' 3. Assert - Vérifier le résultat
    If resultatObtenu = resultatAttendu Then
        Debug.Print "✓ Test réussi : " & "TestNomDeLaFonction"
    Else
        Debug.Print "✗ Test échoué : " & "TestNomDeLaFonction"
        Debug.Print "  Attendu : " & resultatAttendu
        Debug.Print "  Obtenu  : " & resultatObtenu
    End If
End Sub
```

## Exemple concret - Tester une fonction simple

Supposons que vous avez cette fonction :

```vba
Function CalculerTVA(prixHT As Double, tauxTVA As Double) As Double
    CalculerTVA = prixHT * (tauxTVA / 100)
End Function
```

Voici comment créer un test unitaire pour cette fonction :

```vba
Sub TestCalculerTVA_CasNormal()
    ' 1. Arrange - Préparer les données de test
    Dim prixHT As Double
    Dim tauxTVA As Double
    Dim resultatAttendu As Double

    prixHT = 100
    tauxTVA = 20
    resultatAttendu = 20  ' 100 * (20/100) = 20

    ' 2. Act - Exécuter la fonction à tester
    Dim resultatObtenu As Double
    resultatObtenu = CalculerTVA(prixHT, tauxTVA)

    ' 3. Assert - Vérifier le résultat
    If resultatObtenu = resultatAttendu Then
        Debug.Print "✓ TestCalculerTVA_CasNormal : RÉUSSI"
    Else
        Debug.Print "✗ TestCalculerTVA_CasNormal : ÉCHOUÉ"
        Debug.Print "  Attendu : " & resultatAttendu
        Debug.Print "  Obtenu  : " & resultatObtenu
    End If
End Sub
```

## Créer une fonction d'aide pour les assertions

Pour simplifier l'écriture des tests, vous pouvez créer une fonction d'aide :

```vba
Sub AssertEqual(valeurObtenue As Variant, valeurAttendue As Variant, nomTest As String)
    If valeurObtenue = valeurAttendue Then
        Debug.Print "✓ " & nomTest & " : RÉUSSI"
    Else
        Debug.Print "✗ " & nomTest & " : ÉCHOUÉ"
        Debug.Print "  Attendu : " & valeurAttendue
        Debug.Print "  Obtenu  : " & valeurObtenue
    End If
End Sub
```

Maintenant vos tests deviennent plus simples :

```vba
Sub TestCalculerTVA_CasNormal()
    ' Arrange
    Dim prixHT As Double, tauxTVA As Double
    prixHT = 100
    tauxTVA = 20

    ' Act
    Dim resultat As Double
    resultat = CalculerTVA(prixHT, tauxTVA)

    ' Assert
    AssertEqual resultat, 20, "TestCalculerTVA_CasNormal"
End Sub
```

## Tester différents cas de figure

Une bonne pratique consiste à tester plusieurs scénarios pour chaque fonction :

### Cas normal (Happy Path)
```vba
Sub TestCalculerTVA_CasNormal()
    Dim resultat As Double
    resultat = CalculerTVA(100, 20)
    AssertEqual resultat, 20, "TestCalculerTVA_CasNormal"
End Sub
```

### Cas limites (Edge Cases)
```vba
Sub TestCalculerTVA_PrixZero()
    Dim resultat As Double
    resultat = CalculerTVA(0, 20)
    AssertEqual resultat, 0, "TestCalculerTVA_PrixZero"
End Sub

Sub TestCalculerTVA_TVAZero()
    Dim resultat As Double
    resultat = CalculerTVA(100, 0)
    AssertEqual resultat, 0, "TestCalculerTVA_TVAZero"
End Sub
```

### Cas d'erreur
```vba
Sub TestCalculerTVA_ValeurNegative()
    ' Ici, on pourrait tester comment la fonction réagit
    ' aux valeurs négatives (selon les règles métier)
    Dim resultat As Double
    resultat = CalculerTVA(-100, 20)
    AssertEqual resultat, -20, "TestCalculerTVA_ValeurNegative"
End Sub
```

## Tester des fonctions qui retournent des chaînes

```vba
Function FormaterNomComplet(prenom As String, nom As String) As String
    FormaterNomComplet = UCase(nom) & ", " & prenom
End Function

Sub TestFormaterNomComplet()
    Dim resultat As String
    resultat = FormaterNomComplet("Jean", "Dupont")
    AssertEqual resultat, "DUPONT, Jean", "TestFormaterNomComplet"
End Sub
```

## Tester des fonctions qui manipulent des dates

```vba
Function AjouterJoursOuvrables(dateDebut As Date, nombreJours As Integer) As Date
    Dim dateActuelle As Date
    Dim joursAjoutes As Integer

    dateActuelle = dateDebut
    joursAjoutes = 0

    Do While joursAjoutes < nombreJours
        dateActuelle = dateActuelle + 1
        ' Ignore les week-ends (samedi = 7, dimanche = 1)
        If Weekday(dateActuelle) <> 1 And Weekday(dateActuelle) <> 7 Then
            joursAjoutes = joursAjoutes + 1
        End If
    Loop

    AjouterJoursOuvrables = dateActuelle
End Function

Sub TestAjouterJoursOuvrables()
    ' Test : ajouter 5 jours ouvrables à partir d'un lundi
    Dim dateDebut As Date
    Dim resultat As Date
    Dim attendu As Date

    dateDebut = DateSerial(2024, 1, 8)  ' Lundi 8 janvier 2024
    attendu = DateSerial(2024, 1, 15)   ' Lundi 15 janvier 2024

    resultat = AjouterJoursOuvrables(dateDebut, 5)

    AssertEqual resultat, attendu, "TestAjouterJoursOuvrables"
End Sub
```

## Créer une suite de tests

Pour organiser vos tests, créez une procédure qui exécute tous les tests :

```vba
Sub ExecuterTousLesTests()
    Debug.Print "=== DÉBUT DES TESTS ==="
    Debug.Print ""

    ' Tests pour CalculerTVA
    Debug.Print "Tests CalculerTVA :"
    TestCalculerTVA_CasNormal
    TestCalculerTVA_PrixZero
    TestCalculerTVA_TVAZero
    Debug.Print ""

    ' Tests pour FormaterNomComplet
    Debug.Print "Tests FormaterNomComplet :"
    TestFormaterNomComplet
    Debug.Print ""

    ' Tests pour AjouterJoursOuvrables
    Debug.Print "Tests AjouterJoursOuvrables :"
    TestAjouterJoursOuvrables
    Debug.Print ""

    Debug.Print "=== FIN DES TESTS ==="
End Sub
```

## Améliorer le système de test avec un compteur

```vba
' Variables globales pour compter les résultats
Dim testsExecutes As Integer
Dim testsReussis As Integer

Sub InitialiserTests()
    testsExecutes = 0
    testsReussis = 0
End Sub

Sub AssertEqualAvecCompteur(valeurObtenue As Variant, valeurAttendue As Variant, nomTest As String)
    testsExecutes = testsExecutes + 1

    If valeurObtenue = valeurAttendue Then
        testsReussis = testsReussis + 1
        Debug.Print "✓ " & nomTest & " : RÉUSSI"
    Else
        Debug.Print "✗ " & nomTest & " : ÉCHOUÉ"
        Debug.Print "  Attendu : " & valeurAttendue
        Debug.Print "  Obtenu  : " & valeurObtenue
    End If
End Sub

Sub AfficherResultatsFinaux()
    Debug.Print ""
    Debug.Print "=== RÉSUMÉ ==="
    Debug.Print "Tests exécutés : " & testsExecutes
    Debug.Print "Tests réussis  : " & testsReussis
    Debug.Print "Tests échoués  : " & (testsExecutes - testsReussis)

    If testsReussis = testsExecutes Then
        Debug.Print "🎉 TOUS LES TESTS SONT PASSÉS !"
    Else
        Debug.Print "⚠️ Certains tests ont échoué"
    End If
End Sub
```

## Tester des procédures qui modifient des données

Pour tester des procédures qui modifient des cellules Excel :

```vba
Sub RemplirTableauMultiplication(feuille As Worksheet, taille As Integer)
    Dim i As Integer, j As Integer

    For i = 1 To taille
        For j = 1 To taille
            feuille.Cells(i, j).Value = i * j
        Next j
    Next i
End Sub

Sub TestRemplirTableauMultiplication()
    ' Arrange - Créer une feuille de test
    Dim feuilleTest As Worksheet
    Set feuilleTest = Worksheets.Add
    feuilleTest.Name = "TestTemp"

    ' Act - Exécuter la procédure
    RemplirTableauMultiplication feuilleTest, 3

    ' Assert - Vérifier quelques valeurs
    AssertEqual feuilleTest.Cells(1, 1).Value, 1, "TestTableau_1x1"
    AssertEqual feuilleTest.Cells(2, 3).Value, 6, "TestTableau_2x3"
    AssertEqual feuilleTest.Cells(3, 3).Value, 9, "TestTableau_3x3"

    ' Nettoyage - Supprimer la feuille de test
    Application.DisplayAlerts = False
    feuilleTest.Delete
    Application.DisplayAlerts = True
End Sub
```

## Bonnes pratiques pour les tests unitaires

### Nommage des tests
Utilisez des noms descriptifs qui expliquent ce qui est testé :
- `TestCalculerTVA_CasNormal`
- `TestCalculerTVA_PrixZero`
- `TestFormaterNom_AvecEspaces`

### Tests indépendants
Chaque test doit pouvoir s'exécuter seul :
```vba
' Bon - le test prépare ses propres données
Sub TestCalculerRemise()
    Dim prix As Double
    prix = 100
    ' ... reste du test
End Sub

' Mauvais - dépend d'une variable globale
Dim prixGlobal As Double

Sub TestCalculerRemise()
    ' Utilise prixGlobal - problématique !
End Sub
```

### Nettoyage après les tests
Si vos tests créent des fichiers, feuilles, ou modifient l'état d'Excel, nettoyez après :
```vba
Sub TestAvecNettoyage()
    ' Test qui modifie des données

    ' Nettoyage à la fin
    Range("A1:C10").Clear
End Sub
```

### Tests rapides
Les tests doivent s'exécuter rapidement. Évitez :
- Les boucles très longues
- Les interactions avec des fichiers externes
- Les calculs très complexes

## Limitations des tests unitaires simples en VBA

**Pas de framework dédié** : VBA n'a pas de framework de test intégré comme d'autres langages.

**Gestion d'erreurs manuelle** : Vous devez gérer vous-même les erreurs et exceptions.

**Pas de parallélisation** : Les tests s'exécutent un par un.

**Interface limitée** : Pas d'interface graphique sophistiquée pour les résultats.

## Quand écrire des tests

**Fonctions critiques** : Toute fonction importante pour votre application.

**Calculs complexes** : Fonctions avec des formules mathématiques ou logiques complexes.

**Fonctions réutilisables** : Code que vous utilisez dans plusieurs endroits.

**Après correction de bugs** : Créez un test pour vous assurer que le bug ne revient pas.

**Avant refactoring** : Tests pour vous assurer que les modifications ne cassent rien.

## Exemple complet d'organisation

```vba
' =====================================
' MODULE DE FONCTIONS À TESTER
' =====================================

Function CalculerRemise(prix As Double, pourcentage As Double) As Double
    If prix < 0 Or pourcentage < 0 Or pourcentage > 100 Then
        CalculerRemise = -1  ' Erreur
    Else
        CalculerRemise = prix * (pourcentage / 100)
    End If
End Function

' =====================================
' MODULE DE TESTS
' =====================================

' Variables pour les statistiques
Dim totalTests As Integer
Dim testsOK As Integer

Sub ExecuterTousLesTestsRemise()
    InitialiserCompteurs

    Debug.Print "=== TESTS POUR CalculerRemise ==="

    TestCalculerRemise_CasNormal
    TestCalculerRemise_PrixZero
    TestCalculerRemise_RemiseZero
    TestCalculerRemise_PrixNegatif
    TestCalculerRemise_RemiseNegative
    TestCalculerRemise_RemiseTropElevee

    AfficherBilan
End Sub

Sub InitialiserCompteurs()
    totalTests = 0
    testsOK = 0
End Sub

Sub TestCalculerRemise_CasNormal()
    Dim resultat As Double
    resultat = CalculerRemise(100, 10)
    Verifier resultat, 10, "CasNormal"
End Sub

Sub TestCalculerRemise_PrixZero()
    Dim resultat As Double
    resultat = CalculerRemise(0, 10)
    Verifier resultat, 0, "PrixZero"
End Sub

Sub TestCalculerRemise_RemiseZero()
    Dim resultat As Double
    resultat = CalculerRemise(100, 0)
    Verifier resultat, 0, "RemiseZero"
End Sub

Sub TestCalculerRemise_PrixNegatif()
    Dim resultat As Double
    resultat = CalculerRemise(-100, 10)
    Verifier resultat, -1, "PrixNegatif"
End Sub

Sub TestCalculerRemise_RemiseNegative()
    Dim resultat As Double
    resultat = CalculerRemise(100, -10)
    Verifier resultat, -1, "RemiseNegative"
End Sub

Sub TestCalculerRemise_RemiseTropElevee()
    Dim resultat As Double
    resultat = CalculerRemise(100, 150)
    Verifier resultat, -1, "RemiseTropElevee"
End Sub

Sub Verifier(obtenu As Double, attendu As Double, nomTest As String)
    totalTests = totalTests + 1

    If obtenu = attendu Then
        testsOK = testsOK + 1
        Debug.Print "✓ " & nomTest & " : RÉUSSI"
    Else
        Debug.Print "✗ " & nomTest & " : ÉCHOUÉ (attendu: " & attendu & ", obtenu: " & obtenu & ")"
    End If
End Sub

Sub AfficherBilan()
    Debug.Print ""
    Debug.Print "=== BILAN ==="
    Debug.Print totalTests & " tests exécutés"
    Debug.Print testsOK & " tests réussis"
    Debug.Print (totalTests - testsOK) & " tests échoués"

    If testsOK = totalTests Then
        Debug.Print "🎉 SUCCÈS TOTAL !"
    End If
End Sub
```

Les tests unitaires simples en VBA vous permettent de créer un filet de sécurité autour de votre code. Bien qu'ils ne remplacent pas des frameworks de test professionnels, ils constituent un excellent point de départ pour améliorer la qualité et la fiabilité de vos programmes VBA.

⏭️
