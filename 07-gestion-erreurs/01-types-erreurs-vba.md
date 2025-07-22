🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.1. Types d'erreurs en VBA

## Introduction aux types d'erreurs

Comprendre les différents types d'erreurs est essentiel pour bien les gérer. En VBA, les erreurs ne sont pas toutes identiques : certaines peuvent être évitées en écrivant mieux le code, d'autres surviennent pendant l'exécution et ne peuvent être anticipées qu'en partie. Identifier le type d'erreur vous aide à choisir la meilleure stratégie pour la traiter.

**Analogie simple :**
Imaginez que vous conduisez une voiture. Il y a différents types de problèmes :
- **Erreur de syntaxe** = Ne pas savoir comment démarrer (problème de connaissance)
- **Erreur d'exécution** = Tomber en panne d'essence en route (problème pendant le trajet)
- **Erreur logique** = Prendre la mauvaise direction (le trajet se déroule, mais pas comme prévu)

---

## 1. Erreurs de syntaxe (Syntax Errors)

### Qu'est-ce qu'une erreur de syntaxe ?

Les **erreurs de syntaxe** sont des fautes dans l'écriture du code VBA. Elles empêchent le code de s'exécuter car VBA ne comprend pas ce que vous voulez dire. C'est comme écrire une phrase avec une grammaire incorrecte.

### Caractéristiques des erreurs de syntaxe

- **Détectées avant l'exécution** : VBA les repère dès que vous tapez le code
- **Soulignées en rouge** : L'éditeur VBA marque les erreurs en temps réel
- **Code ne peut pas s'exécuter** : Impossible de lancer la macro tant que l'erreur persiste

### Exemples courants d'erreurs de syntaxe

#### 1. Parenthèses manquantes ou mal placées

```vba
' INCORRECT - Parenthèse fermante manquante
If Range("A1").Value > 10 Then
    MsgBox("Valeur élevée"
End If

' CORRECT
If Range("A1").Value > 10 Then
    MsgBox("Valeur élevée")
End If
```

#### 2. Mots-clés mal orthographiés

```vba
' INCORRECT - "Thn" au lieu de "Then"
If Range("A1").Value > 10 Thn
    MsgBox "Valeur élevée"
End If

' CORRECT
If Range("A1").Value > 10 Then
    MsgBox "Valeur élevée"
End If
```

#### 3. Structure incomplète

```vba
' INCORRECT - If sans End If
If Range("A1").Value > 10 Then
    MsgBox "Valeur élevée"
' End If manquant

' CORRECT
If Range("A1").Value > 10 Then
    MsgBox "Valeur élevée"
End If
```

#### 4. Guillemets non fermés

```vba
' INCORRECT - Guillemet fermant manquant
MsgBox "Bonjour monde

' CORRECT
MsgBox "Bonjour monde"
```

#### 5. Déclaration de variable incorrecte

```vba
' INCORRECT - Syntaxe invalide
Dim As Integer monNombre

' CORRECT
Dim monNombre As Integer
```

### Comment éviter les erreurs de syntaxe

1. **Utilisez l'indentation** pour visualiser la structure
2. **Activez la vérification automatique** : Outils > Options > Éditeur
3. **Écrivez ligne par ligne** et testez régulièrement
4. **Utilisez l'auto-complétion** de VBA
5. **Relisez votre code** avant de l'exécuter

---

## 2. Erreurs d'exécution (Runtime Errors)

### Qu'est-ce qu'une erreur d'exécution ?

Les **erreurs d'exécution** surviennent pendant que le code s'exécute. Le code est syntaxiquement correct, mais quelque chose empêche une instruction de se dérouler normalement. C'est comme suivre une recette correctement écrite mais découvrir qu'un ingrédient est périmé.

### Caractéristiques des erreurs d'exécution

- **Surviennent pendant l'exécution** : Le code commence à s'exécuter puis s'arrête
- **Peuvent être imprévisibles** : Dépendent des conditions du moment
- **Message d'erreur affiché** : VBA affiche un numéro et une description
- **Arrêtent l'exécution** : Le programme s'interrompt à la ligne problématique

### Exemples courants d'erreurs d'exécution

#### 1. Erreur 9 : "Subscript out of range" (Indice hors limites)

```vba
' Cette erreur survient si la feuille "Inexistante" n'existe pas
Sub ExempleErreur9()
    Worksheets("Inexistante").Range("A1").Value = "Test"  ' Erreur 9
End Sub

' Autres causes d'erreur 9
Sub AutresErreurs9()
    ' Accéder à un élément de tableau hors limites
    Dim monTableau(1 To 5) As Integer
    monTableau(10) = 100  ' Erreur 9 - l'index 10 n'existe pas

    ' Accéder à un classeur fermé
    Workbooks("FichierFermé.xlsx").Activate  ' Erreur 9
End Sub
```

#### 2. Erreur 11 : "Division by zero" (Division par zéro)

```vba
Sub ExempleErreur11()
    Dim resultat As Double
    Dim diviseur As Double

    diviseur = 0
    resultat = 10 / diviseur  ' Erreur 11

    Range("A1").Value = resultat
End Sub
```

#### 3. Erreur 13 : "Type mismatch" (Non-correspondance de type)

```vba
Sub ExempleErreur13()
    Dim nombre As Integer

    ' Tentative d'assigner du texte à une variable numérique
    nombre = "Bonjour"  ' Erreur 13

    ' Ou tentative de calcul avec du texte
    Range("A1").Value = "Texte"
    Dim resultat As Double
    resultat = Range("A1").Value * 2  ' Erreur 13 si A1 contient du texte
End Sub
```

#### 4. Erreur 1004 : Erreur définie par l'application ou par l'objet

```vba
Sub ExempleErreur1004()
    ' Tentative de copier vers une plage de taille différente
    Range("A1:A3").Copy Range("B1:B5")  ' Erreur 1004

    ' Tentative d'accéder à une plage invalide
    Range("A0").Select  ' Erreur 1004 - A0 n'existe pas

    ' Tentative de modification d'une feuille protégée
    ActiveSheet.Protect "motdepasse"
    Range("A1").Value = "Test"  ' Erreur 1004 si la feuille est protégée
End Sub
```

#### 5. Erreur 53 : "File not found" (Fichier non trouvé)

```vba
Sub ExempleErreur53()
    ' Tentative d'ouverture d'un fichier inexistant
    Workbooks.Open "C:\FichierInexistant.xlsx"  ' Erreur 53
End Sub
```

#### 6. Erreur 70 : "Permission denied" (Autorisation refusée)

```vba
Sub ExempleErreur70()
    ' Tentative de modification d'un fichier en lecture seule
    ' ou ouvert par un autre utilisateur
    Workbooks.Open "C:\FichierEnLectureSeule.xlsx"
    ActiveWorkbook.Save  ' Erreur 70 si le fichier est en lecture seule
End Sub
```

---

## 3. Erreurs logiques (Logic Errors)

### Qu'est-ce qu'une erreur logique ?

Les **erreurs logiques** sont les plus sournoises. Le code s'exécute sans erreur, mais ne fait pas ce que vous vouliez. C'est comme suivre parfaitement une recette, mais se tromper d'ingrédient - le plat se prépare, mais le goût n'est pas celui attendu.

### Caractéristiques des erreurs logiques

- **Aucun message d'erreur** : VBA ne détecte rien d'anormal
- **Code s'exécute complètement** : Pas d'interruption
- **Résultats incorrects** : Les calculs ou actions ne correspondent pas à l'intention
- **Difficiles à détecter** : Nécessitent une vérification manuelle des résultats

### Exemples courants d'erreurs logiques

#### 1. Condition incorrecte

```vba
Sub ExempleErreurLogique1()
    Dim note As Integer
    note = 85

    ' INCORRECT - Intention : afficher "Réussi" si note >= 60
    If note <= 60 Then  ' Erreur logique : devrait être >=
        MsgBox "Réussi"
    Else
        MsgBox "Échoué"
    End If
    ' Résultat : affiche "Échoué" alors que 85 devrait être "Réussi"
End Sub
```

#### 2. Boucle infinie ou mal contrôlée

```vba
Sub ExempleErreurLogique2()
    Dim i As Integer
    i = 1

    ' INCORRECT - Boucle qui ne s'arrête jamais
    Do While i <= 10
        Range("A" & i).Value = i
        ' OUBLI : i = i + 1  (la variable i ne change jamais)
    Loop
    ' Cette boucle continue indéfiniment car i reste toujours 1
End Sub
```

#### 3. Mauvais ordre d'opérations

```vba
Sub ExempleErreurLogique3()
    ' Intention : calculer une remise de 10% puis ajouter une taxe de 5%
    Dim prix As Double
    Dim prixFinal As Double

    prix = 100

    ' INCORRECT - L'ordre des opérations change le résultat
    prixFinal = prix * 1.05 * 0.9  ' Taxe puis remise
    ' Résultat : 94.5

    ' CORRECT - Remise puis taxe
    ' prixFinal = prix * 0.9 * 1.05
    ' Résultat : 94.5 (même résultat par coïncidence, mais concept différent)
End Sub
```

#### 4. Confusion entre références relatives et absolues

```vba
Sub ExempleErreurLogique4()
    Dim i As Integer

    ' Intention : copier A1 vers B1, B2, B3...
    For i = 1 To 3
        Range("A1").Copy Range("B" & i)
    Next i

    ' Ce code fonctionne, mais si l'intention était de copier
    ' A1 vers B1, A2 vers B2, A3 vers B3, c'est une erreur logique
    ' Il faudrait : Range("A" & i).Copy Range("B" & i)
End Sub
```

#### 5. Calculs avec des types de données inappropriés

```vba
Sub ExempleErreurLogique5()
    ' Intention : calculer une moyenne
    Dim total As Integer  ' ERREUR : Integer peut causer des arrondis
    Dim moyenne As Integer  ' ERREUR : Integer pour une moyenne

    total = 7 + 8 + 9
    moyenne = total / 3  ' Résultat : 8 au lieu de 8.33

    ' CORRECT : utiliser Double pour les calculs décimaux
    ' Dim total As Double
    ' Dim moyenne As Double
End Sub
```

---

## 4. Erreurs de compilation (Compile Errors)

### Qu'est-ce qu'une erreur de compilation ?

Les **erreurs de compilation** surviennent quand VBA tente de "préparer" votre code pour l'exécution. Elles sont détectées quand vous essayez d'exécuter le code ou quand vous compilez explicitement (Débogage > Compiler).

### Caractéristiques des erreurs de compilation

- **Détectées avant l'exécution complète** : VBA vérifie le code avant de le lancer
- **Empêchent l'exécution** : Le code ne peut pas démarrer
- **Souvent liées aux déclarations** : Variables, procédures, références

### Exemples courants d'erreurs de compilation

#### 1. Variable non déclarée (en mode Option Explicit)

```vba
Option Explicit  ' Force la déclaration de toutes les variables

Sub ExempleCompilation1()
    monVariable = 10  ' Erreur : Variable non déclarée
    Range("A1").Value = monVariable
End Sub

' CORRECT
Sub ExempleCompilation1Correct()
    Dim monVariable As Integer
    monVariable = 10
    Range("A1").Value = monVariable
End Sub
```

#### 2. Procédure non trouvée

```vba
Sub ExempleCompilation2()
    Call MaProcedureInexistante  ' Erreur : Procédure non trouvée
End Sub
```

#### 3. Référence d'objet manquante

```vba
' Si une référence à une bibliothèque est manquante
Sub ExempleCompilation3()
    Dim regex As RegExp  ' Erreur si la référence Microsoft VBScript est absente
End Sub
```

---

## 5. Comment identifier le type d'erreur

### Moment de détection

```
┌─────────────────────────────────────────────────────────────┐
│ QUAND L'ERREUR EST-ELLE DÉTECTÉE ?                          │
├─────────────────────────────────────────────────────────────┤
│ Pendant la frappe        → Erreur de syntaxe               │
│ Avant l'exécution        → Erreur de compilation           │
│ Pendant l'exécution      → Erreur d'exécution              │
│ Après l'exécution        → Erreur logique                  │
└─────────────────────────────────────────────────────────────┘
```

### Messages d'erreur typiques

#### Erreurs de syntaxe
- "Erreur de syntaxe"
- "Structure If sans End If correspondant"
- "Instruction incorrecte en dehors de Type"

#### Erreurs de compilation
- "Variable non définie"
- "Sub ou Function non définie"
- "Référence de projet non valide"

#### Erreurs d'exécution
- "Erreur d'exécution '9': L'indice n'appartient pas à la sélection"
- "Erreur d'exécution '11': Division par zéro"
- "Erreur d'exécution '13': Non-correspondance de type"

### Couleurs dans l'éditeur VBA

- **Rouge** : Erreurs de syntaxe détectées immédiatement
- **Surligné en jaune** : Ligne où s'est arrêtée l'exécution (erreur d'exécution)
- **Pas de couleur spéciale** : Erreurs logiques (difficiles à détecter)

---

## 6. Stratégies de prévention par type d'erreur

### Pour les erreurs de syntaxe

1. **Activez Option Explicit** : Écrivez `Option Explicit` en haut de vos modules
2. **Utilisez l'indentation** : Rendez la structure du code visible
3. **Écrivez petit à petit** : Testez fréquemment
4. **Utilisez l'auto-complétion** : Laissez VBA vous guider

```vba
' Bonne pratique : structure claire
Sub ExempleBonnePratique()
    Dim i As Integer

    For i = 1 To 10
        If Cells(i, 1).Value > 0 Then
            Cells(i, 2).Value = "Positif"
        Else
            Cells(i, 2).Value = "Négatif ou nul"
        End If
    Next i
End Sub
```

### Pour les erreurs d'exécution

1. **Vérifiez l'existence** avant d'utiliser
2. **Validez les données** avant les calculs
3. **Utilisez des gestionnaires d'erreur** : `On Error`

```vba
Sub PreventionErreurExecution()
    ' Vérifier l'existence d'une feuille
    Dim feuilleExiste As Boolean
    feuilleExiste = False

    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = "MaFeuille" Then
            feuilleExiste = True
            Exit For
        End If
    Next ws

    If feuilleExiste Then
        Worksheets("MaFeuille").Range("A1").Value = "OK"
    Else
        MsgBox "La feuille 'MaFeuille' n'existe pas"
    End If
End Sub
```

### Pour les erreurs logiques

1. **Testez avec des données connues** : Utilisez des exemples dont vous connaissez le résultat
2. **Utilisez Debug.Print** : Affichez les valeurs intermédiaires
3. **Décomposez les calculs complexes** : Étape par étape
4. **Relisez votre logique** : Expliquez votre code à voix haute

```vba
Sub PreventionErreurLogique()
    Dim prix As Double
    Dim remise As Double
    Dim taxe As Double
    Dim prixFinal As Double

    prix = 100
    remise = 0.1    ' 10%
    taxe = 0.05     ' 5%

    ' Calculer étape par étape pour éviter les erreurs logiques
    Dim prixApresRemise As Double
    prixApresRemise = prix * (1 - remise)
    Debug.Print "Prix après remise : " & prixApresRemise

    prixFinal = prixApresRemise * (1 + taxe)
    Debug.Print "Prix final : " & prixFinal

    Range("A1").Value = prixFinal
End Sub
```

---

## 7. Récapitulatif des types d'erreurs

### Tableau comparatif

| Type d'erreur | Quand détectée | Effet | Exemple typique |
|---------------|----------------|-------|-----------------|
| **Syntaxe** | Pendant la frappe | Code ne compile pas | Parenthèse manquante |
| **Compilation** | Avant exécution | Code ne démarre pas | Variable non déclarée |
| **Exécution** | Pendant exécution | Code s'arrête avec message | Division par zéro |
| **Logique** | Après vérification | Résultat incorrect | Condition inversée |

### Priorités de gestion

1. **Éliminez d'abord** les erreurs de syntaxe et compilation
2. **Prévenez** les erreurs d'exécution avec des vérifications
3. **Testez soigneusement** pour détecter les erreurs logiques
4. **Ajoutez des gestionnaires** pour les erreurs d'exécution imprévisibles

### Points clés à retenir

- **Les erreurs ne sont pas vos ennemies** : elles vous aident à améliorer votre code
- **Chaque type nécessite une approche différente** : prévention, gestion, ou test
- **La pratique améliore la détection** : plus vous programmez, plus vous anticipez
- **Un bon débogage commence par une bonne compréhension** des types d'erreurs

Dans la section suivante, nous apprendrons à utiliser `On Error Resume Next` pour gérer les erreurs d'exécution de manière contrôlée.

⏭️
