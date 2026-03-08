🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.1 Désactivation des calculs et de l'affichage

## Introduction

Lorsque vous exécutez une macro VBA qui modifie de nombreuses cellules, Excel effectue automatiquement plusieurs tâches en arrière-plan qui ralentissent considérablement l'exécution :
- Il met à jour l'affichage à chaque modification
- Il recalcule toutes les formules après chaque changement
- Il actualise l'écran en permanence

En désactivant temporairement ces fonctionnalités pendant l'exécution de votre macro, vous pouvez obtenir des gains de performance spectaculaires, souvent de l'ordre de 10 à 100 fois plus rapide !

## Désactivation de la mise à jour de l'écran

### Principe de base

Par défaut, Excel redessine l'écran à chaque modification pour que l'utilisateur voit les changements en temps réel. Pendant l'exécution d'une macro, cette mise à jour constante est inutile et très coûteuse en performance.

### La propriété ScreenUpdating

```vba
Sub ExempleScreenUpdating()
    ' Désactiver la mise à jour de l'écran
    Application.ScreenUpdating = False

    ' Votre code qui modifie des cellules
    Dim i As Integer
    For i = 1 To 1000
        Cells(i, 1).Value = "Ligne " & i
    Next i

    ' Réactiver la mise à jour de l'écran
    Application.ScreenUpdating = True
End Sub
```

### Points importants à retenir

**Toujours réactiver ScreenUpdating :** Il est crucial de remettre `ScreenUpdating = True` à la fin de votre macro. Si vous oubliez, Excel restera "figé" et l'utilisateur ne verra plus les modifications.

**Gestion des erreurs :** Utilisez toujours une gestion d'erreur pour vous assurer que `ScreenUpdating` soit réactivé même si votre macro rencontre un problème :

```vba
Sub MacroSecurisee()
    On Error GoTo GestionErreur

    Application.ScreenUpdating = False

    ' Votre code ici
    Dim i As Integer
    For i = 1 To 1000
        Cells(i, 1).Value = i * 2
    Next i

    Application.ScreenUpdating = True
    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur dans la macro : " & Err.Description
End Sub
```

## Désactivation du recalcul automatique

### Comprendre le recalcul automatique

Excel recalcule automatiquement toutes les formules du classeur chaque fois qu'une cellule est modifiée. Si votre feuille contient de nombreuses formules complexes, ce recalcul constant peut considérablement ralentir votre macro.

### La propriété Calculation

```vba
Sub ExempleCalculation()
    ' Sauvegarder le mode de calcul actuel
    Dim modeCalculOriginal As XlCalculation
    modeCalculOriginal = Application.Calculation

    ' Passer en mode manuel
    Application.Calculation = xlCalculationManual

    ' Votre code qui modifie des données
    Dim i As Integer
    For i = 1 To 500
        Cells(i, 1).Value = i
        Cells(i, 2).Formula = "=A" & i & "*2"
    Next i

    ' Restaurer le mode de calcul original
    Application.Calculation = modeCalculOriginal
End Sub
```

### Les différents modes de calcul

**xlCalculationAutomatic :** Mode par défaut, Excel recalcule automatiquement à chaque modification.

**xlCalculationManual :** Les formules ne sont recalculées que lorsque l'utilisateur appuie sur F9 ou que le code appelle explicitement le recalcul.

**xlCalculationSemiautomatic :** Excel recalcule automatiquement, sauf les tableaux de données.

### Forcer un recalcul quand nécessaire

Si vous avez besoin que certaines formules soient recalculées pendant votre macro :

```vba
Sub RecalculPartiel()
    Application.Calculation = xlCalculationManual

    ' Modifier des données
    Range("A1:A100").Value = 1

    ' Recalculer seulement une plage spécifique
    Range("B1:B100").Calculate

    ' Ou recalculer toute la feuille
    ActiveSheet.Calculate

    ' Ou recalculer tout le classeur
    Application.Calculate

    Application.Calculation = xlCalculationAutomatic
End Sub
```

## Désactivation des événements

### Pourquoi désactiver les événements

Les événements Excel (comme Worksheet_Change, Workbook_SheetChange) se déclenchent automatiquement lors des modifications. Pendant l'exécution d'une macro, ces événements peuvent :
- Ralentir l'exécution
- Créer des boucles infinies
- Déclencher des actions non désirées

### La propriété EnableEvents

```vba
Sub ExempleEnableEvents()
    ' Désactiver les événements
    Application.EnableEvents = False

    ' Modifier des cellules sans déclencher d'événements
    Range("A1:A1000").Value = "Données modifiées"

    ' Réactiver les événements
    Application.EnableEvents = True
End Sub
```

## Combinaison des optimisations

### Le modèle standard d'optimisation

Voici un modèle que vous pouvez réutiliser dans toutes vos macros nécessitant une optimisation :

```vba
Sub MacroOptimisee()
    ' Sauvegarder les paramètres actuels
    Dim screenUpdate As Boolean
    Dim calculationMode As XlCalculation
    Dim enableEvents As Boolean

    screenUpdate = Application.ScreenUpdating
    calculationMode = Application.Calculation
    enableEvents = Application.EnableEvents

    ' Gestion d'erreur pour restaurer les paramètres
    On Error GoTo RestaurerParametres

    ' Optimisations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ========================================
    ' VOTRE CODE PRINCIPAL ICI
    ' ========================================

    Dim i As Long
    For i = 1 To 10000
        Cells(i, 1).Value = "Données " & i
        Cells(i, 2).Value = i * 2
    Next i

    ' Restaurer les paramètres
RestaurerParametres:
    Application.ScreenUpdating = screenUpdate
    Application.Calculation = calculationMode
    Application.EnableEvents = enableEvents

    ' Afficher un message d'erreur si nécessaire
    If Err.Number <> 0 Then
        MsgBox "Erreur : " & Err.Description
    End If
End Sub
```

### Fonction utilitaire réutilisable

Vous pouvez créer des fonctions utilitaires pour simplifier l'activation/désactivation :

```vba
Sub ActiverOptimisations()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Sub DesactiverOptimisations()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub MonTraitement()
    On Error GoTo Nettoyage

    ActiverOptimisations

    ' Votre code ici
    ' ...

    DesactiverOptimisations
    Exit Sub

Nettoyage:
    DesactiverOptimisations
    MsgBox "Erreur : " & Err.Description
End Sub
```

## Autres optimisations d'affichage

### Désactivation des alertes

Pendant l'exécution de macros, Excel peut afficher des boîtes de dialogue d'avertissement (suppression de feuilles, remplacement de données, etc.). Vous pouvez les désactiver temporairement :

```vba
Sub SansAlertes()
    Application.DisplayAlerts = False

    ' Code qui pourrait générer des alertes
    Worksheets("Feuille1").Delete  ' Pas d'alerte de confirmation

    Application.DisplayAlerts = True
End Sub
```

### Masquage des barres d'état

Si vous effectuez de nombreuses opérations, masquer temporairement la barre d'état peut apporter un petit gain :

```vba
Sub SansBarreEtat()
    Application.DisplayStatusBar = False

    ' Votre code

    Application.DisplayStatusBar = True
End Sub
```

## Mesurer l'impact des optimisations

### Comparaison avant/après

```vba
Sub TestPerformance()
    Dim tempsDebut As Double, tempsFin As Double

    ' Test SANS optimisations
    tempsDebut = Timer

    Dim i As Long
    For i = 1 To 5000
        Cells(i, 1).Value = i
    Next i

    tempsFin = Timer
    Debug.Print "Sans optimisation : " & Format(tempsFin - tempsDebut, "0.00") & " secondes"

    ' Effacer les données
    Range("A:A").Clear

    ' Test AVEC optimisations
    tempsDebut = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = 1 To 5000
        Cells(i, 1).Value = i
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    tempsFin = Timer
    Debug.Print "Avec optimisation : " & Format(tempsFin - tempsDebut, "0.00") & " secondes"
End Sub
```

## Recommandations importantes

### Quand utiliser ces optimisations

- Macros modifiant plus de quelques centaines de cellules
- Boucles avec de nombreuses itérations
- Opérations sur des plages importantes
- Import/export de gros volumes de données

### Quand NE PAS les utiliser

- Macros très courtes (quelques cellules modifiées)
- Code nécessitant un feedback visuel en temps réel
- Macros interactives où l'utilisateur doit voir les changements

### Bonnes pratiques de sécurité

1. **Toujours utiliser une gestion d'erreur** pour restaurer les paramètres
2. **Sauvegarder les paramètres originaux** avant de les modifier
3. **Tester rigoureusement** vos macros avec les optimisations
4. **Documenter** dans vos commentaires que les optimisations sont activées

## Résumé

La désactivation temporaire des calculs et de l'affichage est l'une des techniques d'optimisation les plus simples et efficaces en VBA. En appliquant ces trois paramètres de base :

```vba
Application.ScreenUpdating = False  
Application.Calculation = xlCalculationManual  
Application.EnableEvents = False  
```

Vous pouvez obtenir des gains de performance considérables avec un effort minimal. N'oubliez jamais de les restaurer en fin de macro, et utilisez toujours une gestion d'erreur appropriée pour garantir la stabilité de votre application.

⏭️
