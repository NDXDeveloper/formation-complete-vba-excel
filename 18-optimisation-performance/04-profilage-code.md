🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.4 Profilage du code

## Introduction

Le profilage de code consiste à analyser les performances de votre programme pour identifier précisément où il passe le plus de temps. C'est comme faire un diagnostic médical de votre code : au lieu de deviner où sont les problèmes, vous mesurez objectivement pour savoir exactement quoi optimiser.

Sans profilage, vous risquez d'optimiser les mauvaises parties de votre code et de perdre du temps sur des améliorations qui n'apportent aucun gain visible.

## Pourquoi profiler votre code VBA ?

### Éviter les optimisations inutiles

```vba
Sub ExempleOptimisationInutile()
    ' Vous pourriez penser que cette boucle est le problème
    Dim i As Long
    For i = 1 To 1000
        Debug.Print i
    Next i

    ' Alors que le vrai problème est cette ligne qui prend 90% du temps
    Application.Wait Now + TimeValue("00:00:05")  ' Attente de 5 secondes !
End Sub
```

### Identifier les goulots d'étranglement réels

Le profilage vous permet de découvrir que parfois :
- 80% du temps est passé dans 20% du code
- Une fonction appelée une seule fois peut être plus problématique qu'une boucle de 1000 itérations
- Les opérations que vous pensiez rapides sont en fait très lentes

### Mesurer l'impact des optimisations

Sans mesures, vous ne savez pas si vos optimisations sont efficaces :

```vba
Sub MesurerImpactOptimisation()
    Dim avant As Double, apres As Double

    ' Mesure AVANT optimisation
    avant = Timer
    ' Votre code original ici
    Debug.Print "Avant optimisation : " & Format(Timer - avant, "0.00") & "s"

    ' Mesure APRÈS optimisation
    apres = Timer
    ' Votre code optimisé ici
    Debug.Print "Après optimisation : " & Format(Timer - apres, "0.00") & "s"
    Debug.Print "Gain : " & Format(avant / apres, "0.0") & "x plus rapide"
End Sub
```

## Techniques de base pour profiler

### 1. Chronométrage simple avec Timer

La fonction `Timer` est votre outil de base pour mesurer les performances :

```vba
Function ChronometrerSection() As Double
    Dim debut As Double
    debut = Timer

    ' Section de code à mesurer
    Dim i As Long
    For i = 1 To 100000
        ' Simulation d'un traitement
        Cells(1, 1).Value = i
    Next i

    ChronometrerSection = Timer - debut
End Function
```

### 2. Profilage par sections

Divisez votre code en sections et mesurez chacune :

```vba
Sub ProfilageParSections()
    Dim tempsTotal As Double, tempsSection As Double
    Dim debut As Double

    debut = Timer
    Debug.Print "=== DÉBUT DU PROFILAGE ==="

    ' Section 1 : Initialisation
    tempsSection = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Debug.Print "Section 1 (Init) : " & Format(Timer - tempsSection, "0.000") & "s"

    ' Section 2 : Lecture des données
    tempsSection = Timer
    Dim donnees As Variant
    donnees = Range("A1:E10000").Value
    Debug.Print "Section 2 (Lecture) : " & Format(Timer - tempsSection, "0.000") & "s"

    ' Section 3 : Traitement
    tempsSection = Timer
    Dim i As Long
    For i = 1 To UBound(donnees, 1)
        donnees(i, 1) = donnees(i, 1) * 2
    Next i
    Debug.Print "Section 3 (Traitement) : " & Format(Timer - tempsSection, "0.000") & "s"

    ' Section 4 : Écriture
    tempsSection = Timer
    Range("F1:J10000").Value = donnees
    Debug.Print "Section 4 (Écriture) : " & Format(Timer - tempsSection, "0.000") & "s"

    ' Section 5 : Finalisation
    tempsSection = Timer
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Debug.Print "Section 5 (Final) : " & Format(Timer - tempsSection, "0.000") & "s"

    tempsTotal = Timer - debut
    Debug.Print "=== TEMPS TOTAL : " & Format(tempsTotal, "0.000") & "s ==="
End Sub
```

### 3. Profilage de fonctions individuelles

```vba
Function ProfilerFonction(nomFonction As String, Optional parametres As String = "") As Double
    Dim debut As Double
    debut = Timer

    Debug.Print "Début " & nomFonction & "(" & parametres & ") - " & Format(Now, "hh:nn:ss.000")

    ' La fonction retourne son temps d'exécution
    ProfilerFonction = Timer - debut

    Debug.Print "Fin " & nomFonction & " - Durée : " & Format(ProfilerFonction, "0.000") & "s"
End Function

Sub ExempleUtilisationProfiler()
    Dim tempsFonction As Double

    tempsFonction = ProfilerFonction("TraitementDonnees", "1000 lignes")
    ' Votre fonction TraitementDonnees ici
    tempsFonction = Timer - tempsFonction  ' Calculer le temps réel

    If tempsFonction > 1 Then
        Debug.Print "ATTENTION : " & "TraitementDonnees" & " est lente !"
    End If
End Sub
```

## Techniques avancées de profilage

### 1. Profilage automatique avec une classe

```vba
' Module de classe : ClasseProfiler
Private nomOperation As String
Private tempsDebut As Double

Property Let Operation(nom As String)
    nomOperation = nom
    tempsDebut = Timer
    Debug.Print ">>> DÉBUT : " & nomOperation & " - " & Format(Now, "hh:nn:ss.000")
End Property

Property Get TempsEcoule() As Double
    TempsEcoule = Timer - tempsDebut
End Property

Sub Terminer()
    Dim duree As Double
    duree = Timer - tempsDebut
    Debug.Print "<<< FIN : " & nomOperation & " - Durée : " & Format(duree, "0.000") & "s"

    If duree > 0.5 Then
        Debug.Print "⚠️  PERFORMANCE : " & nomOperation & " prend du temps !"
    End If
End Sub
```

Usage de la classe :

```vba
Sub UtiliserClasseProfiler()
    Dim profiler As ClasseProfiler
    Set profiler = New ClasseProfiler

    ' Profiler la lecture
    profiler.Operation = "Lecture fichier Excel"
    Dim donnees As Variant
    donnees = Range("A1:Z10000").Value
    profiler.Terminer

    ' Profiler le traitement
    profiler.Operation = "Traitement des données"
    ' Votre traitement ici
    profiler.Terminer

    Set profiler = Nothing
End Sub
```

### 2. Compteur d'appels de fonctions

```vba
' Variables globales pour le comptage
Public compteurAppels As Collection
Public tempsTotal As Collection

Sub InitialiserCompteurs()
    Set compteurAppels = New Collection
    Set tempsTotal = New Collection
End Sub

Sub EnregistrerAppel(nomFonction As String, duree As Double)
    Dim nbAppels As Long
    Dim tempsCumule As Double

    ' Compter les appels
    On Error Resume Next
    nbAppels = compteurAppels(nomFonction)
    tempsCumule = tempsTotal(nomFonction)
    On Error GoTo 0

    ' Incrémenter
    nbAppels = nbAppels + 1
    tempsCumule = tempsCumule + duree

    ' Sauvegarder
    compteurAppels.Remove nomFonction
    tempsTotal.Remove nomFonction
    compteurAppels.Add nbAppels, nomFonction
    tempsTotal.Add tempsCumule, nomFonction
End Sub

Sub AfficherStatistiques()
    Dim i As Long
    Dim nomFonction As String

    Debug.Print "=== STATISTIQUES D'APPELS ==="
    For i = 1 To compteurAppels.Count
        nomFonction = compteurAppels.Keys()(i)  ' Si disponible
        Debug.Print nomFonction & " : " & compteurAppels(nomFonction) & " appels, " & _
                   Format(tempsTotal(nomFonction), "0.000") & "s total"
    Next i
End Sub
```

### 3. Profilage de boucles avec échantillonnage

```vba
Sub ProfilerBoucle()
    Dim i As Long
    Dim tempsDebut As Double, tempsEchantillon As Double
    Dim nbEchantillons As Long

    nbEchantillons = 100  ' Mesurer tous les 100 itérations
    tempsDebut = Timer

    For i = 1 To 10000
        ' Votre traitement ici
        Cells(i, 1).Value = i * 2

        ' Échantillonnage
        If i Mod nbEchantillons = 0 Then
            tempsEchantillon = Timer - tempsDebut
            Debug.Print "Itération " & i & " : " & _
                       Format(tempsEchantillon / i * 1000, "0.0") & "ms par itération"
        End If
    Next i

    Debug.Print "Temps moyen final : " & _
               Format((Timer - tempsDebut) / 10000 * 1000, "0.0") & "ms par itération"
End Sub
```

## Analyse des résultats de profilage

### 1. Identifier les goulots d'étranglement

```vba
Sub AnalyserResultats()
    ' Exemple de résultats de profilage
    Debug.Print "=== ANALYSE DES PERFORMANCES ==="
    Debug.Print "Section 1 (Init) : 0.001s (0.1%)"
    Debug.Print "Section 2 (Lecture) : 0.050s (5%)"
    Debug.Print "Section 3 (Traitement) : 0.850s (85%) ← GOULOT !"
    Debug.Print "Section 4 (Écriture) : 0.080s (8%)"
    Debug.Print "Section 5 (Final) : 0.019s (1.9%)"
    Debug.Print "TOTAL : 1.000s"
    Debug.Print ""
    Debug.Print "💡 RECOMMANDATION : Optimiser la Section 3"
End Sub
```

### 2. Calculer les pourcentages

```vba
Function CalculerPourcentages(temps() As Double) As String()
    Dim total As Double
    Dim pourcentages() As String
    Dim i As Long

    ' Calculer le total
    For i = 0 To UBound(temps)
        total = total + temps(i)
    Next i

    ' Calculer les pourcentages
    ReDim pourcentages(UBound(temps))
    For i = 0 To UBound(temps)
        pourcentages(i) = Format(temps(i) / total * 100, "0.0") & "%"
    Next i

    CalculerPourcentages = pourcentages
End Function
```

### 3. Générer un rapport de performance

```vba
Sub GenererRapportPerformance()
    Dim rapport As String
    Dim fso As Object
    Dim fichier As Object

    ' Créer le rapport
    rapport = "RAPPORT DE PERFORMANCE - " & Format(Now, "dd/mm/yyyy hh:nn") & vbCrLf
    rapport = rapport & String(50, "=") & vbCrLf & vbCrLf

    rapport = rapport & "SECTIONS ANALYSÉES :" & vbCrLf
    rapport = rapport & "- Initialisation : 0.001s (0.1%)" & vbCrLf
    rapport = rapport & "- Lecture données : 0.050s (5.0%)" & vbCrLf
    rapport = rapport & "- Traitement : 0.850s (85.0%) ← CRITIQUE" & vbCrLf
    rapport = rapport & "- Écriture : 0.080s (8.0%)" & vbCrLf
    rapport = rapport & "- Finalisation : 0.019s (1.9%)" & vbCrLf & vbCrLf

    rapport = rapport & "RECOMMANDATIONS :" & vbCrLf
    rapport = rapport & "1. Optimiser la section Traitement (85% du temps)" & vbCrLf
    rapport = rapport & "2. Utiliser des tableaux au lieu de cellules individuelles" & vbCrLf
    rapport = rapport & "3. Réduire les interactions avec Excel" & vbCrLf

    ' Sauvegarder le rapport
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fichier = fso.CreateTextFile(ThisWorkbook.Path & "\RapportPerformance.txt", True)
    fichier.Write rapport
    fichier.Close

    Debug.Print "Rapport sauvegardé : " & ThisWorkbook.Path & "\RapportPerformance.txt"
End Sub
```

## Outils de mesure spécialisés

### 1. Mesure de la consommation mémoire

```vba
Sub MesurerMemoire()
    ' Estimation simple de l'utilisation mémoire
    Dim avant As Long, apres As Long

    ' Avant traitement
    DoEvents  ' Forcer le nettoyage
    avant = 0  ' Placeholder - nécessiterait des API Windows

    ' Votre code ici
    Dim grosTableau(1 To 100000) As Double
    Dim i As Long
    For i = 1 To 100000
        grosTableau(i) = i
    Next i

    ' Après traitement
    DoEvents
    apres = 0  ' Placeholder

    Debug.Print "Consommation mémoire estimée : " & (apres - avant) & " octets"

    ' Libérer
    Erase grosTableau
End Sub
```

### 2. Mesure de la fréquence d'appel

```vba
' Module global pour surveiller les appels
Public appelsFonctions As Object

Sub InitialiserSurveillance()
    Set appelsFonctions = CreateObject("Scripting.Dictionary")
End Sub

Sub SurveillerAppel(nomFonction As String)
    If appelsFonctions.Exists(nomFonction) Then
        appelsFonctions(nomFonction) = appelsFonctions(nomFonction) + 1
    Else
        appelsFonctions.Add nomFonction, 1
    End If
End Sub

Sub AfficherAppels()
    Dim cle As Variant
    Debug.Print "=== FRÉQUENCE D'APPELS ==="
    For Each cle In appelsFonctions.Keys
        Debug.Print cle & " : " & appelsFonctions(cle) & " appels"
    Next cle
End Sub
```

### 3. Profilage conditionnel

```vba
Public Const MODE_DEBUG As Boolean = True  ' Activer/désactiver le profilage

Sub ProfilerSiDebug(nomSection As String)
    Static tempsDebut As Double
    Static sectionActuelle As String

    If Not MODE_DEBUG Then Exit Sub

    If sectionActuelle <> "" Then
        ' Terminer la section précédente
        Debug.Print "FIN " & sectionActuelle & " : " & Format(Timer - tempsDebut, "0.000") & "s"
    End If

    ' Commencer nouvelle section
    sectionActuelle = nomSection
    tempsDebut = Timer
    Debug.Print "DÉBUT " & nomSection & " : " & Format(Now, "hh:nn:ss.000")
End Sub
```

## Bonnes pratiques de profilage

### 1. Quoi mesurer et quand

```vba
Sub BonnesPratiquesMesure()
    ' ✓ MESURER : Les boucles importantes
    Dim tempsDebut As Double
    tempsDebut = Timer

    For i = 1 To 100000
        ' Traitement important
    Next i

    If Timer - tempsDebut > 0.1 Then  ' Seulement si significatif
        Debug.Print "Boucle importante : " & Format(Timer - tempsDebut, "0.000") & "s"
    End If

    ' ✓ MESURER : Les opérations externes (fichiers, base de données)
    tempsDebut = Timer
    Workbooks.Open "GrosFichier.xlsx"
    Debug.Print "Ouverture fichier : " & Format(Timer - tempsDebut, "0.000") & "s"

    ' ✗ NE PAS MESURER : Les opérations triviales
    ' Debug.Print Timer  ' Inutile pour des micro-opérations
End Sub
```

### 2. Minimiser l'impact du profilage

```vba
' Fonction optimisée pour le profilage
Sub ProfilerOptimise(nomSection As String, Optional terminer As Boolean = False)
    Static sections As Object
    Static temps As Object

    If sections Is Nothing Then
        Set sections = CreateObject("Scripting.Dictionary")
        Set temps = CreateObject("Scripting.Dictionary")
    End If

    If terminer Then
        If temps.Exists(nomSection) Then
            Debug.Print nomSection & " : " & Format(Timer - temps(nomSection), "0.000") & "s"
            temps.Remove nomSection
        End If
    Else
        temps(nomSection) = Timer
    End If
End Sub
```

### 3. Profilage en production vs développement

```vba
#If DEBUG_MODE Then
    ' Code de profilage détaillé en développement
    Public Const PROFILING_ACTIF As Boolean = True
#Else
    ' Profilage minimal en production
    Public Const PROFILING_ACTIF As Boolean = False
#End If

Sub ProfilageConditionnel()
    If PROFILING_ACTIF Then
        Debug.Print "Début traitement - " & Now
    End If

    ' Votre code principal

    If PROFILING_ACTIF Then
        Debug.Print "Fin traitement - " & Now
    End If
End Sub
```

## Interpréter les résultats

### 1. Seuils de performance acceptables

```vba
Sub EvaluerPerformances(duree As Double, nomOperation As String)
    Select Case duree
        Case Is < 0.1
            Debug.Print "✅ " & nomOperation & " : EXCELLENT (" & Format(duree, "0.000") & "s)"
        Case 0.1 To 0.5
            Debug.Print "👍 " & nomOperation & " : BON (" & Format(duree, "0.000") & "s)"
        Case 0.5 To 2
            Debug.Print "⚠️ " & nomOperation & " : MOYEN (" & Format(duree, "0.000") & "s)"
        Case Is > 2
            Debug.Print "🔴 " & nomOperation & " : LENT (" & Format(duree, "0.000") & "s) - À OPTIMISER"
    End Select
End Sub
```

### 2. Comparaison avant/après optimisation

```vba
Sub ComparerPerformances()
    Dim avant As Double, apres As Double, gain As Double

    ' Mesure AVANT
    avant = ChronometrerCodeOriginal()

    ' Mesure APRÈS
    apres = ChronometrerCodeOptimise()

    ' Analyse du gain
    gain = avant / apres

    Debug.Print "=== COMPARAISON PERFORMANCES ==="
    Debug.Print "Avant : " & Format(avant, "0.000") & "s"
    Debug.Print "Après : " & Format(apres, "0.000") & "s"

    If gain > 1.2 Then
        Debug.Print "🎉 OPTIMISATION RÉUSSIE : " & Format(gain, "0.0") & "x plus rapide"
    ElseIf gain > 1.05 Then
        Debug.Print "👍 Amélioration légère : +" & Format((gain - 1) * 100, "0") & "%"
    Else
        Debug.Print "❌ Pas d'amélioration significative"
    End If
End Sub
```

## Template de profilage réutilisable

```vba
Sub TemplateProfilageComplet()
    Dim tempsTotal As Double, tempsSection As Double
    Dim debut As Double

    debut = Timer
    Debug.Print "================================"
    Debug.Print "DÉBUT PROFILAGE : " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Debug.Print "================================"

    On Error GoTo GestionErreur

    ' SECTION 1
    tempsSection = Timer
    Debug.Print "Section 1 - Initialisation..."
    ' Votre code Section 1
    Debug.Print "  ✓ Terminé en " & Format(Timer - tempsSection, "0.000") & "s"

    ' SECTION 2
    tempsSection = Timer
    Debug.Print "Section 2 - Traitement principal..."
    ' Votre code Section 2
    Debug.Print "  ✓ Terminé en " & Format(Timer - tempsSection, "0.000") & "s"

    ' SECTION 3
    tempsSection = Timer
    Debug.Print "Section 3 - Finalisation..."
    ' Votre code Section 3
    Debug.Print "  ✓ Terminé en " & Format(Timer - tempsSection, "0.000") & "s"

    ' BILAN FINAL
    tempsTotal = Timer - debut
    Debug.Print "================================"
    Debug.Print "TEMPS TOTAL : " & Format(tempsTotal, "0.000") & "s"

    If tempsTotal > 5 Then
        Debug.Print "⚠️  PERFORMANCE : Code lent, optimisation recommandée"
    ElseIf tempsTotal < 0.1 Then
        Debug.Print "🚀 PERFORMANCE : Code très rapide !"
    Else
        Debug.Print "✅ PERFORMANCE : Code correcte"
    End If

    Debug.Print "FIN PROFILAGE : " & Format(Now, "hh:nn:ss")
    Debug.Print "================================"
    Exit Sub

GestionErreur:
    Debug.Print "❌ ERREUR pendant le profilage : " & Err.Description
    Debug.Print "Temps écoulé avant erreur : " & Format(Timer - debut, "0.000") & "s"
End Sub
```

## Résumé

Le profilage est essentiel pour optimiser efficacement votre code VBA. Les points clés à retenir :

1. **Mesurez avant d'optimiser** - Ne devinez jamais, mesurez toujours
2. **Identifiez les vrais goulots** - 80% des problèmes viennent de 20% du code
3. **Utilisez Timer pour des mesures simples** - Outil de base suffisant pour la plupart des cas
4. **Profiler par sections** - Divisez votre code en parties mesurables
5. **Analysez les résultats** - Concentrez-vous sur les sections qui prennent le plus de temps
6. **Validez vos optimisations** - Mesurez l'impact réel de vos améliorations

Le profilage vous permet de passer d'optimisations "à l'aveugle" à des améliorations ciblées et efficaces, garantissant que votre temps d'optimisation soit bien investi.

⏭️
