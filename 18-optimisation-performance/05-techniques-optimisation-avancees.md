🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.5 Techniques d'optimisation avancées

## Introduction

Cette section présente des techniques d'optimisation plus sophistiquées qui peuvent transformer radicalement les performances de vos applications VBA. Ces méthodes vont au-delà des optimisations de base et exploitent des fonctionnalités avancées d'Excel et de Windows pour obtenir des gains de performance exceptionnels.

Bien que ces techniques soient "avancées", elles restent accessibles avec les bonnes explications et peuvent faire la différence entre une application VBA acceptable et une solution vraiment professionnelle.

## 1. Optimisation par manipulation directe de la mémoire

### Utilisation des API Windows pour les performances

Certaines opérations peuvent être accélérées en utilisant directement les fonctions du système Windows :

```vba
' Déclarations d'API pour l'optimisation
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub ExempleCopieMemoireRapide()
    ' Copie ultra-rapide de gros tableaux
    Dim source(1 To 100000) As Long
    Dim destination(1 To 100000) As Long
    Dim i As Long

    ' Remplir le tableau source
    For i = 1 To 100000
        source(i) = i
    Next i

    ' Méthode traditionnelle (lente)
    Dim tempsDebut As Double
    tempsDebut = Timer
    For i = 1 To 100000
        destination(i) = source(i)
    Next i
    Debug.Print "Copie traditionnelle : " & Format(Timer - tempsDebut, "0.000") & "s"

    ' Méthode optimisée avec API (très rapide)
    tempsDebut = Timer
    CopyMemory destination(1), source(1), 100000 * 4  ' 4 octets par Long
    Debug.Print "Copie avec API : " & Format(Timer - tempsDebut, "0.000") & "s"
End Sub
```

### Pause optimisée avec Sleep

```vba
Sub PauseOptimisee()
    ' Au lieu d'utiliser Application.Wait qui bloque Excel
    ' ÉVITER : Application.Wait Now + TimeValue("00:00:01")

    ' PRÉFÉRER : Sleep qui libère le processeur
    Sleep 1000  ' Pause de 1 seconde sans bloquer Excel
    DoEvents    ' Permettre à Excel de traiter les événements
End Sub
```

## 2. Optimisation des calculs avec WorksheetFunction

### Exploiter la puissance des fonctions Excel natives

Excel a des fonctions intégrées ultra-optimisées. Il est souvent plus rapide de les utiliser que de recréer la logique en VBA :

```vba
Sub ComparaisonCalculs()
    Dim donnees(1 To 10000) As Double
    Dim i As Long, somme As Double
    Dim tempsDebut As Double

    ' Remplir les données de test
    For i = 1 To 10000
        donnees(i) = i * 1.5
    Next i

    ' Méthode 1 : Boucle VBA (lente)
    tempsDebut = Timer
    somme = 0
    For i = 1 To 10000
        somme = somme + donnees(i)
    Next i
    Debug.Print "Somme VBA : " & Format(Timer - tempsDebut, "0.000") & "s, Résultat : " & somme

    ' Méthode 2 : Fonction Excel native (rapide)
    tempsDebut = Timer
    ' Écrire les données dans une plage temporaire
    Range("Z1:Z10000").Value = Application.Transpose(donnees)
    ' Utiliser la fonction native
    somme = Application.WorksheetFunction.Sum(Range("Z1:Z10000"))
    Debug.Print "Somme Excel : " & Format(Timer - tempsDebut, "0.000") & "s, Résultat : " & somme

    ' Nettoyer
    Range("Z1:Z10000").Clear
End Sub
```

### Fonctions statistiques avancées

```vba
Sub FonctionsStatistiquesOptimisees()
    Dim plage As Range
    Set plage = Range("A1:A10000")

    ' Au lieu de calculer manuellement
    Dim moyenne As Double, ecartType As Double, median As Double

    ' Utiliser les fonctions Excel optimisées
    moyenne = Application.WorksheetFunction.Average(plage)
    ecartType = Application.WorksheetFunction.StDev(plage)
    median = Application.WorksheetFunction.Median(plage)

    Debug.Print "Moyenne : " & moyenne
    Debug.Print "Écart-type : " & ecartType
    Debug.Print "Médiane : " & median
End Sub
```

## 3. Optimisation des accès aux plages avec Union et Intersect

### Opérations sur plusieurs plages simultanément

```vba
Sub OptimisationPlagesMultiples()
    Dim plage1 As Range, plage2 As Range, plage3 As Range
    Dim plageUnion As Range

    Set plage1 = Range("A1:A1000")
    Set plage2 = Range("C1:C1000")
    Set plage3 = Range("E1:E1000")

    ' Méthode inefficace : traiter chaque plage séparément
    Dim tempsDebut As Double
    tempsDebut = Timer
    plage1.Font.Bold = True
    plage2.Font.Bold = True
    plage3.Font.Bold = True
    Debug.Print "Méthode séparée : " & Format(Timer - tempsDebut, "0.000") & "s"

    ' Méthode optimisée : Union pour traiter en une fois
    tempsDebut = Timer
    Set plageUnion = Union(plage1, plage2, plage3)
    plageUnion.Font.Bold = False  ' Reset
    Debug.Print "Méthode Union : " & Format(Timer - tempsDebut, "0.000") & "s"
End Sub
```

### Intersection pour des opérations conditionnelles

```vba
Sub OptimisationAvecIntersect()
    Dim plageSource As Range, plageFiltre As Range, plageResultat As Range

    Set plageSource = Range("A1:A10000")
    Set plageFiltre = Range("A1:A5000")  ' Seulement la première moitié

    ' Utiliser Intersect pour traiter seulement la partie commune
    Set plageResultat = Intersect(plageSource, plageFiltre)

    If Not plageResultat Is Nothing Then
        plageResultat.Interior.Color = RGB(255, 255, 0)  ' Surligner en jaune
        Debug.Print "Plage traitée : " & plageResultat.Address
    End If
End Sub
```

## 4. Optimisation avec les collections et dictionnaires

### Dictionnaire pour les recherches ultra-rapides

```vba
Sub RechercheAvecDictionnaire()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, tempsDebut As Double

    ' Remplir le dictionnaire (une seule fois)
    tempsDebut = Timer
    For i = 1 To 100000
        dict.Add "Clé" & i, "Valeur" & i
    Next i
    Debug.Print "Remplissage dictionnaire : " & Format(Timer - tempsDebut, "0.000") & "s"

    ' Recherches ultra-rapides
    tempsDebut = Timer
    For i = 1 To 10000
        If dict.Exists("Clé" & (i * 5)) Then
            ' Traitement si trouvé
        End If
    Next i
    Debug.Print "10000 recherches : " & Format(Timer - tempsDebut, "0.000") & "s"

    ' Comparaison avec recherche linéaire dans un tableau
    Dim tableau(1 To 100000) As String
    For i = 1 To 100000
        tableau(i) = "Clé" & i
    Next i

    tempsDebut = Timer
    For i = 1 To 1000  ' Seulement 1000 pour éviter l'attente
        Dim j As Long, trouve As Boolean
        trouve = False
        For j = 1 To 100000
            If tableau(j) = "Clé" & (i * 5) Then
                trouve = True
                Exit For
            End If
        Next j
    Next i
    Debug.Print "1000 recherches linéaires : " & Format(Timer - tempsDebut, "0.000") & "s"
End Sub
```

### Collection pour maintenir l'ordre avec performance

```vba
Sub CollectionOptimisee()
    Dim col As Collection
    Set col = New Collection

    Dim i As Long, tempsDebut As Double

    ' Ajout optimisé avec clé
    tempsDebut = Timer
    For i = 1 To 10000
        col.Add "Valeur" & i, "Clé" & i
    Next i
    Debug.Print "Ajout collection : " & Format(Timer - tempsDebut, "0.000") & "s"

    ' Accès direct par clé (rapide)
    tempsDebut = Timer
    For i = 1 To 1000
        Dim valeur As String
        valeur = col("Clé" & (i * 5))
    Next i
    Debug.Print "Accès par clé : " & Format(Timer - tempsDebut, "0.000") & "s"
End Sub
```

## 5. Optimisation des E/S (Entrées/Sorties)

### Lecture/Écriture de fichiers optimisée

```vba
Sub LectureOptimiseeFichier()
    Dim nomFichier As String
    Dim contenu As String
    Dim tempsDebut As Double

    nomFichier = ThisWorkbook.Path & "\test.txt"

    ' Méthode optimisée : lire tout le fichier en une fois
    tempsDebut = Timer

    Dim numeroFichier As Integer
    numeroFichier = FreeFile

    Open nomFichier For Binary As #numeroFichier
    contenu = Space$(LOF(numeroFichier))  ' Allouer toute la mémoire nécessaire
    Get #numeroFichier, , contenu         ' Lire tout d'un coup
    Close #numeroFichier

    Debug.Print "Lecture optimisée : " & Format(Timer - tempsDebut, "0.000") & "s"
    Debug.Print "Taille lue : " & Len(contenu) & " caractères"
End Sub

Sub EcritureOptimiseeFichier()
    Dim nomFichier As String
    Dim contenu As String
    Dim i As Long, tempsDebut As Double

    nomFichier = ThisWorkbook.Path & "\test_sortie.txt"

    ' Construire le contenu en mémoire d'abord
    tempsDebut = Timer
    Dim lignes() As String
    ReDim lignes(1 To 10000)

    For i = 1 To 10000
        lignes(i) = "Ligne " & i & " avec des données importantes"
    Next i

    contenu = Join(lignes, vbCrLf)

    ' Écrire tout en une fois
    Dim numeroFichier As Integer
    numeroFichier = FreeFile

    Open nomFichier For Output As #numeroFichier
    Print #numeroFichier, contenu
    Close #numeroFichier

    Debug.Print "Écriture optimisée : " & Format(Timer - tempsDebut, "0.000") & "s"
End Sub
```

### Optimisation des imports/exports Excel

```vba
Sub ImportOptimise()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim tempsDebut As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    tempsDebut = Timer

    ' Ouvrir le fichier source sans mise à jour
    Set wbSource = Workbooks.Open(ThisWorkbook.Path & "\source.xlsx", UpdateLinks:=False, ReadOnly:=True)
    Set wsSource = wbSource.Worksheets(1)
    Set wsDestination = ThisWorkbook.Worksheets(1)

    ' Copie optimisée en bloc
    Dim plageSource As Range
    Set plageSource = wsSource.UsedRange

    If Not plageSource Is Nothing Then
        ' Copier les valeurs seulement (plus rapide)
        wsDestination.Range("A1").Resize(plageSource.Rows.Count, plageSource.Columns.Count).Value = plageSource.Value
    End If

    ' Fermer sans sauvegarder
    wbSource.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Debug.Print "Import optimisé : " & Format(Timer - tempsDebut, "0.000") & "s"
End Sub
```

## 6. Optimisation des formules et calculs

### Formules dynamiques optimisées

```vba
Sub FormulesOptimisees()
    Dim tempsDebut As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Méthode inefficace : formule cellule par cellule
    tempsDebut = Timer
    Dim i As Long
    For i = 1 To 1000
        Cells(i, 1).Formula = "=ROW()*2"
    Next i
    Debug.Print "Formules individuelles : " & Format(Timer - tempsDebut, "0.000") & "s"

    ' Méthode optimisée : formule pour toute la plage
    ' Excel ajuste automatiquement les références relatives pour chaque cellule
    tempsDebut = Timer
    Range("B1:B1000").Formula = "=ROW()*2"
    Debug.Print "Formule en bloc : " & Format(Timer - tempsDebut, "0.000") & "s"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

### Conversion de formules en valeurs

```vba
Sub ConvertirFormulesEnValeurs()
    Dim plage As Range
    Dim tempsDebut As Double

    Set plage = Range("A1:A10000")

    ' S'assurer que les formules sont calculées
    Application.Calculate

    tempsDebut = Timer

    ' Conversion optimisée
    plage.Value = plage.Value  ' Remplace les formules par leurs valeurs

    Debug.Print "Conversion formules->valeurs : " & Format(Timer - tempsDebut, "0.000") & "s"
End Sub
```

## 7. Optimisation de l'interface utilisateur

### Gestion optimisée des UserForms

```vba
Sub ChargementUserFormOptimise()
    Dim frm As UserForm1
    Set frm = New UserForm1

    ' Charger toutes les données en une fois
    Dim donnees As Variant
    donnees = Range("A1:A1000").Value

    ' Remplir la ListBox en une fois au lieu d'ajouter item par item
    frm.ListBox1.List = donnees

    ' Forcer le rafraîchissement et afficher
    frm.Repaint  ' Repaint est une méthode, pas une propriété
    frm.Show
End Sub
```

### Optimisation des contrôles

```vba
' Dans le code du UserForm
Private Sub OptimiserControles()
    ' Désactiver les événements pendant la mise à jour
    Application.EnableEvents = False

    ' Remplir plusieurs contrôles efficacement
    Me.ComboBox1.List = Array("Option1", "Option2", "Option3")
    Me.ListBox1.List = Range("Data").Value

    ' Mettre à jour plusieurs propriétés ensemble
    With Me.TextBox1
        .Text = "Valeur par défaut"
        .Font.Size = 12
        .BackColor = RGB(240, 240, 240)
    End With

    Application.EnableEvents = True
End Sub
```

## 8. Techniques de mise en cache

### Cache de résultats pour éviter les recalculs

```vba
' Variables globales pour le cache
Private cacheResultats As Object  
Private cacheHeures As Object  

Sub InitialiserCache()
    Set cacheResultats = CreateObject("Scripting.Dictionary")
    Set cacheHeures = CreateObject("Scripting.Dictionary")
End Sub

Function CalculComplexeAvecCache(valeur As Double) As Double
    Dim cle As String
    cle = CStr(valeur)

    ' Vérifier si le résultat est déjà en cache
    If cacheResultats.Exists(cle) Then
        ' Vérifier si le cache n'est pas trop ancien (5 minutes)
        If DateDiff("n", cacheHeures(cle), Now) < 5 Then
            CalculComplexeAvecCache = cacheResultats(cle)
            Exit Function
        End If
    End If

    ' Calcul coûteux (simulation)
    Dim resultat As Double
    Dim i As Long
    For i = 1 To 100000
        resultat = resultat + Sin(valeur + i) * Cos(valeur - i)
    Next i

    ' Mettre en cache
    cacheResultats(cle) = resultat
    cacheHeures(cle) = Now

    CalculComplexeAvecCache = resultat
End Function
```

### Cache de données Excel

```vba
Private cacheDonnees As Variant  
Private cacheValide As Boolean  

Function ObtenirDonneesAvecCache() As Variant
    ' Retourner le cache si valide
    If cacheValide Then
        ObtenirDonneesAvecCache = cacheDonnees
        Exit Function
    End If

    ' Charger les données si pas en cache
    cacheDonnees = Range("A1:Z1000").Value
    cacheValide = True

    ObtenirDonneesAvecCache = cacheDonnees
End Function

Sub InvaliderCache()
    cacheValide = False
End Sub
```

## 9. Optimisation de la gestion des événements

### Événements conditionnels

```vba
' Dans le module de la feuille de calcul
Private suspendreEvenements As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
    If suspendreEvenements Then Exit Sub

    ' Traiter seulement les changements importants
    If Not Intersect(Target, Range("A:C")) Is Nothing Then
        ' Logique d'événement optimisée
        suspendreEvenements = True

        ' Traitement en lot
        Dim cellule As Range
        For Each cellule In Target.Cells
            If cellule.Column <= 3 Then  ' Colonnes A, B, C seulement
                ' Traitement optimisé
            End If
        Next cellule

        suspendreEvenements = False
    End If
End Sub
```

### Événements différés

```vba
Private timerEvenement As Date  
Private Const DELAI_EVENEMENT = 0.5 / 86400  ' 0.5 seconde en jours  

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Différer l'exécution pour éviter trop d'appels
    timerEvenement = Now + DELAI_EVENEMENT

    Application.OnTime timerEvenement, "ExecuterEvenementDiffere"
End Sub

Sub ExecuterEvenementDiffere()
    ' Code qui sera exécuté après le délai
    Debug.Print "Événement différé exécuté : " & Now
End Sub
```

## 10. Template d'optimisation complète

```vba
Sub TemplateOptimisationComplete()
    ' === INITIALISATION ===
    Dim config As Object
    Set config = SauvegarderConfiguration()

    Dim tempsDebut As Double, tempsSection As Double
    tempsDebut = Timer

    On Error GoTo Nettoyage

    ' === OPTIMISATIONS SYSTÈME ===
    AppliquerOptimisations

    ' === SECTION 1 : PRÉPARATION ===
    tempsSection = Timer
    Debug.Print "Préparation des données..."

    ' Initialiser le cache si nécessaire
    If cacheResultats Is Nothing Then InitialiserCache

    ' Préparer les structures de données
    Dim donnees As Variant, resultats As Variant
    donnees = ObtenirDonneesAvecCache()

    Debug.Print "Préparation terminée : " & Format(Timer - tempsSection, "0.000") & "s"

    ' === SECTION 2 : TRAITEMENT PRINCIPAL ===
    tempsSection = Timer
    Debug.Print "Traitement principal..."

    ' Utiliser les techniques optimisées appropriées
    ReDim resultats(1 To UBound(donnees, 1), 1 To 5)

    Dim i As Long
    For i = 1 To UBound(donnees, 1)
        ' Traitement optimisé avec cache
        resultats(i, 1) = CalculComplexeAvecCache(donnees(i, 1))
        resultats(i, 2) = donnees(i, 2) * 1.1
        ' ... autres calculs
    Next i

    Debug.Print "Traitement terminé : " & Format(Timer - tempsSection, "0.000") & "s"

    ' === SECTION 3 : FINALISATION ===
    tempsSection = Timer
    Debug.Print "Finalisation..."

    ' Écriture optimisée des résultats
    Range("F1").Resize(UBound(resultats, 1), UBound(resultats, 2)).Value = resultats

    Debug.Print "Finalisation terminée : " & Format(Timer - tempsSection, "0.000") & "s"

    ' === NETTOYAGE ===
Nettoyage:
    RestaurerConfiguration config

    Dim tempsTotal As Double
    tempsTotal = Timer - tempsDebut

    Debug.Print "================================"
    Debug.Print "OPTIMISATION COMPLÈTE TERMINÉE"
    Debug.Print "Temps total : " & Format(tempsTotal, "0.000") & "s"

    If tempsTotal < 1 Then
        Debug.Print "🚀 Performance EXCELLENTE !"
    ElseIf tempsTotal < 5 Then
        Debug.Print "✅ Performance BONNE"
    Else
        Debug.Print "⚠️ Performance à améliorer"
    End If
    Debug.Print "================================"

    If Err.Number <> 0 Then
        Debug.Print "❌ Erreur : " & Err.Description
    End If
End Sub

Function SauvegarderConfiguration() As Object
    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    config("ScreenUpdating") = Application.ScreenUpdating
    config("Calculation") = Application.Calculation
    config("EnableEvents") = Application.EnableEvents
    config("DisplayAlerts") = Application.DisplayAlerts

    Set SauvegarderConfiguration = config
End Function

Sub AppliquerOptimisations()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Sub RestaurerConfiguration(config As Object)
    Application.ScreenUpdating = config("ScreenUpdating")
    Application.Calculation = config("Calculation")
    Application.EnableEvents = config("EnableEvents")
    Application.DisplayAlerts = config("DisplayAlerts")
End Sub
```

## Résumé des techniques avancées

### Gains de performance typiques

1. **API Windows** : 10-50x plus rapide pour certaines opérations système
2. **WorksheetFunction** : 5-20x plus rapide que les boucles VBA équivalentes
3. **Dictionnaires** : 100-1000x plus rapide que les recherches linéaires
4. **Union/Intersect** : 3-10x plus rapide pour les opérations sur plages multiples
5. **Mise en cache** : Gains variables selon la complexité des calculs
6. **E/S optimisées** : 5-50x plus rapide pour les gros fichiers

### Checklist d'optimisation avancée

```vba
' ✓ Utiliser les API Windows pour les opérations système
' ✓ Exploiter WorksheetFunction pour les calculs
' ✓ Implémenter des dictionnaires pour les recherches
' ✓ Utiliser Union/Intersect pour les plages multiples
' ✓ Mettre en cache les résultats coûteux
' ✓ Optimiser les E/O avec lecture/écriture en bloc
' ✓ Convertir les formules en valeurs quand possible
' ✓ Différer les événements non critiques
' ✓ Utiliser un template complet avec gestion d'erreur
```

Ces techniques avancées, appliquées judicieusement selon vos besoins spécifiques, peuvent transformer une application VBA lente en solution haute performance. L'important est de profiler d'abord pour identifier où ces optimisations auront le plus d'impact.

⏭️ [19. Débogage et tests](/19-debogage-tests/)
