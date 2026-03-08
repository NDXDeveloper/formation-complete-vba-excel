🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.3 Gestion de la mémoire

## Introduction

La gestion de la mémoire est un aspect crucial mais souvent négligé en VBA. Une mauvaise gestion peut provoquer des ralentissements, des blocages, voire des plantages d'Excel. Cette section vous apprendra à comprendre comment VBA utilise la mémoire et comment optimiser son usage pour des applications plus stables et performantes.

## Comprendre la mémoire en VBA

### Qu'est-ce que la mémoire en programmation ?

La mémoire est l'espace de stockage temporaire que votre ordinateur utilise pour faire fonctionner les programmes. Quand vous lancez Excel et VBA, ils "réservent" une partie de cette mémoire pour stocker :
- Les variables que vous créez
- Les objets que vous manipulez
- Les données temporaires des calculs

### Types de mémoire en VBA

**Mémoire de pile (Stack) :** Stocke les variables locales et les paramètres de fonctions. Gérée automatiquement par VBA.

**Mémoire de tas (Heap) :** Stocke les objets, les tableaux dynamiques et les chaînes de caractères. Nécessite une gestion plus attentive.

### Signes d'un problème de mémoire

- Excel devient très lent
- Messages d'erreur "Mémoire insuffisante"
- Excel se bloque ou plante
- L'ordinateur devient globalement lent
- Erreur "Overflow" ou "Out of memory"

## Les variables et leur impact mémoire

### Types de variables et consommation mémoire

```vba
Sub TypesVariablesMemoire()
    ' Variables numériques (tailles fixes)
    Dim unByte As Byte          ' 1 octet
    Dim unInteger As Integer    ' 2 octets
    Dim unLong As Long          ' 4 octets
    Dim unSingle As Single      ' 4 octets
    Dim unDouble As Double      ' 8 octets

    ' Variables texte (taille variable)
    Dim unTexte As String       ' 4 octets + longueur du texte

    ' Variables objets (taille variable)
    Dim unRange As Range        ' 4 octets + données de l'objet
    Dim unWorkbook As Workbook  ' 4 octets + données de l'objet
End Sub
```

### Choisir le bon type de variable

```vba
Sub BonChoixTypes()
    ' MAUVAIS : Integer peut déborder facilement
    Dim compteurPetit As Integer  ' Limité à 32 767

    ' BON : Long pour les compteurs
    Dim compteurGrand As Long     ' Peut aller jusqu'à 2 milliards

    ' MAUVAIS : Variant consomme plus de mémoire
    Dim valeurVariant As Variant  ' 16 octets minimum

    ' BON : Type spécifique selon le besoin
    Dim valeurDouble As Double    ' 8 octets seulement

    ' MAUVAIS : String de taille fixe inutile
    Dim nomFixe As String * 255   ' 255 octets même si on utilise 5 caractères

    ' BON : String dynamique
    Dim nomDynamique As String    ' Prend seulement la place nécessaire
End Sub
```

## Gestion des objets

### Le problème des références d'objets

Quand vous créez des références à des objets Excel, VBA maintient ces objets en mémoire même après leur utilisation.

```vba
Sub ProblemeReferencesObjets()
    Dim ws As Worksheet
    Dim plage As Range
    Dim cell As Range

    Set ws = Worksheets("Feuil1")
    Set plage = ws.Range("A1:Z1000")

    ' Cette boucle crée 26000 références d'objets !
    For Each cell In plage
        ' Traitement des cellules
        cell.Value = cell.Value * 2
    Next cell

    ' Problème : toutes les références restent en mémoire
End Sub
```

### Solution : Libérer les références d'objets

```vba
Sub LibererReferencesObjets()
    Dim ws As Worksheet
    Dim plage As Range
    Dim cell As Range

    Set ws = Worksheets("Feuil1")
    Set plage = ws.Range("A1:Z1000")

    For Each cell In plage
        cell.Value = cell.Value * 2
    Next cell

    ' IMPORTANT : Libérer les références
    Set cell = Nothing
    Set plage = Nothing
    Set ws = Nothing
End Sub
```

### Règle d'or pour les objets

```vba
Sub RegleOrObjets()
    Dim monObjet As Object

    ' 1. Créer la référence
    Set monObjet = CreateObject("Excel.Application")

    ' 2. Utiliser l'objet
    ' ... votre code ici

    ' 3. TOUJOURS libérer avec Set = Nothing
    Set monObjet = Nothing
End Sub
```

## Gestion des tableaux

### Tableaux et consommation mémoire

Les tableaux peuvent consommer énormément de mémoire selon leur taille et leur type.

```vba
Sub CalculerTailleTableau()
    ' Exemple : tableau de 1000 x 1000 doubles
    Dim grandTableau(1 To 1000, 1 To 1000) As Double

    ' Calcul de la mémoire : 1000 × 1000 × 8 octets = 8 MB !
    Debug.Print "Ce tableau consomme 8 MB de mémoire"

    ' Avec des Variants, ce serait 16 MB !
    Dim tableauVariant(1 To 1000, 1 To 1000) As Variant
End Sub
```

### Dimensionnement intelligent des tableaux

```vba
Sub DimensionnementIntelligent()
    Dim donnees() As Variant
    Dim tailleNecessaire As Long

    ' Calculer la taille exacte nécessaire
    tailleNecessaire = ActiveSheet.UsedRange.Rows.Count

    ' Dimensionner juste ce qu'il faut
    ReDim donnees(1 To tailleNecessaire, 1 To 5)

    ' Traitement...

    ' Libérer le tableau après utilisation
    Erase donnees
End Sub
```

### Redimensionnement efficace avec ReDim Preserve

```vba
Sub RedimensionnementEfficace()
    Dim tableau() As String
    Dim taille As Long

    ' Commencer petit
    taille = 100
    ReDim tableau(1 To taille)

    ' Si on a besoin de plus d'espace, doubler la taille
    ' (plus efficace que d'augmenter de 1 à chaque fois)
    If taille < 200 Then
        taille = taille * 2
        ReDim Preserve tableau(1 To taille)
    End If

    ' Redimensionner à la taille finale pour économiser la mémoire
    ReDim Preserve tableau(1 To 150)  ' Si on n'en utilise que 150
End Sub
```

## Gestion des chaînes de caractères

### Les chaînes consomment beaucoup de mémoire

```vba
Sub ProblemeChaines()
    Dim texte As String
    Dim i As Long

    ' TRÈS INEFFICACE : Recrée la chaîne complète à chaque itération
    For i = 1 To 10000
        texte = texte & "Ligne " & i & vbCrLf
    Next i

    ' À la fin, 'texte' peut faire plusieurs MB !
End Sub
```

### Solution : Utiliser un tableau puis Join

```vba
Sub ChainesOptimisees()
    Dim lignes() As String
    Dim i As Long
    Dim texteComplet As String

    ' Stocker chaque ligne dans un tableau
    ReDim lignes(1 To 10000)

    For i = 1 To 10000
        lignes(i) = "Ligne " & i
    Next i

    ' Assembler en une seule fois (très rapide)
    texteComplet = Join(lignes, vbCrLf)

    ' Libérer le tableau
    Erase lignes
End Sub
```

## Surveillance de l'utilisation mémoire

### Fonction pour surveiller la mémoire

```vba
' Déclaration d'API Windows pour surveiller la mémoire
#If VBA7 Then
    Private Declare PtrSafe Function GetProcessMemoryInfo Lib "psapi.dll" _
        (ByVal hProcess As LongPtr, ByRef ppsmemCounters As Any, ByVal cb As Long) As Long
    Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32.dll" () As LongPtr
#Else
    Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" _
        (ByVal hProcess As Long, ByRef ppsmemCounters As Any, ByVal cb As Long) As Long
    Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
#End If

Function ObtenirUtilisationMemoire() As Long
    ' Version simplifiée : utilise les informations VBA disponibles
    ' Retourne une estimation de l'utilisation mémoire
    ObtenirUtilisationMemoire = 0  ' Placeholder - implémentation complète nécessiterait des API
End Function
```

### Surveillance simple avec Timer

```vba
Sub SurveillanceMemoire()
    Dim tempsDebut As Double
    Dim tableauTest() As Double

    tempsDebut = Timer

    ' Créer un gros tableau pour tester
    ReDim tableauTest(1 To 100000)
    Debug.Print "Création tableau : " & Format(Timer - tempsDebut, "0.00") & "s"

    ' Libérer le tableau
    tempsDebut = Timer
    Erase tableauTest
    Debug.Print "Libération tableau : " & Format(Timer - tempsDebut, "0.00") & "s"
End Sub
```

## Bonnes pratiques de gestion mémoire

### Déclarer les variables au bon endroit

```vba
Sub BonnePorteeVariables()
    ' MAUVAIS : Variables globales gardées en mémoire
    ' Public grandeVariable(1 To 100000) As Double

    ' BON : Variables locales libérées automatiquement
    Dim variableLocale(1 To 1000) As Double

    ' Traitement...

    ' La variable est automatiquement libérée à la fin de la Sub
End Sub
```

### Utiliser des constantes quand c'est possible

```vba
Sub UtiliserConstantes()
    ' BON : Les constantes ne consomment pas de mémoire d'exécution
    Const MAX_LIGNES As Long = 10000
    Const NOM_FICHIER As String = "MonFichier.xlsx"

    ' Au lieu de variables qui occupent de la mémoire
    Dim maxLignes As Long
    Dim nomFichier As String
    maxLignes = 10000
    nomFichier = "MonFichier.xlsx"
End Sub
```

### Éviter les variables globales inutiles

```vba
' ÉVITER : Variables globales qui restent en mémoire
Public grosTableauGlobal(1 To 50000) As Variant  
Public chaineGlobale As String  

Sub MauvaiseGestionGlobale()
    ' Ces variables restent en mémoire pendant toute la session Excel
    grosTableauGlobal(1) = "Données"
    chaineGlobale = "Très longue chaîne..."
End Sub

' PRÉFÉRER : Variables locales ou passage par paramètres
Sub BonneGestionLocale()
    Dim tableauLocal(1 To 1000) As Variant
    Dim chaineLocale As String

    ' Traitement...

    ' Variables automatiquement libérées à la fin
End Sub
```

## Gestion des collections et dictionnaires

### Collections : attention à la croissance

```vba
Sub GestionCollections()
    Dim maCollection As Collection
    Set maCollection = New Collection

    ' Ajouter des éléments
    Dim i As Long
    For i = 1 To 10000
        maCollection.Add "Élément " & i, "Clé" & i
    Next i

    ' IMPORTANT : Vider la collection avant de la détruire
    Dim j As Long
    For j = maCollection.Count To 1 Step -1
        maCollection.Remove j
    Next j

    Set maCollection = Nothing
End Sub
```

### Alternative avec Dictionary (si disponible)

```vba
Sub GestionDictionary()
    ' Nécessite une référence à "Microsoft Scripting Runtime"
    Dim monDict As Object
    Set monDict = CreateObject("Scripting.Dictionary")

    ' Ajouter des éléments
    Dim i As Long
    For i = 1 To 10000
        monDict.Add "Clé" & i, "Valeur " & i
    Next i

    ' Nettoyer le dictionnaire
    monDict.RemoveAll
    Set monDict = Nothing
End Sub
```

## Techniques avancées de gestion mémoire

### Pooling d'objets pour réutilisation

```vba
' Module de classe pour gérer un pool d'objets Range
Private poolRanges As Collection

Sub InitialiserPool()
    Set poolRanges = New Collection
End Sub

Function ObtenirRange() As Range
    If poolRanges.Count > 0 Then
        ' Réutiliser un Range existant
        Set ObtenirRange = poolRanges(1)
        poolRanges.Remove 1
    Else
        ' Créer un nouveau Range si le pool est vide
        Set ObtenirRange = Nothing  ' Sera créé selon le besoin
    End If
End Function

Sub RendreRange(r As Range)
    ' Remettre le Range dans le pool pour réutilisation
    poolRanges.Add r
End Sub
```

### Gestion des gros volumes de données

```vba
Sub TraiterGrosVolume()
    Dim donnees As Variant
    Dim resultat As Variant
    Dim i As Long, j As Long

    ' Charger les données par blocs pour éviter les problèmes mémoire
    Const TAILLE_BLOC = 10000

    For i = 1 To 100000 Step TAILLE_BLOC
        ' Charger un bloc
        donnees = Range("A" & i & ":E" & (i + TAILLE_BLOC - 1)).Value

        ' Traiter le bloc
        ReDim resultat(1 To UBound(donnees, 1), 1 To 3)

        For j = 1 To UBound(donnees, 1)
            resultat(j, 1) = donnees(j, 1) * 2
            resultat(j, 2) = donnees(j, 2) & " traité"
            resultat(j, 3) = j
        Next j

        ' Écrire le résultat
        Range("F" & i & ":H" & (i + UBound(resultat, 1) - 1)).Value = resultat

        ' Libérer la mémoire du bloc
        Erase donnees
        Erase resultat

        ' Forcer le nettoyage mémoire (optionnel)
        DoEvents
    Next i
End Sub
```

## Diagnostic et résolution des problèmes mémoire

### Identifier les fuites mémoire

```vba
Sub DetecterFuitesMemoire()
    Dim avant As Long, apres As Long

    ' Mesure avant
    avant = FreeMemory()  ' Fonction hypothétique

    ' Code suspect de fuite mémoire
    Dim i As Long
    For i = 1 To 1000
        Dim obj As Object
        Set obj = CreateObject("Excel.Application")
        ' OUBLI : Set obj = Nothing  <- Fuite mémoire !
    Next i

    ' Mesure après
    apres = FreeMemory()

    If apres < avant Then
        Debug.Print "Fuite mémoire détectée : " & (avant - apres) & " octets"
    End If
End Sub
```

### Nettoyage manuel de la mémoire

```vba
Sub NettoyageMemoire()
    ' Forcer la libération des objets détruits
    DoEvents

    ' Dans certains cas, forcer le garbage collector
    ' (Note : VBA n'a pas de garbage collector explicite comme .NET)

    ' Vider les variables globales si nécessaire
    ' Erase tableauGlobal

    ' Fermer les objets externes
    ' Application.Quit (pour des applications externes)
End Sub
```

## Cas particuliers et pièges à éviter

### Piège 1 : Les boucles avec objets

```vba
' PIÈGE : Création d'objets dans une boucle
Sub PiegeObjetsEnBoucle()
    Dim i As Long

    Dim ws As Worksheet

    For i = 1 To 1000
        Set ws = Worksheets(1)  ' MAUVAIS : Réassigné inutilement à chaque itération
        ' Traitement...
        ' OUBLI : Set ws = Nothing
    Next i
End Sub

' SOLUTION : Déclaration en dehors de la boucle
Sub SolutionObjetsEnBoucle()
    Dim i As Long
    Dim ws As Worksheet

    Set ws = Worksheets(1)  ' Une seule déclaration

    For i = 1 To 1000
        ' Traitement avec ws...
    Next i

    Set ws = Nothing  ' Libération unique
End Sub
```

### Piège 2 : Les variants avec gros volumes

```vba
' PIÈGE : Variants pour de gros tableaux
Sub PiegeVariants()
    Dim donnees As Variant
    donnees = Range("A1:Z10000").Value  ' Peut consommer beaucoup de mémoire

    ' Les Variants stockent aussi le type de chaque cellule !
End Sub

' SOLUTION : Types spécifiques quand possible
Sub SolutionTypes()
    ' Si vous savez que ce sont des nombres
    Dim nombres() As Double
    ' Charger et convertir manuellement si nécessaire
End Sub
```

### Piège 3 : Collections qui grandissent indéfiniment

```vba
' PIÈGE : Collection qui grandit sans limite
Public cacheGlobal As Collection

Sub PiegeCache()
    If cacheGlobal Is Nothing Then Set cacheGlobal = New Collection

    ' Ajouter toujours sans jamais nettoyer
    cacheGlobal.Add Now, "timestamp" & cacheGlobal.Count
    ' Cette collection va grandir indéfiniment !
End Sub

' SOLUTION : Nettoyer périodiquement
Sub SolutionCache()
    If cacheGlobal Is Nothing Then Set cacheGlobal = New Collection

    ' Nettoyer si la collection devient trop grosse
    If cacheGlobal.Count > 1000 Then
        Set cacheGlobal = New Collection
    End If

    cacheGlobal.Add Now, "timestamp" & cacheGlobal.Count
End Sub
```

## Résumé des bonnes pratiques

### Règles essentielles

1. **Toujours libérer les objets** avec `Set objet = Nothing`
2. **Dimensionner les tableaux juste nécessaire** avec `ReDim`
3. **Utiliser `Erase`** pour libérer les gros tableaux
4. **Éviter les variables globales** inutiles
5. **Choisir le bon type de variable** selon le besoin
6. **Traiter par blocs** les très gros volumes
7. **Nettoyer les collections** avant destruction

### Checklist de gestion mémoire

```vba
Sub ChecklistMemoire()
    ' ✓ Variables déclarées avec le bon type
    Dim compteur As Long  ' Pas Integer

    ' ✓ Objets avec gestion de libération
    Dim ws As Worksheet
    On Error GoTo Nettoyage
    Set ws = ActiveSheet

    ' Votre traitement...

    ' ✓ Libération systématique
Nettoyage:
    Set ws = Nothing

    ' ✓ Tableaux libérés après usage
    ' Erase monTableau

    ' ✓ Collections vidées (pas de méthode Clear en VBA)
    ' Set maCollection = New Collection  ' Réinitialiser
End Sub
```

### Surveillance continue

```vba
Sub TemplateAvecSurveillance()
    Dim tempsDebut As Double

    tempsDebut = Timer
    Debug.Print "Début traitement - " & Now

    ' Votre code optimisé ici

    Debug.Print "Fin traitement - Durée : " & Format(Timer - tempsDebut, "0.00") & "s"

    ' Vérification finale
    If Timer - tempsDebut > 10 Then
        Debug.Print "ATTENTION : Traitement long, vérifier l'optimisation mémoire"
    End If
End Sub
```

Une bonne gestion de la mémoire en VBA n'est pas seulement une question de performance, c'est aussi une question de stabilité. En appliquant ces bonnes pratiques, vous créerez des applications VBA plus robustes et plus professionnelles.

⏭️ [Profilage du code](/18-optimisation-performance/04-profilage-code.md)
