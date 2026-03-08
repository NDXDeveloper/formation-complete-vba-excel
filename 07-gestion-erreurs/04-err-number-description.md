🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.4. Err.Number et Err.Description

## Introduction à l'objet Err

L'objet **Err** est le détective de VBA qui enquête sur les erreurs. Quand une erreur survient, VBA enregistre automatiquement des informations précieuses dans cet objet : quel type d'erreur s'est produit (**Err.Number**) et une explication en langage clair (**Err.Description**). Ces informations vous permettent de comprendre exactement ce qui s'est mal passé et de réagir de manière appropriée.

**Analogie simple :**
Imaginez l'objet Err comme un médecin qui examine un patient malade. **Err.Number** serait le code de diagnostic médical (exemple : "J11" pour grippe), et **Err.Description** serait l'explication en français ("infection virale saisonnière"). Les deux informations sont complémentaires : le code pour identifier précisément le problème, la description pour le comprendre.

---

## Err.Number - Le code d'identification de l'erreur

### Qu'est-ce qu'Err.Number ?

**Err.Number** est un nombre entier qui identifie de manière unique chaque type d'erreur VBA. C'est comme un code postal : chaque erreur a son propre numéro, ce qui permet de l'identifier sans ambiguïté.

### Valeurs importantes d'Err.Number

#### 0 = Aucune erreur

```vba
Sub VerifierAbsenceErreur()
    On Error Resume Next

    Range("A1").Value = "Test"  ' Opération simple qui réussit

    If Err.Number = 0 Then
        MsgBox "Aucune erreur - tout va bien !"
    Else
        MsgBox "Erreur détectée : " & Err.Number
    End If

    On Error GoTo 0
End Sub
```

#### Erreurs courantes et leurs numéros

```vba
Sub ExemplesErreursCourantes()
    On Error Resume Next

    ' Erreur 9 : Subscript out of range
    Worksheets("FeuilleInexistante").Range("A1").Value = "Test"
    If Err.Number = 9 Then
        MsgBox "Erreur 9 : Élément inexistant (feuille, plage, etc.)"
        Err.Clear
    End If

    ' Erreur 11 : Division by zero
    Dim resultat As Double
    resultat = 10 / 0
    If Err.Number = 11 Then
        MsgBox "Erreur 11 : Division par zéro"
        Err.Clear
    End If

    ' Erreur 13 : Type mismatch
    Dim nombre As Integer
    nombre = "Texte"
    If Err.Number = 13 Then
        MsgBox "Erreur 13 : Types de données incompatibles"
        Err.Clear
    End If

    On Error GoTo 0
End Sub
```

### Tableau des erreurs VBA les plus fréquentes

| Numéro | Nom anglais | Description française |
|--------|-------------|----------------------|
| **0** | No error | Aucune erreur |
| **9** | Subscript out of range | Indice hors limites |
| **11** | Division by zero | Division par zéro |
| **13** | Type mismatch | Non-correspondance de type |
| **53** | File not found | Fichier non trouvé |
| **70** | Permission denied | Autorisation refusée |
| **91** | Object variable not set | Variable objet non définie |
| **438** | Object doesn't support this property/method | L'objet ne prend pas en charge cette propriété/méthode |
| **1004** | Application-defined or object-defined error | Erreur définie par l'application ou l'objet |

---

## Err.Description - L'explication en français

### Qu'est-ce qu'Err.Description ?

**Err.Description** est une chaîne de caractères qui explique l'erreur dans un langage compréhensible. Contrairement au numéro qui est technique, la description est destinée à être lue par un humain.

### Exemples de descriptions courantes

```vba
Sub ExemplesDescriptions()
    On Error Resume Next

    ' Générer différentes erreurs pour voir leurs descriptions

    ' Erreur 9
    Worksheets("Inexistant").Select
    Debug.Print "Erreur " & Err.Number & ": " & Err.Description
    ' Affiche : "Erreur 9: L'indice n'appartient pas à la sélection"
    Err.Clear

    ' Erreur 11
    Dim test As Double
    test = 5 / 0
    Debug.Print "Erreur " & Err.Number & ": " & Err.Description
    ' Affiche : "Erreur 11: Division par zéro"
    Err.Clear

    ' Erreur 1004
    Range("A0").Select  ' A0 n'existe pas
    Debug.Print "Erreur " & Err.Number & ": " & Err.Description
    ' Affiche : "Erreur 1004: Erreur définie par l'application ou l'objet"
    Err.Clear

    On Error GoTo 0
End Sub
```

### Descriptions en fonction de la langue

Les descriptions d'erreur s'affichent dans la langue de votre version d'Excel :

```vba
Sub DescriptionsMultilingues()
    On Error Resume Next

    ' Division par zéro
    Dim resultat As Double
    resultat = 10 / 0

    ' En français : "Division par zéro"
    ' En anglais : "Division by zero"
    ' En espagnol : "División por cero"

    Debug.Print "Langue détectée : " & Err.Description

    On Error GoTo 0
End Sub
```

---

## Utilisation pratique combinée

### Pattern de base : Number + Description

```vba
Sub PatternDeBase()
    On Error GoTo GestionErreur

    ' Code qui peut générer une erreur
    Worksheets("Données").Range("A1").Value = Range("B1").Value / Range("C1").Value

    MsgBox "Opération réussie !"
    Exit Sub

GestionErreur:
    Dim messageComplet As String
    messageComplet = "Une erreur est survenue :" & vbCrLf & _
                     "Code : " & Err.Number & vbCrLf & _
                     "Description : " & Err.Description

    MsgBox messageComplet, vbCritical, "Erreur"
End Sub
```

### Gestion spécialisée selon le numéro

```vba
Sub GestionSpecialisee()
    On Error GoTo GestionErreur

    ' Tentative d'ouverture d'un fichier
    Workbooks.Open "C:\MonFichier.xlsx"

    ' Si on arrive ici, tout va bien
    MsgBox "Fichier ouvert avec succès"
    Exit Sub

GestionErreur:
    Select Case Err.Number
        Case 53  ' File not found
            MsgBox "Le fichier n'existe pas." & vbCrLf & _
                   "Vérifiez le chemin : C:\MonFichier.xlsx"

        Case 70  ' Permission denied
            MsgBox "Accès refusé au fichier." & vbCrLf & _
                   "Le fichier est peut-être ouvert par un autre utilisateur."

        Case 1004  ' Application error
            MsgBox "Erreur Excel : " & Err.Description & vbCrLf & _
                   "Le fichier est peut-être corrompu."

        Case Else
            MsgBox "Erreur inattendue :" & vbCrLf & _
                   "Code : " & Err.Number & vbCrLf & _
                   "Description : " & Err.Description
    End Select
End Sub
```

---

## Propriétés supplémentaires de l'objet Err

### Err.Source - Source de l'erreur

```vba
Sub ExempleErrSource()
    On Error GoTo GestionErreur

    ' Générer une erreur
    Range("A1").Value = 10 / 0

    Exit Sub

GestionErreur:
    MsgBox "Source de l'erreur : " & Err.Source
    ' Affiche généralement "VBAProject" ou le nom de l'application
End Sub
```

### Err.HelpFile et Err.HelpContext - Aide contextuelle

```vba
Sub ExempleAideContextuelle()
    On Error GoTo GestionErreur

    ' Code avec erreur
    Worksheets("Inexistant").Select

    Exit Sub

GestionErreur:
    MsgBox "Erreur : " & Err.Description & vbCrLf & _
           "Fichier d'aide : " & Err.HelpFile & vbCrLf & _
           "Contexte d'aide : " & Err.HelpContext
End Sub
```

---

## Méthodes de l'objet Err

### Err.Clear - Effacer les informations d'erreur

```vba
Sub ExempleErrClear()
    On Error Resume Next

    ' Première erreur
    Worksheets("FeuilleInexistante").Range("A1").Value = "Test"
    MsgBox "Première erreur - Numéro : " & Err.Number  ' 9

    ' Effacer l'erreur
    Err.Clear
    MsgBox "Après Clear - Numéro : " & Err.Number      ' 0

    ' Nouvelle erreur
    Dim test As Double
    test = 10 / 0
    MsgBox "Nouvelle erreur - Numéro : " & Err.Number  ' 11

    On Error GoTo 0
End Sub
```

### Err.Raise - Générer une erreur personnalisée

```vba
Sub ExempleErrRaise()
    On Error GoTo GestionErreur

    Dim age As Integer
    age = -5  ' Valeur invalide

    ' Vérification et génération d'erreur personnalisée
    If age < 0 Then
        Err.Raise Number:=9999, _
                  Description:="L'âge ne peut pas être négatif", _
                  Source:="ValidationAge"
    End If

    MsgBox "Âge valide : " & age
    Exit Sub

GestionErreur:
    MsgBox "Erreur personnalisée :" & vbCrLf & _
           "Code : " & Err.Number & vbCrLf & _
           "Description : " & Err.Description & vbCrLf & _
           "Source : " & Err.Source
End Sub
```

---

## Techniques avancées d'analyse

### Journalisation détaillée des erreurs

```vba
Sub JournalisationDetaillee()
    On Error GoTo GestionErreur

    ' Code principal
    Dim i As Integer
    For i = 1 To 10
        Cells(i, 1).Value = 100 / Cells(i, 2).Value
    Next i

    Exit Sub

GestionErreur:
    ' Créer un journal détaillé
    Dim ligneJournal As String
    ligneJournal = Format(Now, "yyyy-mm-dd hh:mm:ss") & vbTab & _
                   Err.Number & vbTab & _
                   Err.Description & vbTab & _
                   Err.Source & vbTab & _
                   "Ligne approximative : " & i

    ' Enregistrer dans une feuille de journal
    If WorksheetExists("Journal") Then
        Dim dernieLigne As Long
        dernieLigne = Worksheets("Journal").Cells(Rows.Count, 1).End(xlUp).Row + 1
        Worksheets("Journal").Cells(dernieLigne, 1).Value = ligneJournal
    Else
        Debug.Print ligneJournal
    End If

    ' Décider de la suite
    If MsgBox("Erreur ligne " & i & ": " & Err.Description & vbCrLf & _
              "Continuer avec la ligne suivante ?", vbYesNo) = vbYes Then
        Cells(i, 1).Value = "Erreur"
        Resume Next
    End If
End Sub

Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    WorksheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function
```

### Diagnostic automatique selon le type d'erreur

```vba
Sub DiagnosticAutomatique()
    On Error GoTo GestionErreur

    ' Code qui peut générer diverses erreurs
    ProcessData
    Exit Sub

GestionErreur:
    Dim diagnostic As String

    Select Case Err.Number
        Case 9  ' Subscript out of range
            diagnostic = "Diagnostic : Élément inexistant" & vbCrLf & _
                        "Solutions possibles :" & vbCrLf & _
                        "- Vérifier que la feuille existe" & vbCrLf & _
                        "- Vérifier l'orthographe du nom" & vbCrLf & _
                        "- Vérifier les indices de tableau"

        Case 11  ' Division by zero
            diagnostic = "Diagnostic : Division par zéro" & vbCrLf & _
                        "Solutions possibles :" & vbCrLf & _
                        "- Vérifier que le diviseur n'est pas zéro" & vbCrLf & _
                        "- Ajouter une condition IF avant la division"

        Case 13  ' Type mismatch
            diagnostic = "Diagnostic : Types incompatibles" & vbCrLf & _
                        "Solutions possibles :" & vbCrLf & _
                        "- Vérifier le type des variables" & vbCrLf & _
                        "- Utiliser des fonctions de conversion (CInt, CDbl, etc.)"

        Case 53  ' File not found
            diagnostic = "Diagnostic : Fichier introuvable" & vbCrLf & _
                        "Solutions possibles :" & vbCrLf & _
                        "- Vérifier le chemin complet" & vbCrLf & _
                        "- Vérifier que le fichier existe" & vbCrLf & _
                        "- Vérifier les droits d'accès"

        Case Else
            diagnostic = "Diagnostic : Erreur non répertoriée" & vbCrLf & _
                        "Consulter la documentation Microsoft VBA"
    End Select

    MsgBox "ERREUR " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           diagnostic, vbCritical, "Diagnostic d'erreur"
End Sub

Sub ProcessData()
    ' Code exemple qui peut générer des erreurs
    Worksheets("Données").Range("A1").Value = 10 / 0
End Sub
```

---

## Création d'erreurs personnalisées

### Système d'erreurs numérotées

```vba
' Constantes pour erreurs personnalisées (à mettre en début de module)
Const ERR_AGE_NEGATIF = 9001  
Const ERR_EMAIL_INVALIDE = 9002  
Const ERR_DONNEES_MANQUANTES = 9003  

Sub ValidationAvecErreursPersonnalisees()
    On Error GoTo GestionErreur

    Dim age As Integer
    Dim email As String

    age = Range("A1").Value
    email = Range("B1").Value

    ' Validations avec erreurs personnalisées
    If age < 0 Then
        Err.Raise ERR_AGE_NEGATIF, "ValidationAge", "L'âge ne peut pas être négatif"
    End If

    If InStr(email, "@") = 0 Then
        Err.Raise ERR_EMAIL_INVALIDE, "ValidationEmail", "L'email doit contenir un @"
    End If

    If Range("C1").Value = "" Then
        Err.Raise ERR_DONNEES_MANQUANTES, "ValidationDonnees", "Le nom est obligatoire"
    End If

    MsgBox "Toutes les validations sont OK !"
    Exit Sub

GestionErreur:
    Select Case Err.Number
        Case ERR_AGE_NEGATIF
            MsgBox "Erreur de saisie : " & Err.Description & vbCrLf & _
                   "Veuillez saisir un âge positif en A1"
            Range("A1").Select

        Case ERR_EMAIL_INVALIDE
            MsgBox "Erreur de format : " & Err.Description & vbCrLf & _
                   "Veuillez saisir un email valide en B1"
            Range("B1").Select

        Case ERR_DONNEES_MANQUANTES
            MsgBox "Donnée manquante : " & Err.Description & vbCrLf & _
                   "Veuillez saisir un nom en C1"
            Range("C1").Select

        Case Else
            MsgBox "Erreur système : " & Err.Number & " - " & Err.Description
    End Select
End Sub
```

---

## Bonnes pratiques avec Err.Number et Err.Description

### 1. Toujours vérifier Err.Number avant Err.Description

```vba
Sub VerificationCorrecte()
    On Error Resume Next

    ' Opération potentiellement problématique
    Worksheets("FeuilleInexistante").Range("A1").Value = "Test"

    ' CORRECT : Vérifier d'abord le numéro
    If Err.Number <> 0 Then
        MsgBox "Erreur détectée : " & Err.Description
        Err.Clear
    End If

    On Error GoTo 0
End Sub
```

### 2. Effacer les erreurs après traitement

```vba
Sub NettoyageErreurs()
    On Error Resume Next

    Dim i As Integer
    For i = 1 To 5
        ' Tentative d'opération
        Range("A" & i).Value = 100 / Range("B" & i).Value

        If Err.Number <> 0 Then
            Range("A" & i).Value = "Erreur"
            Err.Clear  ' IMPORTANT : Nettoyer pour la prochaine itération
        End If
    Next i

    On Error GoTo 0
End Sub
```

### 3. Messages d'erreur conviviaux

```vba
Function MessageConvivial(numeroErreur As Long, description As String) As String
    Select Case numeroErreur
        Case 9
            MessageConvivial = "Un élément demandé n'existe pas. " & _
                              "Vérifiez les noms de feuilles et plages."

        Case 11
            MessageConvivial = "Impossible de diviser par zéro. " & _
                              "Vérifiez vos données de calcul."

        Case 13
            MessageConvivial = "Les données ne sont pas dans le bon format. " & _
                              "Vérifiez que les nombres sont bien des nombres."

        Case 53
            MessageConvivial = "Le fichier demandé est introuvable. " & _
                              "Vérifiez le chemin et l'existence du fichier."

        Case Else
            MessageConvivial = "Erreur technique : " & description
    End Select
End Function

Sub UtiliserMessageConvivial()
    On Error GoTo GestionErreur

    ' Code qui peut générer une erreur
    Range("A1").Value = 10 / 0

    Exit Sub

GestionErreur:
    MsgBox MessageConvivial(Err.Number, Err.Description), _
           vbExclamation, "Information"
End Sub
```

---

## Débogage et tests avec Err

### Forcer des erreurs pour tester

```vba
Sub TesterGestionErreurs()
    ' Test de différents scénarios d'erreur pour valider la gestion

    MsgBox "Test 1 : Division par zéro"
    TestErreur1

    MsgBox "Test 2 : Feuille inexistante"
    TestErreur2

    MsgBox "Test 3 : Type incorrect"
    TestErreur3

    MsgBox "Tests terminés"
End Sub

Sub TestErreur1()
    On Error GoTo GestionErreur
    Dim x As Double
    x = 10 / 0
    Exit Sub
GestionErreur:
    Debug.Print "Test 1 - Erreur " & Err.Number & ": " & Err.Description
End Sub

Sub TestErreur2()
    On Error GoTo GestionErreur
    Worksheets("TestInexistant").Range("A1").Value = "Test"
    Exit Sub
GestionErreur:
    Debug.Print "Test 2 - Erreur " & Err.Number & ": " & Err.Description
End Sub

Sub TestErreur3()
    On Error GoTo GestionErreur
    Dim nombre As Integer
    nombre = "Pas un nombre"
    Exit Sub
GestionErreur:
    Debug.Print "Test 3 - Erreur " & Err.Number & ": " & Err.Description
End Sub
```

---

## Récapitulatif

### Points essentiels à retenir

1. **Err.Number** identifie le type d'erreur avec un code numérique unique
2. **Err.Description** explique l'erreur en langage compréhensible
3. **Err.Number = 0** signifie qu'il n'y a aucune erreur
4. **Err.Clear** efface les informations d'erreur stockées
5. **Err.Raise** permet de créer des erreurs personnalisées
6. **Toujours vérifier Err.Number avant d'utiliser Err.Description**

### Erreurs VBA les plus fréquentes à mémoriser

| Code | Situation typique | Action recommandée |
|------|------------------|-------------------|
| **9** | Feuille/plage inexistante | Vérifier existence avant utilisation |
| **11** | Division par zéro | Contrôler le diviseur |
| **13** | Mauvais type de données | Valider et convertir les données |
| **53** | Fichier introuvable | Vérifier chemin et existence |
| **1004** | Erreur Excel générale | Analyser le contexte spécifique |

### Modèle de code recommandé

```vba
Sub ModeleRecommande()
    On Error GoTo GestionErreur

    ' Code principal

    Exit Sub

GestionErreur:
    ' Log technique pour développeur
    Debug.Print "Erreur " & Err.Number & ": " & Err.Description

    ' Message convivial pour utilisateur
    Select Case Err.Number
        Case 9, 11, 13, 53, 1004
            ' Gestion spécialisée des erreurs courantes
        Case Else
            ' Gestion générique
            MsgBox "Une erreur inattendue s'est produite." & vbCrLf & _
                   "Code : " & Err.Number & vbCrLf & _
                   "Description : " & Err.Description
    End Select
End Sub
```

Maîtriser `Err.Number` et `Err.Description` vous donne les outils pour créer une gestion d'erreurs précise et informative. Dans la section suivante, nous découvrirons les bonnes pratiques générales pour créer du code VBA robuste et professionnel.

⏭️ [Bonnes pratiques de gestion d'erreurs](/07-gestion-erreurs/05-bonnes-pratiques-gestion-erreurs.md)
