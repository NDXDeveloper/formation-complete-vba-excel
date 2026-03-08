🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.5. Bonnes pratiques de gestion d'erreurs

## Introduction aux bonnes pratiques

Les **bonnes pratiques** de gestion d'erreurs sont l'ensemble des règles et techniques qui transforment votre code VBA fragile en un code professionnel, robuste et maintenable. Ces pratiques sont le résultat de l'expérience collective de milliers de développeurs et vous éviteront de nombreux pièges et frustrations.

**Analogie simple :**
Pensez aux bonnes pratiques comme aux règles de sécurité routière. Vous pouvez techniquement conduire sans ceinture de sécurité, mais suivre les règles vous protège des accidents et rend la conduite plus sûre pour tous. De même, vous pouvez écrire du code VBA qui fonctionne sans gestion d'erreurs, mais suivre les bonnes pratiques protège votre code et vos utilisateurs des problèmes imprévus.

---

## Principe 1 : Prévention avant correction

### La règle d'or : Mieux vaut prévenir que guérir

La meilleure gestion d'erreur est celle qui évite que l'erreur se produise. **Prévenez les erreurs** en vérifiant les conditions avant d'exécuter des opérations risquées.

#### ❌ Code fragile (sans prévention)

```vba
Sub CodeFragile()
    ' Aucune vérification - code dangereux
    Dim resultat As Double
    resultat = Range("A1").Value / Range("B1").Value
    Range("C1").Value = resultat

    ' Copier vers une autre feuille
    Worksheets("Résultats").Range("A1").Value = resultat
End Sub
```

#### ✅ Code robuste (avec prévention)

```vba
Sub CodeRobuste()
    On Error GoTo GestionErreur

    ' Vérifications préventives
    If Range("A1").Value = "" Then
        MsgBox "Veuillez saisir une valeur en A1"
        Exit Sub
    End If

    If Range("B1").Value = 0 Then
        MsgBox "Division par zéro impossible - vérifiez la cellule B1"
        Exit Sub
    End If

    If Not IsNumeric(Range("A1").Value) Or Not IsNumeric(Range("B1").Value) Then
        MsgBox "Les cellules A1 et B1 doivent contenir des nombres"
        Exit Sub
    End If

    ' Vérifier l'existence de la feuille de destination
    If Not FeuilleExiste("Résultats") Then
        If MsgBox("La feuille 'Résultats' n'existe pas. La créer ?", vbYesNo) = vbYes Then
            Worksheets.Add.Name = "Résultats"
        Else
            Exit Sub
        End If
    End If

    ' Maintenant on peut faire le calcul en sécurité
    Dim resultat As Double
    resultat = Range("A1").Value / Range("B1").Value
    Range("C1").Value = resultat
    Worksheets("Résultats").Range("A1").Value = resultat

    MsgBox "Calcul effectué avec succès !"
    Exit Sub

GestionErreur:
    MsgBox "Erreur inattendue : " & Err.Description
End Sub

Function FeuilleExiste(nomFeuille As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(nomFeuille)
    FeuilleExiste = (Err.Number = 0)
    On Error GoTo 0
End Function
```

### Techniques de prévention courantes

#### 1. Validation des données d'entrée

```vba
Function ValiderDonnees(valeur As Variant, typeAttendu As String) As Boolean
    Select Case LCase(typeAttendu)
        Case "nombre"
            ValiderDonnees = IsNumeric(valeur) And valeur <> ""

        Case "texte"
            ValiderDonnees = (VarType(valeur) = vbString) And Len(Trim(CStr(valeur))) > 0

        Case "date"
            ValiderDonnees = IsDate(valeur)

        Case "positif"
            ValiderDonnees = IsNumeric(valeur) And valeur > 0

        Case Else
            ValiderDonnees = False
    End Select
End Function

Sub UtiliserValidation()
    If Not ValiderDonnees(Range("A1").Value, "nombre") Then
        MsgBox "La cellule A1 doit contenir un nombre valide"
        Range("A1").Select
        Exit Sub
    End If

    ' Continuer avec les données validées
    Dim nombre As Double
    nombre = Range("A1").Value
    ' ... reste du traitement
End Sub
```

#### 2. Vérification d'existence des objets

```vba
Function ClasseurOuvert(nomClasseur As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(nomClasseur)
    ClasseurOuvert = (Err.Number = 0)
    On Error GoTo 0
End Function

Function PlagNommeeExiste(nomPlage As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Range(nomPlage)
    PlagNommeeExiste = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub ExempleVerifications()
    ' Vérifier avant d'utiliser
    If Not ClasseurOuvert("Données.xlsx") Then
        MsgBox "Veuillez d'abord ouvrir le fichier Données.xlsx"
        Exit Sub
    End If

    If Not PlagNommeeExiste("ZoneSaisie") Then
        MsgBox "La plage nommée 'ZoneSaisie' n'existe pas"
        Exit Sub
    End If

    ' Maintenant on peut utiliser en sécurité
    Range("ZoneSaisie").Value = "Données traitées"
End Sub
```

---

## Principe 2 : Structure cohérente du code

### Modèle standard de procédure

Adoptez une structure cohérente pour toutes vos procédures avec gestion d'erreurs :

```vba
Sub ModeleStandard()
    '=== DÉCLARATIONS ===
    Dim variable1 As String
    Dim variable2 As Integer
    Dim resultat As Boolean

    '=== INITIALISATION ===
    variable1 = ""
    variable2 = 0
    resultat = False

    '=== VALIDATION DES PRÉREQUIS ===
    If Range("A1").Value = "" Then
        MsgBox "Donnée manquante en A1"
        Exit Sub
    End If

    '=== ACTIVATION GESTION D'ERREUR ===
    On Error GoTo GestionErreur

    '=== TRAITEMENT PRINCIPAL ===
    ' Votre logique métier ici
    variable2 = Range("A1").Value * 2
    variable1 = "Traitement terminé"
    resultat = True

    '=== FINALISATION NORMALE ===
    On Error GoTo 0
    Range("B1").Value = variable1
    MsgBox "Opération réussie"
    Exit Sub

    '=== GESTION D'ERREUR ===
GestionErreur:
    MsgBox "Erreur dans ModeleStandard : " & Err.Description
    On Error GoTo 0
    ' Nettoyage si nécessaire
End Sub
```

### Zones de gestion d'erreur limitées

Ne placez `On Error` que dans les zones réellement risquées :

```vba
Sub ZonesLimitees()
    '=== Zone normale (pas de gestion d'erreur) ===
    Dim donnees As String
    donnees = Range("A1").Value

    If donnees = "" Then
        MsgBox "Aucune donnée à traiter"
        Exit Sub
    End If

    '=== Zone risquée 1 : Ouverture fichier ===
    On Error GoTo ErreurFichier
    Workbooks.Open "C:\Données.xlsx"
    On Error GoTo 0

    '=== Zone normale ===
    Range("B1").Value = "Fichier ouvert"

    '=== Zone risquée 2 : Calcul ===
    On Error GoTo ErreurCalcul
    Range("C1").Value = CDbl(donnees) * 2
    On Error GoTo 0

    '=== Fin normale ===
    MsgBox "Traitement terminé"
    Exit Sub

ErreurFichier:
    MsgBox "Impossible d'ouvrir le fichier"
    Exit Sub

ErreurCalcul:
    MsgBox "Erreur de calcul - vérifiez que A1 contient un nombre"
    Exit Sub
End Sub
```

---

## Principe 3 : Messages d'erreur informatifs

### Messages orientés utilisateur

Vos messages d'erreur doivent être compréhensibles par vos utilisateurs, pas seulement par vous :

#### ❌ Messages techniques peu utiles

```vba
' ÉVITEZ ces messages
MsgBox "Erreur 1004"  
MsgBox Err.Description  ' Souvent cryptique  
MsgBox "Erreur dans la procédure"  
```

#### ✅ Messages informatifs et utiles

```vba
Function CreerMessageErreur(contexte As String, errNum As Long, errDesc As String) As String
    Dim message As String

    message = "Une erreur s'est produite " & contexte & vbCrLf & vbCrLf

    Select Case errNum
        Case 9  ' Subscript out of range
            message = message & "Problème : Un élément demandé n'existe pas" & vbCrLf & _
                     "Solution : Vérifiez les noms de feuilles et de plages"

        Case 11  ' Division by zero
            message = message & "Problème : Division par zéro impossible" & vbCrLf & _
                     "Solution : Vérifiez que le diviseur n'est pas zéro"

        Case 13  ' Type mismatch
            message = message & "Problème : Format de données incorrect" & vbCrLf & _
                     "Solution : Vérifiez que les nombres sont bien des nombres"

        Case 53  ' File not found
            message = message & "Problème : Fichier introuvable" & vbCrLf & _
                     "Solution : Vérifiez le chemin et l'existence du fichier"

        Case Else
            message = message & "Erreur technique : " & errDesc & vbCrLf & _
                     "Code : " & errNum
    End Select

    CreerMessageErreur = message
End Function

Sub UtiliserMessageInformatif()
    On Error GoTo GestionErreur

    Range("A1").Value = 10 / 0
    Exit Sub

GestionErreur:
    MsgBox CreerMessageErreur("lors du calcul", Err.Number, Err.Description), _
           vbExclamation, "Information"
End Sub
```

### Messages avec actions proposées

Proposez des solutions ou des actions à l'utilisateur :

```vba
Sub MessagesAvecActions()
    On Error GoTo GestionErreur

    Worksheets("Données").Range("A1").Value = "Test"
    Exit Sub

GestionErreur:
    If Err.Number = 9 Then  ' Feuille inexistante
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("La feuille 'Données' n'existe pas." & vbCrLf & _
                        "Voulez-vous la créer maintenant ?", _
                        vbYesNoCancel + vbQuestion, "Feuille manquante")

        Select Case reponse
            Case vbYes
                Worksheets.Add.Name = "Données"
                Resume  ' Reprendre à la ligne qui a causé l'erreur

            Case vbNo
                MsgBox "Opération annulée"

            Case vbCancel
                ' Ne rien faire
        End Select
    Else
        MsgBox "Erreur inattendue : " & Err.Description
    End If
End Sub
```

---

## Principe 4 : Journalisation et traçabilité

### Journal simple dans Excel

```vba
Sub AjouterAuJournal(typeEvenement As String, description As String)
    On Error Resume Next

    ' Créer la feuille de journal si elle n'existe pas
    Dim wsJournal As Worksheet
    Set wsJournal = Worksheets("Journal_Erreurs")

    If wsJournal Is Nothing Then
        Set wsJournal = Worksheets.Add
        wsJournal.Name = "Journal_Erreurs"
        ' En-têtes
        wsJournal.Range("A1:D1").Value = Array("Date/Heure", "Type", "Description", "Utilisateur")
        wsJournal.Range("A1:D1").Font.Bold = True
    End If

    ' Ajouter l'entrée
    Dim nouvelleLigne As Long
    nouvelleLigne = wsJournal.Cells(Rows.Count, 1).End(xlUp).Row + 1

    wsJournal.Cells(nouvelleLigne, 1).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    wsJournal.Cells(nouvelleLigne, 2).Value = typeEvenement
    wsJournal.Cells(nouvelleLigne, 3).Value = description
    wsJournal.Cells(nouvelleLigne, 4).Value = Application.UserName

    On Error GoTo 0
End Sub

Sub ExempleAvecJournal()
    On Error GoTo GestionErreur

    AjouterAuJournal "INFO", "Début du traitement"

    ' Code principal
    Range("A1").Value = 10 / 0

    AjouterAuJournal "INFO", "Traitement terminé avec succès"
    Exit Sub

GestionErreur:
    AjouterAuJournal "ERREUR", "Erreur " & Err.Number & ": " & Err.Description
    MsgBox "Une erreur s'est produite. Consultez le journal pour plus de détails."
End Sub
```

### Informations de contexte

Enregistrez des informations utiles pour le débogage :

```vba
Sub TraçabilitéAvancee()
    On Error GoTo GestionErreur

    Dim contexte As String
    Dim etapeActuelle As String

    etapeActuelle = "Initialisation"
    contexte = "Utilisateur: " & Application.UserName & ", Fichier: " & ActiveWorkbook.Name

    etapeActuelle = "Lecture données"
    Dim valeur1 As Variant
    valeur1 = Range("A1").Value
    contexte = contexte & ", Valeur A1: " & valeur1

    etapeActuelle = "Calcul"
    Dim resultat As Double
    resultat = valeur1 / Range("B1").Value

    etapeActuelle = "Finalisation"
    Range("C1").Value = resultat

    Exit Sub

GestionErreur:
    Dim messageDebug As String
    messageDebug = "ERREUR dans l'étape: " & etapeActuelle & vbCrLf & _
                   "Contexte: " & contexte & vbCrLf & _
                   "Erreur: " & Err.Number & " - " & Err.Description

    Debug.Print messageDebug
    AjouterAuJournal "ERREUR", messageDebug

    MsgBox "Erreur dans l'étape: " & etapeActuelle & vbCrLf & _
           "Détails enregistrés dans le journal"
End Sub
```

---

## Principe 5 : Nettoyage et libération des ressources

### Libération systématique des objets

```vba
Sub LibérationRessources()
    ' Déclarations
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fichierOuvert As Boolean

    Set wb = Nothing
    Set ws = Nothing
    fichierOuvert = False

    On Error GoTo GestionErreur

    ' Ouverture fichier
    Set wb = Workbooks.Open("C:\Données.xlsx")
    fichierOuvert = True

    Set ws = wb.Worksheets("Import")

    ' Traitement
    ws.Range("A1").Value = "Traitement en cours"

    ' Nettoyage normal
    wb.Close SaveChanges:=True
    Set wb = Nothing
    Set ws = Nothing
    fichierOuvert = False

    MsgBox "Traitement terminé"
    Exit Sub

GestionErreur:
    ' Nettoyage en cas d'erreur
    On Error Resume Next  ' Éviter les erreurs en cascade pendant le nettoyage

    If fichierOuvert And Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    Set wb = Nothing
    Set ws = Nothing

    On Error GoTo 0
    MsgBox "Erreur : " & Err.Description & vbCrLf & "Ressources libérées"
End Sub
```

### Pattern de nettoyage avec Finally

```vba
Sub PatternFinally()
    ' Variables de ressources
    Dim wb As Workbook
    Dim fichierModifie As Boolean
    Dim erreurSurvenue As Boolean

    ' Initialisation
    Set wb = Nothing
    fichierModifie = False
    erreurSurvenue = False

    On Error GoTo GestionErreur

    '=== TRY - Code principal ===
    Set wb = Workbooks.Open("C:\Test.xlsx")
    wb.Worksheets(1).Range("A1").Value = "Modification"
    fichierModifie = True

    ' Simulation d'erreur pour test
    ' Dim test As Double: test = 1 / 0

    GoTo Finally

GestionErreur:
    '=== CATCH - Gestion d'erreur ===
    erreurSurvenue = True
    MsgBox "Erreur rencontrée : " & Err.Description

Finally:
    '=== FINALLY - Nettoyage obligatoire ===
    On Error Resume Next

    If Not wb Is Nothing Then
        If fichierModifie And Not erreurSurvenue Then
            wb.Save
            Debug.Print "Fichier sauvegardé"
        End If
        wb.Close
        Debug.Print "Fichier fermé"
    End If

    Set wb = Nothing
    On Error GoTo 0

    If Not erreurSurvenue Then
        MsgBox "Opération terminée avec succès"
    End If
End Sub
```

---

## Principe 6 : Code défensif

### Validation des paramètres de fonction

```vba
Function CalculerPourcentage(valeur As Variant, total As Variant) As Variant
    ' Validation défensive des paramètres
    If Not IsNumeric(valeur) Then
        CalculerPourcentage = "Erreur: Valeur non numérique"
        Exit Function
    End If

    If Not IsNumeric(total) Then
        CalculerPourcentage = "Erreur: Total non numérique"
        Exit Function
    End If

    If total = 0 Then
        CalculerPourcentage = "Erreur: Division par zéro"
        Exit Function
    End If

    If total < 0 Then
        CalculerPourcentage = "Erreur: Total négatif"
        Exit Function
    End If

    ' Calcul sécurisé
    CalculerPourcentage = (valeur / total) * 100
End Function

Sub UtiliserCalculPourcentage()
    Dim resultat As Variant

    resultat = CalculerPourcentage(Range("A1").Value, Range("B1").Value)

    If IsNumeric(resultat) Then
        Range("C1").Value = resultat
        MsgBox "Pourcentage calculé : " & Format(resultat, "0.00") & "%"
    Else
        Range("C1").Value = resultat  ' Afficher le message d'erreur
        MsgBox resultat
    End If
End Sub
```

### Vérifications d'état système

```vba
Function VerifierEnvironnement() As Boolean
    Dim problemes As String

    ' Vérifier la version d'Excel
    If Val(Application.Version) < 14 Then  ' Excel 2010 = version 14
        problemes = problemes & "- Version Excel trop ancienne" & vbCrLf
    End If

    ' Vérifier que des classeurs sont ouverts
    If Workbooks.Count = 0 Then
        problemes = problemes & "- Aucun classeur ouvert" & vbCrLf
    End If

    ' Vérifier que la feuille active n'est pas protégée
    If ActiveSheet.ProtectContents Then
        problemes = problemes & "- Feuille protégée" & vbCrLf
    End If

    ' Retourner le résultat
    If problemes = "" Then
        VerifierEnvironnement = True
    Else
        MsgBox "Problèmes détectés :" & vbCrLf & problemes
        VerifierEnvironnement = False
    End If
End Function

Sub MacroAvecVerification()
    ' Vérification préalable
    If Not VerifierEnvironnement() Then
        Exit Sub
    End If

    ' Code principal
    MsgBox "Environnement OK - exécution de la macro"
End Sub
```

---

## Principe 7 : Tests et robustesse

### Fonction de test automatisé

```vba
Sub TesterGestionErreurs()
    Dim testsReussis As Integer
    Dim testsTotal As Integer

    testsTotal = 0
    testsReussis = 0

    ' Test 1 : Division par zéro
    testsTotal = testsTotal + 1
    If TesterDivisionParZero() Then testsReussis = testsReussis + 1

    ' Test 2 : Feuille inexistante
    testsTotal = testsTotal + 1
    If TesterFeuilleInexistante() Then testsReussis = testsReussis + 1

    ' Test 3 : Type incorrect
    testsTotal = testsTotal + 1
    If TesterTypeIncorrect() Then testsReussis = testsReussis + 1

    ' Résultats
    MsgBox "Tests terminés : " & testsReussis & "/" & testsTotal & " réussis"
End Sub

Function TesterDivisionParZero() As Boolean
    On Error GoTo ErreurAttendue

    Dim resultat As Double
    resultat = 10 / 0

    ' Si on arrive ici, le test a échoué (pas d'erreur détectée)
    TesterDivisionParZero = False
    Debug.Print "ÉCHEC : Division par zéro non détectée"
    Exit Function

ErreurAttendue:
    If Err.Number = 11 Then
        TesterDivisionParZero = True
        Debug.Print "SUCCÈS : Division par zéro correctement détectée"
    Else
        TesterDivisionParZero = False
        Debug.Print "ÉCHEC : Erreur inattendue " & Err.Number
    End If
End Function

Function TesterFeuilleInexistante() As Boolean
    On Error GoTo ErreurAttendue

    Worksheets("FeuilleTestInexistante").Range("A1").Value = "Test"

    ' Si on arrive ici, le test a échoué
    TesterFeuilleInexistante = False
    Debug.Print "ÉCHEC : Accès feuille inexistante non détecté"
    Exit Function

ErreurAttendue:
    If Err.Number = 9 Then
        TesterFeuilleInexistante = True
        Debug.Print "SUCCÈS : Feuille inexistante correctement détectée"
    Else
        TesterFeuilleInexistante = False
        Debug.Print "ÉCHEC : Erreur inattendue " & Err.Number
    End If
End Function

Function TesterTypeIncorrect() As Boolean
    On Error GoTo ErreurAttendue

    Dim nombre As Integer
    nombre = "Pas un nombre"

    ' Si on arrive ici, le test a échoué
    TesterTypeIncorrect = False
    Debug.Print "ÉCHEC : Type incorrect non détecté"
    Exit Function

ErreurAttendue:
    If Err.Number = 13 Then
        TesterTypeIncorrect = True
        Debug.Print "SUCCÈS : Type incorrect correctement détecté"
    Else
        TesterTypeIncorrect = False
        Debug.Print "ÉCHEC : Erreur inattendue " & Err.Number
    End If
End Function
```

---

## Liste de contrôle des bonnes pratiques

### ✅ Checklist pour chaque procédure

Avant de considérer votre code comme terminé, vérifiez :

#### **Prévention**
- [ ] Les données d'entrée sont-elles validées ?
- [ ] L'existence des objets est-elle vérifiée ?
- [ ] Les divisions par zéro sont-elles évitées ?
- [ ] Les types de données sont-ils appropriés ?

#### **Structure**
- [ ] Y a-t-il un `Exit Sub/Function` avant la section d'erreur ?
- [ ] La gestion d'erreur est-elle limitée aux zones risquées ?
- [ ] Les variables sont-elles initialisées ?
- [ ] Le code est-il bien commenté ?

#### **Messages**
- [ ] Les messages d'erreur sont-ils compréhensibles ?
- [ ] Des solutions sont-elles proposées à l'utilisateur ?
- [ ] Les erreurs sont-elles journalisées pour le débogage ?

#### **Nettoyage**
- [ ] Les ressources sont-elles libérées en cas d'erreur ?
- [ ] Les fichiers ouverts sont-ils fermés ?
- [ ] Les variables objet sont-elles remises à Nothing ?

#### **Tests**
- [ ] Le code a-t-il été testé avec des données incorrectes ?
- [ ] Les différents types d'erreur ont-ils été simulés ?
- [ ] Le comportement en cas d'erreur est-il prévisible ?

---

## Récapitulatif des bonnes pratiques

### Les 7 principes fondamentaux

1. **PRÉVENTION** : Validez avant d'exécuter
2. **STRUCTURE** : Adoptez un modèle cohérent
3. **MESSAGES** : Informez clairement l'utilisateur
4. **JOURNALISATION** : Enregistrez pour déboguer
5. **NETTOYAGE** : Libérez les ressources
6. **DÉFENSIF** : Anticipez tous les cas
7. **TESTS** : Validez votre gestion d'erreurs

### Évolution progressive

#### **Niveau débutant**
- Utilisez `On Error Resume Next` pour les vérifications simples
- Ajoutez des messages d'erreur de base
- Validez les données critiques

#### **Niveau intermédiaire**
- Passez à `On Error GoTo` pour plus de contrôle
- Implémentez la journalisation
- Créez des fonctions de validation réutilisables

#### **Niveau avancé**
- Développez un système de gestion centralisé
- Implémentez des tests automatisés
- Créez des erreurs personnalisées métier

### Conseil final

La gestion d'erreurs robuste ne s'apprend pas en un jour. Commencez par appliquer les principes de base, puis enrichissez progressivement vos techniques. Chaque erreur que vous gérez correctement est un pas vers un code plus professionnel et plus fiable.

**Rappelez-vous** : un code qui gère bien les erreurs est la marque d'un développeur qui pense à ses utilisateurs et à la maintenance future de son application.

⏭️
