🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.3 Instructions de contrôle (Exit, GoTo)

## Introduction

Les **instructions de contrôle** permettent de modifier le flux normal d'exécution de votre programme. Elles vous donnent la possibilité de "sortir" prématurément d'une boucle ou d'une procédure, ou de "sauter" à une autre partie du code. Ces outils sont puissants mais doivent être utilisés avec parcimonie pour maintenir un code lisible et maintenable.

### Analogie du parcours en voiture

Imaginez que vous conduisez sur une route :
- **Exit** = Prendre une sortie d'autoroute pour quitter rapidement
- **GoTo** = Faire un détour en sautant directement à une autre route

Ces instructions changent votre itinéraire prévu, ce qui peut être utile dans certaines situations mais peut aussi créer de la confusion si utilisé trop souvent.

## L'instruction Exit

### Concept et utilité

`Exit` permet de **sortir immédiatement** d'une structure (boucle, procédure, fonction) avant sa fin normale. C'est comme appuyer sur un bouton d'urgence qui vous fait sortir instantanément.

### Types d'Exit disponibles

- `Exit Sub` : Sortir d'une procédure
- `Exit Function` : Sortir d'une fonction
- `Exit For` : Sortir d'une boucle For
- `Exit Do` : Sortir d'une boucle Do

## Exit Sub (Sortir d'une procédure)

### Usage basique

```vba
Sub VerifierConditions()
    Dim age As Integer
    age = InputBox("Votre âge ?")

    ' Vérification immédiate avec sortie
    If age < 0 Then
        MsgBox "Âge invalide !"
        Exit Sub  ' Sort immédiatement de la procédure
    End If

    ' Ce code ne s'exécute que si age >= 0
    MsgBox "Âge valide : " & age & " ans"

    ' Autres traitements...
    Range("A1").Value = "Utilisateur de " & age & " ans"
End Sub
```

### Validation d'entrée avec Exit Sub

```vba
Sub TraiterCommande()
    Dim commande As String
    commande = InputBox("Entrez une commande (NOUVEAU/OUVRIR/FERMER) :")

    ' Vérifications avec sorties anticipées
    If commande = "" Then
        MsgBox "Aucune commande saisie"
        Exit Sub
    End If

    If LCase(commande) <> "nouveau" And LCase(commande) <> "ouvrir" And LCase(commande) <> "fermer" Then
        MsgBox "Commande non reconnue : " & commande
        Exit Sub
    End If

    ' Traitement principal (seulement si les validations passent)
    Select Case LCase(commande)
        Case "nouveau"
            MsgBox "Création d'un nouveau document"
        Case "ouvrir"
            MsgBox "Ouverture d'un document"
        Case "fermer"
            MsgBox "Fermeture du document"
    End Select
End Sub
```

### Exit Sub vs conditions imbriquées

```vba
' ❌ Sans Exit Sub - Conditions imbriquées complexes
Sub ExempleSansExit()
    Dim fichier As String
    fichier = InputBox("Nom du fichier :")

    If fichier <> "" Then
        If Len(fichier) > 3 Then
            If Right(fichier, 4) = ".txt" Then
                ' Traitement principal très indenté
                MsgBox "Fichier valide : " & fichier
                Range("A1").Value = fichier
            Else
                MsgBox "Le fichier doit avoir l'extension .txt"
            End If
        Else
            MsgBox "Le nom de fichier est trop court"
        End If
    Else
        MsgBox "Nom de fichier requis"
    End If
End Sub

' ✅ Avec Exit Sub - Plus lisible
Sub ExempleAvecExit()
    Dim fichier As String
    fichier = InputBox("Nom du fichier :")

    ' Validations avec sorties anticipées
    If fichier = "" Then
        MsgBox "Nom de fichier requis"
        Exit Sub
    End If

    If Len(fichier) <= 3 Then
        MsgBox "Le nom de fichier est trop court"
        Exit Sub
    End If

    If Right(fichier, 4) <> ".txt" Then
        MsgBox "Le fichier doit avoir l'extension .txt"
        Exit Sub
    End If

    ' Traitement principal (sans indentation excessive)
    MsgBox "Fichier valide : " & fichier
    Range("A1").Value = fichier
End Sub
```

## Exit Function (Sortir d'une fonction)

### Retour anticipé avec valeur

```vba
Function CalculerRemise(montant As Double, typeClient As String) As Double
    ' Validations avec retours anticipés
    If montant <= 0 Then
        MsgBox "Montant invalide"
        CalculerRemise = 0
        Exit Function
    End If

    If typeClient = "" Then
        MsgBox "Type de client requis"
        CalculerRemise = 0
        Exit Function
    End If

    ' Calcul principal
    Select Case LCase(typeClient)
        Case "vip"
            CalculerRemise = montant * 0.15
        Case "regulier"
            CalculerRemise = montant * 0.05
        Case "nouveau"
            CalculerRemise = montant * 0.10
        Case Else
            MsgBox "Type de client inconnu"
            CalculerRemise = 0
            Exit Function
    End Select
End Function
```

### Fonction avec gestion d'erreur

```vba
Function DivisionSecurisee(dividende As Double, diviseur As Double) As Variant
    ' Vérification de division par zéro
    If diviseur = 0 Then
        DivisionSecurisee = "Erreur : Division par zéro"
        Exit Function
    End If

    ' Vérification de très petites valeurs
    If Abs(diviseur) < 0.000001 Then
        DivisionSecurisee = "Erreur : Diviseur trop petit"
        Exit Function
    End If

    ' Calcul normal
    DivisionSecurisee = dividende / diviseur
End Function

Sub TestDivision()
    MsgBox DivisionSecurisee(10, 2)     ' 5
    MsgBox DivisionSecurisee(10, 0)     ' Erreur : Division par zéro
    MsgBox DivisionSecurisee(10, 0.0000001)  ' Erreur : Diviseur trop petit
End Sub
```

## Exit For (Sortir d'une boucle For)

### Recherche avec arrêt anticipé

```vba
Sub ChercherValeur()
    Dim valeurCherchee As String
    Dim i As Integer
    Dim trouve As Boolean

    valeurCherchee = InputBox("Valeur à chercher :")

    For i = 1 To 1000  ' Chercher dans les 1000 premières lignes
        If Range("A" & i).Value = valeurCherchee Then
            MsgBox "Valeur trouvée à la ligne " & i
            Range("A" & i).Select
            trouve = True
            Exit For  ' Arrêter la recherche dès qu'on trouve
        End If
    Next i

    If Not trouve Then
        MsgBox "Valeur non trouvée"
    End If
End Sub
```

### Traitement avec limite d'erreurs

```vba
Sub TraiterDonneesAvecLimite()
    Dim i As Integer
    Dim erreurs As Integer
    Dim maxErreurs As Integer

    maxErreurs = 5

    For i = 1 To 100
        ' Simuler un traitement qui peut échouer
        If Rnd() < 0.1 Then  ' 10% de chance d'erreur
            erreurs = erreurs + 1
            MsgBox "Erreur lors du traitement de la ligne " & i

            ' Arrêter si trop d'erreurs
            If erreurs >= maxErreurs Then
                MsgBox "Trop d'erreurs (" & erreurs & "). Arrêt du traitement."
                Exit For
            End If
        Else
            Range("A" & i).Value = "Ligne " & i & " - OK"
        End If
    Next i

    MsgBox "Traitement terminé. Erreurs : " & erreurs
End Sub
```

### Boucles imbriquées avec Exit

```vba
Sub ChercherDansTableau()
    Dim ligne As Integer, colonne As Integer
    Dim valeurCherchee As String
    Dim trouve As Boolean

    valeurCherchee = InputBox("Valeur à chercher :")

    For ligne = 1 To 50
        For colonne = 1 To 10
            If Cells(ligne, colonne).Value = valeurCherchee Then
                MsgBox "Trouvé en ligne " & ligne & ", colonne " & colonne
                Cells(ligne, colonne).Select
                trouve = True
                Exit For  ' Sort de la boucle interne
            End If
        Next colonne

        If trouve Then Exit For  ' Sort de la boucle externe
    Next ligne

    If Not trouve Then
        MsgBox "Valeur non trouvée dans le tableau"
    End If
End Sub
```

## Exit Do (Sortir d'une boucle Do)

### Saisie utilisateur avec abandon

```vba
Sub SaisieAvecAbandon()
    Dim reponse As String
    Dim tentatives As Integer

    Do
        tentatives = tentatives + 1
        reponse = InputBox("Entrez 'OK' pour continuer (tentative " & tentatives & "/3) :")

        ' Permettre à l'utilisateur d'annuler
        If reponse = "" Then
            MsgBox "Opération annulée par l'utilisateur"
            Exit Do
        End If

        ' Vérifier la réponse
        If LCase(reponse) = "ok" Then
            MsgBox "Parfait ! Vous pouvez continuer."
            Exit Do
        End If

        ' Limiter les tentatives
        If tentatives >= 3 Then
            MsgBox "Trop de tentatives. Abandon."
            Exit Do
        End If

        MsgBox "Réponse incorrecte. Réessayez."
    Loop
End Sub
```

### Traitement de fichier avec interruption

```vba
Sub TraiterFichierAvecInterruption()
    Dim ligne As Integer
    Dim donnee As String

    ligne = 1

    Do
        donnee = Range("A" & ligne).Value

        ' Arrêter si cellule vide (fin des données)
        If donnee = "" Then
            MsgBox "Fin des données atteinte à la ligne " & ligne
            Exit Do
        End If

        ' Arrêter si marqueur spécial trouvé
        If donnee = "STOP" Then
            MsgBox "Marqueur STOP trouvé. Arrêt du traitement."
            Exit Do
        End If

        ' Traitement normal
        Range("B" & ligne).Value = "Traité : " & donnee
        ligne = ligne + 1

        ' Sécurité : éviter boucle infinie
        If ligne > 10000 Then
            MsgBox "Limite de sécurité atteinte (10000 lignes)"
            Exit Do
        End If
    Loop
End Sub
```

## L'instruction GoTo

### Concept et controverses

`GoTo` permet de "sauter" directement à une autre ligne du code, identifiée par une **étiquette**. Cette instruction est controversée car elle peut rendre le code difficile à suivre et à maintenir.

### Syntaxe de base

```vba
Sub ExempleGoTo()
    MsgBox "Début"
    GoTo EtiquetteTest
    MsgBox "Cette ligne ne sera jamais exécutée"

EtiquetteTest:
    MsgBox "Arrivé à l'étiquette"
End Sub
```

### Usage acceptable : Gestion d'erreur simple

```vba
Sub ExempleGestionErreur()
    Dim fichier As String

    fichier = InputBox("Nom de fichier :")

    If fichier = "" Then GoTo ErreurFichier
    If Len(fichier) < 3 Then GoTo ErreurFichier
    If InStr(fichier, ".") = 0 Then GoTo ErreurFichier

    ' Traitement normal
    MsgBox "Fichier valide : " & fichier
    Range("A1").Value = fichier
    GoTo Fin

ErreurFichier:
    MsgBox "Erreur : Nom de fichier invalide"

Fin:
    MsgBox "Procédure terminée"
End Sub
```

### Simulation de Continue (dans les boucles)

```vba
Sub SimulerContinue()
    Dim i As Integer

    For i = 1 To 10
        ' Sauter les nombres pairs
        If i Mod 2 = 0 Then GoTo SuivantIteration

        ' Traiter seulement les nombres impairs
        MsgBox "Nombre impair : " & i
        Range("A" & ((i + 1) / 2)).Value = i

SuivantIteration:
    Next i
End Sub
```

## Pourquoi éviter GoTo ?

### Problème de lisibilité

```vba
' ❌ Code difficile à suivre avec GoTo
Sub ExempleProblematique()
    Dim x As Integer
    x = 1

Debut:
    If x > 5 Then GoTo Fin
    If x = 3 Then GoTo Special
    Range("A" & x).Value = x
    x = x + 1
    GoTo Debut

Special:
    Range("A" & x).Value = "SPECIAL"
    x = x + 1
    GoTo Debut

Fin:
    MsgBox "Terminé"
End Sub

' ✅ Code équivalent plus clair sans GoTo
Sub ExempleClaire()
    Dim x As Integer

    For x = 1 To 5
        If x = 3 Then
            Range("A" & x).Value = "SPECIAL"
        Else
            Range("A" & x).Value = x
        End If
    Next x

    MsgBox "Terminé"
End Sub
```

## Alternatives modernes à GoTo

### Utiliser des structures conditionnelles

```vba
' ❌ Avec GoTo
Sub AncienneMethode()
    Dim age As Integer
    age = InputBox("Age ?")

    If age < 0 Then GoTo ErreurAge
    If age > 150 Then GoTo ErreurAge

    MsgBox "Age valide"
    GoTo Fin

ErreurAge:
    MsgBox "Age invalide"

Fin:  
End Sub  

' ✅ Avec If...Else
Sub MethodeModerne()
    Dim age As Integer
    age = InputBox("Age ?")

    If age < 0 Or age > 150 Then
        MsgBox "Age invalide"
    Else
        MsgBox "Age valide"
    End If
End Sub
```

### Utiliser des procédures séparées

```vba
' ❌ Avec GoTo et code mélangé
Sub TraitementComplexeAvecGoTo()
    Dim etape As Integer
    etape = InputBox("Étape (1, 2, ou 3) ?")

    If etape = 1 Then GoTo Etape1
    If etape = 2 Then GoTo Etape2
    If etape = 3 Then GoTo Etape3
    GoTo Erreur

Etape1:
    MsgBox "Traitement étape 1"
    Range("A1").Value = "Étape 1 terminée"
    GoTo Fin

Etape2:
    MsgBox "Traitement étape 2"
    Range("A1").Value = "Étape 2 terminée"
    GoTo Fin

Etape3:
    MsgBox "Traitement étape 3"
    Range("A1").Value = "Étape 3 terminée"
    GoTo Fin

Erreur:
    MsgBox "Étape invalide"

Fin:  
End Sub  

' ✅ Avec procédures séparées
Sub TraitementModerne()
    Dim etape As Integer
    etape = InputBox("Étape (1, 2, ou 3) ?")

    Select Case etape
        Case 1
            ExecuterEtape1
        Case 2
            ExecuterEtape2
        Case 3
            ExecuterEtape3
        Case Else
            MsgBox "Étape invalide"
    End Select
End Sub

Sub ExecuterEtape1()
    MsgBox "Traitement étape 1"
    Range("A1").Value = "Étape 1 terminée"
End Sub

Sub ExecuterEtape2()
    MsgBox "Traitement étape 2"
    Range("A1").Value = "Étape 2 terminée"
End Sub

Sub ExecuterEtape3()
    MsgBox "Traitement étape 3"
    Range("A1").Value = "Étape 3 terminée"
End Sub
```

## Exemple pratique complet

### Système de validation robuste

```vba
Sub SystemeValidationComplete()
    ' Exemple combinant Exit et bonnes pratiques

    Dim nom As String, email As String, age As Integer

    ' Saisie et validation du nom
    nom = InputBox("Votre nom :")
    If nom = "" Then
        MsgBox "Le nom est obligatoire"
        Exit Sub
    End If

    If Len(nom) < 2 Then
        MsgBox "Le nom doit contenir au moins 2 caractères"
        Exit Sub
    End If

    ' Saisie et validation de l'email
    email = InputBox("Votre email :")
    If email = "" Then
        MsgBox "L'email est obligatoire"
        Exit Sub
    End If

    If InStr(email, "@") = 0 Or InStr(email, ".") = 0 Then
        MsgBox "Format d'email invalide"
        Exit Sub
    End If

    ' Saisie et validation de l'âge
    Dim ageTexte As String
    ageTexte = InputBox("Votre âge :")

    If Not IsNumeric(ageTexte) Then
        MsgBox "L'âge doit être un nombre"
        Exit Sub
    End If

    age = CInt(ageTexte)
    If age < 0 Or age > 150 Then
        MsgBox "Âge invalide (doit être entre 0 et 150)"
        Exit Sub
    End If

    ' Si toutes les validations passent, enregistrer les données
    Range("A1").Value = "Nom:"
    Range("B1").Value = nom
    Range("A2").Value = "Email:"
    Range("B2").Value = email
    Range("A3").Value = "Âge:"
    Range("B3").Value = age

    MsgBox "Données enregistrées avec succès !"
End Sub
```

## Bonnes pratiques pour les instructions de contrôle

### 1. Préférer Exit aux conditions imbriquées

```vba
' ✅ Plus lisible avec Exit
Sub BonnePratique()
    If condition1 = False Then Exit Sub
    If condition2 = False Then Exit Sub
    If condition3 = False Then Exit Sub

    ' Code principal sans indentation excessive
    ExecuterTraitementPrincipal
End Sub
```

### 2. Toujours documenter les Exit

```vba
Sub ProcedureDocumentee()
    Dim fichier As String
    fichier = InputBox("Fichier :")

    ' Validation : nom de fichier obligatoire
    If fichier = "" Then
        MsgBox "Nom de fichier requis"
        Exit Sub  ' Arrêt si pas de nom
    End If

    ' Validation : extension correcte
    If Right(fichier, 4) <> ".txt" Then
        MsgBox "Extension .txt requise"
        Exit Sub  ' Arrêt si mauvaise extension
    End If

    ' Traitement principal...
End Sub
```

### 3. Éviter GoTo sauf cas exceptionnels

```vba
' ✅ GoTo acceptable pour gestion d'erreur simple
Sub GestionErreurAcceptable()
    On Error GoTo GestionErreur

    ' Code qui peut générer une erreur
    Workbooks.Open "fichier_inexistant.xlsx"
    Exit Sub

GestionErreur:
    MsgBox "Erreur : " & Err.Description
End Sub
```

### 4. Utiliser des noms d'étiquettes explicites

```vba
' ❌ Étiquettes peu claires
Sub MauvaisesEtiquettes()
    If erreur Then GoTo A
    ' ...
A:
    MsgBox "Erreur"
End Sub

' ✅ Étiquettes explicites
Sub BonnesEtiquettes()
    If erreur Then GoTo GestionErreur
    ' ...

GestionErreur:
    MsgBox "Erreur détectée"
End Sub
```

## Récapitulatif des concepts clés

1. **Exit Sub/Function** : Sortie anticipée pour éviter l'indentation excessive
2. **Exit For/Do** : Arrêt de boucle quand la condition est remplie
3. **GoTo** : À éviter en général, acceptable pour gestion d'erreur simple
4. **Validation** : Utiliser Exit pour valider les entrées en début de procédure
5. **Lisibilité** : Préférer les structures modernes aux sauts de code
6. **Documentation** : Commenter les Exit pour expliquer pourquoi
7. **Alternatives** : Utiliser des procédures séparées plutôt que GoTo

Les instructions de contrôle sont des outils puissants qui, utilisés avec sagesse, peuvent considérablement améliorer la lisibilité et la robustesse de votre code. L'objectif est toujours de créer un code facile à comprendre et à maintenir !

⏭️ [6. Manipulation des objets Excel](/06-manipulation-objets-excel/)
