🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.1 Instructions conditionnelles

## Introduction

Les **instructions conditionnelles** permettent à votre programme de prendre des décisions et d'exécuter différents blocs de code selon que certaines conditions sont vraies ou fausses. C'est ce qui rend votre code "intelligent" en lui permettant de s'adapter aux différentes situations.

### Analogie du carrefour

Imaginez que vous arrivez à un carrefour :
- **SI** le feu est vert, **ALORS** vous avancez
- **SINON SI** le feu est orange, **ALORS** vous ralentissez
- **SINON** (feu rouge), vous vous arrêtez

C'est exactement ce que font les instructions conditionnelles : elles examinent une situation (la condition) et choisissent l'action appropriée.

## 5.1.1 If...Then...Else

### Structure de base

La structure `If...Then...Else` est la condition la plus fondamentale en programmation.

```vba
If condition Then
    ' Instructions à exécuter si la condition est vraie
Else
    ' Instructions à exécuter si la condition est fausse
End If
```

### Exemple simple

```vba
Sub ExempleSimple()
    Dim age As Integer
    age = InputBox("Quel est votre âge ?")

    If age >= 18 Then
        MsgBox "Vous êtes majeur !"
    Else
        MsgBox "Vous êtes mineur."
    End If
End Sub
```

### If simple (sans Else)

Parfois, vous n'avez besoin d'agir que si une condition est vraie :

```vba
Sub VerifierCelluleVide()
    If Range("A1").Value = "" Then
        MsgBox "La cellule A1 est vide !"
    End If

    ' Le programme continue ici dans tous les cas
    MsgBox "Vérification terminée"
End Sub
```

### If sur une seule ligne

Pour des actions simples, vous pouvez écrire tout sur une ligne :

```vba
Sub ExempleUneLigne()
    Dim nombre As Integer
    nombre = 15

    If nombre > 10 Then MsgBox "Nombre supérieur à 10"
    If nombre < 0 Then Range("A1").Value = "Négatif"
End Sub
```

### Conditions multiples avec ElseIf

Pour tester plusieurs conditions successives :

```vba
Sub CategoriserAge()
    Dim age As Integer
    age = InputBox("Votre âge ?")

    If age < 0 Then
        MsgBox "Âge invalide !"
    ElseIf age < 13 Then
        MsgBox "Vous êtes un enfant"
    ElseIf age < 18 Then
        MsgBox "Vous êtes un adolescent"
    ElseIf age < 65 Then
        MsgBox "Vous êtes un adulte"
    Else
        MsgBox "Vous êtes un senior"
    End If
End Sub
```

### Conditions complexes avec And, Or, Not

#### Utilisation de And (ET)
```vba
Sub VerifierAcces()
    Dim age As Integer
    Dim permis As String

    age = InputBox("Votre âge ?")
    permis = InputBox("Avez-vous le permis ? (oui/non)")

    If age >= 18 And LCase(permis) = "oui" Then
        MsgBox "Vous pouvez conduire !"
    Else
        MsgBox "Vous ne pouvez pas conduire."
    End If
End Sub
```

#### Utilisation de Or (OU)
```vba
Sub VerifierJourWeekend()
    Dim jour As String
    jour = Format(Date, "dddd")

    If jour = "samedi" Or jour = "dimanche" Then
        MsgBox "C'est le weekend !"
    Else
        MsgBox "C'est un jour de semaine."
    End If
End Sub
```

#### Utilisation de Not (NON)
```vba
Sub VerifierCelluleNonVide()
    If Not Range("A1").Value = "" Then
        MsgBox "La cellule A1 contient : " & Range("A1").Value
    Else
        MsgBox "La cellule A1 est vide"
    End If
End Sub
```

### Conditions imbriquées

Vous pouvez mettre des conditions à l'intérieur d'autres conditions :

```vba
Sub CalculerRemise()
    Dim montant As Double
    Dim estMembre As String
    Dim remise As Double

    montant = InputBox("Montant de l'achat ?")
    estMembre = InputBox("Êtes-vous membre ? (oui/non)")

    If montant > 100 Then
        If LCase(estMembre) = "oui" Then
            remise = montant * 0.15  ' 15% pour les membres
            MsgBox "Remise membre : " & remise & "€"
        Else
            remise = montant * 0.10  ' 10% pour les non-membres
            MsgBox "Remise standard : " & remise & "€"
        End If
    Else
        MsgBox "Pas de remise pour les achats inférieurs à 100€"
    End If
End Sub
```

### Exemples pratiques avec Excel

#### Formater selon le contenu
```vba
Sub FormaterSelonValeur()
    Dim cellule As Range
    Set cellule = Range("A1")

    If IsNumeric(cellule.Value) Then
        If cellule.Value > 0 Then
            cellule.Font.Color = RGB(0, 128, 0)  ' Vert pour positif
        ElseIf cellule.Value < 0 Then
            cellule.Font.Color = RGB(255, 0, 0)  ' Rouge pour négatif
        Else
            cellule.Font.Color = RGB(0, 0, 0)    ' Noir pour zéro
        End If
    Else
        cellule.Font.Color = RGB(128, 128, 128)  ' Gris pour texte
    End If
End Sub
```

#### Valider des données
```vba
Sub ValiderEmail()
    Dim email As String
    email = Range("B2").Value

    If email = "" Then
        MsgBox "L'email ne peut pas être vide !"
        Range("B2").Select
    ElseIf InStr(email, "@") = 0 Then
        MsgBox "L'email doit contenir un @"
        Range("B2").Select
    ElseIf InStr(email, ".") = 0 Then
        MsgBox "L'email doit contenir un point"
        Range("B2").Select
    Else
        MsgBox "Email valide !"
        Range("B2").Font.Color = RGB(0, 128, 0)  ' Vert
    End If
End Sub
```

## 5.1.2 Select Case

### Quand utiliser Select Case ?

Quand vous devez comparer une même variable à plusieurs valeurs différentes, `Select Case` est plus lisible que plusieurs `ElseIf` :

```vba
' ❌ Difficile à lire avec ElseIf
If jour = 1 Then
    MsgBox "Lundi"
ElseIf jour = 2 Then
    MsgBox "Mardi"
ElseIf jour = 3 Then
    MsgBox "Mercredi"
' ... etc

' ✅ Plus clair avec Select Case
Select Case jour
    Case 1
        MsgBox "Lundi"
    Case 2
        MsgBox "Mardi"
    Case 3
        MsgBox "Mercredi"
End Select
```

### Structure de base

```vba
Select Case variable_à_tester
    Case valeur1
        ' Instructions pour valeur1
    Case valeur2
        ' Instructions pour valeur2
    Case Else
        ' Instructions par défaut
End Select
```

### Exemple simple

```vba
Sub AfficherJourSemaine()
    Dim numeroJour As Integer
    numeroJour = Weekday(Date)

    Select Case numeroJour
        Case 1
            MsgBox "Aujourd'hui c'est dimanche"
        Case 2
            MsgBox "Aujourd'hui c'est lundi"
        Case 3
            MsgBox "Aujourd'hui c'est mardi"
        Case 4
            MsgBox "Aujourd'hui c'est mercredi"
        Case 5
            MsgBox "Aujourd'hui c'est jeudi"
        Case 6
            MsgBox "Aujourd'hui c'est vendredi"
        Case 7
            MsgBox "Aujourd'hui c'est samedi"
        Case Else
            MsgBox "Jour invalide"
    End Select
End Sub
```

### Plusieurs valeurs pour un même cas

```vba
Sub CategoriserMois()
    Dim mois As Integer
    mois = Month(Date)

    Select Case mois
        Case 12, 1, 2
            MsgBox "C'est l'hiver"
        Case 3, 4, 5
            MsgBox "C'est le printemps"
        Case 6, 7, 8
            MsgBox "C'est l'été"
        Case 9, 10, 11
            MsgBox "C'est l'automne"
        Case Else
            MsgBox "Mois invalide"
    End Select
End Sub
```

### Plages de valeurs avec To

```vba
Sub CategoriserNote()
    Dim note As Integer
    note = InputBox("Entrez votre note sur 20 :")

    Select Case note
        Case 0 To 7
            MsgBox "Insuffisant"
        Case 8 To 9
            MsgBox "Passable"
        Case 10 To 11
            MsgBox "Assez bien"
        Case 12 To 13
            MsgBox "Bien"
        Case 14 To 15
            MsgBox "Très bien"
        Case 16 To 20
            MsgBox "Excellent"
        Case Else
            MsgBox "Note invalide (doit être entre 0 et 20)"
    End Select
End Sub
```

### Conditions avec Is

```vba
Sub AnalyserTemperature()
    Dim temperature As Integer
    temperature = InputBox("Température actuelle ?")

    Select Case temperature
        Case Is < 0
            MsgBox "Il gèle !"
        Case Is < 10
            MsgBox "Il fait froid"
        Case 10 To 20
            MsgBox "Température agréable"
        Case Is > 30
            MsgBox "Il fait très chaud !"
        Case Else
            MsgBox "Température normale"
    End Select
End Sub
```

### Utilisation avec du texte

```vba
Sub TraiterCommande()
    Dim action As String
    action = InputBox("Que voulez-vous faire ? (nouveau/ouvrir/sauver/quitter)")

    Select Case LCase(action)  ' LCase pour ignorer la casse
        Case "nouveau", "new"
            MsgBox "Création d'un nouveau document"
            ' Code pour nouveau document

        Case "ouvrir", "open"
            MsgBox "Ouverture d'un document"
            ' Code pour ouvrir

        Case "sauver", "save", "sauvegarder"
            MsgBox "Sauvegarde du document"
            ' Code pour sauvegarder

        Case "quitter", "exit", "quit"
            MsgBox "Fermeture de l'application"
            ' Code pour quitter

        Case Else
            MsgBox "Commande non reconnue"
    End Select
End Sub
```

### Exemples pratiques avec Excel

#### Formater selon le type de données
```vba
Sub FormaterSelonType()
    Dim cellule As Range
    Set cellule = Selection.Cells(1, 1)  ' Première cellule sélectionnée

    Select Case True  ' Astuce : utiliser True pour tester différentes conditions
        Case IsEmpty(cellule)
            cellule.Interior.Color = RGB(255, 255, 0)  ' Jaune pour vide

        Case IsDate(cellule.Value)
            ' Important : tester IsDate AVANT IsNumeric
            ' car les dates passent aussi le test IsNumeric
            cellule.Interior.Color = RGB(0, 0, 255)    ' Bleu pour date
            cellule.NumberFormat = "dd/mm/yyyy"

        Case IsNumeric(cellule.Value)
            cellule.Interior.Color = RGB(0, 255, 0)    ' Vert pour nombre
            cellule.NumberFormat = "0.00"

        Case Else
            cellule.Interior.Color = RGB(255, 255, 255) ' Blanc pour texte
    End Select
End Sub
```

#### Menu simple
```vba
Sub MenuPrincipal()
    Dim choix As String

    choix = InputBox("Choisissez une option :" & vbNewLine & _
                     "1 - Effacer la feuille" & vbNewLine & _
                     "2 - Créer un tableau" & vbNewLine & _
                     "3 - Sauvegarder" & vbNewLine & _
                     "4 - Quitter")

    Select Case choix
        Case "1"
            If MsgBox("Effacer toute la feuille ?", vbYesNo) = vbYes Then
                Cells.ClearContents
                MsgBox "Feuille effacée"
            End If

        Case "2"
            CreerTableauSample
            MsgBox "Tableau créé en A1"

        Case "3"
            ActiveWorkbook.Save
            MsgBox "Document sauvegardé"

        Case "4"
            MsgBox "Au revoir !"

        Case Else
            MsgBox "Option invalide"
    End Select
End Sub

Sub CreerTableauSample()
    Range("A1:C1").Value = Array("Nom", "Âge", "Ville")
    Range("A1:C1").Font.Bold = True
End Sub
```

## Comparaison If vs Select Case

### Utilisez If quand :
- Vous testez différentes variables
- Vous avez des conditions complexes (And, Or)
- Vous avez peu de conditions (2-3)
- Les conditions ne suivent pas un pattern logique

```vba
' ✅ If est approprié ici
If age >= 18 And permis = True And vue >= 8 Then
    MsgBox "Peut conduire"
ElseIf temperature < 0 Or pluie = True Then
    MsgBox "Conditions dangereuses"
End If
```

### Utilisez Select Case quand :
- Vous testez la même variable contre plusieurs valeurs
- Vous avez beaucoup de conditions similaires
- Les valeurs suivent un pattern (1,2,3... ou "rouge","vert","bleu"...)
- Vous voulez un code plus lisible

```vba
' ✅ Select Case est approprié ici
Select Case codeErreur
    Case 1001
        MsgBox "Erreur de fichier"
    Case 1002
        MsgBox "Erreur de réseau"
    Case 1003 To 1010
        MsgBox "Erreur système"
    Case Else
        MsgBox "Erreur inconnue"
End Select
```

## Erreurs courantes à éviter

### 1. Oublier End If
```vba
' ❌ Incorrect
If condition Then
    MsgBox "Test"
' Manque End If !

' ✅ Correct
If condition Then
    MsgBox "Test"
End If
```

### 2. Confusion entre = et ==
```vba
' ✅ En VBA, utilisez un seul =
If nom = "Marie" Then
    MsgBox "Bonjour Marie"
End If
```

### 3. Oublier Case Else
```vba
' ❌ Que se passe-t-il si jour = 8 ?
Select Case jour
    Case 1 To 7
        MsgBox "Jour valide"
End Select

' ✅ Gérer tous les cas
Select Case jour
    Case 1 To 7
        MsgBox "Jour valide"
    Case Else
        MsgBox "Jour invalide"
End Select
```

### 4. Conditions inaccessibles
```vba
' ❌ Le deuxième Case ne sera jamais atteint
Select Case age
    Case Is > 10
        MsgBox "Plus de 10 ans"
    Case Is > 18
        MsgBox "Majeur"  ' Jamais exécuté !
End Select

' ✅ Ordre correct
Select Case age
    Case Is > 18
        MsgBox "Majeur"
    Case Is > 10
        MsgBox "Plus de 10 ans"
    Case Else
        MsgBox "10 ans ou moins"
End Select
```

## Bonnes pratiques

### 1. Indentation claire
```vba
' ✅ Bien indenté, facile à lire
If condition1 Then
    If condition2 Then
        MsgBox "Les deux conditions sont vraies"
    Else
        MsgBox "Seule la première condition est vraie"
    End If
Else
    MsgBox "La première condition est fausse"
End If
```

### 2. Conditions positives
```vba
' ❌ Difficile à comprendre
If Not Not estValide Then

' ✅ Plus clair
If estValide Then
```

### 3. Variables explicites
```vba
' ❌ Peu clair
If x > 18 Then

' ✅ Plus clair
Dim age As Integer  
age = x  
If age > 18 Then  
```

### 4. Commentaires pour la logique complexe
```vba
Sub CalculerTarif()
    Dim age As Integer
    Dim jour As String

    age = InputBox("Âge ?")
    jour = Format(Date, "dddd")

    ' Tarifs spéciaux le weekend et pour les seniors
    If (jour = "samedi" Or jour = "dimanche") And age >= 65 Then
        MsgBox "Tarif senior weekend : 5€"
    ElseIf age >= 65 Then
        MsgBox "Tarif senior : 8€"
    ElseIf jour = "samedi" Or jour = "dimanche" Then
        MsgBox "Tarif weekend : 12€"
    Else
        MsgBox "Tarif normal : 15€"
    End If
End Sub
```

## Récapitulatif des concepts clés

1. **If...Then...Else** : Structure fondamentale pour les décisions
2. **ElseIf** : Pour tester plusieurs conditions successives
3. **And, Or, Not** : Combiner des conditions complexes
4. **Select Case** : Alternative élégante pour tester une variable contre plusieurs valeurs
5. **Case Else** : Toujours prévoir un cas par défaut
6. **Indentation** : Rendre le code lisible et compréhensible
7. **Conditions positives** : Préférer les formulations claires

Les instructions conditionnelles sont la base de la logique de programmation. Elles permettent à vos programmes de prendre des décisions intelligentes et de s'adapter à toutes les situations !

⏭️
