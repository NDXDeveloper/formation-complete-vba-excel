🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.4 Valeurs de retour des fonctions

## Introduction

Nous avons vu que les procédures (Sub) exécutent des actions, mais ne retournent pas de valeur. Les fonctions (Function), quant à elles, sont conçues pour **calculer quelque chose et retourner le résultat**. Cette section vous apprendra tout ce qu'il faut savoir sur les valeurs de retour des fonctions.

## Comprendre les valeurs de retour

### Analogie de la calculatrice

Imaginez que vous utilisez une calculatrice :
- Vous tapez "5 + 3"
- La calculatrice **calcule** (traite l'information)
- Elle **affiche "8"** (retourne le résultat)

C'est exactement ce que fait une fonction VBA : elle prend des données en entrée, effectue un traitement, et **retourne un résultat** que vous pouvez utiliser ailleurs.

### Différence fondamentale avec les procédures

```vba
' PROCÉDURE - Fait quelque chose, ne retourne rien
Sub AfficherCalcul()
    MsgBox 5 + 3  ' Affiche directement le résultat
End Sub

' FONCTION - Calcule et retourne une valeur
Function Additionner(a As Integer, b As Integer) As Integer
    Additionner = a + b  ' Retourne le résultat
End Function
```

## Syntaxe de base des fonctions

### Structure obligatoire

```vba
Function NomDeLaFonction(paramètres) As TypeDeRetour
    ' Instructions de calcul
    NomDeLaFonction = valeur_à_retourner
End Function
```

### Éléments essentiels

1. **Function** : Mot-clé pour démarrer une fonction
2. **Nom de la fonction** : Identifiant unique
3. **Paramètres** : Données d'entrée (optionnels)
4. **As TypeDeRetour** : Spécifie le type de valeur retournée
5. **Assignation** : `NomFonction = résultat`
6. **End Function** : Marque la fin de la fonction

## Types de retour courants

### Types de base

```vba
' Retourne un nombre entier
Function CalculerAge(anneeNaissance As Integer) As Integer
    CalculerAge = Year(Date) - anneeNaissance
End Function

' Retourne un nombre décimal
Function CalculerTVA(prix As Double) As Double
    CalculerTVA = prix * 0.2
End Function

' Retourne du texte
Function FormaterNom(prenom As String, nom As String) As String
    FormaterNom = UCase(nom) & ", " & prenom
End Function

' Retourne vrai/faux
Function EstMajeur(age As Integer) As Boolean
    If age >= 18 Then
        EstMajeur = True
    Else
        EstMajeur = False
    End If
End Function

' Retourne une date
Function ProchaineLundi() As Date
    Dim aujourd_hui As Date
    Dim joursRestants As Integer
    aujourd_hui = Date
    ' Weekday avec vbMonday : Lundi=1, Dimanche=7
    joursRestants = (9 - Weekday(aujourd_hui, vbMonday)) Mod 7
    If joursRestants = 0 Then joursRestants = 7  ' Si on est lundi, prochain lundi = +7
    ProchaineLundi = aujourd_hui + joursRestants
End Function
```

## Comment retourner une valeur

### Méthode principale : Assignation au nom de la fonction

```vba
Function Multiplier(x As Double, y As Double) As Double
    Multiplier = x * y  ' La valeur est retournée
End Function
```

### Retour anticipé avec Exit Function

Pour quitter une fonction avant la fin (par exemple en cas d'erreur), utilisez `Exit Function` après avoir assigné la valeur de retour :

```vba
Function Diviser(dividende As Double, diviseur As Double) As Double
    If diviseur = 0 Then
        Diviser = 0  ' Assignation de la valeur de retour
        Exit Function  ' Sortie anticipée de la fonction
    End If
    Diviser = dividende / diviseur
End Function
```

## Exemples progressifs

### Exemple 1 : Fonction simple de calcul

```vba
Function CalculerRemise(prix As Double, pourcentage As Double) As Double
    CalculerRemise = prix * (pourcentage / 100)
End Function
```

**Comment l'utiliser :**
```vba
Sub TestRemise()
    Dim prix_original As Double
    Dim remise_calculee As Double

    prix_original = 100
    remise_calculee = CalculerRemise(prix_original, 15)  ' 15% de remise

    MsgBox "Remise de " & remise_calculee & "€ sur " & prix_original & "€"
End Sub
```

### Exemple 2 : Fonction avec logique conditionnelle

```vba
Function DeterminerCategorie(age As Integer) As String
    If age < 13 Then
        DeterminerCategorie = "Enfant"
    ElseIf age < 18 Then
        DeterminerCategorie = "Adolescent"
    ElseIf age < 65 Then
        DeterminerCategorie = "Adulte"
    Else
        DeterminerCategorie = "Senior"
    End If
End Function
```

**Comment l'utiliser :**
```vba
Sub ClasserPersonnes()
    MsgBox "Une personne de 8 ans est : " & DeterminerCategorie(8)
    MsgBox "Une personne de 25 ans est : " & DeterminerCategorie(25)
    MsgBox "Une personne de 70 ans est : " & DeterminerCategorie(70)
End Sub
```

### Exemple 3 : Fonction de validation

```vba
Function EstEmailValide(email As String) As Boolean
    ' Vérification basique : contient @ et un point après @
    If InStr(email, "@") > 0 And InStr(email, ".") > InStr(email, "@") Then
        EstEmailValide = True
    Else
        EstEmailValide = False
    End If
End Function
```

**Comment l'utiliser :**
```vba
Sub ValiderEmails()
    Dim emails As Variant
    Dim i As Integer

    emails = Array("test@example.com", "email.invalide", "autre@domain.fr")

    For i = 0 To UBound(emails)
        If EstEmailValide(emails(i)) Then
            MsgBox emails(i) & " est valide"
        Else
            MsgBox emails(i) & " est invalide"
        End If
    Next i
End Sub
```

## Utiliser les fonctions dans Excel

### Comme formules personnalisées

Une fois créées, vos fonctions peuvent être utilisées directement dans les cellules Excel :

```vba
Function ConvertirCelsiusFahrenheit(celsius As Double) As Double
    ConvertirCelsiusFahrenheit = (celsius * 9 / 5) + 32
End Function
```

**Dans Excel :**
- Tapez `=ConvertirCelsiusFahrenheit(25)` dans une cellule
- Résultat : 77 (25°C = 77°F)

### Fonctions pour analyser des données

```vba
Function CompterMots(texte As String) As Integer
    Dim mots As Variant
    If Len(Trim(texte)) = 0 Then
        CompterMots = 0
    Else
        mots = Split(Trim(texte), " ")
        CompterMots = UBound(mots) + 1
    End If
End Function
```

**Dans Excel :**
- `=CompterMots("Bonjour tout le monde")` retourne 4

## Fonctions avec plusieurs types de retour

### Utilisation de Variant pour flexibilité

```vba
Function AnalyserNombre(valeur As Variant) As Variant
    If IsNumeric(valeur) Then
        If valeur > 0 Then
            AnalyserNombre = "Positif"
        ElseIf valeur < 0 Then
            AnalyserNombre = "Négatif"
        Else
            AnalyserNombre = "Zéro"
        End If
    Else
        AnalyserNombre = "Pas un nombre"
    End If
End Function
```

## Gestion des erreurs dans les fonctions

### Retour de valeurs d'erreur

```vba
Function DiviserSecurise(dividende As Double, diviseur As Double) As Variant
    If diviseur = 0 Then
        DiviserSecurise = "Erreur : Division par zéro"
    Else
        DiviserSecurise = dividende / diviseur
    End If
End Function
```

### Utilisation des codes d'erreur Excel

```vba
Function RacineCarree(nombre As Double) As Variant
    If nombre < 0 Then
        RacineCarree = CVErr(xlErrNum)  ' #NUM! dans Excel
    Else
        RacineCarree = Sqr(nombre)
    End If
End Function
```

## Fonctions imbriquées

### Utiliser une fonction dans une autre

```vba
Function CalculerPrixTTC(prixHT As Double) As Double
    CalculerPrixTTC = prixHT + CalculerTVA(prixHT)
End Function

Function CalculerTVA(prix As Double) As Double
    CalculerTVA = prix * 0.2
End Function
```

### Fonctions qui appellent d'autres fonctions

```vba
Function AnalyseComplete(age As Integer) As String
    Dim categorie As String
    Dim statut_majeur As String

    categorie = DeterminerCategorie(age)

    If EstMajeur(age) Then
        statut_majeur = "Majeur"
    Else
        statut_majeur = "Mineur"
    End If

    AnalyseComplete = categorie & " - " & statut_majeur
End Function
```

## Exemples pratiques pour Excel

### Fonction de formatage de données

```vba
Function FormaterTelephone(numero As String) As String
    ' Enlève tous les espaces et caractères spéciaux
    Dim numeroNettoye As String
    Dim i As Integer

    For i = 1 To Len(numero)
        If IsNumeric(Mid(numero, i, 1)) Then
            numeroNettoye = numeroNettoye & Mid(numero, i, 1)
        End If
    Next i

    ' Formate en XX.XX.XX.XX.XX si 10 chiffres
    If Len(numeroNettoye) = 10 Then
        FormaterTelephone = Mid(numeroNettoye, 1, 2) & "." & _
                           Mid(numeroNettoye, 3, 2) & "." & _
                           Mid(numeroNettoye, 5, 2) & "." & _
                           Mid(numeroNettoye, 7, 2) & "." & _
                           Mid(numeroNettoye, 9, 2)
    Else
        FormaterTelephone = "Format invalide"
    End If
End Function
```

### Fonction de calcul financier

```vba
Function CalculerInteret(capital As Double, taux As Double, duree As Integer) As Double
    ' Calcul d'intérêt simple : Capital × Taux × Durée
    CalculerInteret = capital * (taux / 100) * duree
End Function

Function CalculerCapitalFinal(capital As Double, taux As Double, duree As Integer) As Double
    CalculerCapitalFinal = capital + CalculerInteret(capital, taux, duree)
End Function
```

## Erreurs courantes à éviter

### 1. Oublier d'assigner la valeur de retour

```vba
' ❌ Incorrect - Aucune valeur retournée
Function Additionner(a As Integer, b As Integer) As Integer
    Dim resultat As Integer
    resultat = a + b
    ' Oublié : Additionner = resultat
End Function

' ✅ Correct
Function Additionner(a As Integer, b As Integer) As Integer
    Additionner = a + b
End Function
```

### 2. Type de retour incorrect

```vba
' ❌ Incorrect - Retourne un nombre mais déclaré comme String
Function Calculer() As String
    Calculer = 5 + 3  ' Erreur de type !
End Function

' ✅ Correct
Function Calculer() As Integer
    Calculer = 5 + 3
End Function
```

### 3. Chemins de code sans retour

```vba
' ❌ Incorrect - Pas de retour si age < 0
Function CategoriserAge(age As Integer) As String
    If age >= 0 And age < 18 Then
        CategoriserAge = "Mineur"
    ElseIf age >= 18 Then
        CategoriserAge = "Majeur"
    End If
    ' Que se passe-t-il si age < 0 ?
End Function

' ✅ Correct
Function CategoriserAge(age As Integer) As String
    If age < 0 Then
        CategoriserAge = "Âge invalide"
    ElseIf age < 18 Then
        CategoriserAge = "Mineur"
    Else
        CategoriserAge = "Majeur"
    End If
End Function
```

## Bonnes pratiques

### 1. Noms de fonctions expressifs

```vba
' ❌ Peu clair
Function Calc(x As Double) As Double

' ✅ Clair
Function CalculerPourcentage(valeur As Double) As Double
```

### 2. Une responsabilité par fonction

```vba
' ✅ Fonction focused sur un seul calcul
Function CalculerTVA(prixHT As Double) As Double
    CalculerTVA = prixHT * 0.2
End Function

' ✅ Fonction séparée pour le prix TTC
Function CalculerPrixTTC(prixHT As Double) As Double
    CalculerPrixTTC = prixHT + CalculerTVA(prixHT)
End Function
```

### 3. Gestion des cas limites

```vba
Function CalculerMoyenne(valeurs As Range) As Variant
    If valeurs.Count = 0 Then
        CalculerMoyenne = "Aucune donnée"
        Exit Function
    End If

    CalculerMoyenne = Application.WorksheetFunction.Average(valeurs)
End Function
```

### 4. Documentation des fonctions

```vba
Function CalculerRemise(prix As Double, pourcentage As Double) As Double
    ' Calcule le montant de la remise
    ' prix : Prix original du produit
    ' pourcentage : Pourcentage de remise (ex: 15 pour 15%)
    ' Retourne : Montant de la remise en euros

    CalculerRemise = prix * (pourcentage / 100)
End Function
```

## Récapitulatif des concepts clés

1. **Les fonctions retournent toujours une valeur** définie par leur type
2. **L'assignation se fait avec le nom de la fonction** : `NomFonction = valeur`
3. **Spécifiez le type de retour** avec `As TypeDeDonnees`
4. **Gérez tous les cas possibles** pour éviter les retours vides
5. **Une fonction = une responsabilité** pour un code maintenant
6. **Testez vos fonctions** avec différents types de données
7. **Documentez le but et les paramètres** de vos fonctions

Les fonctions sont des outils puissants qui transforment vos calculs en blocs réutilisables, rendant votre code plus efficace et plus professionnel !

⏭️
