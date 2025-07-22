🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.1. InputBox et MsgBox

## Introduction

Les boîtes de dialogue `InputBox` et `MsgBox` sont les outils les plus simples et les plus couramment utilisés pour créer une interaction avec l'utilisateur en VBA. Elles permettent de demander des informations à l'utilisateur et d'afficher des messages sans avoir besoin de créer des formulaires complexes.

## MsgBox - Afficher des messages

### Qu'est-ce qu'une MsgBox ?

La fonction `MsgBox` affiche une boîte de dialogue avec un message et attend que l'utilisateur clique sur un bouton pour continuer. C'est l'outil idéal pour :
- Informer l'utilisateur d'un résultat
- Demander une confirmation
- Afficher des messages d'erreur ou d'avertissement

### Syntaxe de base

```vba
MsgBox "Votre message ici"
```

### Exemple simple

```vba
Sub MonPremierMessage()
    MsgBox "Bonjour ! Ceci est mon premier message."
End Sub
```

### Syntaxe complète

```vba
MsgBox(Prompt, [Buttons], [Title], [HelpFile], [Context])
```

**Paramètres :**
- **Prompt** : Le message à afficher (obligatoire)
- **Buttons** : Le type de boutons à afficher (optionnel)
- **Title** : Le titre de la boîte de dialogue (optionnel)
- **HelpFile** : Fichier d'aide (rarement utilisé)
- **Context** : Contexte d'aide (rarement utilisé)

### Types de boutons disponibles

| Constante | Valeur | Description |
|-----------|--------|-------------|
| `vbOKOnly` | 0 | Bouton OK uniquement (par défaut) |
| `vbOKCancel` | 1 | Boutons OK et Annuler |
| `vbAbortRetryIgnore` | 2 | Boutons Abandonner, Recommencer, Ignorer |
| `vbYesNoCancel` | 3 | Boutons Oui, Non, Annuler |
| `vbYesNo` | 4 | Boutons Oui et Non |
| `vbRetryCancel` | 5 | Boutons Recommencer et Annuler |

### Icônes disponibles

| Constante | Valeur | Description |
|-----------|--------|-------------|
| `vbCritical` | 16 | Icône d'erreur critique |
| `vbQuestion` | 32 | Icône de question |
| `vbExclamation` | 48 | Icône d'avertissement |
| `vbInformation` | 64 | Icône d'information |

### Exemples avec différents types de boutons

```vba
Sub ExemplesMsgBox()
    ' Message simple
    MsgBox "Opération terminée !", vbInformation, "Information"

    ' Message de confirmation
    MsgBox "Êtes-vous sûr de vouloir continuer ?", vbYesNo + vbQuestion, "Confirmation"

    ' Message d'erreur
    MsgBox "Une erreur s'est produite !", vbCritical, "Erreur"

    ' Message d'avertissement
    MsgBox "Attention : cette action est irréversible.", vbExclamation, "Avertissement"
End Sub
```

### Récupérer la réponse de l'utilisateur

La fonction `MsgBox` retourne une valeur qui indique quel bouton l'utilisateur a cliqué :

| Constante | Valeur | Bouton cliqué |
|-----------|--------|---------------|
| `vbOK` | 1 | OK |
| `vbCancel` | 2 | Annuler |
| `vbAbort` | 3 | Abandonner |
| `vbRetry` | 4 | Recommencer |
| `vbIgnore` | 5 | Ignorer |
| `vbYes` | 6 | Oui |
| `vbNo` | 7 | Non |

```vba
Sub ExempleReponse()
    Dim reponse As Integer

    reponse = MsgBox("Voulez-vous sauvegarder ?", vbYesNo + vbQuestion, "Sauvegarde")

    If reponse = vbYes Then
        MsgBox "Vous avez choisi de sauvegarder."
    Else
        MsgBox "Vous avez choisi de ne pas sauvegarder."
    End If
End Sub
```

## InputBox - Demander une saisie

### Qu'est-ce qu'une InputBox ?

La fonction `InputBox` affiche une boîte de dialogue qui permet à l'utilisateur de saisir une valeur. Elle est parfaite pour :
- Demander un nom ou un mot de passe
- Récupérer une valeur numérique
- Obtenir une donnée simple de l'utilisateur

### Syntaxe de base

```vba
InputBox("Votre question ici")
```

### Exemple simple

```vba
Sub MonPremiereInputBox()
    Dim nom As String
    nom = InputBox("Quel est votre nom ?")
    MsgBox "Bonjour " & nom & " !"
End Sub
```

### Syntaxe complète

```vba
InputBox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context])
```

**Paramètres :**
- **Prompt** : La question à poser (obligatoire)
- **Title** : Le titre de la boîte de dialogue (optionnel)
- **Default** : Valeur par défaut (optionnel)
- **XPos, YPos** : Position de la boîte (optionnel)
- **HelpFile, Context** : Aide (rarement utilisé)

### Exemples avec paramètres

```vba
Sub ExemplesInputBox()
    Dim nom As String
    Dim age As String
    Dim ville As String

    ' InputBox avec titre
    nom = InputBox("Entrez votre nom :", "Identification")

    ' InputBox avec valeur par défaut
    age = InputBox("Entrez votre âge :", "Âge", "25")

    ' InputBox avec titre et valeur par défaut
    ville = InputBox("Dans quelle ville habitez-vous ?", "Localisation", "Paris")

    ' Affichage des résultats
    MsgBox "Nom : " & nom & vbCrLf & "Âge : " & age & vbCrLf & "Ville : " & ville
End Sub
```

### Gestion de l'annulation

Quand l'utilisateur clique sur "Annuler" ou appuie sur Échap, `InputBox` retourne une chaîne vide. Il est important de tester cette situation :

```vba
Sub GestionAnnulation()
    Dim saisie As String

    saisie = InputBox("Entrez votre nom :", "Saisie obligatoire")

    If saisie = "" Then
        MsgBox "Aucune saisie effectuée. Opération annulée."
    Else
        MsgBox "Bonjour " & saisie & " !"
    End If
End Sub
```

### Validation des données saisies

Il est recommandé de vérifier que la saisie correspond à ce qui est attendu :

```vba
Sub ValidationSaisie()
    Dim ageTexte As String
    Dim age As Integer

    ageTexte = InputBox("Entrez votre âge :", "Âge")

    ' Vérification que c'est un nombre
    If IsNumeric(ageTexte) Then
        age = CInt(ageTexte)
        If age > 0 And age < 150 Then
            MsgBox "Votre âge est : " & age & " ans"
        Else
            MsgBox "L'âge doit être entre 1 et 149 ans."
        End If
    Else
        MsgBox "Veuillez entrer un nombre valide."
    End If
End Sub
```

## Combinaison InputBox et MsgBox

Voici un exemple qui combine les deux pour créer une interaction complète :

```vba
Sub CalculatriceSalaire()
    Dim salaireBrut As String
    Dim tauxCotisation As String
    Dim salaireNet As Double

    ' Demande du salaire brut
    salaireBrut = InputBox("Entrez votre salaire brut mensuel :", "Calculatrice de salaire")

    If salaireBrut = "" Then
        MsgBox "Calcul annulé.", vbInformation
        Exit Sub
    End If

    ' Vérification que c'est un nombre
    If Not IsNumeric(salaireBrut) Then
        MsgBox "Veuillez entrer un montant valide.", vbExclamation
        Exit Sub
    End If

    ' Demande du taux de cotisation
    tauxCotisation = InputBox("Entrez le taux de cotisation (%) :", "Taux de cotisation", "23")

    If tauxCotisation = "" Then
        MsgBox "Calcul annulé.", vbInformation
        Exit Sub
    End If

    If Not IsNumeric(tauxCotisation) Then
        MsgBox "Veuillez entrer un taux valide.", vbExclamation
        Exit Sub
    End If

    ' Calcul du salaire net
    salaireNet = CDbl(salaireBrut) * (1 - CDbl(tauxCotisation) / 100)

    ' Affichage du résultat
    MsgBox "Salaire brut : " & salaireBrut & " €" & vbCrLf & _
           "Taux de cotisation : " & tauxCotisation & " %" & vbCrLf & _
           "Salaire net : " & Format(salaireNet, "0.00") & " €", _
           vbInformation, "Résultat du calcul"
End Sub
```

## Conseils et bonnes pratiques

### Messages clairs et concis
- Utilisez un langage simple et compréhensible
- Évitez le jargon technique
- Soyez précis sur ce que vous demandez

### Gestion des erreurs
- Toujours vérifier si l'utilisateur a annulé
- Valider les données saisies avant de les utiliser
- Afficher des messages d'erreur explicites

### Amélioration de l'expérience utilisateur
- Proposez des valeurs par défaut sensées
- Utilisez des titres descriptifs
- Choisissez les bonnes icônes pour le contexte

### Limites à connaître
- `InputBox` ne permet que la saisie de texte simple
- Pas de validation en temps réel
- Interface limitée (pas de mise en forme)
- Pour des besoins plus complexes, utilisez les UserForms

## Fonctions utiles pour le traitement des données

### Fonctions de validation
- `IsNumeric()` : vérifie si une chaîne est un nombre
- `IsDate()` : vérifie si une chaîne est une date
- `Len()` : retourne la longueur d'une chaîne

### Fonctions de conversion
- `CInt()`, `CLng()`, `CDbl()` : conversion en entier, long, double
- `CStr()` : conversion en chaîne de caractères
- `CDate()` : conversion en date

### Fonctions de formatage
- `Format()` : formatage des nombres et dates
- `UCase()`, `LCase()` : conversion en majuscules/minuscules
- `Trim()` : suppression des espaces en début et fin

---

Les `InputBox` et `MsgBox` sont les fondations de l'interaction utilisateur en VBA. Maîtriser ces outils vous permettra de créer des macros interactives et conviviales. Dans la section suivante, nous découvrirons comment créer des interfaces plus sophistiquées avec les UserForms.

⏭️
