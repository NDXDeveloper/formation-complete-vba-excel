🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.1 Différence entre Sub et Function

## Vue d'ensemble

En VBA, il existe deux types principaux de blocs de code réutilisables : les **procédures (Sub)** et les **fonctions (Function)**. Bien qu'ils puissent sembler similaires au premier regard, ils ont des rôles et des comportements distincts qu'il est essentiel de comprendre.

## Les procédures (Sub)

### Définition

Une **procédure** (ou **Sub**) est un bloc de code qui exécute une série d'actions spécifiques. Elle agit comme un "robot" qui suit des instructions précises pour accomplir une tâche.

### Caractéristiques principales

- **Ne retourne pas de valeur** : Une procédure fait quelque chose, mais ne renvoie pas de résultat
- **Peut modifier l'environnement** : Elle peut changer des données, afficher des messages, formater des cellules, etc.
- **S'exécute du début à la fin** : Elle suit les instructions dans l'ordre
- **Peut recevoir des paramètres** : On peut lui fournir des informations pour personnaliser son comportement

### Syntaxe de base

```vba
Sub NomDeLaProcedure()
    ' Instructions à exécuter
End Sub
```

### Exemples concrets

**Exemple 1 : Afficher un message**
```vba
Sub DireBonjour()
    MsgBox "Bonjour ! Comment allez-vous ?"
End Sub
```

**Exemple 2 : Formater une cellule**
```vba
Sub FormaterCelluleA1()
    Range("A1").Font.Bold = True
    Range("A1").Font.Color = RGB(255, 0, 0)  ' Rouge
    Range("A1").Value = "Titre important"
End Sub
```

**Exemple 3 : Effacer une plage de cellules**
```vba
Sub EffacerDonnees()
    Range("A1:D10").ClearContents
    MsgBox "Les données ont été effacées !"
End Sub
```

## Les fonctions (Function)

### Définition

Une **fonction** (ou **Function**) est un bloc de code qui effectue un calcul ou une opération et **retourne un résultat**. Elle agit comme une "machine" qui prend des ingrédients (paramètres) et produit un résultat.

### Caractéristiques principales

- **Retourne toujours une valeur** : Le but principal est de calculer et renvoyer un résultat
- **Peut être utilisée dans des formules** : Comme les fonctions Excel (SOMME, MOYENNE, etc.)
- **Généralement ne modifie pas l'environnement** : Elle calcule sans changer les données existantes
- **Peut recevoir des paramètres** : Nécessaires pour effectuer ses calculs

### Syntaxe de base

```vba
Function NomDeLaFonction() As TypeDeRetour
    ' Instructions de calcul
    NomDeLaFonction = resultat  ' Assignation du résultat
End Function
```

### Exemples concrets

**Exemple 1 : Calculer l'aire d'un rectangle**
```vba
Function AireRectangle(longueur As Double, largeur As Double) As Double
    AireRectangle = longueur * largeur
End Function
```

**Exemple 2 : Déterminer si un nombre est pair**
```vba
Function EstPair(nombre As Integer) As Boolean
    If nombre Mod 2 = 0 Then
        EstPair = True
    Else
        EstPair = False
    End If
End Function
```

**Exemple 3 : Convertir des degrés Celsius en Fahrenheit**
```vba
Function CelsiusVersFahrenheit(celsius As Double) As Double
    CelsiusVersFahrenheit = (celsius * 9 / 5) + 32
End Function
```

## Comparaison détaillée

| Aspect | Procédure (Sub) | Fonction (Function) |
|--------|-----------------|-------------------|
| **Objectif principal** | Exécuter des actions | Calculer et retourner une valeur |
| **Valeur de retour** | Aucune | Toujours une valeur |
| **Utilisation typique** | Automatiser des tâches | Effectuer des calculs |
| **Dans une formule Excel** | Impossible | Possible |
| **Modification des données** | Courante | Généralement évitée |
| **Appel dans le code** | `Call NomProcedure()` ou `NomProcedure` | `resultat = NomFonction()` |

## Analogies pour mieux comprendre

### La procédure : comme un employé qui exécute des tâches

Imaginez un employé de bureau à qui vous demandez :
- "Imprimez ce document"
- "Envoyez cet email"
- "Classez ces dossiers"

L'employé **fait** ces actions, mais ne vous **retourne** rien de tangible. C'est exactement ce que fait une procédure.

### La fonction : comme une calculatrice

Quand vous utilisez une calculatrice :
- Vous entrez des nombres (paramètres)
- Elle effectue un calcul
- Elle vous **affiche le résultat**

Vous pouvez ensuite utiliser ce résultat pour autre chose. C'est le principe d'une fonction.

## Quand utiliser Sub ou Function ?

### Utilisez une **procédure (Sub)** quand :
- Vous voulez **automatiser une tâche**
- Vous devez **modifier des données** (formater, supprimer, déplacer)
- Vous voulez **afficher des informations** à l'utilisateur
- Vous **organisez une séquence d'actions**

**Exemples d'usage :** Sauvegarder un fichier, générer un rapport, nettoyer des données, créer un graphique.

### Utilisez une **fonction (Function)** quand :
- Vous devez **calculer quelque chose**
- Vous voulez **traiter des données** et obtenir un résultat
- Vous souhaitez **créer une formule personnalisée** pour Excel
- Vous devez **convertir ou transformer** des valeurs

**Exemples d'usage :** Calculer une remise, déterminer l'âge d'une personne, convertir des unités, valider un format.

## Points importants à retenir

1. **Une fonction DOIT retourner une valeur**, une procédure ne le fait jamais
2. **Les fonctions peuvent être utilisées dans les cellules Excel** comme des formules personnalisées
3. **Les procédures sont parfaites pour l'automatisation** de tâches répétitives
4. **Une bonne pratique** : les fonctions calculent, les procédures agissent
5. **Vous pouvez appeler des fonctions depuis des procédures** et vice versa

## Exemple illustratif complet

Voici un exemple qui montre les deux concepts travaillant ensemble :

```vba
' Fonction qui calcule une remise
Function CalculerRemise(prixOriginal As Double, pourcentageRemise As Double) As Double
    CalculerRemise = prixOriginal * (pourcentageRemise / 100)
End Function

' Procédure qui utilise la fonction et affiche le résultat
Sub AfficherPrixAvecRemise()
    Dim prix As Double
    Dim remise As Double
    Dim montantRemise As Double
    Dim prixFinal As Double

    prix = 100
    remise = 15  ' 15% de remise

    ' Utilisation de la fonction
    montantRemise = CalculerRemise(prix, remise)
    prixFinal = prix - montantRemise

    ' La procédure affiche le résultat
    MsgBox "Prix original: " & prix & "€" & vbNewLine & _
           "Remise (" & remise & "%): " & montantRemise & "€" & vbNewLine & _
           "Prix final: " & prixFinal & "€"
End Sub
```

Dans cet exemple :
- La **fonction** `CalculerRemise` fait un calcul et retourne le montant de la remise
- La **procédure** `AfficherPrixAvecRemise` utilise cette fonction et affiche le résultat à l'utilisateur

Cette combinaison illustre parfaitement comment les deux types de blocs de code peuvent travailler ensemble pour créer des solutions efficaces et bien organisées.

⏭️ [Création de procédures simples](/04-procedures-et-fonctions/02-creation-procedures-simples.md)
