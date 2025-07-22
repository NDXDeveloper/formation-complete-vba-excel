🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.3 Paramètres et arguments

## Introduction

Jusqu'à présent, vous avez créé des procédures qui font toujours la même chose. Mais que se passerait-il si vous vouliez que votre procédure s'adapte à différentes situations ? C'est exactement le rôle des **paramètres** et des **arguments** !

## Comprendre les concepts

### Analogie de la machine à café

Imaginez une machine à café :
- Sans paramètres : elle ne ferait qu'un seul type de café, toujours identique
- Avec paramètres : vous pouvez choisir le type de café, la quantité de sucre, la taille de la tasse

De la même façon, les paramètres permettent à vos procédures de recevoir des "instructions" pour s'adapter à vos besoins.

### Définitions importantes

- **Paramètre** : Une "boîte" que vous définissez dans votre procédure pour recevoir une information
- **Argument** : La valeur réelle que vous "mettez dans la boîte" quand vous appelez la procédure

```vba
' "nom" est un PARAMÈTRE
Sub DireBonjour(nom As String)
    MsgBox "Bonjour " & nom & " !"
End Sub

' Quand vous appelez la procédure :
' DireBonjour("Marie")
' "Marie" est un ARGUMENT
```

## Syntaxe de base

### Procédure avec un paramètre

```vba
Sub NomProcedure(nomParametre As TypeDeDonnees)
    ' Instructions utilisant nomParametre
End Sub
```

### Procédure avec plusieurs paramètres

```vba
Sub NomProcedure(param1 As Type1, param2 As Type2, param3 As Type3)
    ' Instructions utilisant param1, param2, param3
End Sub
```

## Exemples progressifs

### Exemple 1 : Procédure avec un paramètre simple

```vba
Sub AfficherMessage(texte As String)
    MsgBox texte
End Sub
```

**Comment l'utiliser :**
```vba
Sub TestAffichage()
    AfficherMessage "Ceci est mon message personnalisé"
    AfficherMessage "Un autre message différent"
End Sub
```

### Exemple 2 : Procédure avec deux paramètres

```vba
Sub EcrireDansCellule(adresse As String, contenu As String)
    Range(adresse).Value = contenu
End Sub
```

**Comment l'utiliser :**
```vba
Sub TestEcriture()
    EcrireDansCellule "A1", "Titre du rapport"
    EcrireDansCellule "B2", "Données importantes"
    EcrireDansCellule "C5", "Total"
End Sub
```

### Exemple 3 : Procédure avec différents types de paramètres

```vba
Sub FormaterCellule(adresse As String, texte As String, taille As Integer, couleur As Long)
    With Range(adresse)
        .Value = texte
        .Font.Size = taille
        .Font.Color = couleur
        .Font.Bold = True
    End With
End Sub
```

**Comment l'utiliser :**
```vba
Sub TestFormatage()
    FormaterCellule "A1", "TITRE PRINCIPAL", 16, RGB(255, 0, 0)    ' Rouge
    FormaterCellule "A2", "Sous-titre", 12, RGB(0, 0, 255)        ' Bleu
End Sub
```

## Types de paramètres courants

### Types de base
```vba
Sub ExempleTypes(texte As String, _
                nombre As Integer, _
                decimal As Double, _
                vrai_faux As Boolean, _
                date_heure As Date)

    MsgBox "Texte reçu : " & texte
    MsgBox "Nombre reçu : " & nombre
    MsgBox "Décimal reçu : " & decimal
    MsgBox "Booléen reçu : " & vrai_faux
    MsgBox "Date reçue : " & date_heure
End Sub
```

### Utilisation avec des valeurs concrètes
```vba
Sub AppelerExempleTypes()
    ExempleTypes "Bonjour", 25, 3.14, True, Date
End Sub
```

## Paramètres optionnels

Parfois, vous voulez qu'un paramètre soit facultatif. Utilisez le mot-clé `Optional` :

```vba
Sub DireBonjour(nom As String, Optional politesse As String = "Monsieur/Madame")
    MsgBox "Bonjour " & politesse & " " & nom
End Sub
```

**Utilisation :**
```vba
Sub TestParametresOptionels()
    DireBonjour "Martin"                    ' Utilise la valeur par défaut
    DireBonjour "Sophie", "Mademoiselle"    ' Utilise la valeur fournie
End Sub
```

## Exemples pratiques utiles

### Exemple 1 : Créer un en-tête personnalisé

```vba
Sub CreerEntete(titre As String, sous_titre As String, couleur_fond As Long)
    ' Titre principal
    Range("A1").Value = titre
    Range("A1").Font.Size = 18
    Range("A1").Font.Bold = True
    Range("A1").Interior.Color = couleur_fond

    ' Sous-titre
    Range("A2").Value = sous_titre
    Range("A2").Font.Size = 12
    Range("A2").Font.Italic = True

    ' Ajuster les colonnes
    Range("A:A").AutoFit
End Sub
```

**Utilisation :**
```vba
Sub CreerDifferentsRapports()
    CreerEntete "RAPPORT FINANCIER", "Trimestre 1 - 2024", RGB(173, 216, 230)  ' Bleu clair
    ' OU
    CreerEntete "ANALYSE MARKETING", "Campagne Été", RGB(144, 238, 144)       ' Vert clair
End Sub
```

### Exemple 2 : Remplir une plage de données

```vba
Sub RemplirPlage(cellule_debut As String, cellule_fin As String, valeur As String)
    Range(cellule_debut & ":" & cellule_fin).Value = valeur
    MsgBox "Plage " & cellule_debut & ":" & cellule_fin & " remplie avec : " & valeur
End Sub
```

**Utilisation :**
```vba
Sub TestRemplissage()
    RemplirPlage "A1", "A10", "Produit"
    RemplirPlage "B1", "B10", "0"
    RemplirPlage "C1", "C10", "En attente"
End Sub
```

### Exemple 3 : Calculer et afficher une remise

```vba
Sub AfficherRemise(prix_original As Double, pourcentage_remise As Double, nom_produit As String)
    Dim montant_remise As Double
    Dim prix_final As Double

    montant_remise = prix_original * (pourcentage_remise / 100)
    prix_final = prix_original - montant_remise

    MsgBox "Produit : " & nom_produit & vbNewLine & _
           "Prix original : " & prix_original & "€" & vbNewLine & _
           "Remise (" & pourcentage_remise & "%) : " & montant_remise & "€" & vbNewLine & _
           "Prix final : " & prix_final & "€"
End Sub
```

**Utilisation :**
```vba
Sub CalculerRemises()
    AfficherRemise 100, 15, "Ordinateur portable"
    AfficherRemise 50, 20, "Clavier mécanique"
    AfficherRemise 25, 10, "Souris ergonomique"
End Sub
```

## Passage de paramètres par valeur vs par référence

### Par valeur (ByVal) - Comportement par défaut

```vba
Sub ModifierParValeur(ByVal nombre As Integer)
    nombre = nombre + 10
    MsgBox "Dans la procédure : " & nombre
End Sub

Sub TestParValeur()
    Dim monNombre As Integer
    monNombre = 5

    ModifierParValeur monNombre
    MsgBox "Après la procédure : " & monNombre  ' Toujours 5 !
End Sub
```

### Par référence (ByRef)

```vba
Sub ModifierParReference(ByRef nombre As Integer)
    nombre = nombre + 10
    MsgBox "Dans la procédure : " & nombre
End Sub

Sub TestParReference()
    Dim monNombre As Integer
    monNombre = 5

    ModifierParReference monNombre
    MsgBox "Après la procédure : " & monNombre  ' Maintenant 15 !
End Sub
```

**Règle simple :**
- **ByVal** : La procédure reçoit une "copie" - l'original ne change pas
- **ByRef** : La procédure reçoit l'"original" - peut le modifier

## Erreurs courantes à éviter

### 1. Oublier de spécifier le type

```vba
' ❌ Incorrect
Sub MaProcedure(nom)  ' Type manquant
    MsgBox nom
End Sub

' ✅ Correct
Sub MaProcedure(nom As String)
    MsgBox nom
End Sub
```

### 2. Mauvais ordre des arguments

```vba
Sub AfficherInfo(nom As String, age As Integer)
    MsgBox nom & " a " & age & " ans"
End Sub

Sub Test()
    ' ❌ Incorrect - ordre inversé
    AfficherInfo 25, "Paul"  ' Erreur !

    ' ✅ Correct
    AfficherInfo "Paul", 25
End Sub
```

### 3. Types incompatibles

```vba
Sub TraiterNombre(nombre As Integer)
    MsgBox nombre * 2
End Sub

Sub Test()
    ' ❌ Incorrect - texte au lieu de nombre
    TraiterNombre "Bonjour"  ' Erreur !

    ' ✅ Correct
    TraiterNombre 10
End Sub
```

## Bonnes pratiques

### 1. Noms de paramètres explicites

```vba
' ❌ Peu clair
Sub Proc(a As String, b As Integer, c As Boolean)

' ✅ Clair
Sub FormaterTexte(texte As String, taille As Integer, gras As Boolean)
```

### 2. Ordre logique des paramètres

```vba
' ✅ Du plus important au moins important
Sub CreerFacture(nom_client As String, montant As Double, Optional tva As Double = 0.2)
```

### 3. Utiliser des valeurs par défaut sensées

```vba
Sub SauvegarderFichier(nom_fichier As String, Optional format As String = "xlsx")
    ' Si format n'est pas spécifié, utilise Excel par défaut
End Sub
```

### 4. Documenter vos paramètres

```vba
Sub CalculerStatistiques(donnees As Range, Optional inclure_moyenne As Boolean = True)
    ' donnees : Plage de cellules contenant les valeurs numériques
    ' inclure_moyenne : Si True, calcule aussi la moyenne (défaut : True)

    ' Code de la procédure...
End Sub
```

## Appel de procédures avec paramètres

### Méthode 1 : Appel direct (recommandée)

```vba
Sub AppelDirect()
    FormaterCellule "A1", "Titre", 16, RGB(255, 0, 0)
End Sub
```

### Méthode 2 : Avec le mot-clé Call

```vba
Sub AppelAvecCall()
    Call FormaterCellule("A1", "Titre", 16, RGB(255, 0, 0))
End Sub
```

### Méthode 3 : Avec noms de paramètres (pour plus de clarté)

```vba
Sub AppelAvecNoms()
    FormaterCellule adresse:="A1", texte:="Titre", taille:=16, couleur:=RGB(255, 0, 0)
End Sub
```

## Récapitulatif des concepts clés

1. **Les paramètres rendent vos procédures flexibles** et réutilisables
2. **Spécifiez toujours le type** de chaque paramètre
3. **L'ordre des arguments** doit correspondre à l'ordre des paramètres
4. **Optional** permet de créer des paramètres facultatifs
5. **ByVal vs ByRef** détermine si l'original peut être modifié
6. **Noms explicites** rendent votre code plus lisible
7. **Documentation** aide à comprendre le rôle de chaque paramètre

Les paramètres transforment vos procédures de "robots rigides" en "assistants intelligents" qui s'adaptent à vos besoins spécifiques !

⏭️
