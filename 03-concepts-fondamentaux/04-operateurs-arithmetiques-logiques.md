🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.4 Opérateurs arithmétiques et logiques

## Introduction

Les opérateurs sont les outils mathématiques et logiques de VBA. Imaginez-les comme une calculatrice intégrée dans votre programme : ils vous permettent d'effectuer des calculs, de comparer des valeurs, et de prendre des décisions. Cette section vous donnera tous les outils nécessaires pour manipuler les données et créer une logique intelligente dans vos programmes.

## Qu'est-ce qu'un opérateur ?

### Définition simple

**Un opérateur** = Un symbole qui effectue une opération sur une ou plusieurs valeurs

**Analogies pratiques :**
- **Calculatrice** : Les boutons +, -, ×, ÷
- **Comparaison** : "Plus grand que", "égal à"
- **Logique** : "ET", "OU", "NON"

**Structure générale :**
```vba
Résultat = Valeur1 Opérateur Valeur2
```

**Exemples simples :**
```vba
Total = 10 + 5                    ' Addition
EstGrand = Age > 18               ' Comparaison
EstValide = (x > 0) And (y < 100) ' Logique
```

## Opérateurs arithmétiques

### Addition (+)

**Utilisation :**
```vba
Dim Somme As Double
Somme = 10 + 15                   ' Résultat : 25
Somme = 3.5 + 2.7                 ' Résultat : 6.2
Somme = Range("A1").Value + Range("B1").Value
```

**Avec des variables :**
```vba
Dim Prix As Double
Dim TVA As Double
Dim Total As Double

Prix = 100.0
TVA = 20.0
Total = Prix + TVA                ' Résultat : 120.0
```

**Concaténation de chaînes :**
```vba
' ATTENTION : + peut aussi concaténer du texte
Dim Texte As String
Texte = "Bonjour" + " " + "le monde"    ' "Bonjour le monde"
' Mais il vaut mieux utiliser & pour le texte
```

### Soustraction (-)

**Utilisation :**
```vba
Dim Difference As Double
Difference = 20 - 8               ' Résultat : 12
Difference = 5.5 - 2.3           ' Résultat : 3.2
```

**Calculs pratiques :**
```vba
Dim PrixInitial As Double
Dim Remise As Double
Dim PrixFinal As Double

PrixInitial = 200.0
Remise = 30.0
PrixFinal = PrixInitial - Remise  ' Résultat : 170.0
```

**Nombres négatifs :**
```vba
Dim Nombre As Double
Nombre = -15                      ' Nombre négatif
Dim Resultat As Double
Resultat = 10 - (-5)             ' Résultat : 15 (double négation)
```

### Multiplication (*)

**Utilisation :**
```vba
Dim Produit As Double
Produit = 6 * 7                   ' Résultat : 42
Produit = 2.5 * 4                 ' Résultat : 10.0
```

**Calculs financiers :**
```vba
Dim Quantite As Integer
Dim PrixUnitaire As Double
Dim SousTotal As Double

Quantite = 5
PrixUnitaire = 12.50
SousTotal = Quantite * PrixUnitaire    ' Résultat : 62.5
```

**Avec pourcentages :**
```vba
Dim MontantHT As Double
Dim TauxTVA As Double
Dim MontantTVA As Double

MontantHT = 100.0
TauxTVA = 0.20                    ' 20%
MontantTVA = MontantHT * TauxTVA  ' Résultat : 20.0
```

### Division (/)

**Division décimale :**
```vba
Dim Quotient As Double
Quotient = 15 / 4                 ' Résultat : 3.75
Quotient = 10 / 3                 ' Résultat : 3.33333...
```

**Moyennes :**
```vba
Dim Note1 As Double, Note2 As Double, Note3 As Double
Dim Moyenne As Double

Note1 = 15.0
Note2 = 12.0
Note3 = 18.0
Moyenne = (Note1 + Note2 + Note3) / 3    ' Résultat : 15.0
```

**Attention à la division par zéro :**
```vba
Dim x As Double
x = 10 / 0                        ' ERREUR : Division par zéro !

' Solution : vérifier avant
Dim Diviseur As Double
Diviseur = Range("B1").Value
If Diviseur <> 0 Then
    x = 10 / Diviseur
Else
    MsgBox "Impossible de diviser par zéro"
End If
```

### Division entière (\)

**Retourne la partie entière du quotient :**
```vba
Dim Resultat As Integer
Resultat = 15 \ 4                 ' Résultat : 3 (pas 3.75)
Resultat = 10 \ 3                 ' Résultat : 3 (pas 3.33)
Resultat = 20 \ 6                 ' Résultat : 3 (pas 3.33)
```

**Utilisation pratique :**
```vba
' Calculer combien de boîtes complètes pour 47 articles
' si chaque boîte contient 12 articles
Dim NombreArticles As Integer
Dim ArticlesParBoite As Integer
Dim BoitesCompletes As Integer

NombreArticles = 47
ArticlesParBoite = 12
BoitesCompletes = NombreArticles \ ArticlesParBoite    ' Résultat : 3
```

### Modulo (Mod)

**Retourne le reste de la division :**
```vba
Dim Reste As Integer
Reste = 15 Mod 4                  ' Résultat : 3 (15 = 4×3 + 3)
Reste = 10 Mod 3                  ' Résultat : 1 (10 = 3×3 + 1)
Reste = 20 Mod 5                  ' Résultat : 0 (division exacte)
```

**Applications pratiques :**

**Vérifier les nombres pairs/impairs :**
```vba
Dim Nombre As Integer
Nombre = 17
If Nombre Mod 2 = 0 Then
    MsgBox "Nombre pair"
Else
    MsgBox "Nombre impair"        ' Affiché pour 17
End If
```

**Créer des groupes cycliques :**
```vba
' Alterner les couleurs de lignes : une ligne sur deux
Dim i As Integer
For i = 1 To 10
    If i Mod 2 = 0 Then
        Cells(i, 1).Interior.Color = vbLightGray    ' Lignes paires
    Else
        Cells(i, 1).Interior.Color = vbWhite        ' Lignes impaires
    End If
Next i
```

### Puissance (^)

**Élévation à la puissance :**
```vba
Dim Resultat As Double
Resultat = 2 ^ 3                  ' Résultat : 8 (2×2×2)
Resultat = 5 ^ 2                  ' Résultat : 25 (5×5)
Resultat = 9 ^ 0.5                ' Résultat : 3 (racine carrée)
```

**Calculs financiers :**
```vba
' Intérêts composés : Capital × (1 + Taux)^Années
Dim Capital As Double
Dim Taux As Double
Dim Annees As Integer
Dim Montant As Double

Capital = 1000.0
Taux = 0.05                       ' 5%
Annees = 10
Montant = Capital * (1 + Taux) ^ Annees    ' Résultat : ~1628.89
```

## Priorité des opérateurs arithmétiques

### Ordre de calcul

**Priorité (du plus prioritaire au moins prioritaire) :**
1. **Parenthèses** : `()`
2. **Puissance** : `^`
3. **Multiplication et Division** : `*` et `/` et `\`
4. **Modulo** : `Mod`
5. **Addition et Soustraction** : `+` et `-`

### Exemples de priorité

**Sans parenthèses :**
```vba
Dim x As Double
x = 2 + 3 * 4                     ' Résultat : 14 (pas 20 !)
' Calcul : 2 + (3 * 4) = 2 + 12 = 14
```

**Avec parenthèses :**
```vba
Dim x As Double
x = (2 + 3) * 4                   ' Résultat : 20
' Calcul : (2 + 3) * 4 = 5 * 4 = 20
```

**Cas complexe :**
```vba
Dim x As Double
x = 2 + 3 * 4 ^ 2 - 1            ' Résultat : 49
' Calcul : 2 + 3 * (4 ^ 2) - 1 = 2 + 3 * 16 - 1 = 2 + 48 - 1 = 49
```

**Même priorité (de gauche à droite) :**
```vba
Dim x As Double
x = 20 / 4 * 3                    ' Résultat : 15
' Calcul : (20 / 4) * 3 = 5 * 3 = 15
```

### Bonnes pratiques avec les parenthèses

**Clarifier l'intention :**
```vba
' Difficile à comprendre
Resultat = a + b * c / d - e

' Plus clair avec parenthèses
Resultat = a + ((b * c) / d) - e
```

**Calculs financiers explicites :**
```vba
' Prix TTC
PrixTTC = PrixHT * (1 + TauxTVA)

' Intérêts composés
Montant = Capital * ((1 + TauxMensuel) ^ NombreMois)
```

## Opérateurs de comparaison

### Égalité (=)

**Test d'égalité :**
```vba
Dim EstEgal As Boolean
EstEgal = (10 = 10)               ' Résultat : True
EstEgal = (5 = 7)                 ' Résultat : False
EstEgal = (Range("A1").Value = "Bonjour")
```

**Avec des variables :**
```vba
Dim Age As Integer
Age = Range("B1").Value
If Age = 18 Then
    MsgBox "Vous êtes majeur"
End If
```

**Attention avec les décimaux :**
```vba
Dim x As Double
x = 0.1 + 0.2
If x = 0.3 Then                   ' Peut être False à cause de la précision !
    MsgBox "Égaux"
Else
    MsgBox "Pas égaux"            ' Souvent affiché
End If

' Solution : vérifier avec une tolérance
If Abs(x - 0.3) < 0.0001 Then
    MsgBox "Pratiquement égaux"
End If
```

### Inégalité (<>)

**Test de différence :**
```vba
Dim EstDifferent As Boolean
EstDifferent = (10 <> 5)          ' Résultat : True
EstDifferent = (7 <> 7)           ' Résultat : False
```

**Usage pratique :**
```vba
Dim Nom As String
Nom = Range("A1").Value
If Nom <> "" Then
    MsgBox "Nom saisi : " & Nom
Else
    MsgBox "Aucun nom saisi"
End If
```

### Supérieur (>) et Supérieur ou égal (>=)

**Comparaisons numériques :**
```vba
Dim EstSuperieur As Boolean
EstSuperieur = (10 > 5)           ' Résultat : True
EstSuperieur = (3 > 8)            ' Résultat : False
EstSuperieur = (5 >= 5)           ' Résultat : True
EstSuperieur = (4 >= 7)           ' Résultat : False
```

**Validation de seuils :**
```vba
Dim Montant As Double
Montant = Range("C1").Value
If Montant > 1000 Then
    MsgBox "Montant élevé"
ElseIf Montant >= 100 Then
    MsgBox "Montant moyen"
Else
    MsgBox "Montant faible"
End If
```

### Inférieur (<) et Inférieur ou égal (<=)

**Comparaisons numériques :**
```vba
Dim EstInferieur As Boolean
EstInferieur = (5 < 10)           ' Résultat : True
EstInferieur = (8 < 3)            ' Résultat : False
EstInferieur = (5 <= 5)           ' Résultat : True
EstInferieur = (7 <= 4)           ' Résultat : False
```

**Contrôle de limites :**
```vba
Dim Age As Integer
Age = Range("B1").Value
If Age < 18 Then
    MsgBox "Mineur"
ElseIf Age <= 65 Then
    MsgBox "Adulte actif"
Else
    MsgBox "Senior"
End If
```

### Comparaison de chaînes

**Ordre alphabétique :**
```vba
Dim Resultat As Boolean
Resultat = ("A" < "B")            ' Résultat : True
Resultat = ("Apple" < "Banana")   ' Résultat : True
Resultat = ("Z" > "A")            ' Résultat : True
```

**Sensibilité à la casse :**
```vba
Dim Resultat As Boolean
Resultat = ("a" = "A")            ' Résultat : False (sensible à la casse)

' Pour ignorer la casse :
Resultat = (UCase("a") = UCase("A"))    ' Résultat : True
```

**Comparaison pratique :**
```vba
Dim Nom1 As String, Nom2 As String
Nom1 = Range("A1").Value
Nom2 = Range("B1").Value

If UCase(Nom1) = UCase(Nom2) Then
    MsgBox "Noms identiques (casse ignorée)"
End If
```

## Opérateurs logiques

### AND (ET logique)

**Les deux conditions doivent être vraies :**
```vba
Dim Resultat As Boolean
Resultat = (True And True)        ' Résultat : True
Resultat = (True And False)       ' Résultat : False
Resultat = (False And True)       ' Résultat : False
Resultat = (False And False)      ' Résultat : False
```

**Usage pratique :**
```vba
Dim Age As Integer
Dim Permis As Boolean

Age = Range("A1").Value
Permis = Range("B1").Value

If (Age >= 18) And (Permis = True) Then
    MsgBox "Peut conduire"
Else
    MsgBox "Ne peut pas conduire"
End If
```

**Validation multiple :**
```vba
Dim Montant As Double
Dim Stock As Integer

Montant = Range("C1").Value
Stock = Range("D1").Value

If (Montant > 0) And (Stock > 0) And (Montant <= 10000) Then
    MsgBox "Commande valide"
Else
    MsgBox "Commande invalide"
End If
```

### OR (OU logique)

**Au moins une condition doit être vraie :**
```vba
Dim Resultat As Boolean
Resultat = (True Or True)         ' Résultat : True
Resultat = (True Or False)        ' Résultat : True
Resultat = (False Or True)        ' Résultat : True
Resultat = (False Or False)       ' Résultat : False
```

**Conditions alternatives :**
```vba
Dim TypeClient As String
TypeClient = Range("A1").Value

If (TypeClient = "VIP") Or (TypeClient = "Premium") Then
    MsgBox "Client prioritaire"
Else
    MsgBox "Client standard"
End If
```

**Validation souple :**
```vba
Dim Email As String
Dim Telephone As String

Email = Range("B1").Value
Telephone = Range("C1").Value

If (Email <> "") Or (Telephone <> "") Then
    MsgBox "Contact possible"
Else
    MsgBox "Aucun moyen de contact"
End If
```

### NOT (NON logique)

**Inverse la condition :**
```vba
Dim Resultat As Boolean
Resultat = Not True               ' Résultat : False
Resultat = Not False              ' Résultat : True
```

**Usage pratique :**
```vba
Dim EstVide As Boolean
EstVide = (Range("A1").Value = "")

If Not EstVide Then
    MsgBox "Cellule remplie"
Else
    MsgBox "Cellule vide"
End If
```

**Simplification de conditions :**
```vba
' Au lieu de :
If EstActif = False Then

' Écrivez :
If Not EstActif Then
```

### Combinaisons complexes

**Parenthèses pour grouper :**
```vba
Dim Age As Integer
Dim Permis As Boolean
Dim Experience As Integer

Age = 25
Permis = True
Experience = 2

' Peut conduire si : (majeur ET a le permis) ET (expérience >= 1 OU âge >= 25)
If ((Age >= 18) And Permis) And ((Experience >= 1) Or (Age >= 25)) Then
    MsgBox "Peut conduire"
End If
```

**Logique métier complexe :**
```vba
Dim EstClient As Boolean
Dim MontantCommande As Double
Dim EstEnStock As Boolean
Dim ModePaiement As String

' Commande acceptée si :
' Client ET (montant > 0) ET en stock ET (paiement carte OU montant < 500)
If EstClient And (MontantCommande > 0) And EstEnStock And _
   ((ModePaiement = "Carte") Or (MontantCommande < 500)) Then
    MsgBox "Commande acceptée"
End If
```

## Opérateurs de chaînes

### Concaténation (&)

**Assembler du texte :**
```vba
Dim NomComplet As String
Dim Prenom As String, Nom As String

Prenom = "Jean"
Nom = "Dupont"
NomComplet = Prenom & " " & Nom   ' Résultat : "Jean Dupont"
```

**Avec des nombres :**
```vba
Dim Message As String
Dim Age As Integer

Age = 25
Message = "Vous avez " & Age & " ans"    ' "Vous avez 25 ans"
```

**Construction de messages :**
```vba
Dim Produit As String
Dim Prix As Double
Dim Description As String

Produit = Range("A1").Value
Prix = Range("B1").Value
Description = "Produit : " & Produit & " - Prix : " & Prix & "€"
Range("C1").Value = Description
```

### Différence entre + et &

**Recommandation : Utilisez & pour le texte :**
```vba
' Avec & (recommandé pour le texte)
Resultat = "Hello" & " " & "World"       ' "Hello World"

' Avec + (peut créer des confusions)
Resultat = "Hello" + " " + "World"       ' Fonctionne mais déconseillé

' Problème potentiel avec +
Dim x As Variant, y As Variant
x = "5"
y = "3"
Resultat1 = x + y                        ' "53" (concaténation)
Resultat2 = CInt(x) + CInt(y)           ' 8 (addition)
```

## Opérateurs d'affectation

### Affectation simple (=)

**Attribution de valeur :**
```vba
Dim x As Integer
x = 10                            ' x prend la valeur 10
x = x + 5                         ' x devient 15 (10 + 5)
```

### Opérateurs d'affectation composés

**VBA ne supporte pas les opérateurs composés comme +=, -=, etc.**

**Au lieu de :**
```vba
' Ceci n'existe PAS en VBA
x += 5                            ' ERREUR !
y *= 2                            ' ERREUR !
```

**Utilisez :**
```vba
x = x + 5                         ' Addition et affectation
y = y * 2                         ' Multiplication et affectation
Total = Total + Montant           ' Accumulation
```

## Évaluation des expressions

### Ordre d'évaluation

**VBA évalue dans cet ordre :**
1. **Parenthèses** : De l'intérieur vers l'extérieur
2. **Opérateurs arithmétiques** : Selon leur priorité
3. **Opérateurs de comparaison** : De gauche à droite
4. **Opérateurs logiques** : NOT, puis AND, puis OR

**Exemple complexe :**
```vba
Dim Resultat As Boolean
Resultat = (10 + 5 > 12) And (Not (3 * 2 = 5)) Or (True)

' Évaluation étape par étape :
' 1. Parenthèses internes : 10 + 5 = 15, 3 * 2 = 6
' 2. Comparaisons : 15 > 12 = True, 6 = 5 = False
' 3. NOT : Not False = True
' 4. AND : True And True = True
' 5. OR : True Or True = True
' Résultat final : True
```

### Court-circuit (Short-circuit)

**VBA évalue parfois partiellement :**
```vba
' Si la première condition est False, VBA peut ignorer la suite avec AND
If (x > 0) And (10 / x > 2) Then
    ' Si x <= 0, la division ne sera pas évaluée (évite l'erreur)
End If

' Avec OR, si la première condition est True, la suite peut être ignorée
If (EstAdmin = True) Or (Age >= 18) Then
    ' Si EstAdmin est True, Age n'est pas vérifié
End If
```

## Erreurs courantes avec les opérateurs

### Confusion = et ==

**En VBA, utilisez = pour l'affectation ET la comparaison :**
```vba
x = 10                            ' Affectation
If x = 10 Then                    ' Comparaison (même symbole !)
```

**Contexte détermine l'usage :**
```vba
' Dans une affectation
Variable = Expression

' Dans une condition
If Variable = Valeur Then
```

### Priorité mal comprise

**Erreur courante :**
```vba
If x = 1 Or 2 Then               ' INCORRECT !
' VBA comprend : If (x = 1) Or (2) Then
' 2 est toujours True, donc condition toujours vraie
```

**Correction :**
```vba
If (x = 1) Or (x = 2) Then       ' CORRECT
```

### Division par zéro

**Problème :**
```vba
Resultat = 10 / 0                 ' ERREUR d'exécution !
```

**Solution :**
```vba
If Diviseur <> 0 Then
    Resultat = 10 / Diviseur
Else
    MsgBox "Division par zéro impossible"
End If
```

### Comparaison de décimaux

**Problème de précision :**
```vba
Dim x As Double
x = 0.1 + 0.2
If x = 0.3 Then                   ' Peut échouer !
```

**Solution :**
```vba
If Abs(x - 0.3) < 0.000001 Then  ' Comparaison avec tolérance
```

## Utilisation pratique dans Excel

### Calculs sur les cellules

```vba
Sub CalculerTotalCommande()
    Dim Quantite As Integer
    Dim PrixUnitaire As Double
    Dim Remise As Double
    Dim Total As Double

    Quantite = Range("B2").Value
    PrixUnitaire = Range("C2").Value
    Remise = Range("D2").Value / 100     ' Conversion pourcentage

    Total = (Quantite * PrixUnitaire) * (1 - Remise)
    Range("E2").Value = Total
End Sub
```

### Validation conditionnelle

```vba
Sub ValiderDonnees()
    Dim Age As Integer
    Dim Salaire As Double
    Dim EstValide As Boolean

    Age = Range("A1").Value
    Salaire = Range("B1").Value

    EstValide = (Age >= 18) And (Age <= 65) And (Salaire > 0)

    If EstValide Then
        Range("C1").Value = "VALIDE"
        Range("C1").Interior.Color = vbGreen
    Else
        Range("C1").Value = "INVALIDE"
        Range("C1").Interior.Color = vbRed
    End If
End Sub
```

### Formatage conditionnel

```vba
Sub FormaterSelonValeur()
    Dim i As Integer
    Dim Valeur As Double

    For i = 1 To 10
        Valeur = Cells(i, 1).Value

        If Valeur > 100 Then
            Cells(i, 1).Interior.Color = vbGreen
        ElseIf Valeur >= 50 Then
            Cells(i, 1).Interior.Color = vbYellow
        Else
            Cells(i, 1).Interior.Color = vbRed
        End If
    Next i
End Sub
```

## Résumé

Les opérateurs sont les outils de calcul et de logique en VBA :

**Opérateurs arithmétiques :**
- **Addition** : `+` (nombres et concaténation)
- **Soustraction** : `-`
- **Multiplication** : `*`
- **Division** : `/` (décimale) et `\` (entière)
- **Modulo** : `Mod` (reste de division)
- **Puissance** : `^`

**Opérateurs de comparaison :**
- **Égalité** : `=`, **Inégalité** : `<>`
- **Supérieur** : `>`, `>=`
- **Inférieur** : `<`, `<=`

**Opérateurs logiques :**
- **AND** : Les deux conditions vraies
- **OR** : Au moins une condition vraie
- **NOT** : Inverse la condition

**Priorité des opérateurs :**
1. **Parenthèses** `()`
2. **Puissance** `^`
3. **Multiplication/Division** `*` `/` `\` `Mod`
4. **Addition/Soustraction** `+` `-`
5. **Comparaison** `=` `<>` `<` `>` `<=` `>=`
6. **Logique** `NOT` puis `AND` puis `OR`

**Bonnes pratiques :**
- **Parenthèses** : Pour clarifier les expressions complexes
- **Vérification** : Division par zéro, valeurs nulles
- **Tolérance** : Pour les comparaisons de décimaux
- **& pour texte** : Préférer & à + pour la concaténation

**À retenir :**
- **Testez vos expressions** dans la fenêtre immédiate
- **Parenthèses** : En cas de doute, utilisez-les
- **Logique claire** : Décomposez les conditions complexes
- **Gestion d'erreurs** : Anticipez les cas limites

Dans la section suivante, nous découvrirons l'art de commenter votre code pour le rendre compréhensible et maintenable.

⏭️
