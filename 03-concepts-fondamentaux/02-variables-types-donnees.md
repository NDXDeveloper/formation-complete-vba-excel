🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.2 Variables et types de données

## Introduction

Les variables sont comme des boîtes étiquetées où vous pouvez stocker des informations. Imaginez votre bureau : vous avez différents tiroirs pour ranger différents types d'objets. En VBA, c'est pareil ! Vous avez différents types de variables pour stocker différents types d'informations : nombres, texte, dates, etc.

## Qu'est-ce qu'une variable ?

### Définition simple

**Une variable** = Un emplacement de stockage nommé dans la mémoire de l'ordinateur

**Analogies pratiques :**
- **Boîte étiquetée** : Vous pouvez y mettre quelque chose et le retrouver plus tard
- **Tiroir** : Vous savez ce qu'il contient grâce à son étiquette
- **Post-it** : Une note où vous écrivez une information pour vous en souvenir

### Pourquoi utiliser des variables ?

**Stockage temporaire :**
```vba
' Sans variable (répétitif) :
Range("A1").Value = Range("B1").Value + Range("C1").Value
Range("A2").Value = (Range("B1").Value + Range("C1").Value) * 1.2

' Avec variable (plus clair) :
Dim Total As Double
Total = Range("B1").Value + Range("C1").Value
Range("A1").Value = Total
Range("A2").Value = Total * 1.2
```

**Réutilisation :**
- **Calculer une fois** : Stocker le résultat dans une variable
- **Utiliser plusieurs fois** : Réutiliser la variable sans recalculer
- **Maintenir la cohérence** : Même valeur utilisée partout

**Lisibilité du code :**
```vba
' Difficile à comprendre :
If Range("E5").Value * 0.2 > 1000 Then

' Plus clair :
Dim Salaire As Double
Dim TauxTVA As Double
Salaire = Range("E5").Value
TauxTVA = 0.2
If Salaire * TauxTVA > 1000 Then
```

## Déclaration des variables

### Syntaxe de base

**Format général :**
```vba
Dim NomVariable As TypeDeDonnées
```

**Exemples simples :**
```vba
Dim Age As Integer              ' Pour stocker un âge
Dim Nom As String              ' Pour stocker un nom
Dim Prix As Double             ' Pour stocker un prix
Dim EstValide As Boolean       ' Pour stocker vrai/faux
```

### Déclaration vs Affectation

**Déclaration** : Créer la variable
```vba
Dim MonNombre As Integer       ' Je crée une boîte appelée "MonNombre"
```

**Affectation** : Donner une valeur
```vba
MonNombre = 25                 ' Je mets 25 dans la boîte
```

**Les deux en même temps** (optionnel) :
```vba
Dim MonNombre As Integer
MonNombre = 25

' Ou en deux lignes séparées pour plus de clarté
```

### Option Explicit

**Qu'est-ce que c'est ?**
Une instruction qui force à déclarer toutes les variables avant de les utiliser.

**Où la placer :**
```vba
Option Explicit                 ' Tout en haut du module, avant toute procédure

Sub MaProcedure()
    Dim x As Integer           ' Déclaration obligatoire
    x = 10                     ' Utilisation autorisée
End Sub
```

**Pourquoi l'utiliser ?**
- **Évite les erreurs de frappe** : `MonNombr` au lieu de `MonNombre`
- **Code plus propre** : Variables explicitement déclarées
- **Débogage facilité** : VBA signale les variables non déclarées
- **Bonne pratique** : Recommandé par tous les développeurs

## Types de données fondamentaux

### Integer (Nombre entier)

**Utilisation :**
- **Compteurs** : Pour compter des éléments
- **Positions** : Numéros de lignes, colonnes
- **Quantités** : Nombre d'articles, d'employés

**Plage de valeurs :** -32 768 à 32 767

**Exemples :**
```vba
Dim NombreEmployes As Integer
Dim LigneActuelle As Integer
Dim Compteur As Integer

NombreEmployes = 150
LigneActuelle = 5
Compteur = 0
```

**Quand l'utiliser :**
- ✅ Âges, quantités, compteurs
- ❌ Calculs financiers (utilisez Double)

### Long (Nombre entier long)

**Utilisation :**
- **Grandes quantités** : Numéros de références, codes postaux
- **Dates** : Stockage interne des dates
- **Compteurs importants** : Au-delà de 32 767

**Plage de valeurs :** -2 147 483 648 à 2 147 483 647

**Exemples :**
```vba
Dim NumeroCommande As Long
Dim CodePostal As Long
Dim NombreVentes As Long

NumeroCommande = 2024001534
CodePostal = 75001
NombreVentes = 45000
```

### Double (Nombre décimal)

**Utilisation :**
- **Calculs financiers** : Prix, montants, taux
- **Mesures** : Longueurs, poids, pourcentages
- **Résultats de division** : Moyennes, ratios

**Précision :** ~15 chiffres significatifs

**Exemples :**
```vba
Dim PrixUnitaire As Double
Dim TauxTVA As Double
Dim Moyenne As Double

PrixUnitaire = 19.99
TauxTVA = 0.20
Moyenne = 15.75
```

**Notation des décimales :**
```vba
Dim x As Double
x = 3.14159                    ' Correct (point décimal)
x = 3,14159                    ' INCORRECT en VBA !
```

### String (Chaîne de caractères)

**Utilisation :**
- **Texte** : Noms, descriptions, commentaires
- **Codes** : Références textuelles
- **Messages** : Affichages pour l'utilisateur

**Exemples :**
```vba
Dim NomClient As String
Dim Adresse As String
Dim Message As String

NomClient = "Dupont"
Adresse = "123 Rue de la Paix"
Message = "Commande validée"
```

**Chaînes vides :**
```vba
Dim Nom As String
Nom = ""                       ' Chaîne vide
' ou
Nom = String(0, " ")          ' Autre façon de créer une chaîne vide
```

**Longueur variable :**
```vba
Dim Texte As String
Texte = "Court"               ' 5 caractères
Texte = "Beaucoup plus long"  ' 19 caractères (même variable)
```

### Boolean (Booléen)

**Utilisation :**
- **États** : Vrai/Faux, Oui/Non, Actif/Inactif
- **Conditions** : Résultats de tests
- **Drapeaux** : Marquer si quelque chose s'est passé

**Valeurs possibles :** `True` (Vrai) ou `False` (Faux)

**Exemples :**
```vba
Dim EstValide As Boolean
Dim CommandeTerminee As Boolean
Dim EstEnStock As Boolean

EstValide = True
CommandeTerminee = False
EstEnStock = (Range("B1").Value > 0)    ' Résultat d'un test
```

**Utilisation dans les conditions :**
```vba
If EstValide Then
    ' Faire quelque chose si c'est vrai
End If

If Not CommandeTerminee Then
    ' Faire quelque chose si c'est faux
End If
```

### Date

**Utilisation :**
- **Dates** : Dates de naissance, de commande, d'échéance
- **Heures** : Heures de début, de fin
- **Combinaison** : Date et heure ensemble

**Format interne :** VBA stocke les dates comme des nombres

**Exemples :**
```vba
Dim DateNaissance As Date
Dim DateCommande As Date
Dim HeureDebut As Date

DateNaissance = #1/15/1990#           ' Format américain
DateCommande = DateValue("15/01/2024") ' Conversion depuis texte
HeureDebut = TimeValue("09:30:00")    ' Heure uniquement
```

**Fonctions utiles :**
```vba
Dim Aujourd As Date
Dim Maintenant As Date

Aujourd = Date                 ' Date du jour
Maintenant = Now              ' Date et heure actuelles
```

### Variant (Type universel)

**Utilisation :**
- **Type inconnu** : Quand vous ne savez pas à l'avance
- **Données mixtes** : Peut contenir n'importe quoi
- **Compatibilité** : Avec Excel qui renvoie souvent des Variant

**Particularités :**
```vba
Dim MaVariable As Variant
' ou simplement :
Dim MaVariable                 ' Variant par défaut

MaVariable = 10               ' Nombre
MaVariable = "Texte"          ' Texte
MaVariable = True             ' Booléen
MaVariable = Date             ' Date
```

**Avantages :**
- **Flexibilité** : Peut tout contenir
- **Simplicité** : Pas besoin de connaître le type à l'avance

**Inconvénients :**
- **Performance** : Plus lent que les types spécifiques
- **Mémoire** : Consomme plus de mémoire
- **Erreurs** : Plus difficile à déboguer

## Affectation de valeurs

### Opérateur d'affectation

**Symbole = :**
```vba
Variable = Valeur              ' La valeur va dans la variable
```

**Direction obligatoire :**
```vba
x = 10                        ' Correct : 10 va dans x
10 = x                        ' INCORRECT : impossible !
```

### Types d'affectation

**Valeurs littérales :**
```vba
Dim Age As Integer
Dim Nom As String
Dim Prix As Double

Age = 25                      ' Nombre
Nom = "Pierre"                ' Texte
Prix = 19.99                  ' Décimal
```

**Résultats de calculs :**
```vba
Dim Total As Double
Dim Moyenne As Double

Total = 100 + 200 + 50        ' Calcul simple
Moyenne = Total / 3           ' Division
```

**Valeurs de cellules Excel :**
```vba
Dim Valeur As Double
Dim Texte As String

Valeur = Range("A1").Value    ' Récupère la valeur de A1
Texte = Range("B1").Value     ' Récupère le texte de B1
```

**Autres variables :**
```vba
Dim x As Integer
Dim y As Integer

x = 10
y = x                         ' y prend la valeur de x (10)
```

## Portée des variables

### Variables locales (dans une procédure)

**Déclaration dans une procédure :**
```vba
Sub MaProcedure()
    Dim MonNombre As Integer   ' Variable locale
    MonNombre = 10
    ' MonNombre n'existe que dans cette procédure
End Sub

Sub AutreProcedure()
    ' MonNombre n'est pas accessible ici
    Dim MonNombre As Integer   ' Différente variable, même nom
End Sub
```

**Durée de vie :** Du début à la fin de la procédure

### Variables au niveau du module

**Déclaration en haut du module :**
```vba
' En haut du module, avant toute procédure
Dim VariableModule As Integer

Sub Procedure1()
    VariableModule = 10        ' Accessible
End Sub

Sub Procedure2()
    Debug.Print VariableModule ' Accessible (affiche 10)
End Sub
```

**Durée de vie :** Tant que le projet VBA est ouvert

### Variables publiques (globales)

**Déclaration avec Public :**
```vba
' En haut du module
Public VariableGlobale As Integer

' Accessible depuis tous les modules du projet
```

**Utilisation modérée recommandée :**
- Pratique pour partager des données
- Peut créer des dépendances complexes
- À utiliser avec parcimonie

## Conversion entre types

### Conversion automatique (implicite)

**VBA convertit automatiquement quand c'est possible :**
```vba
Dim Nombre As Integer
Dim Texte As String

Nombre = 10
Texte = Nombre                ' Devient "10"

Dim Decimal As Double
Decimal = Nombre              ' Devient 10.0
```

### Conversion explicite (recommandée)

**Fonctions de conversion :**

**CInt() - Conversion vers Integer :**
```vba
Dim x As Integer
x = CInt("25")                ' Convertit "25" en 25
x = CInt(3.7)                 ' Convertit 3.7 en 4 (arrondi)
```

**CDbl() - Conversion vers Double :**
```vba
Dim Prix As Double
Prix = CDbl("19.99")          ' Convertit "19.99" en 19.99
```

**CStr() - Conversion vers String :**
```vba
Dim Texte As String
Texte = CStr(125)             ' Convertit 125 en "125"
```

**CBool() - Conversion vers Boolean :**
```vba
Dim EstVrai As Boolean
EstVrai = CBool(-1)           ' True
EstVrai = CBool(0)            ' False
```

**CDate() - Conversion vers Date :**
```vba
Dim MaDate As Date
MaDate = CDate("15/01/2024")  ' Convertit le texte en date
```

### Gestion des erreurs de conversion

**Problèmes courants :**
```vba
Dim x As Integer
x = CInt("abc")               ' ERREUR : "abc" n'est pas un nombre
x = CInt("100000")            ' ERREUR : Trop grand pour Integer
```

**Solution avec IsNumeric :**
```vba
Dim Texte As String
Dim Nombre As Integer

Texte = "123"
If IsNumeric(Texte) Then
    Nombre = CInt(Texte)      ' Conversion sécurisée
Else
    MsgBox "Ce n'est pas un nombre valide"
End If
```

## Initialisation des variables

### Valeurs par défaut

**VBA initialise automatiquement :**
```vba
Dim x As Integer              ' x = 0
Dim s As String              ' s = "" (chaîne vide)
Dim b As Boolean             ' b = False
Dim d As Date                ' d = 30/12/1899 (date par défaut)
Dim v As Variant             ' v = Empty
```

### Initialisation explicite recommandée

**Bonne pratique :**
```vba
Sub ExempleInitialisation()
    Dim Compteur As Integer
    Dim Nom As String
    Dim EstValide As Boolean

    ' Initialisation explicite
    Compteur = 0
    Nom = ""
    EstValide = False

    ' Maintenant on peut utiliser les variables en toute sécurité
End Sub
```

## Conventions de nommage

### Règles de base

**Caractères autorisés :**
- **Lettres** : a-z, A-Z
- **Chiffres** : 0-9 (mais pas en premier)
- **Underscore** : _ (mais évité en général)

**Caractères interdits :**
- **Espaces** : Utilisez MaVariable, pas Ma Variable
- **Caractères spéciaux** : @, %, $, #, etc.
- **Mots-clés** : Sub, Function, If, etc.

### Conventions recommandées

**CamelCase (recommandé) :**
```vba
Dim nomClient As String           ' Première lettre minuscule
Dim montantTotalHT As Double      ' Mots suivants avec majuscule
Dim numeroCommande As Long
```

**PascalCase (alternatif) :**
```vba
Dim NomClient As String           ' Toutes les premières lettres en majuscule
Dim MontantTotalHT As Double
Dim NumeroCommande As Long
```

**Préfixes de type (style hongrois) :**
```vba
Dim intAge As Integer             ' int pour Integer
Dim strNom As String             ' str pour String
Dim dblPrix As Double            ' dbl pour Double
Dim blnValide As Boolean         ' bln pour Boolean
```

### Noms significatifs

**Mauvais exemples :**
```vba
Dim x As Integer                  ' Que représente x ?
Dim temp As String               ' Temporaire de quoi ?
Dim i As Integer                 ' Peut convenir pour un compteur simple
```

**Bons exemples :**
```vba
Dim ageClient As Integer         ' Clair et précis
Dim nomFichier As String         ' On comprend l'usage
Dim compteurLignes As Integer    ' Même pour un compteur
```

## Erreurs courantes avec les variables

### Variable non déclarée

**Erreur :**
```vba
Option Explicit
Sub Test()
    MonNombre = 10               ' ERREUR : Variable non définie
End Sub
```

**Solution :**
```vba
Option Explicit
Sub Test()
    Dim MonNombre As Integer     ' Déclaration
    MonNombre = 10               ' Maintenant c'est correct
End Sub
```

### Mauvais type de données

**Erreur :**
```vba
Dim Age As Integer
Age = "vingt-cinq"              ' ERREUR : Type incompatible
```

**Solution :**
```vba
Dim Age As Integer
Age = 25                        ' Nombre entier correct
' ou
Dim Age As String
Age = "vingt-cinq"              ' Chaîne de caractères correcte
```

### Dépassement de capacité

**Erreur :**
```vba
Dim PetitNombre As Integer
PetitNombre = 50000             ' ERREUR : Dépassement (max 32767)
```

**Solution :**
```vba
Dim GrandNombre As Long         ' Plus grande capacité
GrandNombre = 50000             ' Maintenant c'est correct
```

## Variables et Excel

### Récupérer des valeurs d'Excel

**Depuis des cellules :**
```vba
Dim Nom As String
Dim Age As Integer
Dim Salaire As Double

Nom = Range("A1").Value         ' Texte depuis A1
Age = Range("B1").Value         ' Nombre depuis B1
Salaire = Range("C1").Value     ' Décimal depuis C1
```

**Vérification de type :**
```vba
Dim Valeur As Variant
Valeur = Range("A1").Value

If IsNumeric(Valeur) Then
    Dim Nombre As Double
    Nombre = CDbl(Valeur)
Else
    Dim Texte As String
    Texte = CStr(Valeur)
End If
```

### Envoyer des valeurs vers Excel

**Vers des cellules :**
```vba
Dim Nom As String
Dim Age As Integer

Nom = "Dupont"
Age = 35

Range("A1").Value = Nom         ' Met "Dupont" en A1
Range("B1").Value = Age         ' Met 35 en B1
```

## Résumé

Les variables sont les fondations du stockage de données en VBA :

**Concepts clés :**
- **Déclaration** : `Dim NomVariable As Type`
- **Affectation** : `Variable = Valeur`
- **Types principaux** : Integer, Long, Double, String, Boolean, Date
- **Option Explicit** : Force la déclaration (recommandé)

**Types de données selon l'usage :**
- **Integer/Long** : Compteurs, quantités, positions
- **Double** : Calculs financiers, mesures, moyennes
- **String** : Texte, noms, descriptions
- **Boolean** : États vrai/faux, conditions
- **Date** : Dates et heures
- **Variant** : Type flexible (utilisation modérée)

**Bonnes pratiques :**
- **Déclarer explicitement** tous les types
- **Noms significatifs** : comprendre le rôle de la variable
- **Initialisation** : Donner une valeur de départ
- **Conversion sécurisée** : Vérifier avant de convertir

**À retenir :**
- **Option Explicit** : Toujours l'utiliser
- **Types appropriés** : Choisir selon les données
- **Nommage cohérent** : Facilite la maintenance
- **Conversion explicite** : Plus sûre que l'automatique

Dans la section suivante, nous découvrirons les constantes, qui permettent de définir des valeurs qui ne changent jamais dans votre programme.

⏭️
