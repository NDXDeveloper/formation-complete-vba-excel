🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.1. Type de données Date

## Qu'est-ce que le type Date en VBA ?

En VBA, le type de données **Date** est spécialement conçu pour stocker des informations de date et d'heure. C'est un type de données fondamental qui permet de manipuler facilement les valeurs temporelles dans vos programmes.

## Comment VBA stocke-t-il les dates ?

### Le principe du stockage numérique

VBA stocke les dates sous forme de **nombres décimaux** (type Double en interne). Cette approche peut sembler surprenante au premier abord, mais elle est très logique :

- La **partie entière** du nombre représente le nombre de jours écoulés depuis le **1er janvier 1900**
- La **partie décimale** représente la fraction de la journée (l'heure)

### Exemples concrets

```vba
' Ces deux lignes sont équivalentes :
Dim maDate As Date  
maDate = #1/1/2024#          ' Notation littérale de date  
maDate = 45292               ' Équivalent numérique  
```

**Comprenons avec des exemples :**

- `1` = 1er janvier 1900
- `2` = 2 janvier 1900
- `45292` = 1er janvier 2024
- `45292.5` = 1er janvier 2024 à midi (12h00)
- `45292.75` = 1er janvier 2024 à 18h00

### Pourquoi ce système est-il pratique ?

Cette représentation numérique offre plusieurs avantages :

- **Calculs simples** : ajouter 7 à une date donne la date dans une semaine
- **Comparaisons faciles** : une date plus récente a une valeur numérique plus élevée
- **Précision** : peut représenter des dates avec une précision à la seconde

## Déclaration de variables de type Date

### Syntaxe de base

```vba
Dim nomVariable As Date
```

### Exemples de déclarations

```vba
' Déclaration simple
Dim dateNaissance As Date  
Dim heureDebut As Date  
Dim echeance As Date  

' Déclaration avec initialisation
Dim aujourdHui As Date  
aujourdHui = Date    ' Date actuelle (sans l'heure)  

Dim maintenant As Date  
maintenant = Now      ' Date et heure actuelles  
```

## Affectation de valeurs aux variables Date

### 1. Utilisation de littéraux de date

VBA utilise une syntaxe spéciale avec des **dièses (#)** pour définir des valeurs de date littérales :

```vba
Dim maDate As Date

' Format américain (obligatoire dans le code VBA)
maDate = #12/25/2024#        ' 25 décembre 2024  
maDate = #1/15/2024 3:30 PM# ' 15 janvier 2024 à 15h30  

' Attention : toujours utiliser le format MM/DD/YYYY dans le code !
```

**Important :** Dans le code VBA, vous devez toujours utiliser le format américain (mois/jour/année) entre les dièses, même si votre système utilise un autre format régional.

### 2. Utilisation de fonctions

```vba
Dim maDate As Date

' Fonctions de base
maDate = Date        ' Date actuelle (sans heure)  
maDate = Time        ' Heure actuelle (sans date)  
maDate = Now         ' Date et heure actuelles  

' Fonctions de construction
maDate = DateSerial(2024, 12, 25)        ' 25 décembre 2024  
maDate = TimeSerial(14, 30, 0)           ' 14h30m00s  
maDate = DateValue("25/12/2024")         ' Conversion depuis texte  
maDate = TimeValue("14:30:00")           ' Conversion depuis texte  
```

### 3. Affectation depuis des cellules Excel

```vba
Dim maDate As Date

' Récupération depuis une cellule
maDate = Range("A1").Value  
maDate = Cells(1, 1).Value  

' Avec vérification du type
If IsDate(Range("A1").Value) Then
    maDate = Range("A1").Value
End If
```

## Plages de valeurs supportées

Le type Date en VBA peut représenter des dates dans une plage très large :

- **Date minimale** : 1er janvier 100 après J.-C.
- **Date maximale** : 31 décembre 9999 après J.-C.
- **Précision temporelle** : jusqu'à la seconde

```vba
Dim dateAncienne As Date  
Dim dateFuture As Date  

dateAncienne = #1/1/100#      ' 1er janvier de l'an 100  
dateFuture = #12/31/9999#     ' 31 décembre 9999  
```

## Affichage et formatage des dates

### Affichage par défaut

Lorsque vous affichez une variable Date, VBA utilise le format régional de votre système :

```vba
Dim maDate As Date  
maDate = #12/25/2024 2:30 PM#  

Debug.Print maDate           ' Affiche selon format système  
MsgBox maDate                ' Affiche selon format système  
```

### Formatage personnalisé

Vous pouvez contrôler l'affichage avec la fonction `Format` :

```vba
Dim maDate As Date  
maDate = #12/25/2024 2:30 PM#  

Debug.Print Format(maDate, "dd/mm/yyyy")           ' 25/12/2024  
Debug.Print Format(maDate, "dddd dd mmmm yyyy")    ' mercredi 25 décembre 2024  
Debug.Print Format(maDate, "hh:nn:ss")             ' 14:30:00  
Debug.Print Format(maDate, "dd/mm/yyyy hh:nn")     ' 25/12/2024 14:30  
```

## Valeurs spéciales

### Date vide (zéro)

Une variable Date non initialisée a la valeur `0`, qui correspond au **30 décembre 1899** :

```vba
Dim maDate As Date  
Debug.Print maDate    ' Affiche 30/12/1899 00:00:00  

' Test si une date est vide
If maDate = 0 Then
    Debug.Print "Date non définie"
End If
```

### Partie date seule ou heure seule

```vba
' Date sans heure (heure = 00:00:00)
Dim dateSeule As Date  
dateSeule = #12/25/2024#  

' Heure sans date (date = 30/12/1899)
Dim heureSeule As Date  
heureSeule = #2:30 PM#  
```

## Points importants à retenir

**Format dans le code** : Utilisez toujours le format américain (MM/DD/YYYY) entre les dièses dans votre code VBA, indépendamment des paramètres régionaux de votre système.

**Stockage numérique** : Comprendre que les dates sont des nombres vous aidera dans les calculs et comparaisons.

**Précision** : Le type Date peut stocker des informations très précises, mais attention aux arrondis lors de calculs complexes.

**Initialisation** : Une variable Date non initialisée vaut 0 (30 décembre 1899), pas une valeur vide ou nulle.

**Compatibilité Excel** : Les dates VBA sont directement compatibles avec les dates Excel, facilitant les échanges entre votre code et les cellules.

---

*Le type Date est la fondation de toute manipulation temporelle en VBA. Sa compréhension est essentielle pour créer des applications robustes travaillant avec des données temporelles.*

⏭️
