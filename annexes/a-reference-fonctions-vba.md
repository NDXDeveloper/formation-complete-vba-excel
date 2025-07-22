🔝 Retour au [Sommaire](/SOMMAIRE.md)

# A. Référence des fonctions VBA

## Introduction

Cette annexe présente les fonctions VBA les plus couramment utilisées, organisées par catégorie pour faciliter votre recherche. Chaque fonction est accompagnée de sa syntaxe, d'une description claire et d'exemples simples pour vous aider à comprendre son utilisation.

**Comment lire cette référence :**
- **Syntaxe** : La façon correcte d'écrire la fonction
- **Description** : Ce que fait la fonction
- **Paramètres** : Les informations que vous devez fournir à la fonction
- **Retour** : Le type de résultat que la fonction vous donne
- **Exemple** : Un exemple concret d'utilisation

---

## 1. Fonctions de chaînes de caractères (String)

### Len()
**Syntaxe :** `Len(chaîne)`
**Description :** Retourne le nombre de caractères dans une chaîne de texte
**Paramètres :** chaîne - le texte à mesurer
**Retour :** Nombre entier
**Exemple :**
```vba
Dim longueur As Integer
longueur = Len("Bonjour") ' Résultat : 7
```

### Left()
**Syntaxe :** `Left(chaîne, nombre)`
**Description :** Extrait les caractères de gauche d'une chaîne
**Paramètres :** chaîne - le texte source, nombre - combien de caractères à extraire
**Retour :** Chaîne de caractères
**Exemple :**
```vba
Dim debut As String
debut = Left("Bonjour", 3) ' Résultat : "Bon"
```

### Right()
**Syntaxe :** `Right(chaîne, nombre)`
**Description :** Extrait les caractères de droite d'une chaîne
**Paramètres :** chaîne - le texte source, nombre - combien de caractères à extraire
**Retour :** Chaîne de caractères
**Exemple :**
```vba
Dim fin As String
fin = Right("Bonjour", 4) ' Résultat : "jour"
```

### Mid()
**Syntaxe :** `Mid(chaîne, position, [longueur])`
**Description :** Extrait une partie d'une chaîne à partir d'une position donnée
**Paramètres :** chaîne - le texte source, position - où commencer, longueur - combien de caractères (optionnel)
**Retour :** Chaîne de caractères
**Exemple :**
```vba
Dim milieu As String
milieu = Mid("Bonjour", 3, 2) ' Résultat : "nj"
```

### UCase()
**Syntaxe :** `UCase(chaîne)`
**Description :** Convertit une chaîne en majuscules
**Paramètres :** chaîne - le texte à convertir
**Retour :** Chaîne de caractères en majuscules
**Exemple :**
```vba
Dim majuscule As String
majuscule = UCase("bonjour") ' Résultat : "BONJOUR"
```

### LCase()
**Syntaxe :** `LCase(chaîne)`
**Description :** Convertit une chaîne en minuscules
**Paramètres :** chaîne - le texte à convertir
**Retour :** Chaîne de caractères en minuscules
**Exemple :**
```vba
Dim minuscule As String
minuscule = LCase("BONJOUR") ' Résultat : "bonjour"
```

### Trim()
**Syntaxe :** `Trim(chaîne)`
**Description :** Supprime les espaces au début et à la fin d'une chaîne
**Paramètres :** chaîne - le texte à nettoyer
**Retour :** Chaîne de caractères sans espaces inutiles
**Exemple :**
```vba
Dim propre As String
propre = Trim("  Bonjour  ") ' Résultat : "Bonjour"
```

### Replace()
**Syntaxe :** `Replace(chaîne, recherche, remplacement)`
**Description :** Remplace toutes les occurrences d'un texte par un autre
**Paramètres :** chaîne - le texte source, recherche - texte à remplacer, remplacement - nouveau texte
**Retour :** Chaîne de caractères modifiée
**Exemple :**
```vba
Dim nouveau As String
nouveau = Replace("Bon jour", " ", "") ' Résultat : "Bonjour"
```

---

## 2. Fonctions mathématiques

### Abs()
**Syntaxe :** `Abs(nombre)`
**Description :** Retourne la valeur absolue d'un nombre (sans le signe)
**Paramètres :** nombre - le nombre à traiter
**Retour :** Nombre positif
**Exemple :**
```vba
Dim resultat As Double
resultat = Abs(-15.5) ' Résultat : 15.5
```

### Round()
**Syntaxe :** `Round(nombre, [décimales])`
**Description :** Arrondit un nombre au nombre de décimales spécifié
**Paramètres :** nombre - le nombre à arrondir, décimales - nombre de décimales (optionnel, défaut = 0)
**Retour :** Nombre arrondi
**Exemple :**
```vba
Dim arrondi As Double
arrondi = Round(15.678, 2) ' Résultat : 15.68
```

### Int()
**Syntaxe :** `Int(nombre)`
**Description :** Retourne la partie entière d'un nombre
**Paramètres :** nombre - le nombre à traiter
**Retour :** Nombre entier
**Exemple :**
```vba
Dim entier As Integer
entier = Int(15.8) ' Résultat : 15
```

### Sqr()
**Syntaxe :** `Sqr(nombre)`
**Description :** Calcule la racine carrée d'un nombre
**Paramètres :** nombre - le nombre positif dont on veut la racine
**Retour :** Racine carrée du nombre
**Exemple :**
```vba
Dim racine As Double
racine = Sqr(16) ' Résultat : 4
```

### Rnd()
**Syntaxe :** `Rnd([nombre])`
**Description :** Génère un nombre aléatoire entre 0 et 1
**Paramètres :** nombre - optionnel, influence la génération
**Retour :** Nombre décimal entre 0 et 1
**Exemple :**
```vba
Dim aleatoire As Double
aleatoire = Rnd() ' Résultat : par exemple 0.734521
```

---

## 3. Fonctions de date et heure

### Now()
**Syntaxe :** `Now()`
**Description :** Retourne la date et l'heure actuelles
**Paramètres :** Aucun
**Retour :** Date et heure
**Exemple :**
```vba
Dim maintenant As Date
maintenant = Now() ' Résultat : 22/07/2025 14:30:15
```

### Date()
**Syntaxe :** `Date()`
**Description :** Retourne la date actuelle
**Paramètres :** Aucun
**Retour :** Date
**Exemple :**
```vba
Dim aujourd_hui As Date
aujourd_hui = Date() ' Résultat : 22/07/2025
```

### Time()
**Syntaxe :** `Time()`
**Description :** Retourne l'heure actuelle
**Paramètres :** Aucun
**Retour :** Heure
**Exemple :**
```vba
Dim heure_actuelle As Date
heure_actuelle = Time() ' Résultat : 14:30:15
```

### Year()
**Syntaxe :** `Year(date)`
**Description :** Extrait l'année d'une date
**Paramètres :** date - la date dont on veut l'année
**Retour :** Nombre entier (année)
**Exemple :**
```vba
Dim annee As Integer
annee = Year(Date()) ' Résultat : 2025
```

### Month()
**Syntaxe :** `Month(date)`
**Description :** Extrait le mois d'une date
**Paramètres :** date - la date dont on veut le mois
**Retour :** Nombre entier (1 à 12)
**Exemple :**
```vba
Dim mois As Integer
mois = Month(Date()) ' Résultat : 7 (juillet)
```

### Day()
**Syntaxe :** `Day(date)`
**Description :** Extrait le jour d'une date
**Paramètres :** date - la date dont on veut le jour
**Retour :** Nombre entier (1 à 31)
**Exemple :**
```vba
Dim jour As Integer
jour = Day(Date()) ' Résultat : 22
```

### DateAdd()
**Syntaxe :** `DateAdd(intervalle, nombre, date)`
**Description :** Ajoute ou soustrait un intervalle de temps à une date
**Paramètres :** intervalle - type ("d"=jour, "m"=mois, "yyyy"=année), nombre - quantité, date - date de base
**Retour :** Nouvelle date
**Exemple :**
```vba
Dim nouvelle_date As Date
nouvelle_date = DateAdd("d", 10, Date()) ' Ajoute 10 jours à aujourd'hui
```

---

## 4. Fonctions de conversion

### CStr()
**Syntaxe :** `CStr(expression)`
**Description :** Convertit une expression en chaîne de caractères
**Paramètres :** expression - valeur à convertir
**Retour :** Chaîne de caractères
**Exemple :**
```vba
Dim texte As String
texte = CStr(123) ' Résultat : "123"
```

### CInt()
**Syntaxe :** `CInt(expression)`
**Description :** Convertit une expression en nombre entier
**Paramètres :** expression - valeur à convertir
**Retour :** Nombre entier
**Exemple :**
```vba
Dim nombre As Integer
nombre = CInt("123") ' Résultat : 123
```

### CDbl()
**Syntaxe :** `CDbl(expression)`
**Description :** Convertit une expression en nombre décimal
**Paramètres :** expression - valeur à convertir
**Retour :** Nombre décimal
**Exemple :**
```vba
Dim decimal As Double
decimal = CDbl("123.45") ' Résultat : 123.45
```

### CDate()
**Syntaxe :** `CDate(expression)`
**Description :** Convertit une expression en date
**Paramètres :** expression - valeur à convertir (texte ou nombre)
**Retour :** Date
**Exemple :**
```vba
Dim ma_date As Date
ma_date = CDate("22/07/2025") ' Résultat : 22/07/2025
```

### Val()
**Syntaxe :** `Val(chaîne)`
**Description :** Convertit les chiffres au début d'une chaîne en nombre
**Paramètres :** chaîne - texte contenant des chiffres
**Retour :** Nombre
**Exemple :**
```vba
Dim numero As Double
numero = Val("123abc") ' Résultat : 123
```

---

## 5. Fonctions de test et logiques

### IsNumeric()
**Syntaxe :** `IsNumeric(expression)`
**Description :** Vérifie si une expression peut être convertie en nombre
**Paramètres :** expression - valeur à tester
**Retour :** True (vrai) ou False (faux)
**Exemple :**
```vba
Dim est_nombre As Boolean
est_nombre = IsNumeric("123") ' Résultat : True
```

### IsDate()
**Syntaxe :** `IsDate(expression)`
**Description :** Vérifie si une expression peut être convertie en date
**Paramètres :** expression - valeur à tester
**Retour :** True (vrai) ou False (faux)
**Exemple :**
```vba
Dim est_date As Boolean
est_date = IsDate("22/07/2025") ' Résultat : True
```

### IsEmpty()
**Syntaxe :** `IsEmpty(expression)`
**Description :** Vérifie si une variable est vide (non initialisée)
**Paramètres :** expression - variable à tester
**Retour :** True (vrai) ou False (faux)
**Exemple :**
```vba
Dim ma_variable As Variant
Dim est_vide As Boolean
est_vide = IsEmpty(ma_variable) ' Résultat : True
```

### IsNull()
**Syntaxe :** `IsNull(expression)`
**Description :** Vérifie si une expression contient la valeur Null
**Paramètres :** expression - valeur à tester
**Retour :** True (vrai) ou False (faux)
**Exemple :**
```vba
Dim est_null As Boolean
est_null = IsNull(Null) ' Résultat : True
```

---

## 6. Fonctions d'interaction

### MsgBox()
**Syntaxe :** `MsgBox(message, [boutons], [titre])`
**Description :** Affiche une boîte de message à l'utilisateur
**Paramètres :** message - texte à afficher, boutons - type de boutons (optionnel), titre - titre de la fenêtre (optionnel)
**Retour :** Indique quel bouton a été cliqué
**Exemple :**
```vba
MsgBox "Bonjour !", , "Mon Application"
```

### InputBox()
**Syntaxe :** `InputBox(message, [titre], [valeur_défaut])`
**Description :** Demande une saisie à l'utilisateur
**Paramètres :** message - question à poser, titre - titre de la fenêtre (optionnel), valeur_défaut - valeur pré-remplie (optionnel)
**Retour :** Texte saisi par l'utilisateur
**Exemple :**
```vba
Dim nom As String
nom = InputBox("Quel est votre nom ?", "Saisie")
```

---

## 7. Fonctions de formatage

### Format()
**Syntaxe :** `Format(expression, [format])`
**Description :** Formate une valeur selon un modèle spécifié
**Paramètres :** expression - valeur à formater, format - modèle de formatage
**Retour :** Chaîne formatée
**Exemple :**
```vba
Dim prix As String
prix = Format(1234.56, "0.00 €") ' Résultat : "1234.56 €"
```

---

## Conseils d'utilisation

### Pour les débutants :
1. **Commencez par les fonctions de base** : Len, Left, Right, Mid pour les chaînes
2. **Testez chaque fonction** dans une procédure simple avant de l'utiliser dans un projet
3. **Utilisez MsgBox** pour afficher les résultats et vérifier que vos fonctions marchent
4. **N'hésitez pas à combiner** plusieurs fonctions simples plutôt qu'une fonction complexe

### Erreurs courantes à éviter :
- **Oublier les guillemets** autour des chaînes de caractères
- **Confondre les paramètres** : vérifiez toujours l'ordre dans la syntaxe
- **Ne pas gérer les erreurs** : utilisez IsNumeric avant CInt par exemple
- **Mélanger les types** : une date n'est pas un texte, un nombre n'est pas une chaîne

Cette référence couvre les fonctions VBA les plus utilisées au quotidien. Gardez-la à portée de main lors de vos développements !

⏭️
