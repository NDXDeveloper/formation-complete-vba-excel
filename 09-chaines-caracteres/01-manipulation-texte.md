🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.1. Manipulation de texte

## Introduction à la manipulation de texte

La manipulation de texte consiste à transformer, modifier ou analyser des chaînes de caractères selon nos besoins. En VBA, nous disposons de nombreuses techniques pour effectuer ces opérations, depuis les plus simples jusqu'aux plus sophistiquées.

## Opérations fondamentales

### Concaténation (assemblage de chaînes)

La concaténation est l'opération qui permet de joindre plusieurs chaînes de caractères pour en former une seule.

#### L'opérateur & (recommandé)
```vba
Dim prenom As String  
Dim nom As String  
Dim nomComplet As String  

prenom = "Marie"  
nom = "Martin"  
nomComplet = prenom & " " & nom  
' Résultat : "Marie Martin"
```

#### L'opérateur + (déconseillé)
```vba
' Fonctionne mais moins fiable
nomComplet = prenom + " " + nom
```

**Pourquoi préférer & ?** L'opérateur & est spécifiquement conçu pour les chaînes, tandis que + peut créer des ambiguïtés avec les opérations numériques.

#### Concaténation avec des nombres
```vba
Dim age As Integer  
Dim message As String  

age = 25  
message = "J'ai " & age & " ans"  
' Résultat : "J'ai 25 ans"
```

### Suppression des espaces

Les espaces indésirables sont un problème fréquent lors du traitement de données. VBA offre trois fonctions principales :

#### Trim() - Supprime les espaces en début et fin
```vba
Dim texte As String  
Dim texteNettoye As String  

texte = "   Bonjour le monde   "  
texteNettoye = Trim(texte)  
' Résultat : "Bonjour le monde"
```

#### LTrim() - Supprime uniquement les espaces à gauche
```vba
Dim texte As String  
texte = "   Bonjour   "  
Dim resultat As String  
resultat = LTrim(texte)  
' Résultat : "Bonjour   "
```

#### RTrim() - Supprime uniquement les espaces à droite
```vba
Dim texte As String  
texte = "   Bonjour   "  
Dim resultat As String  
resultat = RTrim(texte)  
' Résultat : "   Bonjour"
```

### Changement de casse (majuscules/minuscules)

#### UCase() - Convertit en majuscules
```vba
Dim texte As String  
texte = "bonjour VBA"  
Debug.Print UCase(texte)  ' "BONJOUR VBA"  
```

#### LCase() - Convertit en minuscules
```vba
Dim texte As String  
texte = "BONJOUR VBA"  
Debug.Print LCase(texte)  ' "bonjour vba"  
```

#### StrConv() - Conversion avancée
```vba
Dim texte As String  
texte = "bonjour vba"  

' Première lettre de chaque mot en majuscule
Debug.Print StrConv(texte, vbProperCase)  ' "Bonjour Vba"

' Conversion en majuscules (équivalent à UCase)
Debug.Print StrConv(texte, vbUpperCase)   ' "BONJOUR VBA"
```

## Remplacement de texte

### Replace() - Fonction de base
La fonction Replace permet de remplacer toutes les occurrences d'une chaîne par une autre.

#### Syntaxe de base
```vba
Replace(chaîne_source, chaîne_à_chercher, chaîne_de_remplacement)
```

#### Exemples pratiques
```vba
Dim texte As String  
Dim nouveau As String  

texte = "Bonjour le monde, bonjour VBA"
' Remplace tous les "bonjour" par "salut"
nouveau = Replace(texte, "bonjour", "salut")
' Résultat : "Bonjour le monde, salut VBA"

' Attention à la casse !
nouveau = Replace(LCase(texte), "bonjour", "salut")
' Résultat : "salut le monde, salut vba"
```

#### Paramètres avancés de Replace
```vba
' Syntaxe complète
Replace(expression, find, replace, start, count, compare)
```

```vba
Dim texte As String  
texte = "abcABCabc"  

' Remplace avec comparaison insensible à la casse
Dim resultat As String  
resultat = Replace(texte, "abc", "XYZ", 1, 2, vbTextCompare)  
' Résultat : "XYZXYZXYZ"
```

### Suppression de caractères spécifiques
```vba
' Supprimer tous les tirets d'un numéro de téléphone
Dim telephone As String  
telephone = "01-23-45-67-89"  
Dim telephoneClean As String  
telephoneClean = Replace(telephone, "-", "")  
' Résultat : "0123456789"

' Supprimer les espaces multiples
Dim texte As String  
texte = "Mot1    Mot2     Mot3"  
' Première étape : remplacer les espaces multiples par un seul
Do While InStr(texte, "  ") > 0
    texte = Replace(texte, "  ", " ")
Loop
' Résultat : "Mot1 Mot2 Mot3"
```

## Insertion et suppression de caractères

### Insertion de caractères
```vba
' Insérer du texte au milieu d'une chaîne
Dim texte As String  
Dim position As Integer  
Dim nouveau As String  

texte = "BonjourVBA"  
position = 8  ' Après "Bonjour"  

' Technique : diviser, insérer, recombiner
nouveau = Left(texte, position - 1) & " " & Mid(texte, position)
' Résultat : "Bonjour VBA"
```

### Suppression de caractères
```vba
' Supprimer des caractères à une position donnée
Dim texte As String  
Dim debut As Integer  
Dim longueur As Integer  
Dim resultat As String  

texte = "Bonjour123VBA"  
debut = 8     ' Position du premier caractère à supprimer  
longueur = 3  ' Nombre de caractères à supprimer  
resultat = Left(texte, debut - 1) & Mid(texte, debut + longueur)  
' Résultat : "BonjourVBA"
```

## Techniques de nettoyage courantes

### Nettoyage complet d'une chaîne
```vba
Function NettoyerTexte(texte As String) As String
    Dim resultat As String
    resultat = texte

    ' Supprimer les espaces en début et fin
    resultat = Trim(resultat)

    ' Remplacer les espaces multiples par un seul
    Do While InStr(resultat, "  ") > 0
        resultat = Replace(resultat, "  ", " ")
    Loop

    ' Supprimer les caractères de contrôle indésirables
    resultat = Replace(resultat, vbTab, " ")
    resultat = Replace(resultat, vbCrLf, " ")
    resultat = Replace(resultat, vbCr, " ")
    resultat = Replace(resultat, vbLf, " ")

    NettoyerTexte = resultat
End Function
```

### Standardisation de formats
```vba
' Standardiser un nom (première lettre majuscule)
Function StandardiserNom(nom As String) As String
    Dim nomNettoye As String
    nomNettoye = Trim(LCase(nom))
    If Len(nomNettoye) > 0 Then
        StandardiserNom = UCase(Left(nomNettoye, 1)) & Mid(nomNettoye, 2)
    Else
        StandardiserNom = ""
    End If
End Function

' Utilisation
Dim nom As String  
nom = "  DUPONT  "  
Dim nomStandard As String  
nomStandard = StandardiserNom(nom)  
' Résultat : "Dupont"
```

## Manipulation de caractères individuels

### Accès aux caractères
```vba
Dim texte As String  
texte = "VBA"  

Debug.Print Left(texte, 1)    ' "V"  
Debug.Print Mid(texte, 2, 1)  ' "B"  
Debug.Print Right(texte, 1)   ' "A"  
```

### Codes ASCII
```vba
' Obtenir le code ASCII d'un caractère
Debug.Print Asc("A")   ' 65

' Convertir un code ASCII en caractère
Debug.Print Chr(65)     ' "A"
```

## Conseils pratiques pour débutants

### 1. Toujours vérifier si la chaîne est vide
```vba
Dim texte As String  
texte = "..."  

If Len(texte) > 0 Then
    ' Effectuer la manipulation
    texte = UCase(texte)
End If
```

### 2. Prévoir les cas particuliers
```vba
' Attention aux chaînes nulles ou vides
If texte <> "" And Not IsNull(texte) Then
    ' Traitement sécurisé
End If
```

### 3. Tester avec des données réelles
Testez toujours vos manipulations avec des données variées :
- Chaînes vides
- Chaînes avec espaces en début/fin
- Chaînes très longues
- Chaînes contenant des caractères spéciaux

### 4. Documenter les transformations
```vba
' Nettoie et standardise un nom de client
' Entrée : "  jean-PAUL dupont  "
' Sortie : "Jean-Paul Dupont"
Function StandardiserNomClient(nom As String) As String
    ' ... code de la fonction
End Function
```

La manipulation de texte est un art qui s'améliore avec la pratique. Ces techniques de base vous permettront de traiter efficacement la majorité des situations courantes en VBA.

⏭️ [Fonctions String (Len, Mid, Left, Right)](/09-chaines-caracteres/02-fonctions-string.md)
