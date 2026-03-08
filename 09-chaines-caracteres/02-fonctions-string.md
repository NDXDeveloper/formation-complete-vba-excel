🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.2. Fonctions String (Len, Mid, Left, Right)

## Introduction aux fonctions String

Les fonctions de manipulation de chaînes sont les outils de base pour extraire, analyser et transformer du texte en VBA. Ces quatre fonctions - Len, Mid, Left et Right - sont probablement les plus utilisées et constituent la boîte à outils essentielle de tout développeur VBA.

Pensez à ces fonctions comme à des "ciseaux intelligents" qui vous permettent de découper précisément les parties de texte dont vous avez besoin.

## Fonction Len() - Mesurer la longueur

### Principe de base
La fonction `Len()` retourne le nombre de caractères contenus dans une chaîne de caractères.

### Syntaxe
```vba
Len(chaîne_de_caractères)
```

### Exemples simples
```vba
Dim texte As String  
Dim longueur As Integer  

texte = "Bonjour"  
longueur = Len(texte)  
' Résultat : 7

texte = "VBA"  
longueur = Len(texte)  
' Résultat : 3

texte = ""  ' Chaîne vide  
longueur = Len(texte)  
' Résultat : 0
```

### Cas particuliers avec Len()
```vba
' Les espaces comptent !
Dim texte1 As String  
texte1 = "Bon jour"  
Debug.Print Len(texte1)  ' Résultat : 8 (espace inclus)  

' Les caractères spéciaux comptent aussi
Dim texte2 As String  
texte2 = "Bonjour!"  
Debug.Print Len(texte2)  ' Résultat : 8 (le ! compte)  

' Tabulations et retours à la ligne
Dim texte3 As String  
texte3 = "Ligne1" & vbCrLf & "Ligne2"  
Debug.Print Len(texte3)  ' Résultat : 14 (vbCrLf = 2 caractères)  
```

### Utilisations pratiques de Len()
```vba
' Vérifier si une chaîne est vide
If Len(monTexte) = 0 Then
    MsgBox "Le texte est vide !"
End If

' Validation de longueur (ex: mot de passe)
If Len(motDePasse) < 8 Then
    MsgBox "Le mot de passe doit contenir au moins 8 caractères"
End If

' Compter les caractères d'une cellule Excel
Dim longueurCellule As Integer  
longueurCellule = Len(Range("A1").Value)  
```

## Fonction Left() - Extraire depuis la gauche

### Principe de base
La fonction `Left()` extrait un nombre spécifié de caractères depuis le début (gauche) d'une chaîne.

### Syntaxe
```vba
Left(chaîne_de_caractères, nombre_de_caractères)
```

### Exemples simples
```vba
Dim texte As String  
Dim debut As String  

texte = "Bonjour VBA"  
debut = Left(texte, 3)  
' Résultat : "Bon"

debut = Left(texte, 7)
' Résultat : "Bonjour"

debut = Left(texte, 1)
' Résultat : "B"
```

### Gestion des cas limites
```vba
Dim texte As String  
texte = "VBA"  

' Demander plus de caractères que disponible
Dim resultat As String  
resultat = Left(texte, 10)  
' Résultat : "VBA" (pas d'erreur, retourne toute la chaîne)

' Demander 0 caractères
resultat = Left(texte, 0)
' Résultat : "" (chaîne vide)
```

### Utilisations pratiques de Left()
```vba
' Extraire un préfixe
Dim codeArticle As String  
codeArticle = "ART001_Ordinateur"  
Dim prefixe As String  
prefixe = Left(codeArticle, 6)  ' "ART001"  

' Extraire les premiers mots
Dim nomComplet As String  
nomComplet = "Jean-Pierre Dupont"  
Dim prenom As String  
prenom = Left(nomComplet, 11)  ' "Jean-Pierre"  

' Vérifier le début d'une chaîne
If Left(email, 4) = "www." Then
    MsgBox "Ce n'est pas un email valide"
End If
```

## Fonction Right() - Extraire depuis la droite

### Principe de base
La fonction `Right()` extrait un nombre spécifié de caractères depuis la fin (droite) d'une chaîne.

### Syntaxe
```vba
Right(chaîne_de_caractères, nombre_de_caractères)
```

### Exemples simples
```vba
Dim texte As String  
Dim fin As String  

texte = "Bonjour VBA"  
fin = Right(texte, 3)  
' Résultat : "VBA"

fin = Right(texte, 7)
' Résultat : "our VBA"

fin = Right(texte, 1)
' Résultat : "A"
```

### Utilisations pratiques de Right()
```vba
' Extraire une extension de fichier
Dim nomFichier As String  
nomFichier = "document.xlsx"  
Dim extension As String  
extension = Right(nomFichier, 4)  ' "xlsx"  

' Extraire les derniers chiffres
Dim numeroFacture As String  
numeroFacture = "FACT_2024_001234"  
Dim numero As String  
numero = Right(numeroFacture, 6)  ' "001234"  

' Vérifier la fin d'une chaîne
If Right(email, 4) = ".com" Then
    MsgBox "Adresse email .com détectée"
End If
```

## Fonction Mid() - Extraire depuis le milieu

### Principe de base
La fonction `Mid()` est la plus flexible : elle extrait des caractères à partir d'une position donnée.

### Syntaxes
```vba
' Extraire à partir d'une position jusqu'à la fin
Mid(chaîne_de_caractères, position_de_début)

' Extraire un nombre précis de caractères
Mid(chaîne_de_caractères, position_de_début, nombre_de_caractères)
```

### Exemples simples
```vba
Dim texte As String  
texte = "Bonjour VBA"  

' À partir de la position 5 jusqu'à la fin
Dim milieu1 As String  
milieu1 = Mid(texte, 5)  
' Résultat : "our VBA"

' À partir de la position 5, prendre 3 caractères
Dim milieu2 As String  
milieu2 = Mid(texte, 5, 3)  
' Résultat : "our"

' À partir de la position 9, prendre 3 caractères
Dim milieu3 As String  
milieu3 = Mid(texte, 9, 3)  
' Résultat : "VBA"
```

### Compter les positions (important pour les débutants)
```vba
' Position :  1234567891011
Dim texte As String  
texte = "Bonjour VBA"  
'              B = position 1
'              o = position 2
'              n = position 3
'              (espace) = position 8
'              V = position 9
```

### Utilisations avancées de Mid()
```vba
' Extraire le milieu d'un nom de fichier
Dim cheminComplet As String  
cheminComplet = "C:\Documents\MonFichier.xlsx"  
Dim nomSeul As String  
nomSeul = Mid(cheminComplet, 14, 10)  ' "MonFichier"  

' Extraire une partie d'un numéro
Dim numeroTelephone As String  
numeroTelephone = "01.23.45.67.89"  
Dim indicatif As String  
indicatif = Mid(numeroTelephone, 1, 2)   ' "01"  
Dim numero1 As String  
numero1 = Mid(numeroTelephone, 4, 2)     ' "23"  
Dim numero2 As String  
numero2 = Mid(numeroTelephone, 7, 2)     ' "45"  
```

## Combinaison des fonctions

### Techniques courantes de combinaison

#### Extraire le nom de fichier sans extension
```vba
Function ExtraireNomSansExtension(cheminComplet As String) As String
    Dim nomAvecExtension As String
    Dim positionPoint As Integer

    ' Extraire juste le nom du fichier (après le dernier \)
    ' Pour cet exemple, supposons que nous ayons juste "document.xlsx"
    nomAvecExtension = "document.xlsx"

    ' Trouver la position du point
    positionPoint = InStr(nomAvecExtension, ".")

    ' Extraire tout ce qui est avant le point
    If positionPoint > 0 Then
        ExtraireNomSansExtension = Left(nomAvecExtension, positionPoint - 1)
    Else
        ExtraireNomSansExtension = nomAvecExtension
    End If
End Function
' Résultat : "document"
```

#### Extraire le prénom et nom séparément
```vba
Sub SeparerNomPrenom()
    Dim nomComplet As String
    nomComplet = "Marie Dubois"
    Dim positionEspace As Integer
    Dim prenom As String
    Dim nom As String

    ' Trouver la position de l'espace
    positionEspace = InStr(nomComplet, " ")

    If positionEspace > 0 Then
        prenom = Left(nomComplet, positionEspace - 1)  ' "Marie"
        nom = Mid(nomComplet, positionEspace + 1)       ' "Dubois"
    End If
End Sub
```

#### Valider un format de code postal français
```vba
Function ValiderCodePostal(code As String) As Boolean
    ' Un code postal français : 5 chiffres
    If Len(code) = 5 Then
        ' Vérifier que tous les caractères sont des chiffres
        Dim i As Integer
        For i = 1 To 5
            Dim caractere As String
            caractere = Mid(code, i, 1)
            If caractere < "0" Or caractere > "9" Then
                ValiderCodePostal = False
                Exit Function
            End If
        Next i
        ValiderCodePostal = True
    Else
        ValiderCodePostal = False
    End If
End Function
```

## Bonnes pratiques et pièges à éviter

### 1. Vérifier la longueur avant extraction
```vba
' Left() ne cause pas d'erreur si la chaîne est trop courte,
' mais le résultat peut ne pas être celui attendu.
' Exemple : Left("VBA", 10) retourne "VBA" (pas 10 caractères)

' Pour garantir un résultat de longueur précise, vérifier d'abord :
Dim resultat As String  
If Len(texte) >= 10 Then  
    resultat = Left(texte, 10)
Else
    resultat = texte  ' ou une autre logique appropriée
End If
```

### 2. Attention aux positions négatives ou nulles
```vba
' Mid() avec une position <= 0 peut causer des erreurs
' Toujours s'assurer que la position >= 1
If position >= 1 And position <= Len(texte) Then
    resultat = Mid(texte, position, 3)
End If
```

### 3. Gérer les chaînes vides
```vba
' Toujours vérifier si la chaîne n'est pas vide
If Len(texte) > 0 Then
    ' Effectuer les opérations
    resultat = Left(texte, 5)
End If
```

### 4. Documentation claire des positions
```vba
' Bien documenter d'où viennent les positions
Dim codeClient As String  
codeClient = "CLI_2024_001234"  
'                          123456789012345
'                          |    |    |
'                          |    |    +-- Numéro client (position 11-16)
'                          |    +-- Année (position 5-8)
'                          +-- Préfixe (position 1-3)

Dim prefixe As String  
prefixe = Left(codeClient, 3)        ' "CLI"  
Dim annee As String  
annee = Mid(codeClient, 5, 4)        ' "2024"  
Dim numero As String  
numero = Right(codeClient, 6)        ' "001234"  
```

## Fonctions utilitaires pratiques

### Fonction pour extraire l'initiale
```vba
Function ExtraireInitiale(nom As String) As String
    If Len(nom) > 0 Then
        ExtraireInitiale = UCase(Left(nom, 1))
    Else
        ExtraireInitiale = ""
    End If
End Function
```

### Fonction pour extraire les X derniers caractères en sécurité
```vba
Function DerniersCaracteres(texte As String, nombre As Integer) As String
    If Len(texte) >= nombre Then
        DerniersCaracteres = Right(texte, nombre)
    Else
        DerniersCaracteres = texte
    End If
End Function
```

### Fonction pour tronquer avec points de suspension
```vba
Function TronquerTexte(texte As String, longueurMax As Integer) As String
    If Len(texte) <= longueurMax Then
        TronquerTexte = texte
    Else
        TronquerTexte = Left(texte, longueurMax - 3) & "..."
    End If
End Function
' Exemple : TronquerTexte("Texte très long", 10) = "Texte tr..."
```

Ces fonctions String sont vos alliées quotidiennes en VBA. Maîtrisez-les bien, car elles constituent la base de nombreuses opérations plus complexes sur les chaînes de caractères.

⏭️
