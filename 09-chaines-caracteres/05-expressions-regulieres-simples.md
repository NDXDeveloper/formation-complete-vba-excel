🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.5. Expressions régulières simples

## Introduction aux expressions régulières

Les expressions régulières (regex) sont comme des "filtres intelligents" pour le texte. Imaginez que vous cherchez tous les numéros de téléphone dans un document, ou que vous voulez vérifier si une adresse email a le bon format. Les expressions régulières vous permettent de définir des "motifs" (patterns) de recherche très précis.

Pour les débutants, pensez aux expressions régulières comme à des "règles de reconnaissance" : au lieu de chercher exactement "01.23.45.67.89", vous pouvez chercher "n'importe quel numéro qui ressemble à un téléphone français".

**Note importante :** VBA ne dispose pas d'expressions régulières intégrées comme d'autres langages. Nous devrons utiliser l'objet RegExp de Microsoft qui nécessite une référence externe.

## Configuration des expressions régulières en VBA

### Activation de la référence
Pour utiliser les expressions régulières en VBA, vous devez d'abord activer une référence :

1. Dans l'éditeur VBA : Outils → Références
2. Cocher "Microsoft VBScript Regular Expressions 5.5"
3. Cliquer OK

### Méthode alternative (sans référence)
```vba
' Création dynamique de l'objet RegExp
Dim regex As Object  
Set regex = CreateObject("VBScript.RegExp")  
```

## Structure de base d'une expression régulière

### Création et configuration de l'objet RegExp
```vba
Dim regex As Object  
Set regex = CreateObject("VBScript.RegExp")  

' Configuration de base
regex.Global = True         ' Chercher toutes les occurrences (pas seulement la première)  
regex.IgnoreCase = True     ' Ignorer majuscules/minuscules  
regex.Pattern = "motif"     ' Le motif à chercher  
```

### Utilisation basique
```vba
Function TestRegex()
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "VBA"
    regex.IgnoreCase = True

    Dim texte As String
    texte = "J'apprends VBA et vba est formidable"

    ' Test de correspondance
    If regex.Test(texte) Then
        MsgBox "VBA trouvé dans le texte !"
    End If
End Function
```

## Métacaractères de base

Les métacaractères sont des symboles spéciaux qui ont une signification particulière dans les expressions régulières.

### Le point (.) - N'importe quel caractère
```vba
' Pattern : "V.A" trouve "VBA", "V A", "V1A", etc.
regex.Pattern = "V.A"

' Exemples de correspondances :
' "VBA" ✓
' "V A" ✓
' "V1A" ✓
' "VXYA" ✗ (trop de caractères entre V et A)
```

### L'astérisque (*) - Zéro ou plusieurs occurrences
```vba
' Pattern : "Bon*jour" trouve "Bojour", "Bonjour", "Bonnjour", etc.
regex.Pattern = "Bon*jour"

' Exemples :
' "Bojour" ✓ (zéro 'n')
' "Bonjour" ✓ (un 'n')
' "Bonnjour" ✓ (deux 'n')
```

### Le plus (+) - Une ou plusieurs occurrences
```vba
' Pattern : "Bon+jour" trouve "Bonjour", "Bonnjour", mais pas "Bojour"
regex.Pattern = "Bon+jour"

' Exemples :
' "Bojour" ✗ (zéro 'n')
' "Bonjour" ✓ (un 'n')
' "Bonnjour" ✓ (deux 'n')
```

### Le point d'interrogation (?) - Zéro ou une occurrence
```vba
' Pattern : "Bon?jour" trouve "Bojour" et "Bonjour" seulement
regex.Pattern = "Bon?jour"

' Exemples :
' "Bojour" ✓ (zéro 'n')
' "Bonjour" ✓ (un 'n')
' "Bonnjour" ✗ (trop de 'n')
```

## Classes de caractères

### Classes prédéfinies
```vba
' \d = chiffre (0-9)
regex.Pattern = "\d+"  ' Un ou plusieurs chiffres

' \w = caractère alphanumérique (lettres, chiffres, _)
regex.Pattern = "\w+"  ' Un ou plusieurs caractères alphanumériques

' \s = espace (espace, tabulation, retour ligne)
regex.Pattern = "\s+"  ' Un ou plusieurs espaces
```

### Classes personnalisées avec crochets [ ]
```vba
' [abc] = exactement a, b ou c
regex.Pattern = "[abc]+"  ' Une ou plusieurs lettres a, b ou c

' [a-z] = n'importe quelle lettre minuscule
regex.Pattern = "[a-z]+"

' [0-9] = n'importe quel chiffre (équivalent à \d)
regex.Pattern = "[0-9]+"

' [a-zA-Z] = n'importe quelle lettre
regex.Pattern = "[a-zA-Z]+"
```

## Exemples pratiques pour débutants

### Validation d'un code postal français
```vba
Function ValiderCodePostal(code As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    ' Pattern : exactement 5 chiffres
    regex.Pattern = "^[0-9]{5}$"

    ValiderCodePostal = regex.Test(code)
End Function

' Tests
Debug.Print ValiderCodePostal("75001")    ' True  
Debug.Print ValiderCodePostal("1234")     ' False (trop court)  
Debug.Print ValiderCodePostal("12345a")   ' False (contient une lettre)  
```

### Trouver tous les nombres dans un texte
```vba
Function ExtraireNombres(texte As String) As String()
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "\d+"
    regex.Global = True

    Dim matches As Object
    Set matches = regex.Execute(texte)

    ' Créer un tableau avec les résultats
    Dim resultats() As String
    ReDim resultats(matches.Count - 1)

    Dim i As Integer
    For i = 0 To matches.Count - 1
        resultats(i) = matches(i).Value
    Next i

    ExtraireNombres = resultats
End Function

' Utilisation
Dim texte As String  
texte = "J'ai 25 ans et je gagne 1500 euros"  
Dim nombres() As String  
nombres = ExtraireNombres(texte)  
' Résultat : ["25", "1500"]
```

### Validation d'une adresse email simple
```vba
Function ValiderEmailSimple(email As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    ' Pattern simple : caractères@caractères.caractères
    regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    regex.IgnoreCase = True

    ValiderEmailSimple = regex.Test(email)
End Function

' Tests
Debug.Print ValiderEmailSimple("jean.dupont@example.com")  ' True  
Debug.Print ValiderEmailSimple("invalid.email")           ' False  
```

## Ancres de position

### Début et fin de chaîne
```vba
' ^ = début de chaîne
' $ = fin de chaîne

' Vérifier qu'un texte commence par "Bonjour"
regex.Pattern = "^Bonjour"

' Vérifier qu'un texte se termine par "VBA"
regex.Pattern = "VBA$"

' Vérifier qu'un texte est exactement "VBA" (rien d'autre)
regex.Pattern = "^VBA$"
```

### Exemple : validation d'un format exact
```vba
Function ValiderFormatExact(texte As String, pattern As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    ' Ajouter ^ et $ pour correspondance exacte
    regex.Pattern = "^" & pattern & "$"

    ValiderFormatExact = regex.Test(texte)
End Function
```

## Quantificateurs précis

### Accolades { } pour spécifier le nombre exact
```vba
' {n} = exactement n occurrences
regex.Pattern = "[0-9]{5}"  ' Exactement 5 chiffres

' {n,m} = entre n et m occurrences
regex.Pattern = "[a-z]{2,5}"  ' Entre 2 et 5 lettres minuscules

' {n,} = au moins n occurrences
regex.Pattern = "[0-9]{3,}"  ' Au moins 3 chiffres
```

### Exemple : validation d'un numéro de téléphone français
```vba
Function ValiderTelephoneFrancais(numero As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    ' Pattern : 0 suivi de 1 chiffre, puis 8 chiffres (format 01 23 45 67 89)
    ' Accepte les espaces, points ou tirets comme séparateurs
    regex.Pattern = "^0[1-9]([.\s-]?[0-9]{2}){4}$"

    ValiderTelephoneFrancais = regex.Test(numero)
End Function

' Tests
Debug.Print ValiderTelephoneFrancais("01.23.45.67.89")  ' True  
Debug.Print ValiderTelephoneFrancais("01 23 45 67 89")  ' True  
Debug.Print ValiderTelephoneFrancais("0123456789")      ' True  
Debug.Print ValiderTelephoneFrancais("1234567890")      ' False (ne commence pas par 0)  
```

## Groupes et captures

### Parenthèses pour grouper
```vba
' Grouper des éléments ensemble
regex.Pattern = "(VBA|Excel|Word)+"  ' VBA ou Excel ou Word, une ou plusieurs fois

' Exemple d'usage
Function ChercherLogiciels(texte As String) As String()
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "(VBA|Excel|Word|PowerPoint|Access)"
    regex.Global = True
    regex.IgnoreCase = True

    Dim matches As Object
    Set matches = regex.Execute(texte)

    Dim resultats() As String
    ReDim resultats(matches.Count - 1)

    Dim i As Integer
    For i = 0 To matches.Count - 1
        resultats(i) = matches(i).Value
    Next i

    ChercherLogiciels = resultats
End Function
```

## Remplacement avec expressions régulières

### Méthode Replace de RegExp
```vba
Function RemplacerAvecRegex(texte As String, pattern As String, remplacement As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = pattern
    regex.Global = True

    RemplacerAvecRegex = regex.Replace(texte, remplacement)
End Function

' Exemple : masquer les numéros de téléphone
Function MasquerTelephones(texte As String) As String
    MasquerTelephones = RemplacerAvecRegex(texte, "\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{2}", "XX.XX.XX.XX.XX")
End Function

' Test
Dim texte As String  
texte = "Appelez-moi au 01.23.45.67.89 ou au 06.12.34.56.78"  
Debug.Print MasquerTelephones(texte)  
' Résultat : "Appelez-moi au XX.XX.XX.XX.XX ou au XX.XX.XX.XX.XX"
```

## Fonctions utilitaires pour débutants

### Fonction de validation générique
```vba
Function ValiderFormat(valeur As String, pattern As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = pattern
    regex.IgnoreCase = True

    ValiderFormat = regex.Test(valeur)
End Function

' Utilisations
Debug.Print ValiderFormat("AB123", "^[A-Z]{2}[0-9]{3}$")  ' True  
Debug.Print ValiderFormat("hello@test.com", "^.+@.+\..+$")  ' True  
```

### Fonction d'extraction générique
```vba
Function ExtraireMotifs(texte As String, pattern As String) As String()
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = pattern
    regex.Global = True

    Dim matches As Object
    Set matches = regex.Execute(texte)

    If matches.Count = 0 Then
        ExtraireMotifs = Split("")  ' Retourne un tableau avec un élément vide
        Exit Function
    End If

    Dim resultats() As String
    ReDim resultats(matches.Count - 1)

    Dim i As Integer
    For i = 0 To matches.Count - 1
        resultats(i) = matches(i).Value
    Next i

    ExtraireMotifs = resultats
End Function
```

### Fonction de nettoyage avec regex
```vba
Function NettoyerTexteAvecRegex(texte As String) As String
    Dim resultat As String
    resultat = texte

    ' Supprimer tous les caractères non alphanumériques sauf espaces
    resultat = RemplacerAvecRegex(resultat, "[^a-zA-Z0-9\s]", "")

    ' Supprimer les espaces multiples
    resultat = RemplacerAvecRegex(resultat, "\s+", " ")

    ' Supprimer les espaces en début et fin
    resultat = Trim(resultat)

    NettoyerTexteAvecRegex = resultat
End Function
```

## Patterns courants pour débutants

### Collection de patterns utiles
```vba
' Numéro de téléphone français
Const PATTERN_TEL_FR As String = "^0[1-9]([.\s-]?[0-9]{2}){4}$"

' Code postal français
Const PATTERN_CP_FR As String = "^[0-9]{5}$"

' Email simple
Const PATTERN_EMAIL As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"

' Date format JJ/MM/AAAA
Const PATTERN_DATE_FR As String = "^[0-9]{2}/[0-9]{2}/[0-9]{4}$"

' Numéro de sécurité sociale français (approximatif)
Const PATTERN_NUM_SS As String = "^[1-2][0-9]{2}[0-1][0-9][0-9][0-9]{3}[0-9]{3}[0-9]{2}$"
```

### Fonction de validation multiple
```vba
Function ValiderDonnee(valeur As String, typeDonnee As String) As Boolean
    Dim pattern As String

    Select Case LCase(typeDonnee)
        Case "telephone"
            pattern = "^0[1-9]([.\s-]?[0-9]{2}){4}$"
        Case "email"
            pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        Case "codepostal"
            pattern = "^[0-9]{5}$"
        Case "date"
            pattern = "^[0-9]{2}/[0-9]{2}/[0-9]{4}$"
        Case Else
            ValiderDonnee = False
            Exit Function
    End Select

    ValiderDonnee = ValiderFormat(valeur, pattern)
End Function

' Utilisation
Debug.Print ValiderDonnee("01.23.45.67.89", "telephone")  ' True  
Debug.Print ValiderDonnee("test@example.com", "email")     ' True  
```

## Conseils pour débuter avec les regex

### 1. Commencer simple
```vba
' Commencez par des patterns très simples
regex.Pattern = "VBA"        ' Cherche exactement "VBA"  
regex.Pattern = "[0-9]"      ' Cherche n'importe quel chiffre  
regex.Pattern = "[0-9]+"     ' Cherche un ou plusieurs chiffres  
```

### 2. Tester pas à pas
```vba
' Testez chaque ajout au pattern
Sub TesterPattern()
    Dim patterns() As String
    patterns = Split("VBA|[V][B][A]|[VBA]+|^VBA$", "|")

    Dim texte As String
    texte = "J'apprends VBA"
    Dim i As Integer

    For i = 0 To UBound(patterns)
        Debug.Print "Pattern: " & patterns(i) & " -> " & ValiderFormat(texte, patterns(i))
    Next i
End Sub
```

### 3. Documenter les patterns complexes
```vba
Function ValiderNumeroCompte(numero As String) As Boolean
    ' Pattern pour numéro de compte bancaire français
    ' Format : 5 chiffres (banque) + 5 chiffres (guichet) + 11 caractères (compte) + 2 chiffres (clé)
    Dim pattern As String
    pattern = "^[0-9]{5}\s?[0-9]{5}\s?[0-9A-Z]{11}\s?[0-9]{2}$"

    ValiderNumeroCompte = ValiderFormat(numero, pattern)
End Function
```

### 4. Prévoir les cas limites
```vba
' Toujours tester avec des données variées
Sub TestsValidation()
    ' Tester les cas normaux
    Debug.Print ValiderEmailSimple("test@example.com")     ' True

    ' Tester les cas limites
    Debug.Print ValiderEmailSimple("")                     ' False
    Debug.Print ValiderEmailSimple("@")                    ' False
    Debug.Print ValiderEmailSimple("test@")                ' False
    Debug.Print ValiderEmailSimple("@example.com")         ' False
End Sub
```

Les expressions régulières sont un outil puissant mais qui demande de la pratique. Commencez par des patterns simples et augmentez progressivement la complexité. Elles vous feront gagner beaucoup de temps pour valider et extraire des données dans vos projets VBA.

⏭️ [10. Dates et heures](/10-dates-et-heures/)
