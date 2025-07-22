🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.3. Recherche et remplacement

## Introduction à la recherche et au remplacement

La recherche et le remplacement de texte sont des opérations fondamentales dans le traitement de données. Imaginez que vous ayez un document où vous devez remplacer tous les "M." par "Monsieur" ou trouver toutes les occurrences d'un code client spécifique. VBA vous offre des outils puissants pour effectuer ces tâches automatiquement.

Ces opérations sont comme une fonction "Rechercher et remplacer" dans Word, mais programmables et beaucoup plus flexibles.

## Fonction InStr() - Trouver la position d'un texte

### Principe de base
La fonction `InStr()` recherche une chaîne de caractères à l'intérieur d'une autre et retourne la position de la première occurrence trouvée.

### Syntaxe simple
```vba
InStr(chaîne_dans_laquelle_chercher, chaîne_à_chercher)
```

### Exemples de base
```vba
Dim texte As String = "Bonjour le monde VBA"
Dim position As Integer

position = InStr(texte, "le")
' Résultat : 9 (première occurrence de "le")

position = InStr(texte, "VBA")
' Résultat : 18

position = InStr(texte, "Python")
' Résultat : 0 (non trouvé)
```

### Comprendre les résultats d'InStr()
```vba
' Position dans la chaîne : 123456789012345678901
Dim texte As String = "Bonjour le monde VBA"
'                      B=1, o=2, n=3, j=4, o=5, u=6, r=7, (espace)=8, l=9, e=10...

' Quand InStr trouve quelque chose
If InStr(texte, "monde") > 0 Then
    MsgBox "Le mot 'monde' a été trouvé !"
End If

' Quand InStr ne trouve rien
If InStr(texte, "Python") = 0 Then
    MsgBox "Le mot 'Python' n'a pas été trouvé"
End If
```

### Syntaxe complète d'InStr()
```vba
InStr(position_de_début, chaîne_source, chaîne_recherchée, mode_comparaison)
```

### Paramètres de comparaison
```vba
Dim texte As String = "Bonjour Le Monde"

' Recherche sensible à la casse (par défaut)
Dim pos1 As Integer = InStr(texte, "le")
' Résultat : 0 (non trouvé car "le" ≠ "Le")

' Recherche insensible à la casse
Dim pos2 As Integer = InStr(1, texte, "le", vbTextCompare)
' Résultat : 9 (trouve "Le")
```

### Recherche à partir d'une position spécifique
```vba
Dim texte As String = "le chat mange le poisson"
Dim position As Integer

' Première occurrence de "le"
position = InStr(texte, "le")
' Résultat : 1

' Chercher "le" à partir de la position 5
position = InStr(5, texte, "le")
' Résultat : 16 (deuxième "le")
```

## Fonction InStrRev() - Recherche inversée

### Principe de base
`InStrRev()` fonctionne comme `InStr()` mais cherche depuis la fin de la chaîne vers le début.

### Syntaxe
```vba
InStrRev(chaîne_source, chaîne_recherchée)
```

### Exemples pratiques
```vba
Dim cheminFichier As String = "C:\Documents\Projets\MonFichier.xlsx"

' Trouver la dernière occurrence de "\"
Dim dernierePosition As Integer = InStrRev(cheminFichier, "\")
' Résultat : 20 (position du dernier \)

' Extraire juste le nom du fichier
Dim nomFichier As String = Mid(cheminFichier, dernierePosition + 1)
' Résultat : "MonFichier.xlsx"
```

### Cas d'usage typique : séparer nom et extension
```vba
Function SeparerNomExtension(nomComplet As String) As String()
    Dim resultat(1) As String  ' 0 = nom, 1 = extension
    Dim positionPoint As Integer

    positionPoint = InStrRev(nomComplet, ".")

    If positionPoint > 0 Then
        resultat(0) = Left(nomComplet, positionPoint - 1)  ' Nom
        resultat(1) = Mid(nomComplet, positionPoint + 1)   ' Extension
    Else
        resultat(0) = nomComplet  ' Pas d'extension
        resultat(1) = ""
    End If

    SeparerNomExtension = resultat
End Function

' Utilisation
Dim fichier As String = "document.xlsx"
Dim parties() As String = SeparerNomExtension(fichier)
' parties(0) = "document"
' parties(1) = "xlsx"
```

## Fonction Replace() - Remplacement simple

### Principe de base (rappel et approfondissement)
La fonction `Replace()` remplace toutes les occurrences d'une chaîne par une autre.

### Syntaxe complète
```vba
Replace(expression, find, replace, start, count, compare)
```

### Paramètres détaillés
- **expression** : la chaîne source
- **find** : ce qu'il faut chercher
- **replace** : par quoi remplacer
- **start** : position de début (optionnel, défaut = 1)
- **count** : nombre max de remplacements (optionnel, défaut = tous)
- **compare** : mode de comparaison (optionnel)

### Exemples de remplacement de base
```vba
Dim texte As String = "J'aime le Java et le Java est bien"
Dim nouveau As String

' Remplacement simple
nouveau = Replace(texte, "Java", "VBA")
' Résultat : "J'aime le VBA et le VBA est bien"

' Remplacement avec limite
nouveau = Replace(texte, "le", "un", 1, 1)  ' Remplacer seulement le premier "le"
' Résultat : "J'aime un Java et le Java est bien"
```

### Remplacement insensible à la casse
```vba
Dim texte As String = "BONJOUR bonjour Bonjour"
Dim nouveau As String

' Sensible à la casse (défaut)
nouveau = Replace(texte, "bonjour", "salut")
' Résultat : "BONJOUR salut Bonjour"

' Insensible à la casse
nouveau = Replace(texte, "bonjour", "salut", 1, -1, vbTextCompare)
' Résultat : "salut salut salut"
```

## Techniques de recherche avancées

### Recherche de mots entiers
```vba
Function ChercherMotEntier(texte As String, mot As String) As Boolean
    Dim position As Integer = InStr(1, texte, mot, vbTextCompare)

    Do While position > 0
        Dim avantOK As Boolean = True
        Dim apresOK As Boolean = True

        ' Vérifier le caractère précédent
        If position > 1 Then
            Dim charAvant As String = Mid(texte, position - 1, 1)
            avantOK = (charAvant = " " Or charAvant = vbTab Or charAvant = vbCrLf)
        End If

        ' Vérifier le caractère suivant
        If position + Len(mot) <= Len(texte) Then
            Dim charApres As String = Mid(texte, position + Len(mot), 1)
            apresOK = (charApres = " " Or charApres = vbTab Or charApres = vbCrLf)
        End If

        If avantOK And apresOK Then
            ChercherMotEntier = True
            Exit Function
        End If

        ' Chercher l'occurrence suivante
        position = InStr(position + 1, texte, mot, vbTextCompare)
    Loop

    ChercherMotEntier = False
End Function
```

### Compter les occurrences
```vba
Function CompterOccurrences(texte As String, recherche As String) As Integer
    Dim compteur As Integer = 0
    Dim position As Integer = 1

    Do
        position = InStr(position, texte, recherche, vbTextCompare)
        If position > 0 Then
            compteur = compteur + 1
            position = position + Len(recherche)  ' Éviter les chevauchements
        End If
    Loop While position > 0

    CompterOccurrences = compteur
End Function

' Utilisation
Dim texte As String = "le chat mange le poisson que le chien regarde"
Dim nombre As Integer = CompterOccurrences(texte, "le")
' Résultat : 3
```

## Remplacement avancé et conditionnel

### Remplacement avec conditions
```vba
Function RemplacerSelonCondition(texte As String) As String
    Dim resultat As String = texte

    ' Remplacer "M." par "Monsieur" seulement s'il est suivi d'un espace et d'une majuscule
    Dim i As Integer = 1
    Do While i <= Len(resultat)
        Dim pos As Integer = InStr(i, resultat, "M.")
        If pos = 0 Then Exit Do

        ' Vérifier si c'est suivi d'un espace et d'une majuscule
        If pos + 2 <= Len(resultat) Then
            Dim charSuivant As String = Mid(resultat, pos + 2, 1)
            Dim charApres As String = ""
            If pos + 3 <= Len(resultat) Then
                charApres = Mid(resultat, pos + 3, 1)
            End If

            If charSuivant = " " And charApres >= "A" And charApres <= "Z" Then
                resultat = Left(resultat, pos - 1) & "Monsieur" & Mid(resultat, pos + 2)
                i = pos + Len("Monsieur")
            Else
                i = pos + 1
            End If
        Else
            i = pos + 1
        End If
    Loop

    RemplacerSelonCondition = resultat
End Function
```

### Remplacement multiple en une passe
```vba
Function RemplacementMultiple(texte As String) As String
    Dim resultat As String = texte

    ' Tableau des remplacements : ancien -> nouveau
    Dim remplacements As Variant
    remplacements = Array("M.", "Monsieur", "Mme", "Madame", "Dr", "Docteur")

    Dim i As Integer
    For i = 0 To UBound(remplacements) Step 2
        resultat = Replace(resultat, remplacements(i), remplacements(i + 1))
    Next i

    RemplacementMultiple = resultat
End Function
```

## Nettoyage et standardisation de données

### Supprimer les caractères indésirables
```vba
Function NettoyerTexte(texte As String) As String
    Dim resultat As String = texte

    ' Supprimer les caractères de contrôle
    resultat = Replace(resultat, vbCrLf, " ")
    resultat = Replace(resultat, vbCr, " ")
    resultat = Replace(resultat, vbLf, " ")
    resultat = Replace(resultat, vbTab, " ")

    ' Supprimer les espaces multiples
    Do While InStr(resultat, "  ") > 0
        resultat = Replace(resultat, "  ", " ")
    Loop

    ' Supprimer les espaces en début et fin
    NettoyerTexte = Trim(resultat)
End Function
```

### Standardiser les numéros de téléphone
```vba
Function StandardiserTelephone(numero As String) As String
    Dim resultat As String = numero

    ' Supprimer tous les caractères non numériques sauf le +
    Dim i As Integer
    Dim numeroClean As String = ""

    For i = 1 To Len(resultat)
        Dim char As String = Mid(resultat, i, 1)
        If (char >= "0" And char <= "9") Or char = "+" Then
            numeroClean = numeroClean & char
        End If
    Next i

    ' Formatter selon le standard français
    If Left(numeroClean, 1) = "+" Then
        StandardiserTelephone = numeroClean  ' Garder le format international
    ElseIf Len(numeroClean) = 10 Then
        ' Format français : 01.23.45.67.89
        StandardiserTelephone = Left(numeroClean, 2) & "." & _
                               Mid(numeroClean, 3, 2) & "." & _
                               Mid(numeroClean, 5, 2) & "." & _
                               Mid(numeroClean, 7, 2) & "." & _
                               Right(numeroClean, 2)
    Else
        StandardiserTelephone = numeroClean  ' Retourner tel quel si format non reconnu
    End If
End Function
```

## Recherche avec critères multiples

### Fonction pour vérifier plusieurs mots-clés
```vba
Function ContientMotsCles(texte As String, motsCles As String) As Boolean
    ' motsCles séparés par des virgules : "VBA,Excel,macro"
    Dim tableauMots() As String = Split(motsCles, ",")

    Dim i As Integer
    For i = 0 To UBound(tableauMots)
        Dim motCle As String = Trim(tableauMots(i))
        If InStr(1, texte, motCle, vbTextCompare) > 0 Then
            ContientMotsCles = True
            Exit Function
        End If
    Next i

    ContientMotsCles = False
End Function

' Utilisation
Dim description As String = "Formation VBA pour Excel débutants"
If ContientMotsCles(description, "VBA,Excel,Access") Then
    MsgBox "Cette formation concerne nos outils !"
End If
```

## Conseils pratiques pour les débutants

### 1. Toujours vérifier le résultat d'InStr()
```vba
' MAUVAIS
Dim pos As Integer = InStr(texte, "cherche")
Dim resultat As String = Mid(texte, pos)  ' Erreur si pos = 0 !

' BON
Dim pos As Integer = InStr(texte, "cherche")
If pos > 0 Then
    Dim resultat As String = Mid(texte, pos)
End If
```

### 2. Attention à la casse par défaut
```vba
' Par défaut, VBA est sensible à la casse
If InStr(texte, "bonjour") > 0 Then  ' Ne trouvera pas "Bonjour"

' Utiliser vbTextCompare pour ignorer la casse
If InStr(1, texte, "bonjour", vbTextCompare) > 0 Then  ' Trouvera "Bonjour"
```

### 3. Gérer les remplacements en cascade
```vba
' ATTENTION : les remplacements peuvent interagir
Dim texte As String = "123"
texte = Replace(texte, "1", "12")  ' "223"
texte = Replace(texte, "2", "21")  ' "21213" (pas ce qu'on voulait !)

' MIEUX : faire tous les remplacements sur le texte original
Dim original As String = "123"
Dim etape1 As String = Replace(original, "1", "12")
Dim final As String = Replace(etape1, "2", "21")
```

### 4. Documenter les expressions complexes
```vba
' Rechercher un format d'email simple
Function EstEmail(texte As String) As Boolean
    ' Vérifie la présence de @ et d'un point après
    Dim posArobase As Integer = InStr(texte, "@")
    If posArobase = 0 Then
        EstEmail = False
        Exit Function
    End If

    Dim posPoint As Integer = InStr(posArobase, texte, ".")
    EstEmail = (posPoint > posArobase + 1)  ' Au moins un caractère entre @ et .
End Function
```

La maîtrise de la recherche et du remplacement vous permettra de traiter efficacement de grandes quantités de données textuelles et d'automatiser de nombreuses tâches de nettoyage et de standardisation.

⏭️
