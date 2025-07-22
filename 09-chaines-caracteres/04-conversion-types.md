🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.4. Conversion de types

## Introduction à la conversion de types

La conversion de types consiste à transformer une donnée d'un type vers un autre. Imaginez que vous ayez le nombre "123" écrit comme du texte dans une cellule Excel - pour faire des calculs dessus, vous devez le convertir en nombre réel. Ou inversement, vous voulez afficher le résultat d'un calcul (un nombre) dans un message à l'utilisateur (du texte).

Ces conversions sont essentielles car VBA est strict sur les types : on ne peut pas additionner directement du texte et des nombres, par exemple.

## Types de données courants en VBA

Avant de voir les conversions, rappelons les types principaux :

| Type | Description | Exemple |
|------|-------------|---------|
| String | Chaîne de caractères | "Bonjour" |
| Integer | Nombre entier (-32 768 à 32 767) | 123 |
| Long | Nombre entier long | 1234567 |
| Double | Nombre décimal | 123.45 |
| Boolean | Vrai ou Faux | True, False |
| Date | Date et heure | #2024-01-15# |
| Variant | Type variable (peut contenir n'importe quoi) | Tout |

## Conversions depuis String (texte)

### Convertir du texte en nombres

#### CInt() - Conversion en Integer
```vba
Dim texteNombre As String = "123"
Dim nombre As Integer

nombre = CInt(texteNombre)
' Résultat : 123 (de type Integer)

' Utilisation directe
Dim age As Integer = CInt("25")
```

#### CLng() - Conversion en Long
```vba
Dim texteGrandNombre As String = "1234567"
Dim grandNombre As Long = CLng(texteGrandNombre)
' Résultat : 1234567 (de type Long)
```

#### CDbl() - Conversion en Double (décimal)
```vba
Dim texteDecimal As String = "123.45"
Dim nombreDecimal As Double = CDbl(texteDecimal)
' Résultat : 123.45 (de type Double)

' Attention aux séparateurs décimaux selon les paramètres régionaux
Dim prix As String = "19,99"  ' Format français
Dim prixNumerique As Double = CDbl(Replace(prix, ",", "."))
```

#### Val() - Fonction universelle de conversion
```vba
' Val() est plus tolérante que les fonctions C...()
Dim resultat1 As Double = Val("123")      ' 123
Dim resultat2 As Double = Val("123.45")   ' 123.45
Dim resultat3 As Double = Val("123abc")   ' 123 (s'arrête au premier caractère non numérique)
Dim resultat4 As Double = Val("abc123")   ' 0 (ne commence pas par un chiffre)
```

### Convertir du texte en Boolean
```vba
Dim texteBoolean As String = "True"
Dim valeurBoolean As Boolean = CBool(texteBoolean)
' Résultat : True

' Autres exemples
Dim bool1 As Boolean = CBool("False")  ' False
Dim bool2 As Boolean = CBool("1")      ' True
Dim bool3 As Boolean = CBool("0")      ' False
```

### Convertir du texte en Date
```vba
' CDate() pour les conversions de date
Dim texteDate As String = "15/01/2024"
Dim maDate As Date = CDate(texteDate)

' DateValue() pour juste la date (sans heure)
Dim dateSeule As Date = DateValue("15/01/2024")

' TimeValue() pour juste l'heure
Dim heureSeule As Date = TimeValue("14:30:00")
```

## Conversions vers String (texte)

### CStr() - Conversion universelle en texte
```vba
Dim nombre As Integer = 123
Dim texte As String = CStr(nombre)
' Résultat : "123"

Dim decimal As Double = 123.45
Dim texteDecimal As String = CStr(decimal)
' Résultat : "123,45" (selon paramètres régionaux)

Dim dateActuelle As Date = Now
Dim texteDate As String = CStr(dateActuelle)
' Résultat : "22/07/2025 10:30:15" (exemple)
```

### Str() - Conversion de nombre en chaîne
```vba
Dim nombre As Integer = 123
Dim texte As String = Str(nombre)
' Résultat : " 123" (note l'espace au début pour les nombres positifs)

' Ltrim pour supprimer l'espace
Dim texteClean As String = LTrim(Str(nombre))
' Résultat : "123"
```

### Format() - Conversion avec formatage
```vba
Dim nombre As Double = 1234.567

' Formatage de nombres
Dim texte1 As String = Format(nombre, "0.00")        ' "1234,57"
Dim texte2 As String = Format(nombre, "#,##0.00")    ' "1 234,57"
Dim texte3 As String = Format(nombre, "0%")          ' "123457%" (multiplie par 100)

' Formatage de dates
Dim dateActuelle As Date = Now
Dim texte4 As String = Format(dateActuelle, "dd/mm/yyyy")     ' "22/07/2025"
Dim texte5 As String = Format(dateActuelle, "dddd dd mmmm")   ' "mardi 22 juillet"
```

## Conversions entre types numériques

### Conversion sécurisée entre Integer, Long et Double
```vba
' De Double vers Integer (attention à la perte de précision)
Dim decimal As Double = 123.67
Dim entier As Integer = CInt(decimal)  ' 124 (arrondi)

' De Integer vers Long (sans perte)
Dim petit As Integer = 123
Dim grand As Long = CLng(petit)  ' 123

' De Long vers Double (sans perte généralement)
Dim entierLong As Long = 123456
Dim decimalLong As Double = CDbl(entierLong)  ' 123456.0
```

### Fonctions d'arrondi avant conversion
```vba
Dim nombre As Double = 123.67

' Round() - Arrondi mathématique
Dim arrondi As Integer = CInt(Round(nombre))  ' 124

' Int() - Partie entière (tronque)
Dim tronque As Integer = CInt(Int(nombre))    ' 123

' Fix() - Supprime la partie décimale
Dim fixe As Integer = CInt(Fix(nombre))       ' 123
```

## Gestion des erreurs de conversion

### Vérification avant conversion
```vba
Function ConvertirEnNombreSur(texte As String) As Double
    ' Vérifier si la conversion est possible
    If IsNumeric(texte) Then
        ConvertirEnNombreSur = CDbl(texte)
    Else
        ConvertirEnNombreSur = 0  ' Valeur par défaut
    End If
End Function

' Utilisation
Dim texte1 As String = "123.45"  ' Valide
Dim texte2 As String = "abc"     ' Invalide

Dim nombre1 As Double = ConvertirEnNombreSur(texte1)  ' 123.45
Dim nombre2 As Double = ConvertirEnNombreSur(texte2)  ' 0
```

### Utilisation de IsNumeric(), IsDate()
```vba
Dim valeur As String = "123.45"

' Vérifier si c'est un nombre
If IsNumeric(valeur) Then
    Dim nombre As Double = CDbl(valeur)
    MsgBox "Conversion réussie : " & nombre
Else
    MsgBox "Ce n'est pas un nombre valide"
End If

' Vérifier si c'est une date
Dim texteDate As String = "15/01/2024"
If IsDate(texteDate) Then
    Dim maDate As Date = CDate(texteDate)
    MsgBox "Date valide : " & Format(maDate, "dd/mm/yyyy")
End If
```

### Gestion d'erreurs avec On Error
```vba
Function ConversionSecurisee(texte As String) As Variant
    On Error GoTo GestionErreur

    ' Tentative de conversion
    ConversionSecurisee = CDbl(texte)
    Exit Function

GestionErreur:
    ' En cas d'erreur, retourner une valeur spéciale
    ConversionSecurisee = "ERREUR"
End Function
```

## Conversions spéciales et utilitaires

### Conversion de Boolean en texte personnalisé
```vba
Function BooleanVersTexte(valeur As Boolean) As String
    If valeur Then
        BooleanVersTexte = "Oui"
    Else
        BooleanVersTexte = "Non"
    End If
End Function

' Utilisation
Dim estValide As Boolean = True
Dim texte As String = BooleanVersTexte(estValide)  ' "Oui"
```

### Conversion de codes ASCII
```vba
' Caractère vers code ASCII
Dim caractere As String = "A"
Dim code As Integer = Asc(caractere)  ' 65

' Code ASCII vers caractère
Dim nouveauCaractere As String = Chr(65)  ' "A"

' Utile pour générer des caractères spéciaux
Dim guillemet As String = Chr(34)  ' " (guillemet)
Dim retourLigne As String = Chr(13) & Chr(10)  ' Équivalent à vbCrLf
```

## Formatage avancé avec Format()

### Formats numériques courants
```vba
Dim nombre As Double = 1234.567

' Formats prédéfinis
Dim texte1 As String = Format(nombre, "General Number")   ' "1234,567"
Dim texte2 As String = Format(nombre, "Currency")         ' "1 234,57 €"
Dim texte3 As String = Format(nombre, "Standard")         ' "1 234,57"
Dim texte4 As String = Format(nombre, "Fixed")            ' "1234,57"
Dim texte5 As String = Format(nombre, "Percent")          ' "123456,70%"

' Formats personnalisés
Dim texte6 As String = Format(nombre, "000000.00")        ' "001234,57"
Dim texte7 As String = Format(nombre, "#,##0.00 €")       ' "1 234,57 €"
```

### Formats de date courants
```vba
Dim dateActuelle As Date = #7/22/2025 2:30:15 PM#

' Formats prédéfinis
Dim texte1 As String = Format(dateActuelle, "General Date")     ' "22/07/2025 14:30:15"
Dim texte2 As String = Format(dateActuelle, "Long Date")        ' "mardi 22 juillet 2025"
Dim texte3 As String = Format(dateActuelle, "Medium Date")      ' "22-juil-25"
Dim texte4 As String = Format(dateActuelle, "Short Date")       ' "22/07/2025"
Dim texte5 As String = Format(dateActuelle, "Long Time")        ' "14:30:15"
Dim texte6 As String = Format(dateActuelle, "Medium Time")      ' "02:30 PM"
Dim texte7 As String = Format(dateActuelle, "Short Time")       ' "14:30"

' Formats personnalisés
Dim texte8 As String = Format(dateActuelle, "dd/mm/yyyy")       ' "22/07/2025"
Dim texte9 As String = Format(dateActuelle, "dddd dd mmmm yyyy")' "mardi 22 juillet 2025"
Dim texte10 As String = Format(dateActuelle, "hh:nn:ss")        ' "14:30:15"
```

## Cas pratiques courants

### Nettoyer et convertir des données de cellules Excel
```vba
Function ConvertirCelluleEnNombre(cellule As Range) As Double
    Dim valeurTexte As String = CStr(cellule.Value)

    ' Nettoyer la valeur
    valeurTexte = Trim(valeurTexte)                    ' Supprimer espaces
    valeurTexte = Replace(valeurTexte, " ", "")        ' Supprimer espaces internes
    valeurTexte = Replace(valeurTexte, "€", "")        ' Supprimer symbole euro
    valeurTexte = Replace(valeurTexte, ",", ".")       ' Normaliser séparateur décimal

    ' Convertir si possible
    If IsNumeric(valeurTexte) Then
        ConvertirCelluleEnNombre = CDbl(valeurTexte)
    Else
        ConvertirCelluleEnNombre = 0
    End If
End Function
```

### Créer un identifiant unique à partir de données
```vba
Function CreerIdentifiant(nom As String, prenom As String, dateNaissance As Date) As String
    ' Créer un ID au format : DUPONT_JEAN_19900115
    Dim nomClean As String = UCase(Replace(Trim(nom), " ", ""))
    Dim prenomClean As String = UCase(Replace(Trim(prenom), " ", ""))
    Dim dateClean As String = Format(dateNaissance, "yyyymmdd")

    CreerIdentifiant = nomClean & "_" & prenomClean & "_" & dateClean
End Function
```

### Validation et conversion de saisies utilisateur
```vba
Function ValiderEtConvertirAge(saisie As String) As Integer
    ' Nettoyer la saisie
    Dim saisieClean As String = Trim(saisie)

    ' Vérifier si c'est numérique
    If Not IsNumeric(saisieClean) Then
        MsgBox "L'âge doit être un nombre"
        ValiderEtConvertirAge = -1  ' Code d'erreur
        Exit Function
    End If

    ' Convertir et valider la plage
    Dim age As Integer = CInt(saisieClean)
    If age < 0 Or age > 150 Then
        MsgBox "L'âge doit être entre 0 et 150 ans"
        ValiderEtConvertirAge = -1
    Else
        ValiderEtConvertirAge = age
    End If
End Function
```

## Bonnes pratiques pour les débutants

### 1. Toujours valider avant de convertir
```vba
' MAUVAIS
Dim nombre As Integer = CInt(texte)  ' Peut provoquer une erreur

' BON
If IsNumeric(texte) Then
    Dim nombre As Integer = CInt(texte)
Else
    ' Gérer le cas d'erreur
End If
```

### 2. Prévoir des valeurs par défaut
```vba
Function ConversionAvecDefaut(texte As String, valeurDefaut As Double) As Double
    If IsNumeric(texte) Then
        ConversionAvecDefaut = CDbl(texte)
    Else
        ConversionAvecDefaut = valeurDefaut
    End If
End Function
```

### 3. Documenter les formats attendus
```vba
' Convertit une date au format "jj/mm/aaaa" en objet Date
Function ConvertirDateFrancaise(texteDate As String) As Date
    ' Format attendu : "22/07/2025"
    If IsDate(texteDate) Then
        ConvertirDateFrancaise = CDate(texteDate)
    Else
        ConvertirDateFrancaise = #1/1/1900#  ' Date par défaut
    End If
End Function
```

### 4. Tester avec des données variées
Testez toujours vos conversions avec :
- Des valeurs normales
- Des valeurs limites (très grandes, très petites)
- Des valeurs vides
- Des valeurs invalides
- Des formats différents

### 5. Utiliser Variant pour plus de flexibilité
```vba
Function ConversionFlexible(valeur As Variant) As String
    Select Case VarType(valeur)
        Case vbString
            ConversionFlexible = CStr(valeur)
        Case vbInteger, vbLong, vbDouble
            ConversionFlexible = Format(valeur, "0.00")
        Case vbDate
            ConversionFlexible = Format(valeur, "dd/mm/yyyy")
        Case vbBoolean
            ConversionFlexible = IIf(valeur, "Oui", "Non")
        Case Else
            ConversionFlexible = "Non convertible"
    End Select
End Function
```

La maîtrise des conversions de types vous permettra de manipuler efficacement les données dans vos programmes VBA, en évitant les erreurs courantes et en créant des solutions robustes.

⏭️
