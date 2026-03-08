🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.2 Création de fonctions personnalisées (UDF)

## Qu'est-ce qu'une fonction personnalisée (UDF) ?

UDF signifie **User Defined Function** (Fonction Définie par l'Utilisateur). C'est votre propre fonction que vous créez en VBA et qui devient utilisable dans Excel exactement comme les fonctions intégrées (SOMME, MOYENNE, etc.).

### Pourquoi créer ses propres fonctions ?

Imaginez ces situations :
- 🧮 Vous devez calculer la TVA avec des règles spécifiques à votre entreprise
- 📊 Vous voulez une fonction qui compte les cellules selon plusieurs critères complexes
- 💰 Vous avez besoin d'un calcul financier particulier qui n'existe pas dans Excel
- 🔤 Vous voulez transformer du texte d'une façon très précise

Au lieu de répéter le même code VBA partout ou d'utiliser des formules compliquées, vous créez **une seule fonction** que vous pouvez réutiliser partout !

### Avantages des UDF

✅ **Simplicité** : Une formule simple au lieu de calculs complexes  
✅ **Réutilisabilité** : Écrite une fois, utilisée partout  
✅ **Lisibilité** : Le nom de votre fonction explique ce qu'elle fait  
✅ **Maintenance** : Modifier la fonction met à jour tous les calculs  
✅ **Partage** : Vos collègues peuvent utiliser vos fonctions

## Différence entre Sub et Function

### Procédure Sub (action)
```vba
Sub AfficherMessage()
    MsgBox "Bonjour !"
End Sub
```
- **Fait quelque chose** (affiche, modifie, sauvegarde...)
- **Ne retourne pas de valeur**
- S'exécute avec F5 ou un bouton

### Fonction Function (calcul)
```vba
Function CalculerTVA(prixHT As Double) As Double
    CalculerTVA = prixHT * 0.2
End Function
```
- **Calcule et retourne une valeur**
- **Peut être utilisée dans une cellule Excel**
- S'utilise comme =CalculerTVA(100) dans Excel

## Votre première fonction personnalisée

### Exemple 1 : Calculer la TVA
```vba
Function CalculerTVA(prixHT As Double) As Double
    ' Cette fonction calcule la TVA à 20%
    CalculerTVA = prixHT * 0.2
End Function
```

**Comment l'utiliser dans Excel :**
1. Tapez dans une cellule : `=CalculerTVA(100)`
2. Appuyez sur Entrée
3. Résultat : 20

**Explication du code :**
- `Function` : Mot-clé pour créer une fonction
- `CalculerTVA` : Nom de votre fonction
- `prixHT As Double` : Paramètre d'entrée (nombre)
- `As Double` : Type de valeur retournée (nombre)
- `CalculerTVA = ...` : Assigner le résultat au nom de la fonction

### Exemple 2 : Fonction avec plusieurs paramètres
```vba
Function CalculerTVAVariable(prixHT As Double, tauxTVA As Double) As Double
    ' Fonction qui calcule la TVA avec un taux variable
    CalculerTVAVariable = prixHT * (tauxTVA / 100)
End Function
```

**Utilisation :** `=CalculerTVAVariable(100; 20)` → Résultat : 20  
**Utilisation :** `=CalculerTVAVariable(100; 5.5)` → Résultat : 5.5  

### Exemple 3 : Calculer le prix TTC complet
```vba
Function PrixTTC(prixHT As Double, Optional tauxTVA As Double = 20) As Double
    ' Calcule le prix TTC (HT + TVA)
    ' Le taux de TVA est optionnel (20% par défaut)
    PrixTTC = prixHT + (prixHT * tauxTVA / 100)
End Function
```

**Utilisations possibles :**
- `=PrixTTC(100)` → 120 (avec TVA par défaut de 20%)
- `=PrixTTC(100; 10)` → 110 (avec TVA de 10%)

**Note :** `Optional` rend le paramètre facultatif avec une valeur par défaut.

## Fonctions travaillant avec du texte

### Exemple 4 : Mettre en forme un nom complet
```vba
Function FormatNomComplet(prenom As String, nom As String) As String
    ' Met en forme : PRENOM Nom
    FormatNomComplet = UCase(prenom) & " " & StrConv(nom, vbProperCase)
End Function
```

**Utilisation :** `=FormatNomComplet("jean"; "dupont")` → "JEAN Dupont"

### Exemple 5 : Compter les mots dans un texte
```vba
Function CompterMots(texte As String) As Integer
    ' Compte le nombre de mots dans un texte
    Dim motsArray As Variant

    ' Supprimer les espaces en début/fin et diviser par les espaces
    texte = Trim(texte)

    If texte = "" Then
        CompterMots = 0
    Else
        motsArray = Split(texte, " ")
        CompterMots = UBound(motsArray) + 1
    End If
End Function
```

**Utilisation :** `=CompterMots("Bonjour le monde")` → 3

## Fonctions avec logique conditionnelle

### Exemple 6 : Classification par tranche d'âge
```vba
Function CategorieAge(age As Integer) As String
    ' Détermine la catégorie d'âge
    If age < 18 Then
        CategorieAge = "Mineur"
    ElseIf age <= 65 Then
        CategorieAge = "Adulte"
    Else
        CategorieAge = "Senior"
    End If
End Function
```

**Utilisation :** `=CategorieAge(25)` → "Adulte"

### Exemple 7 : Calculer une remise selon le montant
```vba
Function CalculerRemise(montant As Double) As Double
    ' Remise progressive selon le montant d'achat
    Select Case montant
        Case Is < 100
            CalculerRemise = 0          ' Pas de remise
        Case Is < 500
            CalculerRemise = montant * 0.05    ' 5%
        Case Is < 1000
            CalculerRemise = montant * 0.1     ' 10%
        Case Else
            CalculerRemise = montant * 0.15    ' 15%
    End Select
End Function
```

**Utilisations :**
- `=CalculerRemise(50)` → 0
- `=CalculerRemise(300)` → 15
- `=CalculerRemise(1200)` → 180

## Fonctions travaillant avec les plages de cellules

### Exemple 8 : Compter les cellules selon plusieurs critères
```vba
Function CompterSiMultiple(plage As Range, critere1 As String, critere2 As String) As Integer
    ' Compte les cellules qui contiennent critere1 OU critere2
    Dim cellule As Range
    Dim compteur As Integer

    compteur = 0
    For Each cellule In plage
        If InStr(cellule.Value, critere1) > 0 Or InStr(cellule.Value, critere2) > 0 Then
            compteur = compteur + 1
        End If
    Next cellule

    CompterSiMultiple = compteur
End Function
```

**Utilisation :** `=CompterSiMultiple(A1:A10; "pomme"; "poire")` → Compte les cellules contenant "pomme" ou "poire"

### Exemple 9 : Moyenne des nombres pairs uniquement
```vba
Function MoyenneNombresPairs(plage As Range) As Double
    Dim cellule As Range
    Dim somme As Double
    Dim compteur As Integer

    somme = 0
    compteur = 0

    For Each cellule In plage
        If IsNumeric(cellule.Value) Then
            If cellule.Value Mod 2 = 0 Then  ' Nombre pair
                somme = somme + cellule.Value
                compteur = compteur + 1
            End If
        End If
    Next cellule

    If compteur > 0 Then
        MoyenneNombresPairs = somme / compteur
    Else
        MoyenneNombresPairs = 0
    End If
End Function
```

## Gestion des erreurs dans les UDF

### Fonction robuste avec gestion d'erreurs
```vba
Function DivisionSecurisee(dividende As Double, diviseur As Double) As Variant
    ' Division avec gestion de l'erreur de division par zéro

    If diviseur = 0 Then
        DivisionSecurisee = "Erreur : Division par zéro"
    Else
        DivisionSecurisee = dividende / diviseur
    End If
End Function
```

### Fonction avec validation des paramètres
```vba
Function CalculerPourcentage(partie As Double, total As Double) As Variant
    ' Calcule un pourcentage avec validation

    ' Vérifier que les paramètres sont valides
    If total <= 0 Then
        CalculerPourcentage = "Erreur : Le total doit être positif"
        Exit Function
    End If

    If partie < 0 Then
        CalculerPourcentage = "Erreur : La partie ne peut pas être négative"
        Exit Function
    End If

    ' Calcul du pourcentage
    CalculerPourcentage = (partie / total) * 100
End Function
```

## Fonctions avancées : retourner des tableaux

### Exemple 10 : Diviser un nom complet
```vba
Function SeparerNomPrenom(nomComplet As String) As Variant
    ' Sépare un nom complet en prénom et nom
    Dim parties As Variant
    Dim resultat(1 To 2) As String

    parties = Split(Trim(nomComplet), " ")

    If UBound(parties) >= 1 Then
        resultat(1) = parties(0)           ' Prénom
        resultat(2) = parties(UBound(parties)) ' Nom (dernier élément)
    Else
        resultat(1) = nomComplet
        resultat(2) = ""
    End If

    SeparerNomPrenom = resultat
End Function
```

**Utilisation :**
- Sélectionnez 2 cellules horizontales (ex: A1:B1)
- Tapez `=SeparerNomPrenom("Jean Dupont")`
- Appuyez sur Ctrl+Shift+Entrée (formule matricielle)
- Résultat : "Jean" dans A1, "Dupont" dans B1

## Où placer vos fonctions personnalisées ?

### Option 1 : Module standard (recommandé pour débuter)
1. Dans l'éditeur VBA (Alt+F11)
2. Insertion → Module
3. Écrivez vos fonctions dans ce module

### Option 2 : Dans le classeur même
- Les fonctions ne sont disponibles que dans ce classeur
- Parfait pour des fonctions spécifiques à un projet

### Option 3 : Complément Excel (.xlam)
- Pour partager vos fonctions avec d'autres personnes
- Les fonctions deviennent disponibles dans tous les classeurs

## Exemple complet : Système de calcul de commissions

```vba
Function CalculerCommission(chiffreAffaires As Double, anciennete As Integer) As Double
    ' Calcule la commission d'un vendeur selon son CA et son ancienneté
    Dim tauxBase As Double
    Dim bonusAnciennete As Double

    ' Taux de base selon le chiffre d'affaires
    Select Case chiffreAffaires
        Case Is < 10000
            tauxBase = 0.02      ' 2%
        Case Is < 50000
            tauxBase = 0.03      ' 3%
        Case Is < 100000
            tauxBase = 0.04      ' 4%
        Case Else
            tauxBase = 0.05      ' 5%
    End Select

    ' Bonus d'ancienneté
    Select Case anciennete
        Case Is < 2
            bonusAnciennete = 0
        Case Is < 5
            bonusAnciennete = 0.005  ' +0.5%
        Case Else
            bonusAnciennete = 0.01   ' +1%
    End Select

    ' Calcul final
    CalculerCommission = chiffreAffaires * (tauxBase + bonusAnciennete)
End Function

Function StatutVendeur(commission As Double) As String
    ' Détermine le statut du vendeur selon sa commission
    Select Case commission
        Case Is < 1000
            StatutVendeur = "Débutant"
        Case Is < 3000
            StatutVendeur = "Confirmé"
        Case Is < 5000
            StatutVendeur = "Expert"
        Case Else
            StatutVendeur = "Top Performer"
    End Select
End Function
```

**Utilisation pratique :**
- Colonne A : Chiffre d'affaires
- Colonne B : Ancienneté
- Colonne C : `=CalculerCommission(A2;B2)`
- Colonne D : `=StatutVendeur(C2)`

## Comment tester vos fonctions

### Méthode 1 : Directement dans Excel
```
Dans une cellule : =MaFonction(paramètres)
```

### Méthode 2 : Dans l'éditeur VBA
```vba
Sub TesterMesUDF()
    ' Tester vos fonctions avec Debug.Print
    Debug.Print "TVA de 100€ : " & CalculerTVA(100)
    Debug.Print "Prix TTC de 100€ : " & PrixTTC(100)
    Debug.Print "Commission pour 50000€, 3 ans : " & CalculerCommission(50000, 3)

    ' Voir les résultats avec Ctrl+G
End Sub
```

### Méthode 3 : Avec une procédure de test complète
```vba
Sub TestCompletCommissions()
    ' Test avec différents scénarios
    Dim scenarios As Variant
    Dim i As Integer

    scenarios = Array( _
        Array(5000, 1), Array(25000, 3), Array(75000, 6), Array(150000, 10))

    Debug.Print "=== TEST DES COMMISSIONS ==="
    For i = 0 To UBound(scenarios)
        Debug.Print "CA: " & scenarios(i)(0) & "€, Ancienneté: " & scenarios(i)(1) & " ans"
        Debug.Print "Commission: " & CalculerCommission(scenarios(i)(0), scenarios(i)(1)) & "€"
        Debug.Print "Statut: " & StatutVendeur(CalculerCommission(scenarios(i)(0), scenarios(i)(1)))
        Debug.Print "---"
    Next i
End Sub
```

## Conseils pour créer de bonnes UDF

### ✅ Bonnes pratiques

1. **Noms explicites** : `CalculerTVA` plutôt que `TVA`
2. **Paramètres typés** : `As Double`, `As String`, `As Range`
3. **Gestion d'erreurs** : Vérifiez les paramètres d'entrée
4. **Documentation** : Commentez vos fonctions
5. **Tests** : Testez avec différents scénarios

### ⚠️ Pièges à éviter

1. **Fonctions trop complexes** : Une fonction = une tâche précise
2. **Pas de gestion d'erreurs** : Prévoyez les cas d'erreur
3. **Modifications de la feuille** : Les UDF ne doivent que calculer
4. **Noms en conflit** : Évitez les noms des fonctions Excel existantes
5. **Paramètres non typés** : Toujours spécifier les types

### 🎯 Template de base pour vos UDF

```vba
Function MaFonction(parametre1 As Type1, parametre2 As Type2) As TypeRetour
    ' Description : Ce que fait votre fonction
    ' Paramètres :
    '   - parametre1 : Description du paramètre 1
    '   - parametre2 : Description du paramètre 2
    ' Retour : Description de ce qui est retourné

    ' Validation des paramètres
    If [condition d'erreur] Then
        MaFonction = "Erreur : Description"
        Exit Function
    End If

    ' Logique de calcul
    ' ...

    ' Retour du résultat
    MaFonction = resultat
End Function
```

## Rendre vos fonctions disponibles partout

### Créer un complément personnel
1. Sauvegardez votre classeur avec vos UDF
2. Fichier → Enregistrer sous → Type : "Complément Excel (.xlam)"
3. Fichier → Options → Compléments → Gérer les compléments Excel
4. Cochez votre complément

Vos fonctions seront maintenant disponibles dans tous vos classeurs Excel !

## Récapitulatif

Les fonctions personnalisées (UDF) vous permettent de :

- 🔧 **Créer vos propres outils de calcul** adaptés à vos besoins
- 📊 **Simplifier des formules complexes** en une seule fonction claire
- 🔄 **Réutiliser votre code** dans toutes vos feuilles Excel
- 🤝 **Partager vos solutions** avec vos collègues
- ⚡ **Gagner du temps** sur des calculs répétitifs

**Points clés à retenir :**
- Une UDF calcule et retourne une valeur
- Elle s'utilise comme une fonction Excel normale
- Toujours gérer les erreurs possibles
- Nommer clairement et documenter vos fonctions

**Prochaine étape :** Nous verrons maintenant comment automatiser la création et manipulation des graphiques Excel via VBA !

⏭️
