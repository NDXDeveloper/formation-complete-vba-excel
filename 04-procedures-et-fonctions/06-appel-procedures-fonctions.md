🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.6 Appel de procédures et fonctions

## Introduction

Créer des procédures et des fonctions n'est que la première étape. Pour qu'elles soient utiles, vous devez savoir comment les **appeler** (les exécuter) depuis d'autres parties de votre code. Cette section vous apprendra toutes les méthodes pour utiliser efficacement vos procédures et fonctions.

## Comprendre l'appel avec une analogie

### L'analogie du téléphone

Appeler une procédure ou une fonction, c'est comme passer un coup de téléphone :

- **Composer le numéro** = Écrire le nom de la procédure/fonction
- **Transmettre le message** = Passer les arguments (paramètres)
- **Recevoir une réponse** = Récupérer la valeur de retour (pour les fonctions)
- **Raccrocher** = Continuer avec le reste du code

## Appel de procédures (Sub)

### Méthode 1 : Appel direct (recommandée)

La méthode la plus simple et la plus courante :

```vba
Sub ProcedurePrincipale()
    ' Appel direct d'une procédure sans paramètres
    AfficherMessage

    ' Appel direct d'une procédure avec paramètres
    FormaterCellule "A1", "Titre", 16

    ' Plusieurs appels successifs
    EffacerZone "A1:D10"
    CreerEntete "RAPPORT MENSUEL"
    AjouterDate
End Sub

' Les procédures appelées
Sub AfficherMessage()
    MsgBox "Procédure exécutée avec succès !"
End Sub

Sub FormaterCellule(adresse As String, texte As String, taille As Integer)
    With Range(adresse)
        .Value = texte
        .Font.Size = taille
        .Font.Bold = True
    End With
End Sub

Sub EffacerZone(plage As String)
    Range(plage).ClearContents
End Sub

Sub CreerEntete(titre As String)
    Range("A1").Value = titre
    Range("A1").Font.Size = 18
End Sub

Sub AjouterDate()
    Range("A2").Value = "Généré le : " & Format(Date, "dd/mm/yyyy")
End Sub
```

### Méthode 2 : Avec le mot-clé Call

Cette méthode est plus formelle mais moins utilisée :

```vba
Sub ExempleAvecCall()
    ' Avec Call, les paramètres doivent être entre parenthèses
    Call AfficherMessage()
    Call FormaterCellule("B1", "Sous-titre", 12)
    Call EffacerZone("B1:E15")
End Sub
```

**Comparaison des deux méthodes :**
```vba
' ✅ Appel direct (préféré)
FormaterCellule "A1", "Titre", 14

' ✅ Avec Call (plus verbeux)
Call FormaterCellule("A1", "Titre", 14)
```

### Appel depuis différents endroits

**Depuis une autre procédure :**
```vba
Sub ProcedureA()
    MsgBox "Début de la procédure A"
    ProcedureB  ' Appel d'une autre procédure
    MsgBox "Fin de la procédure A"
End Sub

Sub ProcedureB()
    MsgBox "Exécution de la procédure B"
End Sub
```

**Depuis l'éditeur VBA :**
- Placez le curseur dans la procédure
- Appuyez sur **F5** ou cliquez sur **Exécuter**

**Depuis Excel :**
- **Alt + F8** pour ouvrir la liste des macros
- Sélectionnez votre procédure et cliquez **Exécuter**

## Appel de fonctions

### Récupération de la valeur de retour

Les fonctions retournent une valeur que vous devez **récupérer** :

```vba
Sub UtiliserFonctions()
    Dim resultat As Double
    Dim message As String
    Dim estValide As Boolean

    ' Récupération dans une variable
    resultat = CalculerTVA(100)
    MsgBox "TVA : " & resultat & "€"

    ' Utilisation directe dans une expression
    MsgBox "Prix TTC : " & (100 + CalculerTVA(100)) & "€"

    ' Récupération de différents types
    message = FormaterNom("Jean", "Dupont")
    estValide = EstMajeur(25)

    MsgBox message & " - Majeur : " & estValide
End Sub

' Les fonctions utilisées
Function CalculerTVA(prix As Double) As Double
    CalculerTVA = prix * 0.2
End Function

Function FormaterNom(prenom As String, nom As String) As String
    FormaterNom = UCase(nom) & ", " & prenom
End Function

Function EstMajeur(age As Integer) As Boolean
    EstMajeur = (age >= 18)
End Function
```

### Utilisation dans des expressions

```vba
Sub ExpressionsAvecFonctions()
    Dim prix As Double
    Dim prixFinal As Double

    prix = 150

    ' Fonction dans un calcul
    prixFinal = prix + CalculerTVA(prix) - CalculerRemise(prix, 10)

    ' Fonction dans une condition
    If EstMajeur(InputBox("Votre âge ?")) Then
        MsgBox "Accès autorisé"
    Else
        MsgBox "Accès refusé"
    End If

    ' Fonction dans une concatenation
    MsgBox "Client : " & FormaterNom("Marie", "Martin") & _
           " - Montant : " & prixFinal & "€"
End Sub

Function CalculerRemise(prix As Double, pourcentage As Double) As Double
    CalculerRemise = prix * (pourcentage / 100)
End Function
```

### Fonctions dans les cellules Excel

Vos fonctions personnalisées peuvent être utilisées comme des formules Excel :

```vba
' Fonction utilisable dans Excel
Function ConvertirEnMajuscules(texte As String) As String
    ConvertirEnMajuscules = UCase(texte)
End Function

Function CalculerAge(dateNaissance As Date) As Integer
    CalculerAge = Year(Date) - Year(dateNaissance)
End Function

Function EstEmail(texte As String) As Boolean
    EstEmail = (InStr(texte, "@") > 0 And InStr(texte, ".") > 0)
End Function
```

**Dans Excel, tapez :**
- `=ConvertirEnMajuscules("bonjour")` → BONJOUR
- `=CalculerAge("15/03/1990")` → 34 (selon l'année actuelle)
- `=EstEmail("test@email.com")` → VRAI

## Passage de paramètres avancé

### Avec noms des paramètres

Pour plus de clarté, vous pouvez nommer les paramètres :

```vba
Sub ExempleParametresNommes()
    ' Appel classique
    CreerRapport "Ventes", "2024", True, "Excel"

    ' Appel avec noms de paramètres (plus clair)
    CreerRapport titre:="Ventes", _
                 annee:="2024", _
                 inclure_graphique:=True, _
                 format_sortie:="Excel"

    ' Mélange possible (paramètres nommés en fin)
    CreerRapport "Marketing", "2024", inclure_graphique:=False
End Sub

Sub CreerRapport(titre As String, annee As String, inclure_graphique As Boolean, Optional format_sortie As String = "PDF")
    MsgBox "Rapport " & titre & " pour " & annee & _
           " - Graphique : " & inclure_graphique & _
           " - Format : " & format_sortie
End Sub
```

### Paramètres optionnels

```vba
Sub TestParametresOptionnels()
    ' Utilisation avec tous les paramètres
    EnvoyerEmail "jean@email.com", "Rapport", "Voir fichier joint", "Urgent"

    ' Utilisation avec paramètres optionnels omis
    EnvoyerEmail "marie@email.com", "Information", "Message simple"

    ' Avec paramètres nommés (pour sauter des optionnels)
    EnvoyerEmail destinataire:="paul@email.com", _
                 sujet:="Réunion", _
                 priorite:="Haute"  ' Corps omis
End Sub

Sub EnvoyerEmail(destinataire As String, _
                 sujet As String, _
                 Optional corps As String = "", _
                 Optional priorite As String = "Normale")
    MsgBox "Email à : " & destinataire & vbNewLine & _
           "Sujet : " & sujet & vbNewLine & _
           "Corps : " & corps & vbNewLine & _
           "Priorité : " & priorite
End Sub
```

## Appels imbriqués et chaînage

### Procédures qui appellent d'autres procédures

```vba
Sub ProcessusComplet()
    MsgBox "Début du processus"

    InitialiserDonnees
    TraiterDonnees
    GenererRapport

    MsgBox "Processus terminé"
End Sub

Sub InitialiserDonnees()
    MsgBox "Initialisation des données"
    EffacerAnciennesDonnees
    CreerStructure
End Sub

Sub TraiterDonnees()
    MsgBox "Traitement des données"
    ValiderDonnees
    CalculerStatistiques
End Sub

Sub GenererRapport()
    MsgBox "Génération du rapport"
    CreerEntetes
    RemplirDonnees
    FormaterPresentation
End Sub

' Procédures de niveau inférieur
Sub EffacerAnciennesDonnees()
    Range("A:Z").ClearContents
End Sub

Sub CreerStructure()
    Range("A1").Value = "Nom"
    Range("B1").Value = "Âge"
    Range("C1").Value = "Ville"
End Sub

Sub ValiderDonnees()
    MsgBox "Validation en cours..."
End Sub

Sub CalculerStatistiques()
    MsgBox "Calcul des statistiques..."
End Sub

Sub CreerEntetes()
    Range("A1:C1").Font.Bold = True
End Sub

Sub RemplirDonnees()
    Range("A2").Value = "Jean Dupont"
    Range("B2").Value = 30
    Range("C2").Value = "Paris"
End Sub

Sub FormaterPresentation()
    Range("A:C").AutoFit
End Sub
```

### Fonctions qui appellent d'autres fonctions

```vba
Function CalculerPrixFinal(prixBase As Double, categorie As String) As Double
    Dim tva As Double
    Dim remise As Double

    ' Appel de fonctions pour obtenir TVA et remise
    tva = CalculerTVA(prixBase)
    remise = CalculerRemiseParCategorie(prixBase, categorie)

    CalculerPrixFinal = prixBase + tva - remise
End Function

Function CalculerTVA(prix As Double) As Double
    CalculerTVA = prix * 0.2
End Function

Function CalculerRemiseParCategorie(prix As Double, categorie As String) As Double
    Select Case LCase(categorie)
        Case "vip"
            CalculerRemiseParCategorie = prix * 0.15  ' 15% pour VIP
        Case "regulier"
            CalculerRemiseParCategorie = prix * 0.05  ' 5% pour régulier
        Case Else
            CalculerRemiseParCategorie = 0  ' Pas de remise
    End Select
End Function

' Utilisation
Sub TestCalculPrix()
    Dim prix As Double

    prix = CalculerPrixFinal(100, "VIP")
    MsgBox "Prix final VIP : " & prix & "€"

    prix = CalculerPrixFinal(100, "regulier")
    MsgBox "Prix final régulier : " & prix & "€"
End Sub
```

## Gestion des erreurs lors des appels

### Vérification avant appel

```vba
Sub AppelSecurise()
    Dim age As Variant

    age = InputBox("Entrez votre âge :")

    ' Vérification avant d'appeler la fonction
    If IsNumeric(age) Then
        If EstMajeur(CInt(age)) Then
            AutoriserAcces
        Else
            RefuserAcces
        End If
    Else
        MsgBox "Veuillez entrer un nombre valide"
    End If
End Sub

Sub AutoriserAcces()
    MsgBox "Accès autorisé !"
End Sub

Sub RefuserAcces()
    MsgBox "Accès refusé - Vous devez être majeur"
End Sub
```

### Gestion d'erreur dans les fonctions

```vba
Function DivisionSecurisee(dividende As Double, diviseur As Double) As Variant
    If diviseur = 0 Then
        DivisionSecurisee = "Erreur : Division par zéro"
    Else
        DivisionSecurisee = dividende / diviseur
    End If
End Function

Sub TestDivision()
    Dim resultat As Variant

    resultat = DivisionSecurisee(10, 2)
    MsgBox "10 ÷ 2 = " & resultat

    resultat = DivisionSecurisee(10, 0)
    MsgBox "10 ÷ 0 = " & resultat  ' Affichera le message d'erreur
End Sub
```

## Appels depuis différents modules

### Appel de procédures publiques

```vba
' === MODULE1 ===
Public Sub ProcedurePublique()
    MsgBox "Procédure publique du Module1"
End Sub

Private Sub ProcedurePrivee()
    MsgBox "Procédure privée du Module1"
End Sub

' === MODULE2 ===
Sub AppelerDepuisAutreModule()
    ' ✅ Fonctionne - procédure publique
    ProcedurePublique

    ' ❌ Erreur - procédure privée non accessible
    ' ProcedurePrivee  ' Cette ligne causerait une erreur
End Sub
```

### Appel explicite avec nom de module

```vba
' === MODULE1 ===
Public Sub TraiterDonnees()
    MsgBox "Traitement depuis Module1"
End Sub

' === MODULE2 ===
Public Sub TraiterDonnees()
    MsgBox "Traitement depuis Module2"
End Sub

Sub AppelExplicite()
    ' Appel de la procédure du module spécifique
    Module1.TraiterDonnees
    Module2.TraiterDonnees

    ' Sans précision, appelle celle du module actuel
    TraiterDonnees
End Sub
```

## Bonnes pratiques pour les appels

### 1. Vérification des paramètres

```vba
Sub TraiterAge(age As Integer)
    ' Vérification en début de procédure
    If age < 0 Or age > 150 Then
        MsgBox "Âge invalide : " & age
        Exit Sub
    End If

    ' Traitement normal
    If age >= 18 Then
        MsgBox "Personne majeure"
    Else
        MsgBox "Personne mineure"
    End If
End Sub
```

### 2. Documentation des appels complexes

```vba
Sub ProcessusComplexe()
    ' Étape 1 : Préparation des données
    InitialiserEnvironnement

    ' Étape 2 : Validation des entrées
    If Not ValiderDonneesEntree() Then
        MsgBox "Données invalides - Arrêt du processus"
        Exit Sub
    End If

    ' Étape 3 : Traitement principal
    ExecuterCalculsComplexes

    ' Étape 4 : Finalisation
    SauvegarderResultats
    NettoyerEnvironnement
End Sub
```

### 3. Gestion des valeurs de retour

```vba
Sub GererRetoursCorrectement()
    Dim succes As Boolean
    Dim message As String

    ' Récupération et vérification du retour
    succes = ExecuterOperation()

    If succes Then
        message = "Opération réussie"
    Else
        message = "Opération échouée"
    End If

    MsgBox message
End Sub

Function ExecuterOperation() As Boolean
    ' Simulation d'une opération qui peut échouer
    If Rnd() > 0.5 Then
        ExecuterOperation = True
    Else
        ExecuterOperation = False
    End If
End Function
```

## Débogage des appels

### Techniques de suivi

```vba
Sub ProcedureAvecSuivi()
    Debug.Print "Début de ProcedureAvecSuivi"

    Dim resultat As Double
    Debug.Print "Appel de CalculerValeur"
    resultat = CalculerValeur(10, 5)
    Debug.Print "Résultat reçu : " & resultat

    Debug.Print "Appel de AfficherResultat"
    AfficherResultat resultat

    Debug.Print "Fin de ProcedureAvecSuivi"
End Sub

Function CalculerValeur(a As Double, b As Double) As Double
    Debug.Print "Dans CalculerValeur - a=" & a & ", b=" & b
    CalculerValeur = a * b
    Debug.Print "Retour de CalculerValeur : " & CalculerValeur
End Function

Sub AfficherResultat(valeur As Double)
    Debug.Print "Dans AfficherResultat - valeur=" & valeur
    MsgBox "Résultat : " & valeur
End Sub
```

### Points d'arrêt pour le débogage

```vba
Sub ProcedureDeDebug()
    Dim i As Integer

    For i = 1 To 5
        ' Placez un point d'arrêt ici (F9)
        Debug.Print "Itération : " & i

        ' Appelez une fonction à déboguer
        TraiterIteration i
    Next i
End Sub

Sub TraiterIteration(numero As Integer)
    ' Cette procédure sera également déboguée
    Range("A" & numero).Value = "Ligne " & numero
End Sub
```

## Récapitulatif des concepts clés

1. **Appel de procédures** : Nom direct ou avec `Call`
2. **Appel de fonctions** : Récupération obligatoire de la valeur de retour
3. **Paramètres nommés** : Pour plus de clarté dans les appels complexes
4. **Appels imbriqués** : Procédures/fonctions qui s'appellent entre elles
5. **Portée des appels** : Public/Private détermine l'accessibilité
6. **Gestion d'erreurs** : Vérifier les paramètres et les retours
7. **Débogage** : Utiliser Debug.Print et les points d'arrêt

La maîtrise des appels de procédures et fonctions vous permet de créer des programmes modulaires, réutilisables et maintenables. C'est la clé pour transformer vos scripts simples en véritables applications robustes !

⏭️
