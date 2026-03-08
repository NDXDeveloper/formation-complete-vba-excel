🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.5. Validation des données saisies

## Introduction

La validation des données est un aspect crucial de toute interface utilisateur. Elle garantit que les informations saisies par l'utilisateur sont correctes, complètes et dans le format attendu avant d'être traitées par votre application. Une bonne validation améliore la qualité des données, réduit les erreurs et offre une meilleure expérience utilisateur.

## Pourquoi valider les données ?

### Qualité des données
- Garantir que les données respectent les formats requis
- Éviter les valeurs incohérentes ou impossibles
- Maintenir l'intégrité de votre base de données

### Sécurité
- Prévenir les erreurs de traitement
- Éviter les plantages de l'application
- Protéger contre les saisies malveillantes

### Expérience utilisateur
- Donner un feedback immédiat à l'utilisateur
- Guider l'utilisateur dans sa saisie
- Réduire la frustration liée aux erreurs

## Types de validation

### Validation syntaxique
Vérification du format des données (email, téléphone, code postal, etc.).

### Validation sémantique
Vérification de la logique des données (date cohérente, âge réaliste, etc.).

### Validation de présence
Vérification que les champs obligatoires sont remplis.

### Validation de plage
Vérification que les valeurs sont dans les limites acceptables.

## Moments de validation

### Validation en temps réel (événement Change)
La validation se fait pendant la saisie, caractère par caractère.

```vba
Private Sub txtAge_Change()
    ' Supprimer les caractères non numériques
    Dim i As Integer
    Dim nouveauTexte As String

    For i = 1 To Len(txtAge.Text)
        Dim caractere As String
        caractere = Mid(txtAge.Text, i, 1)

        If IsNumeric(caractere) Then
            nouveauTexte = nouveauTexte & caractere
        End If
    Next i

    ' Mettre à jour si nécessaire
    If nouveauTexte <> txtAge.Text Then
        txtAge.Text = nouveauTexte
    End If
End Sub
```

### Validation à la sortie du champ (événement Exit)
La validation se fait quand l'utilisateur quitte le champ.

```vba
Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validation seulement si le champ n'est pas vide
    If txtEmail.Text <> "" Then
        If Not EstEmailValide(txtEmail.Text) Then
            MsgBox "Format d'email invalide. Exemple : nom@domaine.com", _
                   vbExclamation, "Erreur de saisie"

            ' Highlight du champ en erreur
            txtEmail.BackColor = RGB(255, 200, 200)  ' Rouge clair
            Cancel = True  ' Empêche de quitter le champ
        Else
            ' Restaurer l'apparence normale
            txtEmail.BackColor = RGB(255, 255, 255)  ' Blanc
        End If
    End If
End Sub
```

### Validation globale (avant traitement)
La validation se fait avant de traiter l'ensemble du formulaire.

```vba
Private Sub btnValider_Click()
    ' Validation complète avant traitement
    If Not ValiderFormulaire() Then
        Exit Sub  ' Arrêter si validation échoue
    End If

    ' Traitement des données validées
    TraiterDonnees
End Sub

Private Function ValiderFormulaire() As Boolean
    ValiderFormulaire = True

    ' Vérifier chaque champ
    If Not ValiderNom() Then ValiderFormulaire = False
    If Not ValiderEmail() Then ValiderFormulaire = False
    If Not ValiderAge() Then ValiderFormulaire = False

    ' Afficher message global si erreurs
    If Not ValiderFormulaire Then
        MsgBox "Veuillez corriger les erreurs signalées", vbExclamation
    End If
End Function
```

## Validation des champs texte

### Validation de présence (champs obligatoires)

```vba
Private Function ValiderNom() As Boolean
    ValiderNom = True

    ' Vérifier que le champ n'est pas vide
    If Trim(txtNom.Text) = "" Then
        lblErreurNom.Caption = "Le nom est obligatoire"
        lblErreurNom.ForeColor = RGB(255, 0, 0)  ' Rouge
        txtNom.BackColor = RGB(255, 200, 200)
        ValiderNom = False
    Else
        ' Effacer les messages d'erreur
        lblErreurNom.Caption = ""
        txtNom.BackColor = RGB(255, 255, 255)
    End If
End Function
```

### Validation de longueur

```vba
Private Function ValiderMotDePasse() As Boolean
    ValiderMotDePasse = True
    Dim mdp As String
    mdp = txtMotDePasse.Text

    ' Vérifier la longueur minimale
    If Len(mdp) < 8 Then
        lblErreurMdp.Caption = "Le mot de passe doit contenir au moins 8 caractères"
        lblErreurMdp.ForeColor = RGB(255, 0, 0)
        ValiderMotDePasse = False
    ElseIf Len(mdp) > 20 Then
        lblErreurMdp.Caption = "Le mot de passe ne peut pas dépasser 20 caractères"
        lblErreurMdp.ForeColor = RGB(255, 0, 0)
        ValiderMotDePasse = False
    Else
        lblErreurMdp.Caption = "✓ Longueur correcte"
        lblErreurMdp.ForeColor = RGB(0, 150, 0)  ' Vert
    End If
End Function
```

### Validation de format (expressions régulières simplifiées)

```vba
Private Function EstEmailValide(ByVal email As String) As Boolean
    EstEmailValide = False

    ' Vérifications de base
    If InStr(email, "@") = 0 Then Exit Function  ' Pas d'arobase
    If InStr(email, ".") = 0 Then Exit Function  ' Pas de point
    If Left(email, 1) = "@" Then Exit Function   ' Commence par @
    If Right(email, 1) = "@" Then Exit Function  ' Finit par @
    If InStr(email, "..") > 0 Then Exit Function ' Double point
    If InStr(email, "@.") > 0 Then Exit Function ' @. consécutifs
    If InStr(email, ".@") > 0 Then Exit Function ' .@ consécutifs

    ' Vérifier qu'il n'y a qu'un seul @
    Dim compteurArobase As Integer
    Dim i As Integer
    For i = 1 To Len(email)
        If Mid(email, i, 1) = "@" Then
            compteurArobase = compteurArobase + 1
        End If
    Next i

    If compteurArobase <> 1 Then Exit Function

    ' Si toutes les vérifications passent
    EstEmailValide = True
End Function
```

## Validation des données numériques

### Validation de type numérique

```vba
Private Function ValiderAge() As Boolean
    ValiderAge = True

    ' Vérifier que c'est un nombre
    If Not IsNumeric(txtAge.Text) Then
        lblErreurAge.Caption = "L'âge doit être un nombre"
        lblErreurAge.ForeColor = RGB(255, 0, 0)
        ValiderAge = False
        Exit Function
    End If

    ' Convertir et vérifier la plage
    Dim age As Integer
    age = CInt(txtAge.Text)

    If age < 0 Then
        lblErreurAge.Caption = "L'âge ne peut pas être négatif"
        lblErreurAge.ForeColor = RGB(255, 0, 0)
        ValiderAge = False
    ElseIf age > 150 Then
        lblErreurAge.Caption = "L'âge semble irréaliste"
        lblErreurAge.ForeColor = RGB(255, 0, 0)
        ValiderAge = False
    Else
        lblErreurAge.Caption = "✓ Âge valide"
        lblErreurAge.ForeColor = RGB(0, 150, 0)
    End If
End Function
```

### Validation de montants

```vba
Private Function ValiderMontant() As Boolean
    ValiderMontant = True
    Dim montantTexte As String
    montantTexte = Replace(txtMontant.Text, ",", ".")  ' Normaliser les décimales

    ' Vérifier que c'est numérique
    If Not IsNumeric(montantTexte) Then
        lblErreurMontant.Caption = "Le montant doit être un nombre"
        lblErreurMontant.ForeColor = RGB(255, 0, 0)
        ValiderMontant = False
        Exit Function
    End If

    Dim montant As Double
    montant = CDbl(montantTexte)

    ' Vérifier les limites
    If montant < 0 Then
        lblErreurMontant.Caption = "Le montant ne peut pas être négatif"
        lblErreurMontant.ForeColor = RGB(255, 0, 0)
        ValiderMontant = False
    ElseIf montant > 999999.99 Then
        lblErreurMontant.Caption = "Le montant est trop élevé (max: 999 999,99)"
        lblErreurMontant.ForeColor = RGB(255, 0, 0)
        ValiderMontant = False
    Else
        ' Formater correctement le montant
        txtMontant.Text = Format(montant, "0.00")
        lblErreurMontant.Caption = "✓ Montant valide"
        lblErreurMontant.ForeColor = RGB(0, 150, 0)
    End If
End Function
```

### Filtrage de saisie numérique en temps réel

```vba
Private Sub txtPrix_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57    ' Chiffres 0-9
            ' Autorisé
        Case 44, 46      ' Virgule et point (séparateurs décimaux)
            ' Vérifier qu'il n'y en a pas déjà un
            If InStr(txtPrix.Text, ",") > 0 Or InStr(txtPrix.Text, ".") > 0 Then
                KeyAscii = 0  ' Refuser
            Else
                KeyAscii = 44  ' Forcer la virgule
            End If
        Case 8           ' Backspace (Delete n'est pas reçu par KeyPress)
            ' Autorisé
        Case Else
            KeyAscii = 0  ' Refuser tous les autres caractères
    End Select
End Sub
```

## Validation des dates

### Validation de format de date

```vba
Private Function ValiderDate() As Boolean
    ValiderDate = True

    ' Vérifier que c'est une date valide
    If Not IsDate(txtDate.Text) Then
        lblErreurDate.Caption = "Format de date invalide (jj/mm/aaaa)"
        lblErreurDate.ForeColor = RGB(255, 0, 0)
        ValiderDate = False
        Exit Function
    End If

    Dim dateVerif As Date
    dateVerif = CDate(txtDate.Text)

    ' Vérifier que la date n'est pas dans le futur
    If dateVerif > Date Then
        lblErreurDate.Caption = "La date ne peut pas être dans le futur"
        lblErreurDate.ForeColor = RGB(255, 0, 0)
        ValiderDate = False
    ElseIf dateVerif < DateAdd("yyyy", -150, Date) Then
        lblErreurDate.Caption = "La date semble trop ancienne"
        lblErreurDate.ForeColor = RGB(255, 0, 0)
        ValiderDate = False
    Else
        ' Formater la date correctement
        txtDate.Text = Format(dateVerif, "dd/mm/yyyy")
        lblErreurDate.Caption = "✓ Date valide"
        lblErreurDate.ForeColor = RGB(0, 150, 0)
    End If
End Function
```

### Validation de cohérence entre dates

```vba
Private Function ValiderDates() As Boolean
    ValiderDates = True

    ' Vérifier d'abord que les deux dates sont valides
    If Not IsDate(txtDateDebut.Text) Or Not IsDate(txtDateFin.Text) Then
        lblErreurDates.Caption = "Les deux dates doivent être valides"
        lblErreurDates.ForeColor = RGB(255, 0, 0)
        ValiderDates = False
        Exit Function
    End If

    Dim dateDebut As Date, dateFin As Date
    dateDebut = CDate(txtDateDebut.Text)
    dateFin = CDate(txtDateFin.Text)

    ' Vérifier la cohérence
    If dateFin < dateDebut Then
        lblErreurDates.Caption = "La date de fin doit être postérieure à la date de début"
        lblErreurDates.ForeColor = RGB(255, 0, 0)
        ValiderDates = False
    Else
        lblErreurDates.Caption = "✓ Dates cohérentes"
        lblErreurDates.ForeColor = RGB(0, 150, 0)
    End If
End Function
```

## Validation de sélections (ComboBox, ListBox)

### Validation de sélection obligatoire

```vba
Private Function ValiderPays() As Boolean
    ValiderPays = True

    ' Vérifier qu'une sélection a été faite
    If cmbPays.ListIndex = -1 Then
        lblErreurPays.Caption = "Vous devez sélectionner un pays"
        lblErreurPays.ForeColor = RGB(255, 0, 0)
        cmbPays.BackColor = RGB(255, 200, 200)
        ValiderPays = False
    Else
        lblErreurPays.Caption = ""
        cmbPays.BackColor = RGB(255, 255, 255)
    End If
End Function
```

### Validation de sélection multiple

```vba
Private Function ValiderSelectionProduits() As Boolean
    ValiderSelectionProduits = True
    Dim nombreSelections As Integer
    Dim i As Integer

    ' Compter les sélections
    For i = 0 To lstProduits.ListCount - 1
        If lstProduits.Selected(i) Then
            nombreSelections = nombreSelections + 1
        End If
    Next i

    ' Vérifier qu'au moins un produit est sélectionné
    If nombreSelections = 0 Then
        lblErreurProduits.Caption = "Vous devez sélectionner au moins un produit"
        lblErreurProduits.ForeColor = RGB(255, 0, 0)
        ValiderSelectionProduits = False
    ElseIf nombreSelections > 5 Then
        lblErreurProduits.Caption = "Vous ne pouvez pas sélectionner plus de 5 produits"
        lblErreurProduits.ForeColor = RGB(255, 0, 0)
        ValiderSelectionProduits = False
    Else
        lblErreurProduits.Caption = "✓ " & nombreSelections & " produit(s) sélectionné(s)"
        lblErreurProduits.ForeColor = RGB(0, 150, 0)
    End If
End Function
```

## Validation de cohérence globale

### Validation de règles métier

```vba
Private Function ValiderRegleMetier() As Boolean
    ValiderRegleMetier = True

    ' Exemple : Un mineur ne peut pas avoir de carte de crédit
    If IsNumeric(txtAge.Text) And cmbTypeCarte.Text = "Carte de crédit" Then
        If CInt(txtAge.Text) < 18 Then
            lblErreurRegle.Caption = "Les mineurs ne peuvent pas avoir de carte de crédit"
            lblErreurRegle.ForeColor = RGB(255, 0, 0)
            ValiderRegleMetier = False
        End If
    End If

    ' Exemple : Le salaire doit être cohérent avec l'âge
    If IsNumeric(txtAge.Text) And IsNumeric(txtSalaire.Text) Then
        Dim age As Integer, salaire As Double
        age = CInt(txtAge.Text)
        salaire = CDbl(txtSalaire.Text)

        If age < 16 And salaire > 0 Then
            lblErreurRegle.Caption = "Un mineur de moins de 16 ans ne peut pas avoir de salaire"
            lblErreurRegle.ForeColor = RGB(255, 0, 0)
            ValiderRegleMetier = False
        End If
    End If

    If ValiderRegleMetier Then
        lblErreurRegle.Caption = ""
    End If
End Function
```

## Système de validation centralisé

### Classe de validation

```vba
' Dans un module de classe appelé "ValidateurDonnees"
Option Explicit

Private erreurs As Collection

Public Sub InitialiserValidation()
    Set erreurs = New Collection
End Sub

Public Sub AjouterErreur(ByVal champ As String, ByVal message As String)
    erreurs.Add message, champ
End Sub

Public Function ADesErreurs() As Boolean
    ADesErreurs = (erreurs.Count > 0)
End Function

Public Function ObtenirMessagesErreur() As String
    Dim i As Integer
    Dim messages As String

    For i = 1 To erreurs.Count
        messages = messages & "• " & erreurs(i) & vbCrLf
    Next i

    ObtenirMessagesErreur = messages
End Function

Public Sub EffacerErreurs()
    Set erreurs = New Collection
End Sub
```

### Utilisation du validateur centralisé

```vba
Private validateur As ValidateurDonnees

Private Sub UserForm_Initialize()
    Set validateur = New ValidateurDonnees
End Sub

Private Function ValiderFormulaire() As Boolean
    validateur.InitialiserValidation

    ' Valider chaque champ
    ValiderChampNom
    ValiderChampEmail
    ValiderChampAge
    ValiderChampMontant

    ' Afficher les erreurs s'il y en a
    If validateur.ADesErreurs Then
        MsgBox "Erreurs de validation :" & vbCrLf & vbCrLf & _
               validateur.ObtenirMessagesErreur, _
               vbExclamation, "Validation"
        ValiderFormulaire = False
    Else
        ValiderFormulaire = True
    End If
End Function

Private Sub ValiderChampNom()
    If Trim(txtNom.Text) = "" Then
        validateur.AjouterErreur "nom", "Le nom est obligatoire"
        txtNom.BackColor = RGB(255, 200, 200)
    ElseIf Len(txtNom.Text) < 2 Then
        validateur.AjouterErreur "nom", "Le nom doit contenir au moins 2 caractères"
        txtNom.BackColor = RGB(255, 200, 200)
    Else
        txtNom.BackColor = RGB(255, 255, 255)
    End If
End Sub
```

## Feedback visuel pour la validation

### Indicateurs colorés

```vba
Private Sub AppliquerStyleErreur(ByRef controle As MSForms.TextBox, _
                                ByRef lblErreur As MSForms.Label, _
                                ByVal message As String)
    controle.BackColor = RGB(255, 200, 200)        ' Rouge clair
    controle.BorderColor = RGB(255, 0, 0)          ' Bordure rouge
    lblErreur.Caption = message
    lblErreur.ForeColor = RGB(255, 0, 0)           ' Texte rouge
End Sub

Private Sub AppliquerStyleValide(ByRef controle As MSForms.TextBox, _
                                ByRef lblErreur As MSForms.Label)
    controle.BackColor = RGB(255, 255, 255)       ' Blanc
    controle.BorderColor = RGB(0, 150, 0)         ' Bordure verte
    lblErreur.Caption = "✓"
    lblErreur.ForeColor = RGB(0, 150, 0)          ' Texte vert
End Sub
```

### Icônes de validation

```vba
Private Sub AfficherIconeValidation(ByRef lblIcone As MSForms.Label, _
                                   ByVal estValide As Boolean)
    If estValide Then
        lblIcone.Caption = "✓"
        lblIcone.ForeColor = RGB(0, 150, 0)       ' Vert
        lblIcone.Font.Size = 12
    Else
        lblIcone.Caption = "✗"
        lblIcone.ForeColor = RGB(255, 0, 0)       ' Rouge
        lblIcone.Font.Size = 12
    End If
End Sub
```

## Validation progressive et aide à la saisie

### Validation en temps réel avec guidance

```vba
Private Sub txtMotDePasse_Change()
    Dim mdp As String
    mdp = txtMotDePasse.Text
    Dim force As Integer
    force = CalculerForceMotDePasse(mdp)

    ' Mise à jour de l'indicateur de force
    Select Case force
        Case 0 To 2
            lblForce.Caption = "Faible"
            lblForce.ForeColor = RGB(255, 0, 0)
            barreForce.BackColor = RGB(255, 0, 0)
            barreForce.Width = 30
        Case 3 To 4
            lblForce.Caption = "Moyen"
            lblForce.ForeColor = RGB(255, 165, 0)
            barreForce.BackColor = RGB(255, 165, 0)
            barreForce.Width = 60
        Case 5 To 6
            lblForce.Caption = "Fort"
            lblForce.ForeColor = RGB(0, 150, 0)
            barreForce.BackColor = RGB(0, 150, 0)
            barreForce.Width = 90
    End Select
End Sub

Private Function CalculerForceMotDePasse(ByVal mdp As String) As Integer
    Dim force As Integer
    force = 0

    ' Longueur
    If Len(mdp) >= 8 Then force = force + 1
    If Len(mdp) >= 12 Then force = force + 1

    ' Majuscules
    If mdp <> LCase(mdp) Then force = force + 1

    ' Minuscules
    If mdp <> UCase(mdp) Then force = force + 1

    ' Chiffres
    Dim i As Integer
    For i = 1 To Len(mdp)
        If IsNumeric(Mid(mdp, i, 1)) Then
            force = force + 1
            Exit For
        End If
    Next i

    ' Caractères spéciaux
    If Len(mdp) <> Len(Replace(Replace(Replace(mdp, "!", ""), "@", ""), "#", "")) Then
        force = force + 1
    End If

    CalculerForceMotDePasse = force
End Function
```

## Bonnes pratiques de validation

### Messages d'erreur efficaces

```vba
' ✗ Mauvais : message vague
MsgBox "Erreur de saisie"

' ✓ Bon : message précis et constructif
MsgBox "Le code postal doit contenir exactement 5 chiffres." & vbCrLf & _
       "Exemple : 75001", vbExclamation, "Format incorrect"
```

### Ordre de validation logique

```vba
Private Function ValiderFormulaire() As Boolean
    ' Ordre logique : obligatoire -> format -> cohérence -> règles métier

    ' 1. Champs obligatoires
    If Not ValiderChampsObligatoires() Then
        ValiderFormulaire = False
        Exit Function
    End If

    ' 2. Formats
    If Not ValiderFormats() Then
        ValiderFormulaire = False
        Exit Function
    End If

    ' 3. Cohérence entre champs
    If Not ValiderCoherence() Then
        ValiderFormulaire = False
        Exit Function
    End If

    ' 4. Règles métier
    If Not ValiderRegleMetier() Then
        ValiderFormulaire = False
        Exit Function
    End If

    ValiderFormulaire = True
End Function
```

### Performance de la validation

```vba
' Éviter les validations coûteuses dans Change
Private dernierContenu As String

Private Sub txtRecherche_Change()
    ' Éviter les traitements inutiles
    If txtRecherche.Text = dernierContenu Then Exit Sub

    dernierContenu = txtRecherche.Text

    ' Déclencher la recherche seulement après une pause
    TemporiserRecherche
End Sub
```

---

La validation des données est un investissement essentiel qui améliore considérablement la qualité et la robustesse de vos applications VBA. En appliquant ces techniques de manière cohérente, vous créerez des interfaces fiables qui guident l'utilisateur vers des saisies correctes et préviennent les erreurs de traitement. Une validation bien conçue transforme un simple formulaire en une véritable application professionnelle.

⏭️ [13. Événements](/13-evenements/)
