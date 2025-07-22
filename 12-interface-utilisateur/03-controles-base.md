🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.3. Contrôles de base (TextBox, ComboBox, ListBox)

## Introduction

Les contrôles sont les éléments interactifs qui composent votre UserForm. Ils permettent à l'utilisateur de saisir des données, faire des choix et naviguer dans votre interface. Dans cette section, nous allons explorer en détail trois contrôles essentiels : TextBox (zone de texte), ComboBox (liste déroulante) et ListBox (liste de sélection).

Maîtriser ces contrôles vous permettra de créer des interfaces complètes et professionnelles capables de gérer la plupart des besoins de saisie de données.

## TextBox (Zone de texte)

### Qu'est-ce qu'une TextBox ?

Une TextBox est un contrôle qui permet à l'utilisateur de saisir et modifier du texte. C'est l'un des contrôles les plus utilisés dans les formulaires car il offre une grande flexibilité pour la collecte d'informations.

### Ajout d'une TextBox

1. Dans la boîte à outils, cliquez sur l'icône **TextBox**
2. Cliquez et tirez sur votre formulaire pour créer la zone de texte
3. Ajustez la taille selon vos besoins

### Propriétés importantes

#### Propriétés de base

**Name** : Le nom du contrôle utilisé dans le code
```vba
' Exemple de nommage
Name: txtNom          ' Pour un nom
Name: txtEmail        ' Pour un email
Name: txtCommentaire  ' Pour un commentaire
```

**Text** : Le contenu actuel de la zone de texte
```vba
' Lecture du contenu
Dim contenu As String
contenu = txtNom.Text

' Modification du contenu
txtNom.Text = "Nouveau texte"
```

**Value** : Synonyme de Text (même fonction)
```vba
' Équivalent à .Text
txtNom.Value = "Même résultat que .Text"
```

#### Propriétés de formatage

**Font** : Police, taille et style du texte
- Cliquez sur le bouton "..." pour ouvrir la boîte de dialogue des polices
- Ou modifiez directement : Font.Name, Font.Size, Font.Bold

**ForeColor** et **BackColor** : Couleurs du texte et de l'arrière-plan
```vba
' Dans le code, pour changer les couleurs
txtNom.ForeColor = RGB(0, 0, 255)    ' Texte bleu
txtNom.BackColor = RGB(255, 255, 0)  ' Fond jaune
```

**TextAlign** : Alignement du texte
- 1 - fmTextAlignLeft (gauche)
- 2 - fmTextAlignCenter (centré)
- 3 - fmTextAlignRight (droite)

#### Propriétés de comportement

**MaxLength** : Nombre maximum de caractères autorisés
```vba
' Limiter à 50 caractères
MaxLength: 50

' Pas de limite (par défaut)
MaxLength: 0
```

**PasswordChar** : Caractère de masquage pour les mots de passe
```vba
' Masquer avec des étoiles
PasswordChar: *

' Aucun masquage (par défaut)
PasswordChar: (vide)
```

**MultiLine** : Autoriser plusieurs lignes
```vba
' Permettre plusieurs lignes
MultiLine: True

' Une seule ligne (par défaut)
MultiLine: False
```

**ScrollBars** : Barres de défilement (utile avec MultiLine)
- 0 - fmScrollBarsNone (aucune)
- 1 - fmScrollBarsHorizontal (horizontale)
- 2 - fmScrollBarsVertical (verticale)
- 3 - fmScrollBarsBoth (les deux)

**Locked** : Empêcher la modification
```vba
' Zone de texte en lecture seule
Locked: True

' Modification autorisée (par défaut)
Locked: False
```

### Événements principaux

#### Change : Se déclenche à chaque modification

```vba
Private Sub txtNom_Change()
    ' Se déclenche à chaque caractère tapé
    If Len(txtNom.Text) > 0 Then
        ' Première lettre en majuscule
        txtNom.Text = UCase(Left(txtNom.Text, 1)) & Mid(txtNom.Text, 2)
    End If
End Sub
```

#### Exit : Se déclenche en quittant le contrôle

```vba
Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validation de l'email
    If txtEmail.Text <> "" Then
        If InStr(txtEmail.Text, "@") = 0 Then
            MsgBox "Format d'email invalide", vbExclamation
            Cancel = True  ' Empêche de quitter le champ
        End If
    End If
End Sub
```

#### Enter : Se déclenche en entrant dans le contrôle

```vba
Private Sub txtMontant_Enter()
    ' Sélectionner tout le contenu pour faciliter la saisie
    txtMontant.SelStart = 0
    txtMontant.SelLength = Len(txtMontant.Text)
End Sub
```

### Méthodes utiles

```vba
' Donner le focus au contrôle
txtNom.SetFocus

' Sélectionner une partie du texte
txtNom.SelStart = 0      ' Position de début
txtNom.SelLength = 3     ' Nombre de caractères

' Sélectionner tout le texte
txtNom.SelStart = 0
txtNom.SelLength = Len(txtNom.Text)
```

### Exemples pratiques

#### Zone de texte pour nom avec validation

```vba
Private Sub txtNom_Change()
    ' Supprimer les chiffres
    Dim i As Integer
    Dim nouveauTexte As String

    For i = 1 To Len(txtNom.Text)
        Dim caractere As String
        caractere = Mid(txtNom.Text, i, 1)
        If Not IsNumeric(caractere) Then
            nouveauTexte = nouveauTexte & caractere
        End If
    Next i

    If nouveauTexte <> txtNom.Text Then
        txtNom.Text = nouveauTexte
    End If
End Sub
```

#### Zone de texte numérique

```vba
Private Sub txtAge_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Autoriser seulement les chiffres et les touches de contrôle
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 127 Then  ' Backspace et Delete
            KeyAscii = 0  ' Annule la frappe
        End If
    End If
End Sub
```

## ComboBox (Liste déroulante)

### Qu'est-ce qu'une ComboBox ?

Une ComboBox combine une zone de texte et une liste déroulante. L'utilisateur peut soit choisir une valeur dans la liste, soit taper une nouvelle valeur. C'est parfait pour offrir des choix prédéfinis tout en gardant la flexibilité.

### Ajout d'une ComboBox

1. Cliquez sur l'icône **ComboBox** dans la boîte à outils
2. Dessinez le contrôle sur votre formulaire
3. Ajustez la largeur (la hauteur s'ajuste automatiquement)

### Propriétés importantes

#### Style de ComboBox

**Style** : Détermine le comportement de la liste
- 0 - fmStyleDropDownCombo : Liste + saisie libre (par défaut)
- 1 - fmStyleSimpleCombo : Liste toujours visible + saisie
- 2 - fmStyleDropDownList : Liste seulement, pas de saisie libre

```vba
' Pour forcer le choix dans la liste uniquement
Style: 2 - fmStyleDropDownList
```

#### Gestion de la liste

**RowSource** : Source des données (plage Excel)
```vba
' Utiliser une plage nommée
RowSource: "ListePays"

' Utiliser une plage directe
RowSource: "Feuil1!A1:A10"
```

### Méthodes pour gérer les éléments

#### AddItem : Ajouter un élément

```vba
' Ajouter un élément simple
cmbPays.AddItem "France"
cmbPays.AddItem "Espagne"
cmbPays.AddItem "Italie"

' Ajouter à une position spécifique
cmbPays.AddItem "Allemagne", 1  ' Insérer en 2ème position
```

#### RemoveItem : Supprimer un élément

```vba
' Supprimer par index (commence à 0)
cmbPays.RemoveItem 0  ' Supprime le premier élément

' Supprimer l'élément sélectionné
If cmbPays.ListIndex >= 0 Then
    cmbPays.RemoveItem cmbPays.ListIndex
End If
```

#### Clear : Vider la liste

```vba
' Supprimer tous les éléments
cmbPays.Clear
```

### Propriétés de sélection

**ListIndex** : Index de l'élément sélectionné
```vba
' Obtenir l'index sélectionné (-1 si aucune sélection)
Dim index As Integer
index = cmbPays.ListIndex

' Sélectionner un élément par index
cmbPays.ListIndex = 0  ' Sélectionner le premier
```

**Text** ou **Value** : Texte de l'élément sélectionné
```vba
' Obtenir la valeur sélectionnée
Dim selection As String
selection = cmbPays.Text

' Définir une sélection par texte
cmbPays.Text = "France"
```

**ListCount** : Nombre d'éléments dans la liste
```vba
Dim nombreElements As Integer
nombreElements = cmbPays.ListCount
```

### Initialisation d'une ComboBox

#### Méthode 1 : Dans l'événement Initialize du formulaire

```vba
Private Sub UserForm_Initialize()
    ' Remplir la liste des pays
    With cmbPays
        .AddItem "France"
        .AddItem "Espagne"
        .AddItem "Italie"
        .AddItem "Allemagne"
        .AddItem "Belgique"

        ' Sélectionner par défaut
        .ListIndex = 0  ' Sélectionne "France"
    End With
End Sub
```

#### Méthode 2 : Depuis une plage Excel

```vba
Private Sub UserForm_Initialize()
    ' Remplir depuis une plage
    Dim i As Integer
    Dim plage As Range

    Set plage = Worksheets("Données").Range("A1:A5")

    cmbPays.Clear
    For i = 1 To plage.Rows.Count
        If plage.Cells(i, 1).Value <> "" Then
            cmbPays.AddItem plage.Cells(i, 1).Value
        End If
    Next i
End Sub
```

#### Méthode 3 : Avec un tableau

```vba
Private Sub UserForm_Initialize()
    Dim pays() As String
    Dim i As Integer

    ' Définir le tableau
    pays = Split("France,Espagne,Italie,Allemagne,Belgique", ",")

    ' Remplir la ComboBox
    cmbPays.Clear
    For i = 0 To UBound(pays)
        cmbPays.AddItem pays(i)
    Next i
End Sub
```

### Événements principaux

#### Change : Modification de la sélection

```vba
Private Sub cmbPays_Change()
    ' Réagir au changement de sélection
    If cmbPays.Text = "France" Then
        txtDevise.Text = "EUR"
    ElseIf cmbPays.Text = "Royaume-Uni" Then
        txtDevise.Text = "GBP"
    Else
        txtDevise.Text = "EUR"  ' Par défaut
    End If
End Sub
```

## ListBox (Liste de sélection)

### Qu'est-ce qu'une ListBox ?

Une ListBox affiche une liste d'éléments où l'utilisateur peut faire une ou plusieurs sélections. Contrairement à la ComboBox, tous les éléments sont visibles simultanément (dans la limite de la taille du contrôle).

### Ajout d'une ListBox

1. Cliquez sur l'icône **ListBox** dans la boîte à outils
2. Dessinez le contrôle sur votre formulaire
3. Ajustez la taille selon le nombre d'éléments à afficher

### Propriétés importantes

#### Sélection multiple

**MultiSelect** : Type de sélection multiple
- 0 - fmMultiSelectSingle : Sélection unique (par défaut)
- 1 - fmMultiSelectMulti : Sélections multiples avec Ctrl+clic
- 2 - fmMultiSelectExtended : Sélections multiples avec Ctrl et Shift

```vba
' Permettre la sélection multiple
MultiSelect: 1 - fmMultiSelectMulti
```

#### Colonnes multiples

**ColumnCount** : Nombre de colonnes à afficher
```vba
' Afficher 3 colonnes
ColumnCount: 3
```

**ColumnWidths** : Largeur de chaque colonne
```vba
' Largeurs : 100pt, 80pt, 120pt
ColumnWidths: 100pt;80pt;120pt

' Largeurs automatiques (égales)
ColumnWidths: (vide)
```

### Méthodes pour gérer les éléments

#### Ajouter des éléments simples

```vba
Private Sub UserForm_Initialize()
    ' Ajouter des éléments simples
    With lstProduits
        .AddItem "Ordinateur portable"
        .AddItem "Souris"
        .AddItem "Clavier"
        .AddItem "Écran"
        .AddItem "Imprimante"
    End With
End Sub
```

#### Ajouter des éléments multi-colonnes

```vba
Private Sub UserForm_Initialize()
    ' Configuration pour 3 colonnes
    lstProduits.ColumnCount = 3
    lstProduits.ColumnWidths = "100pt;80pt;60pt"

    ' Ajouter un élément
    lstProduits.AddItem "Ordinateur"

    ' Remplir les autres colonnes
    Dim derniereRow As Integer
    derniereRow = lstProduits.ListCount - 1

    lstProduits.List(derniereRow, 1) = "1200€"
    lstProduits.List(derniereRow, 2) = "5"
End Sub
```

#### Remplir depuis un tableau 2D

```vba
Private Sub RemplirListeProduits()
    Dim donnees(4, 2) As String
    Dim i As Integer

    ' Préparation des données
    donnees(0, 0) = "Ordinateur": donnees(0, 1) = "1200€": donnees(0, 2) = "5"
    donnees(1, 0) = "Souris": donnees(1, 1) = "25€": donnees(1, 2) = "20"
    donnees(2, 0) = "Clavier": donnees(2, 1) = "45€": donnees(2, 2) = "15"
    donnees(3, 0) = "Écran": donnees(3, 1) = "300€": donnees(3, 2) = "8"
    donnees(4, 0) = "Imprimante": donnees(4, 1) = "150€": donnees(4, 2) = "3"

    ' Configuration de la ListBox
    With lstProduits
        .ColumnCount = 3
        .ColumnWidths = "100pt;80pt;60pt"
        .Clear

        ' Remplissage
        For i = 0 To 4
            .AddItem donnees(i, 0)
            .List(.ListCount - 1, 1) = donnees(i, 1)
            .List(.ListCount - 1, 2) = donnees(i, 2)
        Next i
    End With
End Sub
```

### Gestion de la sélection

#### Sélection unique

```vba
Private Sub lstProduits_Click()
    If lstProduits.ListIndex >= 0 Then
        Dim produitSelectione As String
        produitSelectione = lstProduits.Text
        MsgBox "Vous avez sélectionné : " & produitSelectione
    End If
End Sub
```

#### Sélection multiple

```vba
Private Sub btnVoirSelection_Click()
    Dim i As Integer
    Dim selections As String

    ' Parcourir tous les éléments
    For i = 0 To lstProduits.ListCount - 1
        If lstProduits.Selected(i) Then
            If selections <> "" Then selections = selections & ", "
            selections = selections & lstProduits.List(i)
        End If
    Next i

    If selections = "" Then
        MsgBox "Aucune sélection"
    Else
        MsgBox "Éléments sélectionnés : " & selections
    End If
End Sub
```

#### Compter les sélections

```vba
Function NombreSelections() As Integer
    Dim i As Integer
    Dim compteur As Integer

    compteur = 0
    For i = 0 To lstProduits.ListCount - 1
        If lstProduits.Selected(i) Then
            compteur = compteur + 1
        End If
    Next i

    NombreSelections = compteur
End Function
```

### Recherche dans une ListBox

```vba
Private Sub txtRecherche_Change()
    Dim i As Integer
    Dim textRecherche As String

    textRecherche = UCase(txtRecherche.Text)

    ' Parcourir la liste et sélectionner le premier élément correspondant
    For i = 0 To lstProduits.ListCount - 1
        If UCase(Left(lstProduits.List(i), Len(textRecherche))) = textRecherche Then
            lstProduits.ListIndex = i
            Exit For
        End If
    Next i
End Sub
```

## Intégration des trois contrôles

### Exemple complet : Formulaire de commande

```vba
' Événement d'initialisation du formulaire
Private Sub UserForm_Initialize()
    ' Initialiser la ComboBox des clients
    With cmbClient
        .AddItem "Entreprise A"
        .AddItem "Entreprise B"
        .AddItem "Entreprise C"
        .ListIndex = 0
    End With

    ' Initialiser la ListBox des produits
    With lstProduits
        .ColumnCount = 3
        .ColumnWidths = "120pt;80pt;60pt"
        .MultiSelect = fmMultiSelectMulti

        .AddItem "Ordinateur portable"
        .List(.ListCount - 1, 1) = "1200 €"
        .List(.ListCount - 1, 2) = "10"

        .AddItem "Souris sans fil"
        .List(.ListCount - 1, 1) = "25 €"
        .List(.ListCount - 1, 2) = "50"

        .AddItem "Clavier mécanique"
        .List(.ListCount - 1, 1) = "85 €"
        .List(.ListCount - 1, 2) = "25"
    End With

    ' Initialiser les zones de texte
    txtQuantite.Text = "1"
    txtCommentaire.MultiLine = True
    txtCommentaire.ScrollBars = fmScrollBarsVertical
End Sub

' Validation et traitement de la commande
Private Sub btnValider_Click()
    ' Vérifier que les champs obligatoires sont remplis
    If cmbClient.Text = "" Then
        MsgBox "Veuillez sélectionner un client.", vbExclamation
        cmbClient.SetFocus
        Exit Sub
    End If

    If NombreSelections = 0 Then
        MsgBox "Veuillez sélectionner au moins un produit.", vbExclamation
        lstProduits.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtQuantite.Text) Or Val(txtQuantite.Text) <= 0 Then
        MsgBox "Veuillez entrer une quantité valide.", vbExclamation
        txtQuantite.SetFocus
        Exit Sub
    End If

    ' Traitement de la commande
    MsgBox "Commande validée pour " & cmbClient.Text & " !" & vbCrLf & _
           "Produits sélectionnés : " & NombreSelections & vbCrLf & _
           "Quantité : " & txtQuantite.Text, vbInformation

    Me.Hide
End Sub
```

## Conseils et bonnes pratiques

### Performance

**Désactiver les événements pendant le remplissage :**
```vba
' Pendant le remplissage d'une grande liste
Application.EnableEvents = False
' ... code de remplissage ...
Application.EnableEvents = True
```

**Utiliser Clear avant de remplir :**
```vba
' Toujours vider avant de remplir
cmbPays.Clear
' ... ajout des éléments ...
```

### Validation

**Valider les données à la saisie :**
```vba
Private Sub txtQuantite_Change()
    ' Supprimer les caractères non numériques
    If Not IsNumeric(txtQuantite.Text) And txtQuantite.Text <> "" Then
        txtQuantite.Text = ""
    End If
End Sub
```

**Valider avant traitement :**
```vba
Private Sub btnValider_Click()
    ' Toujours valider avant de traiter
    If Not ValiderFormulaire() Then
        Exit Sub
    End If
    ' ... traitement ...
End Sub

Private Function ValiderFormulaire() As Boolean
    ValiderFormulaire = True

    If cmbClient.ListIndex = -1 Then
        MsgBox "Sélection client requise"
        ValiderFormulaire = False
    End If

    ' ... autres validations ...
End Function
```

### Accessibilité

**Ordre de tabulation logique :**
- Définissez TabIndex dans l'ordre de saisie logique
- Commencez par 0 pour le premier contrôle

**Raccourcis clavier :**
```vba
' Utilisez la propriété Accelerator des étiquettes
lblNom.Accelerator = "N"  ' Alt+N active le contrôle associé
```

**Messages d'aide :**
```vba
' Utilisez ControlTipText pour l'aide contextuelle
txtEmail.ControlTipText = "Entrez votre adresse email complète"
```

---

Maîtriser ces trois contrôles fondamentaux vous donne les outils nécessaires pour créer des interfaces utilisateur complètes et intuitives. Dans la section suivante, nous explorerons la gestion des événements de ces contrôles pour créer des interactions sophistiquées et réactives.

⏭️
