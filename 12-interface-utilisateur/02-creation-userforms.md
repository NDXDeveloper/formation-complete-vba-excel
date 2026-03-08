🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.2. Création de UserForms

## Introduction

Les UserForms (formulaires utilisateur) représentent la prochaine étape dans la création d'interfaces utilisateur en VBA. Contrairement aux simples `InputBox` et `MsgBox`, les UserForms permettent de créer des interfaces complètes et personnalisées avec de nombreux contrôles, une mise en page sophistiquée et une expérience utilisateur professionnelle.

## Qu'est-ce qu'un UserForm ?

Un UserForm est une fenêtre personnalisée que vous pouvez créer dans l'éditeur VBA. C'est l'équivalent d'une boîte de dialogue ou d'une fenêtre d'application, mais entièrement conçue selon vos besoins. Un UserForm peut contenir :

- Des zones de texte pour la saisie
- Des boutons pour les actions
- Des listes déroulantes pour les choix
- Des cases à cocher et boutons radio
- Des étiquettes pour l'information
- Des images et autres éléments graphiques

## Pourquoi utiliser des UserForms ?

### Avantages par rapport aux InputBox/MsgBox

**Flexibilité complète :**
- Disposition libre des éléments
- Taille et apparence personnalisables
- Plusieurs champs de saisie simultanés

**Meilleure expérience utilisateur :**
- Interface plus intuitive et professionnelle
- Validation en temps réel
- Navigation avec la touche Tab

**Fonctionnalités avancées :**
- Contrôles spécialisés (calendriers, barres de progression)
- Gestion d'événements sophistiquée
- Possibilité de créer des assistants multi-étapes

## Création de votre premier UserForm

### Étape 1 : Accéder à l'éditeur VBA

1. Ouvrez Excel
2. Appuyez sur `Alt + F11` pour ouvrir l'éditeur VBA
3. Dans le menu **Insertion**, cliquez sur **UserForm**

Une nouvelle fenêtre apparaît avec :
- Un formulaire vide (la zone de conception)
- La boîte à outils des contrôles
- La fenêtre des propriétés

### Étape 2 : Comprendre l'interface de conception

**La zone de conception :**
C'est votre formulaire vide où vous allez placer vos contrôles. Vous pouvez le redimensionner en tirant sur les coins.

**La boîte à outils :**
Elle contient tous les contrôles que vous pouvez ajouter :
- Label (étiquette)
- TextBox (zone de texte)
- CommandButton (bouton)
- CheckBox (case à cocher)
- OptionButton (bouton radio)
- ComboBox (liste déroulante)
- ListBox (liste)
- Et bien d'autres...

**La fenêtre des propriétés :**
Elle affiche et permet de modifier les propriétés de l'élément sélectionné (couleur, taille, nom, etc.).

### Étape 3 : Configurer les propriétés du formulaire

Sélectionnez le formulaire (cliquez sur sa barre de titre) et modifiez ces propriétés importantes :

**Propriétés essentielles :**
- **Name** : Le nom du formulaire dans le code (ex: `frmContact`)
- **Caption** : Le titre affiché dans la barre de titre
- **Width** et **Height** : Les dimensions du formulaire
- **StartUpPosition** : Position d'ouverture (recommandé : `2 - CenterScreen`)

```vba
' Exemple de propriétés pour un formulaire de contact
Name: frmContact  
Caption: Formulaire de Contact  
Width: 400  
Height: 300  
StartUpPosition: 2 - CenterScreen  
```

## Ajout de contrôles de base

### Les étiquettes (Labels)

Les étiquettes servent à afficher du texte informatif ou des instructions.

**Comment ajouter :**
1. Cliquez sur l'outil **Label** dans la boîte à outils
2. Cliquez et tirez sur le formulaire pour créer l'étiquette
3. Modifiez la propriété **Caption** pour changer le texte

**Propriétés importantes :**
- **Caption** : Le texte affiché
- **Font** : Police, taille, style
- **ForeColor** : Couleur du texte
- **BackColor** : Couleur de fond

### Les zones de texte (TextBox)

Les zones de texte permettent à l'utilisateur de saisir des informations.

**Propriétés importantes :**
- **Name** : Nom du contrôle (ex: `txtNom`, `txtEmail`)
- **Text** : Valeur par défaut
- **MaxLength** : Nombre maximum de caractères
- **PasswordChar** : Caractère de masquage pour les mots de passe
- **MultiLine** : Autorise plusieurs lignes
- **ScrollBars** : Barres de défilement

### Les boutons (CommandButton)

Les boutons déclenchent des actions quand l'utilisateur clique dessus.

**Propriétés importantes :**
- **Name** : Nom du bouton (ex: `btnOK`, `btnAnnuler`)
- **Caption** : Texte du bouton
- **Default** : Bouton activé par Entrée
- **Cancel** : Bouton activé par Échap

## Création d'un formulaire simple étape par étape

Créons ensemble un formulaire de saisie d'employé :

### Étape 1 : Préparation du formulaire

```vba
' Propriétés du formulaire
Name: frmEmploye  
Caption: Nouveau Employé  
Width: 350  
Height: 250  
StartUpPosition: 2 - CenterScreen  
```

### Étape 2 : Ajout des étiquettes

Ajoutez ces étiquettes avec leurs propriétés :

```vba
' Étiquette "Nom :"
Name: lblNom  
Caption: Nom :  
Left: 20  
Top: 20  

' Étiquette "Prénom :"
Name: lblPrenom  
Caption: Prénom :  
Left: 20  
Top: 50  

' Étiquette "Service :"
Name: lblService  
Caption: Service :  
Left: 20  
Top: 80  
```

### Étape 3 : Ajout des zones de texte

```vba
' Zone de texte pour le nom
Name: txtNom  
Left: 80  
Top: 17  
Width: 200  

' Zone de texte pour le prénom
Name: txtPrenom  
Left: 80  
Top: 47  
Width: 200  

' Zone de texte pour le service
Name: txtService  
Left: 80  
Top: 77  
Width: 200  
```

### Étape 4 : Ajout des boutons

```vba
' Bouton OK
Name: btnOK  
Caption: OK  
Left: 80  
Top: 120  
Width: 60  
Default: True  

' Bouton Annuler
Name: btnAnnuler  
Caption: Annuler  
Left: 160  
Top: 120  
Width: 60  
Cancel: True  
```

## Programmation des événements

### Événement Click des boutons

Double-cliquez sur le bouton OK pour créer automatiquement la procédure d'événement :

```vba
Private Sub btnOK_Click()
    ' Vérification que les champs sont remplis
    If txtNom.Text = "" Or txtPrenom.Text = "" Then
        MsgBox "Veuillez remplir tous les champs obligatoires.", vbExclamation
        Exit Sub
    End If

    ' Traitement des données
    MsgBox "Employé enregistré : " & txtPrenom.Text & " " & txtNom.Text

    ' Fermeture du formulaire
    Unload Me
End Sub

Private Sub btnAnnuler_Click()
    ' Fermeture sans sauvegarde
    Unload Me
End Sub
```

### Événement Initialize du formulaire

Cet événement se déclenche à l'ouverture du formulaire :

```vba
Private Sub UserForm_Initialize()
    ' Initialisation des valeurs par défaut
    txtService.Text = "Informatique"

    ' Focus sur le premier champ
    txtNom.SetFocus
End Sub
```

## Affichage du formulaire

Pour afficher votre formulaire depuis une macro :

```vba
Sub AfficherFormulaireEmploye()
    ' Affichage du formulaire
    frmEmploye.Show
End Sub
```

### Différents modes d'affichage

```vba
' Affichage modal (bloque l'accès au reste de l'application)
frmEmploye.Show vbModal

' Affichage non-modal (l'utilisateur peut accéder au reste)
frmEmploye.Show vbModeless

' Par défaut, Show utilise le mode modal
frmEmploye.Show
```

## Récupération des données du formulaire

### Méthode 1 : Variables publiques

Dans un module standard :

```vba
Public nomEmploye As String  
Public prenomEmploye As String  
Public serviceEmploye As String  

Sub CollecterDonneesEmploye()
    frmEmploye.Show

    ' Après fermeture du formulaire
    If nomEmploye <> "" Then
        MsgBox "Données reçues : " & prenomEmploye & " " & nomEmploye
    End If
End Sub
```

Dans le formulaire :

```vba
Private Sub btnOK_Click()
    If txtNom.Text = "" Or txtPrenom.Text = "" Then
        MsgBox "Veuillez remplir tous les champs obligatoires.", vbExclamation
        Exit Sub
    End If

    ' Enregistrement dans les variables globales
    nomEmploye = txtNom.Text
    prenomEmploye = txtPrenom.Text
    serviceEmploye = txtService.Text

    Me.Hide ' Masque le formulaire sans le détruire
End Sub
```

### Méthode 2 : Propriétés personnalisées

Dans le code du formulaire :

```vba
' Propriétés publiques pour récupérer les données
Public Property Get NomComplet() As String
    NomComplet = txtPrenom.Text & " " & txtNom.Text
End Property

Public Property Get Service() As String
    Service = txtService.Text
End Property
```

## Validation des données

### Validation en temps réel

```vba
Private Sub txtNom_Change()
    ' Supprimer les chiffres du nom
    Dim i As Integer
    Dim nouveauTexte As String

    For i = 1 To Len(txtNom.Text)
        If Not IsNumeric(Mid(txtNom.Text, i, 1)) Then
            nouveauTexte = nouveauTexte & Mid(txtNom.Text, i, 1)
        End If
    Next i

    If nouveauTexte <> txtNom.Text Then
        txtNom.Text = nouveauTexte
    End If
End Sub
```

### Validation à la sortie du champ

```vba
Private Sub txtNom_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(txtNom.Text) < 2 Then
        MsgBox "Le nom doit contenir au moins 2 caractères."
        Cancel = True ' Empêche de quitter le champ
    End If
End Sub
```

## Conseils de conception

### Disposition et ergonomie

**Alignement :**
- Alignez les contrôles de même type
- Utilisez une grille invisible pour la régularité
- Respectez des marges cohérentes

**Espacement :**
- Laissez suffisamment d'espace entre les contrôles
- Groupez les éléments liés
- Évitez la surcharge visuelle

**Navigation :**
- Définissez l'ordre des tabulations (propriété TabIndex)
- Utilisez des raccourcis clavier (propriété Accelerator)
- Marquez clairement les champs obligatoires

### Gestion des erreurs

```vba
Private Sub btnOK_Click()
    On Error GoTo ErreurGestion

    ' Validation des données
    If txtNom.Text = "" Then
        MsgBox "Le nom est obligatoire.", vbExclamation, "Erreur de saisie"
        txtNom.SetFocus
        Exit Sub
    End If

    ' Traitement des données
    ' ... votre code ici ...

    Exit Sub

ErreurGestion:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical
End Sub
```

### Amélioration de l'apparence

**Couleurs et polices :**
- Utilisez des couleurs cohérentes avec votre application
- Choisissez des polices lisibles
- Respectez les conventions Windows

**Taille et redimensionnement :**
- Adaptez la taille aux contenus
- Prévoyez différentes résolutions d'écran
- Testez sur différents systèmes

## Fermeture et nettoyage

### Méthodes de fermeture

```vba
' Masquer le formulaire (reste en mémoire)
Me.Hide

' Décharger le formulaire (libère la mémoire)
Unload Me

' Depuis l'extérieur du formulaire
Unload frmEmploye
```

### Événement QueryClose

```vba
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ' Utilisateur clique sur X
        Dim reponse As Integer
        reponse = MsgBox("Voulez-vous vraiment fermer sans sauvegarder ?", _
                        vbYesNo + vbQuestion, "Confirmation")
        If reponse = vbNo Then
            Cancel = True ' Annule la fermeture
        End If
    End If
End Sub
```

## Bonnes pratiques

### Nommage des contrôles

Utilisez des préfixes clairs :
- `lbl` pour les étiquettes (lblNom)
- `txt` pour les zones de texte (txtNom)
- `btn` pour les boutons (btnOK)
- `cmb` pour les listes déroulantes (cmbService)
- `chk` pour les cases à cocher (chkActif)

### Structure du code

```vba
' En-tête du formulaire avec les déclarations
Option Explicit

' Variables privées du formulaire
Private dataValid As Boolean

' Événements d'initialisation
Private Sub UserForm_Initialize()
    ' Code d'initialisation
End Sub

' Événements des contrôles
Private Sub btnOK_Click()
    ' Code du bouton OK
End Sub

' Procédures privées d'aide
Private Sub ValidateData()
    ' Code de validation
End Sub
```

### Performance et optimisation

- Désactivez le calcul automatique pendant le traitement : `Application.Calculation = xlCalculationManual`
- Utilisez `DoEvents` pour maintenir la réactivité dans les longues opérations
- Libérez les ressources avec `Unload` quand c'est possible

---

Les UserForms ouvrent un monde de possibilités pour créer des interfaces utilisateur sophistiquées. Avec ces bases solides, vous pouvez maintenant créer des formulaires professionnels qui amélioreront considérablement l'expérience de vos utilisateurs. Dans la section suivante, nous explorerons les différents types de contrôles disponibles et leurs utilisations spécifiques.

⏭️
