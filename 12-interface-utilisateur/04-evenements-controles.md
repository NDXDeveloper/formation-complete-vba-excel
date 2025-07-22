🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.4. Événements des contrôles

## Introduction

Les événements sont le cœur de l'interactivité dans vos UserForms. Un événement est une action qui se produit suite à une interaction de l'utilisateur ou à un changement d'état du système. Comprendre et maîtriser les événements vous permettra de créer des interfaces réactives et intelligentes qui répondent de manière appropriée aux actions de l'utilisateur.

## Qu'est-ce qu'un événement ?

### Définition

Un événement est un signal envoyé par un contrôle ou le système quand quelque chose se produit. Par exemple :
- L'utilisateur clique sur un bouton → **événement Click**
- L'utilisateur modifie le contenu d'une zone de texte → **événement Change**
- L'utilisateur ouvre un formulaire → **événement Initialize**

### Comment fonctionnent les événements ?

Quand un événement se produit, VBA recherche automatiquement une procédure correspondante dans votre code. Si cette procédure existe, elle s'exécute automatiquement.

```vba
' Structure d'une procédure d'événement
Private Sub NomDuControle_NomEvenement([Paramètres])
    ' Votre code ici
End Sub
```

### Création automatique des procédures d'événement

**Méthode 1 : Double-clic sur le contrôle**
- Double-cliquez sur le contrôle dans le formulaire
- VBA crée automatiquement la procédure pour l'événement par défaut

**Méthode 2 : Via la fenêtre de code**
1. Sélectionnez le contrôle dans la liste déroulante de gauche
2. Choisissez l'événement dans la liste déroulante de droite
3. VBA crée automatiquement la procédure vide

## Événements des UserForms

### Initialize - Initialisation du formulaire

Se déclenche avant l'affichage du formulaire. C'est l'endroit idéal pour configurer vos contrôles.

```vba
Private Sub UserForm_Initialize()
    ' Configuration initiale du formulaire
    Me.Caption = "Formulaire de saisie - Version 1.0"

    ' Initialisation des contrôles
    txtDate.Text = Format(Date, "dd/mm/yyyy")
    cmbPays.AddItem "France"
    cmbPays.AddItem "Espagne"
    cmbPays.ListIndex = 0

    ' Définir le focus initial
    txtNom.SetFocus
End Sub
```

### Activate - Activation du formulaire

Se déclenche chaque fois que le formulaire devient actif.

```vba
Private Sub UserForm_Activate()
    ' Mise à jour de l'affichage
    lblDateOuverture.Caption = "Ouvert le : " & Format(Now, "dd/mm/yyyy hh:mm")

    ' Vérification de données externes
    If Range("A1").Value = "" Then
        MsgBox "Attention : données manquantes dans la feuille"
    End If
End Sub
```

### QueryClose - Tentative de fermeture

Se déclenche avant la fermeture du formulaire. Permet d'annuler la fermeture ou de demander confirmation.

```vba
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' CloseMode indique comment la fermeture a été déclenchée
    ' 0 = vbFormControlMenu (bouton X)
    ' 1 = vbFormCode (Unload dans le code)

    If CloseMode = vbFormControlMenu Then
        Dim reponse As Integer
        reponse = MsgBox("Voulez-vous vraiment fermer sans sauvegarder ?", _
                        vbYesNo + vbQuestion, "Confirmation")

        If reponse = vbNo Then
            Cancel = True  ' Annule la fermeture
        End If
    End If
End Sub
```

### Terminate - Destruction du formulaire

Se déclenche quand le formulaire est déchargé de la mémoire.

```vba
Private Sub UserForm_Terminate()
    ' Nettoyage final
    MsgBox "Formulaire fermé à " & Format(Now, "hh:mm:ss")

    ' Libération des ressources si nécessaire
    Set objConnexion = Nothing
End Sub
```

## Événements des contrôles TextBox

### Change - Modification du contenu

Se déclenche à chaque modification du texte, même par le code.

```vba
Private Sub txtNom_Change()
    ' Conversion automatique en majuscules
    If Len(txtNom.Text) > 0 Then
        ' Éviter la boucle infinie avec un flag
        Static enCours As Boolean
        If Not enCours Then
            enCours = True
            txtNom.Text = UCase(txtNom.Text)
            enCours = False
        End If
    End If

    ' Mise à jour d'un autre contrôle
    lblCaracteres.Caption = Len(txtNom.Text) & " caractères"
End Sub
```

### Enter - Entrée dans le contrôle

Se déclenche quand le contrôle obtient le focus.

```vba
Private Sub txtMontant_Enter()
    ' Sélectionner tout le contenu pour faciliter la modification
    txtMontant.SelStart = 0
    txtMontant.SelLength = Len(txtMontant.Text)

    ' Changer la couleur de fond pour indiquer le focus
    txtMontant.BackColor = RGB(255, 255, 200)  ' Jaune clair

    ' Afficher une aide contextuelle
    lblAide.Caption = "Entrez un montant en euros (ex: 125.50)"
End Sub
```

### Exit - Sortie du contrôle

Se déclenche quand le contrôle perd le focus. Permet de valider les données.

```vba
Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Restaurer la couleur normale
    txtEmail.BackColor = RGB(255, 255, 255)  ' Blanc

    ' Validation de l'email
    If txtEmail.Text <> "" Then
        If InStr(txtEmail.Text, "@") = 0 Or InStr(txtEmail.Text, ".") = 0 Then
            MsgBox "Format d'email invalide", vbExclamation
            txtEmail.BackColor = RGB(255, 200, 200)  ' Rouge clair
            Cancel = True  ' Empêche de quitter le contrôle
        End If
    End If
End Sub
```

### KeyPress - Frappe d'une touche

Se déclenche pour chaque touche pressée. Permet de filtrer la saisie.

```vba
Private Sub txtAge_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Autoriser seulement les chiffres et les touches de contrôle
    Select Case KeyAscii
        Case 48 To 57      ' Chiffres 0-9
            ' Autorisé
        Case 8, 127        ' Backspace et Delete
            ' Autorisé
        Case 13            ' Entrée
            ' Passer au contrôle suivant
            SendKeys "{TAB}"
            KeyAscii = 0
        Case Else
            ' Refuser la touche
            KeyAscii = 0
            Beep  ' Signal sonore
    End Select
End Sub
```

### KeyDown et KeyUp - Touches spéciales

Permettent de détecter les touches spéciales (F1, Ctrl, Alt, etc.).

```vba
Private Sub txtRecherche_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                ByVal Shift As Integer)
    ' Détecter Ctrl+A pour sélectionner tout
    If KeyCode = 65 And Shift = 2 Then  ' A + Ctrl
        txtRecherche.SelStart = 0
        txtRecherche.SelLength = Len(txtRecherche.Text)
    End If

    ' Détecter F1 pour l'aide
    If KeyCode = 112 Then  ' F1
        MsgBox "Tapez votre recherche et appuyez sur Entrée", vbInformation, "Aide"
    End If
End Sub
```

## Événements des contrôles ComboBox et ListBox

### Click - Sélection d'un élément

Se déclenche quand l'utilisateur sélectionne un élément.

```vba
Private Sub cmbPays_Click()
    ' Mise à jour automatique d'autres contrôles
    Select Case cmbPays.Text
        Case "France"
            txtDevise.Text = "EUR"
            txtIndicatif.Text = "+33"
        Case "Espagne"
            txtDevise.Text = "EUR"
            txtIndicatif.Text = "+34"
        Case "Royaume-Uni"
            txtDevise.Text = "GBP"
            txtIndicatif.Text = "+44"
        Case Else
            txtDevise.Text = ""
            txtIndicatif.Text = ""
    End Select
End Sub
```

### Change - Modification de la sélection

Se déclenche lors d'un changement de sélection ou de texte.

```vba
Private Sub cmbCategorie_Change()
    ' Filtrer une autre liste en fonction de la sélection
    RemplirSousCategories(cmbCategorie.Text)

    ' Réinitialiser les contrôles dépendants
    cmbSousCategorie.ListIndex = -1
    txtDescription.Text = ""
End Sub

Private Sub RemplirSousCategories(ByVal categorie As String)
    cmbSousCategorie.Clear

    Select Case categorie
        Case "Informatique"
            cmbSousCategorie.AddItem "Ordinateurs"
            cmbSousCategorie.AddItem "Périphériques"
            cmbSousCategorie.AddItem "Logiciels"
        Case "Mobilier"
            cmbSousCategorie.AddItem "Bureaux"
            cmbSousCategorie.AddItem "Chaises"
            cmbSousCategorie.AddItem "Rangements"
    End Select
End Sub
```

### DblClick - Double-clic

Permet d'effectuer une action spéciale sur double-clic.

```vba
Private Sub lstClients_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Ouvrir la fiche détaillée du client sélectionné
    If lstClients.ListIndex >= 0 Then
        Dim nomClient As String
        nomClient = lstClients.Text

        ' Ouvrir un autre formulaire ou afficher des détails
        MsgBox "Ouverture de la fiche de : " & nomClient, vbInformation

        ' Ou charger les données dans le formulaire actuel
        ChargerDonneesClient(nomClient)
    End If
End Sub
```

## Événements des boutons (CommandButton)

### Click - Clic principal

L'événement le plus utilisé pour les boutons.

```vba
Private Sub btnValider_Click()
    ' Validation des données
    If Not ValiderFormulaire() Then
        Exit Sub
    End If

    ' Traitement des données
    On Error GoTo ErreurTraitement

    SauvegarderDonnees
    MsgBox "Données sauvegardées avec succès", vbInformation

    ' Fermer le formulaire
    Unload Me

    Exit Sub

ErreurTraitement:
    MsgBox "Erreur lors de la sauvegarde : " & Err.Description, vbCritical
End Sub
```

### MouseDown et MouseUp - Événements de souris détaillés

Permettent de créer des effets visuels personnalisés.

```vba
Private Sub btnSpecial_MouseDown(ByVal Button As Integer, _
                                ByVal Shift As Integer, _
                                ByVal X As Single, ByVal Y As Single)
    ' Effet visuel lors du clic
    btnSpecial.BackColor = RGB(200, 200, 200)  ' Gris foncé
End Sub

Private Sub btnSpecial_MouseUp(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Single, ByVal Y As Single)
    ' Restaurer l'apparence normale
    btnSpecial.BackColor = RGB(240, 240, 240)  ' Gris clair
End Sub
```

## Gestion avancée des événements

### Désactivation temporaire des événements

Parfois, vous devez modifier des contrôles sans déclencher leurs événements.

```vba
Private Sub RemplirFormulaire()
    ' Désactiver les événements temporairement
    Application.EnableEvents = False

    ' Modifications sans déclencher les événements
    txtNom.Text = "Dupont"
    txtPrenom.Text = "Jean"
    cmbPays.ListIndex = 0

    ' Réactiver les événements
    Application.EnableEvents = True

    ' Déclencher manuellement un événement si nécessaire
    Call cmbPays_Change
End Sub
```

### Variables de contrôle des événements

Utiliser des variables pour éviter les boucles infinies.

```vba
Private miseAJourEnCours As Boolean

Private Sub txtPrixHT_Change()
    If miseAJourEnCours Then Exit Sub

    miseAJourEnCours = True

    ' Calculer automatiquement le prix TTC
    If IsNumeric(txtPrixHT.Text) And txtPrixHT.Text <> "" Then
        Dim prixHT As Double
        Dim prixTTC As Double

        prixHT = CDbl(txtPrixHT.Text)
        prixTTC = prixHT * 1.2  ' TVA 20%

        txtPrixTTC.Text = Format(prixTTC, "0.00")
    End If

    miseAJourEnCours = False
End Sub

Private Sub txtPrixTTC_Change()
    If miseAJourEnCours Then Exit Sub

    miseAJourEnCours = True

    ' Calculer automatiquement le prix HT
    If IsNumeric(txtPrixTTC.Text) And txtPrixTTC.Text <> "" Then
        Dim prixTTC As Double
        Dim prixHT As Double

        prixTTC = CDbl(txtPrixTTC.Text)
        prixHT = prixTTC / 1.2  ' Retirer la TVA

        txtPrixHT.Text = Format(prixHT, "0.00")
    End If

    miseAJourEnCours = False
End Sub
```

### Gestion centralisée des événements

Pour plusieurs contrôles similaires, créez une fonction commune.

```vba
Private Sub txtChamp1_Enter()
    GererEntreeChamp txtChamp1
End Sub

Private Sub txtChamp2_Enter()
    GererEntreeChamp txtChamp2
End Sub

Private Sub txtChamp3_Enter()
    GererEntreeChamp txtChamp3
End Sub

Private Sub GererEntreeChamp(ByRef champ As MSForms.TextBox)
    ' Sélectionner tout le contenu
    champ.SelStart = 0
    champ.SelLength = Len(champ.Text)

    ' Changer la couleur de fond
    champ.BackColor = RGB(255, 255, 200)

    ' Mettre à jour l'aide
    lblAide.Caption = "Champ actif : " & champ.Name
End Sub
```

## Événements en cascade et synchronisation

### Synchronisation de contrôles liés

```vba
Private Sub lstCategories_Click()
    ' Synchroniser avec une autre liste
    SynchroniserProduits
End Sub

Private Sub SynchroniserProduits()
    If lstCategories.ListIndex >= 0 Then
        Dim categorieSelectionne As String
        categorieSelectionne = lstCategories.Text

        ' Remplir la liste des produits selon la catégorie
        lstProduits.Clear

        ' Exemple de données
        Select Case categorieSelectionne
            Case "Informatique"
                lstProduits.AddItem "Ordinateur portable"
                lstProduits.AddItem "Souris"
                lstProduits.AddItem "Clavier"
            Case "Mobilier"
                lstProduits.AddItem "Bureau"
                lstProduits.AddItem "Chaise"
                lstProduits.AddItem "Armoire"
        End Select

        ' Réinitialiser la sélection des produits
        lstProduits.ListIndex = -1
    End If
End Sub
```

### Validation en cascade

```vba
Private Sub txtCodePostal_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtCodePostal.Text <> "" Then
        ' Validation du format
        If Not EstCodePostalValide(txtCodePostal.Text) Then
            MsgBox "Code postal invalide", vbExclamation
            Cancel = True
            Exit Sub
        End If

        ' Remplissage automatique de la ville
        txtVille.Text = ObtenirVille(txtCodePostal.Text)
    End If
End Sub

Private Function EstCodePostalValide(ByVal codePostal As String) As Boolean
    ' Validation simple : 5 chiffres
    EstCodePostalValide = (Len(codePostal) = 5 And IsNumeric(codePostal))
End Function

Private Function ObtenirVille(ByVal codePostal As String) As String
    ' Exemple simplifié
    Select Case Left(codePostal, 2)
        Case "75": ObtenirVille = "Paris"
        Case "69": ObtenirVille = "Lyon"
        Case "13": ObtenirVille = "Marseille"
        Case Else: ObtenirVille = ""
    End Select
End Function
```

## Événements pour l'amélioration de l'expérience utilisateur

### Feedback visuel

```vba
Private Sub txtMontant_Enter()
    ' Highlight du champ actif
    txtMontant.BackColor = RGB(255, 255, 200)
    txtMontant.BorderColor = RGB(0, 120, 215)
End Sub

Private Sub txtMontant_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Retour à l'apparence normale
    txtMontant.BackColor = RGB(255, 255, 255)
    txtMontant.BorderColor = RGB(128, 128, 128)
End Sub
```

### Messages d'aide contextuelle

```vba
Private Sub txtEmail_Enter()
    lblAide.Caption = "Format requis : nom@domaine.com"
    lblAide.ForeColor = RGB(0, 120, 215)
End Sub

Private Sub txtTelephone_Enter()
    lblAide.Caption = "Format : 01 23 45 67 89"
    lblAide.ForeColor = RGB(0, 120, 215)
End Sub
```

### Progression et état

```vba
Private Sub btnTraiter_Click()
    ' Désactiver le bouton pendant le traitement
    btnTraiter.Enabled = False
    btnTraiter.Caption = "Traitement en cours..."

    ' Afficher une barre de progression (simulation)
    Dim i As Integer
    For i = 1 To 100
        lblProgression.Caption = "Progression : " & i & "%"
        DoEvents  ' Maintenir la réactivité

        ' Simulation d'un traitement
        Application.Wait Now + TimeValue("0:00:01")
    Next i

    ' Restaurer l'état normal
    btnTraiter.Enabled = True
    btnTraiter.Caption = "Traiter"
    lblProgression.Caption = "Traitement terminé"
End Sub
```

## Bonnes pratiques pour les événements

### Gestion d'erreurs dans les événements

```vba
Private Sub txtMontant_Change()
    On Error GoTo ErreurGestion

    ' Votre code ici
    If IsNumeric(txtMontant.Text) Then
        CalculerTotal
    End If

    Exit Sub

ErreurGestion:
    ' Ne pas afficher d'erreur pour les événements Change
    ' Juste enregistrer dans un log si nécessaire
    Debug.Print "Erreur dans txtMontant_Change : " & Err.Description
End Sub
```

### Performance des événements

```vba
Private Sub lstGrandeListe_Click()
    ' Éviter les traitements lourds dans les événements fréquents
    ' Utiliser un timer ou différer le traitement

    ' Annuler le timer précédent s'il existe
    Application.OnTime TimerProchainAppel, "TraiterSelectionDifferee", , False

    ' Programmer le traitement dans 0.5 seconde
    TimerProchainAppel = Now + TimeValue("0:00:00.5")
    Application.OnTime TimerProchainAppel, "TraiterSelectionDifferee"
End Sub
```

### Documentation des événements

```vba
Private Sub btnCalculer_Click()
    '=================================================================
    ' Procédure : btnCalculer_Click
    ' Description : Calcule les résultats financiers
    ' Prérequis : Les champs montant et taux doivent être remplis
    ' Résultat : Met à jour les champs de résultat
    '=================================================================

    ' Validation des prérequis
    If Not ValiderDonneesCalcul() Then Exit Sub

    ' Calcul et affichage des résultats
    EffectuerCalculs
End Sub
```

---

La maîtrise des événements est essentielle pour créer des interfaces utilisateur réactives et intuitives. En comprenant quand et comment utiliser chaque type d'événement, vous pouvez créer des formulaires qui anticipent les besoins de l'utilisateur et offrent une expérience fluide et professionnelle. Dans la section suivante, nous explorerons les techniques de validation des données pour garantir la qualité des informations saisies.

⏭️
