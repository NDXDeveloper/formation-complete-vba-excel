🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.4. Interface de saisie complète

## Introduction

Une interface de saisie complète permet aux utilisateurs de saisir des données de manière structurée et conviviale. Dans ce chapitre, nous allons créer un système complet avec un formulaire personnalisé (UserForm) qui inclut la validation des données, la gestion des erreurs et l'enregistrement dans Excel.

## Objectifs du projet

Notre interface de saisie permettra de :
- Collecter des informations sur des employés (nom, prénom, poste, salaire, date d'embauche)
- Valider automatiquement les données saisies
- Afficher des messages d'erreur clairs en cas de problème
- Enregistrer les données dans une feuille Excel
- Permettre la modification et la suppression d'enregistrements existants

## Étape 1 : Création du UserForm

### Création du formulaire

1. Dans l'éditeur VBA (Alt+F11), clic droit sur votre projet
2. Choisir **Insérer > UserForm**
3. Nommer le formulaire `frmSaisieEmploye`

### Ajout des contrôles

Ajoutez les contrôles suivants sur votre formulaire :

```vba
' Contrôles nécessaires :
' - 5 Labels pour les titres des champs
' - 4 TextBox pour nom, prénom, poste, salaire
' - 1 ComboBox pour le type de contrat
' - 1 TextBox avec format date pour la date d'embauche
' - 3 CommandButton pour Enregistrer, Modifier, Fermer
' - 1 ListBox pour afficher les employés existants
```

### Configuration des propriétés

```vba
' Propriétés du formulaire
frmSaisieEmploye.Caption = "Gestion des Employés"  
frmSaisieEmploye.Width = 400  
frmSaisieEmploye.Height = 500  

' Noms des contrôles (propriété Name)
txtNom.Name = "txtNom"  
txtPrenom.Name = "txtPrenom"  
txtPoste.Name = "txtPoste"  
txtSalaire.Name = "txtSalaire"  
txtDateEmbauche.Name = "txtDateEmbauche"  
cboContrat.Name = "cboContrat"  
cmdEnregistrer.Name = "cmdEnregistrer"  
cmdModifier.Name = "cmdModifier"  
cmdFermer.Name = "cmdFermer"  
lstEmployes.Name = "lstEmployes"  
```

## Étape 2 : Initialisation du formulaire

### Code d'initialisation

```vba
Private Sub UserForm_Initialize()
    ' Initialiser la ComboBox avec les types de contrat
    cboContrat.AddItem "CDI"
    cboContrat.AddItem "CDD"
    cboContrat.AddItem "Stage"
    cboContrat.AddItem "Freelance"

    ' Définir une valeur par défaut
    cboContrat.Value = "CDI"

    ' Charger la liste des employés existants
    ChargerListeEmployes

    ' Vider tous les champs au démarrage
    ViderChamps
End Sub

Private Sub ViderChamps()
    ' Effacer tous les champs de saisie
    txtNom.Value = ""
    txtPrenom.Value = ""
    txtPoste.Value = ""
    txtSalaire.Value = ""
    txtDateEmbauche.Value = ""
    cboContrat.Value = "CDI"

    ' Remettre le focus sur le premier champ
    txtNom.SetFocus
End Sub
```

## Étape 3 : Validation des données

### Fonction de validation générale

```vba
Private Function ValiderDonnees() As Boolean
    Dim messageErreur As String
    messageErreur = ""

    ' Vérification du nom (obligatoire, minimum 2 caractères)
    If Len(Trim(txtNom.Value)) < 2 Then
        messageErreur = messageErreur & "- Le nom doit contenir au moins 2 caractères" & vbCrLf
    End If

    ' Vérification du prénom (obligatoire, minimum 2 caractères)
    If Len(Trim(txtPrenom.Value)) < 2 Then
        messageErreur = messageErreur & "- Le prénom doit contenir au moins 2 caractères" & vbCrLf
    End If

    ' Vérification du poste (obligatoire)
    If Len(Trim(txtPoste.Value)) = 0 Then
        messageErreur = messageErreur & "- Le poste est obligatoire" & vbCrLf
    End If

    ' Vérification du salaire (doit être un nombre positif)
    If Not IsNumeric(txtSalaire.Value) Or Val(txtSalaire.Value) <= 0 Then
        messageErreur = messageErreur & "- Le salaire doit être un nombre positif" & vbCrLf
    End If

    ' Vérification de la date d'embauche
    If Not IsDate(txtDateEmbauche.Value) Then
        messageErreur = messageErreur & "- La date d'embauche doit être une date valide (jj/mm/aaaa)" & vbCrLf
    Else
        ' Vérifier que la date n'est pas dans le futur
        If CDate(txtDateEmbauche.Value) > Date Then
            messageErreur = messageErreur & "- La date d'embauche ne peut pas être dans le futur" & vbCrLf
        End If
    End If

    ' Vérification du type de contrat
    If cboContrat.Value = "" Then
        messageErreur = messageErreur & "- Le type de contrat doit être sélectionné" & vbCrLf
    End If

    ' Afficher les erreurs s'il y en a
    If messageErreur <> "" Then
        MsgBox "Erreurs de saisie détectées :" & vbCrLf & vbCrLf & messageErreur, _
               vbExclamation, "Données invalides"
        ValiderDonnees = False
    Else
        ValiderDonnees = True
    End If
End Function
```

### Validation en temps réel

```vba
' Validation du salaire pendant la saisie
Private Sub txtSalaire_Change()
    ' Autoriser seulement les chiffres, le point et la virgule
    Dim i As Integer
    Dim nouveauTexte As String

    For i = 1 To Len(txtSalaire.Value)
        If IsNumeric(Mid(txtSalaire.Value, i, 1)) Or _
           Mid(txtSalaire.Value, i, 1) = "." Or _
           Mid(txtSalaire.Value, i, 1) = "," Then
            nouveauTexte = nouveauTexte & Mid(txtSalaire.Value, i, 1)
        End If
    Next i

    ' Remplacer la virgule par un point pour la cohérence
    nouveauTexte = Replace(nouveauTexte, ",", ".")

    If nouveauTexte <> txtSalaire.Value Then
        txtSalaire.Value = nouveauTexte
    End If
End Sub

' Validation de la date pendant la saisie
Private Sub txtDateEmbauche_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtDateEmbauche.Value <> "" Then
        If Not IsDate(txtDateEmbauche.Value) Then
            MsgBox "Format de date invalide. Utilisez le format jj/mm/aaaa", vbExclamation
            Cancel = True
        End If
    End If
End Sub
```

## Étape 4 : Enregistrement des données

### Code d'enregistrement

```vba
Private Sub cmdEnregistrer_Click()
    ' Valider les données avant enregistrement
    If Not ValiderDonnees() Then
        Exit Sub
    End If

    ' Demander confirmation
    If MsgBox("Voulez-vous enregistrer ces données ?", _
              vbQuestion + vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If

    ' Procéder à l'enregistrement
    If EnregistrerEmploye() Then
        MsgBox "Employé enregistré avec succès !", vbInformation
        ViderChamps
        ChargerListeEmployes
    Else
        MsgBox "Erreur lors de l'enregistrement.", vbCritical
    End If
End Sub

Private Function EnregistrerEmploye() As Boolean
    On Error GoTo GestionErreur

    Dim ws As Worksheet
    Dim derniereLigne As Long

    ' Référencer la feuille de destination (la créer si elle n'existe pas)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Employes")
    On Error GoTo GestionErreur

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Employes"
        ' Créer les en-têtes
        CreerEnTetes ws
    End If

    ' Trouver la première ligne libre
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Enregistrer les données
    ws.Cells(derniereLigne, 1).Value = Trim(txtNom.Value)
    ws.Cells(derniereLigne, 2).Value = Trim(txtPrenom.Value)
    ws.Cells(derniereLigne, 3).Value = Trim(txtPoste.Value)
    ws.Cells(derniereLigne, 4).Value = CDbl(txtSalaire.Value)
    ws.Cells(derniereLigne, 5).Value = CDate(txtDateEmbauche.Value)
    ws.Cells(derniereLigne, 6).Value = cboContrat.Value
    ws.Cells(derniereLigne, 7).Value = Now() ' Date de création de l'enregistrement

    ' Formater les cellules
    ws.Cells(derniereLigne, 4).NumberFormat = "#,##0.00 €"
    ws.Cells(derniereLigne, 5).NumberFormat = "dd/mm/yyyy"
    ws.Cells(derniereLigne, 7).NumberFormat = "dd/mm/yyyy hh:mm"

    EnregistrerEmploye = True
    Exit Function

GestionErreur:
    MsgBox "Erreur lors de l'enregistrement : " & Err.Description, vbCritical
    EnregistrerEmploye = False
End Function

Private Sub CreerEnTetes(ws As Worksheet)
    ' Créer les en-têtes de colonnes
    ws.Cells(1, 1).Value = "Nom"
    ws.Cells(1, 2).Value = "Prénom"
    ws.Cells(1, 3).Value = "Poste"
    ws.Cells(1, 4).Value = "Salaire"
    ws.Cells(1, 5).Value = "Date d'embauche"
    ws.Cells(1, 6).Value = "Type de contrat"
    ws.Cells(1, 7).Value = "Date de création"

    ' Formater les en-têtes
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With

    ' Ajuster la largeur des colonnes
    ws.Columns("A:G").AutoFit
End Sub
```

## Étape 5 : Affichage et sélection des données

### Chargement de la liste

```vba
Private Sub ChargerListeEmployes()
    On Error GoTo GestionErreur

    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim i As Long

    ' Vider la liste actuelle
    lstEmployes.Clear

    ' Vérifier si la feuille existe
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Employes")
    On Error GoTo GestionErreur

    If ws Is Nothing Then Exit Sub

    ' Trouver la dernière ligne avec des données
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Charger les données dans la ListBox (à partir de la ligne 2)
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value <> "" Then
            lstEmployes.AddItem ws.Cells(i, 1).Value & " " & _
                              ws.Cells(i, 2).Value & " - " & _
                              ws.Cells(i, 3).Value & " (" & _
                              ws.Cells(i, 6).Value & ")"
        End If
    Next i

    Exit Sub

GestionErreur:
    ' En cas d'erreur, ne rien faire (la liste restera vide)
End Sub
```

### Sélection dans la liste

```vba
Private Sub lstEmployes_Click()
    On Error GoTo GestionErreur

    Dim ws As Worksheet
    Dim ligneSelectionnee As Long

    ' Vérifier qu'un élément est sélectionné
    If lstEmployes.ListIndex = -1 Then Exit Sub

    ' Référencer la feuille
    Set ws = ThisWorkbook.Worksheets("Employes")

    ' Calculer la ligne correspondante (ListIndex + 2 car on commence à la ligne 2)
    ligneSelectionnee = lstEmployes.ListIndex + 2

    ' Charger les données dans le formulaire
    txtNom.Value = ws.Cells(ligneSelectionnee, 1).Value
    txtPrenom.Value = ws.Cells(ligneSelectionnee, 2).Value
    txtPoste.Value = ws.Cells(ligneSelectionnee, 3).Value
    txtSalaire.Value = ws.Cells(ligneSelectionnee, 4).Value
    txtDateEmbauche.Value = Format(ws.Cells(ligneSelectionnee, 5).Value, "dd/mm/yyyy")
    cboContrat.Value = ws.Cells(ligneSelectionnee, 6).Value

    ' Activer le bouton Modifier
    cmdModifier.Enabled = True

    Exit Sub

GestionErreur:
    MsgBox "Erreur lors du chargement des données : " & Err.Description, vbCritical
End Sub
```

## Étape 6 : Modification des données

### Code de modification

```vba
Private Sub cmdModifier_Click()
    ' Vérifier qu'un employé est sélectionné
    If lstEmployes.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un employé à modifier.", vbInformation
        Exit Sub
    End If

    ' Valider les nouvelles données
    If Not ValiderDonnees() Then
        Exit Sub
    End If

    ' Demander confirmation
    If MsgBox("Voulez-vous modifier cet employé ?", _
              vbQuestion + vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If

    ' Procéder à la modification
    If ModifierEmploye() Then
        MsgBox "Employé modifié avec succès !", vbInformation
        ViderChamps
        ChargerListeEmployes
        cmdModifier.Enabled = False
    Else
        MsgBox "Erreur lors de la modification.", vbCritical
    End If
End Sub

Private Function ModifierEmploye() As Boolean
    On Error GoTo GestionErreur

    Dim ws As Worksheet
    Dim ligneAModifier As Long

    ' Référencer la feuille
    Set ws = ThisWorkbook.Worksheets("Employes")

    ' Calculer la ligne à modifier
    ligneAModifier = lstEmployes.ListIndex + 2

    ' Modifier les données
    ws.Cells(ligneAModifier, 1).Value = Trim(txtNom.Value)
    ws.Cells(ligneAModifier, 2).Value = Trim(txtPrenom.Value)
    ws.Cells(ligneAModifier, 3).Value = Trim(txtPoste.Value)
    ws.Cells(ligneAModifier, 4).Value = CDbl(txtSalaire.Value)
    ws.Cells(ligneAModifier, 5).Value = CDate(txtDateEmbauche.Value)
    ws.Cells(ligneAModifier, 6).Value = cboContrat.Value
    ' Note : on ne modifie pas la date de création (colonne 7)

    ModifierEmploye = True
    Exit Function

GestionErreur:
    MsgBox "Erreur lors de la modification : " & Err.Description, vbCritical
    ModifierEmploye = False
End Function
```

## Étape 7 : Finalisation de l'interface

### Bouton de fermeture

```vba
Private Sub cmdFermer_Click()
    ' Vérifier s'il y a des modifications non sauvegardées
    If txtNom.Value <> "" Or txtPrenom.Value <> "" Or _
       txtPoste.Value <> "" Or txtSalaire.Value <> "" Or _
       txtDateEmbauche.Value <> "" Then

        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Des données ont été saisies. Voulez-vous fermer sans enregistrer ?", _
                        vbQuestion + vbYesNo, "Fermeture")

        If reponse = vbNo Then
            Exit Sub
        End If
    End If

    ' Fermer le formulaire
    Unload Me
End Sub
```

### Gestion de la fermeture du formulaire

```vba
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Empêcher la fermeture par la croix si des données sont saisies
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        cmdFermer_Click
    End If
End Sub
```

## Étape 8 : Lancement de l'interface

### Macro de lancement

Créez une macro dans un module standard pour lancer l'interface :

```vba
Sub LancerInterfaceSaisie()
    ' Lancer le formulaire de saisie des employés
    frmSaisieEmploye.Show
End Sub
```

## Améliorations possibles

Cette interface de base peut être améliorée avec :

### Fonctionnalités supplémentaires
- **Recherche** : Ajouter une zone de recherche pour filtrer les employés
- **Suppression** : Ajouter la possibilité de supprimer des enregistrements
- **Export** : Permettre d'exporter les données vers d'autres formats
- **Tri** : Trier la liste des employés par différents critères

### Améliorations techniques
- **Gestion des doublons** : Vérifier qu'un employé n'existe pas déjà
- **Sauvegarde automatique** : Sauvegarder périodiquement les modifications
- **Historique** : Conserver un historique des modifications
- **Permissions** : Gérer les droits d'accès selon les utilisateurs

### Interface utilisateur
- **Thème visuel** : Personnaliser l'apparence du formulaire
- **Raccourcis clavier** : Ajouter des raccourcis pour les actions courantes
- **Aide contextuelle** : Ajouter des bulles d'aide sur les champs
- **Progression** : Afficher une barre de progression lors des opérations longues

## Conclusion

Cette interface de saisie complète démontre les principes essentiels de création d'applications VBA professionnelles : validation des données, gestion des erreurs, interface utilisateur intuitive et manipulation sécurisée des données Excel. Elle constitue une base solide pour développer des applications plus complexes selon les besoins spécifiques de votre organisation.

⏭️ [Macro de nettoyage et formatage](/22-projets-pratiques/05-macro-nettoyage-formatage.md)
