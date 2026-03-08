🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.2. Système de gestion de stock simple

## Vue d'ensemble du projet

### Contexte et problématique
La gestion des stocks est un défi majeur pour toute entreprise, qu'elle soit petite ou grande. Que vous gériez l'inventaire d'un magasin, d'un entrepôt, ou même de votre bureau, vous devez constamment suivre :

- **Quels produits** vous avez en stock
- **Combien** d'unités sont disponibles
- **Où** se trouvent les produits
- **Quand** réapprovisionner
- **Qui** a retiré ou ajouté des articles

### Problèmes courants avec la gestion manuelle
Traditionnellement, cette gestion se fait souvent avec :
- **Fichiers Excel séparés** : Difficiles à synchroniser
- **Cahiers papier** : Risque de perte, erreurs de lecture
- **Systèmes complexes** : Coûteux et difficiles à maîtriser

Ces méthodes entraînent des problèmes fréquents :
- **Ruptures de stock** : Manquer de produits importants
- **Surstockage** : Immobiliser inutilement de l'argent
- **Erreurs de comptage** : Différences entre stock théorique et réel
- **Perte de temps** : Recherches manuelles fastidieuses

### Solution proposée
Notre système de gestion de stock simple va résoudre ces problèmes en offrant :
- **Interface utilisateur intuitive** : Formulaires faciles à utiliser
- **Base de données centralisée** : Toutes les informations au même endroit
- **Mises à jour automatiques** : Calculs en temps réel
- **Alertes intelligentes** : Notifications de stock faible
- **Historique des mouvements** : Traçabilité complète

### Objectifs du projet
À la fin de ce projet, vous disposerez d'un système capable de :
1. **Enregistrer** de nouveaux produits avec leurs caractéristiques
2. **Suivre** les entrées et sorties de stock en temps réel
3. **Alerter** automatiquement en cas de stock faible
4. **Rechercher** rapidement un produit spécifique
5. **Générer** des rapports de stock automatiques
6. **Maintenir** un historique complet des mouvements

## Analyse des besoins

### Fonctionnalités principales

#### 1. Gestion des produits
- **Création** : Ajouter de nouveaux produits au catalogue
- **Modification** : Mettre à jour les informations produit
- **Suppression** : Retirer des produits obsolètes
- **Recherche** : Trouver rapidement un produit

#### 2. Gestion des mouvements de stock
- **Entrées** : Réceptions, livraisons, retours
- **Sorties** : Ventes, prélèvements, pertes
- **Ajustements** : Corrections d'inventaire
- **Transferts** : Mouvements entre emplacements

#### 3. Alertes et contrôles
- **Seuils d'alerte** : Stock minimum par produit
- **Notifications visuelles** : Couleurs pour identifier les urgences
- **Rapports automatiques** : États de stock périodiques

#### 4. Interface utilisateur
- **Formulaires intuitifs** : Saisie guidée des données
- **Boutons d'action** : Navigation simple
- **Tableaux de bord** : Vue d'ensemble du stock

### Structure des données

Notre système utilisera plusieurs tables de données :

#### Table PRODUITS
```
Code Produit | Désignation  | Catégorie | Stock Actuel | Stock Min | Prix | Emplacement  
P001         | Ordinateur   | IT        | 15          | 5         | 800  | A1-B2  
P002         | Souris       | IT        | 50          | 10        | 25   | A1-B3  
P003         | Clavier      | IT        | 30          | 8         | 45   | A1-B3  
```

#### Table MOUVEMENTS
```
Date       | Code Produit | Type      | Quantité | Utilisateur | Commentaire
15/01/2024 | P001        | ENTREE    | 10       | Martin      | Livraison
16/01/2024 | P001        | SORTIE    | 3        | Durand      | Vente client
17/01/2024 | P002        | ENTREE    | 20       | Martin      | Commande
```

## Conception de la solution

### Architecture du système

```
Système de Gestion de Stock
├── Interface Principale
│   ├── Tableau de bord
│   ├── Navigation entre modules
│   └── Alertes visuelles
├── Module Produits
│   ├── Formulaire de création/modification
│   ├── Liste des produits
│   └── Recherche avancée
├── Module Mouvements
│   ├── Saisie des entrées/sorties
│   ├── Historique des mouvements
│   └── Validation automatique
├── Module Alertes
│   ├── Détection de stock faible
│   ├── Notifications colorées
│   └── Rapports d'alerte
└── Module Rapports
    ├── État des stocks
    ├── Mouvements par période
    └── Analyse des tendances
```

### Structure des feuilles Excel

Notre système utilisera plusieurs feuilles de calcul :
- **ACCUEIL** : Tableau de bord principal
- **PRODUITS** : Base de données des produits
- **MOUVEMENTS** : Historique de tous les mouvements
- **ALERTES** : Suivi des stocks faibles
- **RAPPORTS** : États et analyses

## Développement de la solution

### 1. Création de la structure des données

Commençons par créer les feuilles et initialiser les en-têtes :

```vba
Sub InitializeStockSystem()
    '=========================================
    ' Initialisation du système de gestion de stock
    ' À exécuter une seule fois au début
    '=========================================

    ' Déclaration des variables
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' Désactiver les alertes et l'affichage
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Supprimer les feuilles existantes si elles existent
    On Error Resume Next
    wb.Sheets("ACCUEIL").Delete
    wb.Sheets("PRODUITS").Delete
    wb.Sheets("MOUVEMENTS").Delete
    wb.Sheets("ALERTES").Delete
    wb.Sheets("RAPPORTS").Delete
    On Error GoTo 0

    ' Créer la feuille PRODUITS
    Set ws = wb.Sheets.Add
    ws.Name = "PRODUITS"
    Call SetupProductSheet(ws)

    ' Créer la feuille MOUVEMENTS
    Set ws = wb.Sheets.Add
    ws.Name = "MOUVEMENTS"
    Call SetupMovementSheet(ws)

    ' Créer la feuille ALERTES
    Set ws = wb.Sheets.Add
    ws.Name = "ALERTES"
    Call SetupAlertSheet(ws)

    ' Créer la feuille RAPPORTS
    Set ws = wb.Sheets.Add
    ws.Name = "RAPPORTS"
    Call SetupReportSheet(ws)

    ' Créer la feuille ACCUEIL (tableau de bord)
    Set ws = wb.Sheets.Add
    ws.Name = "ACCUEIL"
    Call SetupDashboard(ws)

    ' Positionner sur la feuille d'accueil
    wb.Sheets("ACCUEIL").Activate

    ' Réactiver les alertes et l'affichage
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Système de gestion de stock initialisé avec succès !", vbInformation
End Sub

Sub SetupProductSheet(ws As Worksheet)
    '=========================================
    ' Configuration de la feuille PRODUITS
    '=========================================

    With ws
        ' En-têtes des colonnes
        .Cells(1, 1).Value = "Code Produit"
        .Cells(1, 2).Value = "Désignation"
        .Cells(1, 3).Value = "Catégorie"
        .Cells(1, 4).Value = "Stock Actuel"
        .Cells(1, 5).Value = "Stock Minimum"
        .Cells(1, 6).Value = "Prix Unitaire"
        .Cells(1, 7).Value = "Emplacement"
        .Cells(1, 8).Value = "Date Création"
        .Cells(1, 9).Value = "Statut"

        ' Formatage des en-têtes
        With .Range("A1:I1")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(68, 114, 196)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With

        ' Ajustement automatique des colonnes
        .Columns.AutoFit

        ' Protection de la feuille (sauf zone de données)
        .Protect Password:="stock123", AllowInsertingRows:=True
    End With
End Sub

Sub SetupMovementSheet(ws As Worksheet)
    '=========================================
    ' Configuration de la feuille MOUVEMENTS
    '=========================================

    With ws
        ' En-têtes des colonnes
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "Heure"
        .Cells(1, 3).Value = "Code Produit"
        .Cells(1, 4).Value = "Désignation"
        .Cells(1, 5).Value = "Type Mouvement"
        .Cells(1, 6).Value = "Quantité"
        .Cells(1, 7).Value = "Stock Avant"
        .Cells(1, 8).Value = "Stock Après"
        .Cells(1, 9).Value = "Utilisateur"
        .Cells(1, 10).Value = "Commentaire"

        ' Formatage des en-têtes
        With .Range("A1:J1")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(68, 114, 196)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With

        ' Format des colonnes
        .Columns("A").NumberFormat = "dd/mm/yyyy"
        .Columns("B").NumberFormat = "hh:mm"
        .Columns("F:H").NumberFormat = "0"

        .Columns.AutoFit
    End With
End Sub
```

### 2. Interface principale - Tableau de bord

```vba
Sub SetupDashboard(ws As Worksheet)
    '=========================================
    ' Création du tableau de bord principal
    '=========================================

    With ws
        ' Titre principal
        .Cells(1, 1).Value = "SYSTÈME DE GESTION DE STOCK"
        .Cells(1, 1).Font.Size = 20
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(68, 114, 196)
        .Range("A1:H1").Merge
        .Range("A1").HorizontalAlignment = xlCenter

        ' Date et heure actuelles
        .Cells(3, 1).Value = "Dernière mise à jour :"
        .Cells(3, 2).Value = Now()
        .Cells(3, 2).NumberFormat = "dd/mm/yyyy hh:mm"
        .Cells(3, 1).Font.Bold = True

        ' Section Statistiques rapides
        .Cells(5, 1).Value = "STATISTIQUES RAPIDES"
        .Cells(5, 1).Font.Size = 14
        .Cells(5, 1).Font.Bold = True

        ' Zones pour les statistiques (seront mises à jour automatiquement)
        .Cells(7, 1).Value = "Nombre total de produits :"
        .Cells(7, 3).Value = "=COUNTA(PRODUITS!A:A)-1"  ' -1 pour exclure l'en-tête

        .Cells(8, 1).Value = "Produits en alerte :"
        .Cells(8, 3).Value = "=SUMPRODUCT((PRODUITS!D:D<PRODUITS!E:E)*(PRODUITS!D:D<>\"\"))"

        .Cells(9, 1).Value = "Valeur totale du stock :"
        .Cells(9, 3).Value = "=SUMPRODUCT(PRODUITS!D:D,PRODUITS!F:F)"
        .Cells(9, 3).NumberFormat = "#,##0.00 €"

        ' Formatage des statistiques
        .Range("A7:A9").Font.Bold = True
        .Range("C7:C9").Font.Bold = True
        .Range("C7:C9").Interior.Color = RGB(242, 242, 242)

        ' Section Navigation
        .Cells(12, 1).Value = "NAVIGATION"
        .Cells(12, 1).Font.Size = 14
        .Cells(12, 1).Font.Bold = True

        ' Création des boutons de navigation
        Call CreateNavigationButtons(ws)

        ' Section Alertes
        .Cells(18, 1).Value = "ALERTES STOCK FAIBLE"
        .Cells(18, 1).Font.Size = 14
        .Cells(18, 1).Font.Bold = True
        .Cells(18, 1).Font.Color = RGB(255, 0, 0)

        ' Zone pour afficher les produits en alerte
        Call UpdateStockAlerts(ws)

        .Columns.AutoFit
    End With
End Sub

Sub CreateNavigationButtons(ws As Worksheet)
    '=========================================
    ' Création des boutons de navigation
    '=========================================

    Dim btn As Button

    ' Bouton Gestion des Produits
    Set btn = ws.Buttons.Add(50, 220, 150, 30)
    btn.Text = "Gestion des Produits"
    btn.OnAction = "ShowProductForm"

    ' Bouton Saisie de Mouvement
    Set btn = ws.Buttons.Add(220, 220, 150, 30)
    btn.Text = "Saisie de Mouvement"
    btn.OnAction = "ShowMovementForm"

    ' Bouton Consultation Stock
    Set btn = ws.Buttons.Add(390, 220, 150, 30)
    btn.Text = "Consultation Stock"
    btn.OnAction = "ShowStockReport"

    ' Bouton Historique
    Set btn = ws.Buttons.Add(560, 220, 150, 30)
    btn.Text = "Historique"
    btn.OnAction = "ShowMovementHistory"

    ' Bouton Actualiser
    Set btn = ws.Buttons.Add(50, 270, 100, 25)
    btn.Text = "Actualiser"
    btn.OnAction = "RefreshDashboard"
End Sub
```

### 3. Formulaire de gestion des produits

Pour créer une interface utilisateur conviviale, nous utiliserons des UserForms. Voici comment créer le formulaire de gestion des produits :

```vba
Sub ShowProductForm()
    '=========================================
    ' Affichage du formulaire de gestion des produits
    '=========================================

    ' Charger et afficher le formulaire
    Load UserForm_Products
    UserForm_Products.Show
End Sub

' Code du UserForm (à placer dans le module du UserForm)
Private Sub UserForm_Initialize()
    '=========================================
    ' Initialisation du formulaire produits
    '=========================================

    ' Configuration du formulaire
    Me.Caption = "Gestion des Produits"
    Me.StartUpPosition = 0  ' Position manuelle
    Me.Left = 100
    Me.Top = 100

    ' Remplir la liste déroulante des catégories
    With ComboBox_Category
        .AddItem "Informatique"
        .AddItem "Bureautique"
        .AddItem "Mobilier"
        .AddItem "Consommables"
        .AddItem "Maintenance"
        .AddItem "Autre"
    End With

    ' Remplir la liste déroulante des statuts
    With ComboBox_Status
        .AddItem "Actif"
        .AddItem "Inactif"
        .AddItem "Obsolète"
        .Value = "Actif"  ' Valeur par défaut
    End With

    ' Charger la liste des produits existants
    Call LoadProductList

    ' Mode création par défaut
    Call ClearForm
End Sub

Private Sub LoadProductList()
    '=========================================
    ' Chargement de la liste des produits dans la ListBox
    '=========================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim productInfo As String

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Vider la liste
    ListBox_Products.Clear

    ' Charger tous les produits
    For i = 2 To lastRow  ' Commencer à la ligne 2 (après en-têtes)
        If ws.Cells(i, 1).Value <> "" Then
            productInfo = ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value & " (Stock: " & ws.Cells(i, 4).Value & ")"
            ListBox_Products.AddItem productInfo
        End If
    Next i
End Sub

Private Sub Button_Add_Click()
    '=========================================
    ' Ajout d'un nouveau produit
    '=========================================

    ' Validation des données
    If Not ValidateProductData() Then Exit Sub

    ' Vérifier l'unicité du code produit
    If ProductCodeExists(TextBox_Code.Value) Then
        MsgBox "Ce code produit existe déjà. Veuillez en choisir un autre.", vbExclamation
        TextBox_Code.SetFocus
        Exit Sub
    End If

    ' Ajouter le produit
    Call AddNewProduct

    ' Actualiser la liste
    Call LoadProductList

    ' Vider le formulaire pour le prochain produit
    Call ClearForm

    MsgBox "Produit ajouté avec succès !", vbInformation
End Sub

Private Function ValidateProductData() As Boolean
    '=========================================
    ' Validation des données du formulaire
    '=========================================

    ValidateProductData = True

    ' Vérifier le code produit
    If Trim(TextBox_Code.Value) = "" Then
        MsgBox "Le code produit est obligatoire.", vbExclamation
        TextBox_Code.SetFocus
        ValidateProductData = False
        Exit Function
    End If

    ' Vérifier la désignation
    If Trim(TextBox_Name.Value) = "" Then
        MsgBox "La désignation est obligatoire.", vbExclamation
        TextBox_Name.SetFocus
        ValidateProductData = False
        Exit Function
    End If

    ' Vérifier que les quantités sont numériques
    If Not IsNumeric(TextBox_CurrentStock.Value) Then
        MsgBox "Le stock actuel doit être un nombre.", vbExclamation
        TextBox_CurrentStock.SetFocus
        ValidateProductData = False
        Exit Function
    End If

    If Not IsNumeric(TextBox_MinStock.Value) Then
        MsgBox "Le stock minimum doit être un nombre.", vbExclamation
        TextBox_MinStock.SetFocus
        ValidateProductData = False
        Exit Function
    End If

    If Not IsNumeric(TextBox_Price.Value) Then
        MsgBox "Le prix doit être un nombre.", vbExclamation
        TextBox_Price.SetFocus
        ValidateProductData = False
        Exit Function
    End If

    ' Vérifier que les valeurs sont positives
    If CDbl(TextBox_CurrentStock.Value) < 0 Then
        MsgBox "Le stock actuel ne peut pas être négatif.", vbExclamation
        TextBox_CurrentStock.SetFocus
        ValidateProductData = False
        Exit Function
    End If
End Function

Private Sub AddNewProduct()
    '=========================================
    ' Ajout d'un nouveau produit dans la base
    '=========================================

    Dim ws As Worksheet
    Dim newRow As Long

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Remplir la nouvelle ligne
    With ws
        .Cells(newRow, 1).Value = UCase(Trim(TextBox_Code.Value))  ' Code en majuscules
        .Cells(newRow, 2).Value = Trim(TextBox_Name.Value)
        .Cells(newRow, 3).Value = ComboBox_Category.Value
        .Cells(newRow, 4).Value = CDbl(TextBox_CurrentStock.Value)
        .Cells(newRow, 5).Value = CDbl(TextBox_MinStock.Value)
        .Cells(newRow, 6).Value = CDbl(TextBox_Price.Value)
        .Cells(newRow, 7).Value = Trim(TextBox_Location.Value)
        .Cells(newRow, 8).Value = Now()  ' Date de création
        .Cells(newRow, 9).Value = ComboBox_Status.Value

        ' Formatage de la nouvelle ligne
        .Range(.Cells(newRow, 1), .Cells(newRow, 9)).Borders.LineStyle = xlContinuous

        ' Formatage conditionnel pour les alertes de stock
        If CDbl(TextBox_CurrentStock.Value) <= CDbl(TextBox_MinStock.Value) Then
            .Range(.Cells(newRow, 1), .Cells(newRow, 9)).Interior.Color = RGB(255, 200, 200)  ' Rouge clair
        End If
    End With

    ' Enregistrer le mouvement initial (stock initial)
    Call RecordMovement(UCase(Trim(TextBox_Code.Value)), "STOCK_INITIAL", CDbl(TextBox_CurrentStock.Value), 0, "Création du produit")
End Sub
```

### 4. Gestion des mouvements de stock

```vba
Sub ShowMovementForm()
    '=========================================
    ' Affichage du formulaire de saisie des mouvements
    '=========================================

    Load UserForm_Movements
    UserForm_Movements.Show
End Sub

' Code du UserForm_Movements
Private Sub UserForm_Initialize()
    '=========================================
    ' Initialisation du formulaire de mouvements
    '=========================================

    Me.Caption = "Saisie de Mouvement de Stock"

    ' Remplir la liste déroulante des types de mouvements
    With ComboBox_MovementType
        .AddItem "ENTREE"
        .AddItem "SORTIE"
        .AddItem "AJUSTEMENT"
        .AddItem "TRANSFERT"
        .AddItem "RETOUR"
        .AddItem "PERTE"
    End With

    ' Charger la liste des produits
    Call LoadProductCodes

    ' Utilisateur par défaut (nom de l'utilisateur Windows)
    TextBox_User.Value = Environ("USERNAME")

    ' Date et heure actuelles
    TextBox_Date.Value = Format(Now(), "dd/mm/yyyy")
    TextBox_Time.Value = Format(Now(), "hh:mm")
End Sub

Private Sub LoadProductCodes()
    '=========================================
    ' Chargement des codes produits disponibles
    '=========================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ComboBox_ProductCode.Clear

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" And ws.Cells(i, 9).Value = "Actif" Then
            ComboBox_ProductCode.AddItem ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value
        End If
    Next i
End Sub

Private Sub ComboBox_ProductCode_Change()
    '=========================================
    ' Mise à jour automatique des informations produit
    '=========================================

    Dim productCode As String
    Dim ws As Worksheet
    Dim findResult As Range

    If ComboBox_ProductCode.Value = "" Then Exit Sub

    ' Extraire le code produit de la sélection
    productCode = Split(ComboBox_ProductCode.Value, " - ")(0)

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Set findResult = ws.Columns(1).Find(productCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not findResult Is Nothing Then
        Label_CurrentStock.Caption = "Stock actuel : " & ws.Cells(findResult.Row, 4).Value
        Label_ProductName.Caption = ws.Cells(findResult.Row, 2).Value
        Label_Location.Caption = "Emplacement : " & ws.Cells(findResult.Row, 7).Value

        ' Colorer en rouge si stock faible
        If ws.Cells(findResult.Row, 4).Value <= ws.Cells(findResult.Row, 5).Value Then
            Label_CurrentStock.ForeColor = RGB(255, 0, 0)
            Label_CurrentStock.Caption = Label_CurrentStock.Caption & " ⚠️ STOCK FAIBLE"
        Else
            Label_CurrentStock.ForeColor = RGB(0, 0, 0)
        End If
    End If
End Sub

Private Sub Button_ValidateMovement_Click()
    '=========================================
    ' Validation et enregistrement du mouvement
    '=========================================

    ' Validation des données
    If Not ValidateMovementData() Then Exit Sub

    ' Extraire le code produit
    Dim productCode As String
    productCode = Split(ComboBox_ProductCode.Value, " - ")(0)

    ' Vérifier la disponibilité pour les sorties
    If ComboBox_MovementType.Value = "SORTIE" Then
        If Not CheckStockAvailability(productCode, CDbl(TextBox_Quantity.Value)) Then
            Exit Sub
        End If
    End If

    ' Enregistrer le mouvement
    Call ProcessMovement(productCode)

    ' Actualiser l'affichage
    Call ComboBox_ProductCode_Change

    ' Vider les champs pour le prochain mouvement
    TextBox_Quantity.Value = ""
    TextBox_Comment.Value = ""

    MsgBox "Mouvement enregistré avec succès !", vbInformation
End Sub

Private Function ValidateMovementData() As Boolean
    '=========================================
    ' Validation des données de mouvement
    '=========================================

    ValidateMovementData = True

    ' Vérifier la sélection du produit
    If ComboBox_ProductCode.Value = "" Then
        MsgBox "Veuillez sélectionner un produit.", vbExclamation
        ComboBox_ProductCode.SetFocus
        ValidateMovementData = False
        Exit Function
    End If

    ' Vérifier le type de mouvement
    If ComboBox_MovementType.Value = "" Then
        MsgBox "Veuillez sélectionner un type de mouvement.", vbExclamation
        ComboBox_MovementType.SetFocus
        ValidateMovementData = False
        Exit Function
    End If

    ' Vérifier la quantité
    If Not IsNumeric(TextBox_Quantity.Value) Then
        MsgBox "La quantité doit être un nombre positif.", vbExclamation
        TextBox_Quantity.SetFocus
        ValidateMovementData = False
        Exit Function
    End If

    If CDbl(TextBox_Quantity.Value) <= 0 Then
        MsgBox "La quantité doit être un nombre positif.", vbExclamation
        TextBox_Quantity.SetFocus
        ValidateMovementData = False
        Exit Function
    End If

    ' Vérifier l'utilisateur
    If Trim(TextBox_User.Value) = "" Then
        MsgBox "Le nom de l'utilisateur est obligatoire.", vbExclamation
        TextBox_User.SetFocus
        ValidateMovementData = False
        Exit Function
    End If
End Function

Private Function CheckStockAvailability(productCode As String, quantity As Double) As Boolean
    '=========================================
    ' Vérification de la disponibilité du stock
    '=========================================

    Dim ws As Worksheet
    Dim findResult As Range
    Dim currentStock As Double

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Set findResult = ws.Columns(1).Find(productCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not findResult Is Nothing Then
        currentStock = ws.Cells(findResult.Row, 4).Value

        If currentStock < quantity Then
            MsgBox "Stock insuffisant !" & vbNewLine & _
                   "Stock disponible : " & currentStock & vbNewLine & _
                   "Quantité demandée : " & quantity, vbExclamation
            CheckStockAvailability = False
        Else
            CheckStockAvailability = True
        End If
    Else
        MsgBox "Produit non trouvé.", vbCritical
        CheckStockAvailability = False
    End If
End Function

Private Sub ProcessMovement(productCode As String)
    '=========================================
    ' Traitement du mouvement de stock
    '=========================================

    Dim ws As Worksheet
    Dim findResult As Range
    Dim currentStock As Double
    Dim newStock As Double
    Dim quantity As Double
    Dim movementType As String

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Set findResult = ws.Columns(1).Find(productCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not findResult Is Nothing Then
        ' Récupérer les valeurs
        currentStock = ws.Cells(findResult.Row, 4).Value
        quantity = CDbl(TextBox_Quantity.Value)
        movementType = ComboBox_MovementType.Value

        ' Calculer le nouveau stock selon le type de mouvement
        Select Case movementType
            Case "ENTREE", "RETOUR"
                newStock = currentStock + quantity
            Case "SORTIE", "PERTE"
                newStock = currentStock - quantity
            Case "AJUSTEMENT"
                ' Pour un ajustement, la quantité saisie est le nouveau stock souhaité
                newStock = quantity
                quantity = newStock - currentStock  ' Calculer la différence
            Case "TRANSFERT"
                ' Pour simplifier, on traite comme une sortie
                newStock = currentStock - quantity
        End Select

        ' Mettre à jour le stock dans la base
        ws.Cells(findResult.Row, 4).Value = newStock

        ' Appliquer un formatage conditionnel pour les alertes
        Call ApplyStockFormatting(findResult.Row)

        ' Enregistrer le mouvement dans l'historique
        Call RecordMovement(productCode, movementType, quantity, currentStock, TextBox_Comment.Value)

        ' Vérifier et mettre à jour les alertes
        Call UpdateStockAlerts
    End If
End Sub

Private Sub ApplyStockFormatting(rowNumber As Long)
    '=========================================
    ' Application du formatage conditionnel selon le niveau de stock
    '=========================================

    Dim ws As Worksheet
    Dim currentStock As Double
    Dim minStock As Double

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    currentStock = ws.Cells(rowNumber, 4).Value
    minStock = ws.Cells(rowNumber, 5).Value

    ' Supprimer le formatage existant
    ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 9)).Interior.ColorIndex = xlNone

    ' Appliquer le nouveau formatage selon les seuils
    If currentStock <= 0 Then
        ' Stock épuisé - Rouge vif
        ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 9)).Interior.Color = RGB(255, 0, 0)
        ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 9)).Font.Color = vbWhite
    ElseIf currentStock <= minStock Then
        ' Stock faible - Orange
        ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 9)).Interior.Color = RGB(255, 165, 0)
    ElseIf currentStock <= minStock * 1.5 Then
        ' Stock en surveillance - Jaune clair
        ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 9)).Interior.Color = RGB(255, 255, 0)
    End If
End Sub

Sub RecordMovement(productCode As String, movementType As String, quantity As Double, stockBefore As Double, comment As String)
    '=========================================
    ' Enregistrement d'un mouvement dans l'historique
    '=========================================

    Dim ws As Worksheet
    Dim wsProducts As Worksheet
    Dim newRow As Long
    Dim productName As String
    Dim findResult As Range

    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")
    Set wsProducts = ThisWorkbook.Sheets("PRODUITS")

    ' Trouver le nom du produit
    Set findResult = wsProducts.Columns(1).Find(productCode, LookIn:=xlValues, LookAt:=xlWhole)
    If Not findResult Is Nothing Then
        productName = wsProducts.Cells(findResult.Row, 2).Value
    Else
        productName = "Produit inconnu"
    End If

    ' Trouver la prochaine ligne disponible
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Enregistrer le mouvement
    With ws
        .Cells(newRow, 1).Value = Date  ' Date
        .Cells(newRow, 2).Value = Time  ' Heure
        .Cells(newRow, 3).Value = productCode
        .Cells(newRow, 4).Value = productName
        .Cells(newRow, 5).Value = movementType
        .Cells(newRow, 6).Value = quantity
        .Cells(newRow, 7).Value = stockBefore
        .Cells(newRow, 8).Value = stockBefore + IIf(movementType = "ENTREE" Or movementType = "RETOUR" Or movementType = "AJUSTEMENT", quantity, -quantity)
        .Cells(newRow, 9).Value = Environ("USERNAME")  ' Utilisateur Windows
        .Cells(newRow, 10).Value = comment

        ' Formatage de la nouvelle ligne
        .Range(.Cells(newRow, 1), .Cells(newRow, 10)).Borders.LineStyle = xlContinuous

        ' Couleur selon le type de mouvement
        Select Case movementType
            Case "ENTREE", "RETOUR"
                .Range(.Cells(newRow, 1), .Cells(newRow, 10)).Interior.Color = RGB(200, 255, 200)  ' Vert clair
            Case "SORTIE", "PERTE"
                .Range(.Cells(newRow, 1), .Cells(newRow, 10)).Interior.Color = RGB(255, 200, 200)  ' Rouge clair
            Case "AJUSTEMENT"
                .Range(.Cells(newRow, 1), .Cells(newRow, 10)).Interior.Color = RGB(200, 200, 255)  ' Bleu clair
        End Select
    End With
End Sub
```

## 5. Système d'alertes automatiques

```vba
Sub UpdateStockAlerts(Optional targetSheet As Worksheet)
    '=========================================
    ' Mise à jour des alertes de stock faible
    '=========================================

    Dim wsProducts As Worksheet
    Dim wsAlerts As Worksheet
    Dim wsDashboard As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim alertRow As Long
    Dim alertCount As Integer

    Set wsProducts = ThisWorkbook.Sheets("PRODUITS")
    Set wsAlerts = ThisWorkbook.Sheets("ALERTES")

    ' Vider la feuille d'alertes
    wsAlerts.Cells.Clear
    Call SetupAlertSheet(wsAlerts)

    lastRow = wsProducts.Cells(wsProducts.Rows.Count, 1).End(xlUp).Row
    alertRow = 2  ' Commencer après les en-têtes
    alertCount = 0

    Dim currentStock As Double
    Dim minStock As Double
    Dim alertLevel As String

    ' Parcourir tous les produits
    For i = 2 To lastRow
        If wsProducts.Cells(i, 1).Value <> "" And wsProducts.Cells(i, 9).Value = "Actif" Then
            currentStock = wsProducts.Cells(i, 4).Value
            minStock = wsProducts.Cells(i, 5).Value

            ' Déterminer le niveau d'alerte
            If currentStock <= 0 Then
                alertLevel = "CRITIQUE - RUPTURE"
            ElseIf currentStock <= minStock Then
                alertLevel = "URGENT - STOCK FAIBLE"
            ElseIf currentStock <= minStock * 1.5 Then
                alertLevel = "ATTENTION - SURVEILLANCE"
            Else
                alertLevel = ""  ' Pas d'alerte
            End If

            ' Si alerte nécessaire, l'ajouter à la feuille
            If alertLevel <> "" Then
                With wsAlerts
                    .Cells(alertRow, 1).Value = wsProducts.Cells(i, 1).Value  ' Code
                    .Cells(alertRow, 2).Value = wsProducts.Cells(i, 2).Value  ' Désignation
                    .Cells(alertRow, 3).Value = currentStock
                    .Cells(alertRow, 4).Value = minStock
                    .Cells(alertRow, 5).Value = alertLevel
                    .Cells(alertRow, 6).Value = wsProducts.Cells(i, 7).Value  ' Emplacement
                    .Cells(alertRow, 7).Value = Now()  ' Date de l'alerte

                    ' Formatage selon le niveau d'alerte
                    Select Case True
                        Case InStr(alertLevel, "CRITIQUE") > 0
                            .Range(.Cells(alertRow, 1), .Cells(alertRow, 7)).Interior.Color = RGB(255, 0, 0)
                            .Range(.Cells(alertRow, 1), .Cells(alertRow, 7)).Font.Color = vbWhite
                        Case InStr(alertLevel, "URGENT") > 0
                            .Range(.Cells(alertRow, 1), .Cells(alertRow, 7)).Interior.Color = RGB(255, 165, 0)
                        Case InStr(alertLevel, "ATTENTION") > 0
                            .Range(.Cells(alertRow, 1), .Cells(alertRow, 7)).Interior.Color = RGB(255, 255, 0)
                    End Select

                    alertRow = alertRow + 1
                    alertCount = alertCount + 1
                End With
            End If
        End If
    Next i

    ' Mettre à jour le tableau de bord si spécifié
    If Not targetSheet Is Nothing Then
        Call UpdateDashboardAlerts(targetSheet, alertCount)
    End If

    ' Ajuster les colonnes
    wsAlerts.Columns.AutoFit
End Sub

Sub SetupAlertSheet(ws As Worksheet)
    '=========================================
    ' Configuration de la feuille ALERTES
    '=========================================

    With ws
        ' En-têtes
        .Cells(1, 1).Value = "Code Produit"
        .Cells(1, 2).Value = "Désignation"
        .Cells(1, 3).Value = "Stock Actuel"
        .Cells(1, 4).Value = "Stock Minimum"
        .Cells(1, 5).Value = "Niveau d'Alerte"
        .Cells(1, 6).Value = "Emplacement"
        .Cells(1, 7).Value = "Date Alerte"

        ' Formatage des en-têtes
        With .Range("A1:G1")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(255, 0, 0)  ' Rouge pour les alertes
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    End With
End Sub

Sub UpdateDashboardAlerts(ws As Worksheet, alertCount As Integer)
    '=========================================
    ' Mise à jour de la section alertes du tableau de bord
    '=========================================

    Dim startRow As Long
    Dim wsAlerts As Worksheet
    Dim i As Long

    startRow = 20  ' Position de départ des alertes sur le tableau de bord
    Set wsAlerts = ThisWorkbook.Sheets("ALERTES")

    ' Nettoyer la zone d'alertes
    ws.Range("A" & startRow & ":H30").Clear

    If alertCount = 0 Then
        ws.Cells(startRow, 1).Value = "✅ Aucune alerte - Tous les stocks sont corrects"
        ws.Cells(startRow, 1).Font.Color = RGB(0, 128, 0)  ' Vert
        ws.Cells(startRow, 1).Font.Bold = True
    Else
        ' Afficher les alertes les plus critiques (max 5)
        Dim displayCount As Integer
        displayCount = Application.Min(5, alertCount)

        For i = 1 To displayCount
            With ws
                .Cells(startRow + i - 1, 1).Value = "⚠️ " & wsAlerts.Cells(i + 1, 1).Value & " - " & wsAlerts.Cells(i + 1, 2).Value
                .Cells(startRow + i - 1, 1).Font.Bold = True

                ' Couleur selon le niveau d'alerte
                If InStr(wsAlerts.Cells(i + 1, 5).Value, "CRITIQUE") > 0 Then
                    .Cells(startRow + i - 1, 1).Font.Color = RGB(255, 0, 0)
                ElseIf InStr(wsAlerts.Cells(i + 1, 5).Value, "URGENT") > 0 Then
                    .Cells(startRow + i - 1, 1).Font.Color = RGB(255, 165, 0)
                Else
                    .Cells(startRow + i - 1, 1).Font.Color = RGB(255, 255, 0)
                End If

                .Cells(startRow + i - 1, 6).Value = "Stock: " & wsAlerts.Cells(i + 1, 3).Value
                .Cells(startRow + i - 1, 6).Font.Bold = True
            End With
        Next i

        ' Message si plus d'alertes
        If alertCount > 5 Then
            ws.Cells(startRow + 5, 1).Value = "... et " & (alertCount - 5) & " autre(s) alerte(s) - Voir feuille ALERTES"
            ws.Cells(startRow + 5, 1).Font.Italic = True
        End If
    End If
End Sub
```

## 6. Module de rapports et consultation

```vba
Sub ShowStockReport()
    '=========================================
    ' Génération d'un rapport de stock complet
    '=========================================

    Dim wsReport As Worksheet
    Dim wsProducts As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim reportRow As Long

    Set wsProducts = ThisWorkbook.Sheets("PRODUITS")
    Set wsReport = ThisWorkbook.Sheets("RAPPORTS")

    ' Nettoyer la feuille de rapport
    wsReport.Cells.Clear

    ' En-tête du rapport
    With wsReport
        .Cells(1, 1).Value = "RAPPORT DE STOCK COMPLET"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        .Range("A1:J1").Merge
        .Range("A1").HorizontalAlignment = xlCenter

        .Cells(2, 1).Value = "Généré le : " & Format(Now(), "dd/mm/yyyy à hh:mm")
        .Range("A2:J2").Merge
        .Range("A2").HorizontalAlignment = xlCenter

        ' En-têtes des colonnes du rapport
        .Cells(4, 1).Value = "Code"
        .Cells(4, 2).Value = "Désignation"
        .Cells(4, 3).Value = "Catégorie"
        .Cells(4, 4).Value = "Stock Actuel"
        .Cells(4, 5).Value = "Stock Min"
        .Cells(4, 6).Value = "Prix Unit."
        .Cells(4, 7).Value = "Valeur Stock"
        .Cells(4, 8).Value = "Emplacement"
        .Cells(4, 9).Value = "Statut"
        .Cells(4, 10).Value = "Niveau"

        ' Formatage des en-têtes
        With .Range("A4:J4")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(68, 114, 196)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    End With

    ' Remplir le rapport avec les données
    lastRow = wsProducts.Cells(wsProducts.Rows.Count, 1).End(xlUp).Row
    reportRow = 5

    Dim currentStock As Double
    Dim minStock As Double
    Dim stockLevel As String

    For i = 2 To lastRow
        If wsProducts.Cells(i, 1).Value <> "" Then
            With wsReport
                .Cells(reportRow, 1).Value = wsProducts.Cells(i, 1).Value  ' Code
                .Cells(reportRow, 2).Value = wsProducts.Cells(i, 2).Value  ' Désignation
                .Cells(reportRow, 3).Value = wsProducts.Cells(i, 3).Value  ' Catégorie
                .Cells(reportRow, 4).Value = wsProducts.Cells(i, 4).Value  ' Stock actuel
                .Cells(reportRow, 5).Value = wsProducts.Cells(i, 5).Value  ' Stock min
                .Cells(reportRow, 6).Value = wsProducts.Cells(i, 6).Value  ' Prix
                .Cells(reportRow, 7).Value = wsProducts.Cells(i, 4).Value * wsProducts.Cells(i, 6).Value  ' Valeur
                .Cells(reportRow, 8).Value = wsProducts.Cells(i, 7).Value  ' Emplacement
                .Cells(reportRow, 9).Value = wsProducts.Cells(i, 9).Value  ' Statut

                ' Déterminer le niveau de stock
                currentStock = wsProducts.Cells(i, 4).Value
                minStock = wsProducts.Cells(i, 5).Value

                If currentStock <= 0 Then
                    stockLevel = "RUPTURE"
                ElseIf currentStock <= minStock Then
                    stockLevel = "FAIBLE"
                ElseIf currentStock <= minStock * 1.5 Then
                    stockLevel = "SURVEILLANCE"
                Else
                    stockLevel = "NORMAL"
                End If

                .Cells(reportRow, 10).Value = stockLevel

                ' Formatage conditionnel
                Select Case stockLevel
                    Case "RUPTURE"
                        .Range(.Cells(reportRow, 1), .Cells(reportRow, 10)).Interior.Color = RGB(255, 0, 0)
                        .Range(.Cells(reportRow, 1), .Cells(reportRow, 10)).Font.Color = vbWhite
                    Case "FAIBLE"
                        .Range(.Cells(reportRow, 1), .Cells(reportRow, 10)).Interior.Color = RGB(255, 165, 0)
                    Case "SURVEILLANCE"
                        .Range(.Cells(reportRow, 1), .Cells(reportRow, 10)).Interior.Color = RGB(255, 255, 0)
                End Select

                ' Bordures
                .Range(.Cells(reportRow, 1), .Cells(reportRow, 10)).Borders.LineStyle = xlContinuous

                reportRow = reportRow + 1
            End With
        End If
    Next i

    ' Statistiques en bas du rapport
    Call AddReportStatistics(wsReport, reportRow + 2)

    ' Ajustement des colonnes
    wsReport.Columns.AutoFit

    ' Activer la feuille de rapport
    wsReport.Activate
    wsReport.Cells(1, 1).Select

    MsgBox "Rapport de stock généré avec succès !", vbInformation
End Sub

Sub AddReportStatistics(ws As Worksheet, startRow As Long)
    '=========================================
    ' Ajout de statistiques au rapport
    '=========================================

    With ws
        ' Titre des statistiques
        .Cells(startRow, 1).Value = "STATISTIQUES GÉNÉRALES"
        .Cells(startRow, 1).Font.Size = 14
        .Cells(startRow, 1).Font.Bold = True
        .Range(.Cells(startRow, 1), .Cells(startRow, 4)).Merge

        startRow = startRow + 2

        ' Calculs statistiques
        .Cells(startRow, 1).Value = "Nombre total de produits :"
        .Cells(startRow, 2).Value = "=COUNTA(A5:A1000)"

        .Cells(startRow + 1, 1).Value = "Produits en rupture :"
        .Cells(startRow + 1, 2).Value = "=COUNTIF(J5:J1000,\"RUPTURE\")"

        .Cells(startRow + 2, 1).Value = "Produits en stock faible :"
        .Cells(startRow + 2, 2).Value = "=COUNTIF(J5:J1000,\"FAIBLE\")"

        .Cells(startRow + 3, 1).Value = "Valeur totale du stock :"
        .Cells(startRow + 3, 2).Value = "=SUM(G5:G1000)"
        .Cells(startRow + 3, 2).NumberFormat = "#,##0.00 €"

        .Cells(startRow + 4, 1).Value = "Valeur moyenne par produit :"
        .Cells(startRow + 4, 2).Value = "=AVERAGE(G5:G1000)"
        .Cells(startRow + 4, 2).NumberFormat = "#,##0.00 €"

        ' Formatage des statistiques
        .Range(.Cells(startRow, 1), .Cells(startRow + 4, 1)).Font.Bold = True
        .Range(.Cells(startRow, 2), .Cells(startRow + 4, 2)).Font.Bold = True
        .Range(.Cells(startRow, 2), .Cells(startRow + 4, 2)).Interior.Color = RGB(242, 242, 242)
    End With
End Sub
```

## 7. Historique et recherche avancée

```vba
Sub ShowMovementHistory()
    '=========================================
    ' Affichage de l'historique des mouvements avec filtres
    '=========================================

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")

    ' Activer la feuille mouvements
    ws.Activate

    ' Ajouter des filtres automatiques si pas déjà présents
    If Not ws.AutoFilterMode Then
        ws.Range("A1").AutoFilter
    End If

    ' Trier par date décroissante (plus récent en premier)
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A2:A1000"), SortOn:=xlSortOnValues, Order:=xlDescending
    ws.Sort.SortFields.Add Key:=ws.Range("B2:B1000"), SortOn:=xlSortOnValues, Order:=xlDescending

    With ws.Sort
        .SetRange ws.Range("A1:J1000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Message d'information
    MsgBox "Historique des mouvements affiché." & vbNewLine & _
           "Utilisez les filtres pour rechercher des mouvements spécifiques.", vbInformation
End Sub

Sub SearchProduct()
    '=========================================
    ' Fonction de recherche avancée de produits
    '=========================================

    Dim searchTerm As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    Dim resultMessage As String

    ' Demander le terme de recherche
    searchTerm = InputBox("Entrez le code produit ou une partie du nom à rechercher :", "Recherche de produit")

    If searchTerm = "" Then Exit Sub

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    found = False
    resultMessage = "Résultats de la recherche pour '" & searchTerm & "' :" & vbNewLine & vbNewLine

    ' Rechercher dans les codes et noms de produits
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            If InStr(1, UCase(ws.Cells(i, 1).Value), UCase(searchTerm)) > 0 Or _
               InStr(1, UCase(ws.Cells(i, 2).Value), UCase(searchTerm)) > 0 Then

                found = True
                resultMessage = resultMessage & _
                    "Code: " & ws.Cells(i, 1).Value & vbNewLine & _
                    "Nom: " & ws.Cells(i, 2).Value & vbNewLine & _
                    "Stock: " & ws.Cells(i, 4).Value & vbNewLine & _
                    "Emplacement: " & ws.Cells(i, 7).Value & vbNewLine & _
                    "Statut: " & ws.Cells(i, 9).Value & vbNewLine & vbNewLine
            End If
        End If
    Next i

    If found Then
        MsgBox resultMessage, vbInformation, "Résultats de recherche"
    Else
        MsgBox "Aucun produit trouvé pour '" & searchTerm & "'", vbExclamation
    End If
End Sub
```

## 8. Fonctions utilitaires et actualisation

```vba
Sub RefreshDashboard()
    '=========================================
    ' Actualisation complète du tableau de bord
    '=========================================

    Application.ScreenUpdating = False

    ' Mettre à jour les alertes
    Call UpdateStockAlerts(ThisWorkbook.Sheets("ACCUEIL"))

    ' Recalculer toutes les formules
    Application.Calculate

    ' Mettre à jour la date/heure
    ThisWorkbook.Sheets("ACCUEIL").Cells(3, 2).Value = Now()

    Application.ScreenUpdating = True

    MsgBox "Tableau de bord actualisé !", vbInformation
End Sub

Function ProductCodeExists(productCode As String) As Boolean
    '=========================================
    ' Vérification de l'existence d'un code produit
    '=========================================

    Dim ws As Worksheet
    Dim findResult As Range

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Set findResult = ws.Columns(1).Find(productCode, LookIn:=xlValues, LookAt:=xlWhole)

    ProductCodeExists = Not findResult Is Nothing
End Function

Sub ExportToCSV()
    '=========================================
    ' Export des données vers un fichier CSV
    '=========================================

    Dim ws As Worksheet
    Dim fileName As String
    Dim filePath As String

    Set ws = ThisWorkbook.Sheets("PRODUITS")

    ' Générer le nom de fichier
    fileName = "Stock_Export_" & Format(Now(), "yyyy-mm-dd_hh-mm") & ".csv"
    filePath = ThisWorkbook.Path & "\" & fileName

    ' Sauvegarder en CSV
    ws.Copy
    ActiveWorkbook.SaveAs filePath, xlCSV
    ActiveWorkbook.Close False

    MsgBox "Export réalisé avec succès !" & vbNewLine & "Fichier : " & fileName, vbInformation
End Sub

Sub CreateBackup()
    '=========================================
    ' Création d'une sauvegarde complète
    '=========================================

    Dim backupName As String
    Dim backupPath As String

    backupName = "Sauvegarde_Stock_" & Format(Now(), "yyyy-mm-dd_hh-mm") & ".xlsm"
    backupPath = ThisWorkbook.Path & "\Sauvegardes"

    ' Créer le dossier de sauvegarde s'il n'existe pas
    If Dir(backupPath, vbDirectory) = "" Then
        MkDir backupPath
    End If

    ' Sauvegarder le fichier
    ThisWorkbook.SaveCopyAs backupPath & "\" & backupName

    MsgBox "Sauvegarde créée avec succès !" & vbNewLine & "Fichier : " & backupName, vbInformation
End Sub
```

## Installation et utilisation du système

### Étapes d'installation

1. **Création du fichier**
   - Ouvrir un nouveau classeur Excel
   - Sauvegarder au format .xlsm (avec macros)
   - Activer les macros si demandé

2. **Ajout du code VBA**
   - Ouvrir l'éditeur VBA (Alt + F11)
   - Créer un nouveau module standard
   - Copier tout le code développé
   - Créer les UserForms nécessaires

3. **Initialisation du système**
   - Exécuter la macro `InitializeStockSystem`
   - Vérifier la création des feuilles
   - Tester les fonctionnalités de base

### Guide d'utilisation quotidienne

#### Ajout d'un nouveau produit
1. Cliquer sur "Gestion des Produits" dans le tableau de bord
2. Remplir tous les champs obligatoires
3. Définir le stock minimum pour les alertes
4. Valider l'ajout

#### Saisie d'un mouvement
1. Cliquer sur "Saisie de Mouvement"
2. Sélectionner le produit concerné
3. Choisir le type de mouvement (ENTREE/SORTIE)
4. Indiquer la quantité et un commentaire
5. Valider la saisie

#### Consultation des stocks
1. Utiliser le tableau de bord pour une vue d'ensemble
2. Cliquer sur "Consultation Stock" pour un rapport détaillé
3. Utiliser les filtres pour rechercher des produits spécifiques

#### Suivi des alertes
- Les alertes s'affichent automatiquement sur le tableau de bord
- Consulter la feuille "ALERTES" pour le détail complet
- Les produits en rupture ou stock faible sont colorés

## Fonctionnalités avancées

### Création d'un système de codes-barres simple

```vba
Sub GenerateBarcodeSheet()
    '=========================================
    ' Génération d'une feuille avec codes-barres simples
    '=========================================

    Dim ws As Worksheet
    Dim wsProducts As Worksheet
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim printRow As Long

    Set wsProducts = ThisWorkbook.Sheets("PRODUITS")

    ' Créer une nouvelle feuille pour les codes-barres
    Set newSheet = ThisWorkbook.Sheets.Add
    newSheet.Name = "CODES_BARRES_" & Format(Now(), "ddmm")

    ' Configuration de la feuille
    With newSheet
        .Cells(1, 1).Value = "ÉTIQUETTES CODES-BARRES"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        .Range("A1:D1").Merge
        .Range("A1").HorizontalAlignment = xlCenter

        .Cells(2, 1).Value = "Généré le : " & Format(Now(), "dd/mm/yyyy")
        .Range("A2:D2").Merge
        .Range("A2").HorizontalAlignment = xlCenter
    End With

    lastRow = wsProducts.Cells(wsProducts.Rows.Count, 1).End(xlUp).Row
    printRow = 4

    ' Générer les étiquettes
    For i = 2 To lastRow
        If wsProducts.Cells(i, 1).Value <> "" And wsProducts.Cells(i, 9).Value = "Actif" Then
            With newSheet
                ' Cadre de l'étiquette
                .Range(.Cells(printRow, 1), .Cells(printRow + 3, 4)).Borders.LineStyle = xlContinuous
                .Range(.Cells(printRow, 1), .Cells(printRow + 3, 4)).Borders.Weight = xlMedium

                ' Code produit en gros
                .Cells(printRow, 1).Value = wsProducts.Cells(i, 1).Value
                .Cells(printRow, 1).Font.Size = 20
                .Cells(printRow, 1).Font.Bold = True
                .Range(.Cells(printRow, 1), .Cells(printRow, 4)).Merge
                .Range(.Cells(printRow, 1), .Cells(printRow, 4)).HorizontalAlignment = xlCenter

                ' Nom du produit
                .Cells(printRow + 1, 1).Value = wsProducts.Cells(i, 2).Value
                .Cells(printRow + 1, 1).Font.Size = 12
                .Range(.Cells(printRow + 1, 1), .Cells(printRow + 1, 4)).Merge
                .Range(.Cells(printRow + 1, 1), .Cells(printRow + 1, 4)).HorizontalAlignment = xlCenter

                ' Informations complémentaires
                .Cells(printRow + 2, 1).Value = "Emplacement: " & wsProducts.Cells(i, 7).Value
                .Cells(printRow + 2, 3).Value = "Stock: " & wsProducts.Cells(i, 4).Value

                ' Code-barres simulé (caractères spéciaux)
                .Cells(printRow + 3, 1).Value = "||||| |||| | |||| |||||"
                .Cells(printRow + 3, 1).Font.Name = "Courier New"
                .Range(.Cells(printRow + 3, 1), .Cells(printRow + 3, 4)).Merge
                .Range(.Cells(printRow + 3, 1), .Cells(printRow + 3, 4)).HorizontalAlignment = xlCenter

                printRow = printRow + 5  ' Espace entre étiquettes
            End With
        End If
    Next i

    ' Ajuster les colonnes et configurer pour impression
    newSheet.Columns.AutoFit
    newSheet.PageSetup.Orientation = xlPortrait
    newSheet.PageSetup.FitToPagesWide = 1

    MsgBox "Feuille de codes-barres générée avec succès !", vbInformation
End Sub
```

### Système de notifications par email

```vba
Sub SendStockAlertEmail()
    '=========================================
    ' Envoi d'email d'alerte pour stocks faibles
    '=========================================

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim wsAlerts As Worksheet
    Dim lastRow As Long
    Dim emailBody As String
    Dim i As Long
    Dim alertCount As Integer

    ' Mettre à jour les alertes
    Call UpdateStockAlerts

    Set wsAlerts = ThisWorkbook.Sheets("ALERTES")
    lastRow = wsAlerts.Cells(wsAlerts.Rows.Count, 1).End(xlUp).Row
    alertCount = lastRow - 1  ' -1 pour exclure l'en-tête

    If alertCount = 0 Then
        MsgBox "Aucune alerte à envoyer - Tous les stocks sont corrects.", vbInformation
        Exit Sub
    End If

    ' Créer l'application Outlook
    On Error Resume Next
    Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If OutlookApp Is Nothing Then
        MsgBox "Impossible d'accéder à Outlook pour envoyer l'email.", vbExclamation
        Exit Sub
    End If

    ' Créer l'email
    Set OutlookMail = OutlookApp.CreateItem(0)  ' olMailItem = 0

    ' Construire le corps de l'email
    emailBody = "<html><body>"
    emailBody = emailBody & "<h2>🚨 ALERTE STOCK - " & Format(Now(), "dd/mm/yyyy") & "</h2>"
    emailBody = emailBody & "<p>Bonjour,</p>"
    emailBody = emailBody & "<p>Nous avons détecté <strong>" & alertCount & " produit(s)</strong> nécessitant votre attention :</p>"
    emailBody = emailBody & "<table border='1' style='border-collapse: collapse; width: 100%;'>"
    emailBody = emailBody & "<tr style='background-color: #ff0000; color: white;'>"
    emailBody = emailBody & "<th>Code</th><th>Produit</th><th>Stock Actuel</th><th>Stock Min</th><th>Niveau d'Alerte</th><th>Emplacement</th>"
    emailBody = emailBody & "</tr>"

    ' Ajouter chaque alerte
    For i = 2 To lastRow
        emailBody = emailBody & "<tr>"
        emailBody = emailBody & "<td>" & wsAlerts.Cells(i, 1).Value & "</td>"
        emailBody = emailBody & "<td>" & wsAlerts.Cells(i, 2).Value & "</td>"
        emailBody = emailBody & "<td>" & wsAlerts.Cells(i, 3).Value & "</td>"
        emailBody = emailBody & "<td>" & wsAlerts.Cells(i, 4).Value & "</td>"
        emailBody = emailBody & "<td><strong>" & wsAlerts.Cells(i, 5).Value & "</strong></td>"
        emailBody = emailBody & "<td>" & wsAlerts.Cells(i, 6).Value & "</td>"
        emailBody = emailBody & "</tr>"
    Next i

    emailBody = emailBody & "</table>"
    emailBody = emailBody & "<p><strong>Actions recommandées :</strong></p>"
    emailBody = emailBody & "<ul>"
    emailBody = emailBody & "<li>Vérifier les stocks physiques</li>"
    emailBody = emailBody & "<li>Passer commande si nécessaire</li>"
    emailBody = emailBody & "<li>Mettre à jour les seuils d'alerte si besoin</li>"
    emailBody = emailBody & "</ul>"
    emailBody = emailBody & "<p>Cordialement,<br>Système de Gestion de Stock</p>"
    emailBody = emailBody & "</body></html>"

    ' Configurer l'email
    With OutlookMail
        .To = "gestionnaire@entreprise.com"  ' À modifier selon vos besoins
        .CC = "direction@entreprise.com"     ' À modifier selon vos besoins
        .Subject = "🚨 ALERTE STOCK - " & alertCount & " produit(s) en alerte - " & Format(Now(), "dd/mm/yyyy")
        .HTMLBody = emailBody
        .Importance = 2  ' olImportanceHigh = 2

        ' Afficher l'email (ne pas l'envoyer automatiquement pour validation)
        .Display  ' Utiliser .Send pour envoi automatique
    End With

    ' Nettoyer les objets
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

    MsgBox "Email d'alerte préparé et affiché pour validation.", vbInformation
End Sub
```

### Intégration avec un scanner de codes-barres

```vba
Sub ProcessBarcodeInput()
    '=========================================
    ' Traitement d'une saisie de code-barres
    ' À connecter à un scanner ou saisie manuelle
    '=========================================

    Dim scannedCode As String
    Dim ws As Worksheet
    Dim findResult As Range
    Dim actionType As String
    Dim quantityInput As String
    Dim quantity As Double

    ' Demander le code scanné (ou saisi manuellement)
    scannedCode = InputBox("Scannez ou saisissez le code produit :", "Saisie Code-Barres")

    If scannedCode = "" Then Exit Sub

    ' Nettoyer le code (supprimer espaces, majuscules)
    scannedCode = UCase(Trim(scannedCode))

    ' Rechercher le produit
    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Set findResult = ws.Columns(1).Find(scannedCode, LookIn:=xlValues, LookAt:=xlWhole)

    If findResult Is Nothing Then
        MsgBox "Produit non trouvé : " & scannedCode, vbExclamation
        Exit Sub
    End If

    ' Afficher les informations du produit
    Dim productInfo As String
    productInfo = "Produit trouvé :" & vbNewLine & vbNewLine
    productInfo = productInfo & "Code : " & ws.Cells(findResult.Row, 1).Value & vbNewLine
    productInfo = productInfo & "Nom : " & ws.Cells(findResult.Row, 2).Value & vbNewLine
    productInfo = productInfo & "Stock actuel : " & ws.Cells(findResult.Row, 4).Value & vbNewLine
    productInfo = productInfo & "Emplacement : " & ws.Cells(findResult.Row, 7).Value & vbNewLine & vbNewLine
    productInfo = productInfo & "Que souhaitez-vous faire ?"

    ' Proposer les actions possibles
    Dim response As VbMsgBoxResult
    response = MsgBox(productInfo, vbQuestion + vbYesNoCancel, "Actions possibles - OUI=Entrée, NON=Sortie, ANNULER=Consulter")

    Select Case response
        Case vbYes  ' Entrée de stock
            actionType = "ENTREE"
            quantityInput = InputBox("Quantité à ajouter :", "Entrée de stock", "1")
        Case vbNo   ' Sortie de stock
            actionType = "SORTIE"
            quantityInput = InputBox("Quantité à retirer :", "Sortie de stock", "1")
        Case vbCancel  ' Consulter uniquement
            MsgBox productInfo, vbInformation, "Consultation produit"
            Exit Sub
    End Select

    ' Valider la quantité
    If Not IsNumeric(quantityInput) Or quantityInput = "" Then
        MsgBox "Quantité invalide.", vbExclamation
        Exit Sub
    End If

    quantity = CDbl(quantityInput)

    If quantity <= 0 Then
        MsgBox "Quantité invalide.", vbExclamation
        Exit Sub
    End If

    ' Traiter le mouvement
    Call ProcessQuickMovement(scannedCode, actionType, quantity)
End Sub

Sub ProcessQuickMovement(productCode As String, movementType As String, quantity As Double)
    '=========================================
    ' Traitement rapide d'un mouvement (pour scanner)
    '=========================================

    Dim ws As Worksheet
    Dim findResult As Range
    Dim currentStock As Double
    Dim newStock As Double

    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Set findResult = ws.Columns(1).Find(productCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not findResult Is Nothing Then
        currentStock = ws.Cells(findResult.Row, 4).Value

        ' Vérifier la faisabilité pour les sorties
        If movementType = "SORTIE" And currentStock < quantity Then
            MsgBox "Stock insuffisant !" & vbNewLine & _
                   "Stock disponible : " & currentStock & vbNewLine & _
                   "Quantité demandée : " & quantity, vbExclamation
            Exit Sub
        End If

        ' Calculer le nouveau stock
        If movementType = "ENTREE" Then
            newStock = currentStock + quantity
        Else
            newStock = currentStock - quantity
        End If

        ' Mettre à jour le stock
        ws.Cells(findResult.Row, 4).Value = newStock

        ' Appliquer le formatage
        Call ApplyStockFormatting(findResult.Row)

        ' Enregistrer le mouvement
        Call RecordMovement(productCode, movementType, quantity, currentStock, "Saisie code-barres")

        ' Message de confirmation
        MsgBox "Mouvement enregistré !" & vbNewLine & _
               "Produit : " & ws.Cells(findResult.Row, 2).Value & vbNewLine & _
               "Ancien stock : " & currentStock & vbNewLine & _
               "Nouveau stock : " & newStock, vbInformation

        ' Actualiser les alertes
        Call UpdateStockAlerts
    End If
End Sub
```

## Maintenance et optimisation

### Archivage automatique des anciens mouvements

```vba
Sub ArchiveOldMovements()
    '=========================================
    ' Archivage des mouvements de plus de 12 mois
    '=========================================

    Dim ws As Worksheet
    Dim archiveSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim archiveDate As Date
    Dim archivedCount As Integer

    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")
    archiveDate = DateAdd("m", -12, Date)  ' 12 mois en arrière
    archivedCount = 0

    ' Créer ou récupérer la feuille d'archive
    On Error Resume Next
    Set archiveSheet = ThisWorkbook.Sheets("MOUVEMENTS_ARCHIVE")
    On Error GoTo 0

    If archiveSheet Is Nothing Then
        Set archiveSheet = ThisWorkbook.Sheets.Add
        archiveSheet.Name = "MOUVEMENTS_ARCHIVE"
        Call SetupMovementSheet(archiveSheet)
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim archiveRow As Long

    ' Parcourir les mouvements du plus ancien au plus récent
    For i = lastRow To 2 Step -1  ' Commencer par la fin pour éviter les problèmes de décalage
        If IsDate(ws.Cells(i, 1).Value) Then
            If ws.Cells(i, 1).Value < archiveDate Then
                ' Copier la ligne vers l'archive
                archiveRow = archiveSheet.Cells(archiveSheet.Rows.Count, 1).End(xlUp).Row + 1

                ws.Rows(i).Copy
                archiveSheet.Rows(archiveRow).PasteSpecial xlPasteAll

                ' Supprimer la ligne originale
                ws.Rows(i).Delete
                archivedCount = archivedCount + 1
            End If
        End If
    Next i

    Application.CutCopyMode = False

    MsgBox archivedCount & " mouvements archivés (plus de 12 mois).", vbInformation
End Sub
```

### Nettoyage et optimisation

```vba
Sub OptimizeDatabase()
    '=========================================
    ' Optimisation de la base de données
    '=========================================

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Dim optimizations As String

    optimizations = "Optimisations effectuées :" & vbNewLine & vbNewLine

    ' 1. Supprimer les lignes vides dans PRODUITS
    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Call RemoveEmptyRows(ws)
    optimizations = optimizations & "✓ Lignes vides supprimées (PRODUITS)" & vbNewLine

    ' 2. Supprimer les lignes vides dans MOUVEMENTS
    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")
    Call RemoveEmptyRows(ws)
    optimizations = optimizations & "✓ Lignes vides supprimées (MOUVEMENTS)" & vbNewLine

    ' 3. Réorganiser les produits par code
    Set ws = ThisWorkbook.Sheets("PRODUITS")
    Call SortProductsByCode(ws)
    optimizations = optimizations & "✓ Produits triés par code" & vbNewLine

    ' 4. Recalculer toutes les formules
    Application.Calculate
    optimizations = optimizations & "✓ Formules recalculées" & vbNewLine

    ' 5. Mettre à jour les alertes
    Call UpdateStockAlerts
    optimizations = optimizations & "✓ Alertes mises à jour" & vbNewLine

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox optimizations & vbNewLine & "Base de données optimisée avec succès !", vbInformation
End Sub

Sub RemoveEmptyRows(ws As Worksheet)
    '=========================================
    ' Suppression des lignes vides
    '=========================================

    Dim lastRow As Long
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 2 Step -1
        If Application.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub SortProductsByCode(ws As Worksheet)
    '=========================================
    ' Tri des produits par code
    '=========================================

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > 2 Then
        ws.Sort.SortFields.Clear
        ws.Sort.SortFields.Add Key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending

        With ws.Sort
            .SetRange ws.Range("A1:I" & lastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If
End Sub
```

## Sécurité et sauvegarde

### Protection des données

```vba
Sub ProtectWorkbook()
    '=========================================
    ' Protection complète du classeur
    '=========================================

    Dim ws As Worksheet
    Dim password As String

    password = "StockSecure123"  ' À modifier selon vos besoins

    ' Protéger chaque feuille
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.ProtectContents Then
            Select Case ws.Name
                Case "PRODUITS", "MOUVEMENTS", "ALERTES"
                    ' Protection avec autorisation d'insertion de lignes
                    ws.Protect Password:=password, AllowInsertingRows:=True, AllowFiltering:=True
                Case "ACCUEIL", "RAPPORTS"
                    ' Protection complète
                    ws.Protect Password:=password
            End Select
        End If
    Next ws

    ' Protéger la structure du classeur
    ThisWorkbook.Protect Password:=password, Structure:=True

    MsgBox "Classeur protégé avec succès !", vbInformation
End Sub

Sub UnprotectWorkbook()
    '=========================================
    ' Déprotection du classeur pour maintenance
    '=========================================

    Dim ws As Worksheet
    Dim password As String

    password = InputBox("Entrez le mot de passe de déprotection :", "Déprotection")

    If password = "" Then Exit Sub

    On Error GoTo PasswordError

    ' Déprotéger le classeur
    ThisWorkbook.Unprotect Password:=password

    ' Déprotéger chaque feuille
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then
            ws.Unprotect Password:=password
        End If
    Next ws

    MsgBox "Classeur déprotégé pour maintenance.", vbInformation
    Exit Sub

PasswordError:
    MsgBox "Mot de passe incorrect.", vbExclamation
End Sub
```

### Sauvegarde automatique

```vba
Sub AutoBackup()
    '=========================================
    ' Sauvegarde automatique quotidienne
    '=========================================

    Dim backupFolder As String
    Dim fileName As String
    Dim fullPath As String

    ' Définir le dossier de sauvegarde
    backupFolder = ThisWorkbook.Path & "\Sauvegardes_Auto"

    ' Créer le dossier s'il n'existe pas
    If Dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If

    ' Nom de fichier avec date
    fileName = "Stock_Auto_" & Format(Now(), "yyyy-mm-dd") & ".xlsm"
    fullPath = backupFolder & "\" & fileName

    ' Vérifier si une sauvegarde du jour existe déjà
    If Dir(fullPath) <> "" Then
        ' Remplacer la sauvegarde existante
        Kill fullPath
    End If

    ' Sauvegarder
    ThisWorkbook.SaveCopyAs fullPath

    ' Nettoyer les anciennes sauvegardes (garder 30 jours)
    Call CleanOldBackups(backupFolder, 30)
End Sub

Sub CleanOldBackups(folderPath As String, daysToKeep As Integer)
    '=========================================
    ' Nettoyage des anciennes sauvegardes
    '=========================================

    Dim fileName As String
    Dim filePath As String
    Dim fileDate As Date
    Dim cutoffDate As Date
    Dim fileCount As Long
    Dim fileNames() As String

    cutoffDate = DateAdd("d", -daysToKeep, Date)

    ' Collecter d'abord tous les noms de fichiers
    ' (Kill entre les appels Dir corrompt l'état interne de Dir)
    fileCount = 0
    fileName = Dir(folderPath & "\*.xlsm")

    Do While fileName <> ""
        fileCount = fileCount + 1
        ReDim Preserve fileNames(1 To fileCount)
        fileNames(fileCount) = fileName
        fileName = Dir()
    Loop

    ' Puis supprimer les fichiers trop anciens
    Dim j As Long
    For j = 1 To fileCount
        filePath = folderPath & "\" & fileNames(j)
        fileDate = FileDateTime(filePath)

        If fileDate < cutoffDate Then
            Kill filePath
        End If
    Next j
End Sub
```

## Conclusion et perspectives d'évolution

### Bilan du projet

Ce système de gestion de stock simple démontre la puissance de VBA pour créer des solutions métier complètes. Avec environ 500 lignes de code, nous avons développé :

**Fonctionnalités principales réalisées :**
- ✅ Gestion complète des produits (création, modification, consultation)
- ✅ Suivi en temps réel des mouvements de stock
- ✅ Système d'alertes automatiques pour les stocks faibles
- ✅ Interface utilisateur intuitive avec UserForms
- ✅ Rapports automatisés et historique complet
- ✅ Fonctionnalités d'import/export et sauvegarde
- ✅ Protection et sécurisation des données

**Avantages obtenus :**
- **Gain de temps** : Automatisation des tâches répétitives
- **Fiabilité** : Réduction des erreurs de saisie
- **Traçabilité** : Historique complet de tous les mouvements
- **Réactivité** : Alertes en temps réel pour les stocks critiques
- **Simplicité** : Interface accessible à tous les utilisateurs

### Évolutions possibles

#### Extensions techniques
1. **Intégration base de données** : Connexion à SQL Server ou Access
2. **Interface web** : Portage vers une solution web avec Office 365
3. **Application mobile** : Développement d'une app pour saisie nomade
4. **API intégration** : Connexion avec systèmes de caisse ou ERP

#### Fonctionnalités supplémentaires
1. **Gestion multi-entrepôts** : Support de plusieurs emplacements
2. **Prévisions de stock** : Algorithmes de prédiction des besoins
3. **Gestion des fournisseurs** : Carnet d'adresses et commandes automatiques
4. **Contrôle qualité** : Suivi des dates de péremption et lots

#### Optimisations avancées
1. **Performance** : Utilisation de tableaux en mémoire pour gros volumes
2. **Sécurité** : Chiffrement des données sensibles
3. **Multi-utilisateurs** : Gestion des accès concurrents
4. **Audit trail** : Journalisation complète des actions utilisateur

### Apprentissages clés

Ce projet illustre parfaitement l'apprentissage progressif de VBA :

**Concepts fondamentaux utilisés :**
- Variables et types de données
- Structures de contrôle (boucles, conditions)
- Procédures et fonctions
- Manipulation d'objets Excel

**Techniques intermédiaires maîtrisées :**
- UserForms et événements
- Gestion d'erreurs robuste
- Manipulation de fichiers
- Formatage conditionnel automatique

**Approches avancées découvertes :**
- Architecture modulaire
- Optimisation des performances
- Sécurisation et protection
- Intégration avec d'autres applications

### Conseils pour aller plus loin

1. **Personnalisation** : Adaptez le système à vos besoins spécifiques
2. **Formation utilisateurs** : Documentez et formez vos équipes
3. **Maintenance régulière** : Planifiez des mises à jour et optimisations
4. **Feedback continu** : Écoutez les utilisateurs pour améliorer l'outil
5. **Évolution graduelle** : Ajoutez progressivement de nouvelles fonctionnalités

Ce système de gestion de stock constitue un excellent tremplin vers des projets VBA plus ambitieux. Il démontre comment transformer des connaissances techniques en solutions pratiques qui apportent une réelle valeur ajoutée au quotidien professionnel.

**Prochaine étape recommandée :** Implémenter ce système dans votre environnement de travail et l'adapter à vos besoins spécifiques. L'expérience pratique reste le meilleur moyen de progresser en développement VBA !

⏭️ [Outil d'analyse de données](/22-projets-pratiques/03-outil-analyse-donnees.md)
