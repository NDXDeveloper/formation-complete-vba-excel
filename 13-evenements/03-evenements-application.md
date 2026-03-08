🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.3. Événements d'application

## Que sont les Événements d'Application ?

Les événements d'application sont des événements qui se déclenchent automatiquement lors d'actions effectuées au niveau de l'application Excel elle-même, pas seulement sur un classeur ou une feuille spécifique. Ils permettent de surveiller et de réagir à tout ce qui se passe dans Excel, peu importe le fichier ouvert.

**Caractéristiques importantes :**
- Ils concernent l'application Excel dans son ensemble
- Ils fonctionnent même quand vous changez de classeur
- Ils nécessitent une configuration spéciale avec `WithEvents`
- Ils sont très puissants pour créer des outils globaux

## Différence avec les autres événements

| Type d'événement | Portée | Exemple |
|------------------|--------|---------|
| **Feuille** | Une feuille spécifique | Modification d'une cellule dans Feuil1 |
| **Classeur** | Un classeur spécifique | Ouverture du fichier "Ventes.xlsx" |
| **Application** | Toute l'application Excel | Ouverture de n'importe quel classeur |

## Configuration des Événements d'Application

### Étape 1 : Créer un module de classe

Les événements d'application nécessitent une configuration spéciale :

1. **Insérer un module de classe** :
   - Dans l'éditeur VBA (`Alt + F11`)
   - Clic droit dans l'explorateur de projets
   - `Insertion` → `Module de classe`
   - Renommer en "ClsAppEvents" (ou un nom explicite)

2. **Déclarer l'objet Application avec WithEvents** :
```vba
' Dans le module de classe ClsAppEvents
Public WithEvents xlApp As Application
```

### Étape 2 : Créer une variable globale

Dans un module standard, déclarez une variable pour contenir l'instance :

```vba
' Dans un module standard
Public AppEvents As ClsAppEvents
```

### Étape 3 : Initialiser les événements

```vba
' Dans un module standard ou ThisWorkbook
Sub InitialiserEvenementsApp()
    Set AppEvents = New ClsAppEvents
    Set AppEvents.xlApp = Application
End Sub

' Pour arrêter la surveillance
Sub ArreterEvenementsApp()
    Set AppEvents = Nothing
End Sub
```

## Événements d'Application les plus Utiles

### 1. NewWorkbook - Nouveau Classeur

Se déclenche quand un nouveau classeur est créé :

```vba
' Dans le module de classe ClsAppEvents
Private Sub xlApp_NewWorkbook(ByVal Wb As Workbook)
    MsgBox "Nouveau classeur créé : " & Wb.Name

    ' Personnaliser le nouveau classeur
    Wb.Sheets(1).Range("A1").Value = "Créé le " & Now()
End Sub
```

### 2. WorkbookOpen - Ouverture de Classeur

Se déclenche à l'ouverture de n'importe quel classeur :

```vba
Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    ' Journal de tous les fichiers ouverts
    Debug.Print "Fichier ouvert : " & Wb.FullName & " à " & Now()

    ' Vérifier la sécurité
    If InStr(Wb.FullName, "Temp") > 0 Then
        MsgBox "Attention : Fichier provenant d'un dossier temporaire !", vbExclamation
    End If
End Sub
```

### 3. WorkbookBeforeClose - Avant Fermeture

Se déclenche avant la fermeture de n'importe quel classeur :

```vba
Private Sub xlApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    ' Sauvegarder automatiquement les fichiers non sauvegardés
    If Not Wb.Saved And Wb.Path <> "" Then
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Sauvegarder " & Wb.Name & " ?", vbYesNoCancel)

        Select Case reponse
            Case vbYes
                Wb.Save
            Case vbCancel
                Cancel = True  ' Annuler la fermeture
        End Select
    End If
End Sub
```

### 4. WorkbookBeforeSave - Avant Sauvegarde

Se déclenche avant la sauvegarde de n'importe quel classeur :

```vba
Private Sub xlApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Ajouter automatiquement un horodatage
    On Error Resume Next
    Wb.Sheets(1).Range("Z1").Value = "Dernière sauvegarde : " & Now()
    On Error GoTo 0

    ' Alerter pour les gros fichiers
    If Wb.Sheets.Count > 10 Then
        MsgBox "Attention : Ce classeur contient " & Wb.Sheets.Count & " feuilles.", vbInformation
    End If
End Sub
```

### 5. SheetChange - Modification dans n'importe quelle feuille

Se déclenche lors de modifications dans toute feuille de tout classeur :

```vba
Private Sub xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Journaliser toutes les modifications
    Debug.Print "Modification dans " & Sh.Parent.Name & " - " & Sh.Name & _
                " cellule " & Target.Address & " = " & Target.Value
End Sub
```

### 6. SheetSelectionChange - Changement de sélection global

Se déclenche à chaque changement de sélection dans Excel :

```vba
Private Sub xlApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' Afficher la position dans la barre d'état
    Application.StatusBar = "Fichier: " & Sh.Parent.Name & " | Feuille: " & Sh.Name & _
                           " | Sélection: " & Target.Address
End Sub
```

### 7. WorkbookActivate - Activation de Classeur

Se déclenche quand l'utilisateur passe d'un classeur à un autre :

```vba
Private Sub xlApp_WorkbookActivate(ByVal Wb As Workbook)
    ' Personnaliser l'interface selon le classeur
    If InStr(Wb.Name, "Budget") > 0 Then
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
    Else
        Application.DisplayFormulaBar = False
    End If
End Sub
```

### 8. WindowResize - Redimensionnement de Fenêtre

Se déclenche quand une fenêtre Excel est redimensionnée :

```vba
Private Sub xlApp_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    ' Ajuster le zoom selon la taille de la fenêtre
    If Wn.Width < 400 Then
        Wn.Zoom = 75  ' Petit écran
    ElseIf Wn.Width > 800 Then
        Wn.Zoom = 120  ' Grand écran
    Else
        Wn.Zoom = 100  ' Taille normale
    End If
End Sub
```

## Exemple Complet : Système de Surveillance Global

Voici un exemple complet qui combine plusieurs événements d'application :

### Module de classe (ClsAppEvents) :

```vba
' Module de classe : ClsAppEvents
Public WithEvents xlApp As Application

Private Sub xlApp_NewWorkbook(ByVal Wb As Workbook)
    ' Personnaliser les nouveaux classeurs
    With Wb.Sheets(1)
        .Range("A1").Value = "Nouveau document"
        .Range("A2").Value = "Créé le : " & Format(Now(), "dd/mm/yyyy à hh:mm")
        .Range("A3").Value = "Utilisateur : " & Environ("USERNAME")
        .Range("A1:A3").Font.Bold = True
    End With

    MsgBox "Nouveau classeur configuré automatiquement !", vbInformation
End Sub

Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    ' Journal d'ouverture
    Call EnregistrerActivite("OUVERTURE", Wb.Name)

    ' Vérification de sécurité
    If Wb.HasVBProject Then
        If Wb.VBProject.Protection <> 1 Then
            MsgBox "Attention : Ce fichier contient des macros non protégées !", vbExclamation
        End If
    End If
End Sub

Private Sub xlApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    ' Journal de fermeture
    Call EnregistrerActivite("FERMETURE", Wb.Name)

    ' Nettoyage automatique
    If InStr(Wb.Name, "Temp") > 0 Then
        If MsgBox("Supprimer ce fichier temporaire ?", vbYesNo) = vbYes Then
            ' Marquer pour suppression (à implémenter)
        End If
    End If
End Sub

Private Sub xlApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Vérifications avant sauvegarde
    If Wb.Sheets.Count = 1 And Wb.Sheets(1).UsedRange.Cells.Count = 1 Then
        If MsgBox("Ce classeur semble vide. Continuer la sauvegarde ?", vbYesNo) = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Compteur de modifications (simple)
    Static compteurModifs As Long
    compteurModifs = compteurModifs + 1

    ' Afficher le compteur dans la barre d'état (modulo 10 pour ne pas surcharger)
    If compteurModifs Mod 10 = 0 Then
        Application.StatusBar = "Modifications : " & compteurModifs
    End If
End Sub

' Procédure utilitaire
Private Sub EnregistrerActivite(action As String, nomFichier As String)
    ' Enregistrer dans un fichier journal (version simplifiée)
    Debug.Print Format(Now(), "dd/mm/yyyy hh:mm:ss") & " - " & action & " - " & nomFichier

    ' Dans un vrai projet, on pourrait écrire dans un fichier log
    ' ou une base de données
End Sub
```

### Module standard :

```vba
' Module standard
Public AppEvents As ClsAppEvents

' Démarrer la surveillance
Sub DemarrerSurveillance()
    If AppEvents Is Nothing Then
        Set AppEvents = New ClsAppEvents
        Set AppEvents.xlApp = Application
        MsgBox "Surveillance d'application activée !", vbInformation
    Else
        MsgBox "La surveillance est déjà active.", vbInformation
    End If
End Sub

' Arrêter la surveillance
Sub ArreterSurveillance()
    Set AppEvents = Nothing
    Application.StatusBar = False  ' Effacer la barre d'état
    MsgBox "Surveillance d'application désactivée.", vbInformation
End Sub

' Vérifier l'état de la surveillance
Sub EtatSurveillance()
    If AppEvents Is Nothing Then
        MsgBox "Surveillance : INACTIVE", vbExclamation
    Else
        MsgBox "Surveillance : ACTIVE", vbInformation
    End If
End Sub
```

### Initialisation automatique dans ThisWorkbook :

```vba
' Dans ThisWorkbook
Private Sub Workbook_Open()
    ' Démarrer automatiquement la surveillance des événements d'application
    Call DemarrerSurveillance
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Arrêter la surveillance avant fermeture
    Call ArreterSurveillance
End Sub
```

## Gestion des Erreurs dans les Événements d'Application

```vba
Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    On Error GoTo GestionErreur

    ' Code potentiellement problématique
    Wb.Sheets(1).Range("A1").Value = "Test"

    Exit Sub

GestionErreur:
    Debug.Print "Erreur dans WorkbookOpen : " & Err.Description
    ' Ne pas afficher de MsgBox qui pourrait bloquer Excel
    ' Préférer Debug.Print ou l'écriture dans un fichier log
End Sub
```

## Optimisation et Performance

### Conseils d'optimisation :

1. **Éviter les MsgBox fréquents** dans les événements qui se déclenchent souvent
2. **Utiliser Debug.Print** pour le débogage plutôt que MsgBox
3. **Désactiver temporairement** la surveillance si nécessaire :

```vba
Sub CodeSansEvenements()
    ' Arrêter temporairement la surveillance
    Call ArreterSurveillance

    ' Code qui génère beaucoup d'événements
    For i = 1 To 1000
        Workbooks.Add
        ActiveWorkbook.Close
    Next i

    ' Redémarrer la surveillance
    Call DemarrerSurveillance
End Sub
```

## Applications Pratiques des Événements d'Application

### 1. Système de Backup Automatique

```vba
Private Sub xlApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Créer une copie de sauvegarde
    If Wb.Path <> "" Then  ' Fichier déjà sauvegardé au moins une fois
        Dim cheminBackup As String
        cheminBackup = Wb.Path & "\Backup_" & Format(Now(), "yyyymmdd_hhmmss") & "_" & Wb.Name

        On Error Resume Next
        Wb.SaveCopyAs cheminBackup
        On Error GoTo 0
    End If
End Sub
```

### 2. Surveillance de Sécurité

```vba
Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    ' Vérifier l'origine du fichier
    If InStr(Wb.FullName, "Downloads") > 0 Or InStr(Wb.FullName, "Téléchargements") > 0 Then
        MsgBox "Attention : Fichier provenant du dossier de téléchargements !" & vbCrLf & _
               "Vérifiez sa source avant utilisation.", vbExclamation
    End If
End Sub
```

### 3. Statistiques d'Utilisation

```vba
' Variable au niveau du module de classe (pas dans la procédure !)
Private dictionnaire As Object

Private Sub Class_Initialize()
    Set dictionnaire = CreateObject("Scripting.Dictionary")
End Sub

Private Sub xlApp_WorkbookActivate(ByVal Wb As Workbook)
    ' Incrémenter le compteur pour ce fichier
    If dictionnaire.Exists(Wb.Name) Then
        dictionnaire(Wb.Name) = dictionnaire(Wb.Name) + 1
    Else
        dictionnaire(Wb.Name) = 1
    End If

    Debug.Print Wb.Name & " activé " & dictionnaire(Wb.Name) & " fois"
End Sub
```

## Bonnes Pratiques

### ✅ À faire :
- **Initialiser les événements** dans `Workbook_Open`
- **Nettoyer les références** dans `Workbook_BeforeClose`
- **Gérer les erreurs** pour éviter de planter Excel
- **Utiliser Debug.Print** pour les logs plutôt que MsgBox
- **Tester les performances** avec les événements fréquents

### ❌ À éviter :
- **MsgBox répétitifs** dans les événements fréquents (SheetChange, SelectionChange)
- **Code lourd** qui ralentit toute l'application Excel
- **Oublier de nettoyer** les références aux objets WithEvents
- **Boucles infinies** entre événements
- **Modifications massives** sans désactiver temporairement les événements

## Débogage des Événements d'Application

```vba
' Ajouter du logging pour déboguer
Private Sub xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Debug.Print "DEBUG - SheetChange: " & Sh.Parent.Name & "." & Sh.Name & _
                " - " & Target.Address & " = " & Target.Value
End Sub

' Ou créer un fichier log
Private Sub EcrireLog(message As String)
    Open "C:\Temp\ExcelEvents.log" For Append As #1
    Print #1, Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & message
    Close #1
End Sub
```

## Résumé

Les événements d'application sont un outil puissant pour :
- **Surveiller globalement** l'activité dans Excel
- **Automatiser des tâches** qui concernent tous les fichiers
- **Créer des outils transversaux** qui fonctionnent peu importe le classeur
- **Implémenter des systèmes de sécurité** et de surveillance

Ils nécessitent une configuration plus complexe que les autres événements, mais offrent une puissance et une flexibilité exceptionnelles. Dans la section suivante, nous découvrirons comment créer nos propres événements personnalisés.

⏭️ [Création d'événements personnalisés](/13-evenements/04-creation-evenements-personnalises.md)
