🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.1. Événements de classeur (Workbook_Open, Before_Close)

## Que sont les Événements de Classeur ?

Les événements de classeur sont des événements qui se déclenchent automatiquement lors d'actions effectuées sur un classeur Excel entier. Ils permettent d'exécuter du code VBA automatiquement lorsque certaines situations se présentent, comme l'ouverture ou la fermeture d'un fichier.

**Caractéristiques importantes :**
- Ils concernent le classeur dans son ensemble (pas une feuille spécifique)
- Le code doit être placé dans le module **ThisWorkbook**
- Ils s'exécutent automatiquement sans intervention de l'utilisateur

## Comment accéder aux Événements de Classeur

### Étapes pour créer un événement de classeur :

1. **Ouvrir l'éditeur VBA** : Appuyez sur `Alt + F11`
2. **Localiser ThisWorkbook** : Dans l'explorateur de projets (à gauche), double-cliquez sur "ThisWorkbook"
3. **Sélectionner l'objet Workbook** : Dans la fenêtre de code, cliquez sur la liste déroulante en haut à gauche et choisissez "Workbook"
4. **Choisir l'événement** : Dans la liste déroulante en haut à droite, sélectionnez l'événement souhaité

Une procédure vide se créera automatiquement avec la structure correcte.

## Événement Workbook_Open

### Qu'est-ce que c'est ?
L'événement `Workbook_Open` se déclenche automatiquement **chaque fois que le classeur est ouvert**.

### Syntaxe
```vba
Private Sub Workbook_Open()
    ' Votre code ici
End Sub
```

### Utilisations courantes

**1. Message de bienvenue**
```vba
Private Sub Workbook_Open()
    MsgBox "Bienvenue dans le système de gestion !" & vbCrLf & _
           "Version 2.1 - Dernière mise à jour : Mars 2024"
End Sub
```

**2. Initialisation de l'environnement**
```vba
Private Sub Workbook_Open()
    ' Masquer les onglets de feuilles
    ActiveWindow.DisplayWorkbookTabs = False

    ' Masquer la grille
    ActiveWindow.DisplayGridlines = False

    ' Aller à une feuille spécifique
    Sheets("Accueil").Select
End Sub
```

**3. Vérifications de sécurité**
```vba
Private Sub Workbook_Open()
    ' Vérifier si l'utilisateur a les droits
    If Environ("USERNAME") <> "AdminUser" Then
        MsgBox "Accès non autorisé !"
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub
```

**4. Mise à jour de données**
```vba
Private Sub Workbook_Open()
    ' Mettre à jour la date d'ouverture
    Sheets("Journal").Range("A1").Value = "Dernière ouverture : " & Now()

    ' Actualiser les données externes
    ThisWorkbook.RefreshAll
End Sub
```

### Points importants à retenir

- **Exécution unique** : Se déclenche une seule fois à l'ouverture
- **Avant tout le reste** : S'exécute avant que l'utilisateur puisse interagir
- **Attention aux erreurs** : Une erreur peut empêcher l'ouverture normale du fichier
- **Macros désactivées** : Ne fonctionne pas si l'utilisateur refuse les macros

## Événement Workbook_BeforeClose

### Qu'est-ce que c'est ?
L'événement `Workbook_BeforeClose` se déclenche automatiquement **juste avant que le classeur ne se ferme**, mais permet d'annuler la fermeture si nécessaire.

### Syntaxe
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Votre code ici
    ' Cancel = True  ' Pour annuler la fermeture
End Sub
```

### Le paramètre Cancel

Le paramètre `Cancel` est très important :
- **Cancel = False** (défaut) : La fermeture continue normalement
- **Cancel = True** : La fermeture est annulée, le classeur reste ouvert

### Utilisations courantes

**1. Demande de confirmation**
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim reponse As VbMsgBoxResult

    reponse = MsgBox("Êtes-vous sûr de vouloir fermer le fichier ?", _
                     vbYesNo + vbQuestion, "Confirmation")

    If reponse = vbNo Then
        Cancel = True  ' Annule la fermeture
    End If
End Sub
```

**2. Sauvegarde automatique**
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Sauvegarder automatiquement si des modifications existent
    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
        MsgBox "Le fichier a été sauvegardé automatiquement."
    End If
End Sub
```

**3. Nettoyage et archivage**
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Effacer les données temporaires
    Sheets("Temp").Range("A:Z").Clear

    ' Enregistrer la date de fermeture
    Sheets("Journal").Range("B1").Value = "Fermé le : " & Now()

    ' Restaurer les paramètres Excel
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayGridlines = True
End Sub
```

**4. Vérifications obligatoires**
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Vérifier que tous les champs obligatoires sont remplis
    If Sheets("Données").Range("A1").Value = "" Then
        MsgBox "Attention : Le champ nom ne peut pas être vide !" & vbCrLf & _
               "Fermeture annulée.", vbExclamation
        Cancel = True
    End If
End Sub
```

## Autres Événements de Classeur Utiles

### Workbook_BeforeSave
Se déclenche avant la sauvegarde :
```vba
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Code exécuté avant chaque sauvegarde
    Sheets("Journal").Range("C1").Value = "Sauvegardé le : " & Now()
End Sub
```

### Workbook_AfterSave
Se déclenche après la sauvegarde :
```vba
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Success Then
        MsgBox "Sauvegarde réussie !"
    Else
        MsgBox "Erreur lors de la sauvegarde !"
    End If
End Sub
```

### Workbook_NewSheet
Se déclenche quand une nouvelle feuille est créée :
```vba
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    MsgBox "Nouvelle feuille créée : " & Sh.Name
End Sub
```

## Combinaison d'Événements - Exemple Complet

Voici un exemple réaliste combinant plusieurs événements :

```vba
' Dans le module ThisWorkbook

Private Sub Workbook_Open()
    ' Message de bienvenue avec informations
    MsgBox "Bienvenue " & Environ("USERNAME") & " !" & vbCrLf & _
           "Fichier ouvert le " & Format(Now(), "dd/mm/yyyy à hh:mm")

    ' Aller à la feuille d'accueil
    Sheets("Accueil").Select

    ' Enregistrer l'ouverture dans un journal
    Sheets("Journal").Range("A" & Sheets("Journal").Rows.Count).End(xlUp).Offset(1, 0).Value = _
        "Ouverture : " & Now() & " - " & Environ("USERNAME")
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Nettoyer les données temporaires
    On Error Resume Next
    Sheets("Temp").UsedRange.Clear
    On Error GoTo 0

    ' Enregistrer la fermeture
    Sheets("Journal").Range("A" & Sheets("Journal").Rows.Count).End(xlUp).Offset(1, 0).Value = _
        "Fermeture : " & Now() & " - " & Environ("USERNAME")

    ' Sauvegarder si nécessaire
    If Not ThisWorkbook.Saved Then
        ThisWorkbook.Save
    End If

    ' Message de fin
    MsgBox "À bientôt ! Merci d'avoir utilisé l'application."
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Mettre à jour la date de dernière modification
    Sheets("Informations").Range("B2").Value = "Dernière modification : " & Now()
End Sub
```

## Gestion d'Erreurs dans les Événements

Il est crucial de gérer les erreurs dans les événements de classeur :

```vba
Private Sub Workbook_Open()
    On Error GoTo GestionErreur

    ' Votre code ici
    MsgBox "Initialisation terminée avec succès"

    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'initialisation : " & Err.Description, vbCritical
    ' Ne pas bloquer l'ouverture du fichier
End Sub
```

## Bonnes Pratiques

### ✅ À faire :
- **Garder le code simple** dans `Workbook_Open` pour éviter de ralentir l'ouverture
- **Toujours inclure une gestion d'erreur** pour éviter de bloquer Excel
- **Informer l'utilisateur** des actions automatiques importantes
- **Sauvegarder avant de fermer** si des données importantes ont été modifiées

### ❌ À éviter :
- **Code trop long** dans `Workbook_Open` qui ralentit l'ouverture
- **Boucles infinies** ou code qui ne finit jamais
- **MsgBox multiples** qui agacent l'utilisateur
- **Modifications importantes** sans demander confirmation

## Désactivation Temporaire des Événements

Parfois, vous devrez désactiver temporairement les événements :

```vba
Sub MonCodeSansEvenements()
    Application.EnableEvents = False

    ' Code qui pourrait déclencher des événements
    ThisWorkbook.Save

    Application.EnableEvents = True  ' IMPORTANT : toujours réactiver !
End Sub
```

## Résumé

Les événements de classeur sont des outils puissants pour :
- **Automatiser l'initialisation** de votre application
- **Contrôler la fermeture** et s'assurer de la sauvegarde
- **Maintenir un journal** des activités
- **Créer une expérience utilisateur** personnalisée

Ils constituent la base d'applications Excel professionnelles et interactives. Dans la section suivante, nous découvrirons les événements de feuille qui permettent de réagir aux modifications des données en temps réel.

⏭️
