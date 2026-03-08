🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.2. Événements de feuille (Change, SelectionChange)

## Que sont les Événements de Feuille ?

Les événements de feuille sont des événements qui se déclenchent automatiquement lors d'actions effectuées sur une feuille de calcul spécifique. Contrairement aux événements de classeur qui concernent le fichier entier, les événements de feuille sont liés à une seule feuille et permettent de réagir aux interactions de l'utilisateur avec les cellules.

**Caractéristiques importantes :**
- Ils concernent une feuille spécifique uniquement
- Le code doit être placé dans le module de la feuille concernée
- Ils sont très utiles pour la validation de données en temps réel
- Ils permettent de créer des interfaces interactives

## Comment accéder aux Événements de Feuille

### Étapes pour créer un événement de feuille :

1. **Ouvrir l'éditeur VBA** : Appuyez sur `Alt + F11`
2. **Localiser la feuille** : Dans l'explorateur de projets (à gauche), double-cliquez sur la feuille concernée (ex: "Feuil1 (Feuil1)")
3. **Sélectionner l'objet Worksheet** : Dans la fenêtre de code, cliquez sur la liste déroulante en haut à gauche et choisissez "Worksheet"
4. **Choisir l'événement** : Dans la liste déroulante en haut à droite, sélectionnez l'événement souhaité

Une procédure vide se créera automatiquement avec la structure correcte.

## Événement Worksheet_Change

### Qu'est-ce que c'est ?
L'événement `Worksheet_Change` se déclenche automatiquement **chaque fois qu'une cellule de la feuille est modifiée** par l'utilisateur.

### Syntaxe
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Votre code ici
    ' Target représente la ou les cellules modifiées
End Sub
```

### Le paramètre Target

Le paramètre `Target` est un objet Range qui contient :
- La cellule modifiée (si une seule cellule)
- La plage de cellules modifiées (si plusieurs cellules)

**Important** : `Target` contient la nouvelle valeur, pas l'ancienne.

### Exemples d'utilisation de Worksheet_Change

**1. Afficher quelle cellule a été modifiée**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    MsgBox "La cellule " & Target.Address & " a été modifiée." & vbCrLf & _
           "Nouvelle valeur : " & Target.Value
End Sub
```

**2. Validation automatique de données**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Vérifier si la modification concerne la colonne A
    If Target.Column = 1 Then
        ' Vérifier si la valeur est un nombre positif
        If Not IsNumeric(Target.Value) Or Target.Value < 0 Then
            MsgBox "Erreur : Veuillez saisir un nombre positif !", vbExclamation

            ' Désactiver les événements pour éviter une boucle infinie
            Application.EnableEvents = False
            Target.Value = ""  ' Effacer la valeur incorrecte
            Application.EnableEvents = True
        End If
    End If
End Sub
```

**3. Mise à jour automatique de calculs**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Si une valeur change dans la colonne B (quantité)
    If Target.Column = 2 And Target.Row >= 2 Then
        ' Calculer automatiquement le total dans la colonne D
        Target.Offset(0, 2).Value = Target.Value * Range("C" & Target.Row).Value

        ' Mettre à jour la date de modification
        Target.Offset(0, 3).Value = Now()
    End If
End Sub
```

**4. Mise en forme conditionnelle personnalisée**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Si la modification concerne les cellules A1:A10
    If Not Intersect(Target, Range("A1:A10")) Is Nothing Then
        ' Colorer en rouge si la valeur est supérieure à 100
        If Target.Value > 100 Then
            Target.Interior.Color = RGB(255, 200, 200)  ' Rouge clair
        Else
            Target.Interior.Color = RGB(200, 255, 200)  ' Vert clair
        End If
    End If
End Sub
```

**5. Journal des modifications**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Enregistrer toutes les modifications dans une feuille "Journal"
    Dim ws As Worksheet
    Set ws = Sheets("Journal")

    Dim derniereLigne As Long
    derniereLigne = ws.Range("A" & Rows.Count).End(xlUp).Row + 1

    ' Enregistrer la modification
    ws.Range("A" & derniereLigne).Value = Now()                    ' Date/heure
    ws.Range("B" & derniereLigne).Value = Target.Address           ' Cellule
    ws.Range("C" & derniereLigne).Value = Target.Value             ' Nouvelle valeur
    ws.Range("D" & derniereLigne).Value = Environ("USERNAME")      ' Utilisateur
End Sub
```

### Gérer plusieurs cellules modifiées

Quand l'utilisateur modifie plusieurs cellules en même temps (copier-coller, par exemple), `Target` contient toute la plage :

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cellule As Range

    ' Parcourir chaque cellule modifiée
    For Each cellule In Target
        If cellule.Value <> "" Then
            MsgBox "Cellule " & cellule.Address & " = " & cellule.Value
        End If
    Next cellule
End Sub
```

## Événement Worksheet_SelectionChange

### Qu'est-ce que c'est ?
L'événement `Worksheet_SelectionChange` se déclenche automatiquement **chaque fois que l'utilisateur sélectionne une cellule ou une plage de cellules différente**.

### Syntaxe
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Votre code ici
    ' Target représente la nouvelle sélection
End Sub
```

### Exemples d'utilisation de Worksheet_SelectionChange

**1. Afficher des informations sur la sélection**
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Afficher l'adresse et la valeur dans la barre d'état
    Application.StatusBar = "Sélection : " & Target.Address & _
                           " | Valeur : " & Target.Value
End Sub
```

**2. Mise en surbrillance de ligne et colonne**
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Effacer la mise en forme précédente
    Cells.Interior.Color = xlNone

    ' Surligner la ligne et la colonne de la cellule active
    Target.EntireRow.Interior.Color = RGB(220, 220, 220)    ' Gris clair
    Target.EntireColumn.Interior.Color = RGB(220, 220, 220) ' Gris clair

    ' Surligner la cellule active plus fortement
    Target.Interior.Color = RGB(255, 255, 0)  ' Jaune
End Sub
```

**3. Affichage conditionnel d'aide**
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Afficher de l'aide selon la cellule sélectionnée
    Select Case Target.Address
        Case "$A$1"
            Range("E1").Value = "Aide : Saisissez votre nom ici"
        Case "$A$2"
            Range("E1").Value = "Aide : Saisissez votre âge (nombre)"
        Case "$A$3"
            Range("E1").Value = "Aide : Saisissez votre ville"
        Case Else
            Range("E1").Value = ""
    End Select
End Sub
```

**4. Navigation assistée**
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Si l'utilisateur sélectionne une cellule de la première ligne
    If Target.Row = 1 And Target.Column <= 5 Then
        ' Aller automatiquement à la première cellule vide de cette colonne
        Dim premierVide As Range
        Set premierVide = Target.EntireColumn.Find("", Target.Offset(1, 0))

        If Not premierVide Is Nothing Then
            premierVide.Select
        End If
    End If
End Sub
```

**5. Zoom automatique selon la zone**
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Ajuster le zoom selon la zone sélectionnée
    If Target.Row <= 10 Then
        ActiveWindow.Zoom = 120  ' Zone d'en-tête : zoom plus grand
    ElseIf Target.Row <= 50 Then
        ActiveWindow.Zoom = 100  ' Zone de données : zoom normal
    Else
        ActiveWindow.Zoom = 80   ' Zone de calculs : zoom plus petit
    End If
End Sub
```

## Différences importantes entre Change et SelectionChange

| Aspect | Worksheet_Change | Worksheet_SelectionChange |
|--------|------------------|--------------------------|
| **Déclencheur** | Modification de valeur | Changement de sélection |
| **Fréquence** | Moins fréquent | Très fréquent |
| **Usage typique** | Validation, calculs | Interface, navigation |
| **Performance** | Modéré | Attention requise |

## Éviter les Boucles Infinies

### Problème fréquent avec Worksheet_Change
```vba
' ATTENTION : Ce code crée une boucle infinie !
Private Sub Worksheet_Change(ByVal Target As Range)
    Target.Value = Target.Value * 2  ' Modifie la cellule → redéclenche l'événement !
End Sub
```

### Solutions

**1. Désactiver temporairement les événements**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False  ' Désactiver les événements

    Target.Value = Target.Value * 2   ' Modification sans redéclencher l'événement

    Application.EnableEvents = True   ' Réactiver les événements
End Sub
```

**2. Utiliser une condition de sortie**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Éviter de traiter si la valeur a déjà été modifiée par le code
    If Target.Comment Is Nothing Then
        Target.Value = Target.Value * 2
        Target.AddComment "Modifié par VBA"  ' Marquer comme traité
    End If
End Sub
```

## Optimisation des Événements Fréquents

### Pour Worksheet_SelectionChange (très fréquent)

```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Sortir rapidement si pas dans la zone d'intérêt
    If Target.Row > 100 Or Target.Column > 10 Then Exit Sub

    ' Code optimisé ici
    Application.StatusBar = "Ligne " & Target.Row & ", Colonne " & Target.Column
End Sub
```

## Gestion d'Erreurs dans les Événements de Feuille

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo GestionErreur

    ' Votre code ici
    If IsNumeric(Target.Value) Then
        Target.Offset(0, 1).Value = Target.Value * 1.2
    End If

    Exit Sub

GestionErreur:
    MsgBox "Erreur dans l'événement Change : " & Err.Description, vbCritical
    Application.EnableEvents = True  ' S'assurer que les événements restent actifs
End Sub
```

## Exemple Complet : Feuille de Saisie Interactive

Voici un exemple qui combine plusieurs événements pour créer une feuille interactive :

```vba
' Dans le module de la feuille

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo GestionErreur

    ' Mise en surbrillance subtile de la ligne active
    If Target.Row >= 2 And Target.Row <= 20 Then
        ' Effacer la mise en forme précédente
        Range("A2:E20").Interior.Color = xlNone

        ' Surligner la ligne active
        Range("A" & Target.Row & ":E" & Target.Row).Interior.Color = RGB(240, 248, 255)
    End If

    ' Afficher de l'aide contextuelle
    Select Case Target.Column
        Case 1  ' Colonne A
            Range("G1").Value = "Saisissez le nom du produit"
        Case 2  ' Colonne B
            Range("G1").Value = "Saisissez la quantité (nombre entier)"
        Case 3  ' Colonne C
            Range("G1").Value = "Saisissez le prix unitaire"
        Case Else
            Range("G1").Value = ""
    End Select

    Exit Sub

GestionErreur:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo GestionErreur

    Application.EnableEvents = False

    ' Validation et calculs automatiques pour les lignes de données
    If Target.Row >= 2 And Target.Row <= 20 Then

        Select Case Target.Column
            Case 1  ' Colonne A - Nom du produit
                If Len(Target.Value) < 2 Then
                    MsgBox "Le nom du produit doit contenir au moins 2 caractères", vbExclamation
                    Target.Value = ""
                Else
                    ' Mettre en forme le texte
                    Target.Value = UCase(Left(Target.Value, 1)) & LCase(Mid(Target.Value, 2))
                End If

            Case 2  ' Colonne B - Quantité
                If Not IsNumeric(Target.Value) Or Target.Value <= 0 Then
                    MsgBox "La quantité doit être un nombre positif", vbExclamation
                    Target.Value = ""
                Else
                    ' Arrondir à l'entier
                    Target.Value = Int(Target.Value)
                    ' Calculer le total si le prix existe
                    If IsNumeric(Target.Offset(0, 1).Value) Then
                        Target.Offset(0, 2).Value = Target.Value * Target.Offset(0, 1).Value
                    End If
                End If

            Case 3  ' Colonne C - Prix unitaire
                If Not IsNumeric(Target.Value) Or Target.Value <= 0 Then
                    MsgBox "Le prix doit être un nombre positif", vbExclamation
                    Target.Value = ""
                Else
                    ' Formater en monétaire
                    Target.NumberFormat = "0.00 €"
                    ' Calculer le total si la quantité existe
                    If IsNumeric(Target.Offset(0, -1).Value) Then
                        Target.Offset(0, 1).Value = Target.Value * Target.Offset(0, -1).Value
                        Target.Offset(0, 1).NumberFormat = "0.00 €"
                    End If
                End If
        End Select

        ' Mettre à jour la date de modification
        Target.Offset(0, 4).Value = Now()
        Target.Offset(0, 4).NumberFormat = "dd/mm/yyyy hh:mm"
    End If

    Application.EnableEvents = True
    Exit Sub

GestionErreur:
    Application.EnableEvents = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub
```

## Bonnes Pratiques

### ✅ À faire :
- **Toujours inclure une gestion d'erreur** pour éviter de bloquer les événements
- **Utiliser `Application.EnableEvents = False/True`** lors de modifications par code
- **Optimiser le code** dans `SelectionChange` car il s'exécute très souvent
- **Vérifier la zone concernée** avant d'exécuter le code
- **Donner des retours à l'utilisateur** pour les validations

### ❌ À éviter :
- **Code trop lourd** dans `SelectionChange` qui ralentit la navigation
- **Modifications directes** de `Target` dans `Change` sans désactiver les événements
- **MsgBox dans `SelectionChange`** qui interrompt constamment l'utilisateur
- **Oublier de réactiver les événements** après `Application.EnableEvents = False`

## Désactivation et Réactivation des Événements

```vba
' Désactiver temporairement tous les événements
Sub DesactiverEvenements()
    Application.EnableEvents = False
End Sub

' Réactiver tous les événements
Sub ReactiverEvenements()
    Application.EnableEvents = True
End Sub

' Code type avec protection
Sub MonCodeAvecProtection()
    Application.EnableEvents = False

    On Error GoTo Nettoyage

    ' Votre code ici
    Range("A1").Value = "Test"

Nettoyage:
    Application.EnableEvents = True
End Sub
```

## Résumé

Les événements de feuille sont essentiels pour créer des interfaces Excel interactives et réactives :

- **Worksheet_Change** : Parfait pour la validation de données et les calculs automatiques
- **Worksheet_SelectionChange** : Idéal pour l'aide contextuelle et la navigation assistée

Ces événements, utilisés correctement, transforment une simple feuille de calcul en une véritable application interactive. Dans la section suivante, nous découvrirons les événements d'application qui permettent de surveiller Excel dans son ensemble.

⏭️
