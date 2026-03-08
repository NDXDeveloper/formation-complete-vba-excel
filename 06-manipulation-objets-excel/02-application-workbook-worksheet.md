🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.2. Application, Workbook, Worksheet

## Introduction aux trois objets fondamentaux

Dans cette section, nous allons explorer en détail les trois objets les plus importants du modèle Excel : **Application**, **Workbook**, et **Worksheet**. Ces trois objets forment la base de toute programmation VBA et correspondent à ce que vous manipulez quotidiennement dans Excel :

- **Application** = Excel lui-même (le logiciel)
- **Workbook** = Le fichier Excel (le classeur)
- **Worksheet** = Une feuille dans le classeur

Pensez-y comme à une hiérarchie logique : vous ouvrez Excel (Application), puis vous travaillez dans un fichier (Workbook), et enfin vous saisissez des données dans une feuille (Worksheet).

---

## L'objet Application

### Qu'est-ce que l'objet Application ?

L'objet **Application** représente Excel dans son ensemble. C'est le niveau le plus élevé de la hiérarchie des objets. Quand vous lancez Excel, vous créez une instance de l'objet Application.

### Propriétés importantes de Application

#### 1. Informations sur Excel et l'environnement

```vba
' Connaître la version d'Excel
Debug.Print Application.Version          ' Affiche "16.0" pour Excel 2016/2019/365

' Nom de l'utilisateur connecté
Debug.Print Application.UserName         ' Affiche le nom Windows de l'utilisateur

' Chemin d'installation d'Excel
Debug.Print Application.Path             ' Ex: "C:\Program Files\Microsoft Office\root\Office16"
```

#### 2. Contrôle de l'affichage et des performances

```vba
' Désactiver la mise à jour de l'écran (améliore les performances)
Application.ScreenUpdating = False       ' L'écran ne se rafraîchit plus
' ... votre code ici ...
Application.ScreenUpdating = True        ' Réactiver l'affichage

' Désactiver les alertes système
Application.DisplayAlerts = False        ' Plus de boîtes de dialogue d'avertissement
' ... votre code ici ...
Application.DisplayAlerts = True         ' Réactiver les alertes

' Contrôler les calculs automatiques
Application.Calculation = xlCalculationManual     ' Calculs en mode manuel  
Application.Calculation = xlCalculationAutomatic  ' Calculs automatiques (par défaut)  
```

#### 3. État d'Excel

```vba
' Vérifier si Excel est prêt à recevoir des commandes
If Application.Ready Then
    ' Excel est disponible pour traiter des commandes
End If

' Savoir combien de classeurs sont ouverts
Debug.Print Application.Workbooks.Count
```

### Méthodes importantes de Application

#### 1. Gestion des calculs

```vba
' Forcer le recalcul de toutes les feuilles ouvertes
Application.Calculate

' Recalculer uniquement les cellules modifiées
Application.CalculateUntilAsyncQueriesDone
```

#### 2. Contrôle du temps et des pauses

```vba
' Faire une pause dans l'exécution du code
Application.Wait Now + TimeValue("00:00:03")  ' Pause de 3 secondes

' Permettre à Windows de traiter d'autres tâches
DoEvents  ' Fonction VBA autonome (pas une méthode d'Application)
```

#### 3. Interaction avec l'utilisateur

```vba
' Faire clignoter l'icône Excel dans la barre des tâches
Application.WindowState = xlMinimized  
Application.WindowState = xlNormal  
```

### Bonnes pratiques avec Application

**Optimisation des performances :**
```vba
Sub ExempleOptimisation()
    ' Sauvegarder les états actuels
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation

    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation

    ' Optimiser pendant le traitement
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Votre code ici...

    ' Restaurer les états d'origine
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation
End Sub
```

---

## L'objet Workbook

### Qu'est-ce que l'objet Workbook ?

Un objet **Workbook** représente un fichier Excel ouvert. Chaque fois que vous ouvrez un fichier .xlsx, .xlsm, ou que vous créez un nouveau classeur, vous créez un objet Workbook.

### Accéder aux objets Workbook

#### 1. Le classeur actif
```vba
' Le classeur actuellement sélectionné
Dim monClasseur As Workbook  
Set monClasseur = ActiveWorkbook  

' Ou directement sans variable
ActiveWorkbook.Save
```

#### 2. Un classeur spécifique par son nom
```vba
' Accéder à un classeur ouvert par son nom (avec extension)
Dim classeurData As Workbook  
Set classeurData = Workbooks("Données2024.xlsx")  

' Note : incluez toujours l'extension (.xlsx, .xlsm, etc.)
' Workbooks("Données2024") sans extension peut échouer
```

#### 3. Un classeur par son index
```vba
' Le premier classeur ouvert
Dim premierClasseur As Workbook  
Set premierClasseur = Workbooks(1)  
```

### Propriétés importantes de Workbook

#### 1. Informations sur le fichier

```vba
' Nom du fichier (avec extension)
Debug.Print ActiveWorkbook.Name          ' Ex: "MonFichier.xlsx"

' Chemin du dossier (sans le nom du fichier)
Debug.Print ActiveWorkbook.Path          ' Ex: "C:\MesDocuments"

' Chemin complet (dossier + nom)
Debug.Print ActiveWorkbook.FullName      ' Ex: "C:\MesDocuments\MonFichier.xlsx"

' Vérifier si le classeur a été modifié depuis la dernière sauvegarde
If Not ActiveWorkbook.Saved Then
    MsgBox "Le classeur contient des modifications non sauvegardées"
End If
```

#### 2. État et propriétés du classeur

```vba
' Vérifier si le classeur est protégé
If ActiveWorkbook.ProtectStructure Then
    MsgBox "La structure du classeur est protégée"
End If

' Vérifier si le classeur est en lecture seule
If ActiveWorkbook.ReadOnly Then
    MsgBox "Ce classeur est ouvert en lecture seule"
End If

' Connaître le nombre de feuilles dans le classeur
Debug.Print ActiveWorkbook.Worksheets.Count
```

### Méthodes importantes de Workbook

#### 1. Gestion des fichiers

```vba
' Sauvegarder le classeur
ActiveWorkbook.Save

' Sauvegarder sous un nouveau nom
ActiveWorkbook.SaveAs "C:\NouveauDossier\NouveauNom.xlsx"

' Fermer le classeur
ActiveWorkbook.Close SaveChanges:=True   ' Ferme en sauvegardant  
ActiveWorkbook.Close SaveChanges:=False  ' Ferme sans sauvegarder  
```

#### 2. Gestion des feuilles

```vba
' Ajouter une nouvelle feuille
Dim nouvelleFeuille As Worksheet  
Set nouvelleFeuille = ActiveWorkbook.Worksheets.Add  

' Ajouter une feuille avec un nom spécifique
Set nouvelleFeuille = ActiveWorkbook.Worksheets.Add  
nouvelleFeuille.Name = "Nouvelles Données"  
```

#### 3. Protection et sécurité

```vba
' Protéger la structure du classeur (empêche l'ajout/suppression de feuilles)
ActiveWorkbook.Protect Password:="motdepasse"

' Enlever la protection
ActiveWorkbook.Unprotect Password:="motdepasse"
```

### Ouvrir et créer des classeurs

#### 1. Ouvrir un classeur existant

```vba
' Méthode simple
Workbooks.Open "C:\MesDocuments\MonFichier.xlsx"

' Méthode avec gestion d'erreur
Dim nouveauClasseur As Workbook  
On Error Resume Next  
Set nouveauClasseur = Workbooks.Open("C:\MesDocuments\MonFichier.xlsx")  
If nouveauClasseur Is Nothing Then  
    MsgBox "Impossible d'ouvrir le fichier"
End If  
On Error GoTo 0  
```

#### 2. Créer un nouveau classeur

```vba
' Créer un classeur vide
Dim nouveauClasseur As Workbook  
Set nouveauClasseur = Workbooks.Add  

' Le nouveau classeur devient automatiquement le classeur actif
Debug.Print ActiveWorkbook.Name  ' Affiche quelque chose comme "Classeur1"
```

---

## L'objet Worksheet

### Qu'est-ce que l'objet Worksheet ?

Un objet **Worksheet** représente une feuille de calcul individuelle dans un classeur. C'est sur les feuilles que vous travaillez au quotidien : saisir des données, créer des formules, faire des graphiques.

### Accéder aux objets Worksheet

#### 1. La feuille active

```vba
' La feuille actuellement sélectionnée
Dim maFeuille As Worksheet  
Set maFeuille = ActiveSheet  

' Utilisation directe
ActiveSheet.Name = "Feuille Principale"
```

#### 2. Une feuille spécifique par son nom

```vba
' Accéder à une feuille par son nom
Dim feuilleDonnees As Worksheet  
Set feuilleDonnees = Worksheets("Données")  

' Ou dans un classeur spécifique
Set feuilleDonnees = Workbooks("MonFichier.xlsx").Worksheets("Données")
```

#### 3. Une feuille par son index

```vba
' La première feuille du classeur
Dim premiereFeuille As Worksheet  
Set premiereFeuille = Worksheets(1)  

' La dernière feuille
Dim derniereFeuille As Worksheet  
Set derniereFeuille = Worksheets(Worksheets.Count)  
```

### Propriétés importantes de Worksheet

#### 1. Informations de base

```vba
' Nom de la feuille
Debug.Print ActiveSheet.Name

' Modifier le nom de la feuille
ActiveSheet.Name = "Données 2024"

' Index de la feuille (sa position)
Debug.Print ActiveSheet.Index  ' 1 pour la première feuille, 2 pour la seconde, etc.
```

#### 2. Visibilité et état

```vba
' Masquer une feuille
Worksheets("Données").Visible = xlSheetHidden

' Afficher une feuille masquée
Worksheets("Données").Visible = xlSheetVisible

' Masquer complètement (invisible même dans le menu contextuel)
Worksheets("Données").Visible = xlSheetVeryHidden
```

#### 3. Zone de travail

```vba
' Obtenir la plage de cellules utilisées
Dim plageUtilisee As Range  
Set plageUtilisee = ActiveSheet.UsedRange  
Debug.Print "Zone utilisée : " & plageUtilisee.Address  

' Dernière ligne contenant des données
Dim derniereLigne As Long  
derniereLigne = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row  

' Dernière colonne contenant des données
Dim derniereColonne As Long  
derniereColonne = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column  
```

### Méthodes importantes de Worksheet

#### 1. Activation et sélection

```vba
' Activer une feuille (la rendre visible et active)
Worksheets("Données").Activate

' Sélectionner une feuille (peut être utilisée avec Ctrl+clic pour sélections multiples)
Worksheets("Données").Select
```

#### 2. Copier et déplacer

```vba
' Copier une feuille à la fin du classeur
ActiveSheet.Copy After:=Worksheets(Worksheets.Count)

' Copier avant une feuille spécifique
ActiveSheet.Copy Before:=Worksheets("Résultats")

' Déplacer une feuille
ActiveSheet.Move After:=Worksheets("Données")
```

#### 3. Protection

```vba
' Protéger une feuille
ActiveSheet.Protect Password:="motdepasse"

' Protéger en permettant certaines actions
ActiveSheet.Protect Password:="motdepasse", _
                   AllowInsertingRows:=True, _
                   AllowDeletingRows:=True

' Enlever la protection
ActiveSheet.Unprotect Password:="motdepasse"
```

#### 4. Gestion des feuilles

```vba
' Supprimer une feuille (attention : irréversible !)
Application.DisplayAlerts = False  ' Éviter la boîte de confirmation  
Worksheets("FeuilleASupprimer").Delete  
Application.DisplayAlerts = True  
```

### Travailler avec plusieurs feuilles

#### 1. Parcourir toutes les feuilles

```vba
' Afficher le nom de toutes les feuilles
Dim feuille As Worksheet  
For Each feuille In Worksheets  
    Debug.Print feuille.Name
Next feuille

' Ou avec un index
Dim i As Integer  
For i = 1 To Worksheets.Count  
    Debug.Print Worksheets(i).Name
Next i
```

#### 2. Créer et nommer de nouvelles feuilles

```vba
' Ajouter une feuille avec un nom spécifique
Dim nouvelleFeuille As Worksheet  
Set nouvelleFeuille = Worksheets.Add  
nouvelleFeuille.Name = "Rapport " & Format(Date, "yyyy-mm-dd")  

' Ajouter plusieurs feuilles d'un coup
Worksheets.Add Count:=3  ' Ajoute 3 nouvelles feuilles
```

## Relations entre les trois objets

### Navigation dans la hiérarchie

```vba
' Accès complet (explicite)
Application.Workbooks("MonFichier.xlsx").Worksheets("Données").Range("A1").Value = "Test"

' Accès simplifié (si vous travaillez sur le classeur/feuille actif)
ActiveSheet.Range("A1").Value = "Test"

' Ou encore plus simple
Range("A1").Value = "Test"
```

### Bonnes pratiques pour débuter

1. **Utilisez les objets actifs quand c'est possible** : `ActiveWorkbook`, `ActiveSheet` simplifient le code
2. **Nommez vos feuilles de façon explicite** : "Données", "Résultats", "Paramètres" plutôt que "Feuil1", "Feuil2"
3. **Vérifiez l'existence avant d'accéder** aux objets pour éviter les erreurs
4. **Sauvegardez régulièrement** avec `ActiveWorkbook.Save`

### Exemple pratique complet

```vba
Sub ExempleComplet()
    ' Créer un nouveau classeur
    Dim nouveauClasseur As Workbook
    Set nouveauClasseur = Workbooks.Add

    ' Renommer la première feuille
    nouveauClasseur.Worksheets(1).Name = "Données Principales"

    ' Ajouter une seconde feuille
    Dim feuilleResultats As Worksheet
    Set feuilleResultats = nouveauClasseur.Worksheets.Add
    feuilleResultats.Name = "Résultats"

    ' Sauvegarder le classeur
    nouveauClasseur.SaveAs "C:\MonDossier\NouveauRapport.xlsx"

    MsgBox "Classeur créé avec succès !"
End Sub
```

## Points clés à retenir

- **Application** contrôle Excel dans son ensemble et ses paramètres globaux
- **Workbook** représente un fichier Excel et gère la sauvegarde, l'ouverture, la fermeture
- **Worksheet** représente une feuille individuelle où vous manipulez les données
- Ces trois objets forment la base de toute programmation VBA efficace
- Les objets "actifs" (`ActiveWorkbook`, `ActiveSheet`) sont des raccourcis pratiques
- Toujours penser à la hiérarchie : Application → Workbook → Worksheet → cellules

Dans la section suivante, nous découvrirons comment manipuler les cellules et plages de cellules avec les objets Range et Cells.

⏭️
