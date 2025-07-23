🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.4 Automation avec Access

## Introduction à Access Automation

L'automation avec Access permet de gérer des bases de données directement depuis Excel avec VBA. C'est particulièrement utile pour importer/exporter des données, exécuter des requêtes complexes, gérer de gros volumes de données et créer des rapports basés sur des bases de données relationnelles.

## Pourquoi utiliser Access depuis Excel ?

### Avantages d'Access
- **Gestion de gros volumes** : Access peut traiter des millions d'enregistrements
- **Relations entre tables** : Gestion des clés étrangères et contraintes
- **Requêtes complexes** : SQL avancé avec jointures multiples
- **Intégrité des données** : Validation et contraintes automatiques
- **Performance** : Optimisé pour les opérations de base de données

### Quand utiliser Access depuis Excel
- Vos données Excel dépassent les limites (plus de 1 million de lignes)
- Vous avez besoin de relations complexes entre les données
- Plusieurs utilisateurs doivent accéder aux mêmes données
- Vous voulez centraliser les données et les analyser dans Excel

## Première étape : Créer une connexion avec Access

### Méthode simple pour débuter

```vba
Sub PremierTestAccess()
    ' Créer une connexion avec Access
    Dim accessApp As Object
    Set accessApp = CreateObject("Access.Application")

    ' Rendre Access visible (optionnel)
    accessApp.Visible = True

    ' Ouvrir une base de données existante (changez le chemin)
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"

    ' Afficher un message
    MsgBox "Connexion à Access réussie !"

    ' Fermer et libérer
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing
End Sub
```

**Explication ligne par ligne :**
- `CreateObject("Access.Application")` : Lance Access
- `accessApp.Visible = True` : Rend Access visible (facultatif)
- `OpenCurrentDatabase` : Ouvre une base de données existante
- `CloseCurrentDatabase` : Ferme la base de données
- `accessApp.Quit` : Ferme Access complètement

## Créer une nouvelle base de données Access

```vba
Sub CreerNouvelleBaseDeDonnees()
    Dim accessApp As Object
    Set accessApp = CreateObject("Access.Application")

    ' Créer une nouvelle base de données
    Dim cheminBD As String
    cheminBD = Environ("USERPROFILE") & "\Desktop\NouvelleBD.accdb"

    accessApp.NewCurrentDatabase cheminBD
    accessApp.Visible = True

    MsgBox "Nouvelle base de données créée : " & cheminBD

    ' Ne pas fermer immédiatement pour voir le résultat
    Set accessApp = Nothing
End Sub
```

## Comprendre les objets Access depuis Excel

### Hiérarchie simplifiée d'Access
```
Application (Access lui-même)
├── CurrentDb (Base de données courante)
│   ├── TableDefs (Définitions des tables)
│   ├── QueryDefs (Définitions des requêtes)
│   └── Recordsets (Ensembles d'enregistrements)
```

### Exemple d'accès aux objets

```vba
Sub ExplorerObjetAccess()
    Dim accessApp As Object
    Dim db As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"

    ' Accéder à la base de données courante
    Set db = accessApp.CurrentDb

    ' Lister les tables
    Debug.Print "Tables disponibles :"
    Dim i As Integer
    For i = 0 To db.TableDefs.Count - 1
        If Left(db.TableDefs(i).Name, 4) <> "MSys" Then  ' Ignorer les tables système
            Debug.Print "- " & db.TableDefs(i).Name
        End If
    Next i

    ' Lister les requêtes
    Debug.Print vbCrLf & "Requêtes disponibles :"
    For i = 0 To db.QueryDefs.Count - 1
        Debug.Print "- " & db.QueryDefs(i).Name
    Next i

    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing

    MsgBox "Informations affichées dans la fenêtre Exécution immédiate (Ctrl+G)"
End Sub
```

## Importer des données Excel vers Access

### Créer une table et y insérer des données

```vba
Sub ImporterDonneesVersAccess()
    Dim accessApp As Object
    Dim db As Object
    Dim rs As Object

    Set accessApp = CreateObject("Access.Application")

    ' Créer ou ouvrir la base de données
    Dim cheminBD As String
    cheminBD = Environ("USERPROFILE") & "\Desktop\DonneesExcel.accdb"

    ' Créer la base si elle n'existe pas
    On Error Resume Next
    accessApp.OpenCurrentDatabase cheminBD
    If Err.Number <> 0 Then
        Err.Clear
        accessApp.NewCurrentDatabase cheminBD
    End If
    On Error GoTo 0

    Set db = accessApp.CurrentDb

    ' Créer une table si elle n'existe pas
    On Error Resume Next
    db.Execute "CREATE TABLE Ventes (ID AUTOINCREMENT PRIMARY KEY, Produit TEXT(50), Quantite INTEGER, Prix CURRENCY, DateVente DATETIME)"
    On Error GoTo 0

    ' Ouvrir un recordset pour ajouter des données
    Set rs = db.OpenRecordset("Ventes")

    ' Supposons des données Excel en A2:D6 (Produit, Quantité, Prix, Date)
    Dim i As Integer
    Dim derniereLigne As Integer
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To derniereLigne
        rs.AddNew
        rs.Fields("Produit") = Cells(i, 1).Value
        rs.Fields("Quantite") = Cells(i, 2).Value
        rs.Fields("Prix") = Cells(i, 3).Value
        rs.Fields("DateVente") = Cells(i, 4).Value
        rs.Update
    Next i

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    accessApp.Visible = True  ' Montrer le résultat
    MsgBox (derniereLigne - 1) & " enregistrements importés dans Access !"

    Set accessApp = Nothing
End Sub
```

## Exporter des données Access vers Excel

### Méthode simple : Récupérer toute une table

```vba
Sub ExporterTableAccessVersExcel()
    Dim accessApp As Object
    Dim db As Object
    Dim rs As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"
    Set db = accessApp.CurrentDb

    ' Ouvrir la table
    Set rs = db.OpenRecordset("Ventes")

    ' Vider la feuille Excel actuelle
    Cells.Clear

    ' Ajouter les en-têtes
    Dim col As Integer
    For col = 0 To rs.Fields.Count - 1
        Cells(1, col + 1).Value = rs.Fields(col).Name
    Next col

    ' Ajouter les données
    Dim ligne As Integer
    ligne = 2
    rs.MoveFirst
    Do While Not rs.EOF
        For col = 0 To rs.Fields.Count - 1
            Cells(ligne, col + 1).Value = rs.Fields(col).Value
        Next col
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing

    ' Formater les en-têtes
    Range("A1", Cells(1, col)).Font.Bold = True
    Range("A1", Cells(1, col)).Interior.Color = RGB(200, 200, 200)

    MsgBox "Données exportées avec succès ! " & (ligne - 2) & " enregistrements"
End Sub
```

## Exécuter des requêtes SQL depuis Excel

### Requête SELECT simple

```vba
Sub ExecuterRequeteSQL()
    Dim accessApp As Object
    Dim db As Object
    Dim rs As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"
    Set db = accessApp.CurrentDb

    ' Exécuter une requête SQL
    Dim sql As String
    sql = "SELECT Produit, SUM(Quantite) AS TotalQuantite, SUM(Prix * Quantite) AS ChiffreAffaires " & _
          "FROM Ventes " & _
          "WHERE DateVente >= #01/01/2024# " & _
          "GROUP BY Produit " & _
          "ORDER BY ChiffreAffaires DESC"

    Set rs = db.OpenRecordset(sql)

    ' Vider la feuille
    Cells.Clear

    ' En-têtes
    Cells(1, 1).Value = "Produit"
    Cells(1, 2).Value = "Quantité totale"
    Cells(1, 3).Value = "Chiffre d'affaires"

    ' Données
    Dim ligne As Integer
    ligne = 2
    Do While Not rs.EOF
        Cells(ligne, 1).Value = rs.Fields("Produit").Value
        Cells(ligne, 2).Value = rs.Fields("TotalQuantite").Value
        Cells(ligne, 3).Value = rs.Fields("ChiffreAffaires").Value
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing

    ' Formatage
    Range("A1:C1").Font.Bold = True
    Range("A1:C1").Interior.Color = RGB(200, 200, 200)
    Range("C:C").NumberFormat = "#,##0.00 €"

    MsgBox "Requête exécutée ! " & (ligne - 2) & " résultats"
End Sub
```

### Requêtes avec paramètres depuis Excel

```vba
Sub RequeteAvecParametres()
    Dim accessApp As Object
    Dim db As Object
    Dim rs As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"
    Set db = accessApp.CurrentDb

    ' Paramètres depuis des cellules Excel
    Dim dateDebut As Date
    Dim dateFin As Date
    Dim produitRecherche As String

    dateDebut = Range("E1").Value  ' Date de début en E1
    dateFin = Range("E2").Value    ' Date de fin en E2
    produitRecherche = Range("E3").Value  ' Produit recherché en E3

    ' Construire la requête SQL avec paramètres
    Dim sql As String
    sql = "SELECT * FROM Ventes WHERE "
    sql = sql & "DateVente BETWEEN #" & Format(dateDebut, "mm/dd/yyyy") & "# AND #" & Format(dateFin, "mm/dd/yyyy") & "#"

    If produitRecherche <> "" Then
        sql = sql & " AND Produit LIKE '*" & produitRecherche & "*'"
    End If

    sql = sql & " ORDER BY DateVente DESC"

    Set rs = db.OpenRecordset(sql)

    ' Afficher les résultats (même logique que précédemment)
    Cells.Clear

    ' En-têtes
    Dim col As Integer
    For col = 0 To rs.Fields.Count - 1
        Cells(1, col + 1).Value = rs.Fields(col).Name
    Next col

    ' Données
    Dim ligne As Integer
    ligne = 2
    Do While Not rs.EOF
        For col = 0 To rs.Fields.Count - 1
            Cells(ligne, col + 1).Value = rs.Fields(col).Value
        Next col
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing

    MsgBox "Recherche terminée ! " & (ligne - 2) & " résultats trouvés"
End Sub
```

## Mettre à jour des données dans Access

### Mise à jour simple

```vba
Sub MettreAJourDonneesAccess()
    Dim accessApp As Object
    Dim db As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"
    Set db = accessApp.CurrentDb

    ' Mise à jour avec SQL UPDATE
    Dim sql As String
    sql = "UPDATE Ventes SET Prix = Prix * 1.1 WHERE DateVente < #01/01/2024#"

    db.Execute sql

    MsgBox "Prix augmentés de 10% pour les ventes antérieures à 2024"

    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing
End Sub
```

### Mise à jour basée sur Excel

```vba
Sub MettreAJourDepuisExcel()
    Dim accessApp As Object
    Dim db As Object
    Dim rs As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"
    Set db = accessApp.CurrentDb

    ' Supposons une liste de mises à jour en Excel : A=ID, B=NouveauPrix
    Dim i As Integer
    Dim derniereLigne As Integer
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To derniereLigne
        Dim idProduit As Long
        Dim nouveauPrix As Double

        idProduit = Cells(i, 1).Value
        nouveauPrix = Cells(i, 2).Value

        ' Mettre à jour chaque enregistrement
        Dim sql As String
        sql = "UPDATE Ventes SET Prix = " & nouveauPrix & " WHERE ID = " & idProduit
        db.Execute sql
    Next i

    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing

    MsgBox (derniereLigne - 1) & " enregistrements mis à jour"
End Sub
```

## Créer des rapports Access depuis Excel

```vba
Sub CreerRapportAccess()
    Dim accessApp As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"
    accessApp.Visible = True

    ' Ouvrir un rapport existant (si il existe)
    On Error Resume Next
    accessApp.DoCmd.OpenReport "RapportVentes", 0  ' 0 = Mode Aperçu

    If Err.Number <> 0 Then
        MsgBox "Le rapport 'RapportVentes' n'existe pas. Créez-le d'abord dans Access."
        Err.Clear
    Else
        MsgBox "Rapport ouvert dans Access"
    End If
    On Error GoTo 0

    ' Ne pas fermer Access pour voir le rapport
    Set accessApp = Nothing
End Sub
```

## Sauvegarder et compacter la base de données

```vba
Sub SauvegarderEtCompacterBD()
    Dim accessApp As Object

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\MaBaseDeDonnees.accdb"

    ' Compacter et réparer la base de données
    Dim cheminOriginal As String
    Dim cheminSauvegarde As String

    cheminOriginal = "C:\MonDossier\MaBaseDeDonnees.accdb"
    cheminSauvegarde = "C:\MonDossier\MaBaseDeDonnees_Sauvegarde.accdb"

    accessApp.CloseCurrentDatabase

    ' Compacter vers un nouveau fichier
    accessApp.CompactRepair cheminOriginal, cheminSauvegarde

    MsgBox "Base de données compactée et sauvegardée : " & cheminSauvegarde

    accessApp.Quit
    Set accessApp = Nothing
End Sub
```

## Gestion d'erreurs avec Access

```vba
Sub GestionErreursAccess()
    Dim accessApp As Object
    Dim db As Object

    On Error GoTo GestionErreur

    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase "C:\MonDossier\BaseDeDonneesInexistante.accdb"

    ' Code normal ici...

    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing

    Exit Sub

GestionErreur:
    Select Case Err.Number
        Case 3024
            MsgBox "Fichier de base de données introuvable ou corrompu"
        Case 3050
            MsgBox "Impossible de verrouiller le fichier (peut-être ouvert ailleurs)"
        Case 3343
            MsgBox "Erreur de format de base de données"
        Case Else
            MsgBox "Erreur Access : " & Err.Number & " - " & Err.Description
    End Select

    ' Nettoyage en cas d'erreur
    If Not db Is Nothing Then Set db = Nothing
    If Not accessApp Is Nothing Then
        accessApp.CloseCurrentDatabase
        accessApp.Quit
        Set accessApp = Nothing
    End If
End Sub
```

## Exemple complet : Système de synchronisation Excel-Access

```vba
Sub SystemeSynchronisationExcelAccess()
    Dim accessApp As Object
    Dim db As Object
    Dim rs As Object
    Dim cheminBD As String

    On Error GoTo GestionErreur

    ' === ÉTAPE 1 : INITIALISATION ===
    cheminBD = Environ("USERPROFILE") & "\Desktop\SynchroVentes.accdb"

    Set accessApp = CreateObject("Access.Application")

    ' Créer la base si elle n'existe pas
    On Error Resume Next
    accessApp.OpenCurrentDatabase cheminBD
    If Err.Number <> 0 Then
        Err.Clear
        accessApp.NewCurrentDatabase cheminBD

        ' Créer la structure de table
        Set db = accessApp.CurrentDb
        db.Execute "CREATE TABLE VentesHistory (ID AUTOINCREMENT PRIMARY KEY, " & _
                   "Produit TEXT(50), Quantite INTEGER, PrixUnitaire CURRENCY, " & _
                   "DateVente DATETIME, Vendeur TEXT(30), Region TEXT(20), " & _
                   "DateImport DATETIME)"

        MsgBox "Nouvelle base de données créée avec la structure"
    End If
    On Error GoTo GestionErreur

    Set db = accessApp.CurrentDb

    ' === ÉTAPE 2 : IMPORT DES NOUVELLES DONNÉES ===
    ' Supposons des données Excel en A2:G (Produit, Quantité, Prix, Date, Vendeur, Région)
    Dim derniereLigne As Integer
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    If derniereLigne > 1 Then
        Set rs = db.OpenRecordset("VentesHistory")

        Dim i As Integer
        Dim nbImports As Integer
        nbImports = 0

        For i = 2 To derniereLigne
            ' Vérifier si cette ligne n'existe pas déjà (éviter les doublons)
            Dim critereRecherche As String
            critereRecherche = "Produit = '" & Cells(i, 1).Value & "' AND " & _
                              "DateVente = #" & Format(Cells(i, 4).Value, "mm/dd/yyyy") & "# AND " & _
                              "Vendeur = '" & Cells(i, 5).Value & "'"

            Dim rsVerif As Object
            Set rsVerif = db.OpenRecordset("SELECT COUNT(*) AS Nb FROM VentesHistory WHERE " & critereRecherche)

            If rsVerif.Fields("Nb").Value = 0 Then
                ' Nouvel enregistrement
                rs.AddNew
                rs.Fields("Produit") = Cells(i, 1).Value
                rs.Fields("Quantite") = Cells(i, 2).Value
                rs.Fields("PrixUnitaire") = Cells(i, 3).Value
                rs.Fields("DateVente") = Cells(i, 4).Value
                rs.Fields("Vendeur") = Cells(i, 5).Value
                rs.Fields("Region") = Cells(i, 6).Value
                rs.Fields("DateImport") = Now()
                rs.Update
                nbImports = nbImports + 1
            End If

            rsVerif.Close
            Set rsVerif = Nothing
        Next i

        rs.Close
        Set rs = Nothing

        MsgBox nbImports & " nouveaux enregistrements importés dans Access"
    End If

    ' === ÉTAPE 3 : EXPORT DES STATISTIQUES VERS EXCEL ===
    ' Créer un rapport de synthèse
    Dim sql As String
    sql = "SELECT Region, Vendeur, COUNT(*) AS NbVentes, " & _
          "SUM(Quantite) AS TotalQuantite, " & _
          "SUM(Quantite * PrixUnitaire) AS ChiffreAffaires " & _
          "FROM VentesHistory " & _
          "WHERE DateVente >= #" & Format(DateAdd("m", -1, Date), "mm/dd/yyyy") & "# " & _
          "GROUP BY Region, Vendeur " & _
          "ORDER BY ChiffreAffaires DESC"

    Set rs = db.OpenRecordset(sql)

    ' Créer une nouvelle feuille pour les statistiques
    Dim nouvelleFeuille As Worksheet
    Set nouvelleFeuille = ThisWorkbook.Worksheets.Add
    nouvelleFeuille.Name = "Statistiques_" & Format(Date, "mmyyyy")

    ' En-têtes
    nouvelleFeuille.Cells(1, 1).Value = "Région"
    nouvelleFeuille.Cells(1, 2).Value = "Vendeur"
    nouvelleFeuille.Cells(1, 3).Value = "Nb Ventes"
    nouvelleFeuille.Cells(1, 4).Value = "Total Quantité"
    nouvelleFeuille.Cells(1, 5).Value = "Chiffre d'Affaires"

    ' Données
    Dim ligne As Integer
    ligne = 2
    Do While Not rs.EOF
        nouvelleFeuille.Cells(ligne, 1).Value = rs.Fields("Region").Value
        nouvelleFeuille.Cells(ligne, 2).Value = rs.Fields("Vendeur").Value
        nouvelleFeuille.Cells(ligne, 3).Value = rs.Fields("NbVentes").Value
        nouvelleFeuille.Cells(ligne, 4).Value = rs.Fields("TotalQuantite").Value
        nouvelleFeuille.Cells(ligne, 5).Value = rs.Fields("ChiffreAffaires").Value
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    ' Formatage
    With nouvelleFeuille.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(100, 150, 200)
        .Font.Color = RGB(255, 255, 255)
    End With
    nouvelleFeuille.Range("E:E").NumberFormat = "#,##0.00 €"
    nouvelleFeuille.Columns.AutoFit

    ' === ÉTAPE 4 : SAUVEGARDE ===
    ThisWorkbook.Save

    ' Fermeture Access
    Set db = Nothing
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    Set accessApp = Nothing

    MsgBox "Synchronisation terminée !" & vbCrLf & _
           "• " & nbImports & " nouveaux enregistrements importés" & vbCrLf & _
           "• Statistiques créées dans la feuille : " & nouvelleFeuille.Name & vbCrLf & _
           "• Base de données : " & cheminBD

    Exit Sub

GestionErreur:
    MsgBox "Erreur de synchronisation : " & Err.Description

    ' Nettoyage
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
    If Not accessApp Is Nothing Then
        accessApp.CloseCurrentDatabase
        accessApp.Quit
        Set accessApp = Nothing
    End If
End Sub
```

## Points importants à retenir

### ✅ Bonnes pratiques
- Toujours fermer les Recordsets avec `.Close`
- Utiliser `Set variable = Nothing` pour libérer la mémoire
- Fermer Access avec `CloseCurrentDatabase` et `Quit`
- Toujours gérer les erreurs spécifiques à Access
- Sauvegarder régulièrement les bases de données importantes

### ⚠️ Erreurs courantes à éviter
- Oublier de fermer les Recordsets (fuites mémoire)
- Ne pas gérer les bases de données verrouillées
- Utiliser des chemins de fichiers en dur
- Ne pas vérifier l'existence des tables/requêtes
- Manipulation directe de gros volumes sans optimisation

### 💡 Conseils pour débuter
- Commencez par des bases de données simples (une seule table)
- Testez vos requêtes SQL dans Access avant de les automatiser
- Utilisez la fenêtre Exécution immédiate pour déboguer
- Créez des sauvegardes avant les opérations de mise à jour
- Apprenez les bases du SQL pour maximiser l'efficacité

### 🎯 Utilisations typiques
- **Archivage** : Déplacer les anciennes données Excel vers Access
- **Reporting** : Générer des rapports complexes depuis une base centralisée
- **Synchronisation** : Garder Excel et Access synchronisés
- **Analyse** : Exploiter la puissance SQL pour des analyses avancées
- **Multi-utilisateurs** : Centraliser les données pour plusieurs utilisateurs Excel

### 📊 Différences clés Excel vs Access
- **Excel** : Idéal pour calculs, analyses, graphiques, présentation
- **Access** : Idéal pour stockage, relations, requêtes complexes, intégrité
- **Ensemble** : Excel pour l'interface utilisateur, Access pour les données

L'automation avec Access transforme Excel en véritable front-end pour des applications de gestion de données. Cette combinaison offre le meilleur des deux mondes : la puissance d'analyse d'Excel et la robustesse de gestion de données d'Access !

⏭️
