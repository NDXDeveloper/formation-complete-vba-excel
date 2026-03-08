🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.4 Import/Export de données

## Introduction

L'import et l'export de données sont comme les poumons d'Excel : ils permettent de faire entrer et sortir l'information de manière automatisée. Au lieu de copier-coller manuellement des milliers de lignes, VBA peut le faire pour vous en quelques secondes !

Imaginez que vous êtes le chef d'une gare : les trains (données) arrivent de différentes villes (sources) et repartent vers d'autres destinations (cibles). Votre rôle est d'organiser efficacement ces mouvements.

## Pourquoi automatiser les imports/exports ?

### Avantages de l'automatisation
- **Gain de temps** : Plus de copier-coller fastidieux
- **Fiabilité** : Élimination des erreurs humaines
- **Répétabilité** : Même processus à chaque fois
- **Planification** : Peut se faire automatiquement selon un planning
- **Volume** : Traitement de grandes quantités de données

### Scénarios courants
- **Rapports quotidiens** : Import des ventes de la veille
- **Consolidation** : Fusionner plusieurs fichiers Excel
- **Sauvegarde** : Export vers base de données pour archivage
- **Distribution** : Création de fichiers par région/département
- **Migration** : Transfert entre anciens et nouveaux systèmes

## Import de données depuis Excel

### Import depuis un autre fichier Excel

```vba
Sub ImporterDepuisExcel()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connectionString As String
    Dim cheminFichier As String

    ' Sélection du fichier à importer
    cheminFichier = Application.GetOpenFilename( _
        "Fichiers Excel (*.xlsx), *.xlsx," & _
        "Anciens Excel (*.xls), *.xls", _
        , "Sélectionnez le fichier à importer")

    If cheminFichier = "False" Then
        MsgBox "Import annulé"
        Exit Sub
    End If

    ' Vérification de l'existence du fichier
    If Dir(cheminFichier) = "" Then
        MsgBox "Le fichier sélectionné n'existe pas !"
        Exit Sub
    End If

    ' Configuration de la connexion
    Set conn = New ADODB.Connection
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & cheminFichier & ";" & _
                      "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""

    On Error GoTo GestionErreur

    conn.Open connectionString

    ' Import depuis la première feuille
    Set rs = conn.Execute("SELECT * FROM [Feuil1$]")

    ' Vider la feuille actuelle
    ActiveSheet.Cells.Clear

    ' Copier les en-têtes
    Dim col As Integer
    For col = 0 To rs.Fields.Count - 1
        Cells(1, col + 1).Value = rs.Fields(col).Name
    Next col

    ' Copier les données
    Dim ligne As Long
    ligne = 2

    Do While Not rs.EOF
        For col = 0 To rs.Fields.Count - 1
            Cells(ligne, col + 1).Value = rs.Fields(col).Value
        Next col
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Import terminé : " & (ligne - 2) & " lignes importées depuis " & _
           Dir(cheminFichier)
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'import : " & Err.Description
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub
```

### Import avec sélection de feuille

```vba
Sub ImporterAvecChoixFeuille()
    Dim cheminFichier As String
    Dim nomFeuille As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    ' Sélection du fichier
    cheminFichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    If cheminFichier = "False" Then Exit Sub

    ' Demander le nom de la feuille
    nomFeuille = InputBox("Nom de la feuille à importer ?", "Import", "Feuil1")
    If nomFeuille = "" Then Exit Sub

    ' Ajouter $ si nécessaire
    If Right(nomFeuille, 1) <> "$" Then nomFeuille = nomFeuille & "$"

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes"""

    On Error GoTo GestionErreur

    Set rs = conn.Execute("SELECT * FROM [" & nomFeuille & "]")

    ' Traitement identique au précédent...

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur : Vérifiez que la feuille '" & Replace(nomFeuille, "$", "") & "' existe"
    If Not conn Is Nothing Then conn.Close
End Sub
```

## Import depuis fichiers CSV

### Import CSV simple

```vba
Sub ImporterCSV()
    Dim cheminFichier As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connectionString As String

    ' Sélection du fichier CSV
    cheminFichier = Application.GetOpenFilename( _
        "Fichiers CSV (*.csv), *.csv," & _
        "Fichiers texte (*.txt), *.txt", _
        , "Sélectionnez le fichier CSV")

    If cheminFichier = "False" Then Exit Sub

    ' Le chemin doit pointer vers le DOSSIER, pas le fichier
    Dim cheminDossier As String
    Dim nomFichier As String

    cheminDossier = Left(cheminFichier, InStrRev(cheminFichier, "\"))
    nomFichier = Mid(cheminFichier, InStrRev(cheminFichier, "\") + 1)

    Set conn = New ADODB.Connection
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & cheminDossier & ";" & _
                      "Extended Properties=""Text;HDR=Yes;FMT=Delimited"""

    On Error GoTo GestionErreur

    conn.Open connectionString

    ' Utiliser le nom du fichier dans la requête
    Set rs = conn.Execute("SELECT * FROM [" & nomFichier & "]")

    ' Vider la feuille
    ActiveSheet.Cells.Clear

    ' Import des données
    Dim ligne As Long, col As Integer
    ligne = 1

    ' En-têtes
    For col = 0 To rs.Fields.Count - 1
        Cells(ligne, col + 1).Value = rs.Fields(col).Name
    Next col
    ligne = 2

    ' Données
    Do While Not rs.EOF
        For col = 0 To rs.Fields.Count - 1
            Cells(ligne, col + 1).Value = rs.Fields(col).Value
        Next col
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "CSV importé : " & (ligne - 2) & " lignes"
    Exit Sub

GestionErreur:
    MsgBox "Erreur CSV : " & Err.Description
    If Not conn Is Nothing Then conn.Close
End Sub
```

### Import CSV avec paramètres personnalisés

```vba
Sub ImporterCSVPersonnalise()
    Dim cheminFichier As String
    Dim separateur As String
    Dim avecEntetes As Boolean

    ' Paramètres utilisateur
    cheminFichier = Application.GetOpenFilename("Fichiers CSV (*.csv), *.csv")
    If cheminFichier = "False" Then Exit Sub

    separateur = InputBox("Séparateur (virgule, point-virgule, tabulation) ?", "Import CSV", ",")
    avecEntetes = (MsgBox("Le fichier contient-il des en-têtes ?", vbYesNo + vbQuestion) = vbYes)

    ' Lecture ligne par ligne (méthode alternative)
    Dim numeroFichier As Integer
    Dim ligneTexte As String
    Dim tableau As Variant
    Dim ligne As Long

    numeroFichier = FreeFile
    Open cheminFichier For Input As #numeroFichier

    ligne = 1

    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligneTexte

        ' Séparer les champs
        tableau = Split(ligneTexte, separateur)

        ' Écrire dans Excel
        Dim col As Integer
        For col = 0 To UBound(tableau)
            Cells(ligne, col + 1).Value = Trim(tableau(col))
        Next col

        ligne = ligne + 1
    Loop

    Close #numeroFichier

    MsgBox "Import CSV terminé : " & (ligne - 1) & " lignes"
End Sub
```

## Import depuis base de données

### Import avec filtrage

```vba
Sub ImporterDepuisBaseAvecFiltre()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim dateDebut As Date, dateFin As Date

    ' Paramètres de filtrage
    dateDebut = CDate(InputBox("Date de début (jj/mm/aaaa) ?", "Filtre", Date - 30))
    dateFin = CDate(InputBox("Date de fin (jj/mm/aaaa) ?", "Filtre", Date))

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Requête avec filtre de dates
    sql = "SELECT c.NomClient, cmd.DateCommande, cmd.Montant " & _
          "FROM Clients c INNER JOIN Commandes cmd ON c.ClientID = cmd.ClientID " & _
          "WHERE cmd.DateCommande BETWEEN #" & Format(dateDebut, "mm/dd/yyyy") & "# " & _
          "AND #" & Format(dateFin, "mm/dd/yyyy") & "# " & _
          "ORDER BY cmd.DateCommande"

    Set rs = conn.Execute(sql)

    ' Vider la feuille
    ActiveSheet.Cells.Clear

    ' En-têtes
    Cells(1, 1).Value = "Client"
    Cells(1, 2).Value = "Date"
    Cells(1, 3).Value = "Montant"

    ' Données
    Dim ligne As Long
    ligne = 2

    Do While Not rs.EOF
        Cells(ligne, 1).Value = rs.Fields("NomClient").Value
        Cells(ligne, 2).Value = rs.Fields("DateCommande").Value
        Cells(ligne, 3).Value = rs.Fields("Montant").Value
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' Formatage automatique
    Range("A1:C1").Font.Bold = True
    Range("B:B").NumberFormat = "dd/mm/yyyy"
    Range("C:C").NumberFormat = "#,##0.00 €"
    Columns.AutoFit

    MsgBox "Import terminé : " & (ligne - 2) & " commandes importées"
End Sub
```

## Export vers Excel

### Export vers nouveau fichier Excel

```vba
Sub ExporterVersNouveauExcel()
    Dim nouveauClasseur As Workbook
    Dim feuilleCourante As Worksheet
    Dim nouvelleFeuille As Worksheet
    Dim cheminSauvegarde As String

    Set feuilleCourante = ActiveSheet

    ' Vérifier qu'il y a des données
    If feuilleCourante.UsedRange.Rows.Count = 0 Then
        MsgBox "Aucune donnée à exporter !"
        Exit Sub
    End If

    ' Créer un nouveau classeur
    Set nouveauClasseur = Workbooks.Add
    Set nouvelleFeuille = nouveauClasseur.ActiveSheet

    ' Copier les données
    feuilleCourante.UsedRange.Copy
    nouvelleFeuille.Range("A1").PasteSpecial xlPasteValues
    nouvelleFeuille.Range("A1").PasteSpecial xlPasteFormats

    Application.CutCopyMode = False

    ' Ajuster les colonnes
    nouvelleFeuille.Columns.AutoFit

    ' Sauvegarder
    cheminSauvegarde = Application.GetSaveAsFilename( _
        InitialFilename:="Export_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx", _
        FileFilter:="Fichiers Excel (*.xlsx), *.xlsx")

    If cheminSauvegarde <> "False" Then
        nouveauClasseur.SaveAs cheminSauvegarde
        MsgBox "Export réussi vers : " & cheminSauvegarde
    Else
        nouveauClasseur.Close False
        MsgBox "Export annulé"
    End If

    Set nouveauClasseur = Nothing
    Set nouvelleFeuille = Nothing
    Set feuilleCourante = Nothing
End Sub
```

### Export avec plusieurs feuilles

```vba
Sub ExporterPlusieursFeuilles()
    Dim nouveauClasseur As Workbook
    Dim feuille As Worksheet
    Dim cheminSauvegarde As String
    Dim nbFeuillesExportees As Integer

    ' Demander quelles feuilles exporter
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Exporter toutes les feuilles ?", vbYesNoCancel + vbQuestion)

    If reponse = vbCancel Then Exit Sub

    Set nouveauClasseur = Workbooks.Add

    nbFeuillesExportees = 0

    For Each feuille In ThisWorkbook.Worksheets
        ' Si "Toutes" ou si la feuille contient des données
        If reponse = vbYes Or feuille.UsedRange.Rows.Count > 1 Then

            ' Ajouter une nouvelle feuille
            Dim nouvelleFeuille As Worksheet
            Set nouvelleFeuille = nouveauClasseur.Worksheets.Add
            nouvelleFeuille.Name = feuille.Name

            ' Copier les données
            feuille.UsedRange.Copy
            nouvelleFeuille.Range("A1").PasteSpecial xlPasteValues
            nouvelleFeuille.Range("A1").PasteSpecial xlPasteFormats
            nouvelleFeuille.Columns.AutoFit

            nbFeuillesExportees = nbFeuillesExportees + 1
        End If
    Next feuille

    Application.CutCopyMode = False

    ' Supprimer la feuille par défaut (après avoir ajouté les autres)
    If nbFeuillesExportees > 0 Then
        Application.DisplayAlerts = False
        nouveauClasseur.Worksheets("Feuil1").Delete
        Application.DisplayAlerts = True
    End If

    If nbFeuillesExportees > 0 Then
        cheminSauvegarde = Application.GetSaveAsFilename( _
            InitialFilename:="Export_Complet_" & Format(Now, "yyyymmdd") & ".xlsx", _
            FileFilter:="Fichiers Excel (*.xlsx), *.xlsx")

        If cheminSauvegarde <> "False" Then
            nouveauClasseur.SaveAs cheminSauvegarde
            MsgBox nbFeuillesExportees & " feuilles exportées vers : " & cheminSauvegarde
        Else
            nouveauClasseur.Close False
        End If
    Else
        nouveauClasseur.Close False
        MsgBox "Aucune feuille à exporter"
    End If
End Sub
```

## Export vers CSV

### Export CSV simple

```vba
Sub ExporterVersCSV()
    Dim cheminSauvegarde As String
    Dim feuille As Worksheet
    Dim ligne As Long, col As Long
    Dim texte As String
    Dim numeroFichier As Integer

    Set feuille = ActiveSheet

    ' Vérifier qu'il y a des données
    If feuille.UsedRange.Rows.Count = 0 Then
        MsgBox "Aucune donnée à exporter !"
        Exit Sub
    End If

    ' Choisir le fichier de destination
    cheminSauvegarde = Application.GetSaveAsFilename( _
        InitialFilename:="Export_" & Format(Now, "yyyymmdd") & ".csv", _
        FileFilter:="Fichiers CSV (*.csv), *.csv")

    If cheminSauvegarde = "False" Then Exit Sub

    ' Ouvrir le fichier pour écriture
    numeroFichier = FreeFile
    Open cheminSauvegarde For Output As #numeroFichier

    ' Écrire les données ligne par ligne
    For ligne = 1 To feuille.UsedRange.Rows.Count
        texte = ""

        For col = 1 To feuille.UsedRange.Columns.Count
            Dim valeur As String
            valeur = CStr(feuille.Cells(ligne, col).Value)

            ' Échapper les virgules et guillemets
            If InStr(valeur, ",") > 0 Or InStr(valeur, """") > 0 Then
                valeur = """" & Replace(valeur, """", """""") & """"
            End If

            texte = texte & valeur
            If col < feuille.UsedRange.Columns.Count Then texte = texte & ","
        Next col

        Print #numeroFichier, texte
    Next ligne

    Close #numeroFichier

    MsgBox "Export CSV réussi : " & feuille.UsedRange.Rows.Count & " lignes exportées"
End Sub
```

### Export CSV avec paramètres

```vba
Sub ExporterCSVPersonnalise()
    Dim separateur As String
    Dim avecEntetes As Boolean

    ' Paramètres utilisateur
    separateur = InputBox("Séparateur (virgule, point-virgule, tabulation) ?", "Export CSV", ",")
    If separateur = "tabulation" Then separateur = vbTab

    avecEntetes = (MsgBox("Inclure les en-têtes ?", vbYesNo + vbQuestion) = vbYes)

    Dim cheminSauvegarde As String
    cheminSauvegarde = Application.GetSaveAsFilename( _
        "Export_Personnalise_" & Format(Now, "yyyymmdd") & ".csv", _
        "Fichiers CSV (*.csv), *.csv")

    If cheminSauvegarde = "False" Then Exit Sub

    Dim feuille As Worksheet
    Set feuille = ActiveSheet

    Dim numeroFichier As Integer
    numeroFichier = FreeFile
    Open cheminSauvegarde For Output As #numeroFichier

    Dim ligneDebut As Long
    ligneDebut = IIf(avecEntetes, 1, 2)

    Dim ligne As Long, col As Long
    For ligne = ligneDebut To feuille.UsedRange.Rows.Count
        Dim texte As String
        texte = ""

        For col = 1 To feuille.UsedRange.Columns.Count
            Dim valeur As String
            valeur = CStr(feuille.Cells(ligne, col).Value)

            texte = texte & valeur
            If col < feuille.UsedRange.Columns.Count Then texte = texte & separateur
        Next col

        Print #numeroFichier, texte
    Next ligne

    Close #numeroFichier

    MsgBox "Export CSV personnalisé terminé"
End Sub
```

## Export vers base de données

### Export vers Access

```vba
Sub ExporterVersAccess()
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim feuille As Worksheet
    Dim ligne As Long
    Dim cheminBase As String

    Set feuille = ActiveSheet

    ' Sélectionner la base de données
    cheminBase = Application.GetOpenFilename( _
        "Bases Access (*.accdb), *.accdb," & _
        "Anciennes bases (*.mdb), *.mdb", _
        , "Sélectionnez la base de destination")

    If cheminBase = "False" Then Exit Sub

    ' Demander le nom de la table
    Dim nomTable As String
    nomTable = InputBox("Nom de la table de destination ?", "Export", "NouvelleTable")
    If nomTable = "" Then Exit Sub

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminBase

    On Error GoTo GestionErreur

    ' Créer la table si elle n'existe pas (exemple simple)
    On Error Resume Next
    sql = "CREATE TABLE " & nomTable & " (" & _
          "ID AUTOINCREMENT PRIMARY KEY, " & _
          "Colonne1 TEXT(255), " & _
          "Colonne2 TEXT(255), " & _
          "Colonne3 TEXT(255))"
    conn.Execute sql
    On Error GoTo GestionErreur

    ' Vider la table
    conn.Execute "DELETE FROM " & nomTable

    ' Insérer les données (en ignorant la première ligne d'en-têtes)
    For ligne = 2 To feuille.UsedRange.Rows.Count
        Dim val1 As String, val2 As String, val3 As String
        val1 = SecuriserChaine(CStr(feuille.Cells(ligne, 1).Value))
        val2 = SecuriserChaine(CStr(feuille.Cells(ligne, 2).Value))
        val3 = SecuriserChaine(CStr(feuille.Cells(ligne, 3).Value))

        sql = "INSERT INTO " & nomTable & " (Colonne1, Colonne2, Colonne3) " & _
              "VALUES ('" & val1 & "', '" & val2 & "', '" & val3 & "')"

        conn.Execute sql
    Next ligne

    conn.Close
    Set conn = Nothing

    MsgBox "Export vers Access réussi : " & (feuille.UsedRange.Rows.Count - 1) & " lignes exportées"
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'export : " & Err.Description
    If Not conn Is Nothing Then conn.Close
End Sub

Function SecuriserChaine(texte As String) As String
    SecuriserChaine = Replace(texte, "'", "''")
End Function
```

### Export en lot (batch)

```vba
Sub ExportEnLot()
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim feuille As Worksheet
    Dim ligne As Long

    Set feuille = ActiveSheet
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Utiliser une transaction pour optimiser
    conn.BeginTrans

    On Error GoTo AnnulerTransaction

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandText = "INSERT INTO Clients (Nom, Email, Ville) VALUES (?, ?, ?)"

    ' Préparer les paramètres
    cmd.Parameters.Append cmd.CreateParameter("Nom", adVarChar, adParamInput, 255)
    cmd.Parameters.Append cmd.CreateParameter("Email", adVarChar, adParamInput, 255)
    cmd.Parameters.Append cmd.CreateParameter("Ville", adVarChar, adParamInput, 255)

    ' Insérer en lot
    For ligne = 2 To feuille.UsedRange.Rows.Count
        cmd.Parameters("Nom").Value = feuille.Cells(ligne, 1).Value
        cmd.Parameters("Email").Value = feuille.Cells(ligne, 2).Value
        cmd.Parameters("Ville").Value = feuille.Cells(ligne, 3).Value

        cmd.Execute

        ' Afficher le progrès tous les 100 enregistrements
        If ligne Mod 100 = 0 Then
            Application.StatusBar = "Export en cours : " & ligne & " lignes traitées"
        End If
    Next ligne

    ' Confirmer la transaction
    conn.CommitTrans

    Application.StatusBar = False
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing

    MsgBox "Export en lot terminé : " & (feuille.UsedRange.Rows.Count - 1) & " lignes"
    Exit Sub

AnnulerTransaction:
    conn.RollbackTrans
    Application.StatusBar = False
    MsgBox "Erreur lors de l'export en lot : " & Err.Description
    If Not conn Is Nothing Then conn.Close
End Sub
```

## Automatisation et planification

### Import automatique quotidien

```vba
Sub ImportQuotidienAutomatique()
    Dim cheminFichier As String
    Dim dateDuJour As String

    ' Construire le nom du fichier selon la date
    dateDuJour = Format(Date, "yyyymmdd")
    cheminFichier = "C:\Rapports\Ventes_" & dateDuJour & ".xlsx"

    ' Vérifier si le fichier existe
    If Dir(cheminFichier) = "" Then
        ' Essayer avec la date d'hier
        dateDuJour = Format(Date - 1, "yyyymmdd")
        cheminFichier = "C:\Rapports\Ventes_" & dateDuJour & ".xlsx"

        If Dir(cheminFichier) = "" Then
            MsgBox "Aucun fichier de ventes trouvé pour aujourd'hui ou hier"
            Exit Sub
        End If
    End If

    ' Import automatique
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes"""

    Set rs = conn.Execute("SELECT * FROM [Ventes$]")

    ' Nettoyer la feuille de destination
    Worksheets("Dashboard").Cells.Clear

    ' Copier les données
    Dim ligne As Long, col As Integer
    ligne = 1

    ' En-têtes
    For col = 0 To rs.Fields.Count - 1
        Worksheets("Dashboard").Cells(ligne, col + 1).Value = rs.Fields(col).Name
    Next col
    ligne = 2

    ' Données
    Do While Not rs.EOF
        For col = 0 To rs.Fields.Count - 1
            Worksheets("Dashboard").Cells(ligne, col + 1).Value = rs.Fields(col).Value
        Next col
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' Mettre à jour le timestamp
    Worksheets("Dashboard").Range("A1").AddComment "Dernière mise à jour : " & Now()

    ' Sauvegarder
    ThisWorkbook.Save

    MsgBox "Import automatique terminé : " & (ligne - 2) & " lignes importées"
End Sub
```

### Export planifié avec horodatage

```vba
Sub ExportPlanifieAvecHorodatage()
    Dim cheminExport As String
    Dim horodatage As String
    Dim nomFichier As String
    Dim nbFeuillesCopiees As Integer

    ' Créer un nom de fichier unique
    horodatage = Format(Now, "yyyymmdd_hhmmss")
    nomFichier = "Export_Dashboard_" & horodatage & ".xlsx"
    cheminExport = "C:\Exports\" & nomFichier

    ' Créer le dossier s'il n'existe pas
    If Dir("C:\Exports\", vbDirectory) = "" Then
        MkDir "C:\Exports\"
    End If

    ' Copier le classeur
    Dim nouveauClasseur As Workbook
    Set nouveauClasseur = Workbooks.Add

    nbFeuillesCopiees = 0

    ' Copier seulement les feuilles de données (pas les calculs)
    Dim feuille As Worksheet
    For Each feuille In ThisWorkbook.Worksheets
        If Left(feuille.Name, 1) <> "_" Then ' Ignorer les feuilles qui commencent par _
            feuille.Copy After:=nouveauClasseur.Sheets(nouveauClasseur.Sheets.Count)
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
    Next feuille

    ' Supprimer les feuilles par défaut
    Application.DisplayAlerts = False
    Do While nouveauClasseur.Sheets.Count > nbFeuillesCopiees
        nouveauClasseur.Sheets(1).Delete
    Loop
    Application.DisplayAlerts = True

    ' Sauvegarder
    nouveauClasseur.SaveAs cheminExport
    nouveauClasseur.Close

    Set nouveauClasseur = Nothing

    ' Log de l'export
    Dim logFile As String
    logFile = "C:\Exports\log_exports.txt"

    Dim numeroFichier As Integer
    numeroFichier = FreeFile
    Open logFile For Append As #numeroFichier
    Print #numeroFichier, Now() & " - Export réussi : " & nomFichier
    Close #numeroFichier

    MsgBox "Export planifié terminé : " & cheminExport
End Sub
```

## Gestion des erreurs et reprise

### Import avec gestion d'erreurs robuste

```vba
Sub ImportAvecGestionErreurs()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cheminFichier As String
    Dim nbLignesReussies As Long
    Dim nbLignesEchouees As Long
    Dim ligneErreurs As String

    cheminFichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    If cheminFichier = "False" Then Exit Sub

    Set conn = New ADODB.Connection

    On Error GoTo GestionErreurConnexion
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""

    Set rs = conn.Execute("SELECT * FROM [Feuil1$]")

    nbLignesReussies = 0
    nbLignesEchouees = 0
    ligneErreurs = ""

    Dim ligne As Long
    ligne = 2 ' Commencer après les en-têtes

    Do While Not rs.EOF
        On Error GoTo GestionErreurLigne

        ' Traitement de la ligne
        Dim col As Integer
        For col = 0 To rs.Fields.Count - 1
            Cells(ligne, col + 1).Value = rs.Fields(col).Value
        Next col

        nbLignesReussies = nbLignesReussies + 1
        GoTo LigneSuivante

GestionErreurLigne:
        nbLignesEchouees = nbLignesEchouees + 1
        ligneErreurs = ligneErreurs & ligne & " (" & Err.Description & "), "

LigneSuivante:
        On Error GoTo GestionErreurConnexion
        rs.MoveNext
        ligne = ligne + 1
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' Rapport final
    Dim message As String
    message = "Import terminé :" & vbCrLf & _
              "Lignes réussies : " & nbLignesReussies & vbCrLf & _
              "Lignes échouées : " & nbLignesEchouees

    If nbLignesEchouees > 0 Then
        message = message & vbCrLf & "Lignes en erreur : " & Left(ligneErreurs, Len(ligneErreurs) - 2)
    End If

    MsgBox message
    Exit Sub

GestionErreurConnexion:
    MsgBox "Erreur de connexion : " & Err.Description
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub
```

### Reprise après interruption

```vba
Sub ImportAvecReprise()
    Dim cheminFichier As String
    Dim derniereLigne As Long
    Dim fichierReprise As String

    cheminFichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    If cheminFichier = "False" Then Exit Sub

    ' Fichier de sauvegarde de la progression
    fichierReprise = ThisWorkbook.Path & "\reprise_import.txt"

    ' Vérifier s'il y a une reprise en cours
    If Dir(fichierReprise) <> "" Then
        Dim numeroFichier As Integer
        numeroFichier = FreeFile
        Open fichierReprise For Input As #numeroFichier
        Line Input #numeroFichier, derniereLigne
        Close #numeroFichier

        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Import interrompu détecté à la ligne " & derniereLigne & _
                        ". Reprendre ?", vbYesNoCancel + vbQuestion)

        If reponse = vbCancel Then Exit Sub
        If reponse = vbNo Then derniereLigne = 0
    Else
        derniereLigne = 0
    End If

    ' Import avec sauvegarde de progression
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes"""

    Set rs = conn.Execute("SELECT * FROM [Feuil1$]")

    ' Se positionner à la bonne ligne
    Dim ligne As Long
    ligne = 1
    Do While ligne <= derniereLigne And Not rs.EOF
        rs.MoveNext
        ligne = ligne + 1
    Loop

    ' Continuer l'import
    Do While Not rs.EOF
        ' Sauvegarder la progression tous les 50 enregistrements
        If ligne Mod 50 = 0 Then
            numeroFichier = FreeFile
            Open fichierReprise For Output As #numeroFichier
            Print #numeroFichier, ligne
            Close #numeroFichier

            Application.StatusBar = "Import en cours : ligne " & ligne
        End If

        ' Traitement de la ligne
        Dim col As Integer
        For col = 0 To rs.Fields.Count - 1
            Cells(ligne + 1, col + 1).Value = rs.Fields(col).Value
        Next col

        rs.MoveNext
        ligne = ligne + 1
    Loop

    ' Supprimer le fichier de reprise
    If Dir(fichierReprise) <> "" Then Kill fichierReprise

    Application.StatusBar = False
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Import terminé avec reprise : " & ligne & " lignes traitées"
End Sub
```

## Optimisation des performances

### Import par chunks (morceaux)

```vba
Sub ImportParChunks()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cheminFichier As String
    Dim tailleChunk As Long
    Dim ligneActuelle As Long

    cheminFichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    If cheminFichier = "False" Then Exit Sub

    tailleChunk = 1000 ' Traiter par blocs de 1000 lignes
    ligneActuelle = 1

    ' Désactiver les mises à jour d'écran pour la performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes"""

    ' Charger toutes les lignes en une seule requête
    Set rs = conn.Execute("SELECT * FROM [Feuil1$]")

    ' Traitement par chunks : écrire dans Excel par blocs
    Do While Not rs.EOF
        Dim col As Integer
        For col = 0 To rs.Fields.Count - 1
            Cells(ligneActuelle, col + 1).Value = rs.Fields(col).Value
        Next col

        ligneActuelle = ligneActuelle + 1
        rs.MoveNext

        ' Rafraîchir l'affichage tous les tailleChunk enregistrements
        If ligneActuelle Mod tailleChunk = 0 Then
            Application.StatusBar = "Import en cours : " & ligneActuelle & " lignes traitées"
            DoEvents ' Permettre à l'interface de se rafraîchir
        End If
    Loop

    rs.Close

    ' Réactiver les fonctionnalités
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Import par chunks terminé : " & ligneActuelle & " lignes"
End Sub
```

### Export optimisé avec tableau

```vba
Sub ExportOptimiseAvecTableau()
    Dim feuille As Worksheet
    Dim plageData As Range
    Dim tableauData As Variant
    Dim cheminSauvegarde As String

    Set feuille = ActiveSheet
    Set plageData = feuille.UsedRange

    ' Charger toutes les données en mémoire d'un coup
    tableauData = plageData.Value

    cheminSauvegarde = Application.GetSaveAsFilename( _
        "Export_Optimise_" & Format(Now, "yyyymmdd") & ".csv", _
        "Fichiers CSV (*.csv), *.csv")

    If cheminSauvegarde = "False" Then Exit Sub

    ' Écriture optimisée
    Dim numeroFichier As Integer
    Dim ligne As Long, col As Long
    Dim texte As String

    numeroFichier = FreeFile
    Open cheminSauvegarde For Output As #numeroFichier

    ' Traiter le tableau en mémoire (plus rapide)
    For ligne = 1 To UBound(tableauData, 1)
        texte = ""
        For col = 1 To UBound(tableauData, 2)
            Dim valeur As String
            valeur = CStr(tableauData(ligne, col))

            ' Échapper si nécessaire
            If InStr(valeur, ",") > 0 Then
                valeur = """" & Replace(valeur, """", """""") & """"
            End If

            texte = texte & valeur
            If col < UBound(tableauData, 2) Then texte = texte & ","
        Next col

        Print #numeroFichier, texte

        ' Progrès tous les 1000 enregistrements
        If ligne Mod 1000 = 0 Then
            Application.StatusBar = "Export : " & ligne & "/" & UBound(tableauData, 1)
        End If
    Next ligne

    Close #numeroFichier
    Application.StatusBar = False

    MsgBox "Export optimisé terminé : " & UBound(tableauData, 1) & " lignes"
End Sub
```

## Validation et contrôle qualité

### Import avec validation

```vba
Sub ImportAvecValidation()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cheminFichier As String
    Dim nbLignesValides As Long
    Dim nbLignesInvalides As Long
    Dim erreursValidation As String

    cheminFichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    If cheminFichier = "False" Then Exit Sub

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes"""

    Set rs = conn.Execute("SELECT * FROM [Feuil1$]")

    nbLignesValides = 0
    nbLignesInvalides = 0
    erreursValidation = ""

    Dim ligne As Long
    ligne = 2

    Do While Not rs.EOF
        Dim estValide As Boolean
        Dim messageErreur As String

        estValide = True
        messageErreur = ""

        ' Validation 1 : Email obligatoire et format correct
        Dim email As String
        email = CStr(rs.Fields("Email").Value)
        If email = "" Then
            estValide = False
            messageErreur = messageErreur & "Email manquant; "
        ElseIf InStr(email, "@") = 0 Then
            estValide = False
            messageErreur = messageErreur & "Email invalide; "
        End If

        ' Validation 2 : Âge doit être un nombre entre 18 et 120
        Dim age As Variant
        age = rs.Fields("Age").Value
        If Not IsNumeric(age) Then
            estValide = False
            messageErreur = messageErreur & "Âge non numérique; "
        ElseIf age < 18 Or age > 120 Then
            estValide = False
            messageErreur = messageErreur & "Âge hors limites; "
        End If

        ' Validation 3 : Nom obligatoire
        Dim nom As String
        nom = CStr(rs.Fields("NomClient").Value)
        If Trim(nom) = "" Then
            estValide = False
            messageErreur = messageErreur & "Nom manquant; "
        End If

        If estValide Then
            ' Importer la ligne
            Cells(ligne, 1).Value = nom
            Cells(ligne, 2).Value = email
            Cells(ligne, 3).Value = age
            Cells(ligne, 4).Value = "OK"
            nbLignesValides = nbLignesValides + 1
        Else
            ' Marquer l'erreur
            Cells(ligne, 1).Value = nom
            Cells(ligne, 2).Value = email
            Cells(ligne, 3).Value = age
            Cells(ligne, 4).Value = "ERREUR: " & messageErreur
            Cells(ligne, 4).Font.Color = RGB(255, 0, 0) ' Rouge
            nbLignesInvalides = nbLignesInvalides + 1
        End If

        rs.MoveNext
        ligne = ligne + 1
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' Rapport de validation
    Dim message As String
    message = "Import avec validation terminé :" & vbCrLf & _
              "Lignes valides : " & nbLignesValides & vbCrLf & _
              "Lignes invalides : " & nbLignesInvalides & vbCrLf & vbCrLf & _
              "Consultez la colonne D pour les détails des erreurs."

    MsgBox message
End Sub
```

## Utilitaires et fonctions helper

### Fonction de détection automatique du type de fichier

```vba
Function DetecterTypeFichier(cheminFichier As String) As String
    ' Extraire l'extension complète (après le dernier point)
    Dim extension As String
    extension = LCase(Mid(cheminFichier, InStrRev(cheminFichier, ".")))

    Select Case extension
        Case ".xls"
            DetecterTypeFichier = "Excel2003"
        Case ".xlsx"
            DetecterTypeFichier = "Excel2007+"
        Case ".csv"
            DetecterTypeFichier = "CSV"
        Case ".txt"
            DetecterTypeFichier = "Texte"
        Case ".accdb"
            DetecterTypeFichier = "Access2007+"
        Case ".mdb"
            DetecterTypeFichier = "Access2003"
        Case Else
            DetecterTypeFichier = "Inconnu"
    End Select
End Function

Sub ImportUniversel()
    Dim cheminFichier As String
    Dim typeFichier As String

    cheminFichier = Application.GetOpenFilename( _
        "Tous fichiers supportés,*.xlsx;*.xls;*.csv;*.txt;*.accdb;*.mdb")

    If cheminFichier = "False" Then Exit Sub

    typeFichier = DetecterTypeFichier(cheminFichier)

    Select Case typeFichier
        Case "Excel2007+", "Excel2003"
            ImporterDepuisExcel cheminFichier
        Case "CSV"
            ImporterDepuisCSV cheminFichier
        Case "Access2007+", "Access2003"
            ImporterDepuisAccess cheminFichier
        Case Else
            MsgBox "Type de fichier non supporté : " & typeFichier
    End Select
End Sub

Sub ImporterDepuisExcel(cheminFichier As String)
    ' Votre code d'import Excel ici
    MsgBox "Import Excel depuis : " & cheminFichier
End Sub

Sub ImporterDepuisCSV(cheminFichier As String)
    ' Votre code d'import CSV ici
    MsgBox "Import CSV depuis : " & cheminFichier
End Sub

Sub ImporterDepuisAccess(cheminFichier As String)
    ' Votre code d'import Access ici
    MsgBox "Import Access depuis : " & cheminFichier
End Sub
```

### Gestionnaire de progression visuel

```vba
Sub ImportAvecBarreProgression()
    Dim cheminFichier As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim totalLignes As Long
    Dim ligneActuelle As Long

    cheminFichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    If cheminFichier = "False" Then Exit Sub

    ' Première passe : compter les lignes
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes"""

    Set rs = conn.Execute("SELECT COUNT(*) AS Total FROM [Feuil1$]")
    totalLignes = rs.Fields("Total").Value
    rs.Close

    ' Deuxième passe : import avec progression
    Set rs = conn.Execute("SELECT * FROM [Feuil1$]")

    ligneActuelle = 0

    Do While Not rs.EOF
        ligneActuelle = ligneActuelle + 1

        ' Traitement de la ligne
        Dim col As Integer
        For col = 0 To rs.Fields.Count - 1
            Cells(ligneActuelle + 1, col + 1).Value = rs.Fields(col).Value
        Next col

        ' Mise à jour de la progression
        Dim pourcentage As Double
        pourcentage = (ligneActuelle / totalLignes) * 100

        Application.StatusBar = "Import en cours : " & _
                               Format(pourcentage, "0.0") & "% (" & _
                               ligneActuelle & "/" & totalLignes & ")"

        ' Mettre à jour tous les 100 enregistrements
        If ligneActuelle Mod 100 = 0 Then
            DoEvents ' Permettre à l'interface de se rafraîchir
        End If

        rs.MoveNext
    Loop

    Application.StatusBar = False
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Import terminé : " & totalLignes & " lignes importées"
End Sub
```

## Résumé des bonnes pratiques

✅ **Toujours vérifier** l'existence des fichiers avant import

✅ **Gérer les erreurs** à tous les niveaux (connexion, ligne, validation)

✅ **Optimiser les performances** : désactiver calculs et affichage

✅ **Valider les données** avant import définitif

✅ **Sauvegarder la progression** pour les gros volumes

✅ **Utiliser des transactions** pour garantir la cohérence

✅ **Nettoyer les objets** ADO après utilisation

✅ **Prévoir la reprise** après interruption

✅ **Logger les opérations** pour traçabilité

✅ **Tester avec de petits volumes** avant production

## Points d'attention pour débutants

🚨 **Chemins de fichiers** : Toujours utiliser des chemins complets

🚨 **Formats de dates** : Attention aux différences entre systèmes

🚨 **Caractères spéciaux** : Échapper les guillemets et virgules dans CSV

🚨 **Mémoire** : Les gros volumes peuvent saturer la mémoire

🚨 **Permissions** : Vérifier les droits d'accès aux fichiers/dossiers

🚨 **Versions Excel** : Certains providers ne fonctionnent qu'en 32 ou 64 bits

🚨 **Sauvegarde** : Toujours sauvegarder avant import massif

---

*Avec ces techniques d'import/export, vous pouvez maintenant automatiser complètement vos flux de données ! La prochaine étape sera d'intégrer Power Query pour des transformations encore plus puissantes.*

⏭️
