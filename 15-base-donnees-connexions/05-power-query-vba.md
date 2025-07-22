🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.5 Power Query et VBA

## Introduction à Power Query

Power Query est comme un super-assistant pour Excel qui sait nettoyer, transformer et organiser vos données automatiquement. Imaginez que vous avez un employé très efficace qui peut prendre des données en désordre et les ranger parfaitement selon vos instructions, et qui peut refaire ce travail autant de fois que nécessaire sans jamais se tromper.

Avec VBA, vous pouvez contrôler Power Query de manière programmatique, combinant ainsi la puissance de transformation des données avec l'automatisation du code. C'est comme avoir un chef d'orchestre (VBA) qui dirige un virtuose (Power Query) !

## Pourquoi combiner Power Query et VBA ?

### Avantages de Power Query
- **Interface visuelle** : Transformations par glisser-déposer
- **Apprentissage automatique** : Power Query "apprend" vos étapes
- **Performance** : Optimisé pour traiter de gros volumes
- **Connecteurs** : Se connecte à de nombreuses sources de données
- **Reproductibilité** : Les étapes se répètent automatiquement

### Pourquoi ajouter VBA ?
- **Automatisation complète** : Déclencher les actualisations automatiquement
- **Logique conditionnelle** : Adapter les transformations selon le contexte
- **Interface utilisateur** : Créer des boutons et menus personnalisés
- **Intégration** : Combiner avec d'autres processus VBA
- **Planification** : Exécuter selon des horaires spécifiques

## Concepts de base de Power Query

### Qu'est-ce qu'une requête Power Query ?

Une requête Power Query est une série d'étapes qui transforment des données brutes en données propres et structurées. Pensez-y comme une recette de cuisine :

1. **Source** : Les ingrédients (données brutes)
2. **Étapes** : Les instructions de préparation (transformations)
3. **Résultat** : Le plat final (données nettoyées)

### Le langage M

Power Query utilise un langage appelé "M" pour décrire les transformations. Ne vous inquiétez pas, vous n'avez pas besoin de l'apprendre ! Power Query écrit le code M automatiquement quand vous utilisez l'interface visuelle.

## Accéder à Power Query depuis VBA

### Les objets principaux

Dans VBA, Power Query est accessible via plusieurs objets :

```vba
' WorkbookConnection : Représente une connexion de données
Dim conn As WorkbookConnection

' QueryTable : Pour les requêtes qui créent des tableaux
Dim qt As QueryTable

' ListObject : Pour les tableaux Excel liés aux requêtes
Dim tbl As ListObject
```

### Identifier les requêtes existantes

```vba
Sub ListerRequetesPowerQuery()
    Dim conn As WorkbookConnection
    Dim i As Integer

    Debug.Print "=== Requêtes Power Query dans ce classeur ==="

    i = 1
    For Each conn In ThisWorkbook.Connections
        ' Vérifier si c'est une requête Power Query
        If conn.Type = xlConnectionTypeOLEDB Or conn.Type = xlConnectionTypeODBC Then
            Debug.Print i & ". " & conn.Name
            Debug.Print "   Type: " & TypeConnexion(conn.Type)
            Debug.Print "   Description: " & conn.Description
            Debug.Print "   ---"
            i = i + 1
        End If
    Next conn

    If i = 1 Then
        Debug.Print "Aucune requête Power Query trouvée"
    End If
End Sub

Function TypeConnexion(typeConn As XlConnectionType) As String
    Select Case typeConn
        Case xlConnectionTypeOLEDB
            TypeConnexion = "OLEDB (Power Query)"
        Case xlConnectionTypeODBC
            TypeConnexion = "ODBC"
        Case xlConnectionTypeWEB
            TypeConnexion = "Web"
        Case xlConnectionTypeTEXT
            TypeConnexion = "Texte"
        Case Else
            TypeConnexion = "Autre"
    End Select
End Function
```

## Actualiser les requêtes Power Query

### Actualisation simple

```vba
Sub ActualiserToutesLesRequetes()
    ' Actualiser toutes les connexions de données
    ThisWorkbook.RefreshAll

    MsgBox "Toutes les requêtes ont été actualisées !"
End Sub
```

### Actualisation d'une requête spécifique

```vba
Sub ActualiserRequeteSpecifique()
    Dim nomRequete As String
    Dim conn As WorkbookConnection
    Dim trouve As Boolean

    ' Demander le nom de la requête
    nomRequete = InputBox("Nom de la requête à actualiser ?", "Actualisation", "")

    If nomRequete = "" Then Exit Sub

    trouve = False

    ' Rechercher et actualiser la requête
    For Each conn In ThisWorkbook.Connections
        If conn.Name = nomRequete Then
            On Error GoTo GestionErreur

            conn.Refresh
            trouve = True
            MsgBox "Requête '" & nomRequete & "' actualisée avec succès !"
            Exit For
        End If
    Next conn

    If Not trouve Then
        MsgBox "Requête '" & nomRequete & "' non trouvée !"
    End If

    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'actualisation : " & Err.Description
End Sub
```

### Actualisation avec gestion d'événements

```vba
Sub ActualiserAvecSuivi()
    Dim conn As WorkbookConnection
    Dim nbRequetes As Integer
    Dim nbReussies As Integer
    Dim nbEchouees As Integer

    nbRequetes = ThisWorkbook.Connections.Count
    nbReussies = 0
    nbEchouees = 0

    If nbRequetes = 0 Then
        MsgBox "Aucune requête à actualiser"
        Exit Sub
    End If

    Application.StatusBar = "Actualisation en cours..."

    For Each conn In ThisWorkbook.Connections
        On Error GoTo RequeteSuivante

        Application.StatusBar = "Actualisation de : " & conn.Name
        conn.Refresh
        nbReussies = nbReussies + 1
        GoTo ContinuerBoucle

RequeteSuivante:
        nbEchouees = nbEchouees + 1
        Debug.Print "Erreur avec " & conn.Name & ": " & Err.Description

ContinuerBoucle:
        ' Permettre à l'utilisateur de voir le progrès
        DoEvents
    Next conn

    Application.StatusBar = False

    ' Rapport final
    Dim message As String
    message = "Actualisation terminée :" & vbCrLf & _
              "Réussies : " & nbReussies & vbCrLf & _
              "Échouées : " & nbEchouees & vbCrLf & _
              "Total : " & nbRequetes

    MsgBox message
End Sub
```

## Contrôler les paramètres des requêtes

### Modifier les paramètres d'une requête

Certaines requêtes Power Query acceptent des paramètres que vous pouvez modifier depuis VBA :

```vba
Sub ModifierParametreRequete()
    Dim conn As WorkbookConnection
    Dim nomRequete As String
    Dim nouveauParametre As String

    nomRequete = "MaRequeteParametree" ' Nom de votre requête
    nouveauParametre = InputBox("Nouvelle valeur du paramètre ?", "Paramètre")

    If nouveauParametre = "" Then Exit Sub

    ' Rechercher la connexion
    For Each conn In ThisWorkbook.Connections
        If conn.Name = nomRequete Then
            ' Modifier la formule de la requête (exemple simplifié)
            ' Note : Cette approche dépend de la structure de votre requête

            On Error GoTo GestionErreur

            ' Actualiser avec le nouveau paramètre
            conn.Refresh

            MsgBox "Paramètre mis à jour et requête actualisée !"
            Exit Sub
        End If
    Next conn

    MsgBox "Requête '" & nomRequete & "' non trouvée !"
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de la modification : " & Err.Description
End Sub
```

### Gestion des sources de données dynamiques

```vba
Sub ChangerSourceDonnees()
    Dim conn As WorkbookConnection
    Dim nomRequete As String
    Dim nouvellSource As String
    Dim ancienneFormule As String
    Dim nouvelleFormule As String

    nomRequete = InputBox("Nom de la requête à modifier ?")
    If nomRequete = "" Then Exit Sub

    nouvellSource = Application.GetOpenFilename( _
        "Fichiers Excel (*.xlsx), *.xlsx," & _
        "Fichiers CSV (*.csv), *.csv", _
        , "Sélectionnez la nouvelle source")

    If nouvellSource = "False" Then Exit Sub

    ' Rechercher la connexion
    For Each conn In ThisWorkbook.Connections
        If conn.Name = nomRequete Then
            On Error GoTo GestionErreur

            ' Sauvegarder l'ancienne formule
            ancienneFormule = conn.OLEDBConnection.CommandText

            ' Créer la nouvelle formule (exemple pour Excel)
            nouvelleFormule = Replace(ancienneFormule, _
                              "Source = Excel.Workbook(File.Contents(""", _
                              "Source = Excel.Workbook(File.Contents(""" & nouvellSource & """))")

            ' Appliquer la nouvelle formule
            conn.OLEDBConnection.CommandText = nouvelleFormule

            ' Actualiser
            conn.Refresh

            MsgBox "Source de données modifiée et requête actualisée !"
            Exit Sub
        End If
    Next conn

    MsgBox "Requête non trouvée !"
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors du changement de source : " & Err.Description
End Sub
```

## Créer des requêtes Power Query via VBA

### Créer une connexion simple

```vba
Sub CreerRequeteCSV()
    Dim cheminFichier As String
    Dim nomRequete As String
    Dim conn As WorkbookConnection

    ' Sélectionner le fichier CSV
    cheminFichier = Application.GetOpenFilename( _
        "Fichiers CSV (*.csv), *.csv", _
        , "Sélectionnez un fichier CSV")

    If cheminFichier = "False" Then Exit Sub

    nomRequete = "ImportCSV_" & Format(Now, "yyyymmdd_hhmmss")

    On Error GoTo GestionErreur

    ' Créer la connexion Power Query
    Set conn = ThisWorkbook.Connections.Add2( _
        Name:=nomRequete, _
        Description:="Import CSV via VBA", _
        ConnectionString:="", _
        CommandText:="", _
        lCmdtype:=0)

    ' Configuration spécifique pour CSV
    With conn.OLEDBConnection
        .Connection = "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & cheminFichier
        .CommandType = xlCmdDefault
        .CommandText = nomRequete
    End With

    ' Créer le tableau résultant
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim qt As QueryTable
    Set qt = ws.QueryTables.Add( _
        Connection:=conn, _
        Destination:=ws.Range("A1"))

    ' Actualiser pour charger les données
    qt.Refresh

    MsgBox "Requête CSV créée : " & nomRequete
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de la création : " & Err.Description
End Sub
```

### Créer une requête Excel

```vba
Sub CreerRequeteExcel()
    Dim cheminSource As String
    Dim nomFeuille As String
    Dim nomRequete As String

    ' Paramètres de la requête
    cheminSource = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    If cheminSource = "False" Then Exit Sub

    nomFeuille = InputBox("Nom de la feuille source ?", "Requête Excel", "Feuil1")
    If nomFeuille = "" Then Exit Sub

    nomRequete = "ImportExcel_" & Format(Now, "yyyymmdd_hhmmss")

    On Error GoTo GestionErreur

    ' Créer la requête Power Query
    Dim formulePowerQuery As String
    formulePowerQuery = "let" & vbCrLf & _
                       "    Source = Excel.Workbook(File.Contents(""" & cheminSource & """), null, true)," & vbCrLf & _
                       "    " & nomFeuille & "_Sheet = Source{[Item=""" & nomFeuille & """,Kind=""Sheet""]}[Data]," & vbCrLf & _
                       "    #""Promoted Headers"" = Table.PromoteHeaders(" & nomFeuille & "_Sheet, [PromoteAllScalars=true])" & vbCrLf & _
                       "in" & vbCrLf & _
                       "    #""Promoted Headers"""

    ' Ajouter la connexion
    Dim conn As WorkbookConnection
    Set conn = ThisWorkbook.Connections.Add2( _
        Name:=nomRequete, _
        Description:="Import Excel via VBA", _
        ConnectionString:="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & cheminSource, _
        CommandText:=formulePowerQuery, _
        lCmdtype:=0)

    ' Créer le tableau
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcQuery, _
        Source:=conn, _
        Destination:=ws.Range("A1"))

    tbl.Name = "Tableau_" & nomRequete

    MsgBox "Requête Excel créée : " & nomRequete
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de la création : " & Err.Description
End Sub
```

## Gestion avancée des requêtes

### Actualisation asynchrone

```vba
Sub ActualisationAsynchrone()
    Dim conn As WorkbookConnection

    ' Désactiver les alertes
    Application.DisplayAlerts = False

    For Each conn In ThisWorkbook.Connections
        ' Lancer l'actualisation en arrière-plan
        conn.OLEDBConnection.BackgroundQuery = True
        conn.Refresh
    Next conn

    Application.DisplayAlerts = True

    MsgBox "Actualisation lancée en arrière-plan. " & _
           "Les données se mettront à jour automatiquement."
End Sub
```

### Surveillance de l'état des requêtes

```vba
Sub SurveillerEtatRequetes()
    Dim conn As WorkbookConnection
    Dim etatGlobal As String
    Dim nbEnCours As Integer

    etatGlobal = ""
    nbEnCours = 0

    For Each conn In ThisWorkbook.Connections
        If conn.OLEDBConnection.Refreshing Then
            etatGlobal = etatGlobal & conn.Name & " (en cours)" & vbCrLf
            nbEnCours = nbEnCours + 1
        Else
            etatGlobal = etatGlobal & conn.Name & " (terminé)" & vbCrLf
        End If
    Next conn

    If nbEnCours > 0 Then
        MsgBox "État des requêtes :" & vbCrLf & vbCrLf & etatGlobal & vbCrLf & _
               nbEnCours & " requête(s) en cours d'actualisation"
    Else
        MsgBox "Toutes les requêtes sont à jour !" & vbCrLf & vbCrLf & etatGlobal
    End If
End Sub
```

### Suppression de requêtes

```vba
Sub SupprimerRequete()
    Dim nomRequete As String
    Dim conn As WorkbookConnection
    Dim trouve As Boolean
    Dim reponse As VbMsgBoxResult

    nomRequete = InputBox("Nom de la requête à supprimer ?")
    If nomRequete = "" Then Exit Sub

    trouve = False

    For Each conn In ThisWorkbook.Connections
        If conn.Name = nomRequete Then
            trouve = True

            ' Demander confirmation
            reponse = MsgBox("Voulez-vous vraiment supprimer la requête '" & nomRequete & "' ?", _
                           vbYesNo + vbQuestion, "Confirmation")

            If reponse = vbYes Then
                conn.Delete
                MsgBox "Requête '" & nomRequete & "' supprimée"
            End If

            Exit For
        End If
    Next conn

    If Not trouve Then
        MsgBox "Requête '" & nomRequete & "' non trouvée"
    End If
End Sub
```

## Automatisation avec Power Query

### Actualisation planifiée

```vba
Sub ConfigurerActualisationAutomatique()
    Dim conn As WorkbookConnection
    Dim intervalleMinutes As Integer

    intervalleMinutes = Val(InputBox("Intervalle d'actualisation (en minutes) ?", "Planification", "60"))

    If intervalleMinutes <= 0 Then
        MsgBox "Intervalle invalide"
        Exit Sub
    End If

    For Each conn In ThisWorkbook.Connections
        With conn.OLEDBConnection
            .RefreshPeriod = intervalleMinutes
            .RefreshOnFileOpen = True
            .EnableRefresh = True
        End With
    Next conn

    MsgBox "Actualisation automatique configurée : toutes les " & intervalleMinutes & " minutes"
End Sub
```

### Actualisation conditionnelle

```vba
Sub ActualisationConditionnelle()
    Dim derniereActualisation As Date
    Dim intervalleHeures As Double
    Dim doitActualiser As Boolean

    ' Récupérer la dernière actualisation (stockée dans une cellule)
    On Error Resume Next
    derniereActualisation = CDate(Worksheets("Config").Range("B1").Value)
    On Error GoTo 0

    ' Si pas de date, considérer qu'il faut actualiser
    If derniereActualisation = 0 Then
        doitActualiser = True
    Else
        ' Vérifier si 4 heures se sont écoulées
        intervalleHeures = (Now - derniereActualisation) * 24
        doitActualiser = (intervalleHeures >= 4)
    End If

    If doitActualiser Then
        ' Actualiser toutes les requêtes
        ThisWorkbook.RefreshAll

        ' Enregistrer la nouvelle date
        Worksheets("Config").Range("B1").Value = Now

        MsgBox "Données actualisées !"
    Else
        MsgBox "Actualisation non nécessaire (dernière actualisation : " & _
               Format(derniereActualisation, "dd/mm/yyyy hh:mm") & ")"
    End If
End Sub
```

### Création d'un tableau de bord automatisé

```vba
Sub CreerTableauDeBordAutomatise()
    Dim ws As Worksheet
    Dim conn As WorkbookConnection
    Dim tbl As ListObject

    ' Créer une nouvelle feuille pour le tableau de bord
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Dashboard_" & Format(Now, "yyyymmdd")

    ' En-tête du tableau de bord
    ws.Range("A1").Value = "TABLEAU DE BORD"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "Dernière mise à jour : " & Now()

    ' Section état des requêtes
    ws.Range("A4").Value = "ÉTAT DES REQUÊTES"
    ws.Range("A4").Font.Bold = True

    Dim ligne As Integer
    ligne = 5

    ws.Range("A" & ligne).Value = "Nom"
    ws.Range("B" & ligne).Value = "État"
    ws.Range("C" & ligne).Value = "Dernière actualisation"
    ws.Range("A" & ligne & ":C" & ligne).Font.Bold = True
    ligne = ligne + 1

    ' Lister toutes les requêtes
    For Each conn In ThisWorkbook.Connections
        ws.Range("A" & ligne).Value = conn.Name

        If conn.OLEDBConnection.Refreshing Then
            ws.Range("B" & ligne).Value = "En cours..."
            ws.Range("B" & ligne).Font.Color = RGB(255, 165, 0) ' Orange
        Else
            ws.Range("B" & ligne).Value = "OK"
            ws.Range("B" & ligne).Font.Color = RGB(0, 128, 0) ' Vert
        End If

        On Error Resume Next
        ws.Range("C" & ligne).Value = conn.OLEDBConnection.RefreshDate
        On Error GoTo 0

        ligne = ligne + 1
    Next conn

    ' Formatage automatique
    ws.Columns.AutoFit

    ' Actualiser toutes les requêtes
    Application.StatusBar = "Actualisation du tableau de bord..."
    ThisWorkbook.RefreshAll
    Application.StatusBar = False

    MsgBox "Tableau de bord créé et actualisé !"
End Sub
```

## Gestion des erreurs Power Query

### Détection d'erreurs dans les requêtes

```vba
Sub VerifierErreursRequetes()
    Dim conn As WorkbookConnection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim erreursTrouvees As Boolean
    Dim rapportErreurs As String

    erreursTrouvees = False
    rapportErreurs = "RAPPORT D'ERREURS POWER QUERY" & vbCrLf & vbCrLf

    For Each conn In ThisWorkbook.Connections
        On Error Resume Next

        ' Tenter d'actualiser la requête
        conn.Refresh

        If Err.Number <> 0 Then
            erreursTrouvees = True
            rapportErreurs = rapportErreurs & "ERREUR - " & conn.Name & ":" & vbCrLf
            rapportErreurs = rapportErreurs & "  " & Err.Description & vbCrLf & vbCrLf

            Err.Clear
        End If

        On Error GoTo 0
    Next conn

    If erreursTrouvees Then
        ' Créer une feuille de rapport d'erreurs
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Erreurs_" & Format(Now, "yyyymmdd_hhmmss")
        ws.Range("A1").Value = rapportErreurs
        ws.Range("A1").WrapText = True
        ws.Columns.AutoFit

        MsgBox "Erreurs détectées ! Consultez la feuille : " & ws.Name
    Else
        MsgBox "Aucune erreur détectée. Toutes les requêtes fonctionnent correctement."
    End If
End Sub
```

### Réparation automatique de requêtes

```vba
Sub ReparerRequetesAutomatiquement()
    Dim conn As WorkbookConnection
    Dim nbReparees As Integer
    Dim nbEchecs As Integer

    nbReparees = 0
    nbEchecs = 0

    For Each conn In ThisWorkbook.Connections
        On Error GoTo RequeteSuivante

        ' Essayer de réactiver la requête
        conn.OLEDBConnection.EnableRefresh = True
        conn.OLEDBConnection.BackgroundQuery = False

        ' Tenter une actualisation
        conn.Refresh

        nbReparees = nbReparees + 1
        GoTo ContinuerBoucle

RequeteSuivante:
        nbEchecs = nbEchecs + 1
        Debug.Print "Impossible de réparer : " & conn.Name & " - " & Err.Description
        Err.Clear

ContinuerBoucle:
    Next conn

    MsgBox "Réparation terminée :" & vbCrLf & _
           "Requêtes réparées : " & nbReparees & vbCrLf & _
           "Échecs : " & nbEchecs
End Sub
```

## Optimisation des performances

### Actualisation optimisée

```vba
Sub ActualisationOptimisee()
    ' Désactiver les fonctionnalités pour améliorer les performances
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    Dim debut As Date
    debut = Now

    Application.StatusBar = "Actualisation Power Query en cours..."

    On Error GoTo NettoyageFinal

    ' Actualiser toutes les connexions
    ThisWorkbook.RefreshAll

    ' Attendre la fin de toutes les actualisations
    Dim toutesTerminees As Boolean
    Do
        toutesTerminees = True
        Dim conn As WorkbookConnection

        For Each conn In ThisWorkbook.Connections
            If conn.OLEDBConnection.Refreshing Then
                toutesTerminees = False
                Exit For
            End If
        Next conn

        If Not toutesTerminees Then
            DoEvents
            Application.Wait Now + TimeValue("00:00:01") ' Attendre 1 seconde
        End If
    Loop Until toutesTerminees

NettoyageFinal:
    ' Réactiver les fonctionnalités
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    Dim duree As Double
    duree = (Now - debut) * 24 * 60 ' En minutes

    MsgBox "Actualisation terminée en " & Format(duree, "0.0") & " minutes"
End Sub
```

### Surveillance de la mémoire

```vba
Sub SurveillerUtilisationMemoire()
    ' Note : Cette fonction nécessite des API Windows pour la mémoire complète
    ' Version simplifiée utilisant les informations disponibles en VBA

    Dim conn As WorkbookConnection
    Dim tailleApproximative As Long
    Dim ws As Worksheet

    Debug.Print "=== UTILISATION MÉMOIRE PAR REQUÊTE ==="

    For Each conn In ThisWorkbook.Connections
        tailleApproximative = 0

        ' Estimer la taille en comptant les cellules utilisées
        For Each ws In ThisWorkbook.Worksheets
            Dim tbl As ListObject
            For Each tbl In ws.ListObjects
                If tbl.QueryTable Is Nothing Then GoTo TableauSuivant
                If tbl.QueryTable.Connection.Name = conn.Name Then
                    tailleApproximative = tailleApproximative + _
                        (tbl.Range.Rows.Count * tbl.Range.Columns.Count * 50) ' 50 octets par cellule estimé
                End If
TableauSuivant:
            Next tbl
        Next ws

        Debug.Print conn.Name & ": ~" & Format(tailleApproximative / 1024, "#,##0") & " KB"
    Next conn

    MsgBox "Surveillance terminée. Consultez la fenêtre Exécution pour les détails."
End Sub
```

## Intégration avec d'autres systèmes

### Export automatique après actualisation

```vba
Private Sub Workbook_AfterRefresh(ByVal Success As Boolean)
    ' Cet événement se déclenche après chaque actualisation Power Query

    If Success Then
        ' Actualisation réussie - exporter automatiquement
        Call ExporterDonneesActualisees
    Else
        ' Actualisation échouée - alerter l'utilisateur
        MsgBox "Échec de l'actualisation Power Query. Vérifiez vos connexions."
    End If
End Sub

Sub ExporterDonneesActualisees()
    Dim cheminExport As String
    Dim horodatage As String

    horodatage = Format(Now, "yyyymmdd_hhmmss")
    cheminExport = "C:\Exports\Donnees_" & horodatage & ".xlsx"

    ' Créer le dossier si nécessaire
    If Dir("C:\Exports\", vbDirectory) = "" Then
        MkDir "C:\Exports\"
    End If

    ' Sauvegarder une copie
    ThisWorkbook.SaveCopyAs cheminExport

    ' Log de l'export
    Debug.Print Now & " - Export automatique : " & cheminExport
End Sub
```

### Notification par email

```vba
Sub EnvoyerNotificationActualisation()
    ' Nécessite une référence à Microsoft Outlook Object Library

    On Error GoTo GestionErreur

    Dim OutlookApp As Object
    Dim mail As Object

    Set OutlookApp = CreateObject("Outlook.Application")
    Set mail = OutlookApp.CreateItem(0) ' olMailItem

    With mail
        .To = "manager@entreprise.com"
        .Subject = "Données Power Query actualisées - " & Format(Now, "dd/mm/yyyy hh:mm")
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
                "Les données Power Query ont été actualisées avec succès." & vbCrLf & _
                "Heure d'actualisation : " & Format(Now, "dd/mm/yyyy à hh:mm") & vbCrLf & _
                "Nombre de requêtes traitées : " & ThisWorkbook.Connections.Count & vbCrLf & vbCrLf & _
                "Cordialement," & vbCrLf & _
                "Système automatisé Excel"

        ' Ajouter le fichier en pièce jointe (optionnel)
        .Attachments.Add ThisWorkbook.FullName

        ' Envoyer automatiquement ou afficher pour révision
        .Send ' Pour envoi automatique
        ' .Display ' Pour afficher avant envoi
    End With

    Set mail = Nothing
    Set OutlookApp = Nothing

    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'envoi de l'email : " & Err.Description
    Set mail = Nothing
    Set OutlookApp = Nothing
End Sub
```

## Exemples pratiques complets

### Système de reporting automatisé

```vba
Sub SystemeReportingAutomatise()
    ' Système complet d'actualisation et de reporting

    Dim debut As Date
    Dim fin As Date
    Dim duree As Double
    Dim rapportFinal As String

    debut = Now

    ' 1. Vérification préalable
    If Not VerifierConnexionsDisponibles() Then
        MsgBox "Certaines sources de données ne sont pas disponibles. Arrêt du processus."
        Exit Sub
    End If

    ' 2. Sauvegarde avant actualisation
    Dim cheminSauvegarde As String
    cheminSauvegarde = ThisWorkbook.Path & "\Sauvegarde_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    ThisWorkbook.SaveCopyAs cheminSauvegarde

    ' 3. Actualisation optimisée
    Application.ScreenUpdating = False
    Application.StatusBar = "Actualisation des données en cours..."

    Call ActualiserToutesLesRequetesAvecSuivi

    ' 4. Validation des données
    Dim resultatsValidation As String
    resultatsValidation = ValiderDonneesApresActualisation()

    ' 5. Génération du rapport
    Call GenererRapportAutomatique

    ' 6. Export et distribution
    Call ExporterEtDistribuer

    fin = Now
    duree = (fin - debut) * 24 * 60 ' En minutes

    ' 7. Rapport final
    rapportFinal = "RAPPORT D'ACTUALISATION AUTOMATISÉE" & vbCrLf & vbCrLf & _
                   "Début : " & Format(debut, "dd/mm/yyyy hh:mm:ss") & vbCrLf & _
                   "Fin : " & Format(fin, "dd/mm/yyyy hh:mm:ss") & vbCrLf & _
                   "Durée : " & Format(duree, "0.0") & " minutes" & vbCrLf & vbCrLf & _
                   resultatsValidation

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox rapportFinal

    ' 8. Notification (optionnelle)
    Call EnvoyerNotificationActualisation
End Sub

Function VerifierConnexionsDisponibles() As Boolean
    Dim conn As WorkbookConnection
    Dim toutesDisponibles As Boolean

    toutesDisponibles = True

    For Each conn In ThisWorkbook.Connections
        On Error Resume Next

        ' Tester la connexion
        conn.OLEDBConnection.MaintainConnection = True

        If Err.Number <> 0 Then
            Debug.Print "Connexion indisponible : " & conn.Name
            toutesDisponibles = False
            Err.Clear
        End If

        On Error GoTo 0
    Next conn

    VerifierConnexionsDisponibles = toutesDisponibles
End Function

Sub ActualiserToutesLesRequetesAvecSuivi()
    Dim conn As WorkbookConnection
    Dim i As Integer
    Dim total As Integer

    total = ThisWorkbook.Connections.Count
    i = 0

    For Each conn In ThisWorkbook.Connections
        i = i + 1
        Application.StatusBar = "Actualisation " & i & "/" & total & " : " & conn.Name

        On Error Resume Next
        conn.Refresh

        If Err.Number <> 0 Then
            Debug.Print "Erreur requête " & conn.Name & " : " & Err.Description
            Err.Clear
        End If

        On Error GoTo 0
        DoEvents
    Next conn
End Sub

Function ValiderDonneesApresActualisation() As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim validationOK As Boolean
    Dim rapport As String

    validationOK = True
    rapport = "VALIDATION DES DONNÉES :" & vbCrLf

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.ListRows.Count = 0 Then
                rapport = rapport & "⚠️ Table vide : " & tbl.Name & " (feuille " & ws.Name & ")" & vbCrLf
                validationOK = False
            Else
                rapport = rapport & "✅ " & tbl.Name & " : " & tbl.ListRows.Count & " lignes" & vbCrLf
            End If
        Next tbl
    Next ws

    If validationOK Then
        rapport = rapport & vbCrLf & "✅ Validation globale : SUCCÈS"
    Else
        rapport = rapport & vbCrLf & "❌ Validation globale : PROBLÈMES DÉTECTÉS"
    End If

    ValiderDonneesApresActualisation = rapport
End Function

Sub GenererRapportAutomatique()
    Dim wsRapport As Worksheet
    Dim ligne As Integer

    ' Créer ou vider la feuille de rapport
    On Error Resume Next
    Set wsRapport = ThisWorkbook.Worksheets("Rapport_Auto")
    On Error GoTo 0

    If wsRapport Is Nothing Then
        Set wsRapport = ThisWorkbook.Worksheets.Add
        wsRapport.Name = "Rapport_Auto"
    Else
        wsRapport.Cells.Clear
    End If

    ' En-tête du rapport
    ligne = 1
    wsRapport.Cells(ligne, 1).Value = "RAPPORT D'ACTUALISATION AUTOMATIQUE"
    wsRapport.Cells(ligne, 1).Font.Size = 16
    wsRapport.Cells(ligne, 1).Font.Bold = True

    ligne = ligne + 2
    wsRapport.Cells(ligne, 1).Value = "Date de génération :"
    wsRapport.Cells(ligne, 2).Value = Now()
    wsRapport.Cells(ligne, 2).NumberFormat = "dd/mm/yyyy hh:mm"

    ' Résumé des requêtes
    ligne = ligne + 2
    wsRapport.Cells(ligne, 1).Value = "RÉSUMÉ DES REQUÊTES"
    wsRapport.Cells(ligne, 1).Font.Bold = True

    ligne = ligne + 1
    wsRapport.Cells(ligne, 1).Value = "Nom de la requête"
    wsRapport.Cells(ligne, 2).Value = "État"
    wsRapport.Cells(ligne, 3).Value = "Dernière actualisation"
    wsRapport.Range("A" & ligne & ":C" & ligne).Font.Bold = True

    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        ligne = ligne + 1
        wsRapport.Cells(ligne, 1).Value = conn.Name

        If conn.OLEDBConnection.Refreshing Then
            wsRapport.Cells(ligne, 2).Value = "En cours"
            wsRapport.Cells(ligne, 2).Font.Color = RGB(255, 165, 0)
        Else
            wsRapport.Cells(ligne, 2).Value = "Terminé"
            wsRapport.Cells(ligne, 2).Font.Color = RGB(0, 128, 0)
        End If

        On Error Resume Next
        wsRapport.Cells(ligne, 3).Value = conn.OLEDBConnection.RefreshDate
        wsRapport.Cells(ligne, 3).NumberFormat = "dd/mm/yyyy hh:mm"
        On Error GoTo 0
    Next conn

    ' Formatage
    wsRapport.Columns.AutoFit
    wsRapport.Range("A1").Select
End Sub

Sub ExporterEtDistribuer()
    Dim cheminExport As String
    Dim horodatage As String

    horodatage = Format(Now, "yyyymmdd_hhmmss")

    ' Créer le dossier d'export
    If Dir("C:\Exports\", vbDirectory) = "" Then
        MkDir "C:\Exports\"
    End If

    ' Export Excel complet
    cheminExport = "C:\Exports\Rapport_Complet_" & horodatage & ".xlsx"
    ThisWorkbook.SaveCopyAs cheminExport

    ' Export CSV des données principales (exemple)
    Call ExporterTableauPrincipalEnCSV("C:\Exports\Donnees_Principales_" & horodatage & ".csv")

    Debug.Print "Fichiers exportés :"
    Debug.Print "- " & cheminExport
    Debug.Print "- C:\Exports\Donnees_Principales_" & horodatage & ".csv"
End Sub

Sub ExporterTableauPrincipalEnCSV(cheminCSV As String)
    ' Exporter le premier tableau trouvé en CSV
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim trouve As Boolean

    trouve = False

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.ListRows.Count > 0 Then
                ' Exporter ce tableau
                Call ExporterTableauEnCSV(tbl, cheminCSV)
                trouve = True
                Exit For
            End If
        Next tbl
        If trouve Then Exit For
    Next ws
End Sub

Sub ExporterTableauEnCSV(tbl As ListObject, cheminFichier As String)
    Dim numeroFichier As Integer
    Dim ligne As Long
    Dim col As Long
    Dim texte As String
    Dim valeur As String

    numeroFichier = FreeFile
    Open cheminFichier For Output As #numeroFichier

    ' En-têtes
    texte = ""
    For col = 1 To tbl.ListColumns.Count
        texte = texte & tbl.ListColumns(col).Name
        If col < tbl.ListColumns.Count Then texte = texte & ","
    Next col
    Print #numeroFichier, texte

    ' Données
    For ligne = 1 To tbl.ListRows.Count
        texte = ""
        For col = 1 To tbl.ListColumns.Count
            valeur = CStr(tbl.DataBodyRange.Cells(ligne, col).Value)

            ' Échapper les virgules
            If InStr(valeur, ",") > 0 Then
                valeur = """" & Replace(valeur, """", """""") & """"
            End If

            texte = texte & valeur
            If col < tbl.ListColumns.Count Then texte = texte & ","
        Next col
        Print #numeroFichier, texte
    Next ligne

    Close #numeroFichier
End Sub
```

### Dashboard de monitoring Power Query

```vba
Sub CreerDashboardMonitoring()
    Dim wsDashboard As Worksheet
    Dim ligne As Integer
    Dim col As Integer

    ' Créer ou réinitialiser la feuille dashboard
    On Error Resume Next
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard_PQ")
    On Error GoTo 0

    If wsDashboard Is Nothing Then
        Set wsDashboard = ThisWorkbook.Worksheets.Add
        wsDashboard.Name = "Dashboard_PQ"
    Else
        wsDashboard.Cells.Clear
    End If

    ' Configuration de la feuille
    With wsDashboard
        ' Titre principal
        .Range("A1").Value = "DASHBOARD POWER QUERY"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:E1").Merge
        .Range("A1").HorizontalAlignment = xlCenter

        ' Informations générales
        ligne = 3
        .Cells(ligne, 1).Value = "Dernière mise à jour :"
        .Cells(ligne, 2).Value = Now()
        .Cells(ligne, 2).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(ligne, 1).Font.Bold = True

        ligne = ligne + 1
        .Cells(ligne, 1).Value = "Nombre total de requêtes :"
        .Cells(ligne, 2).Value = ThisWorkbook.Connections.Count
        .Cells(ligne, 1).Font.Bold = True

        ' Section détaillée des requêtes
        ligne = ligne + 3
        .Cells(ligne, 1).Value = "DÉTAIL DES REQUÊTES"
        .Cells(ligne, 1).Font.Size = 14
        .Cells(ligne, 1).Font.Bold = True

        ' En-têtes du tableau
        ligne = ligne + 2
        .Cells(ligne, 1).Value = "Nom"
        .Cells(ligne, 2).Value = "Type"
        .Cells(ligne, 3).Value = "État"
        .Cells(ligne, 4).Value = "Dernière actualisation"
        .Cells(ligne, 5).Value = "Lignes"
        .Cells(ligne, 6).Value = "Action"
        .Range("A" & ligne & ":F" & ligne).Font.Bold = True
        .Range("A" & ligne & ":F" & ligne).Interior.Color = RGB(200, 200, 200)

        ' Données des requêtes
        Dim conn As WorkbookConnection
        For Each conn In ThisWorkbook.Connections
            ligne = ligne + 1

            ' Nom de la requête
            .Cells(ligne, 1).Value = conn.Name

            ' Type de connexion
            .Cells(ligne, 2).Value = TypeConnexion(conn.Type)

            ' État
            If conn.OLEDBConnection.Refreshing Then
                .Cells(ligne, 3).Value = "🔄 En cours"
                .Cells(ligne, 3).Font.Color = RGB(255, 165, 0)
            Else
                .Cells(ligne, 3).Value = "✅ OK"
                .Cells(ligne, 3).Font.Color = RGB(0, 128, 0)
            End If

            ' Dernière actualisation
            On Error Resume Next
            .Cells(ligne, 4).Value = conn.OLEDBConnection.RefreshDate
            .Cells(ligne, 4).NumberFormat = "dd/mm/yyyy hh:mm"
            On Error GoTo 0

            ' Nombre de lignes (estimation)
            .Cells(ligne, 5).Value = EstimerNombreLignes(conn.Name)

            ' Bouton d'action (lien hypertexte pour actualiser)
            .Hyperlinks.Add Anchor:=.Cells(ligne, 6), _
                           Address:="", _
                           SubAddress:="ActualiserRequeteSpecifique", _
                           TextToDisplay:="🔄 Actualiser"
        Next conn

        ' Section statistiques
        ligne = ligne + 3
        .Cells(ligne, 1).Value = "STATISTIQUES"
        .Cells(ligne, 1).Font.Size = 14
        .Cells(ligne, 1).Font.Bold = True

        ligne = ligne + 1
        Call AjouterStatistiques(wsDashboard, ligne)

        ' Formatage général
        .Columns.AutoFit
        .Range("A1").Select
    End With

    ' Actualiser automatiquement toutes les 5 minutes
    Application.OnTime Now + TimeValue("00:05:00"), "ActualiserDashboard"

    MsgBox "Dashboard Power Query créé ! Actualisation automatique toutes les 5 minutes."
End Sub

Function EstimerNombreLignes(nomRequete As String) As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim totalLignes As Long

    totalLignes = 0

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.QueryTable Is Nothing Then GoTo TableauSuivant
            If InStr(tbl.QueryTable.Connection, nomRequete) > 0 Then
                totalLignes = totalLignes + tbl.ListRows.Count
            End If
TableauSuivant:
        Next tbl
    Next ws

    EstimerNombreLignes = totalLignes
End Function

Sub AjouterStatistiques(ws As Worksheet, ligneDebut As Integer)
    Dim ligne As Integer
    Dim nbRequetesOK As Integer
    Dim nbRequetesErreur As Integer
    Dim totalLignes As Long
    Dim conn As WorkbookConnection

    ligne = ligneDebut
    nbRequetesOK = 0
    nbRequetesErreur = 0
    totalLignes = 0

    ' Calculer les statistiques
    For Each conn In ThisWorkbook.Connections
        If conn.OLEDBConnection.Refreshing Then
            ' Considérer comme en cours = OK
            nbRequetesOK = nbRequetesOK + 1
        Else
            ' Vérifier s'il y a eu des erreurs récentes
            On Error Resume Next
            If Err.Number <> 0 Then
                nbRequetesErreur = nbRequetesErreur + 1
                Err.Clear
            Else
                nbRequetesOK = nbRequetesOK + 1
            End If
            On Error GoTo 0
        End If

        totalLignes = totalLignes + EstimerNombreLignes(conn.Name)
    Next conn

    ' Afficher les statistiques
    ws.Cells(ligne, 1).Value = "Requêtes fonctionnelles :"
    ws.Cells(ligne, 2).Value = nbRequetesOK
    ws.Cells(ligne, 2).Font.Color = RGB(0, 128, 0)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Requêtes en erreur :"
    ws.Cells(ligne, 2).Value = nbRequetesErreur
    ws.Cells(ligne, 2).Font.Color = RGB(255, 0, 0)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Total de lignes traitées :"
    ws.Cells(ligne, 2).Value = totalLignes
    ws.Cells(ligne, 2).NumberFormat = "#,##0"
    ligne = ligne + 1

    ' Pourcentage de réussite
    Dim pourcentageReussite As Double
    If ThisWorkbook.Connections.Count > 0 Then
        pourcentageReussite = (nbRequetesOK / ThisWorkbook.Connections.Count) * 100
    End If

    ws.Cells(ligne, 1).Value = "Taux de réussite :"
    ws.Cells(ligne, 2).Value = pourcentageReussite
    ws.Cells(ligne, 2).NumberFormat = "0.0%"

    If pourcentageReussite >= 90 Then
        ws.Cells(ligne, 2).Font.Color = RGB(0, 128, 0)
    ElseIf pourcentageReussite >= 70 Then
        ws.Cells(ligne, 2).Font.Color = RGB(255, 165, 0)
    Else
        ws.Cells(ligne, 2).Font.Color = RGB(255, 0, 0)
    End If
End Sub

Sub ActualiserDashboard()
    ' Actualiser le dashboard automatiquement
    On Error Resume Next
    Call CreerDashboardMonitoring
    On Error GoTo 0

    ' Programmer la prochaine actualisation
    Application.OnTime Now + TimeValue("00:05:00"), "ActualiserDashboard"
End Sub
```

## Conseils et bonnes pratiques

### Gestion des versions Power Query

```vba
Sub VerifierVersionPowerQuery()
    ' Vérifier la compatibilité Power Query
    Dim versionExcel As String
    Dim supportePowerQuery As Boolean

    versionExcel = Application.Version

    ' Power Query est intégré depuis Excel 2016
    If Val(versionExcel) >= 16 Then
        supportePowerQuery = True
        MsgBox "Power Query est supporté dans cette version d'Excel (" & versionExcel & ")"
    Else
        supportePowerQuery = False
        MsgBox "Cette version d'Excel (" & versionExcel & ") ne supporte pas Power Query nativement. " & _
               "Vous devez installer le complément Power Query."
    End If
End Sub
```

### Optimisation des requêtes

```vba
Sub OptimiserRequetesPowerQuery()
    Dim conn As WorkbookConnection
    Dim recommandations As String

    recommandations = "RECOMMANDATIONS D'OPTIMISATION :" & vbCrLf & vbCrLf

    For Each conn In ThisWorkbook.Connections
        ' Vérifier les paramètres de performance
        With conn.OLEDBConnection
            If .BackgroundQuery = False Then
                recommandations = recommandations & "⚠️ " & conn.Name & " : Activer l'actualisation en arrière-plan" & vbCrLf
            End If

            If .RefreshPeriod = 0 Then
                recommandations = recommandations & "💡 " & conn.Name & " : Considérer l'actualisation automatique" & vbCrLf
            End If

            If .MaintainConnection = True Then
                recommandations = recommandations & "💡 " & conn.Name & " : Maintenir la connexion peut consommer de la mémoire" & vbCrLf
            End If
        End With
    Next conn

    If Len(recommandations) > 50 Then
        MsgBox recommandations
    Else
        MsgBox "Toutes les requêtes semblent optimisées !"
    End If
End Sub
```

## Résumé des bonnes pratiques

✅ **Toujours tester** les requêtes avec de petits volumes d'abord

✅ **Utiliser l'actualisation en arrière-plan** pour ne pas bloquer l'interface

✅ **Gérer les erreurs** de connexion et d'actualisation

✅ **Optimiser les performances** en désactivant les fonctionnalités non nécessaires

✅ **Documenter** les requêtes avec des noms et descriptions clairs

✅ **Sauvegarder** avant les modifications importantes

✅ **Surveiller** l'utilisation mémoire avec de gros volumes

✅ **Planifier** les actualisations selon les besoins métier

✅ **Valider** les données après chaque actualisation

✅ **Créer des tableaux de bord** pour surveiller l'état des requêtes

## Points d'attention pour débutants

🚨 **Versions Excel** : Power Query nécessite Excel 2016 ou plus récent

🚨 **Performances** : Les gros volumes peuvent être lents à actualiser

🚨 **Connexions** : Vérifier que les sources de données sont accessibles

🚨 **Mémoire** : Surveiller l'utilisation RAM avec de nombreuses requêtes

🚨 **Sauvegarde** : Toujours sauvegarder avant modifications importantes

🚨 **Droits d'accès** : S'assurer d'avoir les permissions sur les sources

🚨 **Planification** : Éviter les actualisations simultanées multiples

---

*Power Query combiné à VBA vous donne un contrôle total sur vos flux de données ! Vous pouvez maintenant créer des solutions d'automatisation complètes et robustes.*

⏭️
