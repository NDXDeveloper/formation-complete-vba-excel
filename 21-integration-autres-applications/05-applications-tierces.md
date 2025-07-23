🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.5 Applications tierces

## Introduction aux applications tierces

L'automation avec des applications tierces permet d'étendre considérablement les capacités d'Excel en interagissant avec des logiciels externes : navigateurs web, applications PDF, logiciels de comptabilité, systèmes de gestion, etc. Cette fonctionnalité ouvre des possibilités infinies d'intégration.

## Qu'est-ce qu'une application tierce ?

### Définition
Une **application tierce** est tout logiciel qui n'appartient pas à la suite Microsoft Office mais qui peut être contrôlé depuis Excel via VBA. Cela inclut :
- Navigateurs web (Internet Explorer, Chrome, Edge)
- Lecteurs/créateurs PDF (Adobe Acrobat)
- Logiciels de comptabilité (SAP, Sage)
- Applications spécialisées métier
- Services web et APIs

### Méthodes d'interaction
Il existe plusieurs façons d'interagir avec des applications tierces :
1. **COM Automation** (comme avec Office)
2. **API Windows** (contrôle système)
3. **Web scraping** (extraction de données web)
4. **Fichiers intermédiaires** (CSV, XML, JSON)
5. **Services web** (REST, SOAP)

## Automation avec Internet Explorer

### Premier exemple simple

```vba
Sub PremierTestNavigateur()
    ' Créer une instance d'Internet Explorer
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")

    ' Rendre le navigateur visible
    ie.Visible = True

    ' Naviguer vers une page web
    ie.Navigate "https://www.google.com"

    ' Attendre que la page se charge
    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop

    MsgBox "Page Google chargée !"

    ' Fermer le navigateur
    ie.Quit
    Set ie = Nothing
End Sub
```

**Explication :**
- `CreateObject("InternetExplorer.Application")` : Lance Internet Explorer
- `ie.Visible = True` : Rend le navigateur visible
- `ie.Navigate` : Charge une page web
- La boucle `Do While` attend que la page soit complètement chargée
- `ie.Quit` : Ferme le navigateur

### Remplir un formulaire web automatiquement

```vba
Sub RemplirFormulaireWeb()
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True

    ' Aller sur une page avec un formulaire de recherche
    ie.Navigate "https://www.google.com"

    ' Attendre le chargement
    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop

    ' Trouver la zone de recherche et y entrer du texte
    Dim searchBox As Object
    Set searchBox = ie.Document.getElementsByName("q")(0)  ' "q" est le nom du champ de recherche Google

    ' Récupérer le terme de recherche depuis Excel
    Dim termeRecherche As String
    termeRecherche = Range("A1").Value

    ' Remplir et soumettre
    searchBox.Value = termeRecherche
    searchBox.Form.Submit

    MsgBox "Recherche lancée pour : " & termeRecherche

    ' Ne pas fermer immédiatement pour voir le résultat
    Set searchBox = Nothing
    Set ie = Nothing
End Sub
```

### Extraire des données depuis une page web

```vba
Sub ExtraireTableauWeb()
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False  ' Invisible pour un traitement en arrière-plan

    ' Exemple : extraire des données depuis une page financière
    ie.Navigate "https://finance.yahoo.com"

    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop

    ' Extraire le titre de la page
    Dim titrePage As String
    titrePage = ie.Document.Title

    ' Chercher des tableaux dans la page
    Dim tableaux As Object
    Set tableaux = ie.Document.getElementsByTagName("table")

    ' Vider la feuille Excel
    Cells.Clear

    ' Afficher les informations extraites
    Range("A1").Value = "Titre de la page :"
    Range("B1").Value = titrePage
    Range("A2").Value = "Nombre de tableaux trouvés :"
    Range("B2").Value = tableaux.Length

    ie.Quit
    Set tableaux = Nothing
    Set ie = Nothing

    MsgBox "Données extraites avec succès !"
End Sub
```

## Automation avec Microsoft Edge (méthode moderne)

```vba
Sub UtiliserEdge()
    ' Edge nécessite une approche différente via Selenium ou EdgeDriver
    ' Voici un exemple conceptuel

    Dim edge As Object

    ' Tenter de créer une instance Edge
    On Error Resume Next
    Set edge = CreateObject("MsEdge.Application")

    If Err.Number <> 0 Then
        MsgBox "Edge automation nécessite des composants spéciaux." & vbCrLf & _
               "Utilisez Internet Explorer pour débuter."
        Exit Sub
    End If
    On Error GoTo 0

    edge.Visible = True
    edge.Navigate "https://www.microsoft.com"

    ' Logique similaire à IE...

    Set edge = Nothing
End Sub
```

## Création et manipulation de fichiers PDF

### Générer un PDF depuis Excel (via PDFCreator virtuel)

```vba
Sub CreerPDFDepuisExcel()
    ' Sélectionner la plage à convertir
    Dim plageAPDF As Range
    Set plageAPDF = Range("A1:E20")

    ' Chemin de destination
    Dim cheminPDF As String
    cheminPDF = Environ("USERPROFILE") & "\Desktop\Rapport_" & Format(Date, "yyyy-mm-dd") & ".pdf"

    ' Exporter en PDF (fonction Excel native)
    plageAPDF.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=cheminPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProps:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "PDF créé : " & cheminPDF
End Sub
```

### Ouvrir un PDF avec l'application par défaut

```vba
Sub OuvrirPDFExterne()
    Dim cheminPDF As String
    cheminPDF = "C:\MonDossier\MonDocument.pdf"

    ' Vérifier si le fichier existe
    If Dir(cheminPDF) <> "" Then
        ' Ouvrir avec l'application par défaut
        Shell "explorer.exe """ & cheminPDF & """", vbNormalFocus
        MsgBox "PDF ouvert avec l'application par défaut"
    Else
        MsgBox "Fichier PDF introuvable : " & cheminPDF
    End If
End Sub
```

## Interaction avec des applications Windows génériques

### Trouver et activer une fenêtre d'application

```vba
' Déclarations API Windows (à placer en haut du module)
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As LongPtr) As Long

Sub ActiverApplicationExterne()
    ' Exemple : activer la Calculatrice Windows
    Dim hwnd As LongPtr

    ' Chercher la fenêtre de la calculatrice
    hwnd = FindWindow(vbNullString, "Calculatrice")

    If hwnd <> 0 Then
        ' Amener la fenêtre au premier plan
        SetForegroundWindow hwnd
        MsgBox "Calculatrice activée"
    Else
        ' Lancer la calculatrice si elle n'est pas ouverte
        Shell "calc.exe", vbNormalFocus
        MsgBox "Calculatrice lancée"
    End If
End Sub
```

### Envoyer des touches à une application externe

```vba
Sub EnvoyerTouchesApplication()
    ' Lancer le Bloc-notes
    Shell "notepad.exe", vbNormalFocus

    ' Attendre un peu que l'application se lance
    Application.Wait Now + TimeValue("00:00:02")

    ' Envoyer du texte au Bloc-notes
    SendKeys "Bonjour depuis Excel !", True
    SendKeys "{ENTER}", True
    SendKeys "Ceci est généré automatiquement.", True
    SendKeys "{ENTER}", True
    SendKeys "Date : " & Date, True

    MsgBox "Texte envoyé au Bloc-notes"
End Sub
```

## Interaction avec des fichiers CSV/XML/JSON

### Lire un fichier CSV depuis une application externe

```vba
Sub LireFichierCSVExterne()
    ' Chemin vers un fichier CSV généré par une autre application
    Dim cheminCSV As String
    cheminCSV = "C:\Exports\DonneesExternes.csv"

    ' Vérifier l'existence du fichier
    If Dir(cheminCSV) = "" Then
        MsgBox "Fichier CSV introuvable. Vérifiez que l'application externe a généré le fichier."
        Exit Sub
    End If

    ' Ouvrir le fichier CSV
    Workbooks.Open cheminCSV

    ' Copier les données vers notre classeur principal
    Dim sourceWB As Workbook
    Set sourceWB = ActiveWorkbook

    ' Copier toutes les données
    sourceWB.Sheets(1).UsedRange.Copy

    ' Retourner à notre classeur et coller
    ThisWorkbook.Activate
    Range("A1").PasteSpecial Paste:=xlPasteValues

    ' Fermer le fichier CSV
    sourceWB.Close SaveChanges:=False

    ' Nettoyer
    Application.CutCopyMode = False
    Set sourceWB = Nothing

    MsgBox "Données CSV importées avec succès"
End Sub
```

### Créer un fichier XML pour une application externe

```vba
Sub CreerFichierXMLPourExterne()
    Dim cheminXML As String
    cheminXML = Environ("USERPROFILE") & "\Desktop\ExportVersExterne.xml"

    ' Créer le contenu XML
    Dim contenuXML As String
    contenuXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    contenuXML = contenuXML & "<donnees>" & vbCrLf

    ' Ajouter les données Excel (supposons A2:C10)
    Dim i As Integer
    For i = 2 To 10
        If Cells(i, 1).Value <> "" Then
            contenuXML = contenuXML & "  <enregistrement>" & vbCrLf
            contenuXML = contenuXML & "    <nom>" & Cells(i, 1).Value & "</nom>" & vbCrLf
            contenuXML = contenuXML & "    <valeur>" & Cells(i, 2).Value & "</valeur>" & vbCrLf
            contenuXML = contenuXML & "    <date>" & Format(Cells(i, 3).Value, "yyyy-mm-dd") & "</date>" & vbCrLf
            contenuXML = contenuXML & "  </enregistrement>" & vbCrLf
        End If
    Next i

    contenuXML = contenuXML & "</donnees>"

    ' Écrire le fichier
    Dim numeroFichier As Integer
    numeroFichier = FreeFile

    Open cheminXML For Output As numeroFichier
    Print #numeroFichier, contenuXML
    Close numeroFichier

    MsgBox "Fichier XML créé : " & cheminXML & vbCrLf & _
           "Prêt à être utilisé par l'application externe"
End Sub
```

## Surveillance de dossiers et fichiers

### Vérifier l'apparition de nouveaux fichiers

```vba
Sub SurveillerNouveauxFichiers()
    Dim dossierSurveille As String
    dossierSurveille = "C:\DossierEchange\"

    ' Lister les fichiers actuels
    Dim fichierActuel As String
    Dim listeFichiers As String

    fichierActuel = Dir(dossierSurveille & "*.*")
    Do While fichierActuel <> ""
        listeFichiers = listeFichiers & fichierActuel & "|"
        fichierActuel = Dir
    Loop

    ' Afficher la liste
    MsgBox "Fichiers trouvés dans " & dossierSurveille & ":" & vbCrLf & _
           Replace(listeFichiers, "|", vbCrLf)

    ' Pour une surveillance continue, il faudrait utiliser un Timer
    ' ou vérifier périodiquement dans une boucle
End Sub
```

### Traitement automatique de fichiers déposés

```vba
Sub TraiterFichiersDeposes()
    Dim dossierEntree As String
    Dim dossierTraite As String

    dossierEntree = "C:\DossierEntree\"
    dossierTraite = "C:\DossierTraite\"

    ' Créer le dossier de traitement s'il n'existe pas
    If Dir(dossierTraite, vbDirectory) = "" Then
        MkDir dossierTraite
    End If

    Dim fichier As String
    fichier = Dir(dossierEntree & "*.csv")

    Do While fichier <> ""
        ' Traiter chaque fichier CSV trouvé
        Dim cheminComplet As String
        cheminComplet = dossierEntree & fichier

        ' Ouvrir et traiter le fichier
        Workbooks.Open cheminComplet

        ' Faire le traitement nécessaire...
        ' (ajout de colonnes, calculs, formatage, etc.)
        Range("A1").Value = "Traité le " & Now()

        ' Sauvegarder dans le dossier de traitement
        ActiveWorkbook.SaveAs dossierTraite & "Traite_" & fichier
        ActiveWorkbook.Close

        ' Déplacer le fichier original (optionnel)
        Name cheminComplet As dossierTraite & "Original_" & fichier

        ' Fichier suivant
        fichier = Dir
    Loop

    MsgBox "Traitement terminé. Vérifiez le dossier : " & dossierTraite
End Sub
```

## Communication via ligne de commande

### Exécuter des commandes système

```vba
Sub ExecuterCommandeSysteme()
    Dim commande As String
    Dim resultat As String

    ' Exemple : obtenir la liste des processus en cours
    commande = "tasklist /fo csv"

    ' Créer un fichier temporaire pour récupérer le résultat
    Dim fichierTemp As String
    fichierTemp = Environ("TEMP") & "\ResultatCommande.txt"

    ' Exécuter la commande et rediriger vers le fichier
    Shell "cmd /c " & commande & " > """ & fichierTemp & """", vbHide

    ' Attendre un peu que la commande se termine
    Application.Wait Now + TimeValue("00:00:03")

    ' Lire le résultat
    If Dir(fichierTemp) <> "" Then
        Dim numeroFichier As Integer
        numeroFichier = FreeFile

        Open fichierTemp For Input As numeroFichier
        resultat = Input(LOF(numeroFichier), numeroFichier)
        Close numeroFichier

        ' Afficher les premières lignes
        MsgBox "Premières lignes du résultat :" & vbCrLf & Left(resultat, 500) & "..."

        ' Supprimer le fichier temporaire
        Kill fichierTemp
    Else
        MsgBox "Erreur : impossible de récupérer le résultat de la commande"
    End If
End Sub
```

## Gestion d'erreurs pour applications tierces

```vba
Sub GestionErreursApplicationsTierces()
    Dim appExterne As Object

    On Error GoTo GestionErreur

    ' Tentative de connexion à une application qui pourrait ne pas exister
    Set appExterne = CreateObject("ApplicationInconnue.Application")

    ' Code normal si l'application existe...
    appExterne.Visible = True

    Set appExterne = Nothing
    Exit Sub

GestionErreur:
    Select Case Err.Number
        Case 429
            MsgBox "L'application tierce n'est pas disponible ou pas installée"
        Case 70
            MsgBox "Permission refusée - l'application est peut-être déjà ouverte"
        Case 53
            MsgBox "Fichier ou composant introuvable"
        Case Else
            MsgBox "Erreur inattendue : " & Err.Number & " - " & Err.Description
    End Select

    ' Nettoyage
    If Not appExterne Is Nothing Then
        Set appExterne = Nothing
    End If
End Sub
```

## Exemple complet : Système d'échange avec logiciel de comptabilité

```vba
Sub SystemeEchangeComptabilite()
    Dim dossierExport As String
    Dim dossierImport As String
    Dim fichierEcriture As String

    On Error GoTo GestionErreur

    ' === CONFIGURATION ===
    dossierExport = "C:\Compta\Export\"
    dossierImport = "C:\Compta\Import\"
    fichierEcriture = "Ecritures_" & Format(Date, "yyyy-mm-dd") & ".csv"

    ' Créer les dossiers s'ils n'existent pas
    If Dir(dossierExport, vbDirectory) = "" Then MkDir dossierExport
    If Dir(dossierImport, vbDirectory) = "" Then MkDir dossierImport

    ' === EXPORT DES DONNÉES EXCEL ===
    ' Supposons des écritures comptables en colonnes A à F
    Dim derniereLigne As Integer
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    If derniereLigne < 2 Then
        MsgBox "Aucune donnée à exporter"
        Exit Sub
    End If

    ' Créer le fichier CSV pour la comptabilité
    Dim cheminExport As String
    cheminExport = dossierExport & fichierEcriture

    Dim numeroFichier As Integer
    numeroFichier = FreeFile

    Open cheminExport For Output As numeroFichier

    ' En-tête CSV (format attendu par le logiciel de comptabilité)
    Print #numeroFichier, "Date;Compte;Libelle;Debit;Credit;Reference"

    ' Données
    Dim i As Integer
    For i = 2 To derniereLigne
        Dim ligne As String
        ligne = Format(Cells(i, 1).Value, "dd/mm/yyyy") & ";"  ' Date
        ligne = ligne & Cells(i, 2).Value & ";"               ' Compte
        ligne = ligne & Cells(i, 3).Value & ";"               ' Libellé
        ligne = ligne & Replace(Cells(i, 4).Value, ".", ",") & ";"  ' Débit
        ligne = ligne & Replace(Cells(i, 5).Value, ".", ",") & ";"  ' Crédit
        ligne = ligne & Cells(i, 6).Value                     ' Référence

        Print #numeroFichier, ligne
    Next i

    Close numeroFichier

    ' === LANCEMENT DU LOGICIEL DE COMPTABILITÉ ===
    ' Exemple : lancer un logiciel avec le fichier en paramètre
    Dim commandeLogiciel As String
    commandeLogiciel = """C:\Program Files\MonLogicielCompta\Compta.exe"" /import """ & cheminExport & """"

    ' Tenter de lancer (si le logiciel supporte les paramètres de ligne de commande)
    On Error Resume Next
    Shell commandeLogiciel, vbNormalFocus

    If Err.Number <> 0 Then
        ' Si l'auto-lancement échoue, ouvrir manuellement le dossier
        Shell "explorer.exe """ & dossierExport & """", vbNormalFocus
        MsgBox "Fichier d'export créé : " & cheminExport & vbCrLf & _
               "Importez-le manuellement dans votre logiciel de comptabilité."
    Else
        MsgBox "Logiciel de comptabilité lancé avec le fichier d'import"
    End If
    On Error GoTo GestionErreur

    ' === VÉRIFICATION DES RETOURS ===
    ' Attendre et vérifier si le logiciel a généré un fichier de retour
    Application.Wait Now + TimeValue("00:00:10")  ' Attendre 10 secondes

    Dim fichierRetour As String
    fichierRetour = Dir(dossierImport & "Retour_*.csv")

    If fichierRetour <> "" Then
        ' Traiter le fichier de retour
        Workbooks.Open dossierImport & fichierRetour

        ' Copier les informations de retour dans une nouvelle feuille
        Dim nouvelleFeuille As Worksheet
        Set nouvelleFeuille = ThisWorkbook.Worksheets.Add
        nouvelleFeuille.Name = "Retour_Compta_" & Format(Date, "mmdd")

        ActiveWorkbook.Sheets(1).UsedRange.Copy
        nouvelleFeuille.Range("A1").PasteSpecial Paste:=xlPasteValues

        ActiveWorkbook.Close SaveChanges:=False
        Application.CutCopyMode = False

        MsgBox "Échange terminé avec succès !" & vbCrLf & _
               "• Export : " & (derniereLigne - 1) & " écritures" & vbCrLf & _
               "• Retour importé dans la feuille : " & nouvelleFeuille.Name
    Else
        MsgBox "Export terminé. En attente du retour du logiciel de comptabilité..."
    End If

    ' === ARCHIVAGE ===
    ' Marquer les lignes exportées
    Range("G2:G" & derniereLigne).Value = "Exporté le " & Now()

    Exit Sub

GestionErreur:
    MsgBox "Erreur dans l'échange avec la comptabilité : " & Err.Description

    ' Nettoyage
    If numeroFichier > 0 Then Close numeroFichier
End Sub
```

## Points importants à retenir

### ✅ Bonnes pratiques
- Toujours vérifier l'existence des applications tierces avant de les utiliser
- Gérer les erreurs spécifiques à chaque type d'application
- Utiliser des formats d'échange standard (CSV, XML, JSON)
- Tester les intégrations sur de petits volumes d'abord
- Documenter les dépendances externes

### ⚠️ Erreurs courantes à éviter
- Supposer qu'une application tierce est toujours installée
- Ne pas gérer les versions différentes des applications
- Oublier les délais nécessaires pour que les applications externes répondent
- Utiliser des chemins de fichiers codés en dur
- Ne pas nettoyer les fichiers temporaires

### 🔧 Considérations techniques
- **Sécurité** : Les applications tierces peuvent présenter des risques
- **Performance** : L'automation d'applications externes peut être lente
- **Maintenance** : Les mises à jour des applications peuvent casser l'automation
- **Compatibilité** : Tester sur différents environnements

### 💡 Conseils pour débuter
- Commencez par des applications simples (Bloc-notes, Calculatrice)
- Utilisez d'abord les fonctions Windows standard (Shell, SendKeys)
- Testez l'automation manuelle avant d'automatiser
- Gardez des solutions de repli (fichiers d'échange)
- Créez des logs pour diagnostiquer les problèmes

### 🎯 Cas d'usage typiques
- **Échange de données** : Import/export avec logiciels métier
- **Web scraping** : Extraction de données depuis des sites web
- **Automatisation de tâches** : Contrôle d'applications de bureau
- **Intégration système** : Communication avec des services Windows
- **Surveillance** : Monitoring de fichiers et processus

### 🌐 Alternatives modernes
- **API REST** : Plus fiable que l'automation directe
- **Services web** : Integration via XML/JSON
- **Bases de données partagées** : Échange via SQL
- **Cloud services** : APIs Microsoft Graph, Google APIs

L'automation avec des applications tierces est puissante mais complexe. Elle nécessite une bonne compréhension des applications cibles et une gestion d'erreurs robuste. Commencez simple et évoluez progressivement vers des intégrations plus sophistiquées !

⏭️
