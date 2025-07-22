🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.2 Connexion à des bases de données externes

## Introduction

Maintenant que vous maîtrisez les bases d'ADO, il est temps d'apprendre à vous connecter aux différents types de bases de données que vous rencontrerez dans le monde professionnel. Chaque type de base de données a ses propres spécificités, comme des dialectes d'une même langue.

Imaginez que vous voulez visiter différents pays : même si vous parlez anglais, vous devrez vous adapter aux accents locaux et aux expressions particulières. C'est exactement ce qui se passe avec les bases de données !

## Vue d'ensemble des types de bases de données

### Bases de données locales
- **Microsoft Access** : Fichiers .accdb ou .mdb sur votre ordinateur
- **SQLite** : Bases de données légères dans un fichier unique
- **Excel** : Autres fichiers Excel comme sources de données

### Bases de données serveur
- **SQL Server** : La solution Microsoft pour l'entreprise
- **MySQL** : Base de données open source très populaire
- **Oracle** : Solution enterprise pour les grandes organisations
- **PostgreSQL** : Alternative open source robuste

### Bases de données cloud
- **Azure SQL Database** : SQL Server dans le cloud Microsoft
- **Amazon RDS** : Services de bases de données Amazon
- **Google Cloud SQL** : Solutions Google

## Microsoft Access

Access est souvent le premier contact avec les bases de données. C'est comme le "petit frère" d'Excel, mais spécialisé dans le stockage structuré.

### Chaîne de connexion Access

```vba
' Pour Access 2007 et plus récent (.accdb)
Dim connectionString As String
connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=C:\MesDocuments\MaBase.accdb"

' Pour Access 2003 et antérieur (.mdb)
connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=C:\MesDocuments\MaBase.mdb"
```

### Exemple complet Access

```vba
Sub ConnexionAccess()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connectionString As String

    ' Configuration de la connexion
    Set conn = New ADODB.Connection
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=C:\MaBase.accdb"

    On Error GoTo GestionErreur

    ' Connexion à la base Access
    conn.Open connectionString

    ' Vérification de la connexion
    If conn.State = adStateOpen Then
        MsgBox "Connexion réussie à Access !"

        ' Récupération des données
        Set rs = conn.Execute("SELECT * FROM Clients ORDER BY NomClient")

        ' Affichage dans Excel
        Dim ligne As Long
        ligne = 1

        Do While Not rs.EOF
            Cells(ligne, 1).Value = rs.Fields("NomClient").Value
            Cells(ligne, 2).Value = rs.Fields("Email").Value
            ligne = ligne + 1
            rs.MoveNext
        Loop

        rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur de connexion: " & Err.Description
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
End Sub
```

### Access avec mot de passe

```vba
connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=C:\MaBase.accdb;" & _
                  "Jet OLEDB:Database Password=MonMotDePasse"
```

## Microsoft SQL Server

SQL Server est la solution de base de données enterprise de Microsoft. C'est comme passer d'une voiture familiale à un camion : plus puissant, mais aussi plus complexe.

### Types de connexion SQL Server

#### Authentification Windows (recommandée)
```vba
connectionString = "Provider=SQLOLEDB;" & _
                  "Data Source=MonServeur;" & _
                  "Initial Catalog=MaBaseDeDonnees;" & _
                  "Integrated Security=SSPI"
```

#### Authentification SQL Server
```vba
connectionString = "Provider=SQLOLEDB;" & _
                  "Data Source=MonServeur;" & _
                  "Initial Catalog=MaBaseDeDonnees;" & _
                  "User ID=MonUtilisateur;" & _
                  "Password=MonMotDePasse"
```

### Exemple complet SQL Server

```vba
Sub ConnexionSQLServer()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connectionString As String

    Set conn = New ADODB.Connection

    ' Configuration pour SQL Server local
    connectionString = "Provider=SQLOLEDB;" & _
                      "Data Source=localhost;" & _
                      "Initial Catalog=Northwind;" & _
                      "Integrated Security=SSPI"

    On Error GoTo GestionErreur

    ' Tentative de connexion
    conn.ConnectionTimeout = 30  ' Timeout de 30 secondes
    conn.Open connectionString

    If conn.State = adStateOpen Then
        MsgBox "Connexion SQL Server réussie !"

        ' Requête avec paramètres
        Dim sql As String
        sql = "SELECT TOP 10 CustomerID, CompanyName, Country " & _
              "FROM Customers WHERE Country = 'France'"

        Set rs = conn.Execute(sql)

        ' Transfert vers Excel
        Dim ws As Worksheet
        Set ws = ActiveSheet

        ' En-têtes
        ws.Cells(1, 1).Value = "ID Client"
        ws.Cells(1, 2).Value = "Société"
        ws.Cells(1, 3).Value = "Pays"

        ' Données
        Dim ligne As Long
        ligne = 2

        Do While Not rs.EOF
            ws.Cells(ligne, 1).Value = rs.Fields("CustomerID").Value
            ws.Cells(ligne, 2).Value = rs.Fields("CompanyName").Value
            ws.Cells(ligne, 3).Value = rs.Fields("Country").Value
            ligne = ligne + 1
            rs.MoveNext
        Loop

        rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur SQL Server: " & Err.Description & vbCrLf & _
           "Vérifiez que le serveur est accessible et que la base existe."

    ' Nettoyage en cas d'erreur
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

### SQL Server Express (version gratuite)

```vba
' Pour SQL Server Express avec instance nommée
connectionString = "Provider=SQLOLEDB;" & _
                  "Data Source=MonOrdinateur\SQLEXPRESS;" & _
                  "Initial Catalog=MaBase;" & _
                  "Integrated Security=SSPI"
```

## MySQL

MySQL est une base de données très populaire, surtout dans le monde web. Pour vous y connecter, vous devez installer le connecteur ODBC MySQL.

### Installation du connecteur MySQL
1. Téléchargez "MySQL Connector/ODBC" depuis le site MySQL
2. Installez-le sur votre machine
3. Redémarrez Excel

### Chaîne de connexion MySQL

```vba
' Via ODBC Driver
connectionString = "Driver={MySQL ODBC 8.0 Unicode Driver};" & _
                  "Server=localhost;" & _
                  "Database=ma_base;" & _
                  "User=mon_utilisateur;" & _
                  "Password=mon_mot_de_passe;" & _
                  "Option=3"
```

### Exemple MySQL

```vba
Sub ConnexionMySQL()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connectionString As String

    Set conn = New ADODB.Connection

    ' Configuration MySQL
    connectionString = "Driver={MySQL ODBC 8.0 Unicode Driver};" & _
                      "Server=localhost;" & _
                      "Database=boutique;" & _
                      "User=root;" & _
                      "Password=;" & _
                      "Option=3"

    On Error GoTo GestionErreur

    conn.Open connectionString

    If conn.State = adStateOpen Then
        MsgBox "Connexion MySQL réussie !"

        ' Requête MySQL
        Set rs = conn.Execute("SELECT id, nom, prix FROM produits LIMIT 20")

        ' Affichage des résultats
        Dim ligne As Long
        ligne = 1

        Do While Not rs.EOF
            Cells(ligne, 1).Value = rs.Fields("id").Value
            Cells(ligne, 2).Value = rs.Fields("nom").Value
            Cells(ligne, 3).Value = rs.Fields("prix").Value
            ligne = ligne + 1
            rs.MoveNext
        Loop

        rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur MySQL: " & Err.Description
    ' Nettoyage...
End Sub
```

## Fichiers Excel comme source de données

Parfois, vous devez lire des données depuis d'autres fichiers Excel. C'est très utile pour consolider des rapports.

### Chaîne de connexion Excel

```vba
' Pour Excel 2007 et plus récent (.xlsx)
connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=C:\Rapports\Ventes2024.xlsx;" & _
                  "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""

' Pour Excel 2003 (.xls)
connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=C:\Rapports\Ventes2024.xls;" & _
                  "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
```

### Exemple lecture Excel externe

```vba
Sub LireAutreFichierExcel()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connectionString As String

    Set conn = New ADODB.Connection

    ' Chemin vers le fichier Excel à lire
    Dim cheminFichier As String
    cheminFichier = "C:\Rapports\VentesJanvier.xlsx"

    ' Vérifier que le fichier existe
    If Dir(cheminFichier) = "" Then
        MsgBox "Le fichier " & cheminFichier & " n'existe pas !"
        Exit Sub
    End If

    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & cheminFichier & ";" & _
                      "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""

    On Error GoTo GestionErreur

    conn.Open connectionString

    ' Lire la feuille "Ventes" (noter le $ à la fin)
    Set rs = conn.Execute("SELECT * FROM [Ventes$] WHERE Montant > 1000")

    ' Copier les données dans la feuille actuelle
    Dim ligne As Long
    ligne = 1

    ' En-têtes
    Cells(ligne, 1).Value = "Produit"
    Cells(ligne, 2).Value = "Montant"
    Cells(ligne, 3).Value = "Date"
    ligne = 2

    ' Données
    Do While Not rs.EOF
        Cells(ligne, 1).Value = rs.Fields("Produit").Value
        Cells(ligne, 2).Value = rs.Fields("Montant").Value
        Cells(ligne, 3).Value = rs.Fields("DateVente").Value
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Import terminé : " & (ligne - 2) & " lignes importées"
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de la lecture du fichier Excel: " & Err.Description
End Sub
```

## Fichiers CSV et texte

Les fichiers CSV sont très courants pour l'échange de données.

### Chaîne de connexion pour CSV

```vba
' Pour lire un fichier CSV
connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=C:\Donnees\;" & _
                  "Extended Properties=""Text;HDR=Yes;FMT=Delimited"""
```

### Exemple lecture CSV

```vba
Sub LireFichierCSV()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connectionString As String

    Set conn = New ADODB.Connection

    ' Le chemin pointe vers le DOSSIER contenant le CSV
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=C:\Donnees\;" & _
                      "Extended Properties=""Text;HDR=Yes;FMT=Delimited"""

    On Error GoTo GestionErreur

    conn.Open connectionString

    ' Le nom du fichier CSV (sans le chemin)
    Set rs = conn.Execute("SELECT * FROM clients.csv")

    Dim ligne As Long
    ligne = 1

    Do While Not rs.EOF
        Cells(ligne, 1).Value = rs.Fields(0).Value  ' Première colonne
        Cells(ligne, 2).Value = rs.Fields(1).Value  ' Deuxième colonne
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur CSV: " & Err.Description
End Sub
```

## Gestion des chemins et sécurité

### Chemins dynamiques

```vba
Function ObtenirCheminBase() As String
    ' Chemin relatif par rapport au fichier Excel
    Dim cheminExcel As String
    cheminExcel = ThisWorkbook.Path
    ObtenirCheminBase = cheminExcel & "\Donnees\MaBase.accdb"
End Function

Sub UtiliserCheminDynamique()
    Dim connectionString As String
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & ObtenirCheminBase()
    ' ... reste du code
End Sub
```

### Boîte de dialogue pour sélectionner une base

```vba
Sub ChoisirBaseDeDonnees()
    Dim cheminFichier As String

    ' Ouvrir la boîte de dialogue
    cheminFichier = Application.GetOpenFilename( _
        "Bases Access (*.accdb), *.accdb," & _
        "Anciennes bases Access (*.mdb), *.mdb", _
        , "Sélectionnez une base de données")

    If cheminFichier <> "False" Then
        ' L'utilisateur a sélectionné un fichier
        Dim connectionString As String
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                          "Data Source=" & cheminFichier

        ' Utiliser cette chaîne de connexion...
        MsgBox "Base sélectionnée : " & cheminFichier
    Else
        MsgBox "Aucune base sélectionnée"
    End If
End Sub
```

## Optimisation des connexions

### Réutilisation de connexions

```vba
Sub OperationsMultiples()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "votre_chaine_de_connexion"

    ' Opération 1 : Lecture
    Set rs = conn.Execute("SELECT COUNT(*) AS NbClients FROM Clients")
    Debug.Print "Nombre de clients : " & rs.Fields("NbClients").Value
    rs.Close

    ' Opération 2 : Mise à jour
    conn.Execute "UPDATE Clients SET DernierAcces = NOW() WHERE Actif = True"

    ' Opération 3 : Nouvelle lecture
    Set rs = conn.Execute("SELECT * FROM Clients WHERE DernierAcces > #" & Date & "#")
    ' Traitement des résultats...
    rs.Close

    ' Une seule fermeture
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

### Gestion du timeout

```vba
Sub ConnexionAvecTimeout()
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    ' Paramétrer les timeouts
    conn.ConnectionTimeout = 30  ' 30 secondes pour se connecter
    conn.CommandTimeout = 60     ' 60 secondes pour exécuter une commande

    conn.Open "votre_chaine_de_connexion"
    ' ... vos opérations
    conn.Close
End Sub
```

## Résolution de problèmes courants

### Erreur "Provider non trouvé"

```vba
' Vérifiez que les bons providers sont installés
Sub TesterProviders()
    On Error Resume Next

    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    ' Test Access
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=test.accdb"
    If Err.Number = 0 Then
        Debug.Print "Provider Access OK"
        conn.Close
    Else
        Debug.Print "Provider Access manquant : " & Err.Description
    End If
    Err.Clear

    ' Test SQL Server
    conn.Open "Provider=SQLOLEDB;Data Source=test"
    If Err.Number = 0 Then
        Debug.Print "Provider SQL Server OK"
        conn.Close
    Else
        Debug.Print "Provider SQL Server : " & Err.Description
    End If

    Set conn = Nothing
End Sub
```

### Gestion des caractères spéciaux

```vba
Function EchapperChaineSQL(texte As String) As String
    ' Remplacer les apostrophes simples par des doubles
    EchapperChaineSQL = Replace(texte, "'", "''")
End Function

Sub RequeteSecurisee()
    Dim nomClient As String
    nomClient = "O'Connor"  ' Nom avec apostrophe

    Dim sql As String
    sql = "SELECT * FROM Clients WHERE NomClient = '" & _
          EchapperChaineSQL(nomClient) & "'"

    Debug.Print sql  ' Affiche : ... WHERE NomClient = 'O''Connor'
End Sub
```

## Tableaux récapitulatifs

### Providers par type de base

| Type de base | Provider | Notes |
|--------------|----------|-------|
| Access 2007+ | Microsoft.ACE.OLEDB.12.0 | Recommandé |
| Access 2003- | Microsoft.Jet.OLEDB.4.0 | Ancien format |
| SQL Server | SQLOLEDB | Standard |
| Excel | Microsoft.ACE.OLEDB.12.0 | Avec Extended Properties |
| MySQL | Driver ODBC | Nécessite installation |
| Oracle | OraOLEDB | Nécessite client Oracle |

### Paramètres Extended Properties

| Format | Extended Properties | Usage |
|--------|-------------------|-------|
| Excel | "Excel 12.0 Xml;HDR=Yes" | Première ligne = en-têtes |
| CSV | "Text;HDR=Yes;FMT=Delimited" | Fichier délimité |
| Texte fixe | "Text;HDR=Yes;FMT=FixedLength" | Colonnes de largeur fixe |

## Résumé des bonnes pratiques

✅ **Toujours tester** la connexion avant de l'utiliser en production

✅ **Gérer les timeouts** pour éviter les blocages

✅ **Vérifier l'existence** des fichiers avant connexion

✅ **Utiliser des chemins absolus** ou des fonctions pour les chemins relatifs

✅ **Échapper les caractères spéciaux** dans les requêtes

✅ **Réutiliser les connexions** pour plusieurs opérations

✅ **Toujours fermer** les connexions et libérer les objets

✅ **Prévoir la gestion d'erreurs** pour chaque type de base

## Points d'attention

🚨 **Providers** : Vérifiez que les bons providers sont installés

🚨 **Permissions** : Assurez-vous d'avoir les droits d'accès aux bases

🚨 **Réseaux** : Les connexions distantes peuvent être lentes ou instables

🚨 **Versions** : Certains providers ne fonctionnent qu'en 32 ou 64 bits

🚨 **Sécurité** : Ne jamais stocker de mots de passe en dur dans le code

---

*Vous maîtrisez maintenant les connexions aux principales bases de données ! La prochaine étape sera d'apprendre à exploiter la puissance du SQL directement depuis VBA.*

⏭️
