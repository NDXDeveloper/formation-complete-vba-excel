🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.1 ADO (ActiveX Data Objects)

## Qu'est-ce qu'ADO ?

ADO (ActiveX Data Objects) est comme un **traducteur universel** entre VBA et les bases de données. Imaginez que vous voulez parler à quelqu'un qui ne parle pas votre langue : vous avez besoin d'un interprète. ADO joue exactement ce rôle entre Excel et les différentes sources de données.

ADO vous permet de :
- Vous connecter à une base de données
- Récupérer des informations (comme lire un livre)
- Modifier des données (comme écrire dans un cahier)
- Exécuter des commandes complexes

## Pourquoi utiliser ADO ?

### Les avantages d'ADO
- **Universalité** : Fonctionne avec presque toutes les bases de données
- **Puissance** : Permet d'exécuter des requêtes SQL complexes
- **Contrôle** : Vous maîtrisez chaque étape du processus
- **Performance** : Plus rapide que d'autres méthodes pour de gros volumes

### Quand utiliser ADO ?
- Quand vous devez récupérer des données depuis une base de données
- Quand vous voulez automatiser des rapports
- Quand vous avez besoin de traiter beaucoup de données
- Quand vous voulez créer des solutions professionnelles

## Les composants principaux d'ADO

ADO fonctionne avec trois objets principaux. Pensez-y comme aux éléments d'une conversation téléphonique :

### 1. Connection (La ligne téléphonique)
L'objet **Connection** établit le lien avec la base de données, comme composer un numéro de téléphone pour établir la communication.

### 2. Command (Ce que vous voulez dire)
L'objet **Command** contient l'instruction que vous voulez envoyer à la base de données, comme le message que vous voulez transmettre lors de votre appel.

### 3. Recordset (La réponse que vous recevez)
L'objet **Recordset** contient les données renvoyées par la base de données, comme la réponse de votre interlocuteur.

## Activation d'ADO dans VBA

Avant d'utiliser ADO, vous devez l'activer dans votre projet VBA :

### Étapes d'activation
1. Ouvrez l'éditeur VBA (Alt + F11)
2. Allez dans le menu **Outils** → **Références**
3. Cochez **Microsoft ActiveX Data Objects 6.1 Library** (ou la version la plus récente disponible)
4. Cliquez sur **OK**

> **💡 Astuce** : Si vous ne trouvez pas la version 6.1, prenez la version la plus élevée disponible (6.0, 2.8, etc.)

## Votre première connexion ADO

Commençons par un exemple simple pour comprendre le principe :

```vba
Sub PremierExempleADO()
    ' Déclaration des variables
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String

    ' Création de la connexion
    Set conn = New ADODB.Connection

    ' Ouverture de la connexion (exemple avec Access)
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Définition de la requête
    sql = "SELECT * FROM Clients"

    ' Exécution de la requête
    Set rs = conn.Execute(sql)

    ' Affichage du premier nom de client
    If Not rs.EOF Then
        MsgBox "Premier client: " & rs.Fields("NomClient").Value
    End If

    ' Nettoyage
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

### Explication ligne par ligne

```vba
Dim conn As ADODB.Connection
```
Déclare une variable pour stocker notre connexion à la base de données.

```vba
Set conn = New ADODB.Connection
```
Crée une nouvelle connexion (comme décrocher le téléphone).

```vba
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"
```
Ouvre la connexion vers le fichier Access spécifié (comme composer le numéro).

```vba
Set rs = conn.Execute(sql)
```
Exécute la requête et stocke le résultat dans un Recordset.

```vba
rs.Close  
conn.Close  
```
Ferme proprement les objets (comme raccrocher le téléphone).

## Les chaînes de connexion

Une **chaîne de connexion** indique à ADO comment se connecter à votre base de données. C'est comme donner une adresse précise.

### Pour Access
```vba
"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MonFichier.accdb"
```

### Pour Excel
```vba
"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MonFichier.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=Yes"""
```

### Pour SQL Server
```vba
"Provider=SQLOLEDB;Data Source=MonServeur;Initial Catalog=MaBase;Integrated Security=SSPI"
```

## Travailler avec les Recordsets

Un **Recordset** est comme un tableau de données que vous pouvez parcourir. Voici les méthodes principales :

### Navigation dans les données
```vba
Sub ParcoursRecordset()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    ' ... code de connexion ...

    Set rs = conn.Execute("SELECT * FROM Clients")

    ' Parcourir tous les enregistrements
    Do While Not rs.EOF
        Debug.Print rs.Fields("NomClient").Value
        rs.MoveNext  ' Passer au suivant
    Loop

    ' ... nettoyage ...
End Sub
```

### Vérifications importantes
```vba
' Vérifier si des données ont été trouvées
If Not rs.EOF Then
    ' Il y a des données
    Debug.Print rs.Fields("NomClient").Value
Else
    ' Aucune donnée trouvée
    MsgBox "Aucun client trouvé"
End If
```

## Récupérer des données dans Excel

Voici comment transférer des données de la base vers Excel :

```vba
Sub TransfererDonnees()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = ActiveSheet
    Set conn = New ADODB.Connection

    ' Connexion à la base
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Exécution de la requête
    Set rs = conn.Execute("SELECT NomClient, Ville FROM Clients")

    ' En-têtes
    ws.Cells(1, 1).Value = "Nom Client"
    ws.Cells(1, 2).Value = "Ville"

    ' Transfert des données
    ligne = 2
    Do While Not rs.EOF
        ws.Cells(ligne, 1).Value = rs.Fields("NomClient").Value
        ws.Cells(ligne, 2).Value = rs.Fields("Ville").Value
        ligne = ligne + 1
        rs.MoveNext
    Loop

    ' Nettoyage
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

## Gestion des erreurs avec ADO

Il est important de prévoir ce qui peut mal se passer :

```vba
Sub ExempleAvecGestionErreurs()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    On Error GoTo GestionErreur

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    Set rs = conn.Execute("SELECT * FROM Clients")

    ' Votre code ici...

    ' Nettoyage normal
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur: " & Err.Description

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

## Les types de curseurs

Les curseurs déterminent comment vous pouvez naviguer dans vos données :

### Curseur Forward-Only (par défaut)
```vba
Set rs = conn.Execute("SELECT * FROM Clients")
' Vous ne pouvez qu'avancer (MoveNext)
```

### Curseur plus flexible
```vba
Dim rs As ADODB.Recordset  
Set rs = New ADODB.Recordset  

rs.Open "SELECT * FROM Clients", conn, adOpenKeyset, adLockOptimistic

' Maintenant vous pouvez :
rs.MoveFirst  ' Aller au premier  
rs.MoveLast   ' Aller au dernier  
rs.MovePrevious  ' Reculer d'un  
rs.MoveNext   ' Avancer d'un  
```

## Modification de données

ADO permet aussi de modifier les données :

### Ajouter un enregistrement
```vba
Sub AjouterClient()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM Clients", conn, adOpenKeyset, adLockOptimistic

    ' Ajouter un nouveau client
    rs.AddNew
    rs.Fields("NomClient").Value = "Nouveau Client"
    rs.Fields("Ville").Value = "Paris"
    rs.Update

    ' Nettoyage
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

### Modifier un enregistrement existant
```vba
Sub ModifierClient()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Ouvrir un Recordset modifiable (pas conn.Execute qui est en lecture seule)
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM Clients WHERE NomClient = 'Dupont'", conn, _
            adOpenKeyset, adLockOptimistic

    If Not rs.EOF Then
        rs.Fields("Ville").Value = "Lyon"
        rs.Update
    End If

    ' Nettoyage
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

## Optimisation et bonnes pratiques

### 1. Toujours nettoyer les objets
```vba
' Toujours fermer et libérer
If Not rs Is Nothing Then
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End If

If Not conn Is Nothing Then
    If conn.State = adStateOpen Then conn.Close
    Set conn = Nothing
End If
```

### 2. Utiliser des requêtes spécifiques
```vba
' Évitez ceci :
sql = "SELECT * FROM Clients"

' Préférez ceci :
sql = "SELECT NomClient, Ville FROM Clients WHERE Actif = True"
```

### 3. Gérer les connexions efficacement
```vba
' Pour plusieurs opérations, réutilisez la même connexion
Sub OperationsMultiples()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set conn = New ADODB.Connection
    conn.Open "votre_chaine_de_connexion"

    ' Opération 1
    conn.Execute "UPDATE Clients SET Actif = True WHERE Ville = 'Paris'"

    ' Opération 2
    Set rs = conn.Execute("SELECT COUNT(*) FROM Clients")
    Debug.Print "Nombre de clients : " & rs.Fields(0).Value
    rs.Close
    Set rs = Nothing

    ' Une seule fermeture à la fin
    conn.Close
    Set conn = Nothing
End Sub
```

## Résumé des points clés

✅ **ADO** est un pont entre VBA et les bases de données

✅ **Trois objets principaux** : Connection, Command, et Recordset

✅ **Toujours activer** la référence ADO dans VBA

✅ **Les chaînes de connexion** définissent comment se connecter

✅ **Toujours nettoyer** les objets après utilisation

✅ **Gérer les erreurs** pour des solutions robustes

✅ **Optimiser les requêtes** pour de meilleures performances

## Points d'attention pour débutants

🚨 **N'oubliez pas** d'activer la référence ADO

🚨 **Attention** aux chemins de fichiers (utilisez des chemins complets)

🚨 **Vérifiez** toujours si le Recordset contient des données (EOF)

🚨 **Fermez** toujours vos connexions pour éviter les fuites mémoire

🚨 **Testez** vos chaînes de connexion avant de les utiliser dans le code

---

*Avec ADO, vous avez maintenant les clés pour faire communiquer Excel avec le monde des bases de données. La prochaine étape sera d'apprendre à se connecter à différents types de bases de données spécifiques !*

⏭️ [Connexion à des bases de données externes](/15-base-donnees-connexions/02-connexion-bases-donnees-externes.md)
