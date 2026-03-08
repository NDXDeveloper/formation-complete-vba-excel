🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.3 Requêtes SQL depuis VBA

## Introduction au SQL

SQL (Structured Query Language) est le langage universel pour communiquer avec les bases de données. Imaginez SQL comme un ensemble de phrases toutes faites que toutes les bases de données comprennent, peu importe leur "nationalité" (Access, SQL Server, MySQL, etc.).

Avec VBA, vous pouvez construire et exécuter des requêtes SQL dynamiquement, ce qui vous donne une puissance énorme pour manipuler les données. C'est comme avoir un traducteur automatique qui s'adapte à ce que vous voulez dire !

## Pourquoi utiliser SQL dans VBA ?

### Avantages du SQL
- **Précision** : Vous récupérez exactement les données dont vous avez besoin
- **Performance** : Plus rapide que de charger toutes les données puis les filtrer
- **Flexibilité** : Peut s'adapter aux besoins de l'utilisateur en temps réel
- **Puissance** : Permet des calculs complexes directement dans la base

### Comparaison : avec et sans SQL

**Sans SQL (inefficace) :**
```vba
' Charger TOUTE la table puis filtrer dans Excel
Set rs = conn.Execute("SELECT * FROM Commandes")  ' 100 000 lignes !  
Do While Not rs.EOF  
    If rs.Fields("DateCommande") >= DateAdd("m", -1, Date) Then
        ' Traiter seulement ces lignes...
    End If
    rs.MoveNext
Loop
```

**Avec SQL (efficace) :**
```vba
' Récupérer seulement ce qui nous intéresse
sql = "SELECT * FROM Commandes WHERE DateCommande >= DATEADD(month, -1, GETDATE())"  
Set rs = conn.Execute(sql)  ' Peut-être 500 lignes seulement !  
```

## Les types de requêtes SQL

### Requêtes de sélection (SELECT)
Pour **récupérer** des données. C'est comme demander : "Montre-moi..."

### Requêtes d'action
- **INSERT** : Pour **ajouter** des données
- **UPDATE** : Pour **modifier** des données existantes
- **DELETE** : Pour **supprimer** des données

## Requêtes SELECT de base

### Structure générale
```sql
SELECT colonnes  
FROM table  
WHERE conditions  
ORDER BY tri  
```

### Exemple simple
```vba
Sub RequeteSimple()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Requête simple
    sql = "SELECT NomClient, Ville FROM Clients"
    Set rs = conn.Execute(sql)

    ' Affichage des résultats
    Do While Not rs.EOF
        Debug.Print rs.Fields("NomClient").Value & " - " & rs.Fields("Ville").Value
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

### Sélectionner des colonnes spécifiques
```vba
' Toutes les colonnes
sql = "SELECT * FROM Clients"

' Colonnes spécifiques
sql = "SELECT NomClient, Email, Telephone FROM Clients"

' Colonnes avec alias (nom personnalisé)
sql = "SELECT NomClient AS Nom, Email AS AdresseMail FROM Clients"
```

## Filtrage avec WHERE

La clause WHERE est comme un filtre qui ne laisse passer que les données qui correspondent à vos critères.

### Conditions simples
```vba
Sub ExemplesWhere()
    Dim sql As String

    ' Égalité
    sql = "SELECT * FROM Clients WHERE Ville = 'Paris'"

    ' Différent
    sql = "SELECT * FROM Clients WHERE Ville <> 'Paris'"

    ' Comparaison numérique
    sql = "SELECT * FROM Commandes WHERE Montant > 1000"

    ' Plage de valeurs
    sql = "SELECT * FROM Commandes WHERE Montant BETWEEN 500 AND 2000"

    ' Liste de valeurs
    sql = "SELECT * FROM Clients WHERE Ville IN ('Paris', 'Lyon', 'Marseille')"

    ' Valeurs nulles
    sql = "SELECT * FROM Clients WHERE Email IS NOT NULL"
End Sub
```

### Recherche de texte avec LIKE
```vba
Sub ExemplesLike()
    Dim sql As String

    ' Commence par "Dup"
    sql = "SELECT * FROM Clients WHERE NomClient LIKE 'Dup*'"

    ' Finit par "son"
    sql = "SELECT * FROM Clients WHERE NomClient LIKE '*son'"

    ' Contient "Martin"
    sql = "SELECT * FROM Clients WHERE NomClient LIKE '*Martin*'"

    ' Email avec domaine gmail
    sql = "SELECT * FROM Clients WHERE Email LIKE '*@gmail.com'"
End Sub
```

### Combinaison de conditions
```vba
Sub ConditionsComplexes()
    Dim sql As String

    ' ET logique
    sql = "SELECT * FROM Clients WHERE Ville = 'Paris' AND Age > 25"

    ' OU logique
    sql = "SELECT * FROM Clients WHERE Ville = 'Paris' OR Ville = 'Lyon'"

    ' Combinaison avec parenthèses
    sql = "SELECT * FROM Clients WHERE (Ville = 'Paris' OR Ville = 'Lyon') AND Age > 30"

    ' Négation
    sql = "SELECT * FROM Clients WHERE NOT (Ville = 'Paris')"
End Sub
```

## Tri avec ORDER BY

```vba
Sub ExemplesTri()
    Dim sql As String

    ' Tri croissant (A vers Z)
    sql = "SELECT * FROM Clients ORDER BY NomClient"

    ' Tri décroissant (Z vers A)
    sql = "SELECT * FROM Clients ORDER BY NomClient DESC"

    ' Tri sur plusieurs colonnes
    sql = "SELECT * FROM Clients ORDER BY Ville, NomClient"

    ' Tri mixte
    sql = "SELECT * FROM Clients ORDER BY Ville ASC, Age DESC"
End Sub
```

## Limitation du nombre de résultats

```vba
Sub LimiterResultats()
    Dim sql As String

    ' Access : TOP
    sql = "SELECT TOP 10 * FROM Clients ORDER BY DateInscription DESC"

    ' SQL Server : TOP
    sql = "SELECT TOP 10 * FROM Clients ORDER BY DateInscription DESC"

    ' MySQL : LIMIT
    sql = "SELECT * FROM Clients ORDER BY DateInscription DESC LIMIT 10"
End Sub
```

## Fonctions SQL utiles

### Fonctions de texte
```vba
Sub FonctionsTexte()
    Dim sql As String

    ' Longueur du texte
    sql = "SELECT NomClient, LEN(NomClient) AS Longueur FROM Clients"

    ' Extraction de caractères
    sql = "SELECT LEFT(NomClient, 3) AS Initiales FROM Clients"
    sql = "SELECT RIGHT(Email, 10) AS FinEmail FROM Clients"
    sql = "SELECT MID(Telephone, 3, 2) AS Indicatif FROM Clients"

    ' Conversion majuscules/minuscules
    sql = "SELECT UPPER(NomClient) AS NomMajuscule FROM Clients"
    sql = "SELECT LOWER(Email) AS EmailMinuscule FROM Clients"

    ' Suppression des espaces
    sql = "SELECT TRIM(NomClient) AS NomNettoye FROM Clients"
End Sub
```

### Fonctions de date
```vba
Sub FonctionsDate()
    Dim sql As String

    ' Date actuelle
    sql = "SELECT *, NOW() AS DateActuelle FROM Commandes"

    ' Extraction de parties de date
    sql = "SELECT *, YEAR(DateCommande) AS Annee FROM Commandes"
    sql = "SELECT *, MONTH(DateCommande) AS Mois FROM Commandes"
    sql = "SELECT *, DAY(DateCommande) AS Jour FROM Commandes"

    ' Calculs de dates
    sql = "SELECT *, DATEDIFF('d', DateCommande, NOW()) AS JoursEcoules FROM Commandes"
    sql = "SELECT * FROM Commandes WHERE DateCommande > DATEADD('m', -1, NOW())"
End Sub
```

### Fonctions de calcul
```vba
Sub FonctionsCalcul()
    Dim sql As String

    ' Fonctions d'agrégation
    sql = "SELECT COUNT(*) AS NombreClients FROM Clients"
    sql = "SELECT SUM(Montant) AS TotalVentes FROM Commandes"
    sql = "SELECT AVG(Montant) AS MoyenneVentes FROM Commandes"
    sql = "SELECT MIN(Montant) AS VenteMin, MAX(Montant) AS VenteMax FROM Commandes"

    ' Calculs personnalisés
    sql = "SELECT *, Montant * 1.2 AS MontantTTC FROM Commandes"
    sql = "SELECT *, ROUND(Montant / Quantite, 2) AS PrixUnitaire FROM Commandes"
End Sub
```

## Requêtes avec paramètres dynamiques

### Construction de requêtes dynamiques
```vba
Sub RequeteDynamique()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim villeRecherchee As String
    Dim ageMinimum As Integer

    ' Paramètres saisis par l'utilisateur
    villeRecherchee = InputBox("Quelle ville recherchez-vous ?", "Recherche", "Paris")
    ageMinimum = Val(InputBox("Âge minimum ?", "Recherche", "18"))

    ' Construction de la requête
    sql = "SELECT NomClient, Age, Ville FROM Clients " & _
          "WHERE Ville = '" & villeRecherchee & "' " & _
          "AND Age >= " & ageMinimum & " " & _
          "ORDER BY NomClient"

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    Set rs = conn.Execute(sql)

    ' Affichage des résultats dans Excel
    Dim ligne As Long
    ligne = 1

    ' En-têtes
    Cells(ligne, 1).Value = "Nom"
    Cells(ligne, 2).Value = "Âge"
    Cells(ligne, 3).Value = "Ville"
    ligne = 2

    ' Données
    Do While Not rs.EOF
        Cells(ligne, 1).Value = rs.Fields("NomClient").Value
        Cells(ligne, 2).Value = rs.Fields("Age").Value
        Cells(ligne, 3).Value = rs.Fields("Ville").Value
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Recherche terminée : " & (ligne - 2) & " clients trouvés"
End Sub
```

### Fonction pour sécuriser les chaînes
```vba
Function SecuriserChaine(texte As String) As String
    ' Remplace les apostrophes simples par des doubles pour éviter les erreurs SQL
    SecuriserChaine = Replace(texte, "'", "''")
End Function

Sub RequeteSecurisee()
    Dim nomRecherche As String
    Dim sql As String

    nomRecherche = InputBox("Nom du client ?")

    ' Version sécurisée
    sql = "SELECT * FROM Clients WHERE NomClient = '" & _
          SecuriserChaine(nomRecherche) & "'"

    ' Maintenant, même si l'utilisateur saisit "O'Connor", ça fonctionne !
End Sub
```

## Jointures entre tables

Les jointures permettent de combiner des données provenant de plusieurs tables.

### Jointure simple (INNER JOIN)
```vba
Sub JointureSimple()
    Dim sql As String

    ' Récupérer les commandes avec les informations client
    sql = "SELECT c.NomClient, c.Ville, cmd.DateCommande, cmd.Montant " & _
          "FROM Clients c " & _
          "INNER JOIN Commandes cmd ON c.ClientID = cmd.ClientID " & _
          "ORDER BY cmd.DateCommande DESC"

    ' Cette requête combine les tables Clients et Commandes
    ' pour avoir le nom du client avec chaque commande
End Sub
```

### Jointure externe (LEFT JOIN)
```vba
Sub JointureExterne()
    Dim sql As String

    ' Tous les clients, même ceux sans commande
    sql = "SELECT c.NomClient, c.Ville, cmd.DateCommande, cmd.Montant " & _
          "FROM Clients c " & _
          "LEFT JOIN Commandes cmd ON c.ClientID = cmd.ClientID " & _
          "ORDER BY c.NomClient"

    ' LEFT JOIN inclut tous les clients, même sans commande
    ' Les champs de commande seront NULL pour les clients sans commande
End Sub
```

### Exemple pratique de jointure
```vba
Sub RapportVentesParVille()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Requête avec jointure et regroupement
    sql = "SELECT c.Ville, COUNT(cmd.CommandeID) AS NbCommandes, " & _
          "SUM(cmd.Montant) AS TotalVentes " & _
          "FROM Clients c " & _
          "LEFT JOIN Commandes cmd ON c.ClientID = cmd.ClientID " & _
          "GROUP BY c.Ville " & _
          "ORDER BY TotalVentes DESC"

    Set rs = conn.Execute(sql)

    ' Affichage dans Excel
    Dim ligne As Long
    ligne = 1

    ' En-têtes
    Cells(ligne, 1).Value = "Ville"
    Cells(ligne, 2).Value = "Nb Commandes"
    Cells(ligne, 3).Value = "Total Ventes"
    ligne = 2

    ' Données
    Do While Not rs.EOF
        Cells(ligne, 1).Value = rs.Fields("Ville").Value
        Cells(ligne, 2).Value = rs.Fields("NbCommandes").Value
        Cells(ligne, 3).Value = rs.Fields("TotalVentes").Value
        ligne = ligne + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

## Regroupement avec GROUP BY

GROUP BY permet de faire des calculs sur des groupes de données.

```vba
Sub ExemplesGroupBy()
    Dim sql As String

    ' Nombre de clients par ville
    sql = "SELECT Ville, COUNT(*) AS NbClients " & _
          "FROM Clients " & _
          "GROUP BY Ville " & _
          "ORDER BY NbClients DESC"

    ' Ventes mensuelles
    sql = "SELECT YEAR(DateCommande) AS Annee, MONTH(DateCommande) AS Mois, " & _
          "SUM(Montant) AS TotalMensuel " & _
          "FROM Commandes " & _
          "GROUP BY YEAR(DateCommande), MONTH(DateCommande) " & _
          "ORDER BY Annee, Mois"

    ' Avec condition sur les groupes (HAVING)
    sql = "SELECT Ville, COUNT(*) AS NbClients " & _
          "FROM Clients " & _
          "GROUP BY Ville " & _
          "HAVING COUNT(*) > 5 " & _
          "ORDER BY NbClients DESC"
End Sub
```

## Requêtes d'insertion (INSERT)

### Insertion simple
```vba
Sub AjouterClient()
    Dim conn As ADODB.Connection
    Dim sql As String

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Insertion d'un nouveau client
    sql = "INSERT INTO Clients (NomClient, Email, Ville, Age) " & _
          "VALUES ('Nouveau Client', 'email@test.com', 'Paris', 35)"

    conn.Execute sql

    MsgBox "Client ajouté avec succès !"

    conn.Close
    Set conn = Nothing
End Sub
```

### Insertion avec variables
```vba
Sub AjouterClientVariable()
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim nom As String, email As String, ville As String, age As Integer

    ' Récupération des données
    nom = InputBox("Nom du client ?")
    email = InputBox("Email ?")
    ville = InputBox("Ville ?")
    age = Val(InputBox("Âge ?"))

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Construction de la requête avec variables
    sql = "INSERT INTO Clients (NomClient, Email, Ville, Age) " & _
          "VALUES ('" & SecuriserChaine(nom) & "', '" & _
          SecuriserChaine(email) & "', '" & _
          SecuriserChaine(ville) & "', " & age & ")"

    On Error GoTo GestionErreur

    conn.Execute sql
    MsgBox "Client " & nom & " ajouté avec succès !"

    conn.Close
    Set conn = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'ajout : " & Err.Description
    If Not conn Is Nothing Then conn.Close
End Sub
```

## Requêtes de mise à jour (UPDATE)

### Mise à jour simple
```vba
Sub MettreAJourClient()
    Dim conn As ADODB.Connection
    Dim sql As String

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Mise à jour d'un client spécifique
    sql = "UPDATE Clients " & _
          "SET Email = 'nouveau@email.com', Ville = 'Lyon' " & _
          "WHERE NomClient = 'Dupont'"

    conn.Execute sql

    MsgBox "Client mis à jour !"

    conn.Close
    Set conn = Nothing
End Sub
```

### Mise à jour en lot
```vba
Sub MiseAJourEnLot()
    Dim conn As ADODB.Connection
    Dim sql As String

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    ' Augmenter l'âge de tous les clients de Paris
    sql = "UPDATE Clients " & _
          "SET Age = Age + 1 " & _
          "WHERE Ville = 'Paris'"

    conn.Execute sql

    ' Marquer les gros clients
    sql = "UPDATE Clients " & _
          "SET Statut = 'VIP' " & _
          "WHERE ClientID IN (" & _
          "    SELECT ClientID FROM Commandes " & _
          "    GROUP BY ClientID " & _
          "    HAVING SUM(Montant) > 10000" & _
          ")"

    conn.Execute sql

    conn.Close
    Set conn = Nothing
End Sub
```

## Requêtes de suppression (DELETE)

### Suppression avec condition
```vba
Sub SupprimerAnciennesCommandes()
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim reponse As VbMsgBoxResult

    ' Confirmation avant suppression
    reponse = MsgBox("Voulez-vous vraiment supprimer les commandes de plus de 2 ans ?", _
                     vbYesNo + vbQuestion, "Confirmation")

    If reponse = vbYes Then
        Set conn = New ADODB.Connection
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

        ' Suppression des anciennes commandes
        sql = "DELETE FROM Commandes " & _
              "WHERE DateCommande < DATEADD('yyyy', -2, NOW())"

        conn.Execute sql

        MsgBox "Anciennes commandes supprimées !"

        conn.Close
        Set conn = Nothing
    End If
End Sub
```

## Gestion avancée des requêtes

### Utilisation de l'objet Command
```vba
Sub UtiliserCommand()
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandText = "SELECT * FROM Clients WHERE Ville = ? AND Age > ?"
    cmd.CommandType = adCmdText

    ' Ajout de paramètres (plus sûr que la concaténation)
    cmd.Parameters.Append cmd.CreateParameter("Ville", adVarChar, adParamInput, 50, "Paris")
    cmd.Parameters.Append cmd.CreateParameter("Age", adInteger, adParamInput, , 25)

    Set rs = cmd.Execute

    ' Traitement des résultats...

    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub
```

### Requêtes avec transactions
```vba
Sub TransactionExemple()
    Dim conn As ADODB.Connection

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    On Error GoTo AnnulerTransaction

    ' Début de la transaction
    conn.BeginTrans

    ' Plusieurs opérations qui doivent toutes réussir
    conn.Execute "INSERT INTO Clients (NomClient, Ville) VALUES ('Test1', 'Paris')"
    conn.Execute "INSERT INTO Commandes (ClientID, Montant) VALUES (1, 1000)"
    conn.Execute "UPDATE Stock SET Quantite = Quantite - 1 WHERE ProduitID = 5"

    ' Si on arrive ici, tout s'est bien passé
    conn.CommitTrans
    MsgBox "Transaction réussie !"

    conn.Close
    Set conn = Nothing
    Exit Sub

AnnulerTransaction:
    ' En cas d'erreur, annuler toutes les modifications
    conn.RollbackTrans
    MsgBox "Erreur : transaction annulée - " & Err.Description

    conn.Close
    Set conn = Nothing
End Sub
```

## Construction d'un générateur de requêtes

```vba
Sub GenerateurRequetes()
    Dim sql As String
    Dim whereClause As String
    Dim orderClause As String

    ' Interface utilisateur simple
    Dim table As String
    table = InputBox("Table à interroger ?", "Générateur", "Clients")

    Dim colonnes As String
    colonnes = InputBox("Colonnes (séparées par des virgules) ?", "Générateur", "*")

    Dim filtre As String
    filtre = InputBox("Filtre (optionnel) ?", "Générateur", "")

    Dim tri As String
    tri = InputBox("Tri (optionnel) ?", "Générateur", "")

    ' Construction de la requête
    sql = "SELECT " & colonnes & " FROM " & table

    If filtre <> "" Then
        sql = sql & " WHERE " & filtre
    End If

    If tri <> "" Then
        sql = sql & " ORDER BY " & tri
    End If

    ' Affichage de la requête générée
    MsgBox "Requête générée :" & vbCrLf & sql

    ' Exécution (optionnelle)
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Exécuter cette requête ?", vbYesNo + vbQuestion)

    If reponse = vbYes Then
        ExecuterRequete sql
    End If
End Sub

Sub ExecuterRequete(sql As String)
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    On Error GoTo GestionErreur

    Set rs = conn.Execute(sql)

    ' Affichage simple des résultats
    Dim ligne As Long
    ligne = 1

    Do While Not rs.EOF And ligne <= 100  ' Limiter à 100 lignes
        Dim col As Integer
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

    MsgBox "Requête exécutée : " & (ligne - 1) & " lignes affichées"
    Exit Sub

GestionErreur:
    MsgBox "Erreur SQL : " & Err.Description
    If Not conn Is Nothing Then conn.Close
End Sub
```

## Optimisation des requêtes

### Bonnes pratiques
```vba
Sub BonnesPratiquesPerformance()
    Dim sql As String

    ' ✅ Bon : Sélectionner seulement les colonnes nécessaires
    sql = "SELECT NomClient, Email FROM Clients"

    ' ❌ Éviter : Sélectionner tout si pas nécessaire
    ' sql = "SELECT * FROM Clients"

    ' ✅ Bon : Utiliser des index dans les WHERE
    sql = "SELECT * FROM Commandes WHERE ClientID = 123"  ' Si ClientID est indexé

    ' ✅ Bon : Limiter le nombre de résultats
    sql = "SELECT TOP 100 * FROM Commandes ORDER BY DateCommande DESC"

    ' ✅ Bon : Utiliser EXISTS au lieu de IN pour de gros volumes
    sql = "SELECT * FROM Clients c " & _
          "WHERE EXISTS (SELECT 1 FROM Commandes cmd WHERE cmd.ClientID = c.ClientID)"
End Sub
```

### Mesure des performances
```vba
Sub MesurerPerformance()
    Dim debut As Date
    Dim fin As Date
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MaBase.accdb"

    debut = Now()

    ' Utiliser un curseur statique pour obtenir RecordCount
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM GrosseTable WHERE Condition = 'Valeur'", conn, _
            adOpenStatic, adLockReadOnly

    fin = Now()

    Debug.Print "Requête exécutée en " & DateDiff("s", debut, fin) & " secondes"
    Debug.Print "Nombre de lignes : " & rs.RecordCount

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

## Résumé des points clés

✅ **SQL** est le langage universel des bases de données

✅ **SELECT** pour récupérer, **INSERT/UPDATE/DELETE** pour modifier

✅ **WHERE** filtre les données, **ORDER BY** les trie

✅ **Jointures** combinent plusieurs tables

✅ **GROUP BY** permet les calculs par groupes

✅ **Toujours sécuriser** les chaînes de caractères

✅ **Transactions** garantissent la cohérence des données

✅ **Optimiser** en sélectionnant seulement ce qui est nécessaire

## Points d'attention pour débutants

🚨 **Attention aux apostrophes** dans les noms (utilisez la fonction de sécurisation)

🚨 **Testez toujours** vos requêtes avec des petits volumes d'abord

🚨 **Sauvegardez** avant les requêtes UPDATE/DELETE

🚨 **Différences SQL** : chaque base a ses spécificités (TOP vs LIMIT)

🚨 **Performance** : évitez SELECT * sur de gros volumes

🚨 **Gestion d'erreurs** : prévoyez toujours les cas d'échec

---

*Avec ces connaissances SQL, vous pouvez maintenant extraire et manipuler les données avec une précision chirurgicale ! La prochaine étape sera d'apprendre à automatiser les imports et exports de données.*

⏭️
