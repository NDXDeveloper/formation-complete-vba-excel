🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.3 Automation avec Outlook

## Introduction à Outlook Automation

L'automation avec Outlook permet d'envoyer des emails, gérer les contacts, créer des rendez-vous et accéder aux données du calendrier directement depuis Excel avec VBA. C'est particulièrement utile pour automatiser l'envoi de rapports, créer des notifications automatiques ou synchroniser des données.

## Première étape : Créer une connexion avec Outlook

### Méthode simple pour débuter

```vba
Sub PremierTestOutlook()
    ' Créer une connexion avec Outlook
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")

    ' Créer un nouveau message
    Dim mail As Object
    Set mail = outlookApp.CreateItem(0)  ' 0 = Email

    ' Configurer le message
    With mail
        .To = "destinataire@exemple.com"
        .Subject = "Mon premier email automatisé"
        .Body = "Bonjour, ceci est un email envoyé automatiquement depuis Excel !"
        .Display  ' Afficher le message (sans l'envoyer automatiquement)
    End With

    ' Important : Libérer la mémoire
    Set mail = Nothing
    Set outlookApp = Nothing
End Sub
```

**Explication ligne par ligne :**
- `CreateObject("Outlook.Application")` : Se connecte à Outlook
- `outlookApp.CreateItem(0)` : Crée un nouvel email (0 = type email)
- `.To` : Définit le destinataire
- `.Subject` : Définit l'objet du message
- `.Body` : Définit le contenu du message
- `.Display` : Affiche le message à l'écran (permet de vérifier avant envoi)

## Comprendre les types d'éléments Outlook

Outlook gère différents types d'éléments :

```vba
Sub TypesElementsOutlook()
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")

    ' Différents types d'éléments
    ' 0 = Email (MailItem)
    ' 1 = Rendez-vous (AppointmentItem)
    ' 2 = Contact (ContactItem)
    ' 3 = Tâche (TaskItem)
    ' 4 = Note (NoteItem)

    ' Exemple de chaque type
    Dim email As Object
    Set email = outlookApp.CreateItem(0)

    Dim rendezVous As Object
    Set rendezVous = outlookApp.CreateItem(1)

    Dim contact As Object
    Set contact = outlookApp.CreateItem(2)

    ' Ne pas oublier de libérer la mémoire
    Set contact = Nothing
    Set rendezVous = Nothing
    Set email = Nothing
    Set outlookApp = Nothing
End Sub
```

## Envoyer des emails simples

### Email basique

```vba
Sub EnvoyerEmailSimple()
    Dim outlookApp As Object
    Dim mail As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)

    With mail
        .To = "destinataire@exemple.com"
        .Subject = "Rapport automatique"
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
                "Veuillez trouver ci-joint le rapport automatique." & vbCrLf & vbCrLf & _
                "Cordialement," & vbCrLf & _
                "Système automatisé Excel"

        ' Pour envoyer automatiquement (attention !)
        ' .Send

        ' Pour afficher avant envoi (plus sûr)
        .Display
    End With

    Set mail = Nothing
    Set outlookApp = Nothing
End Sub
```

### Email avec plusieurs destinataires

```vba
Sub EmailPlusieursDestinataires()
    Dim outlookApp As Object
    Dim mail As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)

    With mail
        .To = "destinataire1@exemple.com; destinataire2@exemple.com"
        .CC = "copie@exemple.com"
        .BCC = "copie.cachee@exemple.com"
        .Subject = "Email à plusieurs destinataires"
        .Body = "Ce message est envoyé à plusieurs personnes."
        .Display
    End With

    Set mail = Nothing
    Set outlookApp = Nothing
End Sub
```

## Emails avec contenu dynamique depuis Excel

### Incorporer des données Excel dans l'email

```vba
Sub EmailAvecDonneesExcel()
    Dim outlookApp As Object
    Dim mail As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)

    ' Récupérer des données depuis Excel
    Dim nomClient As String
    Dim montantCommande As Double
    Dim dateCommande As Date

    nomClient = Range("A2").Value
    montantCommande = Range("B2").Value
    dateCommande = Range("C2").Value

    ' Créer le contenu personnalisé
    Dim contenu As String
    contenu = "Bonjour " & nomClient & "," & vbCrLf & vbCrLf
    contenu = contenu & "Nous vous confirmons votre commande du " & Format(dateCommande, "dd/mm/yyyy") & vbCrLf
    contenu = contenu & "Montant total : " & Format(montantCommande, "#,##0.00") & " €" & vbCrLf & vbCrLf
    contenu = contenu & "Merci pour votre confiance." & vbCrLf & vbCrLf
    contenu = contenu & "Cordialement," & vbCrLf
    contenu = contenu & "L'équipe commerciale"

    With mail
        .To = Range("D2").Value  ' Email du client en colonne D
        .Subject = "Confirmation de commande - " & nomClient
        .Body = contenu
        .Display
    End With

    Set mail = Nothing
    Set outlookApp = Nothing
End Sub
```

### Envoyer des emails en lot

```vba
Sub EnvoyerEmailsEnLot()
    Dim outlookApp As Object
    Dim mail As Object
    Dim i As Integer
    Dim derniereLigne As Integer

    Set outlookApp = CreateObject("Outlook.Application")

    ' Trouver la dernière ligne avec des données
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    ' Boucle pour chaque ligne de données (à partir de la ligne 2)
    For i = 2 To derniereLigne
        Set mail = outlookApp.CreateItem(0)

        With mail
            .To = Cells(i, 4).Value  ' Email en colonne D
            .Subject = "Information personnalisée pour " & Cells(i, 1).Value
            .Body = "Bonjour " & Cells(i, 1).Value & "," & vbCrLf & vbCrLf & _
                   "Voici vos informations :" & vbCrLf & _
                   "• Statut : " & Cells(i, 2).Value & vbCrLf & _
                   "• Montant : " & Cells(i, 3).Value & " €" & vbCrLf & vbCrLf & _
                   "Cordialement"

            ' Attention : .Send envoie automatiquement !
            ' Pour tester, utilisez .Display d'abord
            .Display
            ' .Send
        End With

        Set mail = Nothing
    Next i

    Set outlookApp = Nothing
    MsgBox "Emails préparés pour " & (derniereLigne - 1) & " destinataires"
End Sub
```

## Ajouter des pièces jointes

### Pièce jointe simple

```vba
Sub EmailAvecPieceJointe()
    Dim outlookApp As Object
    Dim mail As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)

    With mail
        .To = "destinataire@exemple.com"
        .Subject = "Rapport avec pièce jointe"
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
               "Veuillez trouver le rapport en pièce jointe." & vbCrLf & vbCrLf & _
               "Cordialement"

        ' Ajouter une pièce jointe (changez le chemin)
        .Attachments.Add "C:\MonDossier\Rapport.xlsx"

        .Display
    End With

    Set mail = Nothing
    Set outlookApp = Nothing
End Sub
```

### Joindre le fichier Excel actuel

```vba
Sub JoindreFichierActuel()
    Dim outlookApp As Object
    Dim mail As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)

    ' Sauvegarder le fichier actuel d'abord
    ThisWorkbook.Save

    With mail
        .To = "destinataire@exemple.com"
        .Subject = "Fichier Excel en cours - " & ThisWorkbook.Name
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
               "Voici le fichier Excel avec les dernières données." & vbCrLf & vbCrLf & _
               "Cordialement"

        ' Joindre le fichier Excel actuel
        .Attachments.Add ThisWorkbook.FullName

        .Display
    End With

    Set mail = Nothing
    Set outlookApp = Nothing
End Sub
```

## Créer des rendez-vous dans le calendrier

### Rendez-vous simple

```vba
Sub CreerRendezVousSimple()
    Dim outlookApp As Object
    Dim rendezVous As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set rendezVous = outlookApp.CreateItem(1)  ' 1 = Rendez-vous

    With rendezVous
        .Subject = "Réunion projet Excel"
        .Location = "Salle de conférence A"
        .Start = DateAdd("d", 1, Now) + TimeValue("14:00:00")  ' Demain à 14h
        .End = DateAdd("d", 1, Now) + TimeValue("15:00:00")    ' Demain à 15h
        .Body = "Réunion pour discuter de l'avancement du projet d'automatisation Excel."

        ' Définir un rappel (en minutes avant le rendez-vous)
        .ReminderSet = True
        .ReminderMinutesBeforeStart = 15

        .Save  ' Sauvegarder dans le calendrier
    End With

    Set rendezVous = Nothing
    Set outlookApp = Nothing

    MsgBox "Rendez-vous créé dans votre calendrier !"
End Sub
```

### Réunion avec participants

```vba
Sub CreerReunionAvecParticipants()
    Dim outlookApp As Object
    Dim reunion As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set reunion = outlookApp.CreateItem(1)

    With reunion
        .Subject = "Réunion équipe - Résultats du mois"
        .Location = "Salle de réunion principale"
        .Start = DateAdd("d", 2, Now) + TimeValue("10:00:00")  ' Après-demain à 10h
        .End = DateAdd("d", 2, Now) + TimeValue("11:30:00")    ' Jusqu'à 11h30
        .Body = "Ordre du jour :" & vbCrLf & _
               "1. Présentation des résultats" & vbCrLf & _
               "2. Objectifs du mois prochain" & vbCrLf & _
               "3. Questions diverses"

        ' Ajouter des participants
        .Recipients.Add "collegue1@exemple.com"
        .Recipients.Add "collegue2@exemple.com"
        .Recipients.Add "manager@exemple.com"

        ' Résoudre les noms (vérifier que les adresses sont valides)
        .Recipients.ResolveAll

        .Send  ' Envoyer les invitations
    End With

    Set reunion = Nothing
    Set outlookApp = Nothing

    MsgBox "Invitations de réunion envoyées !"
End Sub
```

## Gérer les contacts

### Créer un nouveau contact

```vba
Sub CreerNouveauContact()
    Dim outlookApp As Object
    Dim contact As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set contact = outlookApp.CreateItem(2)  ' 2 = Contact

    With contact
        .FirstName = "Jean"
        .LastName = "Dupont"
        .Email1Address = "jean.dupont@exemple.com"
        .BusinessTelephoneNumber = "01 23 45 67 89"
        .CompanyName = "Entreprise ABC"
        .JobTitle = "Directeur commercial"
        .BusinessAddress = "123 Rue de la Paix" & vbCrLf & "75001 Paris"

        .Save  ' Sauvegarder dans les contacts
    End With

    Set contact = Nothing
    Set outlookApp = Nothing

    MsgBox "Contact créé avec succès !"
End Sub
```

### Créer des contacts depuis Excel

```vba
Sub CreerContactsDepuisExcel()
    Dim outlookApp As Object
    Dim contact As Object
    Dim i As Integer
    Dim derniereLigne As Integer

    Set outlookApp = CreateObject("Outlook.Application")

    ' Supposons les données en colonnes : A=Prénom, B=Nom, C=Email, D=Téléphone, E=Entreprise
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To derniereLigne  ' Commencer à la ligne 2 (ligne 1 = en-têtes)
        Set contact = outlookApp.CreateItem(2)

        With contact
            .FirstName = Cells(i, 1).Value  ' Colonne A
            .LastName = Cells(i, 2).Value   ' Colonne B
            .Email1Address = Cells(i, 3).Value  ' Colonne C
            .BusinessTelephoneNumber = Cells(i, 4).Value  ' Colonne D
            .CompanyName = Cells(i, 5).Value  ' Colonne E

            .Save
        End With

        Set contact = Nothing
    Next i

    Set outlookApp = Nothing
    MsgBox (derniereLigne - 1) & " contacts créés avec succès !"
End Sub
```

## Créer des tâches

### Tâche simple

```vba
Sub CreerTacheSimple()
    Dim outlookApp As Object
    Dim tache As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set tache = outlookApp.CreateItem(3)  ' 3 = Tâche

    With tache
        .Subject = "Finaliser le rapport Excel"
        .Body = "Terminer l'automatisation des rapports mensuels avec VBA."
        .DueDate = DateAdd("d", 7, Date)  ' Échéance dans 7 jours
        .StartDate = Date  ' Commence aujourd'hui
        .Importance = 2  ' 0=Faible, 1=Normale, 2=Élevée
        .ReminderSet = True
        .ReminderTime = DateAdd("d", 6, Date) + TimeValue("09:00:00")  ' Rappel la veille à 9h

        .Save
    End With

    Set tache = Nothing
    Set outlookApp = Nothing

    MsgBox "Tâche créée dans Outlook !"
End Sub
```

## Lire les emails reçus

### Accéder à la boîte de réception

```vba
Sub LireEmailsRecus()
    Dim outlookApp As Object
    Dim boiteReception As Object
    Dim emails As Object
    Dim email As Object
    Dim i As Integer

    Set outlookApp = CreateObject("Outlook.Application")
    Set boiteReception = outlookApp.GetNamespace("MAPI").GetDefaultFolder(6)  ' 6 = Boîte de réception
    Set emails = boiteReception.Items

    ' Lire les 5 derniers emails reçus
    For i = 1 To 5
        If i <= emails.Count Then
            Set email = emails(i)

            Debug.Print "Email " & i & ":"
            Debug.Print "De: " & email.SenderName
            Debug.Print "Objet: " & email.Subject
            Debug.Print "Reçu le: " & email.ReceivedTime
            Debug.Print "---"
        End If
    Next i

    Set email = Nothing
    Set emails = Nothing
    Set boiteReception = Nothing
    Set outlookApp = Nothing

    MsgBox "Informations des emails affichées dans la fenêtre Exécution immédiate (Ctrl+G)"
End Sub
```

## Rechercher des emails spécifiques

```vba
Sub RechercherEmails()
    Dim outlookApp As Object
    Dim boiteReception As Object
    Dim emailsTrouves As Object
    Dim email As Object
    Dim i As Integer

    Set outlookApp = CreateObject("Outlook.Application")
    Set boiteReception = outlookApp.GetNamespace("MAPI").GetDefaultFolder(6)

    ' Rechercher les emails contenant "rapport" dans l'objet
    Set emailsTrouves = boiteReception.Items.Restrict("[Subject] LIKE '%rapport%'")

    Debug.Print "Emails trouvés avec 'rapport' dans l'objet : " & emailsTrouves.Count

    For i = 1 To emailsTrouves.Count
        Set email = emailsTrouves(i)
        Debug.Print "- " & email.Subject & " (de " & email.SenderName & ")"
    Next i

    Set email = Nothing
    Set emailsTrouves = Nothing
    Set boiteReception = Nothing
    Set outlookApp = Nothing
End Sub
```

## Gestion des erreurs avec Outlook

```vba
Sub EnvoyerEmailAvecGestionErreurs()
    Dim outlookApp As Object
    Dim mail As Object

    On Error GoTo GestionErreur

    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)

    With mail
        .To = "test@exemple.com"
        .Subject = "Test avec gestion d'erreurs"
        .Body = "Ceci est un test."
        .Display
    End With

    Set mail = Nothing
    Set outlookApp = Nothing

    MsgBox "Email préparé avec succès !"
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de la création de l'email : " & Err.Description

    ' Nettoyage en cas d'erreur
    If Not mail Is Nothing Then Set mail = Nothing
    If Not outlookApp Is Nothing Then Set outlookApp = Nothing
End Sub
```

## Exemple complet : Système de notification automatique

```vba
Sub SystemeNotificationRapport()
    Dim outlookApp As Object
    Dim mail As Object
    Dim i As Integer
    Dim derniereLigne As Integer
    Dim dateRapport As String
    Dim nbLignes As Integer
    Dim resumeRapport As String

    On Error GoTo GestionErreur

    ' Préparer les données du rapport
    dateRapport = Format(Date, "dd/mm/yyyy")
    derniereLigne = Cells(Rows.Count, 1).End(xlUp).Row
    nbLignes = derniereLigne - 1  ' -1 pour exclure les en-têtes

    ' Créer un résumé des données
    resumeRapport = "Résumé du rapport :" & vbCrLf & vbCrLf
    resumeRapport = resumeRapport & "• Nombre d'entrées : " & nbLignes & vbCrLf

    ' Calculer quelques statistiques simples
    If nbLignes > 0 Then
        Dim total As Double
        For i = 2 To derniereLigne
            total = total + Cells(i, 2).Value  ' Supposons colonne B = montants
        Next i
        resumeRapport = resumeRapport & "• Total calculé : " & Format(total, "#,##0.00") & " €" & vbCrLf
        resumeRapport = resumeRapport & "• Moyenne : " & Format(total / nbLignes, "#,##0.00") & " €" & vbCrLf
    End If

    resumeRapport = resumeRapport & "• Date de génération : " & dateRapport & vbCrLf & vbCrLf
    resumeRapport = resumeRapport & "Le fichier Excel complet est joint à cet email."

    ' Sauvegarder le fichier
    ThisWorkbook.Save

    ' Créer et envoyer l'email
    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)

    With mail
        .To = "manager@exemple.com"
        .CC = "equipe@exemple.com"
        .Subject = "Rapport automatique du " & dateRapport & " - " & nbLignes & " entrées"
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
               "Voici le rapport automatique généré ce jour." & vbCrLf & vbCrLf & _
               resumeRapport & vbCrLf & vbCrLf & _
               "Ce rapport a été généré automatiquement par le système Excel." & vbCrLf & _
               "En cas de questions, merci de contacter l'équipe technique." & vbCrLf & vbCrLf & _
               "Cordialement," & vbCrLf & _
               "Système automatisé de reporting"

        ' Joindre le fichier Excel actuel
        .Attachments.Add ThisWorkbook.FullName

        ' Pour tester : afficher l'email
        .Display

        ' Pour envoyer automatiquement (décommenter quand vous êtes sûr)
        ' .Send
    End With

    ' Créer une tâche de suivi
    Dim tache As Object
    Set tache = outlookApp.CreateItem(3)
    With tache
        .Subject = "Vérifier la réception du rapport du " & dateRapport
        .DueDate = DateAdd("d", 1, Date)  ' Vérifier demain
        .Body = "Vérifier que le rapport automatique a bien été reçu et traité."
        .Save
    End With

    Set tache = Nothing
    Set mail = Nothing
    Set outlookApp = Nothing

    MsgBox "Notification préparée avec succès !" & vbCrLf & _
           "Email prêt à envoyer et tâche de suivi créée."

    Exit Sub

GestionErreur:
    MsgBox "Erreur dans le système de notification : " & Err.Description

    ' Nettoyage
    If Not tache Is Nothing Then Set tache = Nothing
    If Not mail Is Nothing Then Set mail = Nothing
    If Not outlookApp Is Nothing Then Set outlookApp = Nothing
End Sub
```

## Points importants à retenir

### ✅ Bonnes pratiques
- Toujours utiliser `.Display` avant `.Send` pour vérifier le contenu
- Libérer tous les objets avec `Set variable = Nothing`
- Utiliser la gestion d'erreurs pour les opérations Outlook
- Vérifier que les adresses email sont valides avant l'envoi
- Sauvegarder Excel avant de joindre le fichier actuel

### ⚠️ Erreurs courantes à éviter
- Envoyer des emails en masse sans vérification
- Oublier de gérer le cas où Outlook n'est pas installé
- Ne pas vérifier l'existence des fichiers à joindre
- Utiliser des adresses email en dur dans le code
- Créer trop de rendez-vous/tâches en boucle sans contrôle

### 🔒 Considérations de sécurité
- Outlook peut afficher des avertissements de sécurité pour l'automation
- Certaines entreprises bloquent l'envoi automatique d'emails
- Toujours tester avec de vraies adresses que vous contrôlez
- Éviter d'automatiser complètement l'envoi en production sans supervision

### 💡 Conseils pour débuter
- Commencez par créer des emails sans les envoyer (`.Display` seulement)
- Testez avec votre propre adresse email
- Utilisez la fenêtre Exécution immédiate (Ctrl+G) pour voir les résultats
- Gardez des modèles d'emails simples pour débuter

### 🎯 Utilisations typiques
- Envoi automatique de rapports quotidiens/hebdomadaires
- Notifications d'alertes basées sur des données Excel
- Création automatique de rendez-vous récurrents
- Synchronisation de contacts depuis des bases de données
- Système de rappels et de suivi automatisé

L'automation avec Outlook transforme Excel en véritable centre de communication automatisé. Une fois maîtrisée, cette fonctionnalité vous permettra de créer des systèmes de notification très sophistiqués !

⏭️
