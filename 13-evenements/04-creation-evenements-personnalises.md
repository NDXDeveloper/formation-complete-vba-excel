🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.4. Création d'événements personnalisés

## Que sont les Événements Personnalisés ?

Les événements personnalisés sont des événements que vous créez vous-même pour répondre à des besoins spécifiques de votre application. Contrairement aux événements prédéfinis d'Excel (comme `Worksheet_Change`), vous définissez vous-même quand et comment ces événements se déclenchent.

**Pourquoi créer des événements personnalisés ?**
- Créer une communication entre différentes parties de votre code
- Signaler des conditions spéciales dans votre application
- Rendre votre code plus modulaire et organisé
- Permettre à plusieurs procédures de réagir à une même situation

## Concepts de Base

### Analogie Simple
Imaginez que vous organisez une fête :
- **Événement prédéfini** : "Quelqu'un sonne à la porte" (vous ne contrôlez pas quand)
- **Événement personnalisé** : "La fête commence" (vous décidez quand le déclencher)

Dans votre code VBA, vous pouvez créer des "annonces" que d'autres parties de votre programme peuvent "écouter".

## Structure de Base des Événements Personnalisés

### 1. Déclaration de l'événement
```vba
' Dans un module de classe
Public Event NomDeLEvenement(paramètres)
```

### 2. Déclenchement de l'événement
```vba
' Quelque part dans votre code
RaiseEvent NomDeLEvenement(valeurs)
```

### 3. Écoute de l'événement
```vba
' Dans un autre module
Public WithEvents MonObjet As MaClasse

Private Sub MonObjet_NomDeLEvenement(paramètres)
    ' Code qui s'exécute quand l'événement se produit
End Sub
```

## Exemple Simple : Compteur avec Événements

### Étape 1 : Créer la classe avec l'événement

Créez un module de classe appelé "ClsCompteur" :

```vba
' Module de classe : ClsCompteur
Private mValeur As Integer

' Déclaration des événements personnalisés
Public Event ValeurChangee(NouvelleValeur As Integer)  
Public Event SeuilAtteint(Seuil As Integer)  

' Propriété pour lire la valeur
Public Property Get Valeur() As Integer
    Valeur = mValeur
End Property

' Propriété pour modifier la valeur (avec événement)
Public Property Let Valeur(NouvelleValeur As Integer)
    Dim ancienneValeur As Integer
    ancienneValeur = mValeur
    mValeur = NouvelleValeur

    ' Déclencher l'événement à chaque changement
    RaiseEvent ValeurChangee(mValeur)

    ' Déclencher un événement spécial si on atteint 100
    If mValeur >= 100 And ancienneValeur < 100 Then
        RaiseEvent SeuilAtteint(100)
    End If
End Property

' Méthode pour incrémenter
Public Sub Incrementer(Optional Pas As Integer = 1)
    Me.Valeur = mValeur + Pas
End Sub

' Méthode pour décrémenter
Public Sub Decrementer(Optional Pas As Integer = 1)
    Me.Valeur = mValeur - Pas
End Sub
```

### Étape 2 : Utiliser la classe avec les événements

**Important :** `WithEvents` ne peut être utilisé que dans un **module de classe** (y compris ThisWorkbook, les modules de feuilles et les UserForms), jamais dans un module standard.

Créez un module de classe appelé "ClsTestCompteur" :

```vba
' Module de classe : ClsTestCompteur
Public WithEvents MonCompteur As ClsCompteur

Sub TesterCompteur()
    ' Créer une instance du compteur
    Set MonCompteur = New ClsCompteur

    ' Modifier la valeur (déclenche les événements)
    MonCompteur.Valeur = 50
    MonCompteur.Incrementer 30
    MonCompteur.Incrementer 25  ' Dépassera 100

    MsgBox "Valeur finale : " & MonCompteur.Valeur
End Sub

' Gestionnaires d'événements
Private Sub MonCompteur_ValeurChangee(NouvelleValeur As Integer)
    Debug.Print "La valeur a changé : " & NouvelleValeur
End Sub

Private Sub MonCompteur_SeuilAtteint(Seuil As Integer)
    MsgBox "Félicitations ! Vous avez atteint " & Seuil & " !", vbInformation
End Sub
```

Pour lancer le test depuis un module standard :

```vba
' Module standard
Sub LancerTestCompteur()
    Dim test As New ClsTestCompteur
    test.TesterCompteur
End Sub
```

## Exemple Pratique : Système de Progression

Créons un système plus réaliste pour suivre la progression d'une tâche :

### Module de classe : ClsProgression

```vba
' Module de classe : ClsProgression
Private mPourcentage As Integer  
Private mTache As String  
Private mTerminee As Boolean  

' Événements personnalisés
Public Event ProgressionMiseAJour(Pourcentage As Integer, Tache As String)  
Public Event TacheTerminee(TacheName As String)  
Public Event ErreurProgression(MessageErreur As String)  

' Initialisation
Private Sub Class_Initialize()
    mPourcentage = 0
    mTache = ""
    mTerminee = False
End Sub

' Propriétés
Public Property Get Pourcentage() As Integer
    Pourcentage = mPourcentage
End Property

Public Property Get Tache() As String
    Tache = mTache
End Property

Public Property Get EstTerminee() As Boolean
    EstTerminee = mTerminee
End Property

' Méthode pour démarrer une tâche
Public Sub DemarrerTache(NomTache As String)
    If mTerminee Then
        RaiseEvent ErreurProgression("Une tâche est déjà terminée. Réinitialisez d'abord.")
        Exit Sub
    End If

    mTache = NomTache
    mPourcentage = 0
    RaiseEvent ProgressionMiseAJour(mPourcentage, mTache)
End Sub

' Méthode pour mettre à jour la progression
Public Sub MettreAJourProgression(NouveauPourcentage As Integer)
    ' Validation
    If NouveauPourcentage < 0 Or NouveauPourcentage > 100 Then
        RaiseEvent ErreurProgression("Le pourcentage doit être entre 0 et 100")
        Exit Sub
    End If

    If mTerminee Then
        RaiseEvent ErreurProgression("Cette tâche est déjà terminée")
        Exit Sub
    End If

    mPourcentage = NouveauPourcentage
    RaiseEvent ProgressionMiseAJour(mPourcentage, mTache)

    ' Vérifier si terminé
    If mPourcentage = 100 Then
        mTerminee = True
        RaiseEvent TacheTerminee(mTache)
    End If
End Sub

' Méthode pour ajouter à la progression
Public Sub AjouterProgression(Increment As Integer)
    MettreAJourProgression mPourcentage + Increment
End Sub

' Méthode pour réinitialiser
Public Sub Reinitialiser()
    mPourcentage = 0
    mTache = ""
    mTerminee = False
End Sub
```

### Utilisation du système de progression

Créez un module de classe "ClsTestProgression" (rappel : `WithEvents` ne fonctionne que dans un module de classe) :

```vba
' Module de classe : ClsTestProgression
Public WithEvents MaProgression As ClsProgression

Sub TesterProgression()
    ' Créer le système de progression
    Set MaProgression = New ClsProgression

    ' Simuler une tâche
    MaProgression.DemarrerTache "Traitement des données"

    ' Simuler l'avancement
    Dim i As Integer
    For i = 10 To 100 Step 10
        MaProgression.MettreAJourProgression i
        Application.Wait Now + TimeValue("00:00:01")  ' Pause d'1 seconde
    Next i
End Sub

' Gestionnaires d'événements
Private Sub MaProgression_ProgressionMiseAJour(Pourcentage As Integer, Tache As String)
    Application.StatusBar = Tache & " : " & Pourcentage & "% terminé"
    DoEvents  ' Permettre à Excel de se rafraîchir
End Sub

Private Sub MaProgression_TacheTerminee(TacheName As String)
    Application.StatusBar = False  ' Effacer la barre d'état
    MsgBox "Tâche terminée : " & TacheName, vbInformation
End Sub

Private Sub MaProgression_ErreurProgression(MessageErreur As String)
    MsgBox "Erreur : " & MessageErreur, vbExclamation
End Sub
```

## Exemple Avancé : Système de Notification Multi-Destinataires

Créons un système où plusieurs objets peuvent "s'abonner" aux notifications :

### Module de classe : ClsNotificateur

```vba
' Module de classe : ClsNotificateur
Private mAbonnes As Collection

' Événements
Public Event NotificationEnvoyee(Message As String, Priorite As Integer)  
Public Event AbonneAjoute(NomAbonne As String)  

Private Sub Class_Initialize()
    Set mAbonnes = New Collection
End Sub

' Ajouter un abonné
Public Sub AjouterAbonne(Nom As String, AdresseEmail As String)
    On Error GoTo GestionErreur

    ' Créer un dictionnaire pour stocker les infos
    Dim abonne As Object
    Set abonne = CreateObject("Scripting.Dictionary")
    abonne("Nom") = Nom
    abonne("Email") = AdresseEmail
    abonne("DateInscription") = Now()

    mAbonnes.Add abonne, Nom
    RaiseEvent AbonneAjoute(Nom)
    Exit Sub

GestionErreur:
    If Err.Number = 457 Then  ' Clé déjà existante
        MsgBox "L'abonné " & Nom & " existe déjà", vbExclamation
    End If
End Sub

' Envoyer une notification
Public Sub EnvoyerNotification(Message As String, Optional Priorite As Integer = 1)
    ' Simuler l'envoi à tous les abonnés
    Debug.Print "=== NOTIFICATION ==="
    Debug.Print "Message : " & Message
    Debug.Print "Priorité : " & Priorite
    Debug.Print "Destinataires : " & mAbonnes.Count

    Dim i As Integer
    For i = 1 To mAbonnes.Count
        Debug.Print "  -> " & mAbonnes(i)("Nom") & " (" & mAbonnes(i)("Email") & ")"
    Next i

    RaiseEvent NotificationEnvoyee(Message, Priorite)
End Sub

' Compter les abonnés
Public Function NombreAbonnes() As Integer
    NombreAbonnes = mAbonnes.Count
End Function
```

### Utilisation du système de notification

```vba
' Module de classe : ClsTestNotifications
Public WithEvents MonNotificateur As ClsNotificateur

Sub TesterNotifications()
    Set MonNotificateur = New ClsNotificateur

    ' Ajouter des abonnés
    MonNotificateur.AjouterAbonne "Jean Dupont", "jean@example.com"
    MonNotificateur.AjouterAbonne "Marie Martin", "marie@example.com"
    MonNotificateur.AjouterAbonne "Pierre Durand", "pierre@example.com"

    ' Envoyer des notifications
    MonNotificateur.EnvoyerNotification "Nouveau rapport disponible", 2
    MonNotificateur.EnvoyerNotification "Maintenance programmée", 3
End Sub

Private Sub MonNotificateur_NotificationEnvoyee(Message As String, Priorite As Integer)
    Dim couleur As String
    Select Case Priorite
        Case 1: couleur = "Verte"
        Case 2: couleur = "Orange"
        Case 3: couleur = "Rouge"
        Case Else: couleur = "Bleue"
    End Select

    MsgBox "Notification " & couleur & " envoyée !" & vbCrLf & Message, vbInformation
End Sub

Private Sub MonNotificateur_AbonneAjoute(NomAbonne As String)
    Debug.Print "Nouvel abonné : " & NomAbonne
End Sub
```

## Événements avec Paramètres Complexes

Vous pouvez passer des objets ou des structures complexes dans vos événements :

**Attention :** Les types personnalisés (Type...End Type) ne peuvent pas être utilisés comme paramètres d'événements en VBA. Utilisez plutôt des paramètres individuels ou un objet (classe) :

```vba
' Module de classe : ClsGestionnaireCommandes
Public Event CommandeCreee(ID As Long, Client As String, Montant As Double, DateCommande As Date)

Private mCommandes As Collection

Private Sub Class_Initialize()
    Set mCommandes = New Collection
End Sub

Public Sub CreerCommande(ID As Long, Client As String, Montant As Double)
    ' Stocker les données dans un dictionnaire
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info("ID") = ID
    info("Client") = Client
    info("Montant") = Montant
    info("DateCommande") = Now()

    mCommandes.Add info, CStr(ID)
    RaiseEvent CommandeCreee(ID, Client, Montant, CDate(info("DateCommande")))
End Sub
```

## Gestion d'Erreurs dans les Événements Personnalisés

```vba
' Dans la classe qui déclenche l'événement
Public Sub OperationRisquee()
    On Error GoTo GestionErreur

    ' Code qui peut générer une erreur
    Dim resultat As Double
    resultat = 10 / 0  ' Division par zéro

    RaiseEvent OperationReussie(resultat)
    Exit Sub

GestionErreur:
    RaiseEvent ErreurSurvenue(Err.Number, Err.Description)
End Sub

' Dans le module qui écoute
Private Sub MonObjet_ErreurSurvenue(NumeroErreur As Long, DescriptionErreur As String)
    MsgBox "Erreur " & NumeroErreur & " : " & DescriptionErreur, vbCritical
End Sub
```

## Avantages des Événements Personnalisés

### ✅ Avantages :
- **Séparation des responsabilités** : Chaque classe fait son travail
- **Réutilisabilité** : Le même événement peut être géré différemment
- **Flexibilité** : Facile d'ajouter de nouveaux gestionnaires
- **Communication asynchrone** : Les objets n'ont pas besoin de se connaître directement
- **Code plus propre** : Évite les dépendances circulaires

### Cas d'usage typiques :
- Systèmes de progression et de monitoring
- Validation de données complexes
- Communication entre formulaires
- Journalisation d'activités
- Notifications et alertes

## Bonnes Pratiques

### ✅ À faire :
- **Nommer clairement** vos événements (utiliser des verbes d'action)
- **Inclure tous les paramètres nécessaires** dans l'événement
- **Gérer les erreurs** dans les gestionnaires d'événements
- **Documenter** le comportement de vos événements
- **Tester** tous les scénarios possibles

### ❌ À éviter :
- **Événements trop fréquents** qui ralentissent l'application
- **Paramètres trop nombreux** (utiliser des types personnalisés si nécessaire)
- **Dépendances circulaires** entre objets avec événements
- **Oublier de gérer** les cas d'erreur
- **Événements sans gestionnaire** (vérifier qu'ils sont utiles)

## Exemple de Nommage Cohérent

```vba
' Bonnes pratiques de nommage
Public Event DonneesChargees(NombreEnregistrements As Long)  
Public Event ErreurChargement(MessageErreur As String)  
Public Event ProgressionMiseAJour(Pourcentage As Integer)  
Public Event ValidationEchouee(Champ As String, Raison As String)  
Public Event SauvegardeTerminee(CheminFichier As String)  

' Éviter les noms vagues
' Public Event Chose()
' Public Event Truc(X As Integer)
' Public Event Event1()
```

## Débogage des Événements Personnalisés

```vba
' Ajouter du logging pour déboguer
Public Sub EnvoyerNotification(Message As String)
    Debug.Print "DÉCLENCHEMENT: EnvoyerNotification - " & Message
    RaiseEvent NotificationEnvoyee(Message)
    Debug.Print "TERMINÉ: EnvoyerNotification"
End Sub

' Dans le gestionnaire
Private Sub MonObjet_NotificationEnvoyee(Message As String)
    Debug.Print "RÉCEPTION: NotificationEnvoyee - " & Message
    ' Votre code ici
    Debug.Print "TRAITEMENT: NotificationEnvoyee terminé"
End Sub
```

## Résumé

Les événements personnalisés permettent de :
- **Créer une architecture modulaire** pour vos applications VBA
- **Implémenter des patterns de communication** avancés
- **Séparer la logique métier** de l'interface utilisateur
- **Rendre votre code plus maintenable** et extensible

Ils représentent un concept avancé qui, une fois maîtrisé, ouvre de nouvelles possibilités pour créer des applications VBA sophistiquées et bien structurées. Dans la section suivante, nous apprendrons à désactiver temporairement les événements pour optimiser les performances.

⏭️
