🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.4. Collections personnalisées

## Qu'est-ce qu'une collection personnalisée ?

Une **collection personnalisée** est un objet qui permet de regrouper et gérer plusieurs objets du même type ou de types similaires. C'est comme un conteneur intelligent qui sait comment organiser, chercher et manipuler un groupe d'éléments.

**Analogie simple :**
Imaginez une **bibliothèque** :
- **Les livres** = les objets individuels (Employee, Product, etc.)
- **La bibliothèque** = la collection qui organise et gère les livres
- **Le bibliothécaire** = les méthodes qui permettent d'ajouter, chercher, retirer des livres

Une collection personnalisée fait la même chose avec vos objets VBA.

## Pourquoi utiliser des collections personnalisées ?

### 1. Organisation logique
Regrouper des objets liés (tous les employés d'une entreprise, tous les produits d'un catalogue)

### 2. Fonctionnalités spécialisées
Ajouter des méthodes de recherche, tri, filtrage spécifiques à votre métier

### 3. Validation centralisée
Contrôler ce qui peut être ajouté ou retiré de la collection

### 4. Interface simplifiée
Cacher la complexité de gestion des objets multiples

### 5. Performance
Optimiser les opérations sur de nombreux objets

## Collection VBA de base vs Collection personnalisée

### Collection VBA standard
```vba
Sub ExempleCollectionStandard()
    ' Collection VBA basique
    Dim employes As New Collection

    ' Ajout d'éléments
    employes.Add "Jean Dupont"
    employes.Add "Marie Martin"
    employes.Add "Paul Dubois"

    ' Accès aux éléments
    Debug.Print employes(1)        ' "Jean Dupont"
    Debug.Print employes.Count     ' 3

    ' Parcours
    Dim i As Integer
    For i = 1 To employes.Count
        Debug.Print employes(i)
    Next i
End Sub
```

**Limites de la collection standard :**
- Pas de validation des types
- Fonctionnalités limitées
- Pas de méthodes métier
- Difficile à maintenir

### Collection personnalisée
Voici ce que nous allons créer : une collection d'employés avec des fonctionnalités avancées.

## Exemple complet : Collection d'employés

### Étape 1 : Classe Employe (rappel simplifié)

```vba
' Module de classe : Employe
Private mNom As String
Private mPrenom As String
Private mSalaire As Double
Private mService As String
Private mDateEmbauche As Date

' Propriétés essentielles
Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(valeur As String)
    mNom = valeur
End Property

Public Property Get Prenom() As String
    Prenom = mPrenom
End Property

Public Property Let Prenom(valeur As String)
    mPrenom = valeur
End Property

Public Property Get Salaire() As Double
    Salaire = mSalaire
End Property

Public Property Let Salaire(valeur As Double)
    mSalaire = valeur
End Property

Public Property Get Service() As String
    Service = mService
End Property

Public Property Let Service(valeur As String)
    mService = valeur
End Property

Public Property Get DateEmbauche() As Date
    DateEmbauche = mDateEmbauche
End Property

Public Property Let DateEmbauche(valeur As Date)
    mDateEmbauche = valeur
End Property

' Méthodes utiles
Public Property Get NomComplet() As String
    NomComplet = mPrenom & " " & mNom
End Property

Public Property Get AncienneteAnnees() As Integer
    AncienneteAnnees = DateDiff("yyyy", mDateEmbauche, Date)
End Property

Public Function VersChaine() As String
    VersChaine = Me.NomComplet & " (" & mService & ") - " & Format(mSalaire, "#,##0") & "€"
End Function
```

### Étape 2 : Classe ListeEmployes (Collection personnalisée)

```vba
' Module de classe : ListeEmployes
Option Explicit

' ========== DONNÉES PRIVÉES ==========
Private mEmployes As Collection      ' Collection interne VBA
Private mNomEntreprise As String

' ========== ÉVÉNEMENTS ==========
Public Event EmployeAjoute(employe As Employe)
Public Event EmployeRetire(employe As Employe)
Public Event ListeVidee()

' ========== INITIALISATION ==========
Private Sub Class_Initialize()
    ' Constructeur - appelé automatiquement à la création
    Set mEmployes = New Collection
    mNomEntreprise = "Mon Entreprise"
End Sub

Private Sub Class_Terminate()
    ' Destructeur - appelé automatiquement à la destruction
    Set mEmployes = Nothing
End Sub

' ========== PROPRIÉTÉS PUBLIQUES ==========

' Nombre d'employés (lecture seule)
Public Property Get Count() As Long
    Count = mEmployes.Count
End Property

' Nom de l'entreprise
Public Property Get NomEntreprise() As String
    NomEntreprise = mNomEntreprise
End Property

Public Property Let NomEntreprise(valeur As String)
    mNomEntreprise = valeur
End Property

' Accès par index (lecture seule)
Public Property Get Item(index As Variant) As Employe
    ' index peut être un numéro (1, 2, 3...) ou une clé (nom)
    Set Item = mEmployes(index)
End Property

' Faire de Item la propriété par défaut
' (permet d'écrire liste(1) au lieu de liste.Item(1))
' Note: Ceci se configure dans les propriétés du module de classe

' ========== MÉTHODES D'AJOUT ==========

' Ajouter un employé existant
Public Sub Ajouter(employe As Employe, Optional cle As String = "")
    ' Validation
    If employe Is Nothing Then
        Err.Raise 5, , "Impossible d'ajouter un employé vide"
        Exit Sub
    End If

    ' Vérifier si l'employé existe déjà
    If Me.Existe(employe.NomComplet) Then
        Err.Raise 5, , "Un employé avec ce nom existe déjà : " & employe.NomComplet
        Exit Sub
    End If

    ' Déterminer la clé
    Dim cleUtilisee As String
    If cle = "" Then
        cleUtilisee = employe.NomComplet
    Else
        cleUtilisee = cle
    End If

    ' Ajouter à la collection interne
    mEmployes.Add employe, cleUtilisee

    ' Déclencher l'événement
    RaiseEvent EmployeAjoute(employe)

    Debug.Print "Employé ajouté : " & employe.NomComplet
End Sub

' Créer et ajouter un employé en une fois
Public Function CreerEmploye(nom As String, prenom As String, salaire As Double, service As String, Optional dateEmbauche As Date) As Employe
    ' Créer l'employé
    Dim nouveauEmploye As New Employe

    ' Configuration
    nouveauEmploye.Nom = nom
    nouveauEmploye.Prenom = prenom
    nouveauEmploye.Salaire = salaire
    nouveauEmploye.Service = service

    If dateEmbauche = 0 Then
        nouveauEmploye.DateEmbauche = Date
    Else
        nouveauEmploye.DateEmbauche = dateEmbauche
    End If

    ' Ajouter à la collection
    Me.Ajouter nouveauEmploye

    ' Retourner la référence
    Set CreerEmploye = nouveauEmploye
End Function

' ========== MÉTHODES DE SUPPRESSION ==========

' Retirer un employé par index ou clé
Public Sub Retirer(index As Variant)
    ' Validation
    If mEmployes.Count = 0 Then
        Err.Raise 5, , "La liste est vide"
        Exit Sub
    End If

    ' Récupérer l'employé avant de le supprimer (pour l'événement)
    Dim employeRetire As Employe
    Set employeRetire = mEmployes(index)

    ' Supprimer de la collection
    mEmployes.Remove index

    ' Déclencher l'événement
    RaiseEvent EmployeRetire(employeRetire)

    Debug.Print "Employé retiré : " & employeRetire.NomComplet
End Sub

' Vider toute la liste
Public Sub Vider()
    ' Supprimer tous les éléments
    Do While mEmployes.Count > 0
        mEmployes.Remove 1
    Loop

    ' Déclencher l'événement
    RaiseEvent ListeVidee

    Debug.Print "Liste d'employés vidée"
End Sub

' ========== MÉTHODES DE RECHERCHE ==========

' Vérifier si un employé existe
Public Function Existe(nomComplet As String) As Boolean
    Dim i As Long
    Existe = False

    For i = 1 To mEmployes.Count
        If mEmployes(i).NomComplet = nomComplet Then
            Existe = True
            Exit Function
        End If
    Next i
End Function

' Chercher un employé par nom complet
Public Function ChercherParNom(nomComplet As String) As Employe
    Dim i As Long
    Set ChercherParNom = Nothing

    For i = 1 To mEmployes.Count
        If mEmployes(i).NomComplet = nomComplet Then
            Set ChercherParNom = mEmployes(i)
            Exit Function
        End If
    Next i
End Function

' Chercher des employés par service
Public Function ChercherParService(service As String) As ListeEmployes
    Dim resultat As New ListeEmployes
    resultat.NomEntreprise = mNomEntreprise & " - Service " & service

    Dim i As Long
    For i = 1 To mEmployes.Count
        If UCase(mEmployes(i).Service) = UCase(service) Then
            resultat.Ajouter mEmployes(i)
        End If
    Next i

    Set ChercherParService = resultat
End Function

' Chercher des employés par salaire minimum
Public Function ChercherParSalaire(salaireMinimum As Double) As ListeEmployes
    Dim resultat As New ListeEmployes
    resultat.NomEntreprise = mNomEntreprise & " - Salaire >= " & Format(salaireMinimum, "#,##0") & "€"

    Dim i As Long
    For i = 1 To mEmployes.Count
        If mEmployes(i).Salaire >= salaireMinimum Then
            resultat.Ajouter mEmployes(i)
        End If
    Next i

    Set ChercherParSalaire = resultat
End Function

' ========== MÉTHODES STATISTIQUES ==========

' Calculer le salaire moyen
Public Function SalaireMoyen() As Double
    If mEmployes.Count = 0 Then
        SalaireMoyen = 0
        Exit Function
    End If

    Dim total As Double
    Dim i As Long

    For i = 1 To mEmployes.Count
        total = total + mEmployes(i).Salaire
    Next i

    SalaireMoyen = total / mEmployes.Count
End Function

' Calculer la masse salariale totale
Public Function MasseSalariale() As Double
    Dim i As Long
    MasseSalariale = 0

    For i = 1 To mEmployes.Count
        MasseSalariale = MasseSalariale + mEmployes(i).Salaire
    Next i
End Function

' Obtenir les services représentés
Public Function ListeServices() As Collection
    Dim services As New Collection
    Dim i As Long

    For i = 1 To mEmployes.Count
        Dim service As String
        service = mEmployes(i).Service

        ' Vérifier si le service existe déjà
        Dim existe As Boolean
        existe = False

        Dim j As Long
        For j = 1 To services.Count
            If services(j) = service Then
                existe = True
                Exit For
            End If
        Next j

        ' Ajouter si nouveau
        If Not existe Then
            services.Add service
        End If
    Next i

    Set ListeServices = services
End Function

' ========== MÉTHODES D'AFFICHAGE ==========

' Afficher tous les employés
Public Sub Afficher()
    Debug.Print "========== " & mNomEntreprise & " =========="
    Debug.Print "Nombre d'employés : " & mEmployes.Count

    If mEmployes.Count = 0 Then
        Debug.Print "Aucun employé"
    Else
        Dim i As Long
        For i = 1 To mEmployes.Count
            Debug.Print i & ". " & mEmployes(i).VersChaine()
        Next i

        Debug.Print "---"
        Debug.Print "Salaire moyen : " & Format(Me.SalaireMoyen(), "#,##0.00") & "€"
        Debug.Print "Masse salariale : " & Format(Me.MasseSalariale(), "#,##0.00") & "€"
    End If

    Debug.Print "=================================="
End Sub

' Générer un rapport par service
Public Sub RapportParService()
    Dim services As Collection
    Set services = Me.ListeServices()

    Debug.Print "========== RAPPORT PAR SERVICE =========="

    Dim i As Long
    For i = 1 To services.Count
        Dim service As String
        service = services(i)

        Dim employesService As ListeEmployes
        Set employesService = Me.ChercherParService(service)

        Debug.Print "--- " & UCase(service) & " ---"
        Debug.Print "Effectif : " & employesService.Count
        Debug.Print "Salaire moyen : " & Format(employesService.SalaireMoyen(), "#,##0.00") & "€"
        Debug.Print ""
    Next i

    Debug.Print "========================================="
End Sub

' ========== MÉTHODES D'ITÉRATION ==========

' Permettre l'utilisation de For Each (nécessite une configuration spéciale)
Public Function NewEnum() As IUnknown
    Set NewEnum = mEmployes.[_NewEnum]
End Function
```

## Utilisation de la collection personnalisée

### Exemple d'utilisation basique

```vba
Sub TestListeEmployes()
    ' Créer la liste
    Dim entreprise As New ListeEmployes
    entreprise.NomEntreprise = "TechCorp SARL"

    ' Ajouter des employés - Méthode 1 : Créer puis ajouter
    Dim emp1 As New Employe
    emp1.Nom = "Dupont"
    emp1.Prenom = "Jean"
    emp1.Salaire = 3500
    emp1.Service = "Informatique"
    emp1.DateEmbauche = #1/15/2020#
    entreprise.Ajouter emp1

    ' Ajouter des employés - Méthode 2 : Créer directement
    entreprise.CreerEmploye "Martin", "Marie", 4200, "Marketing", #3/10/2019#
    entreprise.CreerEmploye "Dubois", "Paul", 3800, "Informatique", #6/5/2021#
    entreprise.CreerEmploye "Leroy", "Sophie", 2900, "Comptabilité", #9/12/2022#
    entreprise.CreerEmploye "Bernard", "Luc", 5200, "Direction", #11/8/2018#

    ' Afficher la liste complète
    entreprise.Afficher

    ' Statistiques
    Debug.Print "Nombre total d'employés : " & entreprise.Count
    Debug.Print "Salaire moyen : " & Format(entreprise.SalaireMoyen(), "#,##0.00") & "€"

    ' Accès par index
    Debug.Print "Premier employé : " & entreprise.Item(1).NomComplet
    Debug.Print "Deuxième employé : " & entreprise(2).NomComplet  ' Syntaxe raccourcie

End Sub
```

### Exemple de recherches

```vba
Sub TestRechercheEmployes()
    ' Créer et remplir la liste (code simplifié)
    Dim entreprise As New ListeEmployes
    ' ... ajout d'employés ...

    ' Recherche par nom
    Dim employe As Employe
    Set employe = entreprise.ChercherParNom("Jean Dupont")
    If Not employe Is Nothing Then
        Debug.Print "Trouvé : " & employe.VersChaine()
    Else
        Debug.Print "Employé non trouvé"
    End If

    ' Recherche par service
    Dim informaticiens As ListeEmployes
    Set informaticiens = entreprise.ChercherParService("Informatique")
    Debug.Print "Employés en informatique : " & informaticiens.Count
    informaticiens.Afficher

    ' Recherche par salaire
    Dim cadres As ListeEmployes
    Set cadres = entreprise.ChercherParSalaire(4000)
    Debug.Print "Employés avec salaire >= 4000€ : " & cadres.Count
    cadres.Afficher

End Sub
```

### Exemple avec événements

```vba
' Dans un module standard
Public WithEvents maListe As ListeEmployes

Sub CreerListeAvecEvenements()
    Set maListe = New ListeEmployes
    maListe.NomEntreprise = "Entreprise avec événements"

    ' Les ajouts déclencheront automatiquement les événements
    maListe.CreerEmploye "Test", "Utilisateur", 3000, "Test"
End Sub

' Gestionnaires d'événements
Private Sub maListe_EmployeAjoute(employe As Employe)
    MsgBox "Nouvel employé ajouté : " & employe.NomComplet
End Sub

Private Sub maListe_EmployeRetire(employe As Employe)
    MsgBox "Employé retiré : " & employe.NomComplet
End Sub

Private Sub maListe_ListeVidee()
    MsgBox "La liste a été vidée"
End Sub
```

### Exemple d'itération avec For Each

```vba
Sub TestIterationForEach()
    Dim entreprise As New ListeEmployes
    ' ... remplir la liste ...

    ' Parcourir avec For Each (si NewEnum est configuré correctement)
    Dim employe As Employe
    For Each employe In entreprise
        Debug.Print employe.NomComplet & " - " & employe.Service
    Next employe

    ' Alternative : parcours classique
    Dim i As Long
    For i = 1 To entreprise.Count
        Debug.Print entreprise(i).NomComplet
    Next i
End Sub
```

## Techniques avancées

### 1. Collection avec tri

```vba
' Ajouter à la classe ListeEmployes
Public Sub TrierParNom()
    ' Tri à bulles simple (pour la démonstration)
    Dim i As Long, j As Long
    Dim temp As Employe

    For i = 1 To mEmployes.Count - 1
        For j = i + 1 To mEmployes.Count
            If mEmployes(i).NomComplet > mEmployes(j).NomComplet Then
                ' Échanger les positions (nécessite une collection temporaire)
                ' Note : VBA Collection ne permet pas l'échange direct
                ' Il faudrait utiliser un Array ou une autre structure
                ' Ceci est une version simplifiée
            End If
        Next j
    Next i
End Sub
```

### 2. Sérialisation (sauvegarde/chargement)

```vba
' Sauvegarder dans un fichier texte
Public Sub SauvegarderVersCSV(cheminFichier As String)
    Dim numeroFichier As Integer
    numeroFichier = FreeFile

    Open cheminFichier For Output As numeroFichier

    ' En-tête
    Print #numeroFichier, "Nom;Prénom;Salaire;Service;DateEmbauche"

    ' Données
    Dim i As Long
    For i = 1 To mEmployes.Count
        With mEmployes(i)
            Print #numeroFichier, .Nom & ";" & .Prenom & ";" & .Salaire & ";" & .Service & ";" & Format(.DateEmbauche, "dd/mm/yyyy")
        End With
    Next i

    Close numeroFichier
    Debug.Print "Liste sauvegardée dans : " & cheminFichier
End Sub
```

### 3. Collection filtrée dynamique

```vba
' Créer une vue filtrée sans copier les objets
Public Function CreerVue(critere As String) As ListeEmployes
    Dim vue As New ListeEmployes
    vue.NomEntreprise = mNomEntreprise & " - Vue : " & critere

    Dim i As Long
    For i = 1 To mEmployes.Count
        ' Ici vous pourriez implémenter une logique de filtrage complexe
        ' basée sur le critère passé en paramètre
        vue.Ajouter mEmployes(i)
    Next i

    Set CreerVue = vue
End Function
```

## Avantages des collections personnalisées

### 1. Type Safety (Sécurité des types)
```vba
' Collection VBA standard
Dim liste As Collection
liste.Add "Texte"
liste.Add 123
liste.Add Date  ' Mélange de types !

' Collection personnalisée
Dim entreprise As ListeEmployes
entreprise.Ajouter monEmploye    ' ✅ Seuls les employés acceptés
entreprise.Ajouter "Texte"       ' ❌ Erreur de compilation
```

### 2. Fonctionnalités métier
```vba
' Opérations spécialisées directement disponibles
entreprise.SalaireMoyen()
entreprise.ChercherParService("IT")
entreprise.RapportParService()
```

### 3. Validation centralisée
```vba
' Toutes les règles de validation dans un seul endroit
' Empêche les doublons, valide les données, etc.
```

### 4. Événements
```vba
' Notification automatique des changements
' Permet de créer des interfaces réactives
```

### 5. Interface claire
```vba
' Code plus lisible et expressif
Dim cadres = entreprise.ChercherParSalaire(4000)
' vs manipulation manuelle d'une Collection basique
```

## Bonnes pratiques

### 1. Nommage
- Classes collection : `Liste` + nom au pluriel (`ListeEmployes`, `ListeProduits`)
- Méthodes : verbes d'action (`Ajouter`, `Retirer`, `Chercher`)

### 2. Validation
- Toujours valider les paramètres
- Gérer les cas limites (liste vide, index invalide)

### 3. Événements
- Déclencher les événements aux bons moments
- Passer les informations pertinentes

### 4. Performance
- Pour de grandes collections, considérer des structures plus efficaces
- Éviter les recherches linéaires répétées

### 5. Documentation
- Documenter l'interface publique
- Expliquer les règles de validation
- Donner des exemples d'utilisation

Les collections personnalisées transforment votre code en rendant la gestion de groupes d'objets intuitive, sûre et puissante. Elles sont essentielles pour créer des applications VBA robustes et maintenables.

⏭️
