🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.5. Modules de classe

## Qu'est-ce qu'un module de classe ?

Un **module de classe** est le conteneur dans lequel vous définissez vos classes en VBA. C'est l'outil qui vous permet de créer vos propres types d'objets personnalisés. Si on reprend l'analogie du plan d'architecte, le module de classe est le **support** sur lequel vous dessinez ce plan.

**Analogie simple :**
- **Module standard** = atelier avec des outils (procédures et fonctions globales)
- **Module de classe** = moule de fabrication (pour créer des objets identiques)
- **Feuille de calcul** = surface de travail avec données
- **UserForm** = interface graphique

Un module de classe est donc spécialement conçu pour définir la structure et le comportement de vos objets.

## Différences entre les types de modules

| Aspect | Module Standard | Module de Classe | Feuille Excel | UserForm |
|--------|----------------|------------------|---------------|----------|
| **Usage** | Procédures globales | Définition d'objets | Données et interface | Interface graphique |
| **Instanciation** | Non | Oui (`New`) | Une seule instance | Oui (`New`) |
| **Variables** | Globales/Locales | Membres d'objet | Cellules | Contrôles |
| **Événements** | Non | Oui (personnalisés) | Oui (Excel) | Oui (contrôles) |
| **Accès** | Direct | Via objets | Direct | Via objets |

## Organisation d'un projet avec modules de classe

### Structure recommandée

```
Mon Projet VBA/
├── Modules Standards/
│   ├── Module1 (Procédures principales)
│   ├── ModuleUtilitaire (Fonctions d'aide)
│   └── ModuleConstantes (Constantes globales)
├── Modules de Classe/
│   ├── Employe (Classe métier)
│   ├── ListeEmployes (Collection)
│   ├── GestionnaireRH (Logique métier)
│   └── ExportateurDonnees (Service)
├── Feuilles Excel/
│   ├── Feuil1 (Interface utilisateur)
│   └── Donnees (Stockage)
└── UserForms/
    ├── FormEmploye (Saisie)
    └── FormRapport (Affichage)
```

## Création et configuration d'un module de classe

### Étape 1 : Créer le module

1. **Éditeur VBA** (Alt + F11)
2. **Clic droit** sur votre projet dans l'Explorateur
3. **Insertion** → **Module de classe**
4. Un nouveau module "Class1" apparaît

### Étape 2 : Configurer les propriétés

Sélectionnez le module de classe et regardez la fenêtre **Propriétés** (F4) :

#### Propriétés importantes

```vba
' Propriétés du module de classe (visible dans la fenêtre Propriétés)

' (Name) = NomDeVotreClasse
' - Détermine le nom de la classe utilisable dans le code
' - Exemple : "Employe", "CompteBancaire", "Produit"

' Instancing = Private (par défaut)
' - Private : Classe utilisable seulement dans ce projet
' - PublicNotCreatable : Visible mais pas instanciable depuis un autre projet
```

**Note :** En VBA (Excel), les modules de classe n'ont que deux propriétés : `(Name)` et `Instancing`. D'autres propriétés comme `Persistable`, `DataBindingBehavior` ou `GlobalMultiUse` existent en VB6 (projets ActiveX) mais ne sont **pas disponibles** dans l'environnement VBA d'Excel.

### Configuration recommandée pour débutants

```vba
' Propriétés typiques pour une classe simple :
' (Name) : Le nom de votre classe
' Instancing : Private (sauf cas spéciaux)
' Tout le reste : valeurs par défaut
```

## Structure d'un module de classe

### Template de base

```vba
' ================================================================
' Module de classe : [NomDeVotreClasse]
' Description : [Description de ce que fait cette classe]
' Auteur : [Votre nom]
' Date : [Date de création]
' ================================================================

Option Explicit

' ========== ÉVÉNEMENTS (si nécessaire) ==========
' Déclaration des événements personnalisés
Public Event [NomEvenement]([parametres])

' ========== VARIABLES PRIVÉES ==========
' Toutes les données internes de l'objet
Private m[NomVariable] As [Type]

' ========== ÉVÉNEMENTS DE CLASSE ==========
' Constructeur et destructeur

Private Sub Class_Initialize()
    ' Code exécuté à la création de l'objet
End Sub

Private Sub Class_Terminate()
    ' Code exécuté à la destruction de l'objet
End Sub

' ========== PROPRIÉTÉS PUBLIQUES ==========
' Interface d'accès aux données

Public Property Get [NomPropriete]() As [Type]
    ' Lecture de la propriété
End Property

Public Property Let [NomPropriete](valeur As [Type])
    ' Écriture de la propriété (types simples)
End Property

Public Property Set [NomPropriete](valeur As [Type])
    ' Écriture de la propriété (objets)
End Property

' ========== MÉTHODES PUBLIQUES ==========
' Actions que l'objet peut effectuer

Public Sub [NomMethode]([parametres])
    ' Méthode d'action
End Sub

Public Function [NomFonction]([parametres]) As [Type]
    ' Méthode de calcul
End Function

' ========== MÉTHODES PRIVÉES ==========
' Fonctionnalités internes (aide)

Private Sub [MethodePrivee]([parametres])
    ' Code d'aide interne
End Sub

Private Function [FonctionPrivee]([parametres]) As [Type]
    ' Calcul interne
End Function
```

## Exemple complet : Module de classe DocumentWord

```vba
' ================================================================
' Module de classe : DocumentWord
' Description : Gestionnaire simplifié pour les documents Word
' ================================================================

Option Explicit

' ========== ÉVÉNEMENTS ==========
Public Event DocumentOuvert(cheminFichier As String)  
Public Event DocumentFerme(cheminFichier As String)  
Public Event ErreurTraitement(description As String)  

' ========== VARIABLES PRIVÉES ==========
Private mCheminFichier As String  
Private mTitre As String  
Private mContenu As String  
Private mEstModifie As Boolean  
Private mEstOuvert As Boolean  
Private mDateCreation As Date  
Private mTailleFichier As Long  

' ========== ÉVÉNEMENTS DE CLASSE ==========

Private Sub Class_Initialize()
    ' Constructeur - appelé automatiquement avec New
    mCheminFichier = ""
    mTitre = "Nouveau document"
    mContenu = ""
    mEstModifie = False
    mEstOuvert = False
    mDateCreation = Now
    mTailleFichier = 0

    Debug.Print "DocumentWord créé : " & Format(Now, "hh:nn:ss")
End Sub

Private Sub Class_Terminate()
    ' Destructeur - appelé automatiquement à la destruction
    If mEstOuvert Then
        Call Me.Fermer()
    End If

    Debug.Print "DocumentWord détruit : " & mTitre
End Sub

' ========== PROPRIÉTÉS PUBLIQUES ==========

' Chemin du fichier (lecture seule après ouverture)
Public Property Get CheminFichier() As String
    CheminFichier = mCheminFichier
End Property

' Titre du document
Public Property Get Titre() As String
    Titre = mTitre
End Property

Public Property Let Titre(valeur As String)
    If Len(Trim(valeur)) > 0 Then
        mTitre = Trim(valeur)
        mEstModifie = True
    Else
        Err.Raise 5, , "Le titre ne peut pas être vide"
    End If
End Property

' Contenu du document
Public Property Get Contenu() As String
    Contenu = mContenu
End Property

Public Property Let Contenu(valeur As String)
    mContenu = valeur
    mEstModifie = True
End Property

' État du document (lecture seule)
Public Property Get EstModifie() As Boolean
    EstModifie = mEstModifie
End Property

Public Property Get EstOuvert() As Boolean
    EstOuvert = mEstOuvert
End Property

Public Property Get DateCreation() As Date
    DateCreation = mDateCreation
End Property

Public Property Get TailleFichier() As Long
    TailleFichier = mTailleFichier
End Property

' Propriétés calculées
Public Property Get NomFichier() As String
    If mCheminFichier <> "" Then
        NomFichier = Mid(mCheminFichier, InStrRev(mCheminFichier, "\") + 1)
    Else
        NomFichier = "Sans nom"
    End If
End Property

Public Property Get Extension() As String
    Dim nom As String
    nom = Me.NomFichier

    If InStr(nom, ".") > 0 Then
        Extension = Mid(nom, InStrRev(nom, ".") + 1)
    Else
        Extension = ""
    End If
End Property

Public Property Get NombreCaracteres() As Long
    NombreCaracteres = Len(mContenu)
End Property

Public Property Get NombreMots() As Long
    If Len(Trim(mContenu)) = 0 Then
        NombreMots = 0
    Else
        NombreMots = UBound(Split(Trim(mContenu), " ")) + 1
    End If
End Property

' ========== MÉTHODES PUBLIQUES ==========

' Créer un nouveau document
Public Sub Nouveau(Optional titre As String = "Nouveau document")
    Call Me.Reinitialiser()
    mTitre = titre
    mEstOuvert = True

    Debug.Print "Nouveau document créé : " & mTitre
End Sub

' Ouvrir un fichier existant
Public Function Ouvrir(cheminFichier As String) As Boolean
    ' Validation du chemin
    If Not Me.FichierExiste(cheminFichier) Then
        RaiseEvent ErreurTraitement("Fichier introuvable : " & cheminFichier)
        Ouvrir = False
        Exit Function
    End If

    ' Simuler l'ouverture (en réalité, il faudrait utiliser l'API Word)
    mCheminFichier = cheminFichier
    mTitre = Me.NomFichier
    mEstOuvert = True
    mEstModifie = False

    ' Simuler la lecture du contenu
    Call Me.LireContenuFichier()

    ' Déclencher l'événement
    RaiseEvent DocumentOuvert(cheminFichier)

    Debug.Print "Document ouvert : " & cheminFichier
    Ouvrir = True
End Function

' Sauvegarder le document
Public Function Sauvegarder(Optional nouveauChemin As String = "") As Boolean
    ' Déterminer le chemin de sauvegarde
    Dim cheminSauvegarde As String
    If nouveauChemin <> "" Then
        cheminSauvegarde = nouveauChemin
        mCheminFichier = nouveauChemin
    ElseIf mCheminFichier <> "" Then
        cheminSauvegarde = mCheminFichier
    Else
        ' Nouveau document sans chemin - demander où sauvegarder
        cheminSauvegarde = Me.DemanderCheminSauvegarde()
        If cheminSauvegarde = "" Then
            Sauvegarder = False
            Exit Function
        End If
        mCheminFichier = cheminSauvegarde
    End If

    ' Simuler la sauvegarde
    Call Me.EcrireContenuFichier(cheminSauvegarde)
    mEstModifie = False

    Debug.Print "Document sauvegardé : " & cheminSauvegarde
    Sauvegarder = True
End Function

' Fermer le document
Public Sub Fermer()
    If mEstModifie Then
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Le document a été modifié. Voulez-vous le sauvegarder ?", _
                        vbYesNoCancel + vbQuestion, "Sauvegarder ?")

        Select Case reponse
            Case vbYes
                If Not Me.Sauvegarder() Then
                    Exit Sub  ' Annuler la fermeture si la sauvegarde échoue
                End If
            Case vbCancel
                Exit Sub  ' Annuler la fermeture
            ' Case vbNo : continuer sans sauvegarder
        End Select
    End If

    ' Déclencher l'événement avant fermeture
    RaiseEvent DocumentFerme(mCheminFichier)

    ' Réinitialiser l'état
    Call Me.Reinitialiser()

    Debug.Print "Document fermé"
End Sub

' Ajouter du texte
Public Sub AjouterTexte(texte As String, Optional sautLigne As Boolean = True)
    If sautLigne And Len(mContenu) > 0 Then
        mContenu = mContenu & vbCrLf & texte
    Else
        mContenu = mContenu & texte
    End If

    mEstModifie = True
End Sub

' Remplacer du texte
Public Function RemplacerTexte(ancien As String, nouveau As String) As Long
    Dim contenuOriginal As String
    contenuOriginal = mContenu

    mContenu = Replace(mContenu, ancien, nouveau)

    ' Compter le nombre de remplacements
    Dim nbRemplacements As Long
    If Len(ancien) <> Len(nouveau) Then
        nbRemplacements = (Len(contenuOriginal) - Len(mContenu)) / (Len(ancien) - Len(nouveau))
    Else
        ' Même longueur : comparer les contenus pour savoir s'il y a eu des changements
        nbRemplacements = IIf(contenuOriginal <> mContenu, _
            (Len(contenuOriginal) - Len(Replace(contenuOriginal, ancien, ""))) / Len(ancien), 0)
    End If

    If nbRemplacements > 0 Then
        mEstModifie = True
        Debug.Print nbRemplacements & " remplacement(s) effectué(s)"
    End If

    RemplacerTexte = nbRemplacements
End Function

' Chercher du texte
Public Function ChercherTexte(recherche As String) As Long
    ChercherTexte = InStr(1, mContenu, recherche, vbTextCompare)
End Function

' Vider le contenu
Public Sub Vider()
    mContenu = ""
    mEstModifie = True
End Sub

' Afficher les informations du document
Public Sub AfficherInfos()
    Debug.Print "========== INFORMATIONS DOCUMENT =========="
    Debug.Print "Titre : " & mTitre
    Debug.Print "Fichier : " & IIf(mCheminFichier = "", "Non sauvegardé", mCheminFichier)
    Debug.Print "État : " & IIf(mEstOuvert, "Ouvert", "Fermé")
    Debug.Print "Modifié : " & IIf(mEstModifie, "Oui", "Non")
    Debug.Print "Création : " & Format(mDateCreation, "dd/mm/yyyy hh:nn:ss")
    Debug.Print "Caractères : " & Format(Me.NombreCaracteres, "#,##0")
    Debug.Print "Mots : " & Format(Me.NombreMots, "#,##0")
    Debug.Print "Taille : " & Me.FormatTaille(mTailleFichier)
    Debug.Print "=========================================="
End Sub

' Exporter vers un format simple
Public Function ExporterVersTexte(cheminDestination As String) As Boolean
    On Error GoTo GestionErreur

    Dim numeroFichier As Integer
    numeroFichier = FreeFile

    Open cheminDestination For Output As numeroFichier
    Print #numeroFichier, "Titre: " & mTitre
    Print #numeroFichier, "Date: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Print #numeroFichier, String(50, "-")
    Print #numeroFichier, mContenu
    Close numeroFichier

    Debug.Print "Document exporté vers : " & cheminDestination
    ExporterVersTexte = True
    Exit Function

GestionErreur:
    Close numeroFichier
    RaiseEvent ErreurTraitement("Erreur lors de l'export : " & Err.Description)
    ExporterVersTexte = False
End Function

' ========== MÉTHODES PRIVÉES ==========

Private Sub Reinitialiser()
    mCheminFichier = ""
    mTitre = "Nouveau document"
    mContenu = ""
    mEstModifie = False
    mEstOuvert = False
    mTailleFichier = 0
End Sub

Private Function FichierExiste(chemin As String) As Boolean
    FichierExiste = (Dir(chemin) <> "")
End Function

Private Sub LireContenuFichier()
    ' Simuler la lecture du fichier
    ' En réalité, il faudrait utiliser l'API Word ou lire un fichier texte
    mContenu = "Contenu simulé du fichier : " & Me.NomFichier & vbCrLf & _
               "Chargé le " & Format(Now, "dd/mm/yyyy à hh:nn:ss")

    ' Simuler la taille du fichier
    mTailleFichier = Len(mContenu) * 2  ' Estimation
End Sub

Private Sub EcrireContenuFichier(chemin As String)
    ' Simuler l'écriture du fichier
    ' En réalité, il faudrait utiliser l'API Word ou écrire un fichier texte
    mTailleFichier = Len(mContenu) * 2
    Debug.Print "Simulation : écriture dans " & chemin
End Sub

Private Function DemanderCheminSauvegarde() As String
    ' Simuler une boîte de dialogue de sauvegarde
    ' En réalité, il faudrait utiliser Application.GetSaveAsFilename
    DemanderCheminSauvegarde = "C:\Temp\" & mTitre & ".txt"
End Function

Private Function FormatTaille(taille As Long) As String
    If taille < 1024 Then
        FormatTaille = taille & " octets"
    ElseIf taille < 1048576 Then
        FormatTaille = Format(taille / 1024, "#,##0.0") & " Ko"
    Else
        FormatTaille = Format(taille / 1048576, "#,##0.0") & " Mo"
    End If
End Function
```

## Utilisation du module de classe

**Important :** `WithEvents` doit être déclaré au niveau module (pas dans une procédure), et uniquement dans un module de classe.

```vba
' Module de classe : ClsTestDocument
' WithEvents doit être déclaré au niveau module
Public WithEvents monDoc As DocumentWord

Sub TestDocumentWord()
    ' Créer un document avec événements
    Set monDoc = New DocumentWord

    ' Créer un nouveau document
    monDoc.Nouveau "Mon rapport mensuel"

    ' Ajouter du contenu
    monDoc.AjouterTexte "RAPPORT MENSUEL", True
    monDoc.AjouterTexte String(20, "="), True
    monDoc.AjouterTexte "Ce rapport contient les données du mois.", True
    monDoc.AjouterTexte "Nombre de ventes : 150", True
    monDoc.AjouterTexte "Chiffre d'affaires : 45 000€", True

    ' Afficher les informations
    monDoc.AfficherInfos

    ' Chercher et remplacer
    Dim nbRemplacements As Long
    nbRemplacements = monDoc.RemplacerTexte("150", "180")

    ' Sauvegarder
    monDoc.Sauvegarder "C:\Temp\rapport.txt"

    ' Exporter
    monDoc.ExporterVersTexte "C:\Temp\rapport_export.txt"

    ' Fermer
    monDoc.Fermer
End Sub

' Gestionnaires d'événements
Private Sub monDoc_DocumentOuvert(cheminFichier As String)
    Debug.Print "Événement : Document ouvert - " & cheminFichier
End Sub

Private Sub monDoc_DocumentFerme(cheminFichier As String)
    Debug.Print "Événement : Document fermé - " & cheminFichier
End Sub

Private Sub monDoc_ErreurTraitement(description As String)
    Debug.Print "Événement : Erreur - " & description
End Sub
```

Pour lancer le test depuis un module standard :

```vba
' Module standard
Sub LancerTestDocument()
    Dim test As New ClsTestDocument
    test.TestDocumentWord
End Sub
```

## Bonnes pratiques pour les modules de classe

### 1. Organisation du code

```vba
' ✅ Bonne organisation
' 1. Option Explicit en premier
' 2. Commentaires d'en-tête
' 3. Événements
' 4. Variables privées
' 5. Class_Initialize/Terminate
' 6. Propriétés
' 7. Méthodes publiques
' 8. Méthodes privées
```

### 2. Nommage cohérent

```vba
' ✅ Conventions recommandées
' - Classes : PascalCase (DocumentWord, GestionnaireStock)
' - Variables membres : mNomVariable
' - Propriétés : PascalCase (CheminFichier, EstModifie)
' - Méthodes : PascalCase + Verbe (Ouvrir, Sauvegarder, Chercher)
' - Événements : PascalCase (DocumentOuvert, ErreurTraitement)
```

### 3. Gestion de la mémoire

```vba
' ✅ Libération des ressources
Private Sub Class_Terminate()
    ' Nettoyer les objets
    Set mObjetInterne = Nothing
    ' Fermer les fichiers ouverts
    ' Libérer les ressources système
End Sub
```

### 4. Validation des paramètres

```vba
' ✅ Toujours valider
Public Property Let Nom(valeur As String)
    If Len(Trim(valeur)) = 0 Then
        Err.Raise 5, , "Le nom ne peut pas être vide"
    End If
    mNom = Trim(valeur)
End Property
```

### 5. Documentation

```vba
' ✅ Documenter l'interface publique
Public Function Chercher(critere As String) As Collection
    ' Recherche des éléments selon un critère
    ' Paramètres :
    '   critere : Texte à rechercher (sensible à la casse)
    ' Retour :
    '   Collection des éléments trouvés (vide si aucun)
    ' Exemple :
    '   Dim resultats As Collection
    '   Set resultats = monObjet.Chercher("important")
End Function
```

### 6. Gestion d'erreurs

```vba
' ✅ Gestion cohérente des erreurs
Public Function TraiterFichier(chemin As String) As Boolean
    On Error GoTo GestionErreur

    ' Code de traitement...
    TraiterFichier = True
    Exit Function

GestionErreur:
    ' Log de l'erreur
    Debug.Print "Erreur dans TraiterFichier : " & Err.Description
    ' Déclencher un événement d'erreur
    RaiseEvent ErreurTraitement(Err.Description)
    TraiterFichier = False
End Function
```

## Avantages des modules de classe bien organisés

### 1. Maintenabilité
- Code structuré et prévisible
- Modification facile et sûre
- Recherche rapide des fonctionnalités

### 2. Réutilisabilité
- Classes copiables entre projets
- Interface standardisée
- Documentation intégrée

### 3. Collaboration
- Structure claire pour le travail en équipe
- Conventions de nommage cohérentes
- Séparation claire des responsabilités

### 4. Débogage
- Localisation rapide des problèmes
- Gestion d'erreurs centralisée
- Traçabilité des opérations

Les modules de classe bien structurés sont la fondation d'applications VBA robustes et professionnelles. Ils transforment vos idées en outils réutilisables et maintenables.

⏭️
