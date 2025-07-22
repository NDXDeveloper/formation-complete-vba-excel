🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.4 Fonctions système : Dir, Kill, MkDir, RmDir

## Introduction

VBA propose quatre fonctions système essentielles pour manipuler les fichiers et dossiers directement au niveau du système d'exploitation. Ces fonctions constituent la base de toutes les opérations sur le système de fichiers et sont indispensables pour automatiser la gestion de vos données.

## La fonction Dir

### Vue d'ensemble

La fonction `Dir` est l'outil principal pour explorer le système de fichiers. Elle permet de :
- Vérifier l'existence de fichiers et dossiers
- Lister le contenu d'un répertoire
- Rechercher des fichiers selon des critères
- Parcourir une arborescence

### Syntaxe de base

```vba
Dir(CheminRecherche, [Attributs])
```

- **CheminRecherche** : Le chemin et pattern de recherche
- **Attributs** : Type d'éléments à rechercher (optionnel)

### Vérifier l'existence d'un fichier

```vba
Sub VerifierFichier()
    Dim CheminFichier As String
    Dim Resultat As String

    CheminFichier = "C:\Temp\MonFichier.txt"

    ' Dir retourne le nom du fichier s'il existe, sinon une chaîne vide
    Resultat = Dir(CheminFichier)

    If Resultat <> "" Then
        MsgBox "Le fichier existe : " & Resultat
    Else
        MsgBox "Le fichier n'existe pas"
    End If
End Sub
```

### Vérifier l'existence d'un dossier

```vba
Sub VerifierDossier()
    Dim CheminDossier As String
    Dim Resultat As String

    CheminDossier = "C:\Temp\MonDossier"

    ' Utiliser l'attribut vbDirectory pour chercher des dossiers
    Resultat = Dir(CheminDossier, vbDirectory)

    If Resultat <> "" Then
        MsgBox "Le dossier existe : " & Resultat
    Else
        MsgBox "Le dossier n'existe pas"
    End If
End Sub
```

### Rechercher avec des caractères jokers

```vba
Sub RechercheAvecJokers()
    Dim NomFichier As String

    ' Rechercher tous les fichiers .txt
    NomFichier = Dir("C:\Temp\*.txt")

    If NomFichier <> "" Then
        Debug.Print "Premier fichier .txt trouvé : " & NomFichier

        ' Continuer la recherche
        Do
            NomFichier = Dir  ' Appel sans paramètre pour continuer
            If NomFichier <> "" Then
                Debug.Print "Fichier suivant : " & NomFichier
            End If
        Loop While NomFichier <> ""
    Else
        MsgBox "Aucun fichier .txt trouvé"
    End If
End Sub
```

### Caractères jokers disponibles

```vba
Sub ExemplesJokers()
    Dim Fichier As String

    ' * = n'importe quel nombre de caractères
    Fichier = Dir("C:\Temp\Rapport*.xlsx")  ' Rapport2024.xlsx, RapportVente.xlsx

    ' ? = exactement un caractère
    Fichier = Dir("C:\Temp\File?.txt")      ' File1.txt, FileA.txt

    ' Combinaisons
    Fichier = Dir("C:\Temp\Data_??_*.csv")  ' Data_01_janvier.csv, Data_12_decembre.csv

    Debug.Print "Premier fichier trouvé : " & Fichier
End Sub
```

### Attributs de recherche

```vba
Sub RechercheAvecAttributs()
    Dim Element As String

    ' vbNormal : fichiers normaux (par défaut)
    Element = Dir("C:\Temp\*.*", vbNormal)

    ' vbDirectory : dossiers uniquement
    Element = Dir("C:\Temp\*.*", vbDirectory)

    ' vbHidden : fichiers cachés
    Element = Dir("C:\Temp\*.*", vbHidden)

    ' vbReadOnly : fichiers en lecture seule
    Element = Dir("C:\Temp\*.*", vbReadOnly)

    ' Combinaisons avec + (OU logique)
    Element = Dir("C:\Temp\*.*", vbDirectory + vbHidden)

    Debug.Print "Premier élément trouvé : " & Element
End Sub
```

### Parcourir tous les fichiers d'un dossier

```vba
Sub ParcourrirTousFichiers()
    Dim CheminDossier As String
    Dim NomFichier As String
    Dim Compteur As Integer

    CheminDossier = "C:\Temp\"
    Compteur = 0

    ' Commencer la recherche
    NomFichier = Dir(CheminDossier & "*.*")

    Debug.Print "=== Fichiers dans " & CheminDossier & " ==="

    Do While NomFichier <> ""
        ' Vérifier que c'est bien un fichier (pas un dossier)
        If (GetAttr(CheminDossier & NomFichier) And vbDirectory) = 0 Then
            Compteur = Compteur + 1
            Debug.Print Compteur & ": " & NomFichier
        End If

        ' Passer au fichier suivant
        NomFichier = Dir
    Loop

    MsgBox "Total : " & Compteur & " fichiers"
End Sub
```

## La fonction Kill

### Vue d'ensemble

La fonction `Kill` supprime définitivement des fichiers du système. Elle ne déplace pas vers la corbeille, la suppression est immédiate et irréversible.

### Syntaxe

```vba
Kill CheminFichier
```

### Supprimer un fichier simple

```vba
Sub SupprimerFichier()
    Dim FichierASupprimer As String

    FichierASupprimer = "C:\Temp\FichierInutile.txt"

    ' Vérifier que le fichier existe
    If Dir(FichierASupprimer) <> "" Then
        On Error GoTo ErreurSuppression

        Kill FichierASupprimer
        MsgBox "Fichier supprimé : " & FichierASupprimer

        Exit Sub

ErreurSuppression:
        MsgBox "Impossible de supprimer le fichier : " & Err.Description
    Else
        MsgBox "Le fichier n'existe pas"
    End If
End Sub
```

### Supprimer plusieurs fichiers avec jokers

```vba
Sub SupprimerFichiersParType()
    Dim DossierCible As String
    Dim TypeFichier As String
    Dim Confirmation As VbMsgBoxResult

    DossierCible = "C:\Temp\"
    TypeFichier = "*.tmp"  ' Tous les fichiers temporaires

    ' Demander confirmation
    Confirmation = MsgBox("Supprimer tous les fichiers " & TypeFichier & " ?", _
                         vbYesNo + vbQuestion, "Confirmation")

    If Confirmation = vbYes Then
        On Error GoTo ErreurSuppression

        Kill DossierCible & TypeFichier
        MsgBox "Fichiers " & TypeFichier & " supprimés"

        Exit Sub

ErreurSuppression:
        MsgBox "Erreur lors de la suppression : " & Err.Description
    End If
End Sub
```

### Suppression sécurisée avec vérification

```vba
Sub SuppressionSecurisee()
    Dim FichierASupprimer As String
    Dim NomFichier As String
    Dim Compteur As Integer

    ' Supprimer tous les fichiers .log anciens
    FichierASupprimer = "C:\Temp\*.log"

    ' Compter d'abord les fichiers à supprimer
    NomFichier = Dir(FichierASupprimer)
    Compteur = 0

    Do While NomFichier <> ""
        Compteur = Compteur + 1
        NomFichier = Dir
    Loop

    If Compteur = 0 Then
        MsgBox "Aucun fichier .log à supprimer"
        Exit Sub
    End If

    ' Demander confirmation avec le nombre
    If MsgBox("Supprimer " & Compteur & " fichier(s) .log ?", _
              vbYesNo + vbQuestion) = vbYes Then

        On Error GoTo ErreurSuppression
        Kill FichierASupprimer
        MsgBox Compteur & " fichier(s) supprimé(s)"

        Exit Sub

ErreurSuppression:
        MsgBox "Erreur : " & Err.Description
    End If
End Sub
```

## La fonction MkDir

### Vue d'ensemble

La fonction `MkDir` crée un nouveau dossier. Elle ne peut créer qu'un seul niveau à la fois - le dossier parent doit exister.

### Syntaxe

```vba
MkDir CheminDossier
```

### Créer un dossier simple

```vba
Sub CreerDossierSimple()
    Dim NouveauDossier As String

    NouveauDossier = "C:\Temp\MonNouveauDossier"

    ' Vérifier que le dossier n'existe pas déjà
    If Dir(NouveauDossier, vbDirectory) = "" Then
        On Error GoTo ErreurCreation

        MkDir NouveauDossier
        MsgBox "Dossier créé : " & NouveauDossier

        Exit Sub

ErreurCreation:
        MsgBox "Impossible de créer le dossier : " & Err.Description
    Else
        MsgBox "Le dossier existe déjà"
    End If
End Sub
```

### Créer une hiérarchie de dossiers

```vba
Sub CreerHierarchieDossiers()
    Dim CheminComplet As String
    Dim PartiesChemin As Variant
    Dim CheminPartiel As String
    Dim i As Integer

    CheminComplet = "C:\Temp\Projet\Année2024\Trimestre1\Données"

    ' Diviser le chemin en parties
    PartiesChemin = Split(CheminComplet, "\")

    ' Reconstruire et créer chaque niveau
    For i = 0 To UBound(PartiesChemin)
        If i = 0 Then
            CheminPartiel = PartiesChemin(i) & "\"  ' C:\
        Else
            CheminPartiel = CheminPartiel & PartiesChemin(i) & "\"
        End If

        ' Créer le dossier s'il n'existe pas
        If Dir(CheminPartiel, vbDirectory) = "" And i > 0 Then
            On Error GoTo ErreurCreation
            MkDir Left(CheminPartiel, Len(CheminPartiel) - 1)  ' Enlever le \ final
            Debug.Print "Créé : " & Left(CheminPartiel, Len(CheminPartiel) - 1)
        End If
    Next i

    MsgBox "Hiérarchie créée : " & CheminComplet
    Exit Sub

ErreurCreation:
    MsgBox "Erreur lors de la création : " & Err.Description
End Sub
```

### Créer des dossiers avec noms dynamiques

```vba
Sub CreerDossiersAvecDates()
    Dim DossierBase As String
    Dim DossierAnnee As String
    Dim DossierMois As String
    Dim DossierJour As String

    DossierBase = "C:\Temp\Archives\"
    DossierAnnee = DossierBase & Year(Date) & "\"
    DossierMois = DossierAnnee & Format(Date, "mm-mmmm") & "\"
    DossierJour = DossierMois & Format(Date, "dd") & "\"

    ' Créer la structure jour par jour
    If Dir(DossierBase, vbDirectory) = "" Then MkDir DossierBase
    If Dir(DossierAnnee, vbDirectory) = "" Then MkDir DossierAnnee
    If Dir(DossierMois, vbDirectory) = "" Then MkDir DossierMois
    If Dir(DossierJour, vbDirectory) = "" Then MkDir DossierJour

    MsgBox "Structure créée : " & DossierJour
End Sub
```

## La fonction RmDir

### Vue d'ensemble

La fonction `RmDir` supprime un dossier vide. Le dossier doit être complètement vide (aucun fichier, aucun sous-dossier) pour pouvoir être supprimé.

### Syntaxe

```vba
RmDir CheminDossier
```

### Supprimer un dossier vide

```vba
Sub SupprimerDossierVide()
    Dim DossierASupprimer As String

    DossierASupprimer = "C:\Temp\DossierVide"

    ' Vérifier que le dossier existe
    If Dir(DossierASupprimer, vbDirectory) <> "" Then
        On Error GoTo ErreurSuppression

        RmDir DossierASupprimer
        MsgBox "Dossier supprimé : " & DossierASupprimer

        Exit Sub

ErreurSuppression:
        If Err.Number = 75 Then  ' Path/file access error
            MsgBox "Le dossier n'est pas vide ou est en cours d'utilisation"
        Else
            MsgBox "Erreur : " & Err.Description
        End If
    Else
        MsgBox "Le dossier n'existe pas"
    End If
End Sub
```

### Vérifier qu'un dossier est vide avant suppression

```vba
Function DossierEstVide(CheminDossier As String) As Boolean
    Dim Element As String

    ' Ajouter \ à la fin si nécessaire
    If Right(CheminDossier, 1) <> "\" Then
        CheminDossier = CheminDossier & "\"
    End If

    ' Chercher n'importe quel élément dans le dossier
    Element = Dir(CheminDossier & "*.*", vbDirectory + vbHidden + vbSystem)

    Do While Element <> ""
        ' Ignorer . et ..
        If Element <> "." And Element <> ".." Then
            DossierEstVide = False
            Exit Function
        End If
        Element = Dir
    Loop

    DossierEstVide = True
End Function

Sub SupprimerAvecVerification()
    Dim DossierASupprimer As String

    DossierASupprimer = "C:\Temp\TestDossier"

    If Dir(DossierASupprimer, vbDirectory) <> "" Then
        If DossierEstVide(DossierASupprimer) Then
            RmDir DossierASupprimer
            MsgBox "Dossier vide supprimé"
        Else
            MsgBox "Le dossier contient des éléments, impossible de le supprimer avec RmDir"
        End If
    Else
        MsgBox "Le dossier n'existe pas"
    End If
End Sub
```

### Supprimer une hiérarchie de dossiers vides

```vba
Sub SupprimerHierarchieVide()
    Dim DossiersASupprimer As Variant
    Dim i As Integer

    ' Liste des dossiers dans l'ordre inverse (du plus profond au plus superficiel)
    DossiersASupprimer = Array( _
        "C:\Temp\Projet\Année2024\Trimestre1\Données", _
        "C:\Temp\Projet\Année2024\Trimestre1", _
        "C:\Temp\Projet\Année2024", _
        "C:\Temp\Projet" _
    )

    For i = 0 To UBound(DossiersASupprimer)
        If Dir(DossiersASupprimer(i), vbDirectory) <> "" Then
            If DossierEstVide(DossiersASupprimer(i)) Then
                On Error Resume Next
                RmDir DossiersASupprimer(i)
                If Err.Number = 0 Then
                    Debug.Print "Supprimé : " & DossiersASupprimer(i)
                Else
                    Debug.Print "Échec : " & DossiersASupprimer(i)
                End If
                On Error GoTo 0
            Else
                Debug.Print "Non vide : " & DossiersASupprimer(i)
                Exit For  ' Arrêter si un dossier n'est pas vide
            End If
        End If
    Next i

    MsgBox "Nettoyage terminé"
End Sub
```

## Combinaison des fonctions - Exemples pratiques

### Nettoyage automatique de fichiers temporaires

```vba
Sub NettoyageFichiersTemporaires()
    Dim DossierTemp As String
    Dim FichierTemp As String
    Dim Compteur As Integer
    Dim DateLimite As Date

    DossierTemp = "C:\Temp\"
    DateLimite = Date - 7  ' Fichiers de plus de 7 jours
    Compteur = 0

    ' Rechercher les fichiers .tmp
    FichierTemp = Dir(DossierTemp & "*.tmp")

    Do While FichierTemp <> ""
        ' Vérifier la date du fichier
        If FileDateTime(DossierTemp & FichierTemp) < DateLimite Then
            On Error Resume Next
            Kill DossierTemp & FichierTemp
            If Err.Number = 0 Then
                Compteur = Compteur + 1
                Debug.Print "Supprimé : " & FichierTemp
            End If
            On Error GoTo 0
        End If

        FichierTemp = Dir
    Loop

    MsgBox "Nettoyage terminé. " & Compteur & " fichier(s) supprimé(s)"
End Sub
```

### Organiser les fichiers par date

```vba
Sub OrganiserFichiersParDate()
    Dim DossierSource As String
    Dim NomFichier As String
    Dim DateFichier As Date
    Dim DossierDestination As String
    Dim fso As Object

    DossierSource = "C:\Temp\Desorganise\"
    Set fso = CreateObject("Scripting.FileSystemObject")

    NomFichier = Dir(DossierSource & "*.*")

    Do While NomFichier <> ""
        ' Vérifier que c'est un fichier
        If (GetAttr(DossierSource & NomFichier) And vbDirectory) = 0 Then
            DateFichier = FileDateTime(DossierSource & NomFichier)
            DossierDestination = DossierSource & Format(DateFichier, "yyyy-mm") & "\"

            ' Créer le dossier du mois s'il n'existe pas
            If Dir(DossierDestination, vbDirectory) = "" Then
                MkDir DossierDestination
                Debug.Print "Dossier créé : " & DossierDestination
            End If

            ' Déplacer le fichier
            On Error Resume Next
            fso.MoveFile DossierSource & NomFichier, DossierDestination & NomFichier
            If Err.Number = 0 Then
                Debug.Print "Déplacé : " & NomFichier
            End If
            On Error GoTo 0
        End If

        NomFichier = Dir
    Loop

    Set fso = Nothing
    MsgBox "Organisation par date terminée"
End Sub
```

## Bonnes pratiques avec les fonctions système

### 1. Toujours gérer les erreurs

```vba
On Error GoTo GestionErreur
' Opérations système
Exit Sub

GestionErreur:
Select Case Err.Number
    Case 53  ' File not found
        MsgBox "Fichier introuvable"
    Case 75  ' Path/file access error
        MsgBox "Accès refusé ou dossier non vide"
    Case 76  ' Path not found
        MsgBox "Chemin introuvable"
    Case Else
        MsgBox "Erreur : " & Err.Description
End Select
```

### 2. Vérifier avant d'agir

```vba
' Avant de créer
If Dir(CheminDossier, vbDirectory) = "" Then
    MkDir CheminDossier
End If

' Avant de supprimer
If Dir(CheminFichier) <> "" Then
    Kill CheminFichier
End If
```

### 3. Utiliser des variables pour les chemins

```vba
Dim DossierTravail As String
Dim ExtensionRecherche As String

DossierTravail = "C:\MonApplication\"
ExtensionRecherche = "*.dat"

NomFichier = Dir(DossierTravail & ExtensionRecherche)
```

### 4. Documenter les opérations dangereuses

```vba
Sub SuppressionComplete()
    ' ATTENTION : Cette fonction supprime DÉFINITIVEMENT tous les fichiers .tmp
    ' Aucune récupération possible après exécution

    Dim Confirmation As VbMsgBoxResult
    Confirmation = MsgBox("DANGER : Supprimer tous les fichiers .tmp ?", _
                         vbYesNo + vbCritical, "Confirmation requise")

    If Confirmation = vbYes Then
        Kill "C:\Temp\*.tmp"
    End If
End Sub
```

## Messages d'erreur courants

### Erreur 53 - "File not found"
- **Cause** : Le fichier ou dossier spécifié n'existe pas
- **Solution** : Vérifier avec Dir() avant l'opération

### Erreur 75 - "Path/file access error"
- **Cause** : Fichier ouvert, dossier non vide, ou droits insuffisants
- **Solution** : Fermer les applications, vider le dossier, ou changer les droits

### Erreur 76 - "Path not found"
- **Cause** : Le chemin parent n'existe pas
- **Solution** : Créer d'abord les dossiers parents avec MkDir

---

*Dans la section suivante, nous découvrirons FileDialog, une interface moderne pour permettre aux utilisateurs de sélectionner des fichiers et dossiers de manière conviviale.*

⏭️
