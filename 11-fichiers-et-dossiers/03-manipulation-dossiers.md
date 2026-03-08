🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.3 Manipulation de dossiers

## Introduction

La manipulation de dossiers est essentielle pour organiser automatiquement vos fichiers, créer des structures de répertoires et gérer l'espace de stockage. VBA offre plusieurs outils pour créer, supprimer, naviguer et analyser les dossiers de votre système.

## Comprendre les chemins de dossiers

### Types de chemins

**Chemin absolu** : Indique l'emplacement complet depuis la racine
```vba
"C:\Users\MonNom\Documents\MonProjet\"
```

**Chemin relatif** : Par rapport au dossier courant
```vba
".\SousDossier\"          ' Dossier enfant
"..\DossierParent\"       ' Dossier parent
"..\..\DeuxNiveauxAuDessus\" ' Deux niveaux au-dessus
```

### Séparateurs et conventions

En VBA, utilisez toujours le backslash `\` pour Windows :
```vba
' Correct
"C:\Mes Documents\Projet\"

' Incorrect (peut causer des erreurs)
"C:/Mes Documents/Projet/"
```

## Vérifier l'existence d'un dossier

### Utiliser la fonction Dir avec vbDirectory

```vba
Sub VerifierExistenceDossier()
    Dim CheminDossier As String

    CheminDossier = "C:\Temp\MonDossier"

    ' Vérifier si le dossier existe
    If Dir(CheminDossier, vbDirectory) <> "" Then
        MsgBox "Le dossier existe !"
    Else
        MsgBox "Le dossier n'existe pas."
    End If
End Sub
```

### Fonction personnalisée pour vérifier un dossier

```vba
Function DossierExiste(ByVal CheminDossier As String) As Boolean
    ' Ajouter un backslash à la fin si nécessaire
    If Right(CheminDossier, 1) <> "\" Then
        CheminDossier = CheminDossier & "\"
    End If

    ' Vérifier l'existence
    DossierExiste = (Dir(CheminDossier, vbDirectory) <> "")
End Function

Sub ExempleUtilisationFunction()
    If DossierExiste("C:\Temp\TestDossier") Then
        MsgBox "Le dossier existe"
    Else
        MsgBox "Le dossier n'existe pas"
    End If
End Sub
```

## Créer des dossiers

### Utiliser MkDir pour créer un dossier

```vba
Sub CreerDossierSimple()
    Dim NouveauDossier As String

    NouveauDossier = "C:\Temp\MonNouveauDossier"

    ' Vérifier que le dossier n'existe pas déjà
    If Not DossierExiste(NouveauDossier) Then
        MkDir NouveauDossier
        MsgBox "Dossier créé : " & NouveauDossier
    Else
        MsgBox "Le dossier existe déjà !"
    End If
End Sub
```

### Créer une arborescence complète

```vba
Sub CreerArborescenceComplete()
    Dim CheminBase As String
    Dim SousDossiers As Variant
    Dim i As Integer

    CheminBase = "C:\Temp\MonProjet\"

    ' Liste des sous-dossiers à créer
    SousDossiers = Array("Documents", "Images", "Données", "Rapports", "Archive")

    ' Créer le dossier principal
    If Not DossierExiste(CheminBase) Then
        MkDir CheminBase
    End If

    ' Créer chaque sous-dossier
    For i = 0 To UBound(SousDossiers)
        Dim CheminComplet As String
        CheminComplet = CheminBase & SousDossiers(i)

        If Not DossierExiste(CheminComplet) Then
            MkDir CheminComplet
            Debug.Print "Créé : " & CheminComplet
        End If
    Next i

    MsgBox "Arborescence créée dans : " & CheminBase
End Sub
```

### Créer des dossiers avec dates

```vba
Sub CreerDossierAvecDate()
    Dim DossierBase As String
    Dim DossierDate As String
    Dim CheminComplet As String

    DossierBase = "C:\Temp\Rapports\"
    DossierDate = Format(Date, "yyyy-mm-dd")
    CheminComplet = DossierBase & DossierDate

    ' Créer le dossier de base s'il n'existe pas
    If Not DossierExiste(DossierBase) Then
        MkDir DossierBase
    End If

    ' Créer le dossier avec la date
    If Not DossierExiste(CheminComplet) Then
        MkDir CheminComplet
        MsgBox "Dossier créé : " & CheminComplet
    Else
        MsgBox "Le dossier du jour existe déjà"
    End If
End Sub
```

## Supprimer des dossiers

### Utiliser RmDir pour supprimer un dossier vide

```vba
Sub SupprimerDossierVide()
    Dim DossierASupprimer As String

    DossierASupprimer = "C:\Temp\DossierVide"

    ' Vérifier que le dossier existe
    If DossierExiste(DossierASupprimer) Then
        On Error GoTo ErreurSuppression

        RmDir DossierASupprimer
        MsgBox "Dossier supprimé : " & DossierASupprimer
        Exit Sub

ErreurSuppression:
        MsgBox "Impossible de supprimer le dossier. Il contient peut-être des fichiers."
    Else
        MsgBox "Le dossier n'existe pas."
    End If
End Sub
```

### Supprimer un dossier et son contenu avec FileSystemObject

```vba
Sub SupprimerDossierEtContenu()
    Dim fso As Object
    Dim DossierASupprimer As String

    DossierASupprimer = "C:\Temp\DossierASupprimer"

    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Vérifier que le dossier existe
    If fso.FolderExists(DossierASupprimer) Then
        On Error GoTo ErreurSuppression

        ' Supprimer le dossier et tout son contenu
        fso.DeleteFolder DossierASupprimer
        MsgBox "Dossier et contenu supprimés : " & DossierASupprimer

        Exit Sub

ErreurSuppression:
        MsgBox "Erreur lors de la suppression : " & Err.Description
    Else
        MsgBox "Le dossier n'existe pas."
    End If

    Set fso = Nothing
End Sub
```

## Naviguer dans les dossiers

### Obtenir le dossier courant

```vba
Sub AfficherDossierCourant()
    Dim DossierActuel As String

    DossierActuel = CurDir
    MsgBox "Dossier courant : " & DossierActuel
End Sub
```

### Changer de dossier courant

```vba
Sub ChangerDossierCourant()
    Dim AncienDossier As String
    Dim NouveauDossier As String

    ' Sauvegarder le dossier actuel
    AncienDossier = CurDir

    NouveauDossier = "C:\Temp"

    ' Vérifier que le nouveau dossier existe
    If DossierExiste(NouveauDossier) Then
        ChDir NouveauDossier
        MsgBox "Nouveau dossier courant : " & CurDir

        ' Revenir au dossier précédent
        ChDir AncienDossier
        MsgBox "Retour au dossier : " & CurDir
    Else
        MsgBox "Le dossier " & NouveauDossier & " n'existe pas"
    End If
End Sub
```

## Lister le contenu d'un dossier

### Lister tous les fichiers

```vba
Sub ListerFichiersDossier()
    Dim CheminDossier As String
    Dim NomFichier As String
    Dim Compteur As Integer

    CheminDossier = "C:\Temp\"

    If Not DossierExiste(CheminDossier) Then
        MsgBox "Le dossier n'existe pas !"
        Exit Sub
    End If

    ' Commencer la recherche
    NomFichier = Dir(CheminDossier & "*.*")
    Compteur = 0

    Debug.Print "=== Contenu de " & CheminDossier & " ==="

    Do While NomFichier <> ""
        ' Vérifier si c'est un fichier (pas un dossier)
        If (GetAttr(CheminDossier & NomFichier) And vbDirectory) = 0 Then
            Compteur = Compteur + 1
            Debug.Print Compteur & ": " & NomFichier
        End If

        ' Passer au fichier suivant
        NomFichier = Dir
    Loop

    MsgBox "Nombre de fichiers trouvés : " & Compteur
End Sub
```

### Lister seulement les dossiers

```vba
Sub ListerSousDossiers()
    Dim CheminDossier As String
    Dim NomElement As String
    Dim Compteur As Integer

    CheminDossier = "C:\Temp\"

    If Not DossierExiste(CheminDossier) Then
        MsgBox "Le dossier n'existe pas !"
        Exit Sub
    End If

    ' Rechercher les dossiers
    NomElement = Dir(CheminDossier & "*.*", vbDirectory)
    Compteur = 0

    Debug.Print "=== Sous-dossiers de " & CheminDossier & " ==="

    Do While NomElement <> ""
        ' Vérifier que c'est un dossier et pas "." ou ".."
        If (GetAttr(CheminDossier & NomElement) And vbDirectory) = vbDirectory Then
            If NomElement <> "." And NomElement <> ".." Then
                Compteur = Compteur + 1
                Debug.Print Compteur & ": " & NomElement
            End If
        End If

        NomElement = Dir
    Loop

    MsgBox "Nombre de sous-dossiers : " & Compteur
End Sub
```

### Lister avec filtres

```vba
Sub ListerFichiersAvecFiltre()
    Dim CheminDossier As String
    Dim Extension As String
    Dim NomFichier As String
    Dim Compteur As Integer

    CheminDossier = "C:\Temp\"
    Extension = "*.txt"  ' Seulement les fichiers .txt

    If Not DossierExiste(CheminDossier) Then
        MsgBox "Le dossier n'existe pas !"
        Exit Sub
    End If

    NomFichier = Dir(CheminDossier & Extension)
    Compteur = 0

    Debug.Print "=== Fichiers " & Extension & " dans " & CheminDossier & " ==="

    Do While NomFichier <> ""
        Compteur = Compteur + 1
        Debug.Print Compteur & ": " & NomFichier
        NomFichier = Dir
    Loop

    MsgBox "Fichiers " & Extension & " trouvés : " & Compteur
End Sub
```

## Fonctions utilitaires avancées

### Obtenir des informations sur un dossier

```vba
Sub InformationsDossier()
    Dim fso As Object
    Dim dossier As Object
    Dim CheminDossier As String

    CheminDossier = "C:\Temp"

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(CheminDossier) Then
        Set dossier = fso.GetFolder(CheminDossier)

        Debug.Print "=== Informations sur " & CheminDossier & " ==="
        Debug.Print "Nom : " & dossier.Name
        Debug.Print "Chemin complet : " & dossier.Path
        Debug.Print "Date de création : " & dossier.DateCreated
        Debug.Print "Date de modification : " & dossier.DateLastModified
        Debug.Print "Taille : " & dossier.Size & " octets"
        Debug.Print "Nombre de fichiers : " & dossier.Files.Count
        Debug.Print "Nombre de sous-dossiers : " & dossier.SubFolders.Count

        MsgBox "Informations affichées dans la fenêtre Immédiate"
    Else
        MsgBox "Le dossier n'existe pas"
    End If

    Set dossier = Nothing
    Set fso = Nothing
End Sub
```

### Copier un dossier

```vba
Sub CopierDossier()
    Dim fso As Object
    Dim DossierSource As String
    Dim DossierDestination As String

    DossierSource = "C:\Temp\DossierACopier"
    DossierDestination = "C:\Temp\CopieDossier"

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(DossierSource) Then
        On Error GoTo ErreurCopie

        ' Copier le dossier et tout son contenu
        fso.CopyFolder DossierSource, DossierDestination
        MsgBox "Dossier copié avec succès !"

        Exit Sub

ErreurCopie:
        MsgBox "Erreur lors de la copie : " & Err.Description
    Else
        MsgBox "Le dossier source n'existe pas"
    End If

    Set fso = Nothing
End Sub
```

### Déplacer un dossier

```vba
Sub DeplacerDossier()
    Dim fso As Object
    Dim DossierSource As String
    Dim DossierDestination As String

    DossierSource = "C:\Temp\DossierADeplacer"
    DossierDestination = "C:\Temp\NouvelEmplacement\"

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(DossierSource) Then
        On Error GoTo ErreurDeplacement

        ' Déplacer le dossier
        fso.MoveFolder DossierSource, DossierDestination
        MsgBox "Dossier déplacé avec succès !"

        Exit Sub

ErreurDeplacement:
        MsgBox "Erreur lors du déplacement : " & Err.Description
    Else
        MsgBox "Le dossier source n'existe pas"
    End If

    Set fso = Nothing
End Sub
```

## Exemples pratiques complets

### Organiser des fichiers par extension

```vba
Sub OrganiserFichiersParExtension()
    Dim DossierSource As String
    Dim fso As Object
    Dim NomFichier As String
    Dim Extension As String
    Dim DossierExtension As String

    DossierSource = "C:\Temp\FilesDesorganises\"

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(DossierSource) Then
        MsgBox "Dossier source introuvable !"
        Exit Sub
    End If

    ' Étape 1 : Collecter tous les noms de fichiers
    ' (Ne pas modifier les fichiers pendant l'itération Dir !)
    Dim fichiers() As String
    Dim nbFichiers As Long
    nbFichiers = 0

    NomFichier = Dir(DossierSource & "*.*")
    Do While NomFichier <> ""
        If (GetAttr(DossierSource & NomFichier) And vbDirectory) = 0 Then
            nbFichiers = nbFichiers + 1
            ReDim Preserve fichiers(1 To nbFichiers)
            fichiers(nbFichiers) = NomFichier
        End If
        NomFichier = Dir
    Loop

    ' Étape 2 : Déplacer les fichiers collectés
    Dim i As Long
    For i = 1 To nbFichiers
        Extension = UCase(fso.GetExtensionName(fichiers(i)))

        If Extension <> "" Then
            DossierExtension = DossierSource & Extension & "\"

            If Not fso.FolderExists(DossierExtension) Then
                MkDir DossierExtension
                Debug.Print "Dossier créé : " & DossierExtension
            End If

            fso.MoveFile DossierSource & fichiers(i), DossierExtension & fichiers(i)
            Debug.Print "Déplacé : " & fichiers(i) & " vers " & Extension
        End If
    Next i

    MsgBox "Organisation terminée !"
    Set fso = Nothing
End Sub
```

### Créer une structure de projet

```vba
Sub CreerStructureProjet()
    Dim ProjetNom As String
    Dim CheminBase As String
    Dim Structure As Variant
    Dim i As Integer

    ProjetNom = InputBox("Nom du projet :", "Nouveau Projet", "MonProjet")

    If ProjetNom = "" Then Exit Sub

    CheminBase = "C:\Projets\" & ProjetNom & "\"

    ' Définir la structure des dossiers
    Structure = Array( _
        "01_Documentation", _
        "02_Code_Source", _
        "03_Tests", _
        "04_Ressources\Images", _
        "04_Ressources\Données", _
        "05_Livraison", _
        "06_Archive" _
    )

    ' Créer le dossier principal
    If Not DossierExiste(CheminBase) Then
        MkDir CheminBase
    End If

    ' Créer chaque dossier de la structure
    For i = 0 To UBound(Structure)
        Dim CheminDossier As String
        CheminDossier = CheminBase & Structure(i)

        ' Gérer les sous-dossiers (avec \)
        If InStr(CheminDossier, "\") > InStr(CheminDossier, CheminBase) + Len(CheminBase) Then
            ' Créer le dossier parent d'abord
            Dim DossierParent As String
            DossierParent = Left(CheminDossier, InStrRev(CheminDossier, "\") - 1)

            If Not DossierExiste(DossierParent) Then
                MkDir DossierParent
            End If
        End If

        If Not DossierExiste(CheminDossier) Then
            MkDir CheminDossier
            Debug.Print "Créé : " & CheminDossier
        End If
    Next i

    MsgBox "Structure de projet créée dans :" & vbCrLf & CheminBase
End Sub
```

## Bonnes pratiques

### 1. Toujours vérifier l'existence
```vba
' Avant de créer
If Not DossierExiste(CheminDossier) Then
    MkDir CheminDossier
End If

' Avant de supprimer
If DossierExiste(CheminDossier) Then
    RmDir CheminDossier
End If
```

### 2. Gérer les erreurs
```vba
On Error GoTo GestionErreur
' Opérations sur dossiers
Exit Sub

GestionErreur:  
MsgBox "Erreur : " & Err.Description  
```

### 3. Utiliser des variables pour les chemins
```vba
Dim DossierTravail As String  
DossierTravail = "C:\MonApplication\Données\"  

' Utiliser la variable partout
If DossierExiste(DossierTravail) Then...
```

### 4. Nettoyer les objets
```vba
Set fso = Nothing  
Set dossier = Nothing  
```

## Erreurs courantes et solutions

### "Path not found"
**Cause :** Le chemin parent n'existe pas  
**Solution :** Créer d'abord les dossiers parents  

### "Permission denied"
**Cause :** Droits insuffisants ou dossier en cours d'utilisation  
**Solution :** Vérifier les droits et fermer les applications  

### "Directory not empty"
**Cause :** Tentative de suppression d'un dossier contenant des fichiers avec RmDir  
**Solution :** Utiliser FileSystemObject.DeleteFolder ou vider d'abord le dossier  

---

*Dans la section suivante, nous découvrirons les fonctions système de VBA pour des opérations plus avancées sur les fichiers et dossiers.*

⏭️ [Dir, Kill, MkDir, RmDir](/11-fichiers-et-dossiers/04-dir-kill-mkdir-rmdir.md)
