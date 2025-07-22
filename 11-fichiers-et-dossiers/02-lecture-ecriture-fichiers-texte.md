🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.2 Lecture et écriture de fichiers texte

## Introduction

Maintenant que nous savons ouvrir et fermer des fichiers, nous allons apprendre à lire leur contenu et à y écrire des données. La manipulation de fichiers texte est l'une des tâches les plus courantes en VBA, que ce soit pour traiter des données, créer des rapports ou échanger des informations avec d'autres systèmes.

## Lecture de fichiers texte

### Les fonctions de lecture

VBA propose plusieurs fonctions pour lire le contenu d'un fichier :

- **Line Input #** : Lit une ligne complète
- **Input #** : Lit un nombre spécifique de caractères
- **Input()** : Lit tout le contenu du fichier d'un coup

### Lire ligne par ligne avec Line Input

C'est la méthode la plus courante et la plus sûre pour lire un fichier texte :

```vba
Sub LireFichierLigneParLigne()
    Dim NumFichier As Integer
    Dim UneLigne As String
    Dim CheminFichier As String

    CheminFichier = "C:\Temp\MonFichier.txt"

    ' Vérifier que le fichier existe
    If Dir(CheminFichier) = "" Then
        MsgBox "Le fichier n'existe pas !"
        Exit Sub
    End If

    NumFichier = FreeFile
    Open CheminFichier For Input As #NumFichier

    ' Lire tant qu'on n'est pas à la fin du fichier
    Do Until EOF(NumFichier)
        Line Input #NumFichier, UneLigne

        ' Traiter la ligne (ici, l'afficher)
        Debug.Print UneLigne
    Loop

    Close #NumFichier
    MsgBox "Lecture terminée !"
End Sub
```

### Comprendre EOF (End of File)

La fonction `EOF()` retourne `True` quand on atteint la fin du fichier :

```vba
' Vérifier si on est à la fin du fichier
If EOF(NumFichier) Then
    MsgBox "Fin du fichier atteinte"
End If
```

### Lire tout le fichier d'un coup

Pour les petits fichiers, on peut tout lire en une seule opération :

```vba
Sub LireFichierComplet()
    Dim NumFichier As Integer
    Dim ContenuComplet As String
    Dim CheminFichier As String

    CheminFichier = "C:\Temp\PetitFichier.txt"

    If Dir(CheminFichier) = "" Then
        MsgBox "Fichier introuvable !"
        Exit Sub
    End If

    NumFichier = FreeFile
    Open CheminFichier For Input As #NumFichier

    ' Lire tout le contenu
    ContenuComplet = Input(LOF(NumFichier), NumFichier)

    Close #NumFichier

    ' Afficher le contenu
    MsgBox ContenuComplet
End Sub
```

### La fonction LOF (Length of File)

`LOF()` retourne la taille du fichier en caractères :

```vba
Dim TailleFichier As Long
TailleFichier = LOF(NumFichier)
Debug.Print "Le fichier fait " & TailleFichier & " caractères"
```

## Écriture de fichiers texte

### Les fonctions d'écriture

Pour écrire dans un fichier, VBA propose :

- **Print #** : Écrit du texte avec retour à la ligne automatique
- **Write #** : Écrit des données avec séparateurs automatiques
- **Put #** : Pour l'écriture binaire (moins courant pour le texte)

### Écrire avec Print #

C'est la méthode la plus simple pour écrire du texte :

```vba
Sub EcrireFichierTexte()
    Dim NumFichier As Integer
    Dim CheminFichier As String

    CheminFichier = "C:\Temp\NouveauFichier.txt"

    NumFichier = FreeFile
    Open CheminFichier For Output As #NumFichier

    ' Écrire plusieurs lignes
    Print #NumFichier, "Première ligne de texte"
    Print #NumFichier, "Deuxième ligne de texte"
    Print #NumFichier, "Date de création : " & Now

    Close #NumFichier
    MsgBox "Fichier créé avec succès !"
End Sub
```

### Écrire sans retour à la ligne

Pour écrire sur la même ligne, utiliser une virgule :

```vba
Sub EcrireSurMemeLigne()
    Dim NumFichier As Integer

    NumFichier = FreeFile
    Open "C:\Temp\SurUneLigne.txt" For Output As #NumFichier

    ' Ces éléments seront sur la même ligne
    Print #NumFichier, "Nom: "; "Jean"; " - Age: "; 25

    Close #NumFichier
End Sub
```

### Ajouter du contenu avec Append

Pour ajouter du contenu à la fin d'un fichier existant :

```vba
Sub AjouterContenu()
    Dim NumFichier As Integer
    Dim CheminFichier As String

    CheminFichier = "C:\Temp\Journal.txt"

    NumFichier = FreeFile
    Open CheminFichier For Append As #NumFichier

    ' Ajouter une nouvelle entrée
    Print #NumFichier, Date & " - " & Time & " : Nouvelle entrée"

    Close #NumFichier
    MsgBox "Entrée ajoutée au journal"
End Sub
```

## Exemples pratiques complets

### Créer un fichier de log

```vba
Sub CreerFichierLog()
    Dim NumFichier As Integer
    Dim CheminLog As String
    Dim i As Integer

    CheminLog = "C:\Temp\Journal_" & Format(Date, "yyyy-mm-dd") & ".txt"

    NumFichier = FreeFile
    Open CheminLog For Append As #NumFichier

    ' Ajouter une entrée avec horodatage
    Print #NumFichier, "=== SESSION DÉMARRÉE ==="
    Print #NumFichier, "Date: " & Date
    Print #NumFichier, "Heure: " & Time
    Print #NumFichier, "Utilisateur: " & Environ("USERNAME")
    Print #NumFichier, "=========================="
    Print #NumFichier, ""

    Close #NumFichier
    MsgBox "Fichier de log créé : " & CheminLog
End Sub
```

### Lire et analyser un fichier CSV simple

```vba
Sub LireFichierCSV()
    Dim NumFichier As Integer
    Dim UneLigne As String
    Dim CheminCSV As String
    Dim Compteur As Integer

    CheminCSV = "C:\Temp\Donnees.csv"

    If Dir(CheminCSV) = "" Then
        MsgBox "Fichier CSV introuvable !"
        Exit Sub
    End If

    NumFichier = FreeFile
    Open CheminCSV For Input As #NumFichier

    Compteur = 0

    Do Until EOF(NumFichier)
        Line Input #NumFichier, UneLigne
        Compteur = Compteur + 1

        ' Afficher les premières lignes
        If Compteur <= 5 Then
            Debug.Print "Ligne " & Compteur & ": " & UneLigne
        End If
    Loop

    Close #NumFichier
    MsgBox "Fichier CSV lu. Nombre de lignes : " & Compteur
End Sub
```

### Copier un fichier texte avec modifications

```vba
Sub CopierEtModifierFichier()
    Dim NumSource As Integer, NumDestination As Integer
    Dim UneLigne As String
    Dim CheminSource As String, CheminDestination As String

    CheminSource = "C:\Temp\Original.txt"
    CheminDestination = "C:\Temp\Copie_Modifiee.txt"

    ' Vérifier que le fichier source existe
    If Dir(CheminSource) = "" Then
        MsgBox "Fichier source introuvable !"
        Exit Sub
    End If

    ' Ouvrir les deux fichiers
    NumSource = FreeFile
    Open CheminSource For Input As #NumSource

    NumDestination = FreeFile
    Open CheminDestination For Output As #NumDestination

    ' Ajouter un en-tête
    Print #NumDestination, "=== COPIE CRÉÉE LE " & Date & " ==="
    Print #NumDestination, ""

    ' Copier ligne par ligne en numérotant
    Dim NumeroLigne As Integer
    NumeroLigne = 1

    Do Until EOF(NumSource)
        Line Input #NumSource, UneLigne
        Print #NumDestination, NumeroLigne & ": " & UneLigne
        NumeroLigne = NumeroLigne + 1
    Loop

    ' Fermer les fichiers
    Close #NumSource
    Close #NumDestination

    MsgBox "Copie terminée ! Fichier créé : " & CheminDestination
End Sub
```

## Gestion de l'encodage des caractères

### Problèmes courants avec les accents

Les fichiers texte peuvent avoir différents encodages. En VBA standard, l'encodage par défaut peut poser problème avec les caractères accentués.

### Solution avec un objet FileSystemObject

Pour une meilleure gestion de l'encodage, on peut utiliser l'objet FileSystemObject :

```vba
Sub LireAvecFileSystemObject()
    Dim fso As Object
    Dim fichier As Object
    Dim contenu As String

    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ouvrir le fichier
    Set fichier = fso.OpenTextFile("C:\Temp\AvecAccents.txt", 1) ' 1 = ForReading

    ' Lire tout le contenu
    contenu = fichier.ReadAll

    ' Fermer le fichier
    fichier.Close

    ' Afficher le contenu
    MsgBox contenu

    ' Nettoyer les objets
    Set fichier = Nothing
    Set fso = Nothing
End Sub
```

## Gestion des erreurs courantes

### Exemple avec gestion d'erreurs complète

```vba
Sub LectureAvecGestionErreurs()
    Dim NumFichier As Integer
    Dim UneLigne As String
    Dim CheminFichier As String

    On Error GoTo GestionErreur

    CheminFichier = "C:\Temp\MonFichier.txt"

    ' Vérification préalable
    If Dir(CheminFichier) = "" Then
        MsgBox "Le fichier " & CheminFichier & " n'existe pas."
        Exit Sub
    End If

    NumFichier = FreeFile
    Open CheminFichier For Input As #NumFichier

    Do Until EOF(NumFichier)
        Line Input #NumFichier, UneLigne
        Debug.Print UneLigne
    Loop

    Close #NumFichier
    MsgBox "Lecture réussie !"
    Exit Sub

GestionErreur:
    ' Fermer le fichier en cas d'erreur
    Close #NumFichier
    MsgBox "Erreur lors de la lecture : " & Err.Description
End Sub
```

## Bonnes pratiques pour les fichiers texte

### 1. Toujours vérifier l'existence en lecture
```vba
If Dir(CheminFichier) = "" Then
    MsgBox "Fichier introuvable !"
    Exit Sub
End If
```

### 2. Gérer les gros fichiers ligne par ligne
```vba
' Préférer ceci pour les gros fichiers
Do Until EOF(NumFichier)
    Line Input #NumFichier, UneLigne
    ' Traiter la ligne
Loop

' Éviter ceci pour les gros fichiers
ContenuComplet = Input(LOF(NumFichier), NumFichier)
```

### 3. Utiliser des chemins avec variables
```vba
Dim DossierTravail As String
DossierTravail = "C:\MonProjet\Données\"
CheminFichier = DossierTravail & "Fichier_" & Format(Date, "yyyy-mm-dd") & ".txt"
```

### 4. Ajouter des informations de contexte
```vba
Print #NumFichier, "Fichier créé le : " & Now
Print #NumFichier, "Par l'utilisateur : " & Environ("USERNAME")
Print #NumFichier, "Machine : " & Environ("COMPUTERNAME")
```

## Erreurs fréquentes et solutions

### "Permission denied"
**Cause :** Le fichier est ouvert dans un autre programme
**Solution :** Fermer l'autre programme ou utiliser un autre nom

### "File not found"
**Cause :** Le chemin est incorrect ou le fichier n'existe pas
**Solution :** Vérifier avec `Dir()` avant d'ouvrir

### "Bad file name or number"
**Cause :** Tentative d'utilisation d'un fichier non ouvert
**Solution :** S'assurer que `Open` a été appelé avec succès

### Caractères bizarres (accents)
**Cause :** Problème d'encodage
**Solution :** Utiliser FileSystemObject ou vérifier l'encodage du fichier source

---

*Dans la section suivante, nous apprendrons à manipuler les dossiers pour organiser nos fichiers de manière automatique.*

⏭️
