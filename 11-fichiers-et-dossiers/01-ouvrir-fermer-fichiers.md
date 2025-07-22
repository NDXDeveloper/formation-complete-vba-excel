🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.1 Ouvrir et fermer des fichiers

## Introduction

En VBA, travailler avec des fichiers nécessite de suivre un processus en trois étapes : **ouvrir** le fichier, **traiter** son contenu, puis **fermer** le fichier. Cette approche garantit une gestion propre des ressources système et évite les conflits d'accès.

## Comprendre les modes d'ouverture

Avant d'ouvrir un fichier, il faut définir comment nous souhaitons l'utiliser. VBA propose plusieurs modes :

### Modes de base
- **Input** : Pour lire un fichier existant uniquement
- **Output** : Pour créer un nouveau fichier ou écraser un fichier existant
- **Append** : Pour ajouter du contenu à la fin d'un fichier existant

### Modes avancés
- **Binary** : Pour lire des fichiers non-textuels (images, exécutables)
- **Random** : Pour accéder à des positions spécifiques dans le fichier

## La fonction Open

### Syntaxe de base
```vba
Open "CheminDuFichier" For Mode As #NuméroFichier
```

### Les composants expliqués

**CheminDuFichier** : Le chemin complet ou relatif vers le fichier
```vba
' Chemin absolu
"C:\Mes Documents\Données.txt"

' Chemin relatif
".\Fichiers\Données.txt"
```

**Mode** : Comment nous voulons utiliser le fichier
```vba
' Pour lire seulement
For Input

' Pour écrire (nouveau fichier ou remplacer)
For Output

' Pour ajouter à la fin
For Append
```

**NuméroFichier** : Un numéro unique pour identifier le fichier ouvert
```vba
' Utiliser un numéro libre
As #1

' Ou obtenir automatiquement un numéro libre
As #FreeFile
```

## Exemples concrets d'ouverture

### Ouvrir un fichier pour lecture
```vba
Sub OuvrirPourLecture()
    Dim NumFichier As Integer

    ' Obtenir un numéro de fichier libre
    NumFichier = FreeFile

    ' Ouvrir le fichier en lecture
    Open "C:\Temp\MonFichier.txt" For Input As #NumFichier

    ' Ici on traiterait le contenu du fichier
    ' (nous verrons cela dans la section suivante)

    ' Toujours fermer le fichier
    Close #NumFichier
End Sub
```

### Ouvrir un fichier pour écriture
```vba
Sub OuvrirPourEcriture()
    Dim NumFichier As Integer

    NumFichier = FreeFile

    ' Créer ou remplacer le fichier
    Open "C:\Temp\NouveauFichier.txt" For Output As #NumFichier

    ' Ici on écrirait dans le fichier

    Close #NumFichier
End Sub
```

### Ouvrir un fichier pour ajout
```vba
Sub OuvrirPourAjout()
    Dim NumFichier As Integer

    NumFichier = FreeFile

    ' Ajouter à la fin d'un fichier existant
    Open "C:\Temp\Journal.txt" For Append As #NumFichier

    ' Ici on ajouterait du contenu

    Close #NumFichier
End Sub
```

## La fonction FreeFile

### Pourquoi l'utiliser ?
La fonction `FreeFile` retourne automatiquement le prochain numéro de fichier disponible. C'est une bonne pratique car :

- **Évite les conflits** : Pas de risque d'utiliser un numéro déjà pris
- **Simplifie le code** : Plus besoin de suivre manuellement les numéros
- **Rend le code robuste** : Fonctionne même si d'autres fichiers sont ouverts

### Exemple d'utilisation
```vba
Sub ExempleFreeFile()
    Dim Fichier1 As Integer, Fichier2 As Integer

    ' Obtenir des numéros libres
    Fichier1 = FreeFile  ' Pourrait être 1
    Fichier2 = FreeFile  ' Pourrait être 2

    ' Ouvrir plusieurs fichiers
    Open "C:\Temp\Fichier1.txt" For Input As #Fichier1
    Open "C:\Temp\Fichier2.txt" For Input As #Fichier2

    ' Fermer les fichiers
    Close #Fichier1
    Close #Fichier2
End Sub
```

## Fermer les fichiers

### L'importance de fermer les fichiers
Fermer un fichier est **essentiel** pour :
- **Libérer les ressources** système
- **Permettre à d'autres programmes** d'accéder au fichier
- **Sauvegarder** définitivement les modifications
- **Éviter les erreurs** lors de prochaines ouvertures

### Méthodes pour fermer

**Fermer un fichier spécifique**
```vba
Close #NuméroFichier
```

**Fermer tous les fichiers ouverts**
```vba
Close  ' Sans numéro, ferme tous les fichiers
```

### Exemple avec gestion d'erreur
```vba
Sub ExempleAvecGestionErreur()
    Dim NumFichier As Integer

    On Error GoTo GestionErreur

    NumFichier = FreeFile
    Open "C:\Temp\MonFichier.txt" For Input As #NumFichier

    ' Traitement du fichier ici

    ' Fermeture normale
    Close #NumFichier
    Exit Sub

GestionErreur:
    ' En cas d'erreur, s'assurer que le fichier est fermé
    Close #NumFichier
    MsgBox "Erreur lors de l'ouverture du fichier : " & Err.Description
End Sub
```

## Vérification d'existence avant ouverture

Il est recommandé de vérifier qu'un fichier existe avant de l'ouvrir en lecture :

```vba
Sub OuvrirAvecVerification()
    Dim CheminFichier As String
    Dim NumFichier As Integer

    CheminFichier = "C:\Temp\MonFichier.txt"

    ' Vérifier si le fichier existe
    If Dir(CheminFichier) <> "" Then
        NumFichier = FreeFile
        Open CheminFichier For Input As #NumFichier

        ' Traitement du fichier
        MsgBox "Fichier ouvert avec succès !"

        Close #NumFichier
    Else
        MsgBox "Le fichier n'existe pas : " & CheminFichier
    End If
End Sub
```

## Bonnes pratiques

### 1. Toujours utiliser FreeFile
```vba
' Recommandé
NumFichier = FreeFile
Open CheminFichier For Input As #NumFichier

' À éviter
Open CheminFichier For Input As #1
```

### 2. Gérer les erreurs
```vba
On Error GoTo GestionErreur
' Code d'ouverture et traitement
Close #NumFichier
Exit Sub

GestionErreur:
Close #NumFichier  ' S'assurer de la fermeture
```

### 3. Fermer dans l'ordre inverse d'ouverture
```vba
' Si plusieurs fichiers sont ouverts
Open Fichier1 For Input As #Num1
Open Fichier2 For Input As #Num2

' Les fermer dans l'ordre inverse
Close #Num2
Close #Num1
```

### 4. Utiliser des variables pour les chemins
```vba
' Plus lisible et maintenable
Dim CheminFichier As String
CheminFichier = "C:\Temp\MonFichier.txt"
Open CheminFichier For Input As #NumFichier
```

## Points d'attention

### Erreurs courantes à éviter
- **Oublier de fermer** : Peut bloquer l'accès au fichier
- **Ouvrir un fichier inexistant** en lecture : Génère une erreur
- **Utiliser le même numéro** pour plusieurs fichiers : Conflit garanti
- **Ne pas gérer les erreurs** : Le programme peut planter

### Messages d'erreur typiques
- **"File not found"** : Le fichier n'existe pas (mode Input)
- **"Permission denied"** : Fichier ouvert dans un autre programme
- **"Bad file name or number"** : Numéro de fichier incorrect
- **"Path/File access error"** : Chemin invalide ou droits insuffisants

---

*Dans la prochaine section, nous verrons comment lire et écrire effectivement le contenu des fichiers une fois qu'ils sont ouverts.*

⏭️
