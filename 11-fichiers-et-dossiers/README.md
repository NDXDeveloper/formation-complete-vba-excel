🔝 Retour au [Sommaire](/SOMMAIRE.md)

# Chapitre 11 : Fichiers et Dossiers - Introduction

## Vue d'ensemble

La manipulation de fichiers et de dossiers est une compétence essentielle en VBA qui permet d'automatiser de nombreuses tâches répétitives. Ce chapitre vous apprendra à interagir avec le système de fichiers de Windows pour créer des applications VBA plus puissantes et autonomes.

## Pourquoi manipuler les fichiers et dossiers en VBA ?

### Automatisation des tâches courantes
- **Traitement par lots** : Traiter plusieurs fichiers Excel simultanément
- **Sauvegarde automatique** : Créer des copies de sécurité programmées
- **Organisation** : Ranger automatiquement les fichiers selon des critères
- **Import/Export** : Échanger des données avec d'autres systèmes

### Cas d'usage concrets
- Consolidation de rapports mensuels provenant de différents services
- Création automatique de structures de dossiers pour nouveaux projets
- Nettoyage et archivage de fichiers anciens
- Génération de rapports à partir de fichiers texte ou CSV

## Concepts fondamentaux

### Le système de fichiers Windows
VBA utilise les fonctions intégrées de Windows pour manipuler les fichiers et dossiers. Il est important de comprendre :

- **Chemins absolus** : `C:\Users\NomUtilisateur\Documents\MonFichier.xlsx`
- **Chemins relatifs** : `.\Données\Fichier.txt` (relatif au répertoire courant)
- **Séparateurs** : Windows utilise `\` comme séparateur de dossiers
- **Extensions** : `.xlsx`, `.txt`, `.csv`, etc. définissent le type de fichier

### Types de fichiers manipulables
VBA peut travailler avec tous types de fichiers :
- **Fichiers Office** : Excel (.xlsx, .xls), Word (.docx), PowerPoint (.pptx)
- **Fichiers texte** : .txt, .csv, .log
- **Fichiers de configuration** : .ini, .xml, .json
- **Tout autre format** : images, PDF, etc. (lecture en mode binaire)

## Sécurité et bonnes pratiques

### Précautions importantes
- **Vérification d'existence** : Toujours vérifier qu'un fichier existe avant de le manipuler
- **Gestion des erreurs** : Prévoir les cas d'erreur (fichier verrouillé, accès refusé)
- **Sauvegarde** : Créer des copies avant modification de fichiers importants
- **Chemins valides** : Valider les chemins pour éviter les caractères interdits

### Droits d'accès
- Certains dossiers système nécessitent des privilèges administrateur
- Les fichiers ouverts dans d'autres applications peuvent être verrouillés
- Les réseaux d'entreprise peuvent limiter l'accès à certains dossiers

## Structure du chapitre

Ce chapitre couvre les aspects suivants de la manipulation de fichiers et dossiers :

### 11.1 Ouvrir et fermer des fichiers
Techniques pour accéder aux fichiers en lecture et écriture, gestion des handles de fichiers.

### 11.2 Lecture et écriture de fichiers texte
Méthodes pour lire et écrire du contenu textuel, gestion de l'encodage des caractères.

### 11.3 Manipulation de dossiers
Création, suppression et navigation dans l'arborescence des dossiers.

### 11.4 Fonctions système (Dir, Kill, MkDir, RmDir)
Utilisation des fonctions VBA natives pour les opérations sur le système de fichiers.

### 11.5 FileDialog pour sélection de fichiers
Interface utilisateur pour permettre à l'utilisateur de choisir des fichiers et dossiers.

## Préparation

Avant de commencer, assurez-vous de :
- Avoir une copie de sauvegarde de vos fichiers importants
- Créer un dossier de test pour vos exercices
- Comprendre la structure de dossiers de votre ordinateur
- Avoir les droits d'écriture dans le répertoire de travail

## Exemple simple d'introduction

Voici un premier aperçu de ce que nous allons apprendre :

```vba
Sub ExempleIntroduction()
    ' Vérifier si un fichier existe
    If Dir("C:\Temp\MonFichier.txt") <> "" Then
        MsgBox "Le fichier existe !"
    Else
        MsgBox "Le fichier n'existe pas."
    End If
End Sub
```

Cet exemple simple illustre l'une des opérations les plus courantes : vérifier l'existence d'un fichier avant de le traiter.

---

*Dans les sections suivantes, nous explorerons en détail chaque aspect de la manipulation de fichiers et dossiers, avec de nombreux exemples pratiques et des exercices pour consolider vos connaissances.*

⏭️
