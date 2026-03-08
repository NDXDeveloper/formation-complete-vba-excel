🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.1. Modèle objet Excel

## Comprendre le modèle objet Excel

Le **modèle objet Excel** est la structure organisée qui représente tous les éléments disponibles dans Excel. Imaginez-le comme un plan architectural qui décrit comment tous les composants d'Excel sont reliés entre eux. Pour un débutant, c'est comme apprendre l'organisation d'une grande bibliothèque : il faut d'abord comprendre comment les livres sont classés avant de pouvoir trouver celui que l'on cherche.

## La hiérarchie complète du modèle objet

Voici la structure hiérarchique complète du modèle objet Excel, présentée de manière simple :

```
Application (Excel lui-même)
│
├── Workbooks (Tous les classeurs ouverts)
│   └── Workbook (Un classeur spécifique)
│       │
│       ├── Worksheets (Toutes les feuilles du classeur)
│       │   └── Worksheet (Une feuille spécifique)
│       │       │
│       │       ├── Range (Une plage de cellules)
│       │       │   └── Cell (Une cellule individuelle)
│       │       │
│       │       ├── Shapes (Formes et objets graphiques)
│       │       ├── Charts (Graphiques)
│       │       └── PivotTables (Tableaux croisés dynamiques)
│       │
│       ├── Names (Noms définis dans le classeur)
│       ├── Styles (Styles du classeur)
│       └── VBProject (Le projet VBA du classeur)
│
├── Windows (Fenêtres d'Excel)
├── AddIns (Compléments installés)
└── CommandBars (Barres d'outils et menus)
```

## Les objets fondamentaux expliqués

### 1. Application (L'objet racine)

L'objet **Application** représente Excel lui-même. C'est le point de départ de toute interaction VBA avec Excel.

**Caractéristiques importantes :**
- Il est toujours accessible et unique
- Il contrôle les paramètres globaux d'Excel
- Il permet d'accéder à tous les autres objets

**Exemples de propriétés courantes :**
```vba
Application.Version          ' Version d'Excel (ex: "16.0")  
Application.UserName         ' Nom de l'utilisateur  
Application.ScreenUpdating   ' Active/désactive la mise à jour d'écran  
Application.Calculation      ' Mode de calcul (automatique/manuel)  
```

**Exemples de méthodes courantes :**
```vba
Application.Quit             ' Fermer Excel  
Application.Calculate        ' Recalculer toutes les feuilles ouvertes  
Application.Wait             ' Faire une pause dans l'exécution  
```

### 2. Workbooks (Collection des classeurs)

**Workbooks** est une collection qui contient tous les classeurs actuellement ouverts dans Excel.

**Utilisation typique :**
```vba
Workbooks.Count              ' Nombre de classeurs ouverts  
Workbooks("MonClasseur.xlsx") ' Accéder à un classeur par son nom  
Workbooks(1)                 ' Accéder au premier classeur ouvert  
```

**Méthodes importantes :**
```vba
Workbooks.Open("C:\MesDocuments\Fichier.xlsx")  ' Ouvrir un classeur  
Workbooks.Add                                   ' Créer un nouveau classeur  
```

### 3. Workbook (Un classeur individuel)

Un objet **Workbook** représente un fichier Excel spécifique.

**Propriétés essentielles :**
```vba
ActiveWorkbook.Name          ' Nom du classeur actif  
ActiveWorkbook.Path          ' Chemin du dossier contenant le classeur  
ActiveWorkbook.FullName      ' Chemin complet + nom du classeur  
ActiveWorkbook.Saved         ' True si le classeur est sauvegardé  
```

**Méthodes courantes :**
```vba
ActiveWorkbook.Save          ' Sauvegarder le classeur  
ActiveWorkbook.Close         ' Fermer le classeur  
ActiveWorkbook.SaveAs("C:\NouveauNom.xlsx")  ' Sauvegarder sous un autre nom  
```

### 4. Worksheets (Collection des feuilles)

**Worksheets** contient toutes les feuilles de calcul d'un classeur donné.

**Accès aux feuilles :**
```vba
Worksheets.Count             ' Nombre de feuilles dans le classeur  
Worksheets("Feuil1")         ' Accéder à une feuille par son nom  
Worksheets(1)                ' Accéder à la première feuille  
ActiveSheet                  ' La feuille actuellement active  
```

### 5. Worksheet (Une feuille individuelle)

Un objet **Worksheet** représente une feuille de calcul spécifique.

**Propriétés importantes :**
```vba
ActiveSheet.Name             ' Nom de la feuille active  
ActiveSheet.Visible          ' Visibilité de la feuille  
ActiveSheet.UsedRange        ' Plage de cellules utilisées  
```

**Méthodes utiles :**
```vba
ActiveSheet.Activate         ' Activer la feuille  
ActiveSheet.Copy             ' Copier la feuille  
ActiveSheet.Delete           ' Supprimer la feuille  
ActiveSheet.Protect          ' Protéger la feuille  
```

## Relations entre les objets

### Le principe de navigation hiérarchique

Pour accéder à un objet spécifique, vous devez généralement "descendre" dans la hiérarchie :

```vba
' Méthode complète (explicite)
Application.Workbooks("MonClasseur.xlsx").Worksheets("Feuil1").Range("A1")

' Méthode simplifiée (utilise les objets actifs)
Range("A1")  ' Si vous travaillez sur la feuille active du classeur actif
```

### Les objets "actifs" (raccourcis utiles)

Excel propose des raccourcis pour accéder aux objets actuellement sélectionnés :

- **ActiveWorkbook** : Le classeur actuellement actif
- **ActiveSheet** : La feuille actuellement active
- **ActiveCell** : La cellule actuellement sélectionnée
- **Selection** : Ce qui est actuellement sélectionné

**Exemple pratique :**
```vba
' Au lieu d'écrire :
Application.ActiveWorkbook.ActiveSheet.Range("A1").Value = "Bonjour"

' Vous pouvez écrire :
ActiveSheet.Range("A1").Value = "Bonjour"

' Ou même simplement :
Range("A1").Value = "Bonjour"
```

## Collections vs Objets individuels

### Qu'est-ce qu'une collection ?

Une **collection** est un groupe d'objets du même type. Dans Excel, les collections portent généralement un nom au pluriel :

- **Workbooks** (collection) contient des objets **Workbook**
- **Worksheets** (collection) contient des objets **Worksheet**
- **Cells** (collection) contient des objets **Range** (représentant des cellules)

### Accéder aux éléments d'une collection

Il existe plusieurs façons d'accéder aux éléments d'une collection :

```vba
' Par index numérique (commence à 1)
Worksheets(1)                ' Première feuille  
Worksheets(2)                ' Deuxième feuille  

' Par nom (plus lisible et stable)
Worksheets("Données")        ' Feuille nommée "Données"  
Worksheets("Résultats")      ' Feuille nommée "Résultats"  

' Nombre d'éléments dans la collection
Worksheets.Count             ' Nombre de feuilles
```

## Propriétés et méthodes : la différence

### Les propriétés (caractéristiques)

Les **propriétés** sont les caractéristiques d'un objet. Elles peuvent généralement être lues et modifiées :

```vba
' Lire une propriété
monNom = ActiveSheet.Name

' Modifier une propriété
ActiveSheet.Name = "NouvelleFeuille"
```

### Les méthodes (actions)

Les **méthodes** sont les actions qu'un objet peut effectuer :

```vba
' Méthodes sans paramètres
ActiveWorkbook.Save          ' Sauvegarder  
ActiveSheet.Calculate        ' Recalculer  

' Méthodes avec paramètres
ActiveSheet.Copy After:=Worksheets(2)  ' Copier après la 2ème feuille
```

## Conseils pour débuter avec le modèle objet

### 1. Utilisez l'aide contextuelle

Dans l'éditeur VBA, tapez le nom d'un objet suivi d'un point, et Excel affichera automatiquement la liste des propriétés et méthodes disponibles.

### 2. Commencez simple

Débutez avec les objets de base (Application, ActiveWorkbook, ActiveSheet, Range) avant de vous aventurer vers des objets plus complexes.

### 3. La logique avant la syntaxe

Réfléchissez d'abord à ce que vous voulez faire en termes d'Excel normal, puis traduisez en objets VBA :
- "Je veux modifier la cellule A1" → `Range("A1").Value = "ma valeur"`
- "Je veux renommer la feuille" → `ActiveSheet.Name = "nouveau nom"`

### 4. Testez progressivement

N'hésitez pas à tester chaque ligne de code séparément pour comprendre son effet avant de construire des programmes plus complexes.

## Récapitulatif des concepts clés

- Le **modèle objet Excel** organise tous les éléments d'Excel selon une hiérarchie logique
- **Application** est l'objet racine qui représente Excel lui-même
- Les **collections** (Workbooks, Worksheets) regroupent des objets similaires
- Les **propriétés** sont les caractéristiques des objets (lecture/écriture)
- Les **méthodes** sont les actions que peuvent effectuer les objets
- Les objets "actifs" (ActiveWorkbook, ActiveSheet) sont des raccourcis pratiques
- La navigation se fait en descendant la hiérarchie : Application → Workbook → Worksheet → Range

Maîtriser ce modèle objet est essentiel car il constitue la fondation de toute programmation VBA efficace dans Excel. Dans les sections suivantes, nous approfondirons chacun de ces objets avec des exemples pratiques concrets.

⏭️
