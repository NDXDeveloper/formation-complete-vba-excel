🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.3. Fenêtre de surveillance (Watch)

## Qu'est-ce que la fenêtre de surveillance ?

La fenêtre de surveillance (Watch Window en anglais) est un outil de débogage qui vous permet de **surveiller en permanence** la valeur de variables ou d'expressions spécifiques pendant l'exécution de votre code VBA. C'est comme avoir un tableau de bord qui affiche en temps réel les informations qui vous intéressent.

Imaginez que vous conduisez une voiture : vous avez besoin de surveiller votre vitesse, votre niveau d'essence, et la température du moteur. La fenêtre de surveillance fonctionne de la même manière pour votre code - elle vous montre les "instruments" que vous voulez surveiller pendant que votre programme "roule".

## Pourquoi utiliser la fenêtre de surveillance ?

**Surveillance continue** : Contrairement au survol de souris qui ne fonctionne que pendant les pauses, la fenêtre de surveillance affiche constamment les valeurs, même pendant l'exécution normale.

**Vue d'ensemble** : Vous pouvez surveiller plusieurs variables en même temps dans une seule fenêtre, au lieu de les vérifier une par une.

**Expressions complexes** : Vous pouvez surveiller non seulement des variables simples, mais aussi des calculs, des propriétés d'objets, ou des expressions complexes.

**Détection de changements** : Vous voyez immédiatement quand et comment vos variables changent de valeur.

**Gain de temps** : Plus besoin d'utiliser Debug.Print ou MsgBox pour vérifier des valeurs spécifiques.

## Comment ouvrir la fenêtre de surveillance

### Méthode 1 : Via le menu
1. Dans l'éditeur VBA, allez dans le menu **Affichage**
2. Cliquez sur **Fenêtre Espion** (ou **Watch Window**)
3. La fenêtre s'ouvre généralement en bas de l'écran

### Méthode 2 : Via la barre d'outils
1. Assurez-vous que la barre d'outils **Débogage** est visible
2. Cliquez sur l'icône de la fenêtre de surveillance

### Méthode 3 : Raccourci pendant le débogage
Quand vous êtes en mode débogage (programme en pause), la fenêtre de surveillance est souvent accessible plus facilement via les menus contextuels.

## Comprendre l'interface de la fenêtre de surveillance

La fenêtre de surveillance se présente comme un tableau avec plusieurs colonnes :

**Expression** : Le nom de la variable ou l'expression que vous surveillez
**Valeur** : La valeur actuelle de cette expression
**Type** : Le type de données (Integer, String, Object, etc.)
**Contexte** : La procédure et le module où l'expression est définie

## Comment ajouter une expression à surveiller

### Méthode 1 : Glisser-déposer depuis le code
1. **Sélectionnez** le nom d'une variable dans votre code
2. **Faites-la glisser** vers la fenêtre de surveillance
3. L'expression apparaît automatiquement dans la liste

### Méthode 2 : Clic droit sur une variable
1. **Clic droit** sur une variable dans votre code
2. Sélectionnez **Ajouter un espion** dans le menu contextuel
3. Une boîte de dialogue s'ouvre pour configurer l'espion

### Méthode 3 : Via le menu Débogage
1. Placez votre curseur sur la variable que vous voulez surveiller
2. Allez dans **Débogage** > **Ajouter un espion**
3. Configurez l'espion dans la boîte de dialogue

### Méthode 4 : Ajout manuel
1. **Clic droit** dans la fenêtre de surveillance
2. Sélectionnez **Ajouter un espion**
3. Tapez manuellement l'expression à surveiller

## Configurer un espion - La boîte de dialogue

Quand vous ajoutez un espion via le menu ou le clic droit, une boîte de dialogue s'ouvre avec plusieurs options :

### Expression
C'est ce que vous voulez surveiller. Cela peut être :
- **Une variable simple** : `monNombre`
- **Une propriété d'objet** : `Worksheets("Feuil1").Range("A1").Value`
- **Une expression calculée** : `nombre1 + nombre2`
- **Un élément de tableau** : `monTableau(2)`

### Contexte
**Procédure** : Spécifie dans quelle procédure cette variable existe
**Module** : Spécifie dans quel module chercher la variable

### Type d'espion
**Expression espion** : Surveille simplement la valeur (le plus courant)
**Arrêt si la valeur est True** : Arrête l'exécution quand l'expression devient vraie
**Arrêt si la valeur change** : Arrête l'exécution quand la valeur change

## Types d'expressions que vous pouvez surveiller

### Variables simples
```vba
Dim nombre As Integer
nombre = 42
' Surveillez : nombre
```

### Propriétés d'objets Excel
```vba
' Surveillez : Range("A1").Value
' Surveillez : ActiveSheet.Name
' Surveillez : Workbooks.Count
```

### Expressions calculées
```vba
Dim a As Integer, b As Integer
a = 10
b = 20
' Surveillez : a + b
' Surveillez : a * b / 2
```

### Éléments de tableaux
```vba
Dim notes(1 To 5) As Integer
notes(1) = 15
' Surveillez : notes(1)
' Surveillez : UBound(notes)
```

### Expressions conditionnelles
```vba
Dim age As Integer
age = 25
' Surveillez : age >= 18
' Surveillez : age > 65
```

## Comprendre les valeurs affichées

### Valeurs normales
Les valeurs s'affichent normalement : `42`, `"Bonjour"`, `True`, etc.

### Valeurs spéciales
**<Hors de portée>** : La variable n'existe pas dans le contexte actuel
**<Non défini>** : La variable n'a pas encore été initialisée
**<Erreur d'objet>** : Problème avec une référence d'objet
**<Expression non valide>** : L'expression contient une erreur de syntaxe

### Objets complexes
Pour les objets, vous voyez souvent le type d'objet entre crochets : `[Worksheet]`, `[Range]`, etc.

## Utiliser les espions pendant le débogage

### Surveillance en temps réel
Pendant l'exécution pas à pas (F8), les valeurs dans la fenêtre de surveillance se mettent à jour automatiquement à chaque étape.

### Détection de changements
Quand une valeur change, elle peut être mise en évidence (selon la version de VBA) pour attirer votre attention.

### Exemple pratique
```vba
Sub ExempleSurveillance()
    Dim compteur As Integer
    Dim somme As Integer

    ' Ajoutez 'compteur' et 'somme' à la surveillance

    For compteur = 1 To 5
        somme = somme + compteur
        ' Regardez les valeurs changer dans la fenêtre
    Next compteur

    MsgBox "Somme finale : " & somme
End Sub
```

## Espions conditionnels - Arrêt automatique

### Arrêt si la valeur est True
Particulièrement utile pour surveiller des conditions :
- Expression : `nombre > 100`
- Le programme s'arrête automatiquement quand `nombre` dépasse 100

### Arrêt si la valeur change
Utile pour détecter quand une variable importante est modifiée :
- Expression : `statusImportant`
- Le programme s'arrête dès que cette variable change

### Exemple d'utilisation
```vba
Sub ExempleArretConditionnel()
    Dim valeur As Integer

    ' Créez un espion : valeur > 50 avec "Arrêt si True"

    For valeur = 1 To 100
        ' Le programme s'arrêtera automatiquement quand valeur atteint 51
        Debug.Print valeur
    Next valeur
End Sub
```

## Gérer les espions

### Modifier un espion
1. **Double-cliquez** sur l'espion dans la fenêtre de surveillance
2. La boîte de dialogue de configuration s'ouvre
3. Modifiez l'expression ou les options
4. Cliquez **OK**

### Supprimer un espion
1. **Sélectionnez** l'espion dans la fenêtre
2. Appuyez sur la touche **Suppr**, ou
3. **Clic droit** > **Supprimer l'espion**

### Supprimer tous les espions
1. **Clic droit** dans la fenêtre de surveillance
2. Sélectionnez **Supprimer tous les espions**

## Conseils pratiques pour utiliser la fenêtre de surveillance

### Surveillez les variables clés
Ne surchargez pas la fenêtre. Concentrez-vous sur les 3-5 variables les plus importantes pour votre débogage.

### Utilisez des expressions parlantes
Au lieu de surveiller `x`, surveillez plutôt `prixTotalHT` ou `nombreClients` - des noms qui ont du sens.

### Surveillez les propriétés d'objets
```vba
' Très utile pour surveiller :
Range("A1").Value
ActiveSheet.Name
Workbooks.Count
```

### Surveillez les calculs
```vba
' Au lieu de calculer mentalement, surveillez :
prixHT * tauxTVA
nombreHeures * tarifHoraire
```

### Utilisez les espions conditionnels avec parcimonie
Ils sont puissants mais peuvent ralentir l'exécution. Utilisez-les seulement quand nécessaire.

## Limitations et précautions

### Performances
Trop d'espions peuvent ralentir l'exécution de votre code, surtout avec des expressions complexes.

### Contexte
Les espions ne fonctionnent que quand les variables sont dans le bon contexte (bonne procédure, bon module).

### Expressions complexes
Des expressions très complexes peuvent parfois provoquer des erreurs ou donner des résultats inattendus.

### Sauvegarde
Comme les points d'arrêt, les espions ne sont pas sauvegardés avec votre fichier Excel.

## Cas d'usage typiques

**Débogage de boucles complexes** : Surveillez les variables de contrôle et les calculs à l'intérieur des boucles.

**Suivi d'états d'objets** : Surveillez les propriétés d'objets Excel qui changent pendant l'exécution.

**Validation de calculs** : Surveillez les résultats intermédiaires de calculs complexes.

**Détection de conditions spéciales** : Utilisez les espions conditionnels pour arrêter automatiquement dans des situations particulières.

**Analyse de flux de données** : Suivez comment les données se transforment à travers votre programme.

La fenêtre de surveillance est un outil puissant qui complète parfaitement les points d'arrêt et l'exécution pas à pas. Elle vous donne une vision continue et détaillée de l'état de votre programme, rendant le débogage plus efficace et plus intuitif.

⏭️
