🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.1. Points d'arrêt (Breakpoints)

## Qu'est-ce qu'un point d'arrêt ?

Un point d'arrêt (breakpoint en anglais) est un outil de débogage qui permet d'**interrompre temporairement l'exécution** de votre code VBA à une ligne spécifique. Imaginez que votre code soit comme un film : un point d'arrêt vous permet de mettre le film en pause à un moment précis pour examiner ce qui se passe.

Quand le programme atteint un point d'arrêt, il s'arrête **avant** d'exécuter cette ligne, vous donnant l'opportunité d'examiner l'état de vos variables, de vérifier les valeurs et de comprendre le comportement de votre programme.

## Pourquoi utiliser des points d'arrêt ?

Les points d'arrêt sont particulièrement utiles dans plusieurs situations :

**Vérifier les valeurs des variables** : Vous pouvez voir exactement quelle valeur contient une variable à un moment donné de l'exécution.

**Comprendre le flux d'exécution** : Vous pouvez suivre le chemin que prend votre code, surtout dans des structures complexes avec des conditions et des boucles.

**Identifier les erreurs logiques** : Quand votre code s'exécute sans erreur mais ne produit pas le résultat attendu, les points d'arrêt vous aident à localiser le problème.

**Analyser le comportement step-by-step** : Au lieu de deviner ce qui se passe, vous pouvez observer le comportement réel de votre code.

## Comment placer un point d'arrêt

### Méthode 1 : Clic dans la marge
La méthode la plus simple consiste à **cliquer dans la marge grise** à gauche de la ligne de code où vous voulez placer le point d'arrêt. Vous verrez apparaître un **cercle rouge** et la ligne sera surlignée en rouge.

### Méthode 2 : Touche F9
1. Placez votre curseur sur la ligne où vous voulez créer le point d'arrêt
2. Appuyez sur la touche **F9**
3. Le point d'arrêt apparaît (cercle rouge + surlignage)

### Méthode 3 : Menu Débogage
1. Placez votre curseur sur la ligne souhaitée
2. Allez dans le menu **Débogage**
3. Cliquez sur **Basculer le point d'arrêt**

## À quoi ressemble un point d'arrêt

Quand vous placez un point d'arrêt, vous voyez :
- Un **cercle rouge** dans la marge gauche de l'éditeur
- La **ligne de code surlignée en rouge** (couleur de fond rouge)

Exemple visuel :
```vba
Sub MonProgramme()
    Dim nombre As Integer
●   nombre = 10              ' ← Point d'arrêt ici (cercle rouge)
    nombre = nombre * 2
    MsgBox nombre
End Sub
```

## Comment supprimer un point d'arrêt

### Supprimer un point d'arrêt spécifique
- **Cliquez à nouveau** sur le cercle rouge dans la marge, ou
- Placez le curseur sur la ligne et appuyez sur **F9**, ou
- Menu **Débogage** > **Basculer le point d'arrêt**

### Supprimer tous les points d'arrêt
- Appuyez sur **Ctrl + Maj + F9**, ou
- Menu **Débogage** > **Effacer tous les points d'arrêt**

## Que se passe-t-il quand le code atteint un point d'arrêt ?

Lorsque votre programme s'exécute et rencontre un point d'arrêt :

1. **L'exécution s'interrompt** immédiatement
2. **L'éditeur VBA devient actif** et se met au premier plan
3. La **ligne avec le point d'arrêt est surlignée en jaune** (au lieu du rouge initial)
4. Une **flèche jaune** apparaît dans la marge, indiquant la prochaine ligne à exécuter
5. Vous pouvez maintenant **examiner vos variables** et l'état du programme

## Examiner les variables quand le programme est en pause

Une fois que votre programme est en pause sur un point d'arrêt, vous avez plusieurs moyens d'examiner vos variables :

### Méthode 1 : Survol avec la souris
Passez simplement votre souris au-dessus du nom d'une variable dans le code. Une petite bulle d'information apparaîtra avec la valeur actuelle de cette variable.

### Méthode 2 : Fenêtre d'exécution immédiate
1. Ouvrez la fenêtre d'exécution immédiate (Ctrl + G ou Affichage > Fenêtre d'exécution immédiate)
2. Tapez `? nomDeLaVariable` et appuyez sur Entrée
3. La valeur s'affiche dans la fenêtre

Exemple :
```
? nombre
10
```

### Méthode 3 : Fenêtre des variables locales
Allez dans **Affichage** > **Fenêtre des variables locales** pour voir toutes les variables de la procédure actuelle avec leurs valeurs.

## Reprendre l'exécution après un point d'arrêt

Une fois que vous avez examiné ce que vous vouliez, vous pouvez reprendre l'exécution :

- **F5** ou **Continuer** : Reprend l'exécution normale jusqu'au prochain point d'arrêt
- **F8** ou **Pas à pas détaillé** (Step Into) : Exécute la ligne suivante puis s'arrête à nouveau (entre dans les sous-procédures)
- **Maj + F8** ou **Pas à pas principal** (Step Over) : Exécute la ligne suivante sans entrer dans les sous-procédures
- **Ctrl + Maj + F8** ou **Pas à pas sortant** (Step Out) : Exécute le reste de la procédure actuelle et s'arrête au retour
- **Ctrl + F8** ou **Exécuter jusqu'au curseur** : Exécute le code jusqu'à la ligne où se trouve le curseur

## Bonnes pratiques avec les points d'arrêt

### Placement stratégique
- Placez les points d'arrêt aux **endroits clés** de votre code : début de procédures, avant et après les calculs importants, dans les conditions
- **Évitez de surcharger** votre code avec trop de points d'arrêt

### Points d'arrêt temporaires
- Les points d'arrêt sont **temporaires** par nature - ils disparaissent quand vous fermez le classeur
- Utilisez-les pendant le développement et retirez-les avant de distribuer votre code

### Documentation
- Si vous devez garder des points d'arrêt pour plus tard, **notez leur emplacement** car ils ne sont pas sauvegardés avec le fichier

## Points d'arrêt et boucles

Les points d'arrêt sont particulièrement utiles dans les boucles. Si vous placez un point d'arrêt à l'intérieur d'une boucle, le programme s'arrêtera **à chaque itération**.

```vba
Sub ExempleBoucle()
    Dim i As Integer

    For i = 1 To 5
●       Debug.Print i        ' ← Point d'arrêt ici
    Next i
End Sub
```

Dans cet exemple, le programme s'arrêtera 5 fois, une fois pour chaque valeur de `i`. Cela vous permet d'observer comment la variable `i` évolue à chaque itération.

## Limitations des points d'arrêt

**Lignes non exécutables** : Vous ne pouvez pas placer de points d'arrêt sur certaines lignes comme les déclarations de variables, les commentaires, ou les lignes vides.

**Performances** : Trop de points d'arrêt peuvent ralentir l'exécution, même s'ils ne sont pas atteints.

**Sauvegarde** : Les points d'arrêt ne sont pas sauvegardés avec votre fichier Excel - ils disparaissent à la fermeture.

## Cas d'usage typiques

**Vérification de calculs** : Placez un point d'arrêt après un calcul complexe pour vérifier que le résultat est correct.

**Débogage de conditions** : Placez des points d'arrêt avant et dans des structures `If...Then` pour voir quel chemin prend votre code.

**Analyse de boucles** : Utilisez les points d'arrêt pour comprendre comment vos boucles se comportent et si elles se terminent correctement.

**Vérification de données** : Arrêtez l'exécution pour examiner le contenu de vos variables ou l'état de vos feuilles Excel.

Les points d'arrêt sont votre premier outil de débogage et l'un des plus puissants. Maîtriser leur utilisation vous fera gagner énormément de temps dans l'identification et la résolution des problèmes dans votre code VBA.

⏭️
