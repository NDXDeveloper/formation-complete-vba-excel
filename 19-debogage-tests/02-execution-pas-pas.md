🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.2. Exécution pas à pas

## Qu'est-ce que l'exécution pas à pas ?

L'exécution pas à pas est une technique de débogage qui vous permet d'**exécuter votre code ligne par ligne**, en contrôlant manuellement chaque étape. C'est comme si vous regardiez votre programme au ralenti, instruction par instruction, pour comprendre exactement ce qui se passe à chaque moment.

Imaginez que vous suivez une recette de cuisine : au lieu de faire toute la recette d'un coup, vous faites une étape, vous vérifiez le résultat, puis vous passez à l'étape suivante. L'exécution pas à pas fonctionne de la même manière avec votre code VBA.

## Pourquoi utiliser l'exécution pas à pas ?

**Comprendre le flux d'exécution** : Vous voyez exactement dans quel ordre vos lignes de code s'exécutent, ce qui est particulièrement utile avec des conditions et des boucles.

**Identifier les erreurs logiques** : Quand votre code ne fait pas ce que vous attendez, l'exécution pas à pas vous aide à voir précisément où les choses tournent mal.

**Apprendre comment fonctionne le code** : C'est un excellent moyen d'apprendre et de comprendre le code écrit par d'autres ou même votre propre code ancien.

**Vérifier les valeurs étape par étape** : Vous pouvez observer comment les variables changent au fur et à mesure de l'exécution.

## Les différents types d'exécution pas à pas

VBA propose plusieurs modes d'exécution pas à pas, chacun adapté à des situations différentes :

### 1. Pas à pas détaillé (Step Into) - Touche F8

C'est le mode le plus détaillé. Il exécute **une seule ligne à la fois** et entre dans **toutes** les procédures appelées.

**Quand l'utiliser** : Quand vous voulez voir chaque détail de l'exécution, y compris ce qui se passe dans les procédures que vous appelez.

**Comment procéder** :
1. Placez votre curseur au début de la procédure que vous voulez déboguer
2. Appuyez sur **F8**
3. Chaque pression sur F8 exécute la ligne suivante
4. Une **flèche jaune** indique la prochaine ligne à exécuter

### 2. Pas à pas principal (Step Over) - Touche Maj + F8

Ce mode exécute une ligne à la fois, mais **n'entre pas** dans les procédures appelées. Il exécute ces procédures en entier d'un coup.

**Quand l'utiliser** : Quand vous voulez rester concentré sur la procédure principale sans vous perdre dans les détails des sous-procédures.

**Exemple** :
```vba
Sub ProcedurePrincipale()
    Dim resultat As Integer
    resultat = 5
→   resultat = CalculerDouble(resultat)    ' ← Flèche jaune ici
    MsgBox resultat
End Sub

Function CalculerDouble(nombre As Integer) As Integer
    CalculerDouble = nombre * 2
End Function
```

Avec **Step Over**, la fonction `CalculerDouble` s'exécutera entièrement, et vous passerez directement à la ligne `MsgBox resultat`.

### 3. Pas à pas sortant (Step Out) - Touche Ctrl + Maj + F8

Ce mode **termine l'exécution** de la procédure actuelle et s'arrête dans la procédure qui l'a appelée.

**Quand l'utiliser** : Quand vous êtes "descendu trop profondément" dans les détails d'une procédure et que vous voulez remonter au niveau supérieur.

## Comment démarrer l'exécution pas à pas

### Méthode 1 : Depuis le début d'une procédure
1. Placez votre curseur n'importe où dans la procédure que vous voulez déboguer
2. Appuyez sur **F8**
3. L'exécution commence et s'arrête sur la première ligne exécutable

### Méthode 2 : Depuis un point d'arrêt
1. Placez un point d'arrêt là où vous voulez commencer
2. Exécutez votre code normalement (F5)
3. Quand le programme s'arrête au point d'arrêt, utilisez F8 pour continuer pas à pas

### Méthode 3 : Via le menu
- Menu **Débogage** > **Pas à pas détaillé** (F8)
- Menu **Débogage** > **Pas à pas principal** (Maj + F8)
- Menu **Débogage** > **Pas à pas sortant** (Ctrl + Maj + F8)

## Comprendre les indicateurs visuels

Pendant l'exécution pas à pas, vous verrez plusieurs indicateurs :

**Flèche jaune** : Indique la **prochaine ligne à exécuter**. Cette ligne n'a pas encore été exécutée.

**Surlignage jaune** : La ligne avec la flèche jaune est aussi surlignée en jaune pour la rendre plus visible.

**Changement de couleur** : Quand vous passez d'une procédure à une autre, l'indicateur se déplace pour suivre l'exécution.

## Exemple pratique d'exécution pas à pas

Prenons cet exemple simple :

```vba
Sub ExemplePasAPas()
    Dim nombre1 As Integer
    Dim nombre2 As Integer
    Dim somme As Integer

→   nombre1 = 10                    ' ← 1ère étape - flèche ici au début
    nombre2 = 20                    ' ← 2ème étape après F8
    somme = nombre1 + nombre2       ' ← 3ème étape après F8
    MsgBox "La somme est : " & somme ' ← 4ème étape après F8
End Sub
```

**Étapes d'exécution** :
1. **F8** : La flèche se place sur `nombre1 = 10`
2. **F8** : Exécute `nombre1 = 10`, la flèche passe à `nombre2 = 20`
3. **F8** : Exécute `nombre2 = 20`, la flèche passe à `somme = nombre1 + nombre2`
4. **F8** : Exécute le calcul, la flèche passe à `MsgBox`
5. **F8** : Affiche le message, fin de la procédure

## Gérer les boucles en pas à pas

L'exécution pas à pas est particulièrement utile pour comprendre les boucles :

```vba
Sub ExempleBoucle()
    Dim i As Integer

→   For i = 1 To 3                  ' ← Démarrage
        Debug.Print "Itération " & i
    Next i
    Debug.Print "Boucle terminée"
End Sub
```

**Déroulement** :
- **1er F8** : Initialise la boucle (i = 1)
- **2ème F8** : Entre dans la boucle, exécute le Debug.Print
- **3ème F8** : Va au Next, incrémente i (i = 2)
- **4ème F8** : Retourne au For, vérifie la condition
- **5ème F8** : Execute le Debug.Print pour i = 2
- Et ainsi de suite...

## Examiner les variables pendant l'exécution

Pendant l'exécution pas à pas, vous pouvez examiner vos variables à tout moment :

### Survol avec la souris
Passez votre souris sur n'importe quelle variable pour voir sa valeur actuelle dans une bulle d'information.

### Fenêtre d'exécution immédiate
- Appuyez sur **Ctrl + G** pour ouvrir la fenêtre
- Tapez `? nomDeLaVariable` pour voir sa valeur
- Vous pouvez même **modifier** une variable : `nomDeLaVariable = nouvelleValeur`

### Fenêtre des variables locales
- **Affichage** > **Fenêtre des variables locales**
- Montre toutes les variables de la procédure actuelle avec leurs valeurs
- Se met à jour automatiquement à chaque étape

## Conseils pratiques pour l'exécution pas à pas

### Choisir le bon mode
- **F8 (Step Into)** : Quand vous voulez tout voir en détail
- **Maj + F8 (Step Over)** : Quand vous voulez ignorer les détails des procédures appelées
- **Ctrl + Maj + F8 (Step Out)** : Quand vous voulez sortir d'une procédure

### Alterner les modes
Vous pouvez **changer de mode** en cours d'exécution selon vos besoins. Par exemple, utiliser Step Into pour entrer dans une procédure, puis Step Out pour en sortir.

### Utiliser avec les points d'arrêt
Combinez l'exécution pas à pas avec les points d'arrêt :
- Placez un point d'arrêt pour aller rapidement à la zone problématique
- Utilisez ensuite l'exécution pas à pas pour analyser en détail

### Reprendre l'exécution normale
À tout moment, vous pouvez appuyer sur **F5** pour reprendre l'exécution normale jusqu'au prochain point d'arrêt ou à la fin du programme.

## Situations où l'exécution pas à pas est particulièrement utile

**Débogage de conditions complexes** : Pour voir quel chemin prend votre code dans des structures If...Then...Else imbriquées.

**Analyse de boucles** : Pour comprendre comment vos boucles se comportent et si elles se terminent correctement.

**Vérification de calculs** : Pour suivre étape par étape des calculs complexes.

**Compréhension du code d'autrui** : Pour apprendre comment fonctionne un code que vous n'avez pas écrit.

**Localisation d'erreurs intermittentes** : Pour capturer des erreurs qui ne se produisent que dans certaines conditions.

## Limitations et précautions

**Temps d'exécution** : L'exécution pas à pas prend beaucoup plus de temps que l'exécution normale. Ne l'utilisez que pour déboguer.

**Codes longs** : Pour de très longs programmes, utilisez des points d'arrêt pour aller directement aux zones intéressantes.

**Procédures système** : Vous ne pouvez pas faire du pas à pas dans les procédures intégrées à Excel (comme les fonctions de feuille de calcul).

**État des objets** : Attention, certaines actions peuvent modifier l'état d'Excel pendant le débogage (sélections, calculs, etc.).

L'exécution pas à pas est un outil indispensable pour comprendre et déboguer votre code VBA. Combinée avec l'observation des variables et l'utilisation judicieuse des points d'arrêt, elle vous permet de résoudre même les problèmes les plus complexes.

⏭️ [Fenêtre de surveillance (Watch)](/19-debogage-tests/03-fenetre-surveillance-watch.md)
