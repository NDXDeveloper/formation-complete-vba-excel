🔝 Retour au [Sommaire](/SOMMAIRE.md)

# Chapitre 5 : Structures de contrôle

## Introduction

Les **structures de contrôle** sont les éléments qui permettent à votre programme de prendre des décisions et de répéter des actions. Sans elles, votre code ne pourrait qu'exécuter une série d'instructions dans l'ordre, du début à la fin, comme une recette de cuisine basique. Avec les structures de contrôle, votre programme devient "intelligent" et peut s'adapter aux différentes situations.

### Qu'est-ce qu'une structure de contrôle ?

Une **structure de contrôle** est un bloc de code qui détermine **l'ordre d'exécution** des instructions. Elle permet de :

- **Prendre des décisions** : "Si cette condition est vraie, alors faire ceci, sinon faire cela"
- **Répéter des actions** : "Tant que cette condition est vraie, continuer à faire ceci"
- **Choisir entre plusieurs options** : "Selon la valeur de cette variable, faire l'action A, B ou C"

### Analogie avec la vie quotidienne

#### La prise de décision
Imaginez que vous sortez de chez vous le matin :
- **SI** il pleut, **ALORS** prenez un parapluie
- **SINON** sortez sans parapluie

#### La répétition
Quand vous faites le ménage :
- **TANT QUE** il reste de la vaisselle sale, **FAIRE** : laver un plat
- **POUR CHAQUE** pièce de la maison, **FAIRE** : passer l'aspirateur

#### Le choix multiple
Quand vous choisissez un moyen de transport :
- **SELON** la distance :
  - **SI** moins de 1 km → à pied
  - **SI** 1-5 km → vélo
  - **SI** plus de 5 km → voiture

## Types de structures de contrôle

### 1. Structures conditionnelles (décisions)

Elles permettent d'exécuter différentes instructions selon que certaines conditions sont vraies ou fausses.

**Exemples de situations :**
- Afficher un message différent selon l'âge de l'utilisateur
- Appliquer une remise si le montant d'achat dépasse un seuil
- Vérifier si une cellule est vide avant d'y écrire

**Structures disponibles en VBA :**
- `If...Then...Else` : La structure conditionnelle de base
- `Select Case` : Pour choisir entre plusieurs options

### 2. Structures répétitives (boucles)

Elles permettent de répéter un bloc d'instructions plusieurs fois, selon différents critères.

**Exemples de situations :**
- Parcourir toutes les lignes d'un tableau
- Répéter un calcul jusqu'à obtenir un résultat précis
- Traiter tous les fichiers d'un dossier
- Demander une saisie à l'utilisateur jusqu'à ce qu'elle soit valide

**Structures disponibles en VBA :**
- `For...Next` : Répéter un nombre déterminé de fois
- `For Each...Next` : Parcourir tous les éléments d'une collection
- `Do...Loop` : Répéter tant qu'une condition est vraie/fausse
- `While...Wend` : Version simplifiée de Do...Loop

### 3. Structures de branchement

Elles permettent de modifier l'ordre normal d'exécution du programme.

**Structures disponibles :**
- `Exit` : Sortir prématurément d'une boucle ou d'une procédure
- `GoTo` : Aller directement à une autre ligne (à éviter généralement)

## Pourquoi les structures de contrôle sont-elles essentielles ?

### Sans structures de contrôle
```vba
Sub ExempleSansStructure()
    MsgBox "Bienvenue"
    MsgBox "Vous avez 25 ans"
    MsgBox "Vous êtes majeur"
    MsgBox "Au revoir"
End Sub
```
Ce code fait toujours la même chose, pour tout le monde.

### Avec structures de contrôle
```vba
Sub ExempleAvecStructure()
    Dim age As Integer
    age = InputBox("Quel est votre âge ?")

    MsgBox "Bienvenue"

    ' Structure conditionnelle
    If age >= 18 Then
        MsgBox "Vous êtes majeur"
    Else
        MsgBox "Vous êtes mineur"
    End If

    MsgBox "Au revoir"
End Sub
```
Ce code s'adapte selon l'âge saisi !

## Combinaison des structures

Les vraies applications combinent différentes structures pour créer des logiques complexes :

```vba
Sub ExempleCombinaison()
    Dim i As Integer
    Dim note As Double

    ' Boucle pour traiter plusieurs étudiants
    For i = 1 To 5
        note = InputBox("Note de l'étudiant " & i & " :")

        ' Condition à l'intérieur de la boucle
        If note >= 10 Then
            MsgBox "Étudiant " & i & " : Admis"
        Else
            MsgBox "Étudiant " & i & " : Recalé"
        End If
    Next i
End Sub
```

## Les opérateurs de comparaison

Pour utiliser les structures de contrôle, vous devez comprendre comment créer des **conditions**. Voici les opérateurs essentiels :

### Opérateurs de comparaison
- `=` : Égal à
- `<>` : Différent de
- `<` : Inférieur à
- `>` : Supérieur à
- `<=` : Inférieur ou égal à
- `>=` : Supérieur ou égal à

### Opérateurs logiques
- `And` : ET logique (les deux conditions doivent être vraies)
- `Or` : OU logique (au moins une condition doit être vraie)
- `Not` : NON logique (inverse la condition)

**Exemples :**
```vba
' Comparaisons simples
If age >= 18 Then...  
If nom = "Marie" Then...  
If prix <> 0 Then...  

' Comparaisons complexes
If age >= 18 And age <= 65 Then...  ' Entre 18 et 65 ans  
If ville = "Paris" Or ville = "Lyon" Then...  ' Paris OU Lyon  
If Not EstVide(Range("A1")) Then...  ' Cellule NON vide  
```

## Objectifs de ce chapitre

À la fin de ce chapitre, vous serez capable de :

- **Utiliser les conditions** pour que votre programme prenne des décisions intelligentes
- **Créer des boucles** pour automatiser les tâches répétitives
- **Combiner différentes structures** pour résoudre des problèmes complexes
- **Choisir la bonne structure** selon le contexte de votre problème
- **Éviter les erreurs courantes** comme les boucles infinies
- **Optimiser vos programmes** en utilisant la structure la plus appropriée

### Prérequis

Avant d'aborder ce chapitre, assurez-vous de maîtriser :
- Les variables et types de données
- Les opérateurs arithmétiques et logiques
- La création de procédures et fonctions
- Les concepts de base de la programmation VBA

### Plan du chapitre

Ce chapitre est organisé de manière progressive :

1. **Instructions conditionnelles** : Apprendre à faire des choix dans votre code
   - If...Then...Else pour les décisions simples et complexes
   - Select Case pour les choix multiples

2. **Structures répétitives (boucles)** : Automatiser les tâches répétitives
   - For...Next pour un nombre déterminé d'itérations
   - For Each...Next pour parcourir des collections
   - Do...Loop pour des répétitions conditionnelles
   - While...Wend comme alternative simple

3. **Instructions de contrôle avancées** : Maîtriser le flux d'exécution
   - Exit pour sortir prématurément
   - GoTo et ses alternatives modernes

### Applications pratiques

Tout au long de ce chapitre, vous verrez des exemples concrets d'utilisation :
- Validation de données saisies par l'utilisateur
- Traitement automatique de grandes quantités de données
- Création de menus interactifs
- Automation de tâches Excel répétitives
- Gestion d'erreurs et de cas particuliers

### Méthode d'apprentissage

Pour chaque structure, nous suivrons cette approche :
1. **Explication du concept** avec des analogies
2. **Syntaxe de base** avec des exemples simples
3. **Cas d'usage pratiques** dans le contexte Excel
4. **Erreurs courantes** et comment les éviter
5. **Bonnes pratiques** pour un code efficace

Cette méthode progressive vous permettra de maîtriser chaque structure avant de passer à la suivante, et de comprendre quand et comment les utiliser dans vos projets réels.

Les structures de contrôle sont le cœur de la logique de programmation. Elles transforment vos scripts linéaires en véritables programmes intelligents capables de s'adapter à toute situation !

⏭️
