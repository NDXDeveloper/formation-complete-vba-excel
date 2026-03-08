🔝 Retour au [Sommaire](/SOMMAIRE.md)

# B. Codes d'erreur courants

## Introduction

Lorsque vous programmez en VBA, il est normal de rencontrer des erreurs. Chaque erreur a un numéro (code) et un message qui vous aide à comprendre ce qui ne va pas. Cette annexe présente les erreurs les plus fréquentes que vous rencontrerez en tant que débutant, avec des explications simples et des solutions pratiques.

**Comment utiliser cette référence :**
- **Code d'erreur** : Le numéro que VBA affiche
- **Nom de l'erreur** : Le nom technique de l'erreur
- **Description** : Ce que signifie cette erreur en termes simples
- **Causes courantes** : Pourquoi cette erreur se produit généralement
- **Solutions** : Comment corriger le problème
- **Exemple** : Un cas concret pour mieux comprendre

---

## 1. Erreurs de syntaxe et compilation

### Erreur 1004 - Erreur définie par l'application ou par l'objet
**Description :** Cette erreur générale indique qu'Excel ne peut pas effectuer l'action demandée  
**Causes courantes :**  
- Tentative d'accès à une cellule qui n'existe pas
- Référence à une feuille qui n'existe pas
- Opération non autorisée sur un objet protégé

**Solutions :**
- Vérifiez que les noms de feuilles et cellules existent
- Assurez-vous que les feuilles ne sont pas protégées
- Contrôlez que les plages de cellules sont valides

**Exemple :**
```vba
' ERREUR : cette feuille n'existe peut-être pas
Worksheets("FeuilleInexistante").Range("A1").Value = "Test"

' SOLUTION : vérifiez d'abord l'existence
If WorksheetExists("FeuilleInexistante") Then
    Worksheets("FeuilleInexistante").Range("A1").Value = "Test"
End If
```

### Erreur 9 - Indice hors limites
**Description :** Vous essayez d'accéder à un élément qui n'existe pas dans une collection  
**Causes courantes :**  
- Référence à une feuille par un mauvais numéro
- Accès à un élément de tableau inexistant
- Nom de feuille mal orthographié

**Solutions :**
- Vérifiez les noms et numéros utilisés
- Comptez le nombre d'éléments dans votre collection
- Utilisez les bons indices (commencent souvent à 1 en VBA)

**Exemple :**
```vba
' ERREUR : si vous n'avez que 3 feuilles
Dim maFeuille As Worksheet  
Set maFeuille = Worksheets(5) ' Erreur 9  

' SOLUTION : utilisez un indice valide ou un nom
Set maFeuille = Worksheets(1) ' ou Worksheets("Feuil1")
```

### Erreur 13 - Type incompatible
**Description :** Vous essayez de mettre une valeur d'un type dans une variable d'un autre type  
**Causes courantes :**  
- Mettre du texte dans une variable numérique
- Assigner une date invalide
- Mélanger des types de données

**Solutions :**
- Vérifiez les types de vos variables
- Utilisez les fonctions de conversion (CInt, CStr, etc.)
- Testez les valeurs avant conversion (IsNumeric, IsDate)

**Exemple :**
```vba
' ERREUR : mettre du texte dans un nombre
Dim nombre As Integer  
nombre = "Bonjour" ' Erreur 13  

' SOLUTION : conversion ou vérification
If IsNumeric("123") Then
    nombre = CInt("123")
End If
```

---

## 2. Erreurs d'exécution

### Erreur 91 - Variable objet ou variable de bloc With non définie
**Description :** Vous essayez d'utiliser un objet qui n'a pas été créé ou assigné  
**Causes courantes :**  
- Oublier de définir un objet avec Set
- Objet détruit ou fermé
- Variable d'objet = Nothing

**Solutions :**
- Utilisez toujours "Set" pour assigner des objets
- Vérifiez que l'objet existe avant de l'utiliser
- Initialisez vos variables d'objet

**Exemple :**
```vba
' ERREUR : objet non défini
Dim monClasseur As Workbook  
monClasseur.Close ' Erreur 91  

' SOLUTION : définir l'objet avec Set
Set monClasseur = ActiveWorkbook  
monClasseur.Close  
```

### Erreur 1004 - La méthode Range de l'objet Worksheet a échoué
**Description :** Problème avec la référence à une plage de cellules  
**Causes courantes :**  
- Syntaxe de plage incorrecte
- Nom de plage inexistant
- Caractères interdits dans la référence

**Solutions :**
- Vérifiez la syntaxe des références (A1, B2:C5, etc.)
- Assurez-vous que les plages nommées existent
- Évitez les caractères spéciaux

**Exemple :**
```vba
' ERREUR : syntaxe incorrecte
Range("A1:Z") ' Erreur 1004

' SOLUTION : syntaxe complète
Range("A1:Z10")
```

### Erreur 5 - Appel ou argument de procédure non valide
**Description :** Les paramètres passés à une fonction ne sont pas corrects  
**Causes courantes :**  
- Nombre incorrect de paramètres
- Type de paramètre incorrect
- Valeur de paramètre hors limites

**Solutions :**
- Vérifiez la syntaxe de la fonction
- Contrôlez les types et valeurs des paramètres
- Consultez l'aide VBA pour la fonction

**Exemple :**
```vba
' ERREUR : paramètre manquant
Mid("Bonjour") ' Erreur 5

' SOLUTION : tous les paramètres requis
Mid("Bonjour", 2, 3)
```

---

## 3. Erreurs de logique

### Erreur 11 - Division par zéro
**Description :** Vous essayez de diviser un nombre par zéro  
**Causes courantes :**  
- Variable non initialisée (valeur = 0)
- Calcul qui aboutit à zéro
- Donnée manquante

**Solutions :**
- Vérifiez que le diviseur n'est pas zéro avant la division
- Initialisez vos variables
- Gérez les cas particuliers

**Exemple :**
```vba
' ERREUR : division par zéro
Dim diviseur As Integer ' vaut 0 par défaut  
Dim resultat As Double  
resultat = 10 / diviseur ' Erreur 11  

' SOLUTION : vérification avant division
If diviseur <> 0 Then
    resultat = 10 / diviseur
Else
    MsgBox "Division par zéro impossible"
End If
```

### Erreur 6 - Dépassement de capacité
**Description :** Le résultat d'un calcul dépasse les limites du type de variable  
**Causes courantes :**  
- Résultat trop grand pour le type Integer
- Multiplication de grands nombres
- Boucle infinie qui incrémente

**Solutions :**
- Utilisez des types de variables plus grands (Long, Double)
- Vérifiez vos calculs
- Contrôlez les boucles

**Exemple :**
```vba
' ERREUR : trop grand pour Integer (limite : 32 767)
Dim nombre As Integer  
nombre = 50000 ' Erreur 6  

' SOLUTION : utiliser un type plus grand
Dim grandNombre As Long  
grandNombre = 50000 ' OK  
```

---

## 4. Erreurs de ressources

### Erreur 7 - Mémoire insuffisante
**Description :** VBA n'a plus assez de mémoire pour continuer  
**Causes courantes :**  
- Tableaux très volumineux
- Boucles qui créent trop d'objets
- Variables non libérées

**Solutions :**
- Libérez les variables d'objet (Set objet = Nothing)
- Optimisez vos boucles
- Travaillez par petits blocs de données

**Exemple :**
```vba
' BONNE PRATIQUE : libérer les objets
Dim monClasseur As Workbook  
Set monClasseur = Workbooks.Open("fichier.xlsx")  
' ... travail avec le classeur
monClasseur.Close  
Set monClasseur = Nothing ' Libère la mémoire  
```

### Erreur 70 - Autorisation refusée
**Description :** VBA ne peut pas accéder au fichier ou effectuer l'opération  
**Causes courantes :**  
- Fichier ouvert dans une autre application
- Fichier en lecture seule
- Permissions insuffisantes

**Solutions :**
- Fermez le fichier dans les autres applications
- Vérifiez les propriétés du fichier
- Exécutez Excel en tant qu'administrateur si nécessaire

---

## 5. Conseils pour déboguer

### Comment identifier une erreur
1. **Lisez le message** : VBA vous donne souvent des indices précieux
2. **Notez le numéro** : Chaque erreur a un code spécifique
3. **Regardez la ligne surlignée** : VBA vous montre où ça coince
4. **Vérifiez les variables** : Placez votre souris sur les variables pour voir leur valeur

### Prévenir les erreurs
1. **Déclarez toujours vos variables** : Utilisez `Dim`
2. **Initialisez vos variables** : Donnez-leur une valeur de départ
3. **Vérifiez avant d'agir** : Testez l'existence des objets
4. **Utilisez la gestion d'erreur** : `On Error` pour intercepter les problèmes

### Techniques de débogage
```vba
Sub ExempleDebuggage()
    Dim i As Integer

    ' Affichage pour vérifier les valeurs
    Debug.Print "Valeur de i : " & i

    ' Point d'arrêt : F9 sur cette ligne
    i = i + 1

    ' Message pour suivre l'exécution
    MsgBox "i vaut maintenant : " & i
End Sub
```

### Gestion d'erreur simple
```vba
Sub AvecGestionErreur()
    On Error Resume Next ' Continue malgré les erreurs

    ' Code qui peut générer une erreur
    Worksheets("FeuilleQuiNExistePeutEtre").Activate

    ' Vérification s'il y a eu une erreur
    If Err.Number <> 0 Then
        MsgBox "Erreur : " & Err.Description
        Err.Clear ' Efface l'erreur
    End If

    On Error GoTo 0 ' Remet la gestion d'erreur normale
End Sub
```

---

## Que faire quand vous rencontrez une erreur

### Étapes à suivre :
1. **Ne paniquez pas** : Les erreurs font partie de l'apprentissage
2. **Lisez le message** : Il contient souvent la solution
3. **Vérifiez cette annexe** : Cherchez le code d'erreur
4. **Testez étape par étape** : Commentez une partie du code pour isoler le problème
5. **Demandez de l'aide** : Forums, collègues, documentation

### Erreurs fréquentes des débutants
- Oublier `Set` pour les objets
- Mélanger les types de données
- Références de cellules incorrectes
- Variables non déclarées
- Boucles infinies

### Ressources utiles
- **F1** : Aide VBA intégrée
- **Debug.Print** : Affiche des valeurs dans la fenêtre d'exécution immédiate
- **Points d'arrêt** : F9 pour arrêter l'exécution et examiner
- **Fenêtre variables locales** : Voir toutes les variables actuelles

**Rappelez-vous :** Chaque erreur est une occasion d'apprendre. Plus vous en rencontrerez, plus vous deviendrez expert en VBA !

⏭️ [C. Raccourcis clavier de l'éditeur VBA](/annexes/c-raccourcis-clavier-editeur-vba.md)
