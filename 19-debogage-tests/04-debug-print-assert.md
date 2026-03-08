🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.4. Debug.Print et Assert

## Introduction aux outils de débogage Debug

VBA propose deux outils particulièrement utiles pour le débogage qui font partie de l'objet `Debug` : **Debug.Print** et **Debug.Assert**. Ces outils sont comme des "assistants silencieux" qui vous aident à comprendre ce qui se passe dans votre code sans perturber l'utilisateur final.

Contrairement aux points d'arrêt ou à l'exécution pas à pas qui interrompent votre programme, ces outils fonctionnent **en arrière-plan** et vous fournissent des informations précieuses sans arrêter l'exécution.

## Debug.Print - Votre journal de bord

### Qu'est-ce que Debug.Print ?

`Debug.Print` est une instruction qui affiche des informations dans la **fenêtre d'exécution immédiate** sans interrompre l'exécution de votre programme. C'est comme tenir un journal de bord de ce qui se passe dans votre code.

Imaginez que vous voulez savoir ce qui se passe dans votre programme, mais sans déranger l'utilisateur avec des messages popup. `Debug.Print` vous permet d'écrire des "notes" que vous seul pouvez voir dans la fenêtre d'exécution immédiate.

### Syntaxe de base

```vba
Debug.Print "Votre message ici"  
Debug.Print variable  
Debug.Print "Message : " & variable  
```

### Exemples simples

```vba
Sub ExempleDebugPrint()
    Dim nom As String
    Dim age As Integer

    nom = "Marie"
    age = 25

    Debug.Print "Début du programme"
    Debug.Print "Nom : " & nom
    Debug.Print "Age : " & age
    Debug.Print "Fin du programme"
End Sub
```

Quand vous exécutez ce code, dans la fenêtre d'exécution immédiate (Ctrl+G), vous verrez :
```
Début du programme  
Nom : Marie  
Age : 25  
Fin du programme  
```

### Pourquoi utiliser Debug.Print ?

**Non intrusif** : N'interrompt pas l'exécution et n'affiche rien à l'utilisateur final.

**Suivi des valeurs** : Permet de voir comment les variables évoluent au fil du temps.

**Traçage du flux** : Aide à comprendre dans quel ordre vos procédures s'exécutent.

**Débogage de boucles** : Particulièrement utile pour voir ce qui se passe à chaque itération.

**Historique persistant** : Les messages restent visibles même après l'exécution, contrairement aux variables que vous devez surveiller en temps réel.

## Utilisation avancée de Debug.Print

### Tracer l'exécution des procédures

```vba
Sub ProcedurePrincipale()
    Debug.Print "*** Début ProcedurePrincipale ***"

    Dim resultat As Integer
    resultat = CalculerSomme(10, 20)

    Debug.Print "Résultat reçu : " & resultat
    Debug.Print "*** Fin ProcedurePrincipale ***"
End Sub

Function CalculerSomme(a As Integer, b As Integer) As Integer
    Debug.Print "  -> Entrée dans CalculerSomme avec a=" & a & ", b=" & b

    CalculerSomme = a + b

    Debug.Print "  -> Sortie de CalculerSomme, résultat=" & CalculerSomme
End Function
```

### Surveiller les boucles

```vba
Sub ExempleBoucleAvecDebug()
    Dim i As Integer
    Dim somme As Integer

    Debug.Print "=== Début de la boucle ==="

    For i = 1 To 5
        somme = somme + i
        Debug.Print "Itération " & i & " : somme = " & somme
    Next i

    Debug.Print "=== Fin de la boucle, somme finale = " & somme & " ==="
End Sub
```

### Afficher des informations conditionnelles

```vba
Sub ExempleConditionnel()
    Dim nombre As Integer

    For nombre = 1 To 10
        If nombre Mod 2 = 0 Then
            Debug.Print nombre & " est pair"
        End If

        If nombre > 7 Then
            Debug.Print "Attention : " & nombre & " est supérieur à 7"
        End If
    Next nombre
End Sub
```

### Surveiller les propriétés d'objets

```vba
Sub SurveillerObjetsExcel()
    Debug.Print "Feuille active : " & ActiveSheet.Name
    Debug.Print "Cellule sélectionnée : " & Selection.Address
    Debug.Print "Valeur en A1 : " & Range("A1").Value
    Debug.Print "Nombre de feuilles : " & Worksheets.Count
End Sub
```

## Debug.Assert - Votre garde du corps

### Qu'est-ce que Debug.Assert ?

`Debug.Assert` est un outil qui vérifie qu'une condition est vraie. Si la condition est fausse, le programme **s'arrête immédiatement** à cet endroit, vous permettant d'examiner la situation.

C'est comme avoir un garde du corps qui surveille constamment que tout va bien et qui vous alerte immédiatement si quelque chose d'anormal se produit.

### Syntaxe de base

```vba
Debug.Assert condition
```

Si `condition` est `True`, le programme continue normalement.  
Si `condition` est `False`, le programme s'arrête sur cette ligne.  

### Exemples de base

```vba
Sub ExempleAssert()
    Dim nombre As Integer
    nombre = 5

    ' Cette assertion passera (5 > 0 est vrai)
    Debug.Assert nombre > 0

    ' Cette assertion échouera et arrêtera le programme
    Debug.Assert nombre > 10

    Debug.Print "Cette ligne ne s'exécutera jamais"
End Sub
```

### Pourquoi utiliser Debug.Assert ?

**Validation des hypothèses** : Vérifiez que vos suppositions sur les données sont correctes.

**Détection précoce d'erreurs** : Arrêtez le programme dès qu'une condition anormale est détectée.

**Documentation vivante** : Les assertions documentent vos attentes dans le code.

**Débogage efficace** : Arrêt automatique exactement là où le problème se produit.

## Utilisations pratiques de Debug.Assert

### Valider les paramètres d'entrée

```vba
Function CalculerRacineCarree(nombre As Double) As Double
    ' Vérifier que le nombre n'est pas négatif
    Debug.Assert nombre >= 0

    CalculerRacineCarree = Sqr(nombre)
End Function
```

### Vérifier les états d'objets

```vba
Sub TravailerAvecFeuille()
    ' S'assurer qu'on a bien une feuille active
    Debug.Assert Not ActiveSheet Is Nothing

    ' S'assurer qu'on n'est pas sur une feuille protégée
    Debug.Assert ActiveSheet.ProtectContents = False

    ' Maintenant on peut travailler en sécurité
    Range("A1").Value = "Test"
End Sub
```

### Surveiller les boucles

```vba
Sub ExempleBoucleAvecAssert()
    Dim i As Integer
    Dim compteur As Integer

    For i = 1 To 100
        compteur = compteur + 1

        ' S'assurer que le compteur ne dépasse jamais i
        Debug.Assert compteur <= i

        ' Protection contre les boucles infinies
        Debug.Assert compteur < 1000
    Next i
End Sub
```

### Valider les résultats de calculs

```vba
Function DiviserNombres(dividende As Double, diviseur As Double) As Double
    ' Vérifier qu'on ne divise pas par zéro
    Debug.Assert diviseur <> 0

    DiviserNombres = dividende / diviseur

    ' Vérifier que le résultat est dans une plage raisonnable
    Debug.Assert Not (DiviserNombres > 1E+15)  ' Pas de valeur astronomique
End Function
```

## Combinaison de Debug.Print et Debug.Assert

Ces deux outils se complètent parfaitement :

```vba
Sub ExempleCombinaison()
    Dim tableau(1 To 10) As Integer
    Dim i As Integer
    Dim somme As Integer

    ' Remplir le tableau
    For i = 1 To 10
        tableau(i) = i * 2
        Debug.Print "tableau(" & i & ") = " & tableau(i)
    Next i

    ' Calculer la somme
    For i = 1 To 10
        somme = somme + tableau(i)

        ' Vérifier que la somme augmente toujours
        Debug.Assert somme > 0

        Debug.Print "Après ajout de " & tableau(i) & ", somme = " & somme
    Next i

    ' Vérifier le résultat final
    Debug.Print "Somme finale : " & somme
    Debug.Assert somme = 110  ' 2+4+6+8+10+12+14+16+18+20 = 110
End Sub
```

## La fenêtre d'exécution immédiate

### Comment l'ouvrir
- **Ctrl + G** dans l'éditeur VBA
- Menu **Affichage** > **Fenêtre d'exécution immédiate**

### Utilisation interactive
Vous pouvez aussi utiliser la fenêtre d'exécution immédiate de manière interactive :

```vba
' Tapez directement dans la fenêtre :
? Range("A1").Value
? ActiveSheet.Name
? Date
```

### Nettoyer la fenêtre
Pour effacer le contenu : **Ctrl + A** puis **Suppr**

## Bonnes pratiques

### Pour Debug.Print

**Messages descriptifs** : Utilisez des messages clairs qui expliquent ce que vous affichez.
```vba
' Bien
Debug.Print "Prix total après remise : " & prixFinal

' Moins bien
Debug.Print prixFinal
```

**Structuration des messages** : Utilisez des préfixes pour organiser vos messages.
```vba
Debug.Print "*** DEBUT ProcedureImportante ***"  
Debug.Print "  -> Paramètre reçu : " & parametre  
Debug.Print "  -> Calcul en cours..."  
Debug.Print "  -> Résultat : " & resultat  
Debug.Print "*** FIN ProcedureImportante ***"  
```

**Utilisation temporaire** : N'oubliez pas de retirer ou commenter les Debug.Print avant la version finale.

### Pour Debug.Assert

**Conditions claires** : Écrivez des conditions facilement compréhensibles.
```vba
' Bien
Debug.Assert age >= 0 And age <= 150

' Moins bien
Debug.Assert Not (age < 0 Or age > 150)
```

**Messages d'erreur** : Bien que Debug.Assert n'affiche pas de message, placez un commentaire explicatif.
```vba
Debug.Assert diviseur <> 0  ' Division par zéro interdite
```

**Utilisation en développement** : Les assertions sont principalement pour le développement, pas pour la production.

## Avantages et limitations

### Avantages
- **Simplicité** : Très faciles à utiliser
- **Non intrusifs** : N'affectent pas l'utilisateur final
- **Flexibilité** : Peuvent être ajoutés ou retirés facilement
- **Débogage en temps réel** : Fournissent des informations pendant l'exécution

### Limitations
- **Debug.Print** : Peut ralentir l'exécution si utilisé massivement
- **Debug.Assert** : S'arrête complètement, pas toujours souhaitable
- **Oublis** : Facile d'oublier de les retirer du code final
- **Visibilité** : Les messages ne sont visibles que dans l'éditeur VBA

## Quand utiliser quoi ?

**Debug.Print** pour :
- Suivre le flux d'exécution
- Surveiller l'évolution des variables
- Déboguer des boucles complexes
- Comprendre le comportement du code

**Debug.Assert** pour :
- Valider des conditions critiques
- Vérifier les paramètres d'entrée
- Détecter des états anormaux
- Arrêter immédiatement en cas de problème

**Les deux ensemble** pour :
- Un débogage complet et efficace
- Valider ET tracer le comportement
- Créer un système de débogage robuste

Debug.Print et Debug.Assert sont des outils essentiels pour tout développeur VBA. Ils vous permettent de créer un système de débogage personnalisé, discret mais puissant, qui vous aide à comprendre et à valider le comportement de votre code en temps réel.

⏭️ [Tests unitaires simples](/19-debogage-tests/05-tests-unitaires-simples.md)
