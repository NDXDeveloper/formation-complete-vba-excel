🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.2. Tableaux statiques et dynamiques

## Introduction aux types de tableaux

En VBA, il existe deux grandes catégories de tableaux : les **tableaux statiques** et les **tableaux dynamiques**. La différence fondamentale réside dans la **flexibilité de leur taille** : les tableaux statiques ont une taille fixe définie à l'avance, tandis que les tableaux dynamiques peuvent changer de taille pendant l'exécution du programme.

**Analogie simple :**
- **Tableau statique** = Un parking avec un nombre fixe de places (50 places, par exemple). Une fois construit, impossible d'ajouter ou de retirer des places.
- **Tableau dynamique** = Un parking modulaire où vous pouvez ajouter ou retirer des places selon vos besoins du moment.

Chaque type a ses avantages et ses inconvénients selon le contexte d'utilisation.

---

## Tableaux statiques

### Qu'est-ce qu'un tableau statique ?

Un **tableau statique** est un tableau dont la taille est **définie au moment de la déclaration** et ne peut plus être modifiée par la suite. Cette taille reste constante pendant toute la durée de vie du tableau.

### Déclaration des tableaux statiques

```vba
Sub DeclarationTableauxStatiques()
    ' Tableau statique de 10 éléments (indices 0 à 9)
    Dim notes(9) As Integer

    ' Tableau statique avec indices personnalisés (1 à 12)
    Dim mois(1 To 12) As String

    ' Tableau statique multidimensionnel (3x4)
    Dim grille(1 To 3, 1 To 4) As Double

    ' Une fois déclarés, la taille ne peut plus changer
    notes(0) = 15
    notes(9) = 18
    ' notes(10) = 20  ' ERREUR : indice hors limites

    mois(1) = "Janvier"
    mois(12) = "Décembre"

    MsgBox "Tableaux statiques déclarés et utilisés"
End Sub
```

### Avantages des tableaux statiques

#### 1. **Simplicité d'utilisation**

```vba
Sub SimplicitéStatique()
    ' Déclaration simple et directe
    Dim joursSemaine(1 To 7) As String

    ' Remplissage direct
    joursSemaine(1) = "Lundi"
    joursSemaine(2) = "Mardi"
    joursSemaine(3) = "Mercredi"
    joursSemaine(4) = "Jeudi"
    joursSemaine(5) = "Vendredi"
    joursSemaine(6) = "Samedi"
    joursSemaine(7) = "Dimanche"

    ' Utilisation immédiate
    MsgBox "Aujourd'hui nous sommes " & joursSemaine(2)
End Sub
```

#### 2. **Performance optimale**

```vba
Sub PerformanceStatique()
    ' Allocation mémoire en une seule fois
    Dim donnees(1 To 1000) As Double

    ' Accès direct très rapide
    Dim debut As Double
    debut = Timer

    Dim i As Long
    For i = 1 To 1000
        donnees(i) = i * 3.14159
    Next i

    Debug.Print "Temps de traitement : " & Format(Timer - debut, "0.000") & " secondes"
End Sub
```

#### 3. **Sécurité de la mémoire**

```vba
Sub SecuriteMemoire()
    ' Taille connue à l'avance = pas de surprise de mémoire
    Dim scores(1 To 100) As Integer  ' Exactement 200 octets alloués

    ' Vérification des limites toujours possible
    Debug.Print "Nombre d'éléments : " & (UBound(scores) - LBound(scores) + 1)
    Debug.Print "Indice minimum : " & LBound(scores)
    Debug.Print "Indice maximum : " & UBound(scores)
End Sub
```

### Inconvénients des tableaux statiques

#### 1. **Inflexibilité de taille**

```vba
Sub ProblemeInflexibilite()
    ' Tableau prévu pour 50 étudiants
    Dim notesEtudiants(1 To 50) As Integer

    ' Que faire si on a soudainement 60 étudiants ?
    ' - Soit on redéclare le tableau (perte des données)
    ' - Soit on utilise un tableau plus grand dès le départ (gaspillage)

    ' Solution de contournement : sur-dimensionner
    Dim notesSecurite(1 To 100) As Integer  ' Plus grand "au cas où"

    MsgBox "Tableau surdimensionné pour éviter les problèmes"
End Sub
```

#### 2. **Gaspillage de mémoire**

```vba
Sub GaspillageMemoire()
    ' Si on prévoit 1000 éléments mais n'en utilise que 10
    Dim granTableau(1 To 1000) As Double  ' 8000 octets alloués

    ' Seulement 10 éléments utilisés = 7920 octets gaspillés
    granTableau(1) = 1.5
    granTableau(2) = 2.7
    ' ... seuls quelques éléments sont utilisés

    MsgBox "990 emplacements inutilisés dans le tableau"
End Sub
```

### Cas d'usage idéaux pour les tableaux statiques

```vba
Sub CasUsageStatiques()
    ' 1. Données de taille connue et constante
    Dim joursMois(1 To 31) As Integer        ' Maximum 31 jours
    Dim heuresJournee(0 To 23) As String     ' Toujours 24 heures
    Dim notes(1 To 20) As Double             ' Classe de 20 élèves fixe

    ' 2. Constantes et données de référence
    Dim couleursPrimaires(1 To 3) As String
    couleursPrimaires(1) = "Rouge"
    couleursPrimaires(2) = "Vert"
    couleursPrimaires(3) = "Bleu"

    ' 3. Buffers et caches de taille fixe
    Dim cacheCalculs(1 To 100) As Double    ' Cache pour 100 calculs max

    MsgBox "Tableaux statiques adaptés à ces situations"
End Sub
```

---

## Tableaux dynamiques

### Qu'est-ce qu'un tableau dynamique ?

Un **tableau dynamique** est un tableau dont la taille peut être **modifiée pendant l'exécution** du programme. Vous pouvez l'agrandir, le rétrécir, ou même le redimensionner complètement selon vos besoins.

### Déclaration des tableaux dynamiques

```vba
Sub DeclarationTableauxDynamiques()
    ' Déclaration sans spécifier de taille (parenthèses vides)
    Dim donneesVariable() As String
    Dim nombresFlexibles() As Double
    Dim grilleEvolutive() As Integer

    ' À ce stade, les tableaux ne sont pas encore utilisables
    ' Il faut les redimensionner avec ReDim

    ' Première allocation de taille
    ReDim donneesVariable(1 To 5)
    ReDim nombresFlexibles(0 To 10)
    ReDim grilleEvolutive(1 To 3, 1 To 3)  ' Tableau 2D dynamique

    ' Maintenant ils sont utilisables
    donneesVariable(1) = "Premier élément"
    nombresFlexibles(0) = 3.14159
    grilleEvolutive(1, 1) = 100

    MsgBox "Tableaux dynamiques créés et redimensionnés"
End Sub
```

### L'instruction ReDim

#### Syntaxe de base

```vba
Sub UtilisationReDim()
    Dim tableau() As Integer

    ' Première allocation
    ReDim tableau(1 To 10)
    tableau(1) = 100
    tableau(10) = 200

    ' Redimensionnement (ATTENTION : efface les données)
    ReDim tableau(1 To 20)
    Debug.Print tableau(1)   ' Affiche 0 (données perdues)

    ' Redimensionnement avec préservation des données
    ReDim Preserve tableau(1 To 25)
    tableau(1) = 100  ' On doit remettre les valeurs
    tableau(25) = 250

    MsgBox "Tableau redimensionné plusieurs fois"
End Sub
```

#### ReDim vs ReDim Preserve

```vba
Sub DifferenceRedimPreserve()
    Dim nombres() As Integer

    ' Initialisation
    ReDim nombres(1 To 5)
    nombres(1) = 10
    nombres(2) = 20
    nombres(3) = 30
    nombres(4) = 40
    nombres(5) = 50

    Debug.Print "Avant redimensionnement :"
    Debug.Print "nombres(1) = " & nombres(1)  ' 10
    Debug.Print "nombres(5) = " & nombres(5)  ' 50

    ' ReDim normal (efface les données)
    ReDim nombres(1 To 8)
    Debug.Print "Après ReDim normal :"
    Debug.Print "nombres(1) = " & nombres(1)  ' 0 (données perdues)

    ' Remettre des données
    nombres(1) = 10: nombres(2) = 20: nombres(3) = 30

    ' ReDim Preserve (garde les données)
    ReDim Preserve nombres(1 To 10)
    Debug.Print "Après ReDim Preserve :"
    Debug.Print "nombres(1) = " & nombres(1)  ' 10 (données préservées)
    Debug.Print "nombres(10) = " & nombres(10) ' 0 (nouveau élément)
End Sub
```

### Avantages des tableaux dynamiques

#### 1. **Flexibilité maximale**

```vba
Sub FlexibiliteMaximale()
    Dim listeClients() As String
    Dim nombreClients As Integer

    ' Demander à l'utilisateur combien de clients
    nombreClients = InputBox("Combien de clients à saisir ?")

    ' Adapter la taille exactement aux besoins
    ReDim listeClients(1 To nombreClients)

    ' Remplissage
    Dim i As Integer
    For i = 1 To nombreClients
        listeClients(i) = InputBox("Nom du client " & i & " :")
    Next i

    ' Si besoin d'ajouter un client
    If MsgBox("Ajouter un client supplémentaire ?", vbYesNo) = vbYes Then
        nombreClients = nombreClients + 1
        ReDim Preserve listeClients(1 To nombreClients)
        listeClients(nombreClients) = InputBox("Nom du nouveau client :")
    End If

    MsgBox "Tableau adapté dynamiquement : " & nombreClients & " clients"
End Sub
```

#### 2. **Optimisation de la mémoire**

```vba
Sub OptimisationMemoire()
    Dim donnees() As Double

    ' Lire le nombre de lignes dans Excel
    Dim dernirereLigne As Long
    dernirereLigne = Cells(Rows.Count, 1).End(xlUp).Row

    ' Allouer exactement la mémoire nécessaire
    ReDim donnees(1 To dernirereLigne)

    ' Charger les données
    Dim i As Long
    For i = 1 To dernirereLigne
        donnees(i) = Cells(i, 1).Value
    Next i

    Debug.Print "Mémoire utilisée optimalement pour " & dernirereLigne & " éléments"
End Sub
```

#### 3. **Adaptation aux données variables**

```vba
Sub AdaptationDonneesVariables()
    Dim resultats() As String
    Dim compteur As Integer
    compteur = 0

    ' Traiter les données ligne par ligne
    Dim ligne As Long
    ligne = 1

    Do While Cells(ligne, 1).Value <> ""
        ' Vérifier si la ligne répond aux critères
        If Cells(ligne, 2).Value > 100 Then
            ' Agrandir le tableau d'un élément
            compteur = compteur + 1
            ReDim Preserve resultats(1 To compteur)
            resultats(compteur) = Cells(ligne, 1).Value
        End If
        ligne = ligne + 1
    Loop

    MsgBox "Trouvé " & compteur & " éléments correspondants"
End Sub
```

### Inconvénients des tableaux dynamiques

#### 1. **Performance variable**

```vba
Sub ProblemePerformance()
    Dim donnees() As Integer
    Dim debut As Double

    debut = Timer

    ' LENT : Redimensionnement répétitif
    Dim i As Integer
    For i = 1 To 1000
        ReDim Preserve donnees(1 To i)
        donnees(i) = i
    Next i

    Debug.Print "Temps avec redimensionnements répétés : " & _
                Format(Timer - debut, "0.000") & " secondes"

    ' RAPIDE : Redimensionnement en une fois
    debut = Timer
    ReDim donnees(1 To 1000)
    For i = 1 To 1000
        donnees(i) = i
    Next i

    Debug.Print "Temps avec un seul redimensionnement : " & _
                Format(Timer - debut, "0.000") & " secondes"
End Sub
```

#### 2. **Complexité de gestion**

```vba
Sub ComplexiteGestion()
    Dim tableau() As String
    Dim taille As Integer
    taille = 0

    ' Il faut toujours vérifier si le tableau est initialisé
    On Error Resume Next
    taille = UBound(tableau)
    If Err.Number <> 0 Then
        ' Tableau pas encore initialisé
        ReDim tableau(1 To 1)
        taille = 1
        Err.Clear
    End If
    On Error GoTo 0

    ' Ajouter un élément
    taille = taille + 1
    ReDim Preserve tableau(1 To taille)
    tableau(taille) = "Nouvel élément"

    MsgBox "Gestion plus complexe mais plus flexible"
End Sub
```

### Cas d'usage idéaux pour les tableaux dynamiques

```vba
Sub CasUsageDynamiques()
    ' 1. Lecture de fichiers de taille inconnue
    Dim lignesFichier() As String
    ' Taille déterminée à la lecture

    ' 2. Résultats de recherche ou filtrage
    Dim elementsFiltrés() As Variant
    ' Nombre dépend des critères

    ' 3. Collections évolutives
    Dim historiqueActions() As String
    ' Grandit avec les actions utilisateur

    ' 4. Buffers adaptatifs
    Dim donnéesReseau() As Byte
    ' Taille dépend des données reçues

    MsgBox "Tableaux dynamiques parfaits pour ces cas"
End Sub
```

---

## Comparaison détaillée

### Tableau comparatif

| Aspect | Tableaux Statiques | Tableaux Dynamiques |
|--------|-------------------|---------------------|
| **Déclaration** | `Dim arr(1 To 10) As Integer` | `Dim arr() As Integer` |
| **Initialisation** | Immédiate | Requiert `ReDim` |
| **Modification taille** | ❌ Impossible | ✅ Avec `ReDim` |
| **Performance** | ⭐⭐⭐ Excellente | ⭐⭐ Bonne (dépend usage) |
| **Mémoire** | Fixe (peut gaspiller) | Variable (optimisable) |
| **Complexité** | ⭐ Simple | ⭐⭐ Moyenne |
| **Prévisibilité** | ⭐⭐⭐ Totale | ⭐⭐ Bonne |

### Exemple comparatif pratique

```vba
Sub ComparaisonPratique()
    ' SCÉNARIO : Stocker des notes d'étudiants

    ' === APPROCHE STATIQUE ===
    Dim notesStatiques(1 To 100) As Integer  ' Prévu pour 100 étudiants max
    Dim nbEtudiantsReel As Integer
    nbEtudiantsReel = 25  ' Mais seulement 25 étudiants réels

    ' Avantage : simple à utiliser
    notesStatiques(1) = 15
    notesStatiques(25) = 18

    ' Inconvénient : 75 emplacements gaspillés
    Debug.Print "Statique - Mémoire gaspillée : " & (100 - nbEtudiantsReel) & " emplacements"

    ' === APPROCHE DYNAMIQUE ===
    Dim notesDynamiques() As Integer

    ' Avantage : taille exacte
    ReDim notesDynamiques(1 To nbEtudiantsReel)
    notesDynamiques(1) = 15
    notesDynamiques(25) = 18

    ' Si un nouvel étudiant arrive
    nbEtudiantsReel = nbEtudiantsReel + 1
    ReDim Preserve notesDynamiques(1 To nbEtudiantsReel)
    notesDynamiques(26) = 17

    Debug.Print "Dynamique - Mémoire optimisée : " & nbEtudiantsReel & " emplacements exacts"
End Sub
```

---

## Techniques avancées

### 1. Fonction pour ajouter un élément à un tableau dynamique

```vba
Sub AjouterElement(ByRef tableau() As String, valeur As String)
    Dim nouvelleTaille As Integer

    ' Vérifier si le tableau est initialisé
    On Error Resume Next
    nouvelleTaille = UBound(tableau) + 1
    If Err.Number <> 0 Then
        ' Tableau non initialisé
        nouvelleTaille = 1
        Err.Clear
    End If
    On Error GoTo 0

    ' Redimensionner et ajouter
    ReDim Preserve tableau(1 To nouvelleTaille)
    tableau(nouvelleTaille) = valeur
End Sub

Sub UtiliserAjoutElement()
    Dim malisteList() As String

    Call AjouterElement(malisteList, "Premier")
    Call AjouterElement(malisteList, "Deuxième")
    Call AjouterElement(malisteList, "Troisième")

    Dim i As Integer
    For i = 1 To UBound(malisteList)
        Debug.Print "Element " & i & ": " & malisteList(i)
    Next i
End Sub
```

### 2. Stratégie de redimensionnement par blocs

```vba
Sub RedimensionnementParBlocs()
    Dim donnees() As Integer
    Dim tailleActuelle As Integer
    Dim tailleAllouee As Integer
    Dim blocTaille As Integer

    tailleActuelle = 0
    tailleAllouee = 0
    blocTaille = 10  ' Grandir par blocs de 10

    ' Simulation d'ajout d'éléments
    Dim i As Integer
    For i = 1 To 25
        tailleActuelle = tailleActuelle + 1

        ' Redimensionner seulement si nécessaire
        If tailleActuelle > tailleAllouee Then
            tailleAllouee = tailleAllouee + blocTaille
            ReDim Preserve donnees(1 To tailleAllouee)
            Debug.Print "Redimensionnement à " & tailleAllouee & " pour élément " & i
        End If

        donnees(tailleActuelle) = i * 10
    Next i

    ' Ajuster à la taille finale exacte
    If tailleActuelle < tailleAllouee Then
        ReDim Preserve donnees(1 To tailleActuelle)
        Debug.Print "Ajustement final à " & tailleActuelle
    End If
End Sub
```

### 3. Conversion entre statique et dynamique

```vba
Sub ConversionStatiqueDynamique()
    ' Tableau statique initial
    Dim statique(1 To 5) As String
    statique(1) = "Un": statique(2) = "Deux": statique(3) = "Trois"
    statique(4) = "Quatre": statique(5) = "Cinq"

    ' Conversion en dynamique
    Dim dynamique() As String
    ReDim dynamique(LBound(statique) To UBound(statique))

    Dim i As Integer
    For i = LBound(statique) To UBound(statique)
        dynamique(i) = statique(i)
    Next i

    ' Maintenant on peut redimensionner le dynamique
    ReDim Preserve dynamique(1 To 8)
    dynamique(6) = "Six": dynamique(7) = "Sept": dynamique(8) = "Huit"

    ' Affichage
    For i = 1 To UBound(dynamique)
        Debug.Print "dynamique(" & i & ") = " & dynamique(i)
    Next i
End Sub
```

---

## Guide de choix : Statique ou Dynamique ?

### Utilisez un tableau statique quand :

✅ **La taille est connue et constante**
```vba
Dim joursSemaine(1 To 7) As String
Dim moisAnnee(1 To 12) As String
```

✅ **Performance maximale requise**
```vba
Dim calculsIntensifs(1 To 10000) As Double
```

✅ **Données de référence ou constantes**
```vba
Dim couleursBase(1 To 3) As String
```

✅ **Simplicité prioritaire**
```vba
Dim notes(1 To 20) As Integer  ' Classe de taille fixe
```

### Utilisez un tableau dynamique quand :

✅ **La taille dépend des données**
```vba
Dim lignesExcel() As Variant  ' Dépend du fichier
```

✅ **Le tableau grandit pendant l'exécution**
```vba
Dim historiqueActions() As String  ' Grandit avec l'usage
```

✅ **Optimisation mémoire importante**
```vba
Dim donneesFiltrées() As String  ' Taille inconnue à l'avance
```

✅ **Flexibilité requise**
```vba
Dim resultatsRecherche() As Variant  ' Varie selon critères
```

---

## Récapitulatif

### Points clés à retenir

1. **Tableaux statiques** : Taille fixe, performance optimale, simplicité
2. **Tableaux dynamiques** : Taille variable, flexibilité maximale, complexité accrue
3. **ReDim** : Redimensionne les tableaux dynamiques
4. **ReDim Preserve** : Redimensionne en gardant les données existantes
5. **Choix basé sur le contexte** : Taille connue = statique, taille variable = dynamique

### Modèles de code recommandés

#### Tableau statique type
```vba
' Quand la taille est connue et fixe
Dim donnees(1 To NB_ELEMENTS) As Type
donnees(1) = valeur1
' ... utilisation directe
```

#### Tableau dynamique type
```vba
' Quand la taille varie
Dim donnees() As Type
ReDim donnees(1 To tailleCalculee)
' ou ReDim Preserve pour agrandir
```

### Conseils de performance

- **Évitez** les `ReDim Preserve` répétitifs
- **Préférez** un redimensionnement par blocs
- **Estimez** la taille finale quand c'est possible
- **Testez** les performances sur de gros volumes

Le choix entre statique et dynamique dépend de votre contexte spécifique. Dans le doute, commencez par statique pour la simplicité, puis passez au dynamique si vous avez besoin de flexibilité. Dans la section suivante, nous approfondirons l'utilisation de `ReDim` pour maîtriser complètement les tableaux dynamiques.

⏭️
