🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.1. Classes et objets

## Comprendre les classes et les objets

### Qu'est-ce qu'une classe ?

Une **classe** est comme un **modèle** ou un **plan de construction** qui définit les caractéristiques et les comportements d'un type d'objet. C'est un peu comme le plan d'architecte d'une maison : le plan décrit comment construire la maison, mais ce n'est pas la maison elle-même.

**Analogie simple :**
- La classe `Voiture` est le plan qui décrit qu'une voiture a une marque, une couleur, et peut démarrer ou s'arrêter
- Chaque voiture concrète (la Renault rouge de Paul, la BMW bleue de Marie) est un **objet** créé à partir de cette classe

### Qu'est-ce qu'un objet ?

Un **objet** est une **instance** d'une classe. C'est l'élément concret créé à partir du modèle (la classe). Si la classe est le moule, l'objet est ce qui sort du moule.

**Dans VBA :**
- Vous **définissez** une classe une seule fois
- Vous pouvez **créer** autant d'objets que nécessaire à partir de cette classe

## Créer votre première classe en VBA

### Étape 1 : Insérer un module de classe

1. Dans l'éditeur VBA (Alt + F11)
2. Clic droit sur votre projet dans l'Explorateur de projets
3. **Insertion** → **Module de classe**
4. Un nouveau module appelé "Class1" apparaît

### Étape 2 : Renommer la classe

1. Sélectionnez "Class1" dans l'Explorateur de projets
2. Dans la fenêtre **Propriétés** (F4 si elle n'est pas visible), changez la propriété **(Name)** de "Class1" vers un nom plus parlant, par exemple "Personne"

### Exemple simple : Classe Personne

Voici une classe simple pour représenter une personne :

```vba
' Module de classe : Personne
' Déclaration des variables privées (données de l'objet)
Private mNom As String
Private mAge As Integer

' Propriété pour accéder au nom
Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(valeur As String)
    mNom = valeur
End Property

' Propriété pour accéder à l'âge
Public Property Get Age() As Integer
    Age = mAge
End Property

Public Property Let Age(valeur As Integer)
    If valeur >= 0 Then
        mAge = valeur
    Else
        Err.Raise 5, , "L'âge ne peut pas être négatif"
    End If
End Property

' Méthode pour faire parler la personne
Public Sub Presenter()
    MsgBox "Bonjour, je m'appelle " & mNom & " et j'ai " & mAge & " ans."
End Sub

' Méthode pour calculer l'année de naissance
Public Function AnneeNaissance() As Integer
    AnneeNaissance = Year(Now) - mAge
End Function
```

### Comprendre le code

#### Variables privées
```vba
Private mNom As String
Private mAge As Integer
```
- Le mot-clé `Private` signifie que ces variables ne sont accessibles que depuis l'intérieur de la classe
- La convention `m` au début indique que c'est une variable **membre** de la classe
- Ces variables stockent les **données** de chaque objet

#### Propriétés (Property Get/Let)
```vba
Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(valeur As String)
    mNom = valeur
End Property
```

- `Property Get` : permet de **lire** la valeur d'une propriété
- `Property Let` : permet d'**écrire** (modifier) la valeur d'une propriété
- `Public` : ces propriétés sont accessibles depuis l'extérieur de la classe

#### Méthodes
```vba
Public Sub Presenter()
    MsgBox "Bonjour, je m'appelle " & mNom & " et j'ai " & mAge & " ans."
End Sub
```
- Une méthode est une procédure (`Sub`) ou fonction (`Function`) qui définit ce que l'objet peut **faire**
- `Public` : ces méthodes peuvent être appelées depuis l'extérieur de la classe

## Utiliser votre classe

### Créer des objets (instanciation)

Voici comment utiliser la classe `Personne` dans un module standard :

```vba
Sub TestPersonne()
    ' Déclaration d'une variable objet
    Dim personne1 As Personne

    ' Création de l'objet (instanciation)
    Set personne1 = New Personne

    ' Utilisation des propriétés
    personne1.Nom = "Marie Dupont"
    personne1.Age = 25

    ' Appel d'une méthode
    personne1.Presenter  ' Affiche : "Bonjour, je m'appelle Marie Dupont et j'ai 25 ans."

    ' Utilisation d'une fonction
    MsgBox "Année de naissance : " & personne1.AnneeNaissance()

    ' Libération de la mémoire
    Set personne1 = Nothing
End Sub
```

### Créer plusieurs objets

```vba
Sub TestPlusieursPersonnes()
    ' Création de plusieurs objets
    Dim chef As New Personne
    Dim employe As New Personne

    ' Configuration du chef
    chef.Nom = "Paul Martin"
    chef.Age = 45

    ' Configuration de l'employé
    employe.Nom = "Sophie Leroy"
    employe.Age = 28

    ' Utilisation
    chef.Presenter
    employe.Presenter

    ' Comparaison
    If chef.Age > employe.Age Then
        MsgBox chef.Nom & " est plus âgé que " & employe.Nom
    End If
End Sub
```

## Syntaxes importantes à retenir

### Déclaration et création d'objets

```vba
' Méthode 1 : Déclaration puis création
Dim monObjet As MaClasse
Set monObjet = New MaClasse

' Méthode 2 : Déclaration et création en une ligne
Dim monObjet As New MaClasse

' Méthode 3 : Création directe
Set monObjet = New MaClasse
```

### Libération de la mémoire

```vba
' Toujours libérer les objets quand on n'en a plus besoin
Set monObjet = Nothing
```

### Vérifier si un objet existe

```vba
If Not monObjet Is Nothing Then
    ' L'objet existe, on peut l'utiliser
    monObjet.MaMethode
End If
```

## Différences entre classes et modules standards

| Aspect | Module standard | Module de classe |
|--------|-----------------|------------------|
| **Utilisation** | Procédures globales | Modèles d'objets |
| **Instanciation** | Pas besoin | `Set obj = New MaClasse` |
| **Variables** | Globales ou locales | Membres de l'objet |
| **Copies** | Une seule version | Autant d'objets que nécessaire |
| **Données** | Partagées | Propres à chaque objet |

## Exemple concret : Classe Produit

Voici un exemple plus pratique pour une application de gestion :

```vba
' Module de classe : Produit
Private mReference As String
Private mNom As String
Private mPrix As Double
Private mStock As Integer

' Propriétés
Public Property Get Reference() As String
    Reference = mReference
End Property

Public Property Let Reference(valeur As String)
    mReference = UCase(valeur)  ' Toujours en majuscules
End Property

Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(valeur As String)
    mNom = valeur
End Property

Public Property Get Prix() As Double
    Prix = mPrix
End Property

Public Property Let Prix(valeur As Double)
    If valeur >= 0 Then
        mPrix = valeur
    Else
        Err.Raise 5, , "Le prix ne peut pas être négatif"
    End If
End Property

Public Property Get Stock() As Integer
    Stock = mStock
End Property

Public Property Let Stock(valeur As Integer)
    If valeur >= 0 Then
        mStock = valeur
    Else
        Err.Raise 5, , "Le stock ne peut pas être négatif"
    End If
End Property

' Méthodes
Public Function ValeurTotaleStock() As Double
    ValeurTotaleStock = mPrix * mStock
End Function

Public Sub AjouterStock(quantite As Integer)
    If quantite > 0 Then
        mStock = mStock + quantite
    End If
End Sub

Public Function RetirerStock(quantite As Integer) As Boolean
    If quantite <= mStock Then
        mStock = mStock - quantite
        RetirerStock = True
    Else
        RetirerStock = False
    End If
End Function

Public Sub AfficherInfo()
    Debug.Print "Produit: " & mReference & " - " & mNom
    Debug.Print "Prix: " & Format(mPrix, "0.00") & "€"
    Debug.Print "Stock: " & mStock & " unités"
    Debug.Print "Valeur stock: " & Format(ValeurTotaleStock(), "0.00") & "€"
    Debug.Print "---"
End Sub
```

### Utilisation de la classe Produit

```vba
Sub TestProduit()
    ' Création d'un produit
    Dim produit As New Produit

    ' Configuration
    produit.Reference = "ref001"
    produit.Nom = "Ordinateur portable"
    produit.Prix = 799.99
    produit.Stock = 15

    ' Utilisation
    produit.AfficherInfo

    ' Vente de 3 unités
    If produit.RetirerStock(3) Then
        Debug.Print "Vente réussie, stock restant: " & produit.Stock
    Else
        Debug.Print "Stock insuffisant"
    End If

    ' Réassort
    produit.AjouterStock(10)
    Debug.Print "Après réassort: " & produit.Stock & " unités"
End Sub
```

## Points importants à retenir

### Encapsulation
- Les variables sont `Private` : elles ne sont pas directement accessibles depuis l'extérieur
- L'accès se fait via les propriétés `Public`, ce qui permet de contrôler et valider les valeurs

### Convention de nommage
- Variables membres : préfixe `m` (ex: `mNom`)
- Classes : PascalCase (ex: `Personne`, `Produit`)
- Propriétés et méthodes : PascalCase (ex: `AnneeNaissance`)

### Gestion de la mémoire
- Toujours utiliser `Set` pour affecter des objets
- Libérer avec `Set objet = Nothing` quand terminé
- VBA gère automatiquement la mémoire, mais c'est une bonne pratique

### Validation des données
- Les propriétés `Let` permettent de valider les données avant de les stocker
- Utiliser `Err.Raise` pour signaler les erreurs

Cette approche orientée objet rend votre code plus organisé, plus facile à maintenir et plus proche de la logique métier de votre application.

⏭️
