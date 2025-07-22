🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.2. Propriétés, méthodes et événements

## Introduction

Les propriétés, méthodes et événements constituent l'**interface** d'un objet - c'est-à-dire tout ce qui est accessible depuis l'extérieur de la classe. Ils définissent ce qu'un objet "possède" (propriétés), ce qu'il "sait faire" (méthodes), et ce qui peut "lui arriver" (événements).

**Analogie simple :**
- **Propriétés** = Caractéristiques d'une voiture (couleur, vitesse, nombre de portes)
- **Méthodes** = Actions qu'elle peut effectuer (démarrer, accélérer, freiner)
- **Événements** = Situations qui peuvent se produire (moteur qui chauffe, essence faible)

## Les Propriétés

### Qu'est-ce qu'une propriété ?

Une **propriété** représente une caractéristique ou un attribut d'un objet. Elle permet de lire ou modifier les données stockées dans l'objet, tout en gardant un contrôle sur la façon dont ces données sont manipulées.

### Types de propriétés en VBA

#### 1. Property Get (Lecture)
Permet de **lire** la valeur d'une propriété.

```vba
Public Property Get Nom() As String
    Nom = mNom
End Property
```

#### 2. Property Let (Écriture - Types simples)
Permet d'**écrire** (modifier) la valeur d'une propriété pour les types simples (String, Integer, Double, etc.).

```vba
Public Property Let Age(valeur As Integer)
    If valeur >= 0 And valeur <= 150 Then
        mAge = valeur
    Else
        Err.Raise 5, , "L'âge doit être entre 0 et 150 ans"
    End If
End Property
```

#### 3. Property Set (Écriture - Objets)
Permet d'**assigner** un objet à une propriété.

```vba
Public Property Set Manager(valeur As Personne)
    Set mManager = valeur
End Property

Public Property Get Manager() As Personne
    Set Manager = mManager
End Property
```

### Exemple complet : Classe Employe

```vba
' Module de classe : Employe
Private mNom As String
Private mPrenom As String
Private mSalaire As Double
Private mDateEmbauche As Date
Private mManager As Personne
Private mActif As Boolean

' Propriété Nom avec validation
Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(valeur As String)
    ' Validation : pas de nom vide
    If Len(Trim(valeur)) > 0 Then
        mNom = Trim(valeur)
    Else
        Err.Raise 5, , "Le nom ne peut pas être vide"
    End If
End Property

' Propriété Prénom
Public Property Get Prenom() As String
    Prenom = mPrenom
End Property

Public Property Let Prenom(valeur As String)
    If Len(Trim(valeur)) > 0 Then
        mPrenom = Trim(valeur)
    Else
        Err.Raise 5, , "Le prénom ne peut pas être vide"
    End If
End Property

' Propriété Salaire avec validation
Public Property Get Salaire() As Double
    Salaire = mSalaire
End Property

Public Property Let Salaire(valeur As Double)
    If valeur >= 0 Then
        mSalaire = valeur
    Else
        Err.Raise 5, , "Le salaire ne peut pas être négatif"
    End If
End Property

' Propriété Date d'embauche
Public Property Get DateEmbauche() As Date
    DateEmbauche = mDateEmbauche
End Property

Public Property Let DateEmbauche(valeur As Date)
    If valeur <= Date Then  ' Date d'embauche ne peut pas être dans le futur
        mDateEmbauche = valeur
    Else
        Err.Raise 5, , "La date d'embauche ne peut pas être dans le futur"
    End If
End Property

' Propriété Manager (objet)
Public Property Get Manager() As Personne
    Set Manager = mManager
End Property

Public Property Set Manager(valeur As Personne)
    Set mManager = valeur
End Property

' Propriété Actif (booléen)
Public Property Get Actif() As Boolean
    Actif = mActif
End Property

Public Property Let Actif(valeur As Boolean)
    mActif = valeur
End Property

' Propriété calculée (lecture seule) - Ancienneté en années
Public Property Get AncienneteAnnees() As Integer
    If mDateEmbauche > 0 Then
        AncienneteAnnees = DateDiff("yyyy", mDateEmbauche, Date)
    Else
        AncienneteAnnees = 0
    End If
End Property

' Propriété calculée (lecture seule) - Nom complet
Public Property Get NomComplet() As String
    NomComplet = mPrenom & " " & mNom
End Property
```

### Propriétés en lecture seule

Certaines propriétés ne doivent pas être modifiables directement. On crée alors seulement un `Property Get` :

```vba
' Propriété calculée - ne peut pas être modifiée directement
Public Property Get SalaireAnnuel() As Double
    SalaireAnnuel = mSalaire * 12
End Property

' Propriété système - générée automatiquement
Public Property Get NumeroEmploye() As String
    If mNumeroEmploye = "" Then
        mNumeroEmploye = "EMP" & Format(Now, "yyyymmddhhnnss")
    End If
    NumeroEmploye = mNumeroEmploye
End Property
```

## Les Méthodes

### Qu'est-ce qu'une méthode ?

Une **méthode** définit ce qu'un objet peut **faire**. C'est une procédure (`Sub`) ou une fonction (`Function`) qui appartient à la classe et qui peut manipuler les données de l'objet.

### Types de méthodes

#### 1. Méthodes d'action (Sub)
Effectuent une action sans retourner de valeur.

```vba
Public Sub Promouvoir(nouveauTitre As String, augmentationPourcent As Double)
    mTitre = nouveauTitre
    mSalaire = mSalaire * (1 + augmentationPourcent / 100)

    ' Déclencher un événement (voir plus bas)
    RaiseEvent Promotion(Me, nouveauTitre, augmentationPourcent)
End Sub

Public Sub ChangerManager(nouveauManager As Personne)
    Dim ancienManager As Personne
    Set ancienManager = mManager
    Set mManager = nouveauManager

    ' Log du changement
    Debug.Print Me.NomComplet & " a changé de manager"
End Sub
```

#### 2. Méthodes de calcul (Function)
Effectuent un calcul et retournent une valeur.

```vba
Public Function CalculerPrimeAnciennete() As Double
    Dim annees As Integer
    annees = Me.AncienneteAnnees

    Select Case annees
        Case 0 To 2
            CalculerPrimeAnciennete = 0
        Case 3 To 5
            CalculerPrimeAnciennete = mSalaire * 0.05
        Case 6 To 10
            CalculerPrimeAnciennete = mSalaire * 0.1
        Case Else
            CalculerPrimeAnciennete = mSalaire * 0.15
    End Select
End Function

Public Function EstEligibleFormation() As Boolean
    EstEligibleFormation = (mActif = True) And (Me.AncienneteAnnees >= 1)
End Function
```

#### 3. Méthodes d'information
Affichent ou retournent des informations sur l'objet.

```vba
Public Sub AfficherFichePaie()
    Debug.Print "=== FICHE DE PAIE ==="
    Debug.Print "Employé : " & Me.NomComplet
    Debug.Print "Salaire mensuel : " & Format(mSalaire, "#,##0.00") & " €"
    Debug.Print "Ancienneté : " & Me.AncienneteAnnees & " ans"
    Debug.Print "Prime ancienneté : " & Format(Me.CalculerPrimeAnciennete(), "#,##0.00") & " €"
    Debug.Print "========================"
End Sub

Public Function VersChaine() As String
    VersChaine = Me.NomComplet & " (" & Me.AncienneteAnnees & " ans d'ancienneté)"
End Function
```

### Méthodes avec paramètres optionnels

```vba
Public Sub AugmenterSalaire(Optional pourcentage As Double = 5, Optional motif As String = "Augmentation annuelle")
    Dim ancienSalaire As Double
    ancienSalaire = mSalaire

    mSalaire = mSalaire * (1 + pourcentage / 100)

    ' Log de l'augmentation
    Debug.Print Me.NomComplet & " : " & motif & " - Salaire de " & _
                Format(ancienSalaire, "#,##0") & "€ à " & Format(mSalaire, "#,##0") & "€"
End Sub
```

## Les Événements

### Qu'est-ce qu'un événement ?

Un **événement** est un mécanisme qui permet à un objet de **notifier** d'autres parties du programme qu'une situation particulière s'est produite. C'est comme un "signal" que l'objet envoie.

### Déclaration d'événements

Les événements se déclarent au début du module de classe avec le mot-clé `Event` :

```vba
' Déclaration des événements au début de la classe
Public Event Promotion(employe As Employe, nouveauTitre As String, augmentation As Double)
Public Event ChangementSalaire(employe As Employe, ancienSalaire As Double, nouveauSalaire As Double)
Public Event Depart(employe As Employe, dateDepart As Date, motif As String)
Public Event AnniversaireEmbauche(employe As Employe, nombreAnnees As Integer)
```

### Déclencher des événements (RaiseEvent)

On utilise `RaiseEvent` pour déclencher un événement depuis une méthode :

```vba
Public Sub ModifierSalaire(nouveauSalaire As Double)
    Dim ancienSalaire As Double
    ancienSalaire = mSalaire

    ' Validation
    If nouveauSalaire >= 0 Then
        mSalaire = nouveauSalaire

        ' Déclencher l'événement
        RaiseEvent ChangementSalaire(Me, ancienSalaire, nouveauSalaire)
    Else
        Err.Raise 5, , "Le salaire ne peut pas être négatif"
    End If
End Sub

Public Sub Demissionner(Optional motif As String = "Démission")
    mActif = False

    ' Déclencher l'événement de départ
    RaiseEvent Depart(Me, Date, motif)
End Sub

Public Sub VerifierAnniversaire()
    Dim aujourd As Date
    aujourd = Date

    ' Vérifier si c'est l'anniversaire d'embauche
    If Month(aujourd) = Month(mDateEmbauche) And Day(aujourd) = Day(mDateEmbauche) Then
        RaiseEvent AnniversaireEmbauche(Me, Me.AncienneteAnnees)
    End If
End Sub
```

### Écouter les événements

Pour écouter les événements d'un objet, il faut le déclarer avec `WithEvents` :

```vba
' Dans un module standard ou une classe
Public WithEvents monEmploye As Employe

Sub CreerEmploye()
    Set monEmploye = New Employe

    monEmploye.Nom = "Martin"
    monEmploye.Prenom = "Paul"
    monEmploye.Salaire = 3000
    monEmploye.DateEmbauche = #1/15/2020#
End Sub

' Gestionnaire d'événement - se déclenche automatiquement
Private Sub monEmploye_ChangementSalaire(employe As Employe, ancienSalaire As Double, nouveauSalaire As Double)
    MsgBox "Changement de salaire pour " & employe.NomComplet & vbCrLf & _
           "Ancien : " & Format(ancienSalaire, "#,##0") & "€" & vbCrLf & _
           "Nouveau : " & Format(nouveauSalaire, "#,##0") & "€"
End Sub

Private Sub monEmploye_Promotion(employe As Employe, nouveauTitre As String, augmentation As Double)
    MsgBox "Félicitations ! " & employe.NomComplet & " a été promu(e) : " & nouveauTitre & vbCrLf & _
           "Augmentation : " & augmentation & "%"
End Sub

Private Sub monEmploye_Depart(employe As Employe, dateDepart As Date, motif As String)
    MsgBox "Départ de " & employe.NomComplet & " le " & Format(dateDepart, "dd/mm/yyyy") & vbCrLf & _
           "Motif : " & motif
End Sub
```

## Exemple complet d'utilisation

```vba
Sub TestEmployeComplet()
    ' Créer un employé avec événements
    Dim emp As New Employe

    ' Configuration initiale
    emp.Nom = "Dubois"
    emp.Prenom = "Claire"
    emp.Salaire = 2800
    emp.DateEmbauche = #3/1/2021#
    emp.Actif = True

    ' Afficher les informations
    emp.AfficherFichePaie

    ' Tester les méthodes
    Debug.Print "Éligible formation : " & emp.EstEligibleFormation()
    Debug.Print "Prime ancienneté : " & Format(emp.CalculerPrimeAnciennete(), "#,##0.00") & "€"

    ' Augmentation
    emp.AugmenterSalaire 8, "Excellente performance"

    ' Promouvoir (déclenche un événement si écouté)
    emp.Promouvoir "Chef de projet", 15

    ' Nouvelle fiche de paie
    emp.AfficherFichePaie
End Sub
```

## Bonnes pratiques

### Pour les propriétés
- **Validation** : Toujours valider les données dans les `Property Let/Set`
- **Nommage** : Utiliser des noms clairs et explicites
- **Cohérence** : Si vous créez un `Property Get`, réfléchissez si un `Property Let/Set` est nécessaire
- **Propriétés calculées** : Utilisez des propriétés en lecture seule pour les valeurs calculées

### Pour les méthodes
- **Responsabilité unique** : Chaque méthode doit avoir une responsabilité claire
- **Paramètres** : Utilisez des paramètres optionnels quand approprié
- **Nommage** : Commencez par un verbe d'action (Calculer, Afficher, Modifier, etc.)
- **Documentation** : Commentez les méthodes complexes

### Pour les événements
- **Nommage** : Utilisez des noms qui décrivent ce qui s'est passé
- **Moment** : Déclenchez les événements au bon moment (après validation et modification)
- **Information** : Passez les informations utiles dans les paramètres de l'événement
- **Performance** : N'abusez pas des événements pour éviter la surcharge

## Comparaison avec Excel

Excel utilise massivement ces concepts :

```vba
' Propriétés Excel
Debug.Print ActiveSheet.Name        ' Property Get
ActiveSheet.Name = "Nouveau nom"    ' Property Let

' Méthodes Excel
ActiveSheet.Calculate               ' Sub (méthode d'action)
Debug.Print ActiveSheet.UsedRange.Count  ' Function (méthode de calcul)

' Événements Excel (dans le module de la feuille)
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Se déclenche quand une cellule change
End Sub
```

Cette approche structurée avec propriétés, méthodes et événements rend vos objets VBA aussi professionnels et faciles à utiliser que les objets intégrés d'Excel.

⏭️
