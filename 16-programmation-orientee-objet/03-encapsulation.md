🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.3. Encapsulation

## Qu'est-ce que l'encapsulation ?

L'**encapsulation** est l'un des concepts les plus importants de la programmation orientée objet. C'est l'art de **cacher** les détails internes d'un objet et de ne rendre accessible que ce qui est nécessaire pour l'utiliser.

**Analogie simple :**
Pensez à votre voiture. Pour la conduire, vous avez besoin :
- **Interface publique** : volant, pédales, levier de vitesse (ce que vous pouvez utiliser)
- **Détails cachés** : injection, compression, alternateur (ce qui est caché sous le capot)

Vous n'avez pas besoin de comprendre le fonctionnement interne du moteur pour conduire. L'encapsulation fonctionne de la même manière avec vos objets VBA.

## Pourquoi l'encapsulation est-elle importante ?

### 1. Sécurité des données
Elle empêche la modification accidentelle ou malveillante des données importantes de votre objet.

### 2. Facilité d'utilisation
L'utilisateur de votre classe n'a pas besoin de connaître les détails complexes de son fonctionnement.

### 3. Évolutivité
Vous pouvez modifier l'implémentation interne sans affecter le code qui utilise votre classe.

### 4. Validation automatique
Vous contrôlez exactement comment les données peuvent être modifiées.

## Les niveaux d'accès en VBA

### Private (Privé)
**Accessible uniquement depuis l'intérieur de la classe**

```vba
Private mSoldeCompte As Double      ' Variable privée
Private mNumeroCompte As String     ' Variable privée

Private Sub CalculerInterets()      ' Méthode privée
    ' Cette méthode ne peut être appelée que depuis l'intérieur de la classe
    mSoldeCompte = mSoldeCompte * 1.02
End Sub
```

### Public (Public)
**Accessible depuis n'importe où**

```vba
Public Property Get Solde() As Double    ' Propriété publique
    Solde = mSoldeCompte
End Property

Public Sub Deposer(montant As Double)    ' Méthode publique
    ' Cette méthode peut être appelée depuis l'extérieur
    If montant > 0 Then
        mSoldeCompte = mSoldeCompte + montant
    End If
End Sub
```

### Friend (Ami)
**Accessible uniquement depuis le même projet VBA** (rarement utilisé en pratique)

## Exemple complet : Classe CompteBancaire

Voici un exemple qui illustre parfaitement l'encapsulation :

```vba
' Module de classe : CompteBancaire
Option Explicit

' ========== DONNÉES PRIVÉES ==========
' Ces variables ne sont accessibles que depuis l'intérieur de la classe
Private mNumeroCompte As String
Private mTitulaire As String
Private mSolde As Double
Private mPlafondDecouvert As Double
Private mDateOuverture As Date
Private mNombreTransactions As Long
Private mCompteActif As Boolean

' ========== PROPRIÉTÉS PUBLIQUES ==========

' Numéro de compte (lecture seule)
Public Property Get NumeroCompte() As String
    NumeroCompte = mNumeroCompte
End Property

' Titulaire avec validation
Public Property Get Titulaire() As String
    Titulaire = mTitulaire
End Property

Public Property Let Titulaire(valeur As String)
    ' Validation : nom non vide
    If Len(Trim(valeur)) > 0 Then
        mTitulaire = Trim(valeur)
    Else
        Err.Raise 5, , "Le nom du titulaire ne peut pas être vide"
    End If
End Property

' Solde (lecture seule - ne peut être modifié directement)
Public Property Get Solde() As Double
    Solde = mSolde
End Property

' Plafond de découvert avec validation
Public Property Get PlafondDecouvert() As Double
    PlafondDecouvert = mPlafondDecouvert
End Property

Public Property Let PlafondDecouvert(valeur As Double)
    ' Validation : découvert ne peut pas être positif
    If valeur <= 0 Then
        mPlafondDecouvert = valeur
    Else
        Err.Raise 5, , "Le plafond de découvert doit être négatif ou nul"
    End If
End Property

' Date d'ouverture (lecture seule)
Public Property Get DateOuverture() As Date
    DateOuverture = mDateOuverture
End Property

' Nombre de transactions (lecture seule)
Public Property Get NombreTransactions() As Long
    NombreTransactions = mNombreTransactions
End Property

' État du compte
Public Property Get EstActif() As Boolean
    EstActif = mCompteActif
End Property

' ========== MÉTHODES PUBLIQUES ==========

' Initialisation du compte
Public Sub Initialiser(numeroCompte As String, titulaire As String, Optional soldeInitial As Double = 0)
    ' Validation du numéro de compte
    If Len(Trim(numeroCompte)) = 0 Then
        Err.Raise 5, , "Le numéro de compte ne peut pas être vide"
    End If

    ' Validation du titulaire
    If Len(Trim(titulaire)) = 0 Then
        Err.Raise 5, , "Le nom du titulaire ne peut pas être vide"
    End If

    ' Initialisation
    mNumeroCompte = Trim(numeroCompte)
    mTitulaire = Trim(titulaire)
    mSolde = soldeInitial
    mPlafondDecouvert = -500  ' Découvert autorisé de 500€
    mDateOuverture = Date
    mNombreTransactions = 0
    mCompteActif = True

    ' Log de création
    Call Me.EcrireLog("Compte créé - Solde initial : " & Format(soldeInitial, "#,##0.00") & "€")
End Sub

' Dépôt d'argent
Public Sub Deposer(montant As Double, Optional motif As String = "Dépôt")
    ' Validation
    If Not mCompteActif Then
        Err.Raise 5, , "Opération impossible : compte inactif"
    End If

    If montant <= 0 Then
        Err.Raise 5, , "Le montant du dépôt doit être positif"
    End If

    ' Effectuer l'opération
    mSolde = mSolde + montant
    mNombreTransactions = mNombreTransactions + 1

    ' Log de l'opération
    Call Me.EcrireLog("DÉPÔT : +" & Format(montant, "#,##0.00") & "€ - " & motif & " - Solde : " & Format(mSolde, "#,##0.00") & "€")
End Sub

' Retrait d'argent
Public Function Retirer(montant As Double, Optional motif As String = "Retrait") As Boolean
    ' Validation
    If Not mCompteActif Then
        Err.Raise 5, , "Opération impossible : compte inactif"
        Retirer = False
        Exit Function
    End If

    If montant <= 0 Then
        Err.Raise 5, , "Le montant du retrait doit être positif"
        Retirer = False
        Exit Function
    End If

    ' Vérifier si le retrait est possible (ne pas dépasser le découvert autorisé)
    If (mSolde - montant) < mPlafondDecouvert Then
        ' Retrait impossible
        Call Me.EcrireLog("RETRAIT REFUSÉ : -" & Format(montant, "#,##0.00") & "€ - " & motif & " - Découvert insuffisant")
        Retirer = False
    Else
        ' Retrait possible
        mSolde = mSolde - montant
        mNombreTransactions = mNombreTransactions + 1

        Call Me.EcrireLog("RETRAIT : -" & Format(montant, "#,##0.00") & "€ - " & motif & " - Solde : " & Format(mSolde, "#,##0.00") & "€")
        Retirer = True
    End If
End Function

' Virement vers un autre compte
Public Function Virement(montant As Double, compteDestination As CompteBancaire, Optional motif As String = "Virement") As Boolean
    ' Validation
    If compteDestination Is Nothing Then
        Err.Raise 5, , "Le compte de destination est requis"
        Virement = False
        Exit Function
    End If

    If compteDestination.NumeroCompte = Me.NumeroCompte Then
        Err.Raise 5, , "Impossible de faire un virement vers le même compte"
        Virement = False
        Exit Function
    End If

    ' Effectuer le retrait
    If Me.Retirer(montant, "Virement vers " & compteDestination.NumeroCompte) Then
        ' Effectuer le dépôt sur le compte destination
        compteDestination.Deposer montant, "Virement de " & Me.NumeroCompte
        Virement = True
    Else
        Virement = False
    End If
End Function

' Fermeture du compte
Public Sub FermerCompte()
    If mSolde <> 0 Then
        Err.Raise 5, , "Impossible de fermer un compte avec un solde non nul"
    Else
        mCompteActif = False
        Call Me.EcrireLog("COMPTE FERMÉ")
    End If
End Sub

' Affichage du relevé
Public Sub AfficherReleve()
    Debug.Print "========== RELEVÉ DE COMPTE =========="
    Debug.Print "Numéro : " & mNumeroCompte
    Debug.Print "Titulaire : " & mTitulaire
    Debug.Print "Date d'ouverture : " & Format(mDateOuverture, "dd/mm/yyyy")
    Debug.Print "Solde actuel : " & Format(mSolde, "#,##0.00") & "€"
    Debug.Print "Découvert autorisé : " & Format(mPlafondDecouvert, "#,##0.00") & "€"
    Debug.Print "Nombre de transactions : " & mNombreTransactions
    Debug.Print "Statut : " & IIf(mCompteActif, "Actif", "Fermé")

    ' Afficher l'état du compte
    If mSolde < 0 Then
        Debug.Print "⚠️ ATTENTION : Compte à découvert"
    ElseIf mSolde < 100 Then
        Debug.Print "ℹ️ INFO : Solde faible"
    End If

    Debug.Print "======================================"
End Sub

' ========== MÉTHODES PRIVÉES ==========
' Ces méthodes ne peuvent être appelées que depuis l'intérieur de la classe

Private Sub EcrireLog(message As String)
    ' Méthode privée pour écrire dans le log
    ' L'utilisateur de la classe ne peut pas l'appeler directement
    Debug.Print Format(Now, "dd/mm/yyyy hh:nn:ss") & " [" & mNumeroCompte & "] " & message
End Sub

Private Function ValiderMontant(montant As Double) As Boolean
    ' Méthode privée de validation
    ValiderMontant = (montant > 0 And montant < 1000000)  ' Limite à 1 million
End Function

Private Sub VerifierLimites()
    ' Méthode privée pour vérifier les limites du compte
    If mNombreTransactions > 1000 Then
        Debug.Print "Attention : Plus de 1000 transactions sur ce compte"
    End If
End Sub
```

## Utilisation de la classe encapsulée

```vba
Sub TestEncapsulation()
    ' Création du compte
    Dim compte As New CompteBancaire
    compte.Initialiser "FR123456789", "Jean Dupont", 1000

    ' ========== CE QUI FONCTIONNE (Interface publique) ==========

    ' Lecture des propriétés publiques
    Debug.Print "Numéro : " & compte.NumeroCompte
    Debug.Print "Titulaire : " & compte.Titulaire
    Debug.Print "Solde : " & compte.Solde

    ' Utilisation des méthodes publiques
    compte.Deposer 500, "Salaire"
    compte.Retirer 200, "Courses"
    compte.AfficherReleve

    ' Modification des propriétés autorisées
    compte.PlafondDecouvert = -1000

    ' ========== CE QUI NE FONCTIONNE PAS (Données privées) ==========

    ' Ces lignes provoqueraient des erreurs de compilation :
    ' compte.mSolde = 999999           ' ❌ Variable privée
    ' compte.mNumeroCompte = "HACK"    ' ❌ Variable privée
    ' compte.EcrireLog("Test")         ' ❌ Méthode privée

    ' Ces propriétés ne peuvent pas être modifiées :
    ' compte.Solde = 5000              ' ❌ Pas de Property Let
    ' compte.NumeroCompte = "Nouveau"  ' ❌ Lecture seule

End Sub
```

## Avantages de cette approche encapsulée

### 1. Protection des données critiques
```vba
' ❌ Sans encapsulation (dangereux)
' L'utilisateur pourrait faire :
' compte.Solde = -999999999
' compte.NumeroCompte = ""

' ✅ Avec encapsulation (sécurisé)
' L'utilisateur doit passer par les méthodes contrôlées :
compte.Retirer(200)  ' Vérifie les limites automatiquement
```

### 2. Validation automatique
```vba
' Toute modification passe par la validation
compte.Titulaire = ""  ' ❌ Erreur automatique
compte.PlafondDecouvert = 1000  ' ❌ Erreur automatique (doit être négatif)
```

### 3. Traçabilité
```vba
' Toutes les opérations sont automatiquement enregistrées
' via la méthode privée EcrireLog
```

### 4. Évolution facile
Si vous voulez changer la façon dont les intérêts sont calculés, vous modifiez seulement le code interne :

```vba
' Ancienne version (dans une méthode privée)
Private Sub CalculerInterets()
    mSolde = mSolde * 1.02
End Sub

' Nouvelle version (personne ne le remarque à l'extérieur)
Private Sub CalculerInterets()
    If mSolde > 0 Then
        mSolde = mSolde * 1.025  ' Taux plus avantageux
    End If
End Sub
```

## Techniques d'encapsulation avancées

### 1. Propriétés calculées
```vba
' Propriété calculée (pas de variable membre correspondante)
Public Property Get SoldeFormate() As String
    SoldeFormate = Format(mSolde, "#,##0.00") & "€"
End Property

Public Property Get EstADecouvert() As Boolean
    EstADecouvert = (mSolde < 0)
End Property
```

### 2. Validation dans les propriétés
```vba
Public Property Let PlafondDecouvert(valeur As Double)
    ' Validation complexe
    If valeur > 0 Then
        Err.Raise 5, , "Le découvert ne peut pas être positif"
    ElseIf valeur < -10000 Then
        Err.Raise 5, , "Découvert maximum : -10 000€"
    Else
        ' Log du changement
        If mPlafondDecouvert <> valeur Then
            Call Me.EcrireLog("Changement découvert : " & Format(mPlafondDecouvert, "#,##0") & "€ -> " & Format(valeur, "#,##0") & "€")
        End If
        mPlafondDecouvert = valeur
    End If
End Property
```

### 3. Méthodes d'aide privées
```vba
Private Function EstOperationAutorisee(montant As Double) As Boolean
    ' Logique complexe centralisée
    EstOperationAutorisee = mCompteActif And _
                           montant > 0 And _
                           montant < 50000 And _
                           mNombreTransactions < 1000
End Function

Public Function Retirer(montant As Double, Optional motif As String = "Retrait") As Boolean
    If Me.EstOperationAutorisee(montant) Then
        ' ... logique de retrait
    Else
        Retirer = False
    End If
End Function
```

## Comparaison : Avec et sans encapsulation

### ❌ Sans encapsulation (approche procédurale)
```vba
' Variables globales - dangereuses !
Public soldeCompte As Double
Public numeroCompte As String

Sub RetirerArgent(montant As Double)
    soldeCompte = soldeCompte - montant  ' Aucune validation !
End Sub
```

**Problèmes :**
- N'importe qui peut modifier `soldeCompte` directement
- Aucune validation
- Pas de traçabilité
- Code difficile à maintenir

### ✅ Avec encapsulation
```vba
' Classe CompteBancaire avec encapsulation
' - Données privées protégées
' - Validation automatique
' - Interface claire et contrôlée
' - Évolution facile
```

## Bonnes pratiques d'encapsulation

### 1. Règle générale
- **Tout est Private par défaut**
- Ne rendez Public que ce qui doit vraiment être accessible

### 2. Nommage des variables privées
```vba
' Convention : préfixe 'm' pour 'membre'
Private mNom As String          ' ✅ Bon
Private nom As String           ' ❌ Pas clair
Private m_nom As String         ' ✅ Alternative acceptée
```

### 3. Propriétés vs variables publiques
```vba
' ❌ Éviter les variables publiques
Public Nom As String

' ✅ Préférer les propriétés avec validation
Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(valeur As String)
    If Len(Trim(valeur)) > 0 Then
        mNom = Trim(valeur)
    End If
End Property
```

### 4. Documentation
```vba
' Documentez l'interface publique
Public Sub Deposer(montant As Double, Optional motif As String = "Dépôt")
    ' Effectue un dépôt sur le compte
    ' Paramètres :
    '   montant : Montant à déposer (doit être > 0)
    '   motif   : Description de l'opération
    ' Exceptions :
    '   Erreur 5 si montant <= 0 ou compte inactif
```

L'encapsulation est la fondation d'un code robuste et maintenable. Elle transforme vos données brutes en objets intelligents qui se protègent eux-mêmes et offrent une interface claire et sécurisée.

⏭️
