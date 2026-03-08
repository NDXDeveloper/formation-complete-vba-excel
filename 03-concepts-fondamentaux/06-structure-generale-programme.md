🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.6 Structure générale d'un programme

## Introduction

La structure d'un programme VBA, c'est comme l'architecture d'une maison. Vous ne construisez pas au hasard : vous commencez par les fondations, puis les murs, et enfin le toit. En programmation, c'est pareil ! Il existe un ordre logique et des bonnes pratiques qui rendent votre code solide, lisible et facile à maintenir. Cette section vous donnera les clés pour organiser vos programmes comme un professionnel.

## Qu'est-ce que la structure d'un programme ?

### Définition simple

**La structure d'un programme** = L'organisation logique de votre code pour qu'il soit clair, efficace et maintenable

**Analogies pratiques :**
- **Architecture de maison** : Fondations → Murs → Toit
- **Livre** : Table des matières → Chapitres → Paragraphes
- **Entreprise** : Direction → Services → Employés
- **Recette** : Ingrédients → Préparation → Cuisson

### Pourquoi la structure est-elle importante ?

**Lisibilité :**
```vba
' MAUVAISE structure : tout mélangé
Sub CalculerTout()
    Dim x As Integer, y As String, Total As Double
    x = Range("A1").Value: y = "Bonjour": Total = 0
    For i = 1 To 10: Total = Total + Cells(i, 1).Value: Next i
    Range("B1").Value = Total: MsgBox y & " " & x
End Sub

' BONNE structure : organisée et claire
Sub CalculerTotal()
    ' === DÉCLARATIONS ===
    Dim Total As Double
    Dim i As Integer

    ' === INITIALISATION ===
    Total = 0

    ' === TRAITEMENT PRINCIPAL ===
    For i = 1 To 10
        Total = Total + Cells(i, 1).Value
    Next i

    ' === AFFICHAGE RÉSULTAT ===
    Range("B1").Value = Total
End Sub
```

**Maintenance :**
- **Modifications faciles** : Retrouver rapidement ce qu'il faut changer
- **Ajouts simples** : Savoir où insérer du nouveau code
- **Débogage efficace** : Localiser les problèmes rapidement

**Réutilisabilité :**
- **Modules clairs** : Réutiliser des parties dans d'autres projets
- **Fonctions spécialisées** : Éviter la duplication de code
- **Standards** : Code compréhensible par toute l'équipe

## Structure au niveau du module

### Organisation générale d'un module

**Ordre recommandé de haut en bas :**
```vba
'******************************************************************************
' MODULE : ModuleCalculsFinanciers
' DESCRIPTION : Fonctions de calcul pour la gestion financière
' AUTEUR : Votre Nom
' DATE : 15/01/2024
'******************************************************************************

' ===== 1. DIRECTIVES DE COMPILATION =====
Option Explicit  
Option Compare Text  

' ===== 2. CONSTANTES PUBLIQUES =====
Public Const TAUX_TVA_STANDARD As Double = 0.20  
Public Const DEVISE_DEFAUT As String = "EUR"  

' ===== 3. CONSTANTES PRIVÉES =====
Const SEUIL_ALERTE As Double = 1000.0  
Const MESSAGE_ERREUR As String = "Erreur de calcul"  

' ===== 4. VARIABLES PUBLIQUES =====
Public ModeDebug As Boolean

' ===== 5. VARIABLES PRIVÉES =====
Dim CompteurCalculs As Long  
Dim DerniereErreur As String  

' ===== 6. FONCTIONS PUBLIQUES =====
Public Function CalculerTTC(PrixHT As Double) As Double
    ' Code de la fonction
End Function

' ===== 7. PROCÉDURES PUBLIQUES =====
Public Sub InitialiserModule()
    ' Code d'initialisation
End Sub

' ===== 8. FONCTIONS PRIVÉES =====
Private Function ValiderMontant(Montant As Double) As Boolean
    ' Code de validation
End Function

' ===== 9. PROCÉDURES PRIVÉES =====
Private Sub LoggerErreur(MessageErreur As String)
    ' Code de logging
End Sub
```

### Directives de compilation

**Option Explicit :**
```vba
Option Explicit    ' Force la déclaration de toutes les variables
```

**Option Compare :**
```vba
Option Compare Text        ' Comparaison de texte insensible à la casse  
Option Compare Binary      ' Comparaison binaire (par défaut)  
```

**Option Base :**
```vba
Option Base 1             ' Les tableaux commencent à 1 (rare)  
Option Base 0             ' Les tableaux commencent à 0 (par défaut)  
```

### Déclarations globales

**Constantes au niveau module :**
```vba
' Constantes utilisables dans tout le module
Const MAX_TENTATIVES As Integer = 3  
Const REPERTOIRE_TEMP As String = "C:\Temp\"  
Const VERSION_MODULE As String = "1.2.0"  
```

**Variables partagées :**
```vba
' Variables accessibles à toutes les procédures du module
Dim CompteurGlobal As Long  
Dim ConfigurationActive As String  
Dim TableauResultats() As Double  
```

## Structure d'une procédure

### Anatomie d'une procédure bien structurée

```vba
'******************************************************************************
' PROCÉDURE : CalculerRemiseClient
' DESCRIPTION : Calcule la remise selon le profil client et le montant
' PARAMÈTRES :
'   - TypeClient : "VIP", "Premium", "Standard"
'   - MontantCommande : Montant en euros
' RETOUR : Pourcentage de remise (0 à 1)
' AUTEUR : Votre Nom
' DATE : 15/01/2024
'******************************************************************************
Function CalculerRemiseClient(TypeClient As String, MontantCommande As Double) As Double

    ' ===== DÉCLARATIONS LOCALES =====
    Dim PourcentageRemise As Double
    Dim SeuilVIP As Double
    Dim SeuilPremium As Double

    ' ===== INITIALISATION =====
    PourcentageRemise = 0
    SeuilVIP = 5000
    SeuilPremium = 2000

    ' ===== VALIDATION DES PARAMÈTRES =====
    If MontantCommande <= 0 Then
        CalculerRemiseClient = 0
        Exit Function
    End If

    ' ===== TRAITEMENT PRINCIPAL =====
    Select Case UCase(TypeClient)
        Case "VIP"
            If MontantCommande >= SeuilVIP Then
                PourcentageRemise = 0.15        ' 15% pour VIP gros montant
            Else
                PourcentageRemise = 0.10        ' 10% pour VIP montant normal
            End If

        Case "PREMIUM"
            If MontantCommande >= SeuilPremium Then
                PourcentageRemise = 0.08        ' 8% pour Premium gros montant
            Else
                PourcentageRemise = 0.05        ' 5% pour Premium montant normal
            End If

        Case "STANDARD"
            If MontantCommande >= SeuilVIP Then
                PourcentageRemise = 0.03        ' 3% pour Standard très gros montant
            Else
                PourcentageRemise = 0           ' Pas de remise Standard
            End If

        Case Else
            PourcentageRemise = 0               ' Type client inconnu
    End Select

    ' ===== FINALISATION ET RETOUR =====
    CalculerRemiseClient = PourcentageRemise

End Function
```

### Sections d'une procédure

**1. En-tête documentaire :**
- Description de ce que fait la procédure
- Paramètres d'entrée et de sortie
- Informations d'auteur et de version

**2. Déclarations locales :**
- Variables utilisées uniquement dans cette procédure
- Groupement par type ou usage

**3. Initialisation :**
- Valeurs de départ des variables
- Configuration initiale

**4. Validation :**
- Vérification des paramètres d'entrée
- Gestion des cas d'erreur anticipés

**5. Traitement principal :**
- Logique métier de la procédure
- Algorithme principal

**6. Finalisation :**
- Nettoyage si nécessaire
- Préparation du retour

## Patterns de structure courants

### Pattern de validation

```vba
Sub TraiterCommande(NumeroCommande As Long)
    ' ===== VALIDATION =====
    ' Vérification des prérequis avant traitement
    If NumeroCommande <= 0 Then
        MsgBox "Numéro de commande invalide"
        Exit Sub
    End If

    If Not VerifierExistenceCommande(NumeroCommande) Then
        MsgBox "Commande inexistante"
        Exit Sub
    End If

    ' ===== TRAITEMENT =====
    ' Traitement principal seulement si validation OK
    ' ... code principal ...

End Sub
```

### Pattern d'initialisation

```vba
Sub GenererRapport()
    ' ===== INITIALISATION ENVIRONNEMENT =====
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ===== INITIALISATION VARIABLES =====
    Dim FeuilleRapport As Worksheet
    Dim LigneActuelle As Long
    Set FeuilleRapport = Worksheets.Add
    LigneActuelle = 1

    ' ===== TRAITEMENT =====
    ' ... génération du rapport ...

    ' ===== NETTOYAGE =====
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set FeuilleRapport = Nothing
End Sub
```

### Pattern de boucle structurée

```vba
Sub TraiterDonneesListe()
    ' ===== PRÉPARATION =====
    Dim i As Long
    Dim DerniereLigne As Long
    Dim CompteurErreurs As Long

    DerniereLigne = Range("A" & Rows.Count).End(xlUp).Row
    CompteurErreurs = 0

    ' ===== BOUCLE PRINCIPALE =====
    For i = 2 To DerniereLigne    ' Ligne 1 = en-têtes

        ' Validation de la ligne courante
        If ValiderLigne(i) Then
            ' Traitement de la ligne valide
            TraiterLigne i
        Else
            ' Gestion des erreurs
            CompteurErreurs = CompteurErreurs + 1
            MarquerLigneErreur i
        End If

    Next i

    ' ===== RAPPORT FINAL =====
    AfficherResumeTraitement DerniereLigne - 1, CompteurErreurs
End Sub
```

## Organisation par responsabilité

### Séparation des préoccupations

**Une procédure = une responsabilité :**
```vba
' MAUVAIS : fait trop de choses différentes
Sub ToutFaireEnUneSeuleFois()
    ' Validation des données
    If Range("A1").Value = "" Then Exit Sub

    ' Calculs financiers
    Dim Total As Double
    Total = Range("A1").Value * 1.2

    ' Formatage de l'affichage
    Range("B1").NumberFormat = "0.00 €"

    ' Sauvegarde du fichier
    ActiveWorkbook.Save

    ' Envoi par email
    ' ... code d'envoi email ...
End Sub

' BON : responsabilités séparées
Sub TraiterCommande()
    If Not ValiderDonnees() Then Exit Sub

    Dim Total As Double
    Total = CalculerTotal()

    FormaterAffichage Total
    SauvegarderDonnees
    EnvoyerNotification
End Sub

Function ValiderDonnees() As Boolean
    ' Uniquement la validation
End Function

Function CalculerTotal() As Double
    ' Uniquement les calculs
End Function

Sub FormaterAffichage(Montant As Double)
    ' Uniquement le formatage
End Sub
```

### Hiérarchie des fonctions

**Niveau 1 : Fonction principale (orchestrateur)**
```vba
Sub ProcessusCompletCommande()
    ' Vue d'ensemble du processus
    InitialiserEnvironnement

    If ValiderPrealables() Then
        TraiterToutesLesCommandes
        GenererRapportFinal
    End If

    NettoyerEnvironnement
End Sub
```

**Niveau 2 : Fonctions de section**
```vba
Sub TraiterToutesLesCommandes()
    Dim i As Long
    Dim NombreCommandes As Long

    NombreCommandes = ObtenirNombreCommandes()

    For i = 1 To NombreCommandes
        TraiterUneCommande i
    Next i
End Sub
```

**Niveau 3 : Fonctions de détail**
```vba
Sub TraiterUneCommande(NumeroLigne As Long)
    Dim Commande As Variant    ' Objet ou Type personnalisé

    Commande = LireCommande(NumeroLigne)

    If ValiderCommande(Commande) Then
        CalculerCommande Commande
        SauvegarderCommande Commande
    End If
End Sub
```

## Gestion des erreurs structurée

### Pattern de gestion d'erreurs

```vba
Sub ProcedureAvecGestionErreurs()
    ' ===== INITIALISATION =====
    On Error GoTo GestionErreur
    Dim ResultatOK As Boolean
    ResultatOK = False

    ' ===== TRAITEMENT PRINCIPAL =====
    ' ... code principal ...
    ResultatOK = True

    ' ===== SORTIE NORMALE =====
SortieNormale:
    ' Nettoyage commun
    Application.ScreenUpdating = True
    If ResultatOK Then
        MsgBox "Traitement réussi"
    End If
    Exit Sub

    ' ===== GESTION D'ERREUR =====
GestionErreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description
    Resume SortieNormale
End Sub
```

### Structure avec nettoyage garanti

```vba
Sub ProcedureAvecNettoyage()
    ' ===== DÉCLARATIONS =====
    Dim FeuilleTemp As Worksheet
    Dim AncienCalcul As XlCalculation

    ' ===== INITIALISATION =====
    On Error GoTo NettoyageEtSortie
    Set FeuilleTemp = Nothing
    AncienCalcul = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' ===== TRAITEMENT =====
    Set FeuilleTemp = Worksheets.Add
    ' ... traitement avec la feuille temporaire ...

    ' ===== NETTOYAGE ET SORTIE =====
NettoyageEtSortie:
    ' Code de nettoyage TOUJOURS exécuté
    Application.Calculation = AncienCalcul

    If Not FeuilleTemp Is Nothing Then
        Application.DisplayAlerts = False
        FeuilleTemp.Delete
        Application.DisplayAlerts = True
        Set FeuilleTemp = Nothing
    End If

    If Err.Number <> 0 Then
        MsgBox "Erreur : " & Err.Description
    End If
End Sub
```

## Standards de nommage

### Conventions pour les procédures

**Verbes d'action pour les procédures :**
```vba
Sub CalculerTotal()          ' Calcule quelque chose  
Sub AfficherResultat()       ' Affiche quelque chose  
Sub SauvegarderDonnees()     ' Sauvegarde quelque chose  
Sub ValiderSaisie()          ' Valide quelque chose  
Sub InitialiserModule()      ' Initialise quelque chose  
```

**Noms descriptifs pour les fonctions :**
```vba
Function ObtenirNombreClients() As Long           ' Retourne un nombre  
Function EstValide(Valeur As String) As Boolean   ' Retourne vrai/faux  
Function CalculerRemise(Montant As Double) As Double   ' Retourne un calcul  
```

### Préfixes significatifs

**Par type d'opération :**
```vba
' Validation
Function EstNumerique(Texte As String) As Boolean  
Function PeutEtreConverti(Valeur As Variant) As Boolean  

' Obtention de données
Function ObtenirDerniereLigne() As Long  
Function RecupererParametre(Nom As String) As String  

' Vérification d'état
Function ExisteFeuille(NomFeuille As String) As Boolean  
Function EstOuvert(NomFichier As String) As Boolean  

' Création/Génération
Sub CreerNouvelleCommande()  
Sub GenererRapportMensuel()  

' Nettoyage/Suppression
Sub SupprimerFichiersTemp()  
Sub NettoyerDonneesAnciennes()  
```

## Documentation de structure

### Plan du module

```vba
'******************************************************************************
' MODULE PLAN
'******************************************************************************
'
' FONCTIONS PUBLIQUES :
' =====================
' CalculerTTC(PrixHT) -> Double
'   Calcule le prix TTC à partir du prix HT
'
' ValiderCommande(Commande) -> Boolean
'   Vérifie qu'une commande est valide
'
' GenererFacture(NumCommande) -> String
'   Génère une facture et retourne le nom du fichier
'
' PROCÉDURES PUBLIQUES :
' ======================
' InitialiserModule()
'   Initialise les paramètres du module
'
' TraiterFichierCommandes(CheminFichier)
'   Traite un fichier de commandes en lot
'
' FONCTIONS PRIVÉES :
' ===================
' FormaterMontant(Montant) -> String
' ObtenirTauxTVA() -> Double
' LoggerAction(Message)
'
'******************************************************************************
```

### Diagramme de flux dans les commentaires

```vba
'******************************************************************************
' FLUX DE TRAITEMENT PRINCIPAL
'******************************************************************************
'
' TraiterCommandes()
'     │
'     ├── InitialiserEnvironnement()
'     │
'     ├── Pour chaque commande :
'     │   ├── LireCommande()
'     │   ├── ValiderCommande()
'     │   │   ├── EstMontantValide()
'     │   │   └── EstClientExistant()
'     │   ├── CalculerTotaux()
'     │   │   ├── CalculerSousTotal()
'     │   │   ├── CalculerTVA()
'     │   │   └── CalculerRemise()
'     │   └── SauvegarderCommande()
'     │
'     ├── GenererRapport()
'     └── NettoyerEnvironnement()
'
'******************************************************************************
```

## Modularité et réutilisabilité

### Fonctions utilitaires

**Module d'utilitaires générales :**
```vba
'******************************************************************************
' MODULE : ModuleUtilitaires
' DESCRIPTION : Fonctions utilitaires réutilisables
'******************************************************************************

' Validation
Public Function EstCelluleVide(Cellule As Range) As Boolean
    EstCelluleVide = (Cellule.Value = "" Or IsEmpty(Cellule.Value))
End Function

' Conversion
Public Function ConvertirEnNombre(Texte As String) As Double
    If IsNumeric(Texte) Then
        ConvertirEnNombre = CDbl(Texte)
    Else
        ConvertirEnNombre = 0
    End If
End Function

' Formatage
Public Function FormaterMontantEuro(Montant As Double) As String
    FormaterMontantEuro = Format(Montant, "0.00") & " €"
End Function
```

### Paramétrage centralisé

**Module de configuration :**
```vba
'******************************************************************************
' MODULE : ModuleConfiguration
' DESCRIPTION : Paramètres centralisés de l'application
'******************************************************************************

' Paramètres métier
Public Const TAUX_TVA As Double = 0.20  
Public Const SEUIL_FRANCO_PORT As Double = 100.0  
Public Const REMISE_MAX As Double = 0.30  

' Paramètres techniques
Public Const REPERTOIRE_EXPORT As String = "C:\Exports\"  
Public Const FORMAT_DATE As String = "dd/mm/yyyy"  
Public Const NOM_FEUILLE_CONFIG As String = "Parametres"  

' Fonction de lecture dynamique
Public Function LireParametre(NomParametre As String) As Variant
    On Error GoTo ValeurParDefaut
    LireParametre = Worksheets(NOM_FEUILLE_CONFIG).Range(NomParametre).Value
    Exit Function

ValeurParDefaut:
    ' Valeurs par défaut en cas d'erreur
    Select Case NomParametre
        Case "TauxTVA": LireParametre = TAUX_TVA
        Case "SeuilFrancoPort": LireParametre = SEUIL_FRANCO_PORT
        Case Else: LireParametre = Empty
    End Select
End Function
```

## Optimisation de structure

### Éviter la duplication

**AVANT : Code dupliqué**
```vba
Sub TraiterClientsVIP()
    Dim i As Long
    Application.ScreenUpdating = False
    For i = 1 To 100
        If Cells(i, 3).Value = "VIP" Then
            Cells(i, 4).Value = Cells(i, 2).Value * 0.9
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

Sub TraiterClientsPremium()
    Dim i As Long
    Application.ScreenUpdating = False
    For i = 1 To 100
        If Cells(i, 3).Value = "Premium" Then
            Cells(i, 4).Value = Cells(i, 2).Value * 0.95
        End If
    Next i
    Application.ScreenUpdating = True
End Sub
```

**APRÈS : Code factorisé**
```vba
Sub TraiterClientsVIP()
    TraiterClients "VIP", 0.9
End Sub

Sub TraiterClientsPremium()
    TraiterClients "Premium", 0.95
End Sub

Private Sub TraiterClients(TypeClient As String, FacteurRemise As Double)
    Dim i As Long

    Application.ScreenUpdating = False

    For i = 1 To 100
        If Cells(i, 3).Value = TypeClient Then
            Cells(i, 4).Value = Cells(i, 2).Value * FacteurRemise
        End If
    Next i

    Application.ScreenUpdating = True
End Sub
```

### Optimisation des performances

**Structure optimisée pour les performances :**
```vba
Sub TraiterGrosVolume()
    ' ===== OPTIMISATION DÉBUT =====
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ===== LECTURE EN BLOC =====
    ' Plus rapide que cellule par cellule
    Dim DonneesSource As Variant
    DonneesSource = Range("A1:C1000").Value

    ' ===== TRAITEMENT EN MÉMOIRE =====
    Dim i As Long
    For i = 1 To UBound(DonneesSource, 1)
        ' Traitement sur le tableau en mémoire
        DonneesSource(i, 3) = DonneesSource(i, 2) * 1.2
    Next i

    ' ===== ÉCRITURE EN BLOC =====
    Range("A1:C1000").Value = DonneesSource

    ' ===== RESTAURATION =====
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

## Résumé

Une bonne structure de programme est la clé de la qualité :

**Structure de module :**
- **Ordre logique** : Directives → Constantes → Variables → Procédures
- **Documentation** : En-tête de module et historique
- **Séparation** : Public vs Private selon l'usage

**Structure de procédure :**
- **En-tête** : Documentation de la fonction
- **Sections** : Déclarations → Initialisation → Validation → Traitement → Finalisation
- **Responsabilité unique** : Une procédure = un rôle

**Bonnes pratiques :**
- **Nommage** : Verbes pour actions, noms descriptifs
- **Modularité** : Fonctions réutilisables et spécialisées
- **Gestion d'erreurs** : Structure avec nettoyage garanti
- **Documentation** : Plan du module et flux de traitement

**Patterns utiles :**
- **Validation** : Vérifier avant traiter
- **Initialisation/Nettoyage** : Garantir la cohérence
- **Séparation des responsabilités** : Code maintenable
- **Factorisation** : Éviter la duplication

**À retenir :**
- **Structure = Fondation** : Investissement pour le futur
- **Lisibilité** : Pour vous et vos collègues
- **Maintenance** : Facilite les évolutions
- **Professionnalisme** : Marque d'un code de qualité

Ce chapitre 3 vous a donné toutes les bases fondamentales de VBA. Dans le chapitre suivant, nous découvrirons les procédures et fonctions, qui vous permettront de structurer votre code en blocs réutilisables et efficaces.

⏭️ [4. Les procédures et fonctions](/04-procedures-et-fonctions/)
