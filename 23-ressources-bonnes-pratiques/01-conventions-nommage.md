🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.1 Conventions de nommage

## Introduction

Les conventions de nommage sont l'un des aspects les plus importants des bonnes pratiques en programmation. Elles consistent à établir des règles claires et cohérentes pour nommer vos variables, procédures, fonctions, modules et autres éléments de votre code VBA.

Un bon système de nommage rend votre code immédiatement compréhensible, même pour quelqu'un qui le découvre pour la première fois. C'est comme avoir un système de classement bien organisé : vous retrouvez instantanément ce que vous cherchez.

## Pourquoi les conventions de nommage sont-elles importantes ?

### Lisibilité immédiate
Comparez ces deux exemples :

```vba
' Mauvais exemple
Dim x As Integer
Dim y As String
x = 25
y = "Dupont"

' Bon exemple
Dim ageClient As Integer
Dim nomFamille As String
ageClient = 25
nomFamille = "Dupont"
```

Dans le second exemple, vous comprenez immédiatement l'intention du code.

### Maintenance facilitée
Quand vous reprenez votre code après plusieurs semaines ou mois, des noms explicites vous évitent de devoir décrypter ce que fait chaque élément.

### Collaboration efficace
Si d'autres personnes doivent travailler sur votre code, des conventions claires leur permettent de comprendre rapidement la structure et la logique.

## Règles générales de base

### 1. Utilisez des noms descriptifs

Vos noms doivent exprimer clairement l'utilisation ou le contenu de l'élément :

```vba
' Évitez
Dim d As Date
Dim c As Integer

' Préférez
Dim dateLivraison As Date
Dim compteurLignes As Integer
```

### 2. Évitez les abréviations obscures

```vba
' Difficile à comprendre
Dim nbCmdEnAtnt As Integer

' Plus clair
Dim nombreCommandesEnAttente As Integer
```

### 3. Soyez cohérent

Si vous utilisez "nombre" quelque part, n'utilisez pas "nb" ailleurs pour la même notion :

```vba
' Incohérent
Dim nombreClients As Integer
Dim nbProduits As Integer

' Cohérent
Dim nombreClients As Integer
Dim nombreProduits As Integer
```

## Convention pour les variables

### Variables simples

Utilisez la notation "camelCase" (première lettre minuscule, puis majuscule pour chaque mot) :

```vba
Dim ageUtilisateur As Integer
Dim nomCompletClient As String
Dim dateFinContrat As Date
Dim montantFactureTTC As Double
Dim estClientActif As Boolean
```

### Préfixes pour les types de données (optionnel)

Certains développeurs utilisent des préfixes pour identifier rapidement le type de données. Cette pratique, appelée "notation hongroise", n'est pas obligatoire mais peut être utile :

```vba
Dim strNomClient As String      ' str pour String
Dim intAge As Integer          ' int pour Integer
Dim dblMontant As Double       ' dbl pour Double
Dim blnEstActif As Boolean     ' bln pour Boolean
Dim dtDateNaissance As Date    ' dt pour Date
```

**Note pour débutants :** Choisissez une approche (avec ou sans préfixes) et restez cohérent dans tout votre projet.

### Variables de boucle

Pour les compteurs de boucles, les conventions classiques sont acceptables :

```vba
Dim i As Integer    ' Pour une boucle simple
Dim j As Integer    ' Pour une boucle imbriquée
Dim k As Integer    ' Pour une troisième boucle imbriquée

' Ou plus explicite pour des boucles complexes
Dim indiceLigne As Integer
Dim indiceColonne As Integer
```

## Convention pour les constantes

Les constantes doivent être écrites en MAJUSCULES avec des underscores pour séparer les mots :

```vba
Const TAUX_TVA = 0.2
Const MESSAGE_ERREUR_FICHIER = "Impossible d'ouvrir le fichier"
Const NOMBRE_MAX_TENTATIVES = 3
Const CHEMIN_DOSSIER_SAUVEGARDE = "C:\Sauvegardes\"
```

## Convention pour les procédures et fonctions

### Nommage des procédures (Sub)

Utilisez des verbes qui décrivent l'action effectuée, en PascalCase (première lettre majuscule pour chaque mot) :

```vba
Sub CalculerMontantTotal()
Sub AfficherRapportVentes()
Sub SupprimerLignesVides()
Sub ExporterDonneesCSV()
Sub InitialiserParametres()
```

### Nommage des fonctions (Function)

Les fonctions retournent une valeur, leurs noms peuvent donc être des noms ou des questions :

```vba
Function ObtenirNombreClients() As Integer
Function CalculerRemise(montant As Double) As Double
Function EstClientVIP(idClient As String) As Boolean
Function FormatageDate(uneDate As Date) As String
```

### Paramètres des procédures et fonctions

Utilisez la même convention que pour les variables :

```vba
Sub CalculerFacture(montantHT As Double, tauxTVA As Double, nomClient As String)
Function CalculerAge(dateNaissance As Date) As Integer
```

## Convention pour les modules

### Modules standard

Utilisez des noms descriptifs en PascalCase qui indiquent le domaine fonctionnel :

```vba
' Exemples de noms de modules
ModuleCalculs
ModuleFichiers
ModuleInterfaceUtilisateur
ModuleRapports
ModuleUtilitaires
```

### Modules de classe

Pour les modules de classe, utilisez des noms qui représentent l'objet modélisé :

```vba
Client
Produit
Commande
FactureVente
GestionnaireStock
```

## Convention pour les contrôles d'interface

Si vous créez des UserForms, utilisez des préfixes pour identifier rapidement le type de contrôle :

```vba
' TextBox (zone de texte)
txtNomClient
txtAdresseEmail
txtMontantCommande

' ComboBox (liste déroulante)
cboCategorieProduit
cboVilleClient

' ListBox (zone de liste)
lstProduitsSelectionnes
lstClientsActifs

' CommandButton (bouton)
btnValider
btnAnnuler
btnRechercherClient

' Label (étiquette)
lblTitrePrincipal
lblMessageErreur

' CheckBox (case à cocher)
chkClientVIP
chkLivraisonUrgente
```

## Conventions pour les feuilles Excel

### Noms de feuilles de calcul

Évitez les noms génériques comme "Feuil1", "Feuil2". Utilisez des noms descriptifs :

```vba
' Au lieu de Feuil1, Feuil2, Feuil3
DonneesClients
RapportVentes
Parametres
TableauBord
CalculsIntermediaires
```

### Noms de plages nommées

Utilisez des noms explicites pour vos plages nommées Excel :

```vba
' Au lieu de Zone1, Données1
PlageClients
ListeProduits
TableauVentes2024
ZoneSaisie
CellulesCalculs
```

## Conseils pratiques pour débutants

### 1. Commencez simple
Ne vous compliquez pas la vie au début. L'important est d'être cohérent et descriptif :

```vba
' Simple et efficace pour débuter
Dim nom As String
Dim age As Integer
Dim salaire As Double
```

### 2. Évitez les caractères spéciaux
VBA n'accepte pas tous les caractères. Restez sur les lettres, chiffres et underscores :

```vba
' Évitez
Dim montant€ As Double        ' Le € n'est pas accepté
Dim nom-client As String      ' Le - n'est pas accepté

' Utilisez
Dim montantEuros As Double
Dim nomClient As String
```

### 3. Attention aux mots réservés
Ne donnez pas à vos variables des noms qui sont déjà utilisés par VBA :

```vba
' Évitez ces noms réservés
Dim Date As Date      ' "Date" est une fonction VBA
Dim Name As String    ' "Name" est une propriété VBA
Dim Value As Double   ' "Value" est une propriété VBA

' Utilisez plutôt
Dim dateCommande As Date
Dim nomProduit As String
Dim valeurVente As Double
```

### 4. Longueur raisonnable
Trouvez le bon équilibre entre précision et concision :

```vba
' Trop court
Dim d As Date
Dim n As String

' Trop long
Dim dateDeLaFactureDeVenteDuClientDuMoisDeMarsDeuxMilleVingtQuatre As Date

' Juste ce qu'il faut
Dim dateFacture As Date
Dim nomClientFacture As String
```

## Exemple complet avec bonnes conventions

Voici un exemple de code qui respecte toutes les conventions de nommage :

```vba
' Module : ModuleGestionClients
Option Explicit

' Constantes du module
Const TAUX_REMISE_VIP = 0.15
Const NOMBRE_MAX_COMMANDES = 100

' Procédure pour traiter une commande client
Sub TraiterCommandeClient()
    Dim nomClient As String
    Dim montantCommande As Double
    Dim estClientVIP As Boolean
    Dim montantFinal As Double

    nomClient = "Martin Durand"
    montantCommande = 1500
    estClientVIP = True

    montantFinal = CalculerMontantAvecRemise(montantCommande, estClientVIP)

    AfficherResultatCommande nomClient, montantFinal
End Sub

' Fonction pour calculer le montant avec remise
Function CalculerMontantAvecRemise(montantBase As Double, clientVIP As Boolean) As Double
    Dim montantCalcule As Double

    montantCalcule = montantBase

    If clientVIP Then
        montantCalcule = montantBase * (1 - TAUX_REMISE_VIP)
    End If

    CalculerMontantAvecRemise = montantCalcule
End Function

' Procédure pour afficher le résultat
Sub AfficherResultatCommande(nomDuClient As String, montantAPayer As Double)
    MsgBox "Client : " & nomDuClient & vbCrLf & _
           "Montant à payer : " & Format(montantAPayer, "0.00") & " €"
End Sub
```

## Résumé des bonnes pratiques

1. **Soyez descriptif** : Vos noms doivent expliquer l'usage
2. **Restez cohérent** : Utilisez toujours les mêmes conventions dans tout votre projet
3. **Utilisez camelCase pour les variables** : premierMot, deuxièmeMot
4. **Utilisez PascalCase pour les procédures** : PremierMot, DeuxièmeMot
5. **Utilisez MAJUSCULES pour les constantes** : PREMIERE_CONSTANTE
6. **Évitez les abréviations obscures** : préférez "nombreClients" à "nbCli"
7. **Choisissez des longueurs raisonnables** : ni trop court, ni trop long
8. **Évitez les mots réservés VBA** : ne nommez pas une variable "Date" ou "Name"

En suivant ces conventions, votre code VBA sera beaucoup plus professionnel, lisible et maintenable. C'est un investissement en temps au début qui vous fera gagner énormément de temps par la suite !

⏭️
