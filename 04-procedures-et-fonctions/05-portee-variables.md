🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.5 Portée des variables (Public, Private, Static)

## Introduction

La **portée** d'une variable détermine **où** et **comment longtemps** cette variable peut être utilisée dans votre programme. C'est un concept fondamental qui vous permet de contrôler l'accès et la durée de vie de vos données. Une bonne compréhension de la portée vous évitera de nombreux problèmes et rendra votre code plus organisé.

## Comprendre la portée avec une analogie

### L'analogie de la maison

Imaginez une maison avec différentes pièces :

- **Variables locales** = Objets dans votre chambre (seul vous y avez accès)
- **Variables Private** = Objets dans votre appartement (toute votre famille peut les utiliser)
- **Variables Public** = Objets dans les parties communes de l'immeuble (tous les résidents peuvent les utiliser)
- **Variables Static** = Objets que vous gardez même quand vous quittez temporairement la pièce

## Les différents niveaux de portée

### Vue d'ensemble

```vba
' Variables PUBLIC - Disponibles partout
Public compteurGlobal As Integer

' Variables PRIVATE - Disponibles dans ce module uniquement
Private configurationModule As String

' Dans une procédure...
Sub ExempleProcedure()
    ' Variables LOCALES - Disponibles dans cette procédure uniquement
    Dim compteurLocal As Integer

    ' Variables STATIC - Gardent leur valeur entre les appels
    Static compteurPersistant As Integer
End Sub
```

## 1. Variables locales (Dim)

### Définition et caractéristiques

Les variables **locales** sont déclarées à l'intérieur d'une procédure ou fonction avec le mot-clé `Dim`.

**Caractéristiques :**
- **Portée** : Uniquement dans la procédure où elles sont déclarées
- **Durée de vie** : Détruite à la fin de la procédure
- **Accès** : Impossible depuis d'autres procédures

### Exemple pratique

```vba
Sub CalculerFacture()
    ' Variables locales - visibles uniquement dans cette procédure
    Dim prixHT As Double
    Dim tva As Double
    Dim prixTTC As Double

    prixHT = 100
    tva = prixHT * 0.2
    prixTTC = prixHT + tva

    MsgBox "Prix TTC : " & prixTTC & "€"

    ' À la fin de cette procédure, toutes ces variables sont détruites
End Sub

Sub AutreProcedure()
    ' ❌ ERREUR ! Ces variables n'existent pas ici
    ' MsgBox prixHT  ' Variable non définie !

    ' ✅ Il faut redéclarer des variables locales
    Dim monPrix As Double
    monPrix = 50
    MsgBox monPrix
End Sub
```

### Avantages des variables locales

- **Sécurité** : Pas de conflit entre procédures
- **Mémoire** : Libérée automatiquement
- **Clarté** : Données limitées à leur contexte d'usage

## 2. Variables Private (au niveau module)

### Définition et caractéristiques

Les variables **Private** sont déclarées au début d'un module, avant toute procédure.

**Caractéristiques :**
- **Portée** : Toutes les procédures du même module
- **Durée de vie** : Tant que le module est chargé
- **Accès** : Invisible depuis les autres modules

### Syntaxe et placement

```vba
' === EN HAUT DU MODULE, AVANT TOUTE PROCÉDURE ===
Private nomUtilisateur As String  
Private compteurOperations As Integer  
Private dateDebut As Date  

' === ENSUITE, VOS PROCÉDURES ===
Sub InitialiserSession()
    nomUtilisateur = "Marie Dupont"
    compteurOperations = 0
    dateDebut = Now()
    MsgBox "Session initialisée pour " & nomUtilisateur
End Sub

Sub EnregistrerOperation()
    compteurOperations = compteurOperations + 1
    MsgBox "Opération n°" & compteurOperations & " pour " & nomUtilisateur
End Sub

Sub AfficherStatistiques()
    Dim dureeSession As Double
    dureeSession = Now() - dateDebut

    MsgBox "Utilisateur : " & nomUtilisateur & vbNewLine & _
           "Opérations : " & compteurOperations & vbNewLine & _
           "Durée : " & Format(dureeSession, "hh:mm:ss")
End Sub
```

### Cas d'usage typiques

**Configuration du module :**
```vba
Private cheminFichiers As String  
Private formatExport As String  

Sub ConfigurerModule()
    cheminFichiers = "C:\MonDossier\"
    formatExport = "xlsx"
End Sub

Sub ExporterDonnees()
    ' Utilise les variables de configuration
    ActiveWorkbook.SaveAs cheminFichiers & "Export." & formatExport
End Sub
```

## 3. Variables Public (globales)

### Définition et caractéristiques

Les variables **Public** sont accessibles depuis **tous les modules** de votre projet.

**Caractéristiques :**
- **Portée** : Tout le projet VBA
- **Durée de vie** : Tant que l'application est ouverte
- **Accès** : Visible partout (modules, procédures, fonctions)

### Déclaration et utilisation

```vba
' === MODULE 1 ===
Public versionApplication As String  
Public utilisateurActuel As String  
Public modeDebug As Boolean  

Sub InitialiserApplication()
    versionApplication = "2.1.0"
    utilisateurActuel = "Admin"
    modeDebug = True
    MsgBox "Application initialisée - Version " & versionApplication
End Sub

' === MODULE 2 ===
Sub AfficherInfosUtilisateur()
    ' Accès aux variables publiques depuis un autre module
    MsgBox "Utilisateur connecté : " & utilisateurActuel & vbNewLine & _
           "Version : " & versionApplication
End Sub

Sub GererErreur()
    If modeDebug Then
        MsgBox "Mode debug activé - Affichage des erreurs détaillées"
    End If
End Sub
```

### Précautions avec les variables Public

**⚠️ Attention aux conflits :**
```vba
' Module1
Public compteur As Integer

Sub ProcedureA()
    compteur = 10
End Sub

' Module2
Sub ProcedureB()
    compteur = 20  ' Modifie la même variable !
End Sub

Sub ProcedureC()
    MsgBox compteur  ' Affichera 20, pas 10 !
End Sub
```

## 4. Variables Static

### Définition et caractéristiques

Les variables **Static** conservent leur valeur entre les appels successifs d'une procédure.

**Caractéristiques :**
- **Portée** : Locale à la procédure
- **Durée de vie** : Persiste entre les appels
- **Accès** : Uniquement dans la procédure où elle est déclarée

### Exemple classique : Compteur d'appels

```vba
Sub CompteurAppels()
    Static nombreAppels As Integer

    nombreAppels = nombreAppels + 1
    MsgBox "Cette procédure a été appelée " & nombreAppels & " fois"
End Sub

Sub TestCompteur()
    CompteurAppels  ' Affiche : 1 fois
    CompteurAppels  ' Affiche : 2 fois
    CompteurAppels  ' Affiche : 3 fois
End Sub
```

### Comparaison Dim vs Static

```vba
Sub AvecDim()
    Dim compteur As Integer  ' Remise à 0 à chaque appel
    compteur = compteur + 1
    MsgBox "Dim : " & compteur  ' Affiche toujours 1
End Sub

Sub AvecStatic()
    Static compteur As Integer  ' Garde sa valeur
    compteur = compteur + 1
    MsgBox "Static : " & compteur  ' Affiche 1, puis 2, puis 3...
End Sub
```

### Cas d'usage pratiques pour Static

**Générateur d'ID unique :**
```vba
Function ObtenirNouvelID() As Integer
    Static dernierID As Integer
    dernierID = dernierID + 1
    ObtenirNouvelID = dernierID
End Function

Sub CreerPlusieursIDs()
    MsgBox "ID 1 : " & ObtenirNouvelID()  ' 1
    MsgBox "ID 2 : " & ObtenirNouvelID()  ' 2
    MsgBox "ID 3 : " & ObtenirNouvelID()  ' 3
End Sub
```

**Cache de calcul :**
```vba
Function CalculComplexe(valeur As Double) As Double
    Static derniereValeur As Double
    Static dernierResultat As Double

    ' Si même valeur, retourne le résultat en cache
    If valeur = derniereValeur Then
        CalculComplexe = dernierResultat
        Exit Function
    End If

    ' Sinon, calcule et met en cache
    dernierResultat = valeur * valeur * valeur  ' Calcul complexe
    derniereValeur = valeur
    CalculComplexe = dernierResultat
End Function
```

## Tableau récapitulatif des portées

| Type | Mot-clé | Portée | Durée de vie | Déclaration |
|------|---------|---------|--------------|-------------|
| **Locale** | `Dim` | Procédure uniquement | Fin de procédure | Dans la procédure |
| **Module** | `Private` | Toutes procédures du module | Vie du module | Début de module |
| **Globale** | `Public` | Tout le projet | Vie de l'application | Début de module |
| **Persistante** | `Static` | Procédure uniquement | Vie de l'application | Dans la procédure |

## Exemples pratiques complets

### Exemple 1 : Système de configuration

```vba
' === Variables au niveau module ===
Private cheminBase As String  
Private utilisateurCourant As String  

' === Variables publiques ===
Public modeDebug As Boolean

Sub ConfigurerApplication()
    ' Configuration locale temporaire
    Dim reponse As String

    cheminBase = "C:\MonApplication\"
    utilisateurCourant = Environ("USERNAME")
    modeDebug = True

    reponse = InputBox("Chemin de base:", , cheminBase)
    If reponse <> "" Then cheminBase = reponse

    MsgBox "Configuration terminée pour " & utilisateurCourant
End Sub

Sub SauvegarderFichier(nomFichier As String)
    Dim cheminComplet As String
    cheminComplet = cheminBase & nomFichier

    If modeDebug Then
        MsgBox "Sauvegarde dans : " & cheminComplet
    End If

    ' Code de sauvegarde...
End Sub
```

### Exemple 2 : Compteur de performances

```vba
Function MesurerPerformance() As String
    Static nombreTests As Integer
    Static tempsTotal As Double
    Dim debut As Double

    debut = Timer
    nombreTests = nombreTests + 1

    ' Simulation d'un traitement
    Application.Wait Now + TimeValue("00:00:01")

    tempsTotal = tempsTotal + (Timer - debut)

    MesurerPerformance = "Test " & nombreTests & _
                        " - Temps moyen : " & Format(tempsTotal / nombreTests, "0.00") & "s"
End Function
```

## Bonnes pratiques

### 1. Règle de la portée minimale

**✅ Principe :** Utilisez la portée la plus restrictive possible.

```vba
' ❌ Évitez les variables publiques inutiles
Public tempValue As String

' ✅ Préférez les variables locales
Sub TraiterDonnees()
    Dim tempValue As String  ' Suffit pour cette procédure
    ' ...
End Sub
```

### 2. Nommage selon la portée

```vba
' Variables locales : notation simple
Dim nom As String  
Dim compteur As Integer  

' Variables de module : préfixe explicite
Private m_configuration As String  
Private m_etatModule As Boolean  

' Variables publiques : préfixe global
Public g_versionApp As String  
Public g_utilisateur As String  
```

### 3. Initialisation appropriée

```vba
' ✅ Initialisation des variables de module
Private m_estInitialise As Boolean

Sub InitialiserModule()
    If Not m_estInitialise Then
        ' Code d'initialisation
        m_estInitialise = True
    End If
End Sub
```

### 4. Documentation de la portée

```vba
'===============================================
' VARIABLES DE MODULE (PRIVATE)
'===============================================
Private m_cheminTravail As String    ' Chemin de base pour les fichiers  
Private m_compteurErreurs As Integer ' Nombre d'erreurs dans ce module  

'===============================================
' PROCÉDURES PUBLIQUES
'===============================================
Public Sub ExporterDonnees()
    ' Utilise les variables de module...
End Sub
```

## Erreurs courantes à éviter

### 1. Variables non initialisées

```vba
' ❌ Problème
Private compteur As Integer  ' Vaut 0 par défaut, mais pas explicite

Sub Incrementer()
    compteur = compteur + 1  ' Fonctionne par chance
End Sub

' ✅ Solution
Private compteur As Integer

Sub InitialiserCompteur()
    compteur = 0  ' Initialisation explicite
End Sub
```

### 2. Confusion entre Static et module

```vba
' ❌ Static dans chaque procédure (redondant)
Sub Proc1()
    Static config As String
    config = "valeur"
End Sub

Sub Proc2()
    Static config As String  ' Différente variable !
    MsgBox config  ' Vide !
End Sub

' ✅ Variable de module (partagée)
Private config As String

Sub Proc1()
    config = "valeur"
End Sub

Sub Proc2()
    MsgBox config  ' Affiche "valeur"
End Sub
```

### 3. Abus des variables Public

```vba
' ❌ Trop de variables globales
Public temp1, temp2, temp3, result, counter, flag

' ✅ Utilisation ciblée
Public AppConfig As String  ' Vraiment nécessaire globalement
' Le reste en local ou module selon les besoins
```

## Conseils pour choisir la bonne portée

### Questions à se poser

1. **Cette variable est-elle utilisée dans plusieurs procédures ?**
   - Non → `Dim` (locale)
   - Oui, même module → `Private`
   - Oui, plusieurs modules → `Public`

2. **Dois-je conserver la valeur entre les appels ?**
   - Oui, dans une procédure → `Static`
   - Oui, dans tout le module → `Private`

3. **Cette information concerne-t-elle tout l'application ?**
   - Oui → `Public`
   - Non → Plus restrictif

### Hiérarchie recommandée

1. **D'abord** : Variables locales (`Dim`)
2. **Ensuite** : Variables Static si persistance nécessaire
3. **Puis** : Variables Private si partage dans le module
4. **En dernier** : Variables Public si vraiment globales

## Récapitulatif des concepts clés

1. **Dim** : Variables locales, détruites à la fin de la procédure
2. **Private** : Variables de module, partagées dans un seul module
3. **Public** : Variables globales, accessibles partout
4. **Static** : Variables locales qui persistent entre les appels
5. **Règle d'or** : Utilisez la portée la plus restrictive possible
6. **Organisation** : Déclarez les variables de module en haut
7. **Documentation** : Commentez le rôle des variables partagées

La maîtrise de la portée des variables vous permettra de créer des programmes plus robustes, organisés et faciles à maintenir !

⏭️
