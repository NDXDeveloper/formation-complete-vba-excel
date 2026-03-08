🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.1. Déclaration d'API

## Qu'est-ce qu'une déclaration d'API ?

Une **déclaration d'API** est la façon de dire à VBA qu'il existe une fonction dans Windows que vous souhaitez utiliser. C'est comme donner à VBA le "numéro de téléphone" et les "instructions d'utilisation" d'un service Windows.

**Analogie simple :**
Imaginez que vous voulez commander une pizza :
- **La déclaration** = Noter le numéro de téléphone de la pizzeria et le format de commande
- **L'appel** = Téléphoner effectivement pour commander
- **Les paramètres** = Dire quelle pizza, quelle taille, quelle adresse

En VBA, vous devez d'abord "déclarer" une fonction API avant de pouvoir l'utiliser.

## Syntaxe de base d'une déclaration

### Structure générale

```vba
[Private|Public] Declare [PtrSafe] Function NomFonction Lib "NomBibliothèque" _
    Alias "NomOriginal" (paramètres) As TypeRetour
```

ou

```vba
[Private|Public] Declare [PtrSafe] Sub NomProcédure Lib "NomBibliothèque" _
    Alias "NomOriginal" (paramètres)
```

### Décomposition des éléments

#### 1. Visibilité (Private/Public)
```vba
Private Declare Function...  ' Utilisable seulement dans ce module  
Public Declare Function...   ' Utilisable dans tout le projet  
```

#### 2. PtrSafe (important pour la compatibilité)
```vba
' Pour VBA 64 bits (recommandé pour la compatibilité)
Declare PtrSafe Function...

' Ancienne syntaxe (32 bits seulement)
Declare Function...
```

#### 3. Function vs Sub
```vba
Declare Function...   ' Retourne une valeur  
Declare Sub...        ' Ne retourne rien (équivalent d'une Sub VBA)  
```

#### 4. Lib "Bibliothèque"
```vba
Lib "kernel32"        ' Bibliothèque système de base  
Lib "user32"          ' Interface utilisateur  
Lib "advapi32"        ' Services avancés  
```

#### 5. Alias (optionnel)
```vba
Alias "GetUserNameA"  ' Nom exact de la fonction dans Windows
```

## Exemples de déclarations courantes

### Exemple 1 : GetUserName (obtenir le nom d'utilisateur)

```vba
' Déclaration complète avec gestion 32/64 bits
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If
```

**Explication :**
- `GetUserName` : Nom qu'on donne à la fonction en VBA
- `Lib "advapi32.dll"` : La fonction se trouve dans cette bibliothèque Windows
- `Alias "GetUserNameA"` : Le vrai nom de la fonction (version ANSI)
- `lpBuffer As String` : Paramètre pour recevoir le nom d'utilisateur
- `nSize As Long` : Taille du buffer de réception
- `As Long` : La fonction retourne un nombre (succès/échec)

### Exemple 2 : Sleep (pause en millisecondes)

```vba
' Déclaration simple
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
```

**Explication :**
- `Sub Sleep` : C'est une procédure, elle ne retourne rien
- `Lib "kernel32"` : Bibliothèque système de base
- `dwMilliseconds As Long` : Nombre de millisecondes à attendre

### Exemple 3 : GetSystemMetrics (informations écran)

```vba
' Déclaration pour obtenir les métriques système
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
#End If
```

**Explication :**
- `nIndex As Long` : Code spécifiant quelle métrique obtenir
- `As Long` : Retourne la valeur de la métrique demandée

## Gestion de la compatibilité 32/64 bits

### Problème de compatibilité
VBA existe en version 32 bits et 64 bits. Les déclarations doivent être adaptées selon la version.

### Solution avec compilation conditionnelle

```vba
' Template recommandé pour toutes vos déclarations
#If VBA7 Then
    ' Version VBA 7+ (Office 2010 et plus récent)
    Private Declare PtrSafe Function MaFonction Lib "malib.dll" _
        (paramètres) As TypeRetour
#Else
    ' Version VBA 6 et antérieure (Office 2007 et plus ancien)
    Private Declare Function MaFonction Lib "malib.dll" _
        (paramètres) As TypeRetour
#End If
```

### Constantes utiles pour les types

```vba
' Constantes pour gérer les différences 32/64 bits
#If VBA7 Then
    #If Win64 Then
        Private Const PTR_SIZE = 8  ' 64 bits
    #Else
        Private Const PTR_SIZE = 4  ' 32 bits
    #End If
#Else
    Private Const PTR_SIZE = 4      ' VBA6 = toujours 32 bits
#End If
```

## Types de paramètres courants

### Types simples VBA
```vba
' Types directement compatibles
ByVal monEntier As Long          ' Nombre entier 32 bits  
ByVal monTexte As String         ' Chaîne de caractères  
ByVal monBooleen As Boolean      ' Vrai/Faux  
ByRef maVariable As Long         ' Passage par référence  
```

### Types spéciaux Windows
```vba
' Types spécifiques aux API Windows
ByVal hWnd As Long              ' Handle de fenêtre (32 bits)  
ByVal hWnd As LongPtr           ' Handle de fenêtre (64 bits compatible)  
ByVal lpParam As Long           ' Paramètre générique  
ByVal dwFlags As Long           ' Flags/options  
```

### Gestion des chaînes de caractères

#### Chaînes en entrée (ByVal)
```vba
' Pour passer du texte à l'API
Declare PtrSafe Function MaFonction Lib "malib" _
    (ByVal texte As String) As Long

' Utilisation
Dim resultat As Long  
resultat = MaFonction("Mon texte")  
```

#### Chaînes en sortie (ByRef ou ByVal avec buffer)
```vba
' Pour recevoir du texte de l'API
Declare PtrSafe Function GetInfo Lib "malib" _
    (ByVal buffer As String, ByVal taille As Long) As Long

' Utilisation
Dim buffer As String  
Dim taille As Long  
buffer = Space(255)  ' Créer un buffer de 255 caractères  
taille = 255  
GetInfo buffer, taille  
```

## Déclarations avec structures (Types personnalisés)

### Définir une structure
```vba
' Définition d'une structure Windows courante
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Déclaration d'une API qui utilise cette structure
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowRect Lib "user32" _
        (ByVal hWnd As LongPtr, lpRect As RECT) As Long
#Else
    Private Declare Function GetWindowRect Lib "user32" _
        (ByVal hWnd As Long, lpRect As RECT) As Long
#End If
```

### Utilisation de la structure
```vba
Sub ExempleStructure()
    Dim rectangle As RECT
    Dim handle As LongPtr

    ' handle = ... (obtenir le handle d'une fenêtre)

    If GetWindowRect(handle, rectangle) <> 0 Then
        Debug.Print "Gauche : " & rectangle.Left
        Debug.Print "Haut : " & rectangle.Top
        Debug.Print "Droite : " & rectangle.Right
        Debug.Print "Bas : " & rectangle.Bottom
    End If
End Sub
```

## Bonnes pratiques pour les déclarations

### 1. Toujours utiliser la compilation conditionnelle
```vba
' ✅ Toujours faire ceci
#If VBA7 Then
    Private Declare PtrSafe Function...
#Else
    Private Declare Function...
#End If

' ❌ Éviter ceci (pas compatible)
Private Declare Function...
```

### 2. Utiliser Private sauf nécessité
```vba
' ✅ Par défaut, garder privé
Private Declare PtrSafe Function...

' ✅ Public seulement si utilisé dans plusieurs modules
Public Declare PtrSafe Function...
```

### 3. Noms explicites
```vba
' ✅ Noms clairs
Private Declare PtrSafe Function ObtenirNomUtilisateur...  
Private Declare PtrSafe Function PauseEnMillisecondes...  

' ❌ Noms cryptiques
Private Declare PtrSafe Function GetUN...  
Private Declare PtrSafe Function Slp...  
```

### 4. Commentaires explicatifs
```vba
' ✅ Toujours documenter
' Obtient le nom d'utilisateur Windows actuel
' Paramètres :
'   lpBuffer : Buffer pour recevoir le nom (ByVal)
'   nSize : Taille du buffer en caractères
' Retour :
'   <> 0 si succès, 0 si échec
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If
```

### 5. Regroupement logique
```vba
' ✅ Organiser les déclarations par thème
' ==========================================
' APIs SYSTÈME
' ==========================================
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName...
    Private Declare PtrSafe Function GetComputerName...
#Else
    Private Declare Function GetUserName...
    Private Declare Function GetComputerName...
#End If

' ==========================================
' APIs FENÊTRES
' ==========================================
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow...
    Private Declare PtrSafe Function SetWindowPos...
#Else
    Private Declare Function FindWindow...
    Private Declare Function SetWindowPos...
#End If
```

## Erreurs courantes et solutions

### Erreur 1 : "Declare statements not allowed in object modules"
```vba
' ❌ Problème : Déclaration dans une feuille Excel ou UserForm
' Solution : Mettre les déclarations dans un module standard

' Dans Module1 (pas dans Feuil1 ou UserForm1)
Private Declare PtrSafe Function MaFonction...
```

### Erreur 2 : "Can't find DLL entry point"
```vba
' ❌ Problème : Mauvais nom de fonction ou d'alias
Declare PtrSafe Function GetUserName Lib "advapi32" _
    Alias "GetUserNam" ' <- Erreur de frappe

' ✅ Solution : Vérifier l'orthographe exacte
Declare PtrSafe Function GetUserName Lib "advapi32" _
    Alias "GetUserNameA"  ' <- Nom correct
```

### Erreur 3 : "Bad DLL calling convention"
```vba
' ❌ Problème : Mauvais types de paramètres
Declare PtrSafe Function MaFonction(param As Integer) As Long

' ✅ Solution : Utiliser les bons types Windows
Declare PtrSafe Function MaFonction(ByVal param As Long) As Long
```

### Erreur 4 : Plantage de l'application
```vba
' ❌ Problème : Mauvaise gestion des chaînes
Dim nom As String  
GetUserName nom, 50  ' <- Buffer non initialisé  

' ✅ Solution : Initialiser le buffer correctement
Dim nom As String  
nom = Space(50)      ' <- Créer un buffer de 50 caractères  
GetUserName nom, 50  
```

## Où trouver les déclarations

### 1. Documentation Microsoft
- **MSDN Library** : Documentation officielle des API Windows
- **Platform SDK** : Exemples de déclarations en différents langages

### 2. Sites spécialisés
- **api-guide.com** : Base de données de déclarations VBA
- **vbapi.com** : Convertisseur automatique C++ vers VBA

### 3. Forums et communautés
- **stackoverflow.com** : Questions/réponses avec exemples
- **developpez.com** : Communauté francophone

### 4. Outils de conversion
Certains outils peuvent convertir automatiquement les déclarations C++ en VBA, mais vérifiez toujours le résultat.

## Template de module pour les API

```vba
' ================================================================
' Module : ModuleAPI
' Description : Déclarations d'API Windows courantes
' Auteur : [Votre nom]
' Date : [Date]
' ================================================================

Option Explicit

' ==========================================
' CONSTANTES
' ==========================================
' Constantes pour les valeurs de retour
Private Const API_SUCCESS = 1  
Private Const API_FAILURE = 0  

' Constantes pour les tailles de buffer
Private Const MAX_USERNAME = 256  
Private Const MAX_COMPUTERNAME = 256  

' ==========================================
' TYPES PERSONNALISÉS
' ==========================================
' Structures Windows courantes
Type POINT
    x As Long
    y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' ==========================================
' DÉCLARATIONS D'API
' ==========================================

' APIs Système
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    Private Declare PtrSafe Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#End If

' ==========================================
' FONCTIONS WRAPPER (Interface simplifiée)
' ==========================================

Public Function ObtenirNomUtilisateur() As String
    ' Interface simplifiée pour GetUserName
    Dim buffer As String
    Dim taille As Long

    buffer = Space(MAX_USERNAME)
    taille = MAX_USERNAME

    If GetUserName(buffer, taille) <> 0 Then
        ObtenirNomUtilisateur = Left(buffer, taille - 1)  ' Retirer le caractère null
    Else
        ObtenirNomUtilisateur = ""
    End If
End Function

Public Sub PauseMS(millisecondes As Long)
    ' Interface simplifiée pour Sleep
    If millisecondes > 0 Then
        Sleep millisecondes
    End If
End Sub
```

La déclaration d'API est la première étape cruciale pour utiliser les fonctionnalités Windows avancées. Une déclaration correcte garantit un fonctionnement stable et compatible de vos applications VBA.

⏭️ [API courantes (GetUserName, Sleep)](/17-api-windows/02-api-courantes.md)
