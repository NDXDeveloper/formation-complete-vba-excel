🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.2. API courantes (GetUserName, Sleep)

## Introduction aux API de base

Dans ce chapitre, nous allons explorer les API Windows les plus couramment utilisées et les plus sûres pour débuter. Ces fonctions sont stables, bien documentées et présentent peu de risques pour votre système.

**Pourquoi commencer par ces API ?**
- **Simplicité** : Syntaxe relativement simple à comprendre
- **Sécurité** : Peu de risques de plantage ou de problèmes système
- **Utilité** : Fonctionnalités couramment nécessaires dans les applications
- **Compatibilité** : Fonctionnent sur toutes les versions de Windows

## 1. GetUserName - Obtenir le nom d'utilisateur Windows

### Description
`GetUserName` permet d'obtenir le nom de l'utilisateur actuellement connecté à Windows. C'est différent du nom d'utilisateur Excel ou Office.

### Déclaration complète

```vba
' Déclaration compatible 32/64 bits
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If
```

### Explication des paramètres

#### lpBuffer (String)
- **Rôle** : Buffer (zone mémoire) pour recevoir le nom d'utilisateur
- **Type** : `ByVal lpBuffer As String`
- **Préparation** : Doit être initialisé avec `Space()` avant l'appel

#### nSize (Long)
- **Rôle** : Taille du buffer en caractères
- **Type** : `nSize As Long`
- **Valeur** : Doit correspondre à la taille du buffer créé

#### Valeur de retour (Long)
- **Succès** : Valeur différente de 0
- **Échec** : 0

### Utilisation basique

```vba
Sub ExempleGetUserNameSimple()
    ' Créer un buffer pour recevoir le nom
    Dim nomUtilisateur As String
    Dim tailleBuffer As Long
    Dim resultat As Long

    ' Initialiser le buffer (255 caractères max)
    nomUtilisateur = Space(255)
    tailleBuffer = 255

    ' Appeler l'API
    resultat = GetUserName(nomUtilisateur, tailleBuffer)

    ' Vérifier le succès
    If resultat <> 0 Then
        ' Nettoyer le résultat (retirer les caractères vides)
        nomUtilisateur = Left(nomUtilisateur, tailleBuffer - 1)
        Debug.Print "Utilisateur connecté : " & nomUtilisateur
    Else
        Debug.Print "Erreur lors de la récupération du nom d'utilisateur"
    End If
End Sub
```

### Version améliorée avec gestion d'erreurs

```vba
Function ObtenirNomUtilisateur() As String
    ' Fonction wrapper pour GetUserName avec gestion d'erreurs

    Dim buffer As String
    Dim taille As Long
    Dim resultat As Long

    ' Initialisation
    buffer = Space(256)  ' Buffer de 256 caractères
    taille = 256

    ' Appel de l'API avec gestion d'erreurs
    On Error GoTo GestionErreur

    resultat = GetUserName(buffer, taille)

    If resultat <> 0 Then
        ' Succès : nettoyer et retourner le nom
        ObtenirNomUtilisateur = Trim(Left(buffer, taille - 1))
    Else
        ' Échec : retourner une valeur par défaut
        ObtenirNomUtilisateur = "Utilisateur inconnu"
    End If

    Exit Function

GestionErreur:
    ' En cas d'erreur VBA
    ObtenirNomUtilisateur = "Erreur système"
    Debug.Print "Erreur GetUserName : " & Err.Description
End Function
```

### Utilisation pratique

```vba
Sub UtilisationPratiqueGetUserName()
    Dim utilisateur As String

    ' Obtenir le nom d'utilisateur
    utilisateur = ObtenirNomUtilisateur()

    ' Utiliser le nom dans l'application
    Range("A1").Value = "Bonjour " & utilisateur

    ' Personnaliser un message
    MsgBox "Bienvenue " & utilisateur & " !" & vbCrLf & _
           "Fichier ouvert le " & Format(Now, "dd/mm/yyyy à hh:nn"), _
           vbInformation, "Application Personnalisée"

    ' Logger l'ouverture
    Debug.Print Format(Now, "dd/mm/yyyy hh:nn:ss") & " - Ouverture par " & utilisateur
End Sub
```

## 2. Sleep - Pause précise en millisecondes

### Description
`Sleep` permet de mettre en pause l'exécution du programme pendant un nombre précis de millisecondes. C'est plus précis que `Application.Wait` d'Excel.

### Déclaration complète

```vba
' Déclaration compatible 32/64 bits
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#End If
```

### Explication des paramètres

#### dwMilliseconds (Long)
- **Rôle** : Nombre de millisecondes à attendre
- **Type** : `ByVal dwMilliseconds As Long`
- **Plage** : 0 à environ 2 milliards (2,147,483,647 ms = ~24 jours)
- **Précision** : 1 milliseconde = 0,001 seconde

#### Pas de valeur de retour
- `Sleep` est une `Sub`, elle ne retourne rien
- L'exécution reprend automatiquement après le délai

### Utilisation basique

```vba
Sub ExempleSleepSimple()
    Debug.Print "Début : " & Format(Now, "hh:nn:ss")

    ' Pause de 2 secondes (2000 millisecondes)
    Sleep 2000

    Debug.Print "Fin : " & Format(Now, "hh:nn:ss")

    ' Pause courte de 500 ms
    Sleep 500

    Debug.Print "Après pause courte : " & Format(Now, "hh:nn:ss")
End Sub
```

### Comparaison avec Application.Wait

```vba
Sub ComparaisonMethodesPause()
    Dim debut As Date

    ' Test avec Application.Wait (Excel)
    Debug.Print "=== Test Application.Wait ==="
    debut = Now
    Debug.Print "Début : " & Format(debut, "hh:nn:ss.000")

    Application.Wait Now + TimeValue("00:00:02")  ' 2 secondes

    Debug.Print "Fin : " & Format(Now, "hh:nn:ss.000")
    Debug.Print "Durée réelle : " & Format((Now - debut) * 86400, "0.0") & " secondes"

    ' Pause entre les tests
    Sleep 1000

    ' Test avec Sleep (API Windows)
    Debug.Print "=== Test Sleep API ==="
    debut = Now
    Debug.Print "Début : " & Format(debut, "hh:nn:ss.000")

    Sleep 2000  ' 2000 millisecondes = 2 secondes

    Debug.Print "Fin : " & Format(Now, "hh:nn:ss.000")
    Debug.Print "Durée réelle : " & Format((Now - debut) * 86400, "0.0") & " secondes"
End Sub
```

### Fonctions wrapper utiles

```vba
' Collection de fonctions de pause pratiques
Sub PauseSecondes(secondes As Double)
    ' Pause en secondes (accepte les décimales)
    If secondes > 0 Then
        Sleep CLng(secondes * 1000)
    End If
End Sub

Sub PauseMinutes(minutes As Double)
    ' Pause en minutes
    If minutes > 0 Then
        Sleep CLng(minutes * 60000)
    End If
End Sub

Sub PauseCourte()
    ' Pause courte standard (250 ms)
    Sleep 250
End Sub

Sub PauseMoyenne()
    ' Pause moyenne standard (1 seconde)
    Sleep 1000
End Sub

Sub PauseLongue()
    ' Pause longue standard (3 secondes)
    Sleep 3000
End Sub
```

### Utilisation pratique : Animation simple

```vba
Sub AnimationTexteSimple()
    Dim i As Integer
    Dim cellule As Range

    Set cellule = Range("A1")

    ' Animation de points
    For i = 1 To 10
        cellule.Value = "Chargement" & String(i Mod 4, ".")
        cellule.Font.Color = RGB(0, 100 + (i * 15), 0)  ' Vert progressif

        ' Pause courte pour voir l'animation
        Sleep 300
    Next i

    ' Message final
    cellule.Value = "Terminé !"
    cellule.Font.Color = RGB(0, 150, 0)
    cellule.Font.Bold = True
End Sub
```

## 3. GetComputerName - Nom de l'ordinateur

### Description
Obtient le nom NetBIOS de l'ordinateur local.

### Déclaration et utilisation

```vba
' Déclaration
#If VBA7 Then
    Private Declare PtrSafe Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Function ObtenirNomOrdinateur() As String
    Dim buffer As String
    Dim taille As Long

    buffer = Space(256)
    taille = 256

    If GetComputerName(buffer, taille) <> 0 Then
        ObtenirNomOrdinateur = Left(buffer, taille)
    Else
        ObtenirNomOrdinateur = "Ordinateur inconnu"
    End If
End Function
```

## 4. Beep - Son système

### Description
Émet un son système simple.

### Déclaration et utilisation

```vba
' Déclaration
#If VBA7 Then
    Private Declare PtrSafe Function Beep Lib "kernel32" _
        (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#Else
    Private Declare Function Beep Lib "kernel32" _
        (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#End If

Sub ExempleBeep()
    ' Son aigu court (1000 Hz pendant 200 ms)
    Beep 1000, 200

    Sleep 500  ' Pause entre les sons

    ' Son grave long (400 Hz pendant 800 ms)
    Beep 400, 800
End Sub

Sub AlerteSonore()
    ' Séquence d'alerte
    Dim i As Integer
    For i = 1 To 3
        Beep 800, 150   ' Son court
        Sleep 100       ' Pause courte
    Next i
End Sub
```

## 5. GetSystemMetrics - Informations écran

### Description
Obtient diverses métriques du système (résolution écran, taille des bordures, etc.).

### Déclaration et constantes

```vba
' Déclaration
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
#End If

' Constantes pour les métriques courantes
Private Const SM_CXSCREEN = 0       ' Largeur écran  
Private Const SM_CYSCREEN = 1       ' Hauteur écran  
Private Const SM_CXVSCROLL = 2      ' Largeur barre défilement  
Private Const SM_CYHSCROLL = 3      ' Hauteur barre défilement  
Private Const SM_CYCAPTION = 4      ' Hauteur barre titre  
Private Const SM_CMONITORS = 80     ' Nombre d'écrans  
```

### Utilisation pratique

```vba
Sub InformationsEcran()
    Dim largeur As Long
    Dim hauteur As Long
    Dim nbEcrans As Long

    ' Obtenir les dimensions de l'écran principal
    largeur = GetSystemMetrics(SM_CXSCREEN)
    hauteur = GetSystemMetrics(SM_CYSCREEN)
    nbEcrans = GetSystemMetrics(SM_CMONITORS)

    ' Afficher les informations
    Debug.Print "=== INFORMATIONS ÉCRAN ==="
    Debug.Print "Résolution : " & largeur & " x " & hauteur & " pixels"
    Debug.Print "Nombre d'écrans : " & nbEcrans
    Debug.Print "Rapport d'aspect : " & Format(largeur / hauteur, "0.00")

    ' Adapter l'interface selon la résolution
    If largeur >= 1920 Then
        Debug.Print "Écran haute résolution détecté"
        ' Ajuster la taille des UserForms, etc.
    ElseIf largeur <= 1024 Then
        Debug.Print "Écran basse résolution détecté"
        ' Interface compacte
    End If
End Sub
```

## Application pratique : Système d'information

### Module complet avec toutes les API

```vba
' ================================================================
' Module : InformationsSysteme
' Description : Collecte d'informations système via API Windows
' ================================================================

Option Explicit

' Déclarations d'API
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    Private Declare PtrSafe Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long

    Private Declare PtrSafe Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)

    Private Declare PtrSafe Function Beep Lib "kernel32" _
        (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    Private Declare Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long

    Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)

    Private Declare Function Beep Lib "kernel32" _
        (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#End If

' Constantes
Private Const SM_CXSCREEN = 0  
Private Const SM_CYSCREEN = 1  
Private Const SM_CMONITORS = 80  

' Fonctions wrapper publiques
Public Function ObtenirNomUtilisateur() As String
    Dim buffer As String, taille As Long
    buffer = Space(256): taille = 256

    If GetUserName(buffer, taille) <> 0 Then
        ObtenirNomUtilisateur = Trim(Left(buffer, taille - 1))
    Else
        ObtenirNomUtilisateur = "Inconnu"
    End If
End Function

Public Function ObtenirNomOrdinateur() As String
    Dim buffer As String, taille As Long
    buffer = Space(256): taille = 256

    If GetComputerName(buffer, taille) <> 0 Then
        ObtenirNomOrdinateur = Left(buffer, taille)
    Else
        ObtenirNomOrdinateur = "Inconnu"
    End If
End Function

Public Sub AfficherInformationsSysteme()
    ' Collecte et affichage des informations système
    Dim rapport As String

    rapport = "========== INFORMATIONS SYSTÈME ==========" & vbCrLf
    rapport = rapport & "Utilisateur : " & ObtenirNomUtilisateur() & vbCrLf
    rapport = rapport & "Ordinateur : " & ObtenirNomOrdinateur() & vbCrLf
    rapport = rapport & "Résolution : " & GetSystemMetrics(SM_CXSCREEN) & " x " & GetSystemMetrics(SM_CYSCREEN) & vbCrLf
    rapport = rapport & "Écrans : " & GetSystemMetrics(SM_CMONITORS) & vbCrLf
    rapport = rapport & "Date/Heure : " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    rapport = rapport & "=========================================="

    ' Afficher dans une MsgBox
    MsgBox rapport, vbInformation, "Informations Système"

    ' Aussi dans le debug
    Debug.Print rapport

    ' Son de confirmation
    Beep 1000, 200
End Sub

Public Sub DeconnexionSecurisee()
    ' Démonstration d'une déconnexion avec compte à rebours
    Dim i As Integer
    Dim utilisateur As String

    utilisateur = ObtenirNomUtilisateur()

    ' Avertissement
    If MsgBox("Déconnexion dans 10 secondes pour " & utilisateur & vbCrLf & _
              "Voulez-vous continuer ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    ' Compte à rebours avec sons
    For i = 10 To 1 Step -1
        Debug.Print "Déconnexion dans " & i & " seconde(s)..."

        ' Son plus aigu à mesure qu'on approche de la fin
        Beep 500 + ((11 - i) * 50), 100

        ' Pause d'une seconde
        Sleep 1000
    Next i

    ' Son final
    Beep 1500, 500
    Debug.Print "Simulation de déconnexion pour " & utilisateur
End Sub
```

### Utilisation du module

```vba
Sub TestAPI()
    ' Test simple de toutes les API

    ' Informations de base
    Debug.Print "Utilisateur : " & ObtenirNomUtilisateur()
    Debug.Print "Ordinateur : " & ObtenirNomOrdinateur()

    ' Affichage complet
    AfficherInformationsSysteme

    ' Animation courte
    Debug.Print "Animation en cours..."
    Dim i As Integer
    For i = 1 To 5
        Debug.Print "Étape " & i
        Sleep 500
    Next i

    Debug.Print "Test terminé !"
    Beep 1200, 300  ' Son de fin
End Sub
```

## Bonnes pratiques avec ces API

### 1. Gestion des buffers
```vba
' ✅ Toujours initialiser les buffers
buffer = Space(256)  ' Taille suffisante

' ❌ Éviter les buffers trop petits
buffer = Space(10)   ' Risque de dépassement
```

### 2. Vérification des retours
```vba
' ✅ Toujours vérifier le succès
If GetUserName(buffer, taille) <> 0 Then
    ' Traitement en cas de succès
Else
    ' Gestion de l'échec
End If
```

### 3. Gestion des pauses
```vba
' ✅ Pauses raisonnables
Sleep 100   ' 0,1 seconde - correct

' ❌ Éviter les pauses excessives
Sleep 60000 ' 1 minute - bloque l'interface
```

### 4. Fonctions wrapper
```vba
' ✅ Créer des fonctions simples d'utilisation
Public Function NomUtilisateur() As String
    ' Logique complexe cachée
End Function

' ✅ Interface utilisateur claire
NomUtilisateur()  ' Simple à utiliser
```

Ces API de base constituent une excellente introduction au monde des API Windows. Elles sont sûres, utiles et vous donnent un aperçu de la puissance disponible dans Windows sans risquer de compromettre la stabilité de votre système.

⏭️
