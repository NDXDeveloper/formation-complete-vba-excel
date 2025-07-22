🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.3. Manipulation du registre

## Qu'est-ce que le registre Windows ?

Le **registre Windows** est une base de données hiérarchique qui stocke les paramètres de configuration de Windows et des applications installées. C'est le "centre de contrôle" du système d'exploitation.

**Analogie simple :**
Imaginez le registre comme un **immense classeur** avec des dossiers et sous-dossiers :
- **Ruches principales** = Classeurs principaux (HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)
- **Clés** = Dossiers et sous-dossiers
- **Valeurs** = Fiches individuelles contenant des informations
- **Données** = Le contenu écrit sur chaque fiche

Exemple : Pour stocker la couleur préférée d'un utilisateur
- **Chemin** : `HKEY_CURRENT_USER\Software\MonApp\Preferences`
- **Nom de valeur** : `CouleurPreferee`
- **Données** : `Bleu`

## Structure du registre

### Ruches principales (Hives)

#### HKEY_CURRENT_USER (HKCU)
- **Usage** : Paramètres spécifiques à l'utilisateur actuel
- **Exemples** : Préférences d'applications, fond d'écran, paramètres personnels
- **Sécurité** : Lecture/écriture pour l'utilisateur courant

#### HKEY_LOCAL_MACHINE (HKLM)
- **Usage** : Paramètres globaux du système
- **Exemples** : Logiciels installés, drivers, configuration hardware
- **Sécurité** : Généralement lecture seule pour les utilisateurs standards

#### HKEY_CLASSES_ROOT (HKCR)
- **Usage** : Associations de fichiers et informations COM
- **Exemples** : Programmes par défaut pour .xlsx, .docx

#### Autres ruches (moins utilisées)
- **HKEY_USERS** : Tous les profils utilisateurs
- **HKEY_CURRENT_CONFIG** : Configuration matérielle actuelle

### Types de données du registre

| Type VBA | Type Registre | Description | Exemple |
|----------|---------------|-------------|---------|
| String | REG_SZ | Chaîne de caractères | "MonApplication" |
| Long | REG_DWORD | Nombre entier 32 bits | 12345 |
| String | REG_EXPAND_SZ | Chaîne avec variables | "%USERPROFILE%\Documents" |
| Binary | REG_BINARY | Données binaires | Configurations complexes |

## ⚠️ Précautions importantes

### Risques de la manipulation du registre
- **Instabilité système** : Modifications incorrectes peuvent rendre Windows instable
- **Perte de données** : Suppression accidentelle de clés importantes
- **Problèmes de sécurité** : Accès non autorisé aux paramètres système

### Règles de sécurité essentielles
1. **Toujours sauvegarder** avant toute modification
2. **Travailler uniquement** dans HKEY_CURRENT_USER pour débuter
3. **Éviter HKEY_LOCAL_MACHINE** sauf nécessité absolue
4. **Tester** sur un environnement de développement
5. **Valider** tous les paramètres avant écriture

## Déclarations d'API pour le registre

### API principales

```vba
' Déclarations compatibles 32/64 bits
#If VBA7 Then
    ' Ouvrir une clé
    Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long

    ' Lire une valeur
    Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, _
        lpcbData As Long) As Long

    ' Écrire une valeur
    Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, _
        ByVal cbData As Long) As Long

    ' Créer une clé
    Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
        ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
        phkResult As LongPtr, lpdwDisposition As Long) As Long

    ' Fermer une clé
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As LongPtr) As Long

    ' Supprimer une valeur
    Private Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" _
        Alias "RegDeleteValueA" (ByVal hKey As LongPtr, ByVal lpValueName As String) As Long
#Else
    ' Versions 32 bits (similaires mais sans PtrSafe)
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

    ' ... autres déclarations similaires
#End If
```

### Constantes importantes

```vba
' Ruches principales
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CLASSES_ROOT = &H80000000

' Droits d'accès
Private Const KEY_READ = &H20019
Private Const KEY_WRITE = &H20006
Private Const KEY_ALL_ACCESS = &HF003F

' Types de données
Private Const REG_SZ = 1          ' Chaîne
Private Const REG_DWORD = 4       ' Nombre entier

' Codes de retour
Private Const ERROR_SUCCESS = 0
Private Const ERROR_FILE_NOT_FOUND = 2
```

## Alternative sécurisée : WScript.Shell

### Méthode recommandée pour débuter

Avant d'utiliser les API complexes, VBA offre une alternative plus simple et sûre via l'objet `WScript.Shell` :

```vba
Sub ExempleWScriptShell()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")

    ' Lire une valeur
    Dim valeur As String
    On Error Resume Next
    valeur = objShell.RegRead("HKCU\Software\MonApp\MaValeur")
    If Err.Number = 0 Then
        Debug.Print "Valeur lue : " & valeur
    Else
        Debug.Print "Valeur non trouvée"
        Err.Clear
    End If

    ' Écrire une valeur
    objShell.RegWrite "HKCU\Software\MonApp\MaValeur", "NouvelleDonnee", "REG_SZ"
    Debug.Print "Valeur écrite"

    ' Supprimer une valeur
    On Error Resume Next
    objShell.RegDelete "HKCU\Software\MonApp\MaValeur"
    If Err.Number = 0 Then
        Debug.Print "Valeur supprimée"
    End If

    Set objShell = Nothing
End Sub
```

## Fonctions wrapper avec API natives

### Classe RegistreHelper pour simplifier l'usage

```vba
' ================================================================
' Module : RegistreHelper
' Description : Wrapper sécurisé pour les API du registre
' ================================================================

Option Explicit

' Déclarations d'API (voir section précédente pour la version complète)

Public Function LireValeurRegistre(ruche As LongPtr, chemin As String, nomValeur As String) As Variant
    ' Lit une valeur du registre de manière sécurisée

    Dim hKey As LongPtr
    Dim resultat As Long
    Dim typeValeur As Long
    Dim buffer As String
    Dim tailleBuffer As Long

    ' Initialiser le résultat
    LireValeurRegistre = Empty

    ' Ouvrir la clé en lecture seule
    resultat = RegOpenKeyEx(ruche, chemin, 0, KEY_READ, hKey)
    If resultat <> ERROR_SUCCESS Then
        Debug.Print "Erreur ouverture clé : " & chemin
        Exit Function
    End If

    ' Déterminer la taille nécessaire
    resultat = RegQueryValueEx(hKey, nomValeur, 0, typeValeur, vbNullString, tailleBuffer)
    If resultat <> ERROR_SUCCESS Then
        RegCloseKey hKey
        Debug.Print "Valeur non trouvée : " & nomValeur
        Exit Function
    End If

    ' Créer un buffer et lire la valeur
    buffer = Space(tailleBuffer)
    resultat = RegQueryValueEx(hKey, nomValeur, 0, typeValeur, buffer, tailleBuffer)

    ' Fermer la clé
    RegCloseKey hKey

    If resultat = ERROR_SUCCESS Then
        Select Case typeValeur
            Case REG_SZ
                ' Chaîne de caractères - retirer le caractère null
                LireValeurRegistre = Left(buffer, tailleBuffer - 1)
            Case REG_DWORD
                ' Nombre entier - conversion nécessaire
                LireValeurRegistre = CLng(buffer)
            Case Else
                ' Type non géré
                LireValeurRegistre = buffer
        End Select
    End If
End Function

Public Function EcrireValeurRegistre(ruche As LongPtr, chemin As String, nomValeur As String, valeur As Variant, Optional typeReg As Long = REG_SZ) As Boolean
    ' Écrit une valeur dans le registre

    Dim hKey As LongPtr
    Dim resultat As Long
    Dim disposition As Long
    Dim donnees As String

    ' Initialiser le résultat
    EcrireValeurRegistre = False

    ' Créer ou ouvrir la clé
    resultat = RegCreateKeyEx(ruche, chemin, 0, vbNullString, 0, KEY_WRITE, 0, hKey, disposition)
    If resultat <> ERROR_SUCCESS Then
        Debug.Print "Erreur création/ouverture clé : " & chemin
        Exit Function
    End If

    ' Préparer les données selon le type
    Select Case typeReg
        Case REG_SZ
            donnees = CStr(valeur) & Chr(0)  ' Ajouter le caractère null
        Case REG_DWORD
            donnees = CStr(CLng(valeur))
        Case Else
            donnees = CStr(valeur)
    End Select

    ' Écrire la valeur
    resultat = RegSetValueEx(hKey, nomValeur, 0, typeReg, donnees, Len(donnees))

    ' Fermer la clé
    RegCloseKey hKey

    ' Vérifier le succès
    If resultat = ERROR_SUCCESS Then
        EcrireValeurRegistre = True
        Debug.Print "Valeur écrite : " & nomValeur & " = " & valeur
    Else
        Debug.Print "Erreur écriture : " & nomValeur
    End If
End Function

Public Function SupprimerValeurRegistre(ruche As LongPtr, chemin As String, nomValeur As String) As Boolean
    ' Supprime une valeur du registre

    Dim hKey As LongPtr
    Dim resultat As Long

    ' Initialiser le résultat
    SupprimerValeurRegistre = False

    ' Ouvrir la clé en écriture
    resultat = RegOpenKeyEx(ruche, chemin, 0, KEY_WRITE, hKey)
    If resultat <> ERROR_SUCCESS Then
        Debug.Print "Erreur ouverture clé pour suppression : " & chemin
        Exit Function
    End If

    ' Supprimer la valeur
    resultat = RegDeleteValue(hKey, nomValeur)

    ' Fermer la clé
    RegCloseKey hKey

    ' Vérifier le succès
    If resultat = ERROR_SUCCESS Then
        SupprimerValeurRegistre = True
        Debug.Print "Valeur supprimée : " & nomValeur
    Else
        Debug.Print "Erreur suppression : " & nomValeur
    End If
End Function
```

## Classe de gestion des préférences application

### Exemple pratique complet

```vba
' ================================================================
' Module de classe : PreferencesApp
' Description : Gestion des préférences utilisateur via le registre
' ================================================================

Option Explicit

Private Const CHEMIN_BASE As String = "Software\MonApplication\Preferences"
Private mNomApplication As String

' Initialisation
Private Sub Class_Initialize()
    mNomApplication = "MonApplication"
End Sub

' Propriétés
Public Property Let NomApplication(valeur As String)
    mNomApplication = valeur
End Property

Public Property Get NomApplication() As String
    NomApplication = mNomApplication
End Property

' Méthodes de lecture
Public Function LireCouleur(Optional defaut As String = "Bleu") As String
    LireCouleur = Me.LirePreference("Couleur", defaut)
End Function

Public Function LireTaillePolice(Optional defaut As Long = 12) As Long
    LireTaillePolice = CLng(Me.LirePreference("TaillePolice", CStr(defaut)))
End Function

Public Function LireFenetreMaximisee(Optional defaut As Boolean = False) As Boolean
    Dim valeur As String
    valeur = Me.LirePreference("FenetreMaximisee", IIf(defaut, "1", "0"))
    LireFenetreMaximisee = (valeur = "1")
End Function

Public Function LireDernierFichier(Optional defaut As String = "") As String
    LireDernierFichier = Me.LirePreference("DernierFichier", defaut)
End Function

' Méthodes d'écriture
Public Sub SauvegarderCouleur(couleur As String)
    Me.EcrirePreference "Couleur", couleur
End Sub

Public Sub SauvegarderTaillePolice(taille As Long)
    Me.EcrirePreference "TaillePolice", CStr(taille)
End Sub

Public Sub SauvegarderFenetreMaximisee(maximisee As Boolean)
    Me.EcrirePreference "FenetreMaximisee", IIf(maximisee, "1", "0")
End Sub

Public Sub SauvegarderDernierFichier(chemin As String)
    Me.EcrirePreference "DernierFichier", chemin
End Sub

' Méthodes utilitaires
Public Sub ResetPreferences()
    ' Remet toutes les préférences par défaut
    Me.SauvegarderCouleur "Bleu"
    Me.SauvegarderTaillePolice 12
    Me.SauvegarderFenetreMaximisee False
    Me.SauvegarderDernierFichier ""

    Debug.Print "Préférences remises à zéro"
End Sub

Public Sub AfficherPreferences()
    ' Affiche toutes les préférences actuelles
    Debug.Print "=== PRÉFÉRENCES " & UCase(mNomApplication) & " ==="
    Debug.Print "Couleur : " & Me.LireCouleur()
    Debug.Print "Taille police : " & Me.LireTaillePolice()
    Debug.Print "Fenêtre maximisée : " & IIf(Me.LireFenetreMaximisee(), "Oui", "Non")
    Debug.Print "Dernier fichier : " & Me.LireDernierFichier()
    Debug.Print "=================================="
End Sub

Public Function ExporterPreferences() As String
    ' Exporte les préférences sous forme de texte
    Dim export As String

    export = "[" & mNomApplication & "]" & vbCrLf
    export = export & "Couleur=" & Me.LireCouleur() & vbCrLf
    export = export & "TaillePolice=" & Me.LireTaillePolice() & vbCrLf
    export = export & "FenetreMaximisee=" & IIf(Me.LireFenetreMaximisee(), "1", "0") & vbCrLf
    export = export & "DernierFichier=" & Me.LireDernierFichier() & vbCrLf

    ExporterPreferences = export
End Function

' Méthodes privées (implémentation)
Private Function LirePreference(nom As String, defaut As String) As String
    ' Version sécurisée avec WScript.Shell
    Dim objShell As Object
    Dim chemin As String

    Set objShell = CreateObject("WScript.Shell")
    chemin = "HKCU\" & CHEMIN_BASE & "\" & nom

    On Error Resume Next
    LirePreference = objShell.RegRead(chemin)
    If Err.Number <> 0 Then
        LirePreference = defaut
        Err.Clear
    End If

    Set objShell = Nothing
End Function

Private Sub EcrirePreference(nom As String, valeur As String)
    ' Version sécurisée avec WScript.Shell
    Dim objShell As Object
    Dim chemin As String

    Set objShell = CreateObject("WScript.Shell")
    chemin = "HKCU\" & CHEMIN_BASE & "\" & nom

    On Error Resume Next
    objShell.RegWrite chemin, valeur, "REG_SZ"
    If Err.Number <> 0 Then
        Debug.Print "Erreur écriture préférence : " & nom
        Err.Clear
    End If

    Set objShell = Nothing
End Sub
```

## Utilisation pratique des préférences

### Intégration dans une application

```vba
Sub TestPreferencesApplication()
    ' Créer l'objet de gestion des préférences
    Dim prefs As New PreferencesApp
    prefs.NomApplication = "MonExcelApp"

    ' Première utilisation - charger les préférences
    Debug.Print "=== CHARGEMENT DES PRÉFÉRENCES ==="
    prefs.AfficherPreferences

    ' Simuler l'utilisation de l'application
    Debug.Print "=== MODIFICATION DES PRÉFÉRENCES ==="
    prefs.SauvegarderCouleur "Rouge"
    prefs.SauvegarderTaillePolice 14
    prefs.SauvegarderFenetreMaximisee True
    prefs.SauvegarderDernierFichier "C:\Mes Documents\rapport.xlsx"

    ' Afficher les nouvelles préférences
    prefs.AfficherPreferences

    ' Exporter pour sauvegarde
    Debug.Print "=== EXPORT ==="
    Debug.Print prefs.ExporterPreferences()

    ' Simuler une réinitialisation
    If MsgBox("Voulez-vous remettre les préférences par défaut ?", vbYesNo) = vbYes Then
        prefs.ResetPreferences
        prefs.AfficherPreferences
    End If
End Sub

Sub ChargerPreferencesInterface()
    ' Exemple d'utilisation au démarrage d'une UserForm
    Dim prefs As New PreferencesApp

    ' Charger les préférences sauvegardées
    With prefs
        ' Adapter l'interface selon les préférences
        If .LireFenetreMaximisee() Then
            ' Maximiser la fenêtre Excel
            Application.WindowState = xlMaximized
        End If

        ' Adapter la couleur de fond (exemple)
        Select Case .LireCouleur()
            Case "Rouge"
                ' ActiveSheet.Tab.Color = RGB(255, 0, 0)
            Case "Bleu"
                ' ActiveSheet.Tab.Color = RGB(0, 0, 255)
        End Select

        ' Ouvrir le dernier fichier si disponible
        Dim dernierFichier As String
        dernierFichier = .LireDernierFichier()
        If dernierFichier <> "" And Dir(dernierFichier) <> "" Then
            If MsgBox("Ouvrir le dernier fichier : " & dernierFichier & " ?", vbYesNo) = vbYes Then
                ' Workbooks.Open dernierFichier
                Debug.Print "Ouverture simulée : " & dernierFichier
            End If
        End If
    End With
End Sub
```

## Bonnes pratiques pour le registre

### 1. Sécurité et permissions
```vba
' ✅ Travailler dans HKEY_CURRENT_USER (sûr)
chemin = "HKCU\Software\MonApp\Preferences"

' ❌ Éviter HKEY_LOCAL_MACHINE (nécessite droits admin)
' chemin = "HKLM\Software\MonApp"  ' Dangereux pour débutants
```

### 2. Gestion d'erreurs robuste
```vba
' ✅ Toujours gérer les erreurs
On Error Resume Next
valeur = objShell.RegRead(chemin)
If Err.Number <> 0 Then
    valeur = valeurParDefaut
    Err.Clear
End If
```

### 3. Nettoyage et organisation
```vba
' ✅ Structurer les clés logiquement
' HKCU\Software\[NomSociete]\[NomApp]\[Section]
' Exemple : HKCU\Software\MaSociete\MonApp\Interface
'          HKCU\Software\MaSociete\MonApp\Donnees
```

### 4. Sauvegarde avant modifications
```vba
Sub SauvegarderRegistre()
    ' Exporter une section avant modification
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")

    ' Commande d'export (attention : nécessite droits)
    Dim commande As String
    commande = "reg export ""HKCU\Software\MonApp"" ""C:\Backup\MonApp.reg"""

    ' Attention : à utiliser avec précaution
    ' objShell.Run commande, 0, True

    Set objShell = Nothing
End Sub
```

### 5. Validation des données
```vba
Public Sub EcrirePreferenceSecurisee(nom As String, valeur As String)
    ' Validation des paramètres
    If Len(nom) = 0 Or Len(nom) > 255 Then
        Err.Raise 5, , "Nom de préférence invalide"
        Exit Sub
    End If

    If Len(valeur) > 1024 Then
        Err.Raise 5, , "Valeur trop longue (max 1024 caractères)"
        Exit Sub
    End If

    ' Caractères interdits dans les noms de valeurs
    If InStr(nom, "\") > 0 Or InStr(nom, "/") > 0 Then
        Err.Raise 5, , "Caractères interdits dans le nom"
        Exit Sub
    End If

    ' Écriture sécurisée
    Me.EcrirePreference nom, valeur
End Sub
```

## Alternatives modernes

### 1. Fichiers de configuration
```vba
' Alternative : fichier INI ou JSON
' Plus portable, plus facile à sauvegarder
Sub SauvegarderDansFichier()
    Dim fichierConfig As String
    fichierConfig = Environ("APPDATA") & "\MonApp\config.txt"

    ' Écriture dans un fichier texte simple
    ' Avantage : pas de risque pour le système
End Sub
```

### 2. Propriétés du document
```vba
' Pour Excel : utiliser les propriétés personnalisées
Sub UtiliserProprietesDocument()
    ' Sauvegarder dans le fichier Excel lui-même
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:="MaPreference", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:="MaValeur"
End Sub
```

La manipulation du registre est un outil puissant mais qui nécessite des précautions. Commencez toujours par les méthodes les plus sûres (WScript.Shell) et ne passez aux API natives que si vous avez besoin de fonctionnalités avancées spécifiques.

⏭️
