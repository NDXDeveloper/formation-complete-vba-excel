🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.4. Interaction avec le système

## Introduction

L'interaction avec le système permet à vos applications VBA de communiquer avec Windows et d'autres programmes, de gérer les fenêtres, de lancer des processus et de surveiller l'activité système. C'est un niveau avancé qui transforme VBA en véritable outil d'automatisation système.

**Analogie simple :**
Imaginez que votre application VBA est comme un **chef d'orchestre** :
- **Gestion des fenêtres** = Diriger les musiciens (applications) sur scène
- **Lancement de processus** = Inviter de nouveaux musiciens à rejoindre l'orchestre
- **Surveillance système** = Écouter et coordonner l'ensemble
- **Communication inter-applications** = Faire jouer les instruments ensemble

## Domaines d'interaction système

### 1. Gestion des fenêtres
- Trouver des fenêtres d'autres applications
- Déplacer, redimensionner, minimiser/maximiser
- Mettre au premier plan, cacher des fenêtres
- Envoyer des messages à d'autres applications

### 2. Gestion des processus
- Lancer des programmes externes
- Attendre la fin d'exécution
- Terminer des processus
- Obtenir des informations sur les processus

### 3. Système de fichiers avancé
- Surveiller les changements dans des dossiers
- Obtenir des informations détaillées sur les fichiers
- Gérer les attributs et permissions
- Opérations sur les raccourcis

### 4. Informations système
- État de la mémoire et du processeur
- Informations sur les lecteurs
- Services Windows
- Variables d'environnement avancées

## 1. Gestion des fenêtres

### Déclarations d'API pour les fenêtres

```vba
' API pour la gestion des fenêtres
#If VBA7 Then
    ' Trouver une fenêtre
    Private Declare PtrSafe Function FindWindow Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    ' Trouver une fenêtre enfant
    Private Declare PtrSafe Function FindWindowEx Lib "user32" _
        Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr

    ' Obtenir le titre d'une fenêtre
    Private Declare PtrSafe Function GetWindowText Lib "user32" _
        Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long

    ' Obtenir les dimensions d'une fenêtre
    Private Declare PtrSafe Function GetWindowRect Lib "user32" _
        (ByVal hwnd As LongPtr, lpRect As RECT) As Long

    ' Déplacer/redimensionner une fenêtre
    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

    ' Montrer/cacher une fenêtre
    Private Declare PtrSafe Function ShowWindow Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long

    ' Mettre une fenêtre au premier plan
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Long

    ' Envoyer des messages à une fenêtre
    Private Declare PtrSafe Function SendMessage Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, _
        ByVal wParam As LongPtr, lParam As Any) As LongPtr
#Else
    ' Versions 32 bits (remplacer LongPtr par Long)
    ' ... déclarations similaires
#End If

' Structure pour les coordonnées de fenêtre
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Constantes pour ShowWindow
Private Const SW_HIDE = 0  
Private Const SW_NORMAL = 1  
Private Const SW_MINIMIZED = 2  
Private Const SW_MAXIMIZED = 3  
Private Const SW_RESTORE = 9  

' Constantes pour SetWindowPos
Private Const HWND_TOP = 0  
Private Const HWND_TOPMOST = -1  
Private Const SWP_NOMOVE = &H2  
Private Const SWP_NOSIZE = &H1  

' API pour les métriques système (utilisée par CentrerFenetre)
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
#End If
```

### Classe GestionnaireFenetres

```vba
' ================================================================
' Module de classe : GestionnaireFenetres
' Description : Gestion simplifiée des fenêtres système
' ================================================================

Option Explicit

Public Function TrouverFenetreParTitre(titre As String, Optional exact As Boolean = False) As LongPtr
    ' Trouve une fenêtre par son titre
    Dim hwnd As LongPtr

    If exact Then
        ' Recherche exacte
        hwnd = FindWindow(vbNullString, titre)
    Else
        ' Recherche partielle (plus complexe, nécessite énumération)
        hwnd = Me.ChercherFenetrePartielle(titre)
    End If

    TrouverFenetreParTitre = hwnd

    If hwnd <> 0 Then
        Debug.Print "Fenêtre trouvée : " & titre & " (Handle: " & hwnd & ")"
    Else
        Debug.Print "Fenêtre non trouvée : " & titre
    End If
End Function

Public Function ObtenirTitreFenetre(hwnd As LongPtr) As String
    ' Obtient le titre d'une fenêtre
    Dim buffer As String
    Dim longueur As Long

    buffer = Space(255)
    longueur = GetWindowText(hwnd, buffer, 255)

    If longueur > 0 Then
        ObtenirTitreFenetre = Left(buffer, longueur)
    Else
        ObtenirTitreFenetre = ""
    End If
End Function

Public Function ObtenirPositionFenetre(hwnd As LongPtr) As RECT
    ' Obtient la position et taille d'une fenêtre
    Dim rect As RECT

    If GetWindowRect(hwnd, rect) <> 0 Then
        ObtenirPositionFenetre = rect
        Debug.Print "Position fenêtre - Gauche:" & rect.Left & " Haut:" & rect.Top & _
                   " Largeur:" & (rect.Right - rect.Left) & " Hauteur:" & (rect.Bottom - rect.Top)
    End If
End Function

Public Sub DeplacerFenetre(hwnd As LongPtr, x As Long, y As Long, Optional largeur As Long = -1, Optional hauteur As Long = -1)
    ' Déplace et/ou redimensionne une fenêtre
    Dim flags As Long

    If largeur = -1 Or hauteur = -1 Then
        ' Déplacement seulement
        flags = SWP_NOSIZE
        SetWindowPos hwnd, HWND_TOP, x, y, 0, 0, flags
        Debug.Print "Fenêtre déplacée vers (" & x & ", " & y & ")"
    Else
        ' Déplacement et redimensionnement
        SetWindowPos hwnd, HWND_TOP, x, y, largeur, hauteur, 0
        Debug.Print "Fenêtre déplacée et redimensionnée : (" & x & ", " & y & ") " & largeur & "x" & hauteur
    End If
End Sub

Public Sub AfficherFenetre(hwnd As LongPtr, mode As Long)
    ' Affiche une fenêtre selon le mode spécifié
    ShowWindow hwnd, mode

    Select Case mode
        Case SW_HIDE: Debug.Print "Fenêtre cachée"
        Case SW_NORMAL: Debug.Print "Fenêtre restaurée"
        Case SW_MINIMIZED: Debug.Print "Fenêtre réduite"
        Case SW_MAXIMIZED: Debug.Print "Fenêtre agrandie"
        Case SW_RESTORE: Debug.Print "Fenêtre restaurée"
    End Select
End Sub

Public Sub MettreAuPremierPlan(hwnd As LongPtr)
    ' Met une fenêtre au premier plan
    SetForegroundWindow hwnd
    Debug.Print "Fenêtre mise au premier plan"
End Sub

Public Sub CentrerFenetre(hwnd As LongPtr)
    ' Centre une fenêtre sur l'écran
    Dim rect As RECT
    Dim largeurEcran As Long, hauteurEcran As Long
    Dim largeurFenetre As Long, hauteurFenetre As Long
    Dim x As Long, y As Long

    ' Obtenir la taille de l'écran
    largeurEcran = GetSystemMetrics(0)  ' SM_CXSCREEN
    hauteurEcran = GetSystemMetrics(1)  ' SM_CYSCREEN

    ' Obtenir la taille de la fenêtre
    rect = Me.ObtenirPositionFenetre(hwnd)
    largeurFenetre = rect.Right - rect.Left
    hauteurFenetre = rect.Bottom - rect.Top

    ' Calculer la position centrée
    x = (largeurEcran - largeurFenetre) \ 2
    y = (hauteurEcran - hauteurFenetre) \ 2

    ' Déplacer la fenêtre
    Me.DeplacerFenetre hwnd, x, y
End Sub

Private Function ChercherFenetrePartielle(titrePartiel As String) As LongPtr
    ' Recherche une fenêtre dont le titre contient le texte spécifié
    ' Note : Implémentation simplifiée - en réalité nécessiterait EnumWindows

    ' Pour l'exemple, on essaie quelques applications courantes
    Dim fenetresPossibles As Variant
    Dim i As Integer
    Dim hwnd As LongPtr
    Dim titre As String

    fenetresPossibles = Array("Notepad", "Calculator", "Microsoft Excel", "Microsoft Word")

    For i = 0 To UBound(fenetresPossibles)
        hwnd = FindWindow(vbNullString, fenetresPossibles(i))
        If hwnd <> 0 Then
            titre = Me.ObtenirTitreFenetre(hwnd)
            If InStr(UCase(titre), UCase(titrePartiel)) > 0 Then
                ChercherFenetrePartielle = hwnd
                Exit Function
            End If
        End If
    Next i

    ChercherFenetrePartielle = 0
End Function
```

## 2. Gestion des processus

### Déclarations d'API pour les processus

```vba
' API pour la gestion des processus
#If VBA7 Then
    ' Créer un processus
    Private Declare PtrSafe Function CreateProcess Lib "kernel32" _
        Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

    ' Attendre la fin d'un processus
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long

    ' Fermer un handle
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" _
        (ByVal hObject As LongPtr) As Long

    ' Terminer un processus
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" _
        (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long

    ' Obtenir le code de sortie
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
#Else
    ' Versions 32 bits...
#End If

' Structures pour les processus
Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

' Constantes
Private Const INFINITE = &HFFFFFFFF  
Private Const WAIT_TIMEOUT = &H102  
Private Const WAIT_OBJECT_0 = 0  
```

### Classe GestionnaireProcessus

```vba
' ================================================================
' Module de classe : GestionnaireProcessus
' Description : Gestion des processus système
' ================================================================

Option Explicit

Public Function LancerProgramme(cheminProgramme As String, Optional parametres As String = "", Optional attendre As Boolean = False, Optional visible As Boolean = True) As Long
    ' Lance un programme externe avec options avancées

    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim commandeComplete As String
    Dim resultat As Long
    Dim codeRetour As Long

    ' Préparer la commande
    If parametres <> "" Then
        commandeComplete = """" & cheminProgramme & """ " & parametres
    Else
        commandeComplete = """" & cheminProgramme & """"
    End If

    ' Initialiser la structure STARTUPINFO
    si.cb = Len(si)
    If Not visible Then
        si.dwFlags = 1  ' STARTF_USESHOWWINDOW
        si.wShowWindow = SW_HIDE
    End If

    ' Lancer le processus
    resultat = CreateProcess(vbNullString, commandeComplete, 0, 0, 0, 0, 0, vbNullString, si, pi)

    If resultat = 0 Then
        Debug.Print "Erreur lors du lancement : " & cheminProgramme
        LancerProgramme = -1
        Exit Function
    End If

    Debug.Print "Programme lancé : " & cheminProgramme & " (PID: " & pi.dwProcessId & ")"

    If attendre Then
        ' Attendre la fin du processus
        Debug.Print "Attente de la fin du processus..."
        WaitForSingleObject pi.hProcess, INFINITE

        ' Obtenir le code de retour
        GetExitCodeProcess pi.hProcess, codeRetour
        Debug.Print "Processus terminé avec le code : " & codeRetour

        LancerProgramme = codeRetour
    Else
        LancerProgramme = pi.dwProcessId
    End If

    ' Nettoyer les handles
    CloseHandle pi.hProcess
    CloseHandle pi.hThread
End Function

Public Function LancerProgrammeSimple(commande As String, Optional avecFocus As Boolean = True) As Boolean
    ' Version simplifiée utilisant Shell VBA
    ' Note : Shell VBA retourne toujours immédiatement (pas d'attente).
    ' Pour attendre la fin d'un processus, utilisez LancerProgramme avec attendre:=True.
    On Error GoTo GestionErreur

    Dim style As Integer
    style = IIf(avecFocus, vbNormalFocus, vbNormalNoFocus)

    Shell commande, style
    LancerProgrammeSimple = True

    Debug.Print "Commande exécutée : " & commande
    Exit Function

GestionErreur:
    Debug.Print "Erreur Shell : " & Err.Description
    LancerProgrammeSimple = False
End Function

Public Sub OuvrirFichierAvec(cheminFichier As String, Optional programme As String = "")
    ' Ouvre un fichier avec le programme spécifié ou par défaut
    Dim commande As String

    If programme <> "" Then
        commande = """" & programme & """ """ & cheminFichier & """"
    Else
        ' Utiliser l'association par défaut
        commande = """" & cheminFichier & """"
    End If

    Me.LancerProgrammeSimple commande
    Debug.Print "Fichier ouvert : " & cheminFichier
End Sub

Public Sub OuvrirDossier(cheminDossier As String)
    ' Ouvre un dossier dans l'Explorateur Windows
    Dim commande As String
    commande = "explorer.exe """ & cheminDossier & """"

    Me.LancerProgrammeSimple commande
    Debug.Print "Dossier ouvert : " & cheminDossier
End Sub

Public Sub OuvrirSiteWeb(url As String)
    ' Ouvre une URL dans le navigateur par défaut

    ' Ajouter http:// si nécessaire
    If Left(LCase(url), 7) <> "http://" And Left(LCase(url), 8) <> "https://" Then
        url = "http://" & url
    End If

    ' Shell ne peut pas ouvrir une URL directement,
    ' on utilise explorer.exe qui la transmet au navigateur par défaut
    Dim commande As String
    commande = "explorer.exe """ & url & """"
    Me.LancerProgrammeSimple commande
    Debug.Print "Site web ouvert : " & url
End Sub
```

## 3. Surveillance du système de fichiers

### Surveillance de dossiers

```vba
' ================================================================
' Module de classe : SurveillanceFichiers
' Description : Surveillance des changements dans les dossiers
' ================================================================

Option Explicit

' API pour la surveillance de fichiers
#If VBA7 Then
    Private Declare PtrSafe Function FindFirstChangeNotification Lib "kernel32" _
        Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, _
        ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As LongPtr

    Private Declare PtrSafe Function FindNextChangeNotification Lib "kernel32" _
        (ByVal hChangeHandle As LongPtr) As Long

    Private Declare PtrSafe Function FindCloseChangeNotification Lib "kernel32" _
        (ByVal hChangeHandle As LongPtr) As Long

    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
#Else
    ' Versions 32 bits...
#End If

' Constantes pour la surveillance
Private Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1  
Private Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2  
Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4  
Private Const FILE_NOTIFY_CHANGE_SIZE = &H8  
Private Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10  

Public Sub SurveillerDossier(cheminDossier As String, Optional dureeSecondes As Long = 60)
    ' Surveille les changements dans un dossier pendant une durée donnée

    Dim hNotification As LongPtr
    Dim filtre As Long
    Dim resultat As Long
    Dim finSurveillance As Date

    ' Configuration du filtre de surveillance
    filtre = FILE_NOTIFY_CHANGE_FILE_NAME Or FILE_NOTIFY_CHANGE_LAST_WRITE

    ' Démarrer la surveillance
    hNotification = FindFirstChangeNotification(cheminDossier, 0, filtre)

    If hNotification = -1 Then
        Debug.Print "Erreur : Impossible de surveiller le dossier " & cheminDossier
        Exit Sub
    End If

    Debug.Print "Surveillance démarrée pour : " & cheminDossier
    Debug.Print "Durée : " & dureeSecondes & " secondes"

    finSurveillance = DateAdd("s", dureeSecondes, Now)

    Do While Now < finSurveillance
        ' Attendre un changement (timeout de 1 seconde)
        resultat = WaitForSingleObject(hNotification, 1000)

        If resultat = WAIT_OBJECT_0 Then
            ' Un changement a été détecté
            Debug.Print Format(Now, "hh:nn:ss") & " - Changement détecté dans : " & cheminDossier

            ' Analyser les changements (implémentation simplifiée)
            Me.AnalyserChangements cheminDossier

            ' Préparer la prochaine notification
            FindNextChangeNotification hNotification
        End If

        ' Permettre à Excel de traiter d'autres événements
        DoEvents
    Loop

    ' Arrêter la surveillance
    FindCloseChangeNotification hNotification
    Debug.Print "Surveillance terminée"
End Sub

Private Sub AnalyserChangements(dossier As String)
    ' Analyse simplifiée des changements
    ' En pratique, il faudrait comparer l'état avant/après

    Dim fichier As String
    Dim compteur As Integer

    fichier = Dir(dossier & "\*.*")
    Do While fichier <> ""
        compteur = compteur + 1
        fichier = Dir
    Loop

    Debug.Print "  → " & compteur & " fichier(s) dans le dossier"
End Sub
```

## 4. Informations système avancées

### Classe InformationsSysteme

```vba
' ================================================================
' Module de classe : InformationsSystemeAvancees
' Description : Informations détaillées sur le système
' ================================================================

Option Explicit

' APIs pour les informations système
#If VBA7 Then
    Private Declare PtrSafe Function GetDiskFreeSpaceEx Lib "kernel32" _
        Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, _
        lpFreeBytesAvailable As Currency, lpTotalNumberOfBytes As Currency, _
        lpTotalNumberOfFreeBytes As Currency) As Long

    Private Declare PtrSafe Function GetDriveType Lib "kernel32" _
        Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

    Private Declare PtrSafe Function GetLogicalDriveStrings Lib "kernel32" _
        Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
#Else
    ' Versions 32 bits...
#End If

' Constantes pour les types de lecteurs
Private Const DRIVE_UNKNOWN = 0  
Private Const DRIVE_NO_ROOT_DIR = 1  
Private Const DRIVE_REMOVABLE = 2  
Private Const DRIVE_FIXED = 3  
Private Const DRIVE_REMOTE = 4  
Private Const DRIVE_CDROM = 5  
Private Const DRIVE_RAMDISK = 6  

Public Function ObtenirLecteurs() As Collection
    ' Obtient la liste de tous les lecteurs
    Dim lecteurs As New Collection
    Dim buffer As String
    Dim longueur As Long
    Dim position As Integer
    Dim lecteur As String

    buffer = Space(255)
    longueur = GetLogicalDriveStrings(255, buffer)

    position = 1
    Do While position < longueur
        lecteur = Mid(buffer, position, 3)  ' Format "C:\"
        If lecteur <> "" Then
            lecteurs.Add lecteur
            position = position + 4  ' "C:\" + null = 4 caractères
        Else
            Exit Do
        End If
    Loop

    Set ObtenirLecteurs = lecteurs
End Function

Public Function ObtenirTypeLecteur(lecteur As String) As String
    ' Obtient le type d'un lecteur
    Dim typeLecteur As Long

    typeLecteur = GetDriveType(lecteur)

    Select Case typeLecteur
        Case DRIVE_UNKNOWN: ObtenirTypeLecteur = "Inconnu"
        Case DRIVE_NO_ROOT_DIR: ObtenirTypeLecteur = "Invalide"
        Case DRIVE_REMOVABLE: ObtenirTypeLecteur = "Amovible"
        Case DRIVE_FIXED: ObtenirTypeLecteur = "Disque dur"
        Case DRIVE_REMOTE: ObtenirTypeLecteur = "Réseau"
        Case DRIVE_CDROM: ObtenirTypeLecteur = "CD-ROM/DVD"
        Case DRIVE_RAMDISK: ObtenirTypeLecteur = "RAM Disk"
        Case Else: ObtenirTypeLecteur = "Autre"
    End Select
End Function

Public Function ObtenirEspaceDisque(lecteur As String) As String
    ' Obtient les informations d'espace disque
    Dim espaceLibre As Currency
    Dim espaceTotal As Currency
    Dim espaceLibreTotal As Currency
    Dim resultat As Long

    resultat = GetDiskFreeSpaceEx(lecteur, espaceLibre, espaceTotal, espaceLibreTotal)

    ' Note : le type Currency divise automatiquement par 10000.
    ' Il faut donc multiplier par 10000 pour obtenir la taille réelle en octets.
    If resultat <> 0 Then
        Dim rapport As String
        rapport = "Lecteur " & lecteur & vbCrLf
        rapport = rapport & "Total : " & Me.FormatTaille(espaceTotal * 10000) & vbCrLf
        rapport = rapport & "Libre : " & Me.FormatTaille(espaceLibre * 10000) & vbCrLf
        rapport = rapport & "Utilisé : " & Me.FormatTaille((espaceTotal - espaceLibre) * 10000) & vbCrLf
        rapport = rapport & "% Libre : " & Format((espaceLibre / espaceTotal) * 100, "0.0") & "%"

        ObtenirEspaceDisque = rapport
    Else
        ObtenirEspaceDisque = "Erreur lors de la lecture de " & lecteur
    End If
End Function

Private Function FormatTaille(taille As Currency) As String
    ' Formate une taille en octets vers une unité lisible
    Dim KB As Currency: KB = 1024@
    Dim MB As Currency: MB = 1048576@       ' 1024 * 1024
    Dim GB As Currency: GB = 1073741824@    ' 1024 * 1024 * 1024

    If taille >= GB Then
        FormatTaille = Format(taille / GB, "0.0") & " Go"
    ElseIf taille >= MB Then
        FormatTaille = Format(taille / MB, "0.0") & " Mo"
    ElseIf taille >= KB Then
        FormatTaille = Format(taille / KB, "0.0") & " Ko"
    Else
        FormatTaille = Format(taille, "0") & " octets"
    End If
End Function

Public Sub AfficherRapportSysteme()
    ' Génère un rapport complet du système
    Dim rapport As String
    Dim lecteurs As Collection
    Dim lecteur As Variant

    rapport = "========== RAPPORT SYSTÈME ==========" & vbCrLf
    rapport = rapport & "Date : " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    rapport = rapport & "Utilisateur : " & Environ("USERNAME") & vbCrLf
    rapport = rapport & "Ordinateur : " & Environ("COMPUTERNAME") & vbCrLf
    rapport = rapport & "OS : " & Environ("OS") & vbCrLf & vbCrLf

    rapport = rapport & "=== LECTEURS ===" & vbCrLf
    Set lecteurs = Me.ObtenirLecteurs()

    For Each lecteur In lecteurs
        rapport = rapport & "Lecteur " & lecteur & " (" & Me.ObtenirTypeLecteur(lecteur) & ")" & vbCrLf
        If Me.ObtenirTypeLecteur(lecteur) = "Disque dur" Then
            rapport = rapport & Me.ObtenirEspaceDisque(lecteur) & vbCrLf
        End If
        rapport = rapport & vbCrLf
    Next lecteur

    rapport = rapport & "====================================="

    Debug.Print rapport

    ' Aussi dans un fichier texte
    Me.SauvegarderRapport rapport
End Sub

Private Sub SauvegarderRapport(contenu As String)
    ' Sauvegarde le rapport dans un fichier
    Dim fichier As String
    Dim numeroFichier As Integer

    fichier = Environ("TEMP") & "\RapportSysteme_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    numeroFichier = FreeFile

    Open fichier For Output As numeroFichier
    Print #numeroFichier, contenu
    Close numeroFichier

    Debug.Print "Rapport sauvegardé : " & fichier
End Sub
```

## Application complète : Gestionnaire système

### Module principal d'automatisation

```vba
Sub DemonstrationInteractionSysteme()
    ' Démonstration complète des interactions système

    Debug.Print "=== DÉMONSTRATION INTERACTION SYSTÈME ==="

    ' 1. Gestion des fenêtres
    Debug.Print vbCrLf & "1. GESTION DES FENÊTRES"
    Dim gestFenetres As New GestionnaireFenetres

    ' Lancer le Bloc-notes
    Shell "notepad.exe", vbNormalFocus
    Sleep 2000  ' Attendre que la fenêtre s'ouvre

    ' Trouver et manipuler la fenêtre du Bloc-notes
    Dim hwndNotepad As LongPtr
    hwndNotepad = gestFenetres.TrouverFenetreParTitre("Bloc-notes", True)

    If hwndNotepad <> 0 Then
        ' Déplacer et redimensionner la fenêtre
        gestFenetres.DeplacerFenetre hwndNotepad, 100, 100, 800, 600
        Sleep 1000

        ' Centrer la fenêtre
        gestFenetres.CentrerFenetre hwndNotepad
        Sleep 1000

        ' Minimiser puis restaurer
        gestFenetres.AfficherFenetre hwndNotepad, SW_MINIMIZED
        Sleep 2000
        gestFenetres.AfficherFenetre hwndNotepad, SW_RESTORE
        gestFenetres.MettreAuPremierPlan hwndNotepad
    End If

    ' 2. Gestion des processus
    Debug.Print vbCrLf & "2. GESTION DES PROCESSUS"
    Dim gestProcessus As New GestionnaireProcessus

    ' Lancer la Calculatrice et attendre sa fermeture
    Dim pidCalculatrice As Long
    pidCalculatrice = gestProcessus.LancerProgramme("calc.exe", "", False, True)
    Debug.Print "Calculatrice lancée (PID: " & pidCalculatrice & ")"

    ' Ouvrir quelques fichiers et dossiers
    gestProcessus.OuvrirDossier Environ("USERPROFILE") & "\Documents"
    Sleep 1000
    gestProcessus.OuvrirSiteWeb "www.microsoft.com"

    ' 3. Informations système
    Debug.Print vbCrLf & "3. INFORMATIONS SYSTÈME"
    Dim infosSys As New InformationsSystemeAvancees
    infosSys.AfficherRapportSysteme

    ' 4. Surveillance de fichiers (optionnel - commenté pour éviter le blocage)
    ' Debug.Print vbCrLf & "4. SURVEILLANCE FICHIERS"
    ' Dim surveillance As New SurveillanceFichiers
    ' surveillance.SurveillerDossier Environ("USERPROFILE") & "\Desktop", 30

    Debug.Print vbCrLf & "=== DÉMONSTRATION TERMINÉE ==="
End Sub
```

### Utilitaires système pratiques

```vba
Sub NettoyageSystemeAutomatique()
    ' Automatisation de tâches de nettoyage système

    Debug.Print "=== NETTOYAGE SYSTÈME AUTOMATIQUE ==="

    Dim gestProcessus As New GestionnaireProcessus

    ' 1. Vider la corbeille (avec confirmation)
    If MsgBox("Vider la corbeille ?", vbYesNo + vbQuestion) = vbYes Then
        gestProcessus.LancerProgrammeSimple "cmd /c rd /s /q %systemdrive%\$Recycle.bin", True
        Debug.Print "Corbeille vidée"
    End If

    ' 2. Nettoyer les fichiers temporaires
    Debug.Print "Nettoyage des fichiers temporaires..."
    gestProcessus.LancerProgrammeSimple "cmd /c del /q /s %temp%\*.*", True

    ' 3. Lancer le nettoyage de disque Windows
    If MsgBox("Lancer le nettoyage de disque Windows ?", vbYesNo + vbQuestion) = vbYes Then
        gestProcessus.LancerProgramme "cleanmgr.exe", "/sagerun:1", False
    End If

    Debug.Print "Nettoyage terminé"
End Sub

Sub SauvegardeAutomatique()
    ' Automatisation de sauvegarde de documents

    Debug.Print "=== SAUVEGARDE AUTOMATIQUE ==="

    Dim gestProcessus As New GestionnaireProcessus
    Dim dossierSource As String
    Dim dossierDestination As String
    Dim dateSauvegarde As String

    ' Configuration des dossiers
    dossierSource = Environ("USERPROFILE") & "\Documents"
    dateSauvegarde = Format(Now, "yyyy-mm-dd")
    dossierDestination = "D:\Sauvegardes\Documents_" & dateSauvegarde

    ' Créer le dossier de destination
    MkDir dossierDestination

    ' Copier les fichiers (méthode simple)
    Dim commande As String
    commande = "xcopy """ & dossierSource & """ """ & dossierDestination & """ /E /I /Y"

    Debug.Print "Début de la sauvegarde..."
    Debug.Print "Source : " & dossierSource
    Debug.Print "Destination : " & dossierDestination

    gestProcessus.LancerProgrammeSimple commande, True

    Debug.Print "Sauvegarde terminée"

    ' Ouvrir le dossier de destination
    gestProcessus.OuvrirDossier dossierDestination
End Sub

Sub MonitorageRessources()
    ' Surveillance des ressources système

    Debug.Print "=== MONITORAGE RESSOURCES ==="

    Dim gestProcessus As New GestionnaireProcessus
    Dim infosSys As New InformationsSystemeAvancees

    ' Générer un rapport détaillé
    infosSys.AfficherRapportSysteme

    ' Lancer le gestionnaire des tâches pour surveillance en temps réel
    If MsgBox("Ouvrir le gestionnaire des tâches ?", vbYesNo + vbQuestion) = vbYes Then
        gestProcessus.LancerProgramme "taskmgr.exe"
    End If

    ' Lancer l'observateur d'événements
    If MsgBox("Ouvrir l'observateur d'événements ?", vbYesNo + vbQuestion) = vbYes Then
        gestProcessus.LancerProgramme "eventvwr.exe"
    End If

    Debug.Print "Outils de monitorage lancés"
End Sub
```

### Interface utilisateur système

```vba
' ================================================================
' Module : InterfaceSysteme
' Description : Interface utilisateur pour les fonctions système
' ================================================================

Sub MenuPrincipalSysteme()
    ' Menu principal pour les fonctions système

    Dim choix As String
    Dim continuer As Boolean

    continuer = True

    Do While continuer
        choix = InputBox( _
            "=== GESTIONNAIRE SYSTÈME VBA ===" & vbCrLf & vbCrLf & _
            "1 - Informations système" & vbCrLf & _
            "2 - Gestion des fenêtres" & vbCrLf & _
            "3 - Lancement de programmes" & vbCrLf & _
            "4 - Nettoyage automatique" & vbCrLf & _
            "5 - Sauvegarde documents" & vbCrLf & _
            "6 - Monitorage ressources" & vbCrLf & _
            "7 - Démonstration complète" & vbCrLf & _
            "0 - Quitter" & vbCrLf & vbCrLf & _
            "Votre choix :", "Gestionnaire Système")

        Select Case choix
            Case "1"
                Dim infosSys As New InformationsSystemeAvancees
                infosSys.AfficherRapportSysteme
                MsgBox "Rapport généré dans la fenêtre Debug et sauvegardé dans " & Environ("TEMP")

            Case "2"
                MenuGestionFenetres

            Case "3"
                MenuLancementProgrammes

            Case "4"
                NettoyageSystemeAutomatique
                MsgBox "Nettoyage terminé"

            Case "5"
                SauvegardeAutomatique
                MsgBox "Sauvegarde terminée"

            Case "6"
                MonitorageRessources

            Case "7"
                DemonstrationInteractionSysteme
                MsgBox "Démonstration terminée - Consultez la fenêtre Debug"

            Case "0", ""
                continuer = False

            Case Else
                MsgBox "Choix invalide", vbExclamation
        End Select
    Loop

    MsgBox "Au revoir !", vbInformation
End Sub

Sub MenuGestionFenetres()
    ' Sous-menu pour la gestion des fenêtres

    Dim gestFenetres As New GestionnaireFenetres
    Dim titre As String
    Dim hwnd As LongPtr

    titre = InputBox("Entrez le titre de la fenêtre à rechercher :", "Recherche de fenêtre")
    If titre = "" Then Exit Sub

    hwnd = gestFenetres.TrouverFenetreParTitre(titre)

    If hwnd = 0 Then
        MsgBox "Fenêtre non trouvée : " & titre, vbExclamation
        Exit Sub
    End If

    Dim action As String
    action = InputBox( _
        "Fenêtre trouvée : " & gestFenetres.ObtenirTitreFenetre(hwnd) & vbCrLf & vbCrLf & _
        "Actions disponibles :" & vbCrLf & _
        "1 - Centrer" & vbCrLf & _
        "2 - Minimiser" & vbCrLf & _
        "3 - Maximiser" & vbCrLf & _
        "4 - Restaurer" & vbCrLf & _
        "5 - Premier plan" & vbCrLf & _
        "6 - Déplacer (100,100)" & vbCrLf & vbCrLf & _
        "Votre choix :", "Actions sur la fenêtre")

    Select Case action
        Case "1": gestFenetres.CentrerFenetre hwnd
        Case "2": gestFenetres.AfficherFenetre hwnd, SW_MINIMIZED
        Case "3": gestFenetres.AfficherFenetre hwnd, SW_MAXIMIZED
        Case "4": gestFenetres.AfficherFenetre hwnd, SW_RESTORE
        Case "5": gestFenetres.MettreAuPremierPlan hwnd
        Case "6": gestFenetres.DeplacerFenetre hwnd, 100, 100
        Case Else: Exit Sub
    End Select

    MsgBox "Action exécutée sur la fenêtre", vbInformation
End Sub

Sub MenuLancementProgrammes()
    ' Sous-menu pour le lancement de programmes

    Dim gestProcessus As New GestionnaireProcessus
    Dim choix As String

    choix = InputBox( _
        "=== LANCEMENT DE PROGRAMMES ===" & vbCrLf & vbCrLf & _
        "1 - Bloc-notes" & vbCrLf & _
        "2 - Calculatrice" & vbCrLf & _
        "3 - Explorateur Windows" & vbCrLf & _
        "4 - Panneau de configuration" & vbCrLf & _
        "5 - Gestionnaire des tâches" & vbCrLf & _
        "6 - Invite de commandes" & vbCrLf & _
        "7 - Programme personnalisé" & vbCrLf & vbCrLf & _
        "Votre choix :", "Lancement de programmes")

    Select Case choix
        Case "1"
            gestProcessus.LancerProgramme "notepad.exe"
        Case "2"
            gestProcessus.LancerProgramme "calc.exe"
        Case "3"
            gestProcessus.OuvrirDossier Environ("USERPROFILE")
        Case "4"
            gestProcessus.LancerProgramme "control.exe"
        Case "5"
            gestProcessus.LancerProgramme "taskmgr.exe"
        Case "6"
            gestProcessus.LancerProgramme "cmd.exe"
        Case "7"
            Dim programme As String
            programme = InputBox("Entrez le chemin du programme à lancer :", "Programme personnalisé")
            If programme <> "" Then
                gestProcessus.LancerProgramme programme
            End If
        Case Else
            Exit Sub
    End Select

    MsgBox "Programme lancé", vbInformation
End Sub
```

## Bonnes pratiques et sécurité

### 1. Gestion des erreurs robuste

```vba
Public Function LancementSecurise(programme As String) As Boolean
    ' Lance un programme avec gestion d'erreurs complète

    On Error GoTo GestionErreur

    ' Vérifier que le fichier existe
    If Dir(programme) = "" Then
        Debug.Print "Erreur : Fichier introuvable - " & programme
        LancementSecurise = False
        Exit Function
    End If

    ' Vérifier l'extension (sécurité de base)
    Dim extension As String
    extension = LCase(Mid(programme, InStrRev(programme, ".")))

    If extension <> ".exe" And extension <> ".com" And extension <> ".bat" Then
        Debug.Print "Attention : Type de fichier inhabituel - " & extension
        If MsgBox("Lancer quand même ?", vbYesNo + vbQuestion) = vbNo Then
            LancementSecurise = False
            Exit Function
        End If
    End If

    ' Lancement avec timeout
    Shell programme, vbNormalFocus
    LancementSecurise = True

    Debug.Print "Programme lancé avec succès : " & programme
    Exit Function

GestionErreur:
    Debug.Print "Erreur lors du lancement : " & Err.Description
    LancementSecurise = False
End Function
```

### 2. Validation des paramètres

```vba
Private Function ValiderChemin(chemin As String) As Boolean
    ' Valide un chemin de fichier ou dossier

    ' Vérifications de base
    If Len(chemin) = 0 Then
        ValiderChemin = False
        Exit Function
    End If

    ' Caractères interdits dans les noms de fichiers/dossiers
    ' Note : ":" est autorisé en position 2 (lettre de lecteur, ex: "C:\")
    Dim caracteresInterdits As String
    caracteresInterdits = "<>""|?*"

    Dim i As Integer
    For i = 1 To Len(caracteresInterdits)
        If InStr(chemin, Mid(caracteresInterdits, i, 1)) > 0 Then
            Debug.Print "Caractère interdit trouvé : " & Mid(caracteresInterdits, i, 1)
            ValiderChemin = False
            Exit Function
        End If
    Next i

    ' Longueur maximale
    If Len(chemin) > 260 Then
        Debug.Print "Chemin trop long (max 260 caractères)"
        ValiderChemin = False
        Exit Function
    End If

    ValiderChemin = True
End Function
```

### 3. Nettoyage des ressources

```vba
Public Sub NettoyageRessources()
    ' Nettoie toutes les ressources système utilisées
    ' À adapter selon les objets et handles utilisés dans votre application

    ' Exemple de nettoyage type :
    ' CloseHandle hMonHandle        ' Fermer les handles ouverts
    ' Set monObjetCOM = Nothing     ' Libérer les objets COM

    ' Permettre le traitement des événements en attente
    DoEvents

    Debug.Print "Nettoyage des ressources terminé"
End Sub
```

### 4. Logging et audit

```vba
Private Sub EcrireLog(message As String, Optional niveau As String = "INFO")
    ' Système de logging simple

    Dim fichierLog As String
    Dim numeroFichier As Integer
    Dim timestamp As String

    fichierLog = Environ("TEMP") & "\SystemeVBA.log"
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    numeroFichier = FreeFile
    Open fichierLog For Append As numeroFichier

    Print #numeroFichier, timestamp & " [" & niveau & "] " & message

    Close numeroFichier

    ' Aussi dans Debug pour développement
    Debug.Print timestamp & " [" & niveau & "] " & message
End Sub
```

## Limitations et considérations

### Compatibilité Windows
- **Versions supportées** : Windows 7 et plus récent recommandé
- **Architecture** : Code 32/64 bits avec compilation conditionnelle
- **Droits utilisateur** : Certaines opérations nécessitent des droits administrateur

### Performance
- **Appels API** : Plus rapides que les équivalents VBA/COM
- **Mémoire** : Attention aux fuites de mémoire avec les handles
- **Threading** : VBA est mono-thread, attention aux blocages

### Sécurité
- **Validation** : Toujours valider les paramètres d'entrée
- **Droits** : Ne pas demander plus de droits que nécessaire
- **Audit** : Logger les opérations sensibles

### Maintenance
- **Documentation** : Documenter tous les appels d'API utilisés
- **Tests** : Tester sur différentes configurations
- **Évolution** : Prévoir les changements de versions Windows

L'interaction avec le système via les API Windows ouvre des possibilités immenses pour vos applications VBA. Utilisez ces fonctionnalités avec prudence et toujours dans l'intérêt de l'utilisateur et de la stabilité du système.

⏭️
