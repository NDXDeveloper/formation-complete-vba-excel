🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.5. Compatibilité entre versions

## Qu'est-ce que la compatibilité entre versions ?

La compatibilité entre versions, c'est s'assurer que votre code VBA **fonctionne correctement** sur différentes versions d'Excel et d'Office. C'est comme s'assurer qu'une recette de cuisine fonctionne dans différents types de fours : votre "recette" VBA doit donner le même résultat, que l'utilisateur ait Excel 2016, 2019, 2021, ou Office 365.

Imaginez que vous développez une solution VBA géniale sur Excel 2021, mais que la moitié de votre entreprise utilise encore Excel 2016. Si votre code ne fonctionne que sur la version récente, vous excluez automatiquement une partie de vos utilisateurs. La compatibilité permet à votre solution d'atteindre le plus large public possible.

## Pourquoi la compatibilité est-elle importante ?

**Adoption maximale** : Plus votre code est compatible, plus il peut être utilisé par des utilisateurs différents dans des environnements variés.

**Coût de maintenance réduit** : Maintenir une seule version compatible coûte moins cher que développer des versions spécifiques pour chaque plateforme.

**Pérennité** : Un code compatible résiste mieux au temps et aux évolutions technologiques.

**Satisfaction utilisateur** : Les utilisateurs apprécient les solutions qui fonctionnent sur leur environnement existant, sans nécessiter de mise à jour.

**Contraintes organisationnelles** : Beaucoup d'entreprises ne peuvent pas mettre à jour immédiatement vers les dernières versions pour des raisons de coût, de politique IT, ou de compatibilité avec d'autres systèmes.

## Vue d'ensemble des versions Excel

### Versions principales et leurs caractéristiques

**Excel 2010** :
- Première version avec l'interface Ruban complète
- VBA 7.0, architecture 64-bit optionnelle
- Fin de support : 2020

**Excel 2013** :
- Nouvelles fonctions de feuille de calcul
- Améliorations de l'interface utilisateur
- VBA reste largement identique à 2010

**Excel 2016** :
- Intégration cloud renforcée
- Nouvelles fonctions comme TEXTJOIN, CONCAT
- Améliorations de performance

**Excel 2019** :
- Améliorations des graphiques
- Types de données liés (Actions, Géographie)
- Nouvelles fonctions comme IFS, SWITCH, MAXIFS, MINIFS

**Excel 2021** :
- Fonctions XLOOKUP, LET
- Fonctions de tableaux dynamiques (UNIQUE, SORT, FILTER, SEQUENCE)
- Types de données étendus

**Microsoft 365 (Office 365)** :
- Mises à jour continues
- Fonctionnalités les plus récentes
- Évolution constante

### Cycles de développement

**Versions boîte (perpétuelles)** : Excel 2019, 2021
- Mises à jour de sécurité uniquement
- Fonctionnalités figées à la sortie

**Microsoft 365 (abonnement)** :
- Nouvelles fonctionnalités ajoutées régulièrement
- Canal de mise à jour configurable (Mensuel, Semi-annuel)

## Principales différences affectant VBA

### Architecture 32-bit vs 64-bit

**Le problème** : Depuis Excel 2010, Office existe en versions 32-bit et 64-bit.

**Impact sur VBA** :
- Les déclarations d'API Windows doivent être adaptées
- Les types de données pour les pointeurs changent
- Certaines bibliothèques externes peuvent être incompatibles

**Code problématique** :
```vba
' Fonctionne uniquement en 32-bit
Private Declare Function GetWindowsDirectory Lib "kernel32" _
    Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
```

**Code compatible** :
```vba
' Fonctionne en 32-bit et 64-bit
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowsDirectory Lib "kernel32" _
        Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
    Private Declare Function GetWindowsDirectory Lib "kernel32" _
        Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If
```

### Nouveaux objets et propriétés

**Fonctionnalités récentes non disponibles dans les anciennes versions** :

```vba
Sub ExempleCompatibilite()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Vérifier si une propriété récente existe avant de l'utiliser
    If ProprieteExiste(ws, "EnableAnimations") Then
        ' Code utilisant la propriété récente
    Else
        ' Code alternatif pour les anciennes versions
    End If
End Sub
```

### Changements dans les constantes

**Nouvelles constantes** : Certaines constantes VBA ont été ajoutées dans les versions récentes.

```vba
Sub UtiliserConstantes()
    Dim formatFichier As Long

    ' Vérifier la version pour choisir le bon format de fichier
    If Val(Application.Version) >= 12 Then  ' Excel 2007+
        formatFichier = xlOpenXMLWorkbook  ' Format .xlsx
    Else
        formatFichier = xlNormal  ' Format .xls (Excel 97-2003)
    End If

    ActiveWorkbook.SaveAs "MonFichier", formatFichier
End Sub
```

### Fonctions de feuille de calcul évoluées

**Nouvelles fonctions** : XLOOKUP, UNIQUE, SORT, etc. ne sont pas disponibles dans toutes les versions.

```vba
Function RechercheCompatible(valeurCherchee As Variant, plageRecherche As Range, plageRetour As Range) As Variant
    ' Utiliser XLOOKUP si disponible, sinon VLOOKUP
    On Error GoTo UtiliserVLookup

    ' Tentative avec XLOOKUP (Excel 365/2021)
    RechercheCompatible = Application.WorksheetFunction.XLookup(valeurCherchee, plageRecherche, plageRetour)
    Exit Function

UtiliserVLookup:
    ' Méthode alternative avec VLOOKUP pour les anciennes versions
    On Error GoTo GestionErreur
    RechercheCompatible = Application.WorksheetFunction.VLookup(valeurCherchee, plageRecherche.Resize(, 2), 2, False)
    Exit Function

GestionErreur:
    RechercheCompatible = "Non trouvé"
End Function
```

## Stratégies de développement compatible

### Développement pour le plus petit dénominateur commun

**Principe** : Développez en utilisant les fonctionnalités disponibles dans la version la plus ancienne que vous devez supporter.

**Avantages** :
- Compatibilité garantie
- Code plus simple
- Moins de tests nécessaires

**Inconvénients** :
- Impossibilité d'utiliser les nouvelles fonctionnalités
- Code parfois moins efficace
- Fonctionnalités limitées

### Développement adaptatif

**Principe** : Détectez la version et adaptez le comportement en conséquence.

```vba
Function ObtenirVersionExcel() As Double
    ObtenirVersionExcel = Val(Application.Version)
End Function

Sub ComportementAdaptatif()
    Dim version As Double
    version = ObtenirVersionExcel()

    Select Case version
        Case Is >= 16  ' Excel 2016 et plus récent
            ' Utiliser les fonctionnalités modernes
            UtiliserFonctionsModernes
        Case Is >= 14  ' Excel 2010-2013
            ' Utiliser les fonctionnalités intermédiaires
            UtiliserFonctionsIntermediaires
        Case Else      ' Excel 2007 et antérieur
            ' Utiliser seulement les fonctionnalités de base
            UtiliserFonctionsDeBase
    End Select
End Sub
```

### Développement modulaire

**Principe** : Séparez les fonctionnalités spécifiques aux versions dans des modules distincts.

```vba
' Module: CompatibiliteExcel
Public Const EXCEL_2010 As Double = 14  
Public Const EXCEL_2013 As Double = 15  
Public Const EXCEL_2016 As Double = 16  

Function EstVersionMinimum(versionMinimum As Double) As Boolean
    EstVersionMinimum = (Val(Application.Version) >= versionMinimum)
End Function

' Module: FonctionsModernes (utilisé seulement si Excel 2016+)
Sub UtiliserTableauxDynamiques()
    If Not EstVersionMinimum(EXCEL_2016) Then
        MsgBox "Cette fonctionnalité nécessite Excel 2016 ou plus récent"
        Exit Sub
    End If

    ' Code utilisant les nouvelles fonctionnalités
End Sub
```

## Techniques de détection de version

### Détection de la version d'Excel

```vba
Function InformationsVersion() As String
    Dim info As String

    info = "Version Excel : " & Application.Version & vbCrLf
    info = info & "Architecture : "

    #If VBA7 Then
        #If Win64 Then
            info = info & "64-bit" & vbCrLf
        #Else
            info = info & "32-bit" & vbCrLf
        #End If
    #Else
        info = info & "Ancienne version (pré-2010)" & vbCrLf
    #End If

    info = info & "Build : " & Application.Build

    InformationsVersion = info
End Function
```

### Détection de fonctionnalités disponibles

```vba
Function FonctionDisponible(nomFonction As String) As Boolean
    On Error GoTo FonctionIndisponible

    ' Tenter d'utiliser la fonction avec des paramètres test
    Select Case UCase(nomFonction)
        Case "XLOOKUP"
            Application.WorksheetFunction.XLookup(1, Range("A1"), Range("B1"))
        Case "UNIQUE"
            Application.WorksheetFunction.Unique(Range("A1:A2"))
        Case Else
            FonctionDisponible = False
            Exit Function
    End Select

    FonctionDisponible = True
    Exit Function

FonctionIndisponible:
    FonctionDisponible = False
End Function
```

### Détection d'objets et propriétés

```vba
Function ProprieteExiste(objet As Object, nomPropriete As String) As Boolean
    On Error GoTo ProprieteInexistante

    Dim valeur As Variant
    valeur = CallByName(objet, nomPropriete, VbGet)

    ProprieteExiste = True
    Exit Function

ProprieteInexistante:
    ProprieteExiste = False
End Function
```

## Gestion des formats de fichiers

### Formats supportés par version

**Excel 2007+** : .xlsx, .xlsm, .xlam  
**Excel 2003 et antérieur** : .xls, .xla  

**Code adaptatif pour la sauvegarde** :
```vba
Sub SauvegardeCompatible(nomFichier As String)
    Dim formatFichier As Long
    Dim extension As String

    If Val(Application.Version) >= 12 Then  ' Excel 2007+
        formatFichier = xlOpenXMLWorkbookMacroEnabled
        extension = ".xlsm"
    Else  ' Excel 2003 et antérieur
        formatFichier = xlExcel8
        extension = ".xls"
    End If

    ActiveWorkbook.SaveAs nomFichier & extension, formatFichier
End Sub
```

### Conversion de formats

```vba
Sub ConvertirPourAnciennesVersions()
    Dim cheminOriginal As String
    Dim cheminCompatible As String

    cheminOriginal = ActiveWorkbook.FullName
    cheminCompatible = Replace(cheminOriginal, ".xlsm", "_compatible.xls")

    ' Sauvegarder en format Excel 97-2003
    ActiveWorkbook.SaveAs cheminCompatible, xlExcel8

    MsgBox "Version compatible créée : " & cheminCompatible
End Sub
```

## Tests de compatibilité

### Environnement de test

**Machines virtuelles** : Utilisez des VM avec différentes versions d'Excel pour tester.

**Utilisateurs pilotes** : Identifiez des utilisateurs avec différentes versions pour les tests.

**Automatisation** : Créez des scripts de test automatique qui vérifient les fonctionnalités sur chaque version.

### Liste de vérification pour les tests

```vba
Sub TestsCompatibilite()
    Dim resultats As String

    resultats = "=== TESTS DE COMPATIBILITÉ ===" & vbCrLf
    resultats = resultats & "Version : " & Application.Version & vbCrLf
    resultats = resultats & "Build : " & Application.Build & vbCrLf & vbCrLf

    ' Test 1 : Fonctions de base
    On Error Resume Next
    Range("A1").Value = "Test"
    If Err.Number = 0 Then
        resultats = resultats & "✓ Accès aux cellules : OK" & vbCrLf
    Else
        resultats = resultats & "✗ Accès aux cellules : ERREUR" & vbCrLf
    End If
    Err.Clear

    ' Test 2 : API Windows (si utilisées)
    Dim versionVBA As String
    #If VBA7 Then
        versionVBA = "VBA7 détecté"
    #Else
        versionVBA = "VBA6 ou antérieur"
    #End If
    resultats = resultats & "Info VBA : " & versionVBA & vbCrLf

    ' Test 3 : Fonctions spécifiques
    If FonctionDisponible("XLOOKUP") Then
        resultats = resultats & "✓ XLOOKUP disponible" & vbCrLf
    Else
        resultats = resultats & "- XLOOKUP non disponible" & vbCrLf
    End If

    ' Afficher les résultats
    MsgBox resultats, vbInformation, "Résultats des tests"
End Sub
```

## Stratégies de déploiement

### Déploiement conditionnel

```vba
Sub VerifierPrerequisAvantInstallation()
    Dim versionMinimum As Double
    versionMinimum = 14  ' Excel 2010

    If Val(Application.Version) < versionMinimum Then
        MsgBox "Cette solution nécessite Excel 2010 ou plus récent." & vbCrLf & _
               "Version détectée : " & Application.Version, vbCritical
        Exit Sub
    End If

    ' Procéder à l'installation
    InstallerSolution
End Sub
```

### Messages d'erreur informatifs

```vba
Sub GererErreurCompatibilite(numeroErreur As Long, descriptionErreur As String)
    Dim message As String

    message = "Une erreur de compatibilité s'est produite :" & vbCrLf & vbCrLf
    message = message & "Erreur : " & descriptionErreur & vbCrLf
    message = message & "Code : " & numeroErreur & vbCrLf & vbCrLf

    message = message & "Informations système :" & vbCrLf
    message = message & "Excel : " & Application.Version & vbCrLf
    message = message & "OS : " & Application.OperatingSystem & vbCrLf & vbCrLf

    message = message & "Solutions possibles :" & vbCrLf
    message = message & "• Mettre à jour vers une version plus récente d'Excel" & vbCrLf
    message = message & "• Contacter le support technique" & vbCrLf
    message = message & "• Utiliser le mode compatibilité"

    MsgBox message, vbExclamation, "Problème de compatibilité"
End Sub
```

## Documentation de compatibilité

### Matrice de compatibilité

Créez un tableau documentant les fonctionnalités par version :

```
MATRICE DE COMPATIBILITÉ - MON OUTIL VBA

Fonctionnalité          | Excel 2010 | Excel 2013 | Excel 2016 | Excel 2019 | Excel 365
-----------------------|------------|------------|------------|------------|----------
Import CSV             |     ✓      |     ✓      |     ✓      |     ✓      |     ✓  
Export PDF             |     ✓      |     ✓      |     ✓      |     ✓      |     ✓  
Recherche avancée      |     ✓      |     ✓      |     ✓      |     ✓      |     ✓  
Tableaux dynamiques    |     -      |     -      |     ✓      |     ✓      |     ✓  
Fonctions modernes     |     -      |     -      |     -      |     ✓      |     ✓  
```

### Notes de version

```
NOTES DE COMPATIBILITÉ v2.1

VERSIONS SUPPORTÉES :
- Excel 2010 (14.0) : Fonctionnalités de base uniquement
- Excel 2013 (15.0) : Support complet sauf tableaux dynamiques
- Excel 2016 (16.0) : Support complet
- Excel 2019+ : Toutes les fonctionnalités

PROBLÈMES CONNUS :
- Excel 2010 : L'export PDF peut être lent sur de gros fichiers
- Excel 2013 : Certains graphiques peuvent s'afficher différemment

SOLUTIONS DE CONTOURNEMENT :
- Pour Excel 2010 : Utiliser l'export XPS comme alternative au PDF
- Pour toutes versions : Redémarrer Excel en cas de problème d'affichage
```

## Maintenance et évolution

### Stratégie de support

**Support étendu** : Maintenir la compatibilité avec plusieurs versions anciennes
- Avantage : Large adoption
- Inconvénient : Code complexe, innovations limitées

**Support progressif** : Abandonner progressivement les anciennes versions
- Avantage : Code plus moderne, nouvelles fonctionnalités
- Inconvénient : Réduction du public cible

### Plan de migration

```vba
Sub PlanifierMigration()
    Dim message As String

    If Val(Application.Version) < 16 Then
        message = "AVERTISSEMENT :" & vbCrLf & vbCrLf
        message = message & "Vous utilisez une version ancienne d'Excel (" & Application.Version & ")." & vbCrLf
        message = message & "Le support de cette version sera arrêté dans 6 mois." & vbCrLf & vbCrLf
        message = message & "Nous recommandons une mise à jour vers Excel 2016 ou plus récent " & vbCrLf
        message = message & "pour bénéficier de toutes les fonctionnalités et du support continu."

        MsgBox message, vbExclamation, "Plan de migration"
    End If
End Sub
```

## Bonnes pratiques

### Développement

**Testez régulièrement** : Ne pas attendre la fin du développement pour tester la compatibilité

**Documentez les dépendances** : Notez clairement quelles fonctionnalités nécessitent quelles versions

**Utilisez la compilation conditionnelle** : Séparez le code spécifique aux versions avec #If

**Prévoyez des alternatives** : Toujours avoir un plan B pour les fonctionnalités modernes

### Déploiement

**Communiquez clairement** : Informez les utilisateurs des prérequis avant l'installation

**Fournissez des versions multiples** : Si possible, créez des versions adaptées aux différents environnements

**Testez en conditions réelles** : Validez sur les vraies configurations utilisateur

**Monitorer l'usage** : Suivez quelles versions sont réellement utilisées pour adapter votre stratégie

La compatibilité entre versions est un défi technique qui nécessite une approche méthodique et une planification soigneuse. En anticipant ces questions dès le début du développement, vous créez des solutions VBA robustes qui peuvent évoluer avec l'écosystème Office et servir efficacement vos utilisateurs, quelle que soit leur configuration.

⏭️
