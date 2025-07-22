🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.3. On Error GoTo

## Introduction à On Error GoTo

L'instruction `On Error GoTo` est une méthode plus sophistiquée de gestion d'erreurs que `On Error Resume Next`. Au lieu d'ignorer les erreurs, elle **redirige l'exécution vers une section spécialisée** de votre code quand une erreur survient. C'est comme avoir un plan d'urgence bien organisé qui s'active automatiquement en cas de problème.

**Analogie simple :**
Imaginez que vous conduisez en suivant un GPS. Avec `On Error Resume Next`, si la route est bloquée, vous passez simplement à l'instruction suivante (potentiellement dans un mur). Avec `On Error GoTo`, le GPS vous dit automatiquement "route bloquée, redirection vers l'itinéraire de secours" et vous guide vers une solution alternative.

---

## Syntaxe de base

### Structure générale

```vba
Sub ExempleStructure()
    On Error GoTo GestionErreur    ' Activer la redirection d'erreur

    ' Votre code principal ici
    Range("A1").Value = 10 / 0     ' Cette ligne causera une erreur
    MsgBox "Cette ligne ne s'exécutera jamais"

    Exit Sub                       ' IMPORTANT : Sortir avant la section d'erreur

GestionErreur:                     ' Étiquette de destination
    MsgBox "Une erreur est survenue : " & Err.Description
    ' Code de gestion d'erreur ici
End Sub
```

### Éléments essentiels

1. **On Error GoTo [Étiquette]** : Active la redirection
2. **Étiquette:** : Point de destination (nom suivi de deux-points)
3. **Exit Sub/Function** : Sortie avant la section d'erreur
4. **Code de gestion** : Actions à effectuer en cas d'erreur

---

## Comment fonctionne On Error GoTo

### Flux d'exécution normal vs avec erreur

#### Exécution normale (sans erreur)

```vba
Sub FluxNormal()
    On Error GoTo GestionErreur

    Range("A1").Value = 10        ' ✓ S'exécute
    Range("A2").Value = 20        ' ✓ S'exécute
    Range("A3").Value = Range("A1").Value + Range("A2").Value  ' ✓ S'exécute
    MsgBox "Calcul terminé"       ' ✓ S'exécute

    Exit Sub                      ' ✓ Sort ici - section d'erreur ignorée

GestionErreur:                    ' ✗ Cette section n'est pas exécutée
    MsgBox "Erreur détectée"
End Sub
```

#### Exécution avec erreur

```vba
Sub FluxAvecErreur()
    On Error GoTo GestionErreur

    Range("A1").Value = 10        ' ✓ S'exécute
    Range("A2").Value = 0         ' ✓ S'exécute
    Range("A3").Value = Range("A1").Value / Range("A2").Value  ' ✗ ERREUR ici
    MsgBox "Calcul terminé"       ' ✗ Jamais exécuté

    Exit Sub                      ' ✗ Jamais atteint

GestionErreur:                    ' ✓ Saut direct ici
    MsgBox "Division par zéro détectée !"
    Range("A3").Value = "Erreur"
End Sub
```

---

## Types d'étiquettes et conventions de nommage

### Noms d'étiquettes courants

```vba
Sub ConventionsNommage()
    On Error GoTo GestionErreur   ' Le plus courant
    ' Ou
    On Error GoTo ErreurHandler   ' Version anglaise
    ' Ou
    On Error GoTo TraiterErreur   ' Plus descriptif

    ' Code principal...

    Exit Sub

GestionErreur:
    ' Code de gestion d'erreur
End Sub
```

### Plusieurs gestionnaires d'erreur

```vba
Sub PlusieurGestionnaires()
    On Error GoTo ErreurFichier

    ' Opérations sur fichiers
    Workbooks.Open "C:\MonFichier.xlsx"

    On Error GoTo ErreurCalcul

    ' Opérations de calcul
    Range("A1").Value = Range("B1").Value / Range("C1").Value

    Exit Sub

ErreurFichier:
    MsgBox "Problème avec le fichier : " & Err.Description
    Exit Sub

ErreurCalcul:
    MsgBox "Erreur de calcul : " & Err.Description
    Range("A1").Value = "N/A"
End Sub
```

---

## Patterns courants de gestion d'erreur

### 1. Pattern simple avec message

```vba
Sub PatternSimple()
    On Error GoTo GestionErreur

    ' Code principal
    Dim resultat As Double
    resultat = Range("A1").Value / Range("B1").Value
    Range("C1").Value = resultat

    MsgBox "Opération réussie"
    Exit Sub

GestionErreur:
    MsgBox "Erreur rencontrée : " & vbCrLf & _
           "Numéro : " & Err.Number & vbCrLf & _
           "Description : " & Err.Description
End Sub
```

### 2. Pattern avec nettoyage

```vba
Sub PatternNettoyage()
    On Error GoTo GestionErreur

    ' Déclarations
    Dim wb As Workbook
    Dim ws As Worksheet

    ' Opérations
    Set wb = Workbooks.Open("C:\Données.xlsx")
    Set ws = wb.Worksheets("Import")

    ' Traitement des données
    ws.Range("A1").Value = "Traitement en cours..."

    ' Nettoyage normal
    wb.Close SaveChanges:=True
    MsgBox "Traitement terminé avec succès"
    Exit Sub

GestionErreur:
    ' Nettoyage en cas d'erreur
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    MsgBox "Erreur pendant le traitement : " & Err.Description
End Sub
```

### 3. Pattern avec tentatives multiples

```vba
Sub PatternTentativesMultiples()
    On Error GoTo GestionErreur

    Dim tentatives As Integer
    tentatives = 0

Retry:
    tentatives = tentatives + 1

    ' Tentative d'opération
    Workbooks.Open "C:\FichierReseau.xlsx"

    MsgBox "Fichier ouvert avec succès"
    Exit Sub

GestionErreur:
    If tentatives < 3 Then
        MsgBox "Tentative " & tentatives & " échouée. Nouvelle tentative..."
        Err.Clear
        Resume Retry
    Else
        MsgBox "Impossible d'ouvrir le fichier après 3 tentatives : " & Err.Description
    End If
End Sub
```

---

## Instructions de reprise

### Resume - Reprendre à la ligne d'erreur

```vba
Sub ExempleResume()
    On Error GoTo GestionErreur

    Dim i As Integer
    For i = 1 To 5
        Range("A" & i).Value = 100 / Range("B" & i).Value
    Next i

    MsgBox "Tous les calculs terminés"
    Exit Sub

GestionErreur:
    ' Mettre une valeur par défaut et continuer
    Range("A" & i).Value = 0
    Resume Next  ' Continue avec la prochaine itération de la boucle
End Sub
```

### Resume Next - Reprendre à la ligne suivante

```vba
Sub ExempleResumeNext()
    On Error GoTo GestionErreur

    Range("A1").Value = "Début"
    Range("FeuillePeutPasExister").Value = "Test"  ' Erreur possible
    Range("A2").Value = "Fin"  ' Cette ligne s'exécutera grâce à Resume Next

    Exit Sub

GestionErreur:
    MsgBox "Erreur ignorée : " & Err.Description
    Resume Next  ' Continue avec Range("A2").Value = "Fin"
End Sub
```

### Resume [Étiquette] - Reprendre à une position spécifique

```vba
Sub ExempleResumeEtiquette()
    On Error GoTo GestionErreur

    ' Tentative d'ouverture du fichier principal
    Workbooks.Open "C:\Principal.xlsx"
    GoTo TraitementDonnees

    Exit Sub

GestionErreur:
    MsgBox "Fichier principal inaccessible, utilisation du fichier de sauvegarde"
    Workbooks.Open "C:\Sauvegarde.xlsx"
    Resume TraitementDonnees

TraitementDonnees:
    ' Code commun pour traiter les données
    MsgBox "Traitement des données en cours..."
End Sub
```

---

## Gestion d'erreurs spécialisée par type

### Gestion selon le numéro d'erreur

```vba
Sub GestionParType()
    On Error GoTo GestionErreur

    ' Code susceptible de générer différents types d'erreurs
    Dim resultat As Double
    resultat = Range("A1").Value / Range("B1").Value
    Worksheets("Données").Range("C1").Value = resultat

    Exit Sub

GestionErreur:
    Select Case Err.Number
        Case 9  ' Subscript out of range
            MsgBox "La feuille 'Données' n'existe pas. Création en cours..."
            Worksheets.Add.Name = "Données"
            Resume  ' Reprendre à la ligne qui a causé l'erreur

        Case 11  ' Division by zero
            MsgBox "Division par zéro détectée"
            Range("C1").Value = "Infini"

        Case 13  ' Type mismatch
            MsgBox "Type de données incorrect dans les cellules"
            Range("C1").Value = "Erreur type"

        Case Else
            MsgBox "Erreur inattendue : " & Err.Number & " - " & Err.Description
    End Select
End Sub
```

### Gestion d'erreurs en cascade

```vba
Sub GestionCascade()
    On Error GoTo ErreurPrincipale

    ' Étape 1 : Ouvrir fichier
    Workbooks.Open "C:\Données.xlsx"

    On Error GoTo ErreurTraitement

    ' Étape 2 : Traiter données
    Range("A1:A10").Formula = "=B1:B10*2"

    On Error GoTo ErreurSauvegarde

    ' Étape 3 : Sauvegarder
    ActiveWorkbook.Save

    MsgBox "Toutes les étapes terminées avec succès"
    Exit Sub

ErreurPrincipale:
    MsgBox "Impossible d'ouvrir le fichier : " & Err.Description
    Exit Sub

ErreurTraitement:
    MsgBox "Erreur pendant le traitement : " & Err.Description
    ' Continuer vers la sauvegarde malgré l'erreur de traitement
    Resume Next

ErreurSauvegarde:
    MsgBox "Impossible de sauvegarder : " & Err.Description
    ' Proposer une sauvegarde alternative
    ActiveWorkbook.SaveAs "C:\Sauvegarde_Urgence.xlsx"
End Sub
```

---

## Techniques avancées

### 1. Journalisation des erreurs

```vba
Sub AvecJournalisation()
    On Error GoTo GestionErreur

    ' Code principal
    Dim i As Integer
    For i = 1 To 100
        Cells(i, 1).Value = i * Cells(i, 2).Value
    Next i

    Exit Sub

GestionErreur:
    ' Enregistrer l'erreur dans un journal
    Dim ligneJournal As String
    ligneJournal = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & _
                   "Erreur " & Err.Number & ": " & Err.Description & _
                   " à la ligne " & Erl

    ' Ajouter au journal (dans une feuille dédiée)
    Worksheets("Journal").Range("A1").End(xlDown).Offset(1, 0).Value = ligneJournal

    ' Proposer à l'utilisateur de continuer ou d'arrêter
    If MsgBox("Erreur rencontrée. Continuer ?", vbYesNo) = vbYes Then
        Resume Next
    End If
End Sub
```

### 2. Gestionnaire d'erreur centralisé

```vba
' Module séparé pour la gestion d'erreurs
Sub GestionnaireErreurCentralise(procedureName As String)
    Dim message As String
    message = "Erreur dans " & procedureName & vbCrLf & _
              "Numéro : " & Err.Number & vbCrLf & _
              "Description : " & Err.Description & vbCrLf & _
              "Heure : " & Format(Now, "yyyy-mm-dd hh:mm:ss")

    ' Log dans fichier ou base de données
    Debug.Print message

    ' Afficher à l'utilisateur
    MsgBox message, vbCritical, "Erreur Application"
End Sub

Sub UtiliserGestionnaireCentralise()
    On Error GoTo GestionErreur

    ' Code principal
    Range("A1").Value = 10 / 0

    Exit Sub

GestionErreur:
    Call GestionnaireErreurCentralise("UtiliserGestionnaireCentralise")
End Sub
```

---

## Bonnes pratiques avec On Error GoTo

### 1. Structure recommandée

```vba
Sub StructureRecommandee()
    ' Déclarations en début
    Dim variable1 As String
    Dim variable2 As Integer

    ' Activation du gestionnaire d'erreur
    On Error GoTo GestionErreur

    ' Code principal bien structuré
    ' ... votre logique ici ...

    ' Sortie normale (OBLIGATOIRE)
    On Error GoTo 0  ' Désactiver la gestion d'erreur
    Exit Sub

    ' Section de gestion d'erreur
GestionErreur:
    ' Traitement de l'erreur
    MsgBox "Erreur : " & Err.Description

    ' Nettoyage si nécessaire
    ' ... code de nettoyage ...

    ' Optionnel : désactiver la gestion d'erreur
    On Error GoTo 0
End Sub
```

### 2. Éviter les pièges courants

#### Piège 1 : Oublier Exit Sub

```vba
' INCORRECT - Risque d'exécuter la section d'erreur normalement
Sub PiegeExitSub()
    On Error GoTo GestionErreur

    Range("A1").Value = "OK"
    ' OUBLI : Exit Sub

GestionErreur:
    MsgBox "Cette section s'exécute toujours !"  ' PROBLÈME
End Sub

' CORRECT
Sub CorrectExitSub()
    On Error GoTo GestionErreur

    Range("A1").Value = "OK"
    Exit Sub  ' IMPORTANT

GestionErreur:
    MsgBox "Cette section ne s'exécute qu'en cas d'erreur"
End Sub
```

#### Piège 2 : Gestionnaire d'erreur dans une boucle infinie

```vba
' ATTENTION : Risque de boucle infinie
Sub AttentionBoucleInfinie()
    On Error GoTo GestionErreur

    Range("A1").Value = 10 / 0  ' Génère toujours une erreur

    Exit Sub

GestionErreur:
    MsgBox "Erreur détectée"
    Resume  ' Retourne à la ligne d'erreur = boucle infinie !
End Sub
```

### 3. Tests et débogage

```vba
Sub CodeAvecDebug()
    On Error GoTo GestionErreur

    Debug.Print "Début de la procédure"

    ' Code principal avec points de contrôle
    Debug.Print "Avant calcul"
    Dim resultat As Double
    resultat = Range("A1").Value / Range("B1").Value
    Debug.Print "Résultat calculé : " & resultat

    Range("C1").Value = resultat
    Debug.Print "Fin normale"
    Exit Sub

GestionErreur:
    Debug.Print "ERREUR - Numéro : " & Err.Number & ", Description : " & Err.Description
    MsgBox "Erreur : " & Err.Description
End Sub
```

---

## Comparaison avec On Error Resume Next

### Quand utiliser chaque méthode

| Critère | On Error GoTo | On Error Resume Next |
|---------|---------------|----------------------|
| **Contrôle** | Maximum | Minimal |
| **Complexité** | Plus complexe | Plus simple |
| **Maintenance** | Plus facile | Plus difficile |
| **Débogage** | Meilleur | Plus difficile |
| **Performance** | Légèrement moins bon | Meilleur |
| **Flexibilité** | Maximum | Limitée |

### Exemple comparatif

```vba
' Avec On Error Resume Next
Sub MethodeResumeNext()
    On Error Resume Next

    Range("A1").Value = 10 / 0
    If Err.Number <> 0 Then
        Range("A1").Value = "Erreur"
        Err.Clear
    End If

    On Error GoTo 0
End Sub

' Avec On Error GoTo
Sub MethodeGoTo()
    On Error GoTo GestionErreur

    Range("A1").Value = 10 / 0

    Exit Sub

GestionErreur:
    Range("A1").Value = "Erreur"
    MsgBox "Division par zéro détectée et corrigée"
End Sub
```

---

## Récapitulatif

### Points clés à retenir

1. **On Error GoTo** redirige vers une section spécialisée en cas d'erreur
2. **Exit Sub/Function** est obligatoire avant la section d'erreur
3. **Resume, Resume Next, Resume [Étiquette]** permettent de reprendre l'exécution
4. **Select Case Err.Number** permet une gestion spécialisée par type d'erreur
5. **Toujours prévoir le nettoyage** des ressources en cas d'erreur
6. **Documenter et journaliser** les erreurs pour faciliter la maintenance

### Modèle type complet

```vba
Sub ModeleComplet()
    ' Déclarations
    Dim resultat As Variant

    ' Activation gestion d'erreur
    On Error GoTo GestionErreur

    ' Code principal
    resultat = FonctionQuiPeutEchouer()

    ' Sortie normale
    On Error GoTo 0
    MsgBox "Succès : " & resultat
    Exit Sub

    ' Gestion d'erreur
GestionErreur:
    Select Case Err.Number
        Case 9
            ' Traitement spécifique erreur 9
        Case 11
            ' Traitement spécifique erreur 11
        Case Else
            ' Traitement générique
            MsgBox "Erreur inattendue : " & Err.Description
    End Select

    ' Nettoyage et désactivation
    On Error GoTo 0
End Sub
```

### Conseil pour progresser

Commencez par des gestionnaires simples avec juste un message d'erreur. Progressivement, ajoutez la gestion spécialisée par type d'erreur, puis les fonctionnalités avancées comme Resume et la journalisation.

`On Error GoTo` est plus puissant qu'`On Error Resume Next` mais nécessite plus de rigueur dans la structure du code. C'est l'outil de choix pour créer des applications VBA robustes et professionnelles.

Dans la section suivante, nous découvrirons en détail les propriétés `Err.Number` et `Err.Description` pour analyser finement les erreurs.

⏭️
