🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.5 Filtres automatiques et avancés

## Introduction : Qu'est-ce que le filtrage de données ?

Le **filtrage** permet de masquer temporairement certaines lignes d'un tableau pour ne voir que celles qui nous intéressent. C'est comme avoir des lunettes spéciales qui ne montrent que ce qu'on veut voir !

### Exemple concret
Imaginez une base de données de 5000 clients avec :
- Nom, Prénom, Âge, Ville, Profession, Chiffre d'affaires

Avec les filtres, vous pouvez instantanément afficher :
- 👥 Seulement les clients de Paris âgés de 25 à 40 ans
- 💼 Les architectes qui génèrent plus de 50 000€
- 🎯 Les prospects contactés la semaine dernière

**Sans VBA :** Vous cliquez manuellement sur chaque filtre.
**Avec VBA :** Vous automatisez des filtres complexes en une ligne de code !

## Types de filtres Excel

### 1. Filtres automatiques (AutoFilter)
- 🔽 Petites flèches déroulantes sur chaque colonne
- ✅ Faciles à utiliser pour des critères simples
- 📊 Parfaits pour l'exploration interactive des données

### 2. Filtres avancés (AdvancedFilter)
- 🎯 Critères complexes avec ET/OU
- 📋 Possibilité de copier les résultats ailleurs
- 🔧 Plus puissants mais plus techniques

## Partie 1 : Filtres automatiques (AutoFilter)

### Activer les filtres automatiques

```vba
Sub ActiverFiltresAutomatiques()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Activer les filtres sur la plage de données
    If Not ws.AutoFilterMode Then
        ws.Range("A1").AutoFilter  ' Active sur toute la région de données
        MsgBox "Filtres automatiques activés !"
    Else
        MsgBox "Les filtres sont déjà activés sur cette feuille."
    End If
End Sub
```

### Créer des données d'exemple pour nos tests

```vba
Sub CreerDonneesExemple()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Nettoyer la feuille
    ws.Cells.Clear

    ' En-têtes
    ws.Range("A1:F1").Value = Array("Nom", "Âge", "Ville", "Profession", "Salaire", "Date_Embauche")

    ' Données d'exemple
    ws.Range("A2:F16").Value = Array( _
        Array("Dupont", 28, "Paris", "Ingénieur", 45000, "01/01/2020"), _
        Array("Martin", 35, "Lyon", "Comptable", 38000, "15/03/2019"), _
        Array("Bernard", 42, "Paris", "Manager", 65000, "10/06/2018"), _
        Array("Petit", 29, "Marseille", "Ingénieur", 42000, "22/09/2021"), _
        Array("Durand", 31, "Paris", "Comptable", 41000, "05/12/2020"), _
        Array("Moreau", 26, "Lyon", "Ingénieur", 39000, "18/04/2022"), _
        Array("Simon", 38, "Toulouse", "Manager", 58000, "30/08/2017"), _
        Array("Michel", 33, "Paris", "Comptable", 43000, "12/02/2019"), _
        Array("Lefebvre", 27, "Marseille", "Ingénieur", 40000, "25/11/2021"), _
        Array("Leroy", 45, "Lyon", "Manager", 70000, "08/07/2016"), _
        Array("Roux", 30, "Paris", "Ingénieur", 46000, "14/05/2020"), _
        Array("Vincent", 36, "Toulouse", "Comptable", 39500, "03/10/2018"), _
        Array("Fournier", 24, "Marseille", "Ingénieur", 37000, "20/01/2023"), _
        Array("Morel", 39, "Paris", "Manager", 62000, "16/09/2017"), _
        Array("Girard", 32, "Lyon", "Comptable", 42500, "28/06/2019"))

    ' Formater les en-têtes
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
        .Borders.LineStyle = xlContinuous
    End With

    ' Ajuster la largeur des colonnes
    ws.Columns("A:F").AutoFit

    MsgBox "Données d'exemple créées ! 15 employés prêts pour le filtrage."
End Sub
```

### Filtres simples : Un critère

#### Filtrer par ville
```vba
Sub FiltrerParParis()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' S'assurer que les filtres sont activés
    If Not ws.AutoFilterMode Then
        ws.Range("A1").AutoFilter
    End If

    ' Filtrer la colonne "Ville" (colonne 3) pour "Paris"
    ws.Range("A1").AutoFilter Field:=3, Criteria1:="Paris"

    MsgBox "Affichage des employés de Paris uniquement."
End Sub
```

#### Filtrer par salaire (supérieur à 45000)
```vba
Sub FiltrerSalaireEleve()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ws.Range("A1").AutoFilter Field:=5, Criteria1:=">45000"

    MsgBox "Affichage des salaires > 45000€."
End Sub
```

#### Filtrer par âge (entre 25 et 35 ans)
```vba
Sub FiltrerTranchesAge()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Deux critères sur la même colonne : >= 25 ET <= 35
    ws.Range("A1").AutoFilter Field:=2, Criteria1:=">=25", _
                              Operator:=xlAnd, Criteria2:="<=35"

    MsgBox "Affichage des employés entre 25 et 35 ans."
End Sub
```

### Filtres multiples : Plusieurs colonnes

#### Ingénieurs de Paris avec bon salaire
```vba
Sub FiltrerIngenieursParis()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Réinitialiser tous les filtres d'abord
    ws.AutoFilter.ShowAllData

    ' Filtrer par profession
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="Ingénieur"

    ' Puis par ville
    ws.Range("A1").AutoFilter Field:=3, Criteria1:="Paris"

    ' Puis par salaire
    ws.Range("A1").AutoFilter Field:=5, Criteria1:=">40000"

    MsgBox "Affichage des ingénieurs parisiens avec salaire > 40000€."
End Sub
```

### Filtres avec critères multiples (OU)

#### Paris OU Lyon
```vba
Sub FiltrerParisOuLyon()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Pour filtrer avec "OU", utiliser un Array
    ws.Range("A1").AutoFilter Field:=3, Criteria1:=Array("Paris", "Lyon"), _
                              Operator:=xlFilterValues

    MsgBox "Affichage des employés de Paris OU Lyon."
End Sub
```

### Gestion avancée des filtres automatiques

#### Compter les lignes visibles après filtrage
```vba
Sub CompterLignesVisibles()
    Dim ws As Worksheet
    Dim plageVisible As Range
    Dim nombreLignes As Long

    Set ws = ActiveSheet

    ' Appliquer un filtre d'abord
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="Ingénieur"

    ' Compter les lignes visibles (sans l'en-tête)
    Set plageVisible = ws.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    nombreLignes = plageVisible.Rows.Count - 1  ' -1 pour exclure l'en-tête

    MsgBox nombreLignes & " ingénieur(s) trouvé(s)."
End Sub
```

#### Copier les données filtrées
```vba
Sub CopierDonneesFiltrees()
    Dim ws As Worksheet
    Dim wsDestination As Worksheet
    Dim plageVisible As Range

    Set ws = ActiveSheet

    ' Appliquer un filtre
    ws.Range("A1").AutoFilter Field:=3, Criteria1:="Paris"

    ' Créer une nouvelle feuille
    Set wsDestination = Worksheets.Add
    wsDestination.Name = "Employés_Paris"

    ' Copier seulement les lignes visibles
    Set plageVisible = ws.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    plageVisible.Copy wsDestination.Range("A1")

    ' Ajuster les colonnes
    wsDestination.Columns.AutoFit

    MsgBox "Employés parisiens copiés dans la nouvelle feuille."
End Sub
```

#### Effacer tous les filtres
```vba
Sub EffacerTousLesFiltres()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.AutoFilterMode Then
        ' Afficher toutes les données
        ws.AutoFilter.ShowAllData
        MsgBox "Tous les filtres ont été supprimés."
    Else
        MsgBox "Aucun filtre n'est actif sur cette feuille."
    End If
End Sub
```

## Partie 2 : Filtres avancés (AdvancedFilter)

### Principe des filtres avancés

Les filtres avancés utilisent une **zone de critères** séparée où vous définissez vos conditions de filtrage.

### Structure de la zone de critères
```
A1: Nom       B1: Âge      C1: Ville      D1: Salaire
A2: Dupont    B2:          C2: Paris      D2: >45000
A3:           B3: >30      C3: Lyon       D3:
```

- **Ligne 1** : Noms des champs (identiques aux en-têtes)
- **Ligne 2** : Premier jeu de critères (ET logique)
- **Ligne 3** : Deuxième jeu de critères (OU logique avec ligne 2)

### Créer votre premier filtre avancé

#### Préparation de la zone de critères
```vba
Sub Preparer ZoneCriteres()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Zone de critères en H1:K3
    ws.Range("H1:K1").Value = Array("Nom", "Âge", "Ville", "Salaire")

    ' Critère 1 : Ingénieurs de Paris avec salaire > 40000
    ws.Range("H2").Value = ""        ' Nom (vide = tous)
    ws.Range("I2").Value = ""        ' Âge (vide = tous)
    ws.Range("J2").Value = "Paris"   ' Ville = Paris
    ws.Range("K2").Value = ">40000"  ' Salaire > 40000

    ' Formater la zone de critères
    With ws.Range("H1:K3")
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(255, 255, 200)  ' Jaune clair
    End With

    ws.Range("H1:K1").Font.Bold = True

    MsgBox "Zone de critères créée en H1:K3"
End Sub
```

#### Appliquer le filtre avancé
```vba
Sub AppliquerFiltreAvance()
    Dim ws As Worksheet
    Dim plageSource As Range
    Dim plageCriteres As Range

    Set ws = ActiveSheet
    Set plageSource = ws.Range("A1:F16")      ' Données source
    Set plageCriteres = ws.Range("H1:K3")     ' Zone de critères

    ' Appliquer le filtre avancé sur place
    plageSource.AdvancedFilter Action:=xlFilterInPlace, _
                               CriteriaRange:=plageCriteres

    MsgBox "Filtre avancé appliqué ! Seules les lignes correspondantes sont visibles."
End Sub
```

### Filtre avancé avec copie des résultats

#### Copier les résultats ailleurs
```vba
Sub FiltreAvanceAvecCopie()
    Dim ws As Worksheet
    Dim plageSource As Range
    Dim plageCriteres As Range
    Dim plageDestination As Range

    Set ws = ActiveSheet
    Set plageSource = ws.Range("A1:F16")
    Set plageCriteres = ws.Range("H1:K3")
    Set plageDestination = ws.Range("A20")    ' Résultats à partir de A20

    ' Nettoyer la zone de destination d'abord
    ws.Range("A20:F50").Clear

    ' Appliquer le filtre avec copie
    plageSource.AdvancedFilter Action:=xlFilterCopy, _
                               CriteriaRange:=plageCriteres, _
                               CopyToRange:=plageDestination

    MsgBox "Résultats copiés à partir de la ligne 20."
End Sub
```

### Critères complexes avec les filtres avancés

#### Critères multiples avec ET logique
```vba
Sub CreerCriteresComplexesET()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Zone de critères pour : Âge > 30 ET Ville = Paris ET Salaire > 45000
    ws.Range("H1:K1").Value = Array("Nom", "Âge", "Ville", "Salaire")
    ws.Range("H2:K2").Value = Array("", ">30", "Paris", ">45000")

    ' Tous les critères sur la même ligne = ET logique
    MsgBox "Critères ET créés : Âge>30 ET Paris ET Salaire>45000"
End Sub
```

#### Critères multiples avec OU logique
```vba
Sub CreerCriteresComplexesOU()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Zone de critères pour : (Paris ET Salaire>45000) OU (Lyon ET Âge<30)
    ws.Range("H1:K1").Value = Array("Nom", "Âge", "Ville", "Salaire")
    ws.Range("H2:K2").Value = Array("", "", "Paris", ">45000")      ' Premier critère
    ws.Range("H3:K3").Value = Array("", "<30", "Lyon", "")         ' Deuxième critère

    ' Critères sur lignes différentes = OU logique
    MsgBox "Critères OU créés : (Paris ET Salaire>45000) OU (Lyon ET Âge<30)"
End Sub
```

### Filtres avec critères de texte

#### Recherche par motif
```vba
Sub FiltrerParMotif()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Chercher tous les noms commençant par "Du"
    ws.Range("H1:K1").Value = Array("Nom", "Âge", "Ville", "Salaire")
    ws.Range("H2").Value = "Du*"    ' * = joker pour "n'importe quoi après Du"

    ' Appliquer le filtre
    ws.Range("A1:F16").AdvancedFilter Action:=xlFilterInPlace, _
                                     CriteriaRange:=ws.Range("H1:K2")

    MsgBox "Affichage des noms commençant par 'Du'"
End Sub
```

#### Filtres avec formules personnalisées
```vba
Sub FiltrerAvecFormule()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Critère avec formule : salaire supérieur à la moyenne
    ws.Range("H1").Value = "Salaire_Sup_Moyenne"
    ws.Range("H2").Value = "=E2>MOYENNE($E$2:$E$16)"

    ' Appliquer le filtre
    ws.Range("A1:F16").AdvancedFilter Action:=xlFilterInPlace, _
                                     CriteriaRange:=ws.Range("H1:H2")

    MsgBox "Affichage des salaires supérieurs à la moyenne"
End Sub
```

## Partie 3 : Automation avancée du filtrage

### Système de filtrage interactif

#### Menu de filtrage personnalisé
```vba
Sub MenuFiltragePrincipal()
    Dim choix As String

    choix = InputBox("Choisissez votre filtre :" & vbCrLf & _
                     "1 - Tous les employés" & vbCrLf & _
                     "2 - Ingénieurs uniquement" & vbCrLf & _
                     "3 - Salaires élevés (>50000)" & vbCrLf & _
                     "4 - Jeunes employés (<30 ans)" & vbCrLf & _
                     "5 - Employés parisiens" & vbCrLf & _
                     "6 - Managers expérimentés", "Menu de filtrage", "1")

    Select Case choix
        Case "1"
            Call AfficherTous
        Case "2"
            Call FiltrerIngenieurs
        Case "3"
            Call FiltrerSalairesEleves
        Case "4"
            Call FiltrerJeunesEmployes
        Case "5"
            Call FiltrerParisiens
        Case "6"
            Call FiltrerManagersExperimentes
        Case Else
            MsgBox "Choix non valide."
    End Select
End Sub

Sub AfficherTous()
    ActiveSheet.AutoFilter.ShowAllData
    MsgBox "Tous les employés affichés."
End Sub

Sub FiltrerIngenieurs()
    ActiveSheet.Range("A1").AutoFilter Field:=4, Criteria1:="Ingénieur"
    MsgBox "Affichage des ingénieurs uniquement."
End Sub

Sub FiltrerSalairesEleves()
    ActiveSheet.Range("A1").AutoFilter Field:=5, Criteria1:=">50000"
    MsgBox "Affichage des salaires > 50000€."
End Sub
```

### Filtrage automatique par boutons

#### Créer des boutons de filtrage
```vba
Sub CreerBoutonsFiltrage()
    Dim ws As Worksheet
    Dim btn As Button

    Set ws = ActiveSheet

    ' Supprimer les anciens boutons
    ws.Buttons.Delete

    ' Bouton "Tous"
    Set btn = ws.Buttons.Add(ws.Range("H5").Left, ws.Range("H5").Top, 80, 25)
    btn.Text = "Tous"
    btn.OnAction = "AfficherTous"

    ' Bouton "Ingénieurs"
    Set btn = ws.Buttons.Add(ws.Range("I5").Left, ws.Range("I5").Top, 80, 25)
    btn.Text = "Ingénieurs"
    btn.OnAction = "FiltrerIngenieurs"

    ' Bouton "Paris"
    Set btn = ws.Buttons.Add(ws.Range("J5").Left, ws.Range("J5").Top, 80, 25)
    btn.Text = "Paris"
    btn.OnAction = "FiltrerParisiens"

    ' Bouton "Salaires +"
    Set btn = ws.Buttons.Add(ws.Range("K5").Left, ws.Range("K5").Top, 80, 25)
    btn.Text = "Salaires +"
    btn.OnAction = "FiltrerSalairesEleves"

    MsgBox "Boutons de filtrage créés !"
End Sub
```

### Rapport automatisé avec filtres

#### Générer un rapport multi-filtres
```vba
Sub GenererRapportMultiFiltres()
    Dim ws As Worksheet
    Dim wsRapport As Worksheet
    Dim ligne As Long

    Set ws = ActiveSheet

    ' Créer la feuille de rapport
    On Error Resume Next
    Worksheets("Rapport_Filtrage").Delete
    On Error GoTo 0

    Set wsRapport = Worksheets.Add
    wsRapport.Name = "Rapport_Filtrage"

    ' Titre du rapport
    wsRapport.Range("A1").Value = "RAPPORT D'ANALYSE DES EMPLOYÉS"
    wsRapport.Range("A1").Font.Size = 16
    wsRapport.Range("A1").Font.Bold = True

    ligne = 3

    ' 1. Ingénieurs
    ws.AutoFilter.ShowAllData
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="Ingénieur"
    Call CopierResultatsFiltre(ws, wsRapport, ligne, "INGÉNIEURS")
    ligne = ligne + 10

    ' 2. Managers
    ws.AutoFilter.ShowAllData
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="Manager"
    Call CopierResultatsFiltre(ws, wsRapport, ligne, "MANAGERS")
    ligne = ligne + 10

    ' 3. Employés parisiens
    ws.AutoFilter.ShowAllData
    ws.Range("A1").AutoFilter Field:=3, Criteria1:="Paris"
    Call CopierResultatsFiltre(ws, wsRapport, ligne, "EMPLOYÉS PARISIENS")

    ' Nettoyer
    ws.AutoFilter.ShowAllData

    MsgBox "Rapport multi-filtres généré dans la feuille 'Rapport_Filtrage'"
End Sub

Sub CopierResultatsFiltre(wsSource As Worksheet, wsDestination As Worksheet, _
                         ligneDebut As Long, titre As String)
    Dim plageVisible As Range

    ' Titre de la section
    wsDestination.Cells(ligneDebut, 1).Value = titre
    wsDestination.Cells(ligneDebut, 1).Font.Bold = True
    wsDestination.Cells(ligneDebut, 1).Interior.Color = RGB(200, 220, 255)

    ' Copier les données filtrées
    Set plageVisible = wsSource.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    plageVisible.Copy wsDestination.Cells(ligneDebut + 1, 1)

    ' Nettoyer le presse-papiers
    Application.CutCopyMode = False
End Sub
```

### Filtrage conditionnel intelligent

#### Filtrage adaptatif selon les données
```vba
Sub FiltrageAdaptatif()
    Dim ws As Worksheet
    Dim moyenneSalaire As Double
    Dim moyenneAge As Double

    Set ws = ActiveSheet

    ' Calculer les moyennes
    moyenneSalaire = Application.WorksheetFunction.Average(ws.Range("E2:E16"))
    moyenneAge = Application.WorksheetFunction.Average(ws.Range("B2:B16"))

    ' Créer des critères dynamiques
    ws.Range("H1:K1").Value = Array("Nom", "Âge", "Ville", "Salaire")
    ws.Range("H2:K2").Value = Array("", ">" & Int(moyenneAge), "", ">" & Int(moyenneSalaire))

    ' Appliquer le filtre
    ws.Range("A1:F16").AdvancedFilter Action:=xlFilterInPlace, _
                                     CriteriaRange:=ws.Range("H1:K2")

    MsgBox "Filtrage adaptatif appliqué :" & vbCrLf & _
           "Âge > " & Int(moyenneAge) & " ans" & vbCrLf & _
           "Salaire > " & Int(moyenneSalaire) & "€"
End Sub
```

## Partie 4 : Gestion des erreurs et optimisation

### Code robuste avec gestion d'erreurs

```vba
Sub FiltrageSécurisé()
    Dim ws As Worksheet
    Dim plageSource As Range

    On Error GoTo GestionErreur

    Set ws = ActiveSheet

    ' Vérifier qu'il y a des données
    If ws.UsedRange.Rows.Count < 2 Then
        MsgBox "Aucune donnée à filtrer."
        Exit Sub
    End If

    ' Vérifier que les filtres peuvent être activés
    Set plageSource = ws.UsedRange

    ' Désactiver les filtres existants si nécessaire
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If

    ' Activer les nouveaux filtres
    plageSource.AutoFilter

    ' Appliquer le filtre
    plageSource.AutoFilter Field:=1, Criteria1:="<>"  ' Non vides

    MsgBox "Filtrage sécurisé appliqué avec succès."
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors du filtrage : " & Err.Description
    ' Remettre en état normal si possible
    On Error Resume Next
    ws.AutoFilter.ShowAllData
End Sub
```

### Optimisation des performances

```vba
Sub FiltrageOptimise()
    ' Désactiver les mises à jour d'écran pour plus de rapidité
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Votre code de filtrage ici
    ActiveSheet.Range("A1").AutoFilter Field:=3, Criteria1:="Paris"

    ' Réactiver les mises à jour
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Filtrage optimisé terminé."
End Sub
```

## Conseils pour maîtriser les filtres

### ✅ Bonnes pratiques

1. **Données structurées** : En-têtes clairs en première ligne
2. **Pas de lignes vides** : Évitez les interruptions dans vos données
3. **Types cohérents** : Même type de données dans chaque colonne
4. **Sauvegarde** : Toujours pouvoir revenir à l'état initial
5. **Tests** : Vérifiez vos critères sur un échantillon d'abord

### ⚠️ Pièges à éviter

1. **Données manquantes** : Les cellules vides peuvent fausser les filtres
2. **Formats incohérents** : "100" et 100 sont différents pour Excel
3. **Plages incorrectes** : Vérifiez que votre plage inclut toutes les données
4. **Filtres imbriqués** : Attention aux filtres qui se cumulent
5. **Oubli de réinitialisation** : Pensez à remettre tous les filtres

### 🛠️ Outils de diagnostic

```vba
Sub DiagnosticFiltres()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Debug.Print "=== DIAGNOSTIC FILTRES ==="
    Debug.Print "Feuille : " & ws.Name
    Debug.Print "Mode AutoFilter : " & ws.AutoFilterMode
    Debug.Print "Plage de données : " & ws.UsedRange.Address
    Debug.Print "Nombre de lignes : " & ws.UsedRange.Rows.Count
    Debug.Print "Nombre de colonnes : " & ws.UsedRange.Columns.Count

    If ws.AutoFilterMode Then
        Debug.Print "Filtres actifs détectés"
    Else
        Debug.Print "Aucun filtre actif"
    End If

    ' Voir avec Ctrl+G dans l'éditeur VBA
End Sub
```

## Opérateurs de filtrage courants

| Opérateur VBA | Signification | Exemple |
|---------------|---------------|---------|
| `"=Paris"` | Égal à | Exactement "Paris" |
| `"<>Paris"` | Différent de | Tout sauf "Paris" |
| `">1000"` | Supérieur à | Plus de 1000 |
| `">=1000"` | Supérieur ou égal | 1000 ou plus |
| `"<1000"` | Inférieur à | Moins de 1000 |
| `"<=1000"` | Inférieur ou égal | 1000 ou moins |
| `"Du*"` | Commence par | Commence par "Du" |
| `"*son"` | Se termine par | Se termine par "son" |
| `"*mart*"` | Contient | Contient "mart" |

## Types d'actions pour AdvancedFilter

| Action VBA | Description |
|------------|-------------|
| `xlFilterInPlace` | Filtre sur place (masque les lignes) |
| `xlFilterCopy` | Copie les résultats ailleurs |

## Récapitulatif

Les filtres Excel automatisés avec VBA vous permettent de :

- 🔍 **Rechercher rapidement** dans de grandes bases de données
- 📊 **Créer des vues personnalisées** de vos données
- 🎯 **Automatiser des analyses répétitives** (rapports mensuels, etc.)
- 📋 **Extraire des sous-ensembles** de données pour traitement
- 🔄 **Standardiser les procédures** de filtrage dans votre équipe

**Points clés à retenir :**
- **AutoFilter** : Simple et rapide pour des critères basiques
- **AdvancedFilter** : Puissant pour des critères complexes avec ET/OU
- **Zone de critères** : Clé du succès des filtres avancés
- **Gestion d'erreurs** : Toujours vérifier vos données avant filtrage
- **Performance** : Désactiver l'affichage pour les gros volumes

**Comparaison AutoFilter vs AdvancedFilter :**

| Critère | AutoFilter | AdvancedFilter |
|---------|------------|----------------|
| **Facilité** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ |
| **Puissance** | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **Critères complexes** | ⭐⭐ | ⭐⭐⭐⭐⭐ |
| **Copie résultats** | ⭐⭐ | ⭐⭐⭐⭐⭐ |
| **Performance** | ⭐⭐⭐⭐ | ⭐⭐⭐ |

**Cas d'usage recommandés :**

🔸 **Utilisez AutoFilter pour :**
- Filtres simples sur une ou deux colonnes
- Exploration interactive des données
- Filtres temporaires et rapides
- Interface utilisateur simple

🔸 **Utilisez AdvancedFilter pour :**
- Critères complexes avec multiple ET/OU
- Extraction de données vers autre emplacement
- Filtres basés sur des formules
- Automatisation de rapports complexes

**Prochaine étape :** Vous maîtrisez maintenant tous les outils avancés d'Excel en VBA ! Ces compétences vous permettront de créer des solutions complètes d'analyse et de traitement de données.

⏭️
