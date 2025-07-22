🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.4 Tableaux croisés dynamiques

## Introduction : Qu'est-ce qu'un tableau croisé dynamique ?

Un **tableau croisé dynamique** (TCD) est l'un des outils les plus puissants d'Excel pour analyser de grandes quantités de données. Il permet de résumer, regrouper et analyser des informations de manière interactive.

### Exemple concret
Imaginez que vous avez 10 000 lignes de ventes avec :
- Date de vente
- Vendeur
- Produit
- Région
- Montant

Un TCD peut instantanément vous dire :
- 📊 Quel vendeur a le meilleur chiffre d'affaires ?
- 📈 Quelles sont les ventes par mois ?
- 🗺️ Quelle région performe le mieux ?
- 📱 Quel produit se vend le plus ?

**Sans VBA :** Vous créez manuellement chaque TCD, un par un.
**Avec VBA :** Vous automatisez la création de dizaines de TCD en quelques secondes !

## Pourquoi automatiser les TCD avec VBA ?

### Avantages de l'automatisation

✅ **Gain de temps** : Créer plusieurs analyses en une fois
✅ **Reproductibilité** : Même analyse chaque mois/semaine
✅ **Cohérence** : Toujours la même mise en forme
✅ **Actualisation automatique** : Données toujours à jour
✅ **Rapports standardisés** : Format uniforme pour toute l'équipe

### Cas d'usage typiques
- 📈 Rapports mensuels automatisés
- 🎯 Tableaux de bord interactifs
- 📊 Analyses multi-critères
- 🔄 Actualisation de données externe

## Anatomie d'un tableau croisé dynamique

### Les 4 zones principales

```
┌─────────────────┬─────────────────┐
│   FILTRES       │                 │
├─────────────────┼─────────────────┤
│                 │    COLONNES     │
│     LIGNES      ├─────────────────┤
│                 │    VALEURS      │
└─────────────────┴─────────────────┘
```

**Zone FILTRES** : Filtre général sur toutes les données
**Zone LIGNES** : Ce qui apparaît en lignes (ex: Produits)
**Zone COLONNES** : Ce qui apparaît en colonnes (ex: Mois)
**Zone VALEURS** : Les calculs (ex: Somme des ventes)

## Votre premier TCD en VBA

### Préparation des données d'exemple
```vba
Sub PreparerDonneesVentes()
    ' Créer des données d'exemple pour nos TCD
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' En-têtes
    ws.Range("A1:E1").Value = Array("Date", "Vendeur", "Produit", "Région", "Montant")

    ' Quelques données d'exemple
    ws.Range("A2:E11").Value = Array( _
        Array("01/01/2024", "Pierre", "Ordinateur", "Nord", 1200), _
        Array("02/01/2024", "Marie", "Souris", "Sud", 25), _
        Array("03/01/2024", "Paul", "Clavier", "Est", 45), _
        Array("04/01/2024", "Pierre", "Écran", "Nord", 300), _
        Array("05/01/2024", "Marie", "Ordinateur", "Sud", 1100), _
        Array("06/01/2024", "Paul", "Souris", "Est", 30), _
        Array("07/01/2024", "Pierre", "Clavier", "Nord", 50), _
        Array("08/01/2024", "Marie", "Écran", "Sud", 280), _
        Array("09/01/2024", "Paul", "Ordinateur", "Est", 1250), _
        Array("10/01/2024", "Pierre", "Souris", "Nord", 28))

    MsgBox "Données d'exemple créées !"
End Sub
```

### Créer votre premier TCD
```vba
Sub CreerPremierTCD()
    Dim ws As Worksheet
    Dim wsDestination As Worksheet
    Dim plageSource As Range
    Dim cache As PivotCache
    Dim tcd As PivotTable

    ' Feuille source (avec les données)
    Set ws = ActiveSheet
    Set plageSource = ws.Range("A1:E11")  ' Ajustez selon vos données

    ' Créer une nouvelle feuille pour le TCD
    Set wsDestination = Worksheets.Add
    wsDestination.Name = "TCD_Ventes_Par_Vendeur"

    ' Créer le cache (données en mémoire)
    Set cache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=plageSource)

    ' Créer le tableau croisé dynamique
    Set tcd = cache.CreatePivotTable( _
        TableDestination:=wsDestination.Range("A1"), _
        TableName:="TCD_Vendeurs")

    ' Configurer le TCD
    With tcd
        ' Ajouter Vendeur en lignes
        .PivotFields("Vendeur").Orientation = xlRowField

        ' Ajouter Montant en valeurs (somme)
        .PivotFields("Montant").Orientation = xlDataField
        .PivotFields("Somme de Montant").Function = xlSum
    End With

    MsgBox "Premier TCD créé ! Ventes par vendeur."
End Sub
```

## TCD plus complexe : Multi-critères

### Ventes par vendeur et par produit
```vba
Sub CreerTCDVendeurProduit()
    Dim ws As Worksheet
    Dim wsDestination As Worksheet
    Dim plageSource As Range
    Dim cache As PivotCache
    Dim tcd As PivotTable

    Set ws = Worksheets("Feuil1")  ' Feuille avec les données
    Set plageSource = ws.Range("A1:E11")

    ' Nouvelle feuille
    Set wsDestination = Worksheets.Add
    wsDestination.Name = "TCD_Vendeur_Produit"

    ' Créer le TCD
    Set cache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=plageSource)

    Set tcd = cache.CreatePivotTable( _
        TableDestination:=wsDestination.Range("A1"), _
        TableName:="TCD_VendeurProduit")

    With tcd
        ' Vendeur en lignes
        .PivotFields("Vendeur").Orientation = xlRowField

        ' Produit en colonnes
        .PivotFields("Produit").Orientation = xlColumnField

        ' Montant en valeurs
        .PivotFields("Montant").Orientation = xlDataField
        .PivotFields("Somme de Montant").Function = xlSum

        ' Région en filtre
        .PivotFields("Région").Orientation = xlPageField
    End With

    MsgBox "TCD Vendeur x Produit créé avec filtre Région !"
End Sub
```

## Personnaliser l'apparence des TCD

### Mise en forme et options
```vba
Sub PersonnaliserTCD()
    Dim tcd As PivotTable

    ' Récupérer le TCD existant
    Set tcd = ActiveSheet.PivotTables("TCD_VendeurProduit")

    With tcd
        ' Options d'affichage
        .ShowTableStyleRowStripes = True       ' Lignes alternées
        .ShowTableStyleColumnStripes = False   ' Pas de colonnes alternées
        .TableStyle2 = "PivotStyleMedium2"     ' Style prédéfini

        ' Sous-totaux
        .PivotFields("Vendeur").Subtotals(1) = True  ' Sous-totaux pour vendeurs

        ' Totaux généraux
        .RowGrand = True    ' Total général en ligne
        .ColumnGrand = True ' Total général en colonne

        ' Format des nombres
        .PivotFields("Somme de Montant").NumberFormat = "# ##0 €"

        ' Titre personnalisé
        .Name = "Analyse Ventes Détaillée"
    End With

    MsgBox "TCD personnalisé !"
End Sub
```

## Actualiser les données

### Actualisation simple
```vba
Sub ActualiserTCD()
    Dim tcd As PivotTable

    ' Actualiser un TCD spécifique
    Set tcd = ActiveSheet.PivotTables(1)  ' Premier TCD de la feuille
    tcd.RefreshTable

    MsgBox "TCD actualisé !"
End Sub
```

### Actualiser tous les TCD du classeur
```vba
Sub ActualiserTousLesTCD()
    Dim ws As Worksheet
    Dim tcd As PivotTable
    Dim compteur As Integer

    compteur = 0

    ' Parcourir toutes les feuilles
    For Each ws In ThisWorkbook.Worksheets
        ' Parcourir tous les TCD de chaque feuille
        For Each tcd In ws.PivotTables
            tcd.RefreshTable
            compteur = compteur + 1
        Next tcd
    Next ws

    MsgBox compteur & " tableau(x) croisé(s) actualisé(s) !"
End Sub
```

## TCD avec calculs personnalisés

### Ajout de champs calculés
```vba
Sub AjouterChampCalcule()
    Dim tcd As PivotTable
    Dim champCalcule As PivotField

    Set tcd = ActiveSheet.PivotTables(1)

    ' Ajouter un champ calculé (ex: Commission = 5% des ventes)
    Set champCalcule = tcd.CalculatedFields.Add( _
        Name:="Commission", _
        Formula:="=Montant*0.05")

    ' Ajouter ce champ aux valeurs
    champCalcule.Orientation = xlDataField

    ' Formater
    tcd.PivotFields("Commission").NumberFormat = "# ##0,00 €"

    MsgBox "Champ 'Commission' ajouté au TCD !"
End Sub
```

## Filtrage automatique des TCD

### Appliquer des filtres par code
```vba
Sub FiltrerTCD()
    Dim tcd As PivotTable
    Dim champFiltre As PivotField

    Set tcd = ActiveSheet.PivotTables(1)
    Set champFiltre = tcd.PivotFields("Région")

    ' Désactiver tous les éléments d'abord
    champFiltre.ClearAllFilters

    ' Activer seulement certaines régions
    champFiltre.PivotItems("Nord").Visible = True
    champFiltre.PivotItems("Sud").Visible = True
    champFiltre.PivotItems("Est").Visible = False  ' Masquer l'Est

    MsgBox "Filtre appliqué : Nord et Sud uniquement"
End Sub
```

### Filtrer par valeurs
```vba
Sub FiltrerParMontant()
    Dim tcd As PivotTable

    Set tcd = ActiveSheet.PivotTables(1)

    ' Filtrer pour ne montrer que les ventes > 100€
    With tcd.PivotFields("Montant")
        .AutoSort xlDescending, "Montant"  ' Trier par montant décroissant
        .PivotFilters.Add Type:=xlValueIsGreaterThan, Value1:=100
    End With

    MsgBox "Affichage des ventes > 100€ uniquement"
End Sub
```

## Créer plusieurs TCD automatiquement

### Générateur de rapports multiples
```vba
Sub CreerRapportsMultiples()
    Dim ws As Worksheet
    Dim plageSource As Range
    Dim cache As PivotCache

    Set ws = Worksheets("Feuil1")  ' Feuille source
    Set plageSource = ws.Range("A1:E11")

    ' Créer le cache une seule fois (plus efficace)
    Set cache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=plageSource)

    ' Rapport 1 : Ventes par région
    Call CreerTCDParRegion(cache)

    ' Rapport 2 : Ventes par produit
    Call CreerTCDParProduit(cache)

    ' Rapport 3 : Évolution mensuelle
    Call CreerTCDParMois(cache)

    MsgBox "3 rapports TCD créés automatiquement !"
End Sub

Sub CreerTCDParRegion(cache As PivotCache)
    Dim wsDestination As Worksheet
    Dim tcd As PivotTable

    Set wsDestination = Worksheets.Add
    wsDestination.Name = "Rapport_Regions"

    Set tcd = cache.CreatePivotTable( _
        TableDestination:=wsDestination.Range("A1"), _
        TableName:="TCD_Regions")

    With tcd
        .PivotFields("Région").Orientation = xlRowField
        .PivotFields("Montant").Orientation = xlDataField
        .PivotFields("Somme de Montant").NumberFormat = "# ##0 €"
    End With
End Sub

Sub CreerTCDParProduit(cache As PivotCache)
    Dim wsDestination As Worksheet
    Dim tcd As PivotTable

    Set wsDestination = Worksheets.Add
    wsDestination.Name = "Rapport_Produits"

    Set tcd = cache.CreatePivotTable( _
        TableDestination:=wsDestination.Range("A1"), _
        TableName:="TCD_Produits")

    With tcd
        .PivotFields("Produit").Orientation = xlRowField
        .PivotFields("Montant").Orientation = xlDataField
        .PivotFields("Somme de Montant").NumberFormat = "# ##0 €"
        ' Trier par montant décroissant
        .PivotFields("Produit").AutoSort xlDescending, "Somme de Montant"
    End With
End Sub

Sub CreerTCDParMois(cache As PivotCache)
    Dim wsDestination As Worksheet
    Dim tcd As PivotTable

    Set wsDestination = Worksheets.Add
    wsDestination.Name = "Rapport_Mensuel"

    Set tcd = cache.CreatePivotTable( _
        TableDestination:=wsDestination.Range("A1"), _
        TableName:="TCD_Mensuel")

    With tcd
        .PivotFields("Date").Orientation = xlRowField
        .PivotFields("Montant").Orientation = xlDataField
        .PivotFields("Somme de Montant").NumberFormat = "# ##0 €"

        ' Grouper les dates par mois
        .PivotFields("Date").LabelRange.Group Start:=True, End:=True, _
            Periods:=Array(False, False, False, False, True, False, False)
    End With
End Sub
```

## Exporter les données d'un TCD

### Copier les résultats vers une nouvelle feuille
```vba
Sub ExporterDonneesTCD()
    Dim tcd As PivotTable
    Dim plageData As Range
    Dim wsExport As Worksheet

    Set tcd = ActiveSheet.PivotTables(1)

    ' Obtenir la plage de données du TCD
    Set plageData = tcd.TableRange1

    ' Créer une nouvelle feuille pour l'export
    Set wsExport = Worksheets.Add
    wsExport.Name = "Export_" & Format(Now, "ddmmyy_hhnn")

    ' Copier les données (valeurs uniquement)
    plageData.Copy
    wsExport.Range("A1").PasteSpecial xlPasteValues

    ' Nettoyer
    Application.CutCopyMode = False

    MsgBox "Données TCD exportées vers : " & wsExport.Name
End Sub
```

## Automatisation complète : Tableau de bord

### Créer un tableau de bord automatisé
```vba
Sub CreerTableauDeBordComplet()
    Dim wsSource As Worksheet
    Dim wsDashboard As Worksheet
    Dim plageSource As Range

    ' Nettoyer et préparer
    Application.ScreenUpdating = False

    Set wsSource = Worksheets("Feuil1")
    Set plageSource = wsSource.UsedRange  ' Toutes les données utilisées

    ' Créer la feuille tableau de bord
    On Error Resume Next
    Worksheets("Tableau_de_Bord").Delete
    On Error GoTo 0

    Set wsDashboard = Worksheets.Add
    wsDashboard.Name = "Tableau_de_Bord"

    ' Titre principal
    With wsDashboard.Range("A1")
        .Value = "TABLEAU DE BORD VENTES"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
    End With
    wsDashboard.Range("A1:H1").Merge

    ' TCD 1 : Top 3 des vendeurs (A3)
    Call CreerTCDTopVendeurs(plageSource, wsDashboard.Range("A3"))

    ' TCD 2 : Répartition par région (E3)
    Call CreerTCDRepartitionRegion(plageSource, wsDashboard.Range("E3"))

    ' TCD 3 : Performance produits (A10)
    Call CreerTCDProduits(plageSource, wsDashboard.Range("A10"))

    ' Actualiser tout
    wsDashboard.PivotTables(1).RefreshTable
    wsDashboard.PivotTables(2).RefreshTable
    wsDashboard.PivotTables(3).RefreshTable

    Application.ScreenUpdating = True
    MsgBox "Tableau de bord créé et actualisé !"
End Sub

Sub CreerTCDTopVendeurs(source As Range, destination As Range)
    Dim cache As PivotCache
    Dim tcd As PivotTable

    Set cache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=source)

    Set tcd = cache.CreatePivotTable( _
        TableDestination:=destination, _
        TableName:="TCD_TopVendeurs")

    With tcd
        .PivotFields("Vendeur").Orientation = xlRowField
        .PivotFields("Montant").Orientation = xlDataField
        .PivotFields("Somme de Montant").NumberFormat = "# ##0 €"
        .PivotFields("Vendeur").AutoSort xlDescending, "Somme de Montant"
        .Name = "Top Vendeurs"
    End With
End Sub
```

## Gestion des erreurs avec les TCD

### Code robuste avec gestion d'erreurs
```vba
Sub CreerTCDSecurise()
    Dim ws As Worksheet
    Dim plageSource As Range
    Dim cache As PivotCache
    Dim tcd As PivotTable

    On Error GoTo GestionErreur

    Set ws = ActiveSheet

    ' Vérifier qu'il y a des données
    If ws.UsedRange.Rows.Count < 2 Then
        MsgBox "Pas assez de données pour créer un TCD."
        Exit Sub
    End If

    Set plageSource = ws.UsedRange

    ' Vérifier que la première ligne contient des en-têtes
    If IsEmpty(ws.Cells(1, 1).Value) Then
        MsgBox "La première ligne doit contenir les en-têtes."
        Exit Sub
    End If

    ' Créer le TCD
    Set cache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=plageSource)

    Set tcd = cache.CreatePivotTable( _
        TableDestination:=Worksheets.Add.Range("A1"), _
        TableName:="TCD_" & Format(Now, "hhmmss"))

    MsgBox "TCD créé avec succès !"
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de la création du TCD : " & Err.Description
End Sub
```

## Conseils pour débuter avec les TCD en VBA

### ✅ Bonnes pratiques

1. **Données bien structurées** : En-têtes en première ligne, pas de lignes vides
2. **Cache unique** : Réutilisez le même PivotCache pour plusieurs TCD
3. **Noms explicites** : Donnez des noms clairs à vos TCD
4. **Gestion d'erreurs** : Vérifiez toujours vos données source
5. **Actualisation** : Pensez à actualiser vos TCD quand les données changent

### ⚠️ Pièges à éviter

1. **Données incomplètes** : Vérifiez que vos colonnes sont complètes
2. **Plages incorrectes** : Assurez-vous que votre plage inclut toutes les données
3. **Noms en double** : Chaque TCD doit avoir un nom unique
4. **Feuilles supprimées** : Ne supprimez pas les feuilles source des TCD
5. **Types de données** : Attention aux dates et nombres mal formatés

### 🛠️ Outils de débogage

```vba
Sub DebugTCD()
    Dim ws As Worksheet
    Dim tcd As PivotTable

    Set ws = ActiveSheet

    Debug.Print "=== INFORMATION TCD ==="
    Debug.Print "Feuille : " & ws.Name
    Debug.Print "Nombre de TCD : " & ws.PivotTables.Count

    For Each tcd In ws.PivotTables
        Debug.Print "TCD : " & tcd.Name
        Debug.Print "Source : " & tcd.SourceData
        Debug.Print "---"
    Next tcd

    ' Voir avec Ctrl+G dans l'éditeur VBA
End Sub
```

## Types de fonctions communes dans les TCD

| Fonction VBA | Description | Exemple d'usage |
|--------------|-------------|-----------------|
| `xlSum` | Somme | Total des ventes |
| `xlAverage` | Moyenne | CA moyen par vendeur |
| `xlCount` | Nombre d'éléments | Nombre de commandes |
| `xlMax` | Valeur maximale | Plus grosse vente |
| `xlMin` | Valeur minimale | Plus petite vente |
| `xlProduct` | Produit | Calculs composés |

## Orientation des champs

| Orientation VBA | Zone du TCD | Utilisation |
|-----------------|-------------|-------------|
| `xlRowField` | Lignes | Catégories principales |
| `xlColumnField` | Colonnes | Sous-catégories |
| `xlDataField` | Valeurs | Calculs et métriques |
| `xlPageField` | Filtres | Filtres généraux |

## Récapitulatif

Les tableaux croisés dynamiques en VBA vous permettent de :

- 📊 **Automatiser l'analyse** de grandes quantités de données
- 🎯 **Créer des rapports standardisés** reproductibles
- ⚡ **Gagner un temps considérable** sur l'analyse manuelle
- 🔄 **Actualiser automatiquement** vos analyses
- 📈 **Générer des tableaux de bord** complets

**Points clés à retenir :**
- Un TCD transforme des données brutes en analyses synthétiques
- VBA permet d'automatiser complètement la création et la mise à jour
- Toujours bien structurer vos données source
- Utilisez un PivotCache pour plusieurs TCD basés sur les mêmes données

**Prochaine étape :** Nous verrons maintenant comment créer des systèmes de filtrage avancés pour traiter efficacement vos bases de données !

⏭️
