🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.3. Outil d'analyse de données

## Vue d'ensemble du projet

### Contexte et problématique
À l'ère du numérique, nous sommes submergés par une quantité énorme de données : ventes, performances, sondages, mesures, statistiques... Ces données brutes, bien qu'importantes, restent souvent inexploitées car leur analyse manuelle est complexe et chronophage.

Les défis courants de l'analyse de données incluent :
- **Volume important** : Milliers de lignes à traiter manuellement
- **Calculs complexes** : Statistiques avancées difficiles à réaliser
- **Visualisation** : Transformer les chiffres en graphiques compréhensibles
- **Comparaisons** : Identifier les tendances et les écarts
- **Répétitivité** : Même analyse à refaire régulièrement

### Problèmes de l'approche manuelle
L'analyse traditionnelle avec Excel présente des limites :
- **Temps considérable** : Heures passées sur des calculs répétitifs
- **Risque d'erreurs** : Formules complexes mal saisies
- **Manque de cohérence** : Analyses différentes selon l'utilisateur
- **Difficulté de mise à jour** : Reprendre tout le travail à chaque nouvelle donnée
- **Présentation** : Graphiques non standardisés, peu professionnels

### Solution proposée
Notre outil d'analyse de données va automatiser complètement ce processus en proposant :
- **Analyse automatique** : Statistiques descriptives instantanées
- **Visualisations dynamiques** : Graphiques générés automatiquement
- **Détection d'anomalies** : Identification des valeurs aberrantes
- **Comparaisons temporelles** : Évolution et tendances
- **Rapports professionnels** : Documents formatés prêts à présenter

### Objectifs du projet
À la fin de ce projet, vous disposerez d'un analyseur capable de :
1. **Importer** des données depuis différentes sources
2. **Calculer automatiquement** toutes les statistiques essentielles
3. **Générer** des graphiques adaptés au type de données
4. **Détecter** les tendances et anomalies
5. **Créer** des tableaux de bord interactifs
6. **Exporter** des rapports d'analyse complets

## Analyse des besoins

### Types d'analyses supportées

#### 1. Analyse descriptive
- **Mesures de tendance centrale** : Moyenne, médiane, mode
- **Mesures de dispersion** : Écart-type, variance, étendue
- **Quartiles et percentiles** : Répartition des données
- **Comptages et fréquences** : Distribution des valeurs

#### 2. Analyse comparative
- **Comparaisons temporelles** : Évolution dans le temps
- **Comparaisons catégorielles** : Différences entre groupes
- **Analyses de corrélation** : Relations entre variables
- **Benchmarking** : Comparaison avec des références

#### 3. Analyse visuelle
- **Graphiques en colonnes** : Comparaisons simples
- **Graphiques en lignes** : Évolutions temporelles
- **Graphiques en secteurs** : Répartitions
- **Histogrammes** : Distributions de fréquences
- **Nuages de points** : Corrélations

#### 4. Détection d'anomalies
- **Valeurs aberrantes** : Points hors normes statistiques
- **Tendances inhabituelles** : Évolutions anormales
- **Données manquantes** : Identification des lacunes
- **Incohérences** : Données contradictoires

### Structure des données supportées

Notre outil sera conçu pour analyser différents types de données :

#### Format standard attendu
```
Date       | Vendeur | Région    | Produit   | Quantité | CA      | Objectif
01/01/2024 | Martin  | Nord      | OrdiA     | 5        | 4000    | 3500
02/01/2024 | Durand  | Sud       | OrdiB     | 3        | 2400    | 2000
03/01/2024 | Martin  | Nord      | OrdiA     | 7        | 5600    | 4000
```

#### Types de colonnes supportées
- **Dates** : Analyses temporelles et saisonnalité
- **Catégories** : Regroupements et comparaisons
- **Nombres** : Calculs statistiques complets
- **Texte** : Comptages et classifications

## Conception de la solution

### Architecture du système

```
Outil d'Analyse de Données
├── Module Principal
│   ├── Interface de démarrage
│   ├── Sélection des données
│   └── Configuration de l'analyse
├── Module Importation
│   ├── Détection automatique du format
│   ├── Nettoyage des données
│   └── Validation de la structure
├── Module Statistiques
│   ├── Calculs descriptifs
│   ├── Analyses de corrélation
│   └── Tests de significativité
├── Module Visualisation
│   ├── Génération automatique de graphiques
│   ├── Adaptation au type de données
│   └── Formatage professionnel
├── Module Anomalies
│   ├── Détection statistique
│   ├── Analyse des tendances
│   └── Signalement des incohérences
└── Module Rapport
    ├── Synthèse automatique
    ├── Recommandations
    └── Export multi-format
```

### Workflow d'analyse

1. **Sélection des données** → Choix de la plage à analyser
2. **Préparation** → Nettoyage et validation automatiques
3. **Analyse descriptive** → Calcul de toutes les statistiques
4. **Analyse visuelle** → Génération des graphiques appropriés
5. **Détection d'anomalies** → Identification des points d'attention
6. **Synthèse** → Création du rapport final
7. **Export** → Sauvegarde et partage des résultats

## Développement de la solution

### 1. Module principal - Interface de démarrage

```vba
Sub StartDataAnalysis()
    '=========================================
    ' Point d'entrée principal de l'outil d'analyse
    '=========================================

    Dim sourceRange As Range
    Dim analysisSheet As Worksheet
    Dim dataHeaders As Variant
    Dim dataTypes As Variant

    ' Message de bienvenue
    MsgBox "Bienvenue dans l'Outil d'Analyse de Données VBA !" & vbNewLine & vbNewLine & _
           "Cet outil va automatiquement :" & vbNewLine & _
           "• Analyser vos données" & vbNewLine & _
           "• Calculer les statistiques" & vbNewLine & _
           "• Créer des graphiques" & vbNewLine & _
           "• Détecter les anomalies" & vbNewLine & _
           "• Générer un rapport complet", vbInformation, "Analyseur de Données"

    ' Étape 1 : Sélection des données
    Set sourceRange = SelectDataRange()
    If sourceRange Is Nothing Then Exit Sub

    ' Étape 2 : Validation des données
    If Not ValidateDataStructure(sourceRange) Then Exit Sub

    ' Étape 3 : Analyse des types de colonnes
    dataHeaders = GetDataHeaders(sourceRange)
    dataTypes = AnalyzeColumnTypes(sourceRange)

    ' Étape 4 : Création de la feuille d'analyse
    Set analysisSheet = CreateAnalysisSheet()

    ' Étape 5 : Lancement de l'analyse complète
    Call PerformCompleteAnalysis(sourceRange, analysisSheet, dataHeaders, dataTypes)

    ' Message de fin
    MsgBox "Analyse terminée avec succès !" & vbNewLine & _
           "Consultez la feuille 'ANALYSE_DONNEES' pour voir les résultats.", vbInformation
End Sub

Function SelectDataRange() As Range
    '=========================================
    ' Sélection intelligente de la plage de données
    '=========================================

    Dim userRange As Range
    Dim response As VbMsgBoxResult

    ' Proposer d'utiliser la sélection actuelle ou toutes les données
    If Selection.Cells.Count > 1 Then
        response = MsgBox("Analyser la sélection actuelle ?" & vbNewLine & vbNewLine & _
                         "OUI = Analyser la sélection" & vbNewLine & _
                         "NON = Sélectionner une autre plage" & vbNewLine & _
                         "ANNULER = Arrêter", vbYesNoCancel + vbQuestion)

        Select Case response
            Case vbYes
                Set userRange = Selection
            Case vbNo
                ' Laisser l'utilisateur sélectionner
                On Error Resume Next
                Set userRange = Application.InputBox("Sélectionnez la plage de données à analyser (avec en-têtes) :", _
                                                   "Sélection des données", Type:=8)
                On Error GoTo 0
            Case vbCancel
                Set SelectDataRange = Nothing
                Exit Function
        End Select
    Else
        ' Détecter automatiquement la plage de données
        Set userRange = ActiveSheet.UsedRange

        If userRange.Rows.Count < 2 Then
            MsgBox "Aucune donnée détectée sur cette feuille.", vbExclamation
            Set SelectDataRange = Nothing
            Exit Function
        End If
    End If

    ' Vérifier que la plage contient des données
    If userRange Is Nothing Then
        Set SelectDataRange = Nothing
    ElseIf userRange.Rows.Count < 2 Then
        MsgBox "La plage sélectionnée doit contenir au moins une ligne d'en-têtes et une ligne de données.", vbExclamation
        Set SelectDataRange = Nothing
    Else
        Set SelectDataRange = userRange
    End If
End Function

Function ValidateDataStructure(dataRange As Range) As Boolean
    '=========================================
    ' Validation de la structure des données
    '=========================================

    Dim col As Long
    Dim emptyHeaders As Long
    Dim emptyColumns As Long

    ValidateDataStructure = True
    emptyHeaders = 0
    emptyColumns = 0

    ' Vérifier les en-têtes
    For col = 1 To dataRange.Columns.Count
        If Trim(dataRange.Cells(1, col).Value) = "" Then
            emptyHeaders = emptyHeaders + 1
        End If

        ' Vérifier que la colonne contient des données
        If Application.CountA(dataRange.Columns(col)) <= 1 Then  ' Seulement l'en-tête
            emptyColumns = emptyColumns + 1
        End If
    Next col

    ' Signaler les problèmes détectés
    If emptyHeaders > 0 Then
        MsgBox "Attention : " & emptyHeaders & " colonne(s) sans en-tête détectée(s)." & vbNewLine & _
               "L'analyse pourrait être incomplète.", vbExclamation
    End If

    If emptyColumns > 0 Then
        MsgBox "Attention : " & emptyColumns & " colonne(s) vide(s) détectée(s)." & vbNewLine & _
               "Ces colonnes seront ignorées dans l'analyse.", vbExclamation
    End If

    ' Vérifier qu'il reste suffisamment de données
    If (dataRange.Columns.Count - emptyColumns) < 1 Then
        MsgBox "Erreur : Aucune colonne de données valide trouvée.", vbCritical
        ValidateDataStructure = False
    End If

    If dataRange.Rows.Count < 3 Then  ' En-tête + au moins 2 lignes de données
        MsgBox "Attention : Peu de données disponibles pour une analyse statistique fiable." & vbNewLine & _
               "Résultats à interpréter avec précaution.", vbExclamation
    End If
End Function

Function GetDataHeaders(dataRange As Range) As Variant
    '=========================================
    ' Extraction des en-têtes de colonnes
    '=========================================

    Dim headers() As String
    Dim col As Long

    ReDim headers(1 To dataRange.Columns.Count)

    For col = 1 To dataRange.Columns.Count
        headers(col) = Trim(dataRange.Cells(1, col).Value)
        If headers(col) = "" Then
            headers(col) = "Colonne_" & col  ' Nom par défaut
        End If
    Next col

    GetDataHeaders = headers
End Function

Function AnalyzeColumnTypes(dataRange As Range) As Variant
    '=========================================
    ' Analyse automatique des types de données par colonne
    '=========================================

    Dim columnTypes() As String
    Dim col As Long
    Dim row As Long
    Dim numericCount As Long
    Dim dateCount As Long
    Dim textCount As Long
    Dim totalRows As Long
    Dim cellValue As Variant

    ReDim columnTypes(1 To dataRange.Columns.Count)
    totalRows = dataRange.Rows.Count - 1  ' Exclure l'en-tête

    For col = 1 To dataRange.Columns.Count
        numericCount = 0
        dateCount = 0
        textCount = 0

        ' Analyser un échantillon de cellules (max 100 pour performance)
        Dim sampleSize As Long
        sampleSize = Application.Min(100, totalRows)

        For row = 2 To sampleSize + 1  ' Commencer après l'en-tête
            cellValue = dataRange.Cells(row, col).Value

            If cellValue <> "" Then
                If IsNumeric(cellValue) Then
                    numericCount = numericCount + 1
                ElseIf IsDate(cellValue) Then
                    dateCount = dateCount + 1
                Else
                    textCount = textCount + 1
                End If
            End If
        Next row

        ' Déterminer le type majoritaire
        If dateCount > numericCount And dateCount > textCount Then
            columnTypes(col) = "DATE"
        ElseIf numericCount > textCount Then
            columnTypes(col) = "NUMERIC"
        Else
            columnTypes(col) = "TEXT"
        End If
    Next col

    AnalyzeColumnTypes = columnTypes
End Function
```

### 2. Module de création de l'espace d'analyse

```vba
Function CreateAnalysisSheet() As Worksheet
    '=========================================
    ' Création de la feuille d'analyse avec structure prédéfinie
    '=========================================

    Dim ws As Worksheet
    Dim existingSheet As Worksheet

    ' Supprimer la feuille existante si elle existe
    On Error Resume Next
    Set existingSheet = ThisWorkbook.Sheets("ANALYSE_DONNEES")
    If Not existingSheet Is Nothing Then
        Application.DisplayAlerts = False
        existingSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Créer la nouvelle feuille
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "ANALYSE_DONNEES"

    ' Configurer l'en-tête principal
    With ws
        .Cells(1, 1).Value = "RAPPORT D'ANALYSE DE DONNÉES"
        .Cells(1, 1).Font.Size = 18
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 51, 102)
        .Range("A1:L1").Merge
        .Range("A1").HorizontalAlignment = xlCenter

        .Cells(2, 1).Value = "Généré automatiquement le " & Format(Now(), "dd/mm/yyyy à hh:mm")
        .Range("A2:L2").Merge
        .Range("A2").HorizontalAlignment = xlCenter
        .Cells(2, 1).Font.Italic = True

        ' Ligne de séparation
        .Range("A3:L3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A3:L3").Borders(xlEdgeBottom).Weight = xlMedium

        ' Sections prédéfinies
        .Cells(5, 1).Value = "1. RÉSUMÉ EXÉCUTIF"
        .Cells(5, 1).Font.Size = 14
        .Cells(5, 1).Font.Bold = True
        .Cells(5, 1).Font.Color = RGB(0, 51, 102)

        .Cells(15, 1).Value = "2. STATISTIQUES DESCRIPTIVES"
        .Cells(15, 1).Font.Size = 14
        .Cells(15, 1).Font.Bold = True
        .Cells(15, 1).Font.Color = RGB(0, 51, 102)

        .Cells(35, 1).Value = "3. ANALYSES VISUELLES"
        .Cells(35, 1).Font.Size = 14
        .Cells(35, 1).Font.Bold = True
        .Cells(35, 1).Font.Color = RGB(0, 51, 102)

        .Cells(55, 1).Value = "4. DÉTECTION D'ANOMALIES"
        .Cells(55, 1).Font.Size = 14
        .Cells(55, 1).Font.Bold = True
        .Cells(55, 1).Font.Color = RGB(0, 51, 102)

        .Cells(70, 1).Value = "5. RECOMMANDATIONS"
        .Cells(70, 1).Font.Size = 14
        .Cells(70, 1).Font.Bold = True
        .Cells(70, 1).Font.Color = RGB(0, 51, 102)
    End With

    Set CreateAnalysisSheet = ws
End Function

Sub PerformCompleteAnalysis(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Orchestration de l'analyse complète
    '=========================================

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Afficher un message de progression
    Application.StatusBar = "Analyse en cours... Calcul des statistiques"

    ' 1. Créer le résumé exécutif
    Call CreateExecutiveSummary(sourceRange, analysisSheet, headers, columnTypes)

    Application.StatusBar = "Analyse en cours... Statistiques descriptives"

    ' 2. Calculer les statistiques descriptives
    Call CalculateDescriptiveStats(sourceRange, analysisSheet, headers, columnTypes)

    Application.StatusBar = "Analyse en cours... Création des graphiques"

    ' 3. Générer les visualisations
    Call CreateVisualizations(sourceRange, analysisSheet, headers, columnTypes)

    Application.StatusBar = "Analyse en cours... Détection d'anomalies"

    ' 4. Détecter les anomalies
    Call DetectAnomalies(sourceRange, analysisSheet, headers, columnTypes)

    Application.StatusBar = "Analyse en cours... Génération des recommandations"

    ' 5. Générer les recommandations
    Call GenerateRecommendations(sourceRange, analysisSheet, headers, columnTypes)

    ' Finalisation
    analysisSheet.Columns.AutoFit
    analysisSheet.Activate
    analysisSheet.Cells(1, 1).Select

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub
```

### 3. Module de résumé exécutif

```vba
Sub CreateExecutiveSummary(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Création du résumé exécutif automatique
    '=========================================

    Dim numericColumns As Long
    Dim textColumns As Long
    Dim dateColumns As Long
    Dim totalRows As Long
    Dim col As Long
    Dim startRow As Long

    startRow = 6  ' Position après le titre de section
    totalRows = sourceRange.Rows.Count - 1  ' Exclure l'en-tête

    ' Compter les types de colonnes
    For col = 1 To UBound(columnTypes)
        Select Case columnTypes(col)
            Case "NUMERIC"
                numericColumns = numericColumns + 1
            Case "TEXT"
                textColumns = textColumns + 1
            Case "DATE"
                dateColumns = dateColumns + 1
        End Select
    Next col

    With analysisSheet
        ' Informations générales sur le dataset
        .Cells(startRow, 1).Value = "📊 APERÇU DES DONNÉES"
        .Cells(startRow, 1).Font.Bold = True
        .Cells(startRow, 1).Font.Size = 12

        .Cells(startRow + 2, 1).Value = "• Nombre total d'enregistrements :"
        .Cells(startRow + 2, 3).Value = Format(totalRows, "#,##0")
        .Cells(startRow + 2, 3).Font.Bold = True

        .Cells(startRow + 3, 1).Value = "• Nombre de variables :"
        .Cells(startRow + 3, 3).Value = UBound(headers)
        .Cells(startRow + 3, 3).Font.Bold = True

        .Cells(startRow + 4, 1).Value = "• Variables numériques :"
        .Cells(startRow + 4, 3).Value = numericColumns
        .Cells(startRow + 4, 3).Font.Bold = True

        .Cells(startRow + 5, 1).Value = "• Variables textuelles :"
        .Cells(startRow + 5, 3).Value = textColumns
        .Cells(startRow + 5, 3).Font.Bold = True

        .Cells(startRow + 6, 1).Value = "• Variables temporelles :"
        .Cells(startRow + 6, 3).Value = dateColumns
        .Cells(startRow + 6, 3).Font.Bold = True

        ' Analyse de complétude
        .Cells(startRow + 8, 1).Value = "📋 QUALITÉ DES DONNÉES"
        .Cells(startRow + 8, 1).Font.Bold = True
        .Cells(startRow + 8, 1).Font.Size = 12

        Dim completenessRate As Double
        Dim totalCells As Long
        Dim emptyCells As Long

        totalCells = (sourceRange.Rows.Count - 1) * sourceRange.Columns.Count  ' Exclure en-têtes
        emptyCells = Application.CountBlank(sourceRange.Offset(1, 0).Resize(sourceRange.Rows.Count - 1))
        completenessRate = (totalCells - emptyCells) / totalCells * 100

        .Cells(startRow + 10, 1).Value = "• Taux de complétude :"
        .Cells(startRow + 10, 3).Value = Format(completenessRate, "0.0") & "%"
        .Cells(startRow + 10, 3).Font.Bold = True

        ' Colorer selon le taux de complétude
        If completenessRate >= 95 Then
            .Cells(startRow + 10, 3).Font.Color = RGB(0, 128, 0)  ' Vert
        ElseIf completenessRate >= 80 Then
            .Cells(startRow + 10, 3).Font.Color = RGB(255, 165, 0)  ' Orange
        Else
            .Cells(startRow + 10, 3).Font.Color = RGB(255, 0, 0)  ' Rouge
        End If

        .Cells(startRow + 11, 1).Value = "• Cellules vides :"
        .Cells(startRow + 11, 3).Value = Format(emptyCells, "#,##0")
        .Cells(startRow + 11, 3).Font.Bold = True

        ' Principales observations
        .Cells(startRow + 13, 1).Value = "🔍 PRINCIPALES OBSERVATIONS"
        .Cells(startRow + 13, 1).Font.Bold = True
        .Cells(startRow + 13, 1).Font.Size = 12

        Call GenerateKeyInsights(sourceRange, analysisSheet, startRow + 15, headers, columnTypes)
    End With
End Sub

Sub GenerateKeyInsights(sourceRange As Range, analysisSheet As Worksheet, startRow As Long, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Génération d'observations clés automatiques
    '=========================================

    Dim col As Long
    Dim insightRow As Long
    Dim maxValue As Double
    Dim minValue As Double
    Dim avgValue As Double
    Dim maxColumn As String
    Dim minColumn As String

    insightRow = startRow

    ' Trouver la colonne numérique avec les plus grandes valeurs
    For col = 1 To UBound(columnTypes)
        If columnTypes(col) = "NUMERIC" Then
            Dim colRange As Range
            Set colRange = sourceRange.Columns(col).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

            Dim colMax As Double
            Dim colMin As Double
            Dim colAvg As Double

            On Error Resume Next
            colMax = Application.Max(colRange)
            colMin = Application.Min(colRange)
            colAvg = Application.Average(colRange)
            On Error GoTo 0

            If col = 1 Or colMax > maxValue Then
                maxValue = colMax
                maxColumn = headers(col)
            End If

            If col = 1 Or colAvg > avgValue Then
                avgValue = colAvg
            End If
        End If
    Next col

    With analysisSheet
        ' Observation sur les valeurs maximales
        If maxColumn <> "" Then
            .Cells(insightRow, 1).Value = "• La variable '" & maxColumn & "' présente la valeur maximale de " & Format(maxValue, "#,##0.00")
            insightRow = insightRow + 1
        End If

        ' Observation sur la période d'analyse (si données temporelles)
        Dim dateCol As Long
        For col = 1 To UBound(columnTypes)
            If columnTypes(col) = "DATE" Then
                dateCol = col
                Exit For
            End If
        Next col

        If dateCol > 0 Then
            Dim dateRange As Range
            Set dateRange = sourceRange.Columns(dateCol).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

            On Error Resume Next
            Dim earliestDate As Date
            Dim latestDate As Date
            earliestDate = Application.Min(dateRange)
            latestDate = Application.Max(dateRange)
            On Error GoTo 0

            If earliestDate <> 0 And latestDate <> 0 Then
                .Cells(insightRow, 1).Value = "• Période d'analyse : du " & Format(earliestDate, "dd/mm/yyyy") & " au " & Format(latestDate, "dd/mm/yyyy")
                insightRow = insightRow + 1

                Dim daysDiff As Long
                daysDiff = latestDate - earliestDate
                .Cells(insightRow, 1).Value = "• Durée couverte : " & daysDiff & " jour(s)"
                insightRow = insightRow + 1
            End If
        End If

        ' Observation sur la variabilité
        .Cells(insightRow, 1).Value = "• " & (sourceRange.Rows.Count - 1) & " enregistrements analysés sur " & UBound(headers) & " variables"
        insightRow = insightRow + 1

        ' Recommandation générale
        .Cells(insightRow + 1, 1).Value = "💡 Ces observations préliminaires seront détaillées dans les sections suivantes."
        .Cells(insightRow + 1, 1).Font.Italic = True
        .Cells(insightRow + 1, 1).Font.Color = RGB(0, 102, 204)
    End With
End Sub
```

### 4. Module de statistiques descriptives

```vba
Sub CalculateDescriptiveStats(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Calcul complet des statistiques descriptives
    '=========================================

    Dim startRow As Long
    Dim col As Long
    Dim statsRow As Long

    startRow = 16  ' Position après le titre de section
    statsRow = startRow + 1

    With analysisSheet
        ' En-têtes du tableau de statistiques
        .Cells(statsRow, 1).Value = "Variable"
        .Cells(statsRow, 2).Value = "Type"
        .Cells(statsRow, 3).Value = "Observations"
        .Cells(statsRow, 4).Value = "Moyenne"
        .Cells(statsRow, 5).Value = "Médiane"
        .Cells(statsRow, 6).Value = "Écart-type"
        .Cells(statsRow, 7).Value = "Minimum"
        .Cells(statsRow, 8).Value = "Maximum"
        .Cells(statsRow, 9).Value = "Q1"
        .Cells(statsRow, 10).Value = "Q3"
        .Cells(statsRow, 11).Value = "Valeurs uniques"
        .Cells(statsRow, 12).Value = "Données manquantes"

        ' Formatage des en-têtes
        With .Range(.Cells(statsRow, 1), .Cells(statsRow, 12))
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(68, 114, 196)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    End With

    ' Calculer les statistiques pour chaque colonne
    For col = 1 To UBound(headers)
        Call CalculateColumnStats(sourceRange, analysisSheet, col, statsRow + col, headers(col), columnTypes(col))
    Next col

    ' Ajouter un résumé des corrélations si plusieurs colonnes numériques
    Call AddCorrelationSummary(sourceRange, analysisSheet, statsRow + UBound(headers) + 3, headers, columnTypes)
End Sub

Sub CalculateColumnStats(sourceRange As Range, analysisSheet As Worksheet, colIndex As Long, targetRow As Long, headerName As String, dataType As String)
    '=========================================
    ' Calcul des statistiques pour une colonne spécifique
    '=========================================

    Dim dataRange As Range
    Dim observations As Long
    Dim missingData As Long
    Dim uniqueValues As Long
    Dim mean As Double
    Dim median As Double
    Dim stdDev As Double
    Dim minVal As Variant
    Dim maxVal As Variant
    Dim q1 As Double
    Dim q3 As Double

    ' Définir la plage de données (exclure l'en-tête)
    Set dataRange = sourceRange.Columns(colIndex).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

    ' Calculs de base
    observations = dataRange.Rows.Count
    missingData = Application.CountBlank(dataRange)

    ' Statistiques selon le type de données
    With analysisSheet
        .Cells(targetRow, 1).Value = headerName
        .Cells(targetRow, 2).Value = dataType
        .Cells(targetRow, 3).Value = observations
        .Cells(targetRow, 12).Value = missingData

        Select Case dataType
            Case "NUMERIC"
                Call CalculateNumericStats(dataRange, mean, median, stdDev, minVal, maxVal, q1, q3, uniqueValues)

                .Cells(targetRow, 4).Value = Round(mean, 2)
                .Cells(targetRow, 5).Value = Round(median, 2)
                .Cells(targetRow, 6).Value = Round(stdDev, 2)
                .Cells(targetRow, 7).Value = minVal
                .Cells(targetRow, 8).Value = maxVal
                .Cells(targetRow, 9).Value = Round(q1, 2)
                .Cells(targetRow, 10).Value = Round(q3, 2)
                .Cells(targetRow, 11).Value = uniqueValues

                ' Formatage numérique
                .Range(.Cells(targetRow, 4), .Cells(targetRow, 10)).NumberFormat = "#,##0.00"

            Case "DATE"
                Call CalculateDateStats(dataRange, minVal, maxVal, uniqueValues)

                .Cells(targetRow, 4).Value = "N/A"
                .Cells(targetRow, 5).Value = "N/A"
                .Cells(targetRow, 6).Value = "N/A"
                .Cells(targetRow, 7).Value = minVal
                .Cells(targetRow, 8).Value = maxVal
                .Cells(targetRow, 9).Value = "N/A"
                .Cells(targetRow, 10).Value = "N/A"
                .Cells(targetRow, 11).Value = uniqueValues

                ' Formatage des dates
                If minVal <> "N/A" Then .Cells(targetRow, 7).NumberFormat = "dd/mm/yyyy"
                If maxVal <> "N/A" Then .Cells(targetRow, 8).NumberFormat = "dd/mm/yyyy"

            Case "TEXT"
                Call CalculateTextStats(dataRange, uniqueValues)

                .Cells(targetRow, 4).Value = "N/A"
                .Cells(targetRow, 5).Value = "N/A"
                .Cells(targetRow, 6).Value = "N/A"
                .Cells(targetRow, 7).Value = "N/A"
                .Cells(targetRow, 8).Value = "N/A"
                .Cells(targetRow, 9).Value = "N/A"
                .Cells(targetRow, 10).Value = "N/A"
                .Cells(targetRow, 11).Value = uniqueValues
        End Select

        ' Formatage conditionnel selon les données manquantes
        If missingData > 0 Then
            Dim missingRate As Double
            missingRate = missingData / observations

            If missingRate > 0.1 Then  ' Plus de 10% manquant
                .Cells(targetRow, 12).Font.Color = RGB(255, 0, 0)  ' Rouge
            ElseIf missingRate > 0.05 Then  ' Plus de 5% manquant
                .Cells(targetRow, 12).Font.Color = RGB(255, 165, 0)  ' Orange
            End If
        End If

        ' Bordures pour toute la ligne
        .Range(.Cells(targetRow, 1), .Cells(targetRow, 12)).Borders.LineStyle = xlContinuous

        ' Coloration alternée
        If targetRow Mod 2 = 0 Then
            .Range(.Cells(targetRow, 1), .Cells(targetRow, 12)).Interior.Color = RGB(242, 242, 242)
        End If
    End With
End Sub

Sub CalculateNumericStats(dataRange As Range, ByRef mean As Double, ByRef median As Double, ByRef stdDev As Double, ByRef minVal As Variant, ByRef maxVal As Variant, ByRef q1 As Double, ByRef q3 As Double, ByRef uniqueValues As Long)
    '=========================================
    ' Calculs statistiques pour données numériques
    '=========================================

    On Error Resume Next

    ' Statistiques de base
    mean = Application.Average(dataRange)
    median = Application.Median(dataRange)
    stdDev = Application.StDev(dataRange)
    minVal = Application.Min(dataRange)
    maxVal = Application.Max(dataRange)

    ' Quartiles
    q1 = Application.Quartile(dataRange, 1)
    q3 = Application.Quartile(dataRange, 3)

    ' Compter les valeurs uniques
    uniqueValues = CountUniqueValues(dataRange)

    On Error GoTo 0

    ' Gestion des erreurs (si toutes les cellules sont vides)
    If IsError(mean) Then mean = 0
    If IsError(median) Then median = 0
    If IsError(stdDev) Then stdDev = 0
    If IsError(minVal) Then minVal = "N/A"
    If IsError(maxVal) Then maxVal = "N/A"
    If IsError(q1) Then q1 = 0
    If IsError(q3) Then q3 = 0
End Sub

Sub CalculateDateStats(dataRange As Range, ByRef minVal As Variant, ByRef maxVal As Variant, ByRef uniqueValues As Long)
    '=========================================
    ' Calculs statistiques pour données temporelles
    '=========================================

    On Error Resume Next

    minVal = Application.Min(dataRange)
    maxVal = Application.Max(dataRange)
    uniqueValues = CountUniqueValues(dataRange)

    On Error GoTo 0

    If IsError(minVal) Or minVal = 0 Then minVal = "N/A"
    If IsError(maxVal) Or maxVal = 0 Then maxVal = "N/A"
End Sub

Sub CalculateTextStats(dataRange As Range, ByRef uniqueValues As Long)
    '=========================================
    ' Calculs statistiques pour données textuelles
    '=========================================

    uniqueValues = CountUniqueValues(dataRange)
End Sub

Function CountUniqueValues(dataRange As Range) As Long
    '=========================================
    ' Comptage des valeurs uniques dans une plage
    '=========================================

    Dim dict As Object
    Dim cell As Range
    Dim cellValue As Variant

    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In dataRange
        cellValue = cell.Value
        If cellValue <> "" Then
            If Not dict.Exists(cellValue) Then
                dict.Add cellValue, 1
            End If
        End If
    Next cell

    CountUniqueValues = dict.Count
    Set dict = Nothing
End Function

Sub AddCorrelationSummary(sourceRange As Range, analysisSheet As Worksheet, startRow As Long, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Ajout d'un résumé des corrélations entre variables numériques
    '=========================================

    Dim numericCols() As Long
    Dim numericCount As Long
    Dim col As Long
    Dim i As Long, j As Long

    ' Identifier les colonnes numériques
    numericCount = 0
    For col = 1 To UBound(columnTypes)
        If columnTypes(col) = "NUMERIC" Then
            numericCount = numericCount + 1
            ReDim Preserve numericCols(1 To numericCount)
            numericCols(numericCount) = col
        End If
    Next col

    ' Ne créer la matrice de corrélation que s'il y a au moins 2 variables numériques
    If numericCount >= 2 Then
        With analysisSheet
            .Cells(startRow, 1).Value = "📊 MATRICE DE CORRÉLATION"
            .Cells(startRow, 1).Font.Bold = True
            .Cells(startRow, 1).Font.Size = 12

            ' En-têtes de la matrice
            .Cells(startRow + 2, 1).Value = "Variables"
            For i = 1 To numericCount
                .Cells(startRow + 2, i + 1).Value = headers(numericCols(i))
                .Cells(startRow + 2 + i, 1).Value = headers(numericCols(i))
            Next i

            ' Formatage des en-têtes
            .Range(.Cells(startRow + 2, 1), .Cells(startRow + 2, numericCount + 1)).Font.Bold = True
            .Range(.Cells(startRow + 3, 1), .Cells(startRow + 2 + numericCount, 1)).Font.Bold = True

            ' Calcul des corrélations
            For i = 1 To numericCount
                For j = 1 To numericCount
                    Dim correlation As Double
                    correlation = CalculateCorrelation(sourceRange, numericCols(i), numericCols(j))

                    .Cells(startRow + 2 + i, j + 1).Value = Round(correlation, 3)

                    ' Formatage conditionnel des corrélations
                    If Abs(correlation) >= 0.7 And i <> j Then
                        .Cells(startRow + 2 + i, j + 1).Font.Color = RGB(255, 0, 0)  ' Rouge pour forte corrélation
                        .Cells(startRow + 2 + i, j + 1).Font.Bold = True
                    ElseIf Abs(correlation) >= 0.4 And i <> j Then
                        .Cells(startRow + 2 + i, j + 1).Font.Color = RGB(255, 165, 0)  ' Orange pour corrélation modérée
                    End If

                    ' Diagonale en gras (corrélation = 1)
                    If i = j Then
                        .Cells(startRow + 2 + i, j + 1).Font.Bold = True
                    End If
                Next j
            Next i

            ' Bordures pour la matrice
            .Range(.Cells(startRow + 2, 1), .Cells(startRow + 2 + numericCount, numericCount + 1)).Borders.LineStyle = xlContinuous

            ' Légende
            .Cells(startRow + numericCount + 5, 1).Value = "Légende : |r| ≥ 0.7 = forte corrélation (rouge), |r| ≥ 0.4 = corrélation modérée (orange)"
            .Cells(startRow + numericCount + 5, 1).Font.Italic = True
            .Cells(startRow + numericCount + 5, 1).Font.Size = 10
        End With
    End If
End Sub

Function CalculateCorrelation(sourceRange As Range, col1 As Long, col2 As Long) As Double
    '=========================================
    ' Calcul du coefficient de corrélation entre deux colonnes
    '=========================================

    Dim range1 As Range
    Dim range2 As Range

    Set range1 = sourceRange.Columns(col1).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)
    Set range2 = sourceRange.Columns(col2).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

    On Error Resume Next
    CalculateCorrelation = Application.Correl(range1, range2)
    On Error GoTo 0

    If IsError(CalculateCorrelation) Then CalculateCorrelation = 0
End Function
```

## 5. Module de visualisations automatiques

```vba
Sub CreateVisualizations(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Création automatique de graphiques adaptés aux données
    '=========================================

    Dim startRow As Long
    Dim chartCount As Long
    Dim col As Long

    startRow = 36  ' Position après le titre de section
    chartCount = 0

    With analysisSheet
        .Cells(startRow, 1).Value = "Les graphiques suivants sont générés automatiquement selon le type de données :"
        .Cells(startRow, 1).Font.Italic = True
    End With

    ' Créer des graphiques pour chaque variable numérique
    For col = 1 To UBound(columnTypes)
        If columnTypes(col) = "NUMERIC" Then
            Call CreateHistogram(sourceRange, analysisSheet, col, headers(col), startRow + 2 + (chartCount * 20))
            chartCount = chartCount + 1
        End If
    Next col

    ' Créer un graphique de distribution pour les données textuelles
    Call CreateCategoryDistribution(sourceRange, analysisSheet, headers, columnTypes, startRow + 2 + (chartCount * 20))
    chartCount = chartCount + 1

    ' Créer un graphique temporel si données de dates
    Call CreateTimeSeriesChart(sourceRange, analysisSheet, headers, columnTypes, startRow + 2 + (chartCount * 20))

    ' Créer un graphique de corrélation si plusieurs variables numériques
    Call CreateCorrelationChart(sourceRange, analysisSheet, headers, columnTypes, startRow + 2 + ((chartCount + 1) * 20))
End Sub

Sub CreateHistogram(sourceRange As Range, analysisSheet As Worksheet, colIndex As Long, columnName As String, chartTop As Long)
    '=========================================
    ' Création d'un histogramme pour une variable numérique
    '=========================================

    Dim dataRange As Range
    Dim chartObject As ChartObject
    Dim chart As Chart
    Dim binCount As Long
    Dim minVal As Double
    Dim maxVal As Double
    Dim binWidth As Double
    Dim bins() As Double
    Dim frequencies() As Long
    Dim i As Long

    Set dataRange = sourceRange.Columns(colIndex).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

    ' Calculer les paramètres de l'histogramme
    minVal = Application.Min(dataRange)
    maxVal = Application.Max(dataRange)
    binCount = Application.Min(10, Int(Sqr(dataRange.Rows.Count)))  ' Règle de Sturges simplifiée
    binWidth = (maxVal - minVal) / binCount

    ' Créer les intervalles
    ReDim bins(0 To binCount)
    ReDim frequencies(0 To binCount - 1)

    For i = 0 To binCount
        bins(i) = minVal + i * binWidth
    Next i

    ' Calculer les fréquences
    Call CalculateFrequencies(dataRange, bins, frequencies)

    ' Créer les données dans la feuille pour le graphique
    Dim dataStartRow As Long
    dataStartRow = chartTop + 15

    With analysisSheet
        .Cells(dataStartRow, 1).Value = "Intervalle"
        .Cells(dataStartRow, 2).Value = "Fréquence"

        For i = 0 To UBound(frequencies)
            .Cells(dataStartRow + 1 + i, 1).Value = Format(bins(i), "0.00") & " - " & Format(bins(i + 1), "0.00")
            .Cells(dataStartRow + 1 + i, 2).Value = frequencies(i)
        Next i

        ' Créer le graphique
        Set chartObject = .ChartObjects.Add(Left:=50, Top:=chartTop, Width:=400, Height:=250)
        Set chart = chartObject.Chart

        With chart
            .ChartType = xlColumnClustered
            .SetSourceData .Parent.Parent.Range(.Parent.Parent.Cells(dataStartRow, 1), .Parent.Parent.Cells(dataStartRow + UBound(frequencies) + 1, 2))
            .HasTitle = True
            .ChartTitle.Text = "Distribution de " & columnName
            .HasLegend = False

            ' Formatage des axes
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Intervalles"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Fréquence"

            ' Couleur des barres
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
        End With

        ' Titre descriptif
        .Cells(chartTop - 2, 1).Value = "📊 HISTOGRAMME - " & UCase(columnName)
        .Cells(chartTop - 2, 1).Font.Bold = True
        .Cells(chartTop - 2, 1).Font.Size = 12
    End With
End Sub

Sub CalculateFrequencies(dataRange As Range, bins() As Double, ByRef frequencies() As Long)
    '=========================================
    ' Calcul des fréquences pour l'histogramme
    '=========================================

    Dim cell As Range
    Dim value As Double
    Dim binIndex As Long

    For Each cell In dataRange
        If IsNumeric(cell.Value) And cell.Value <> "" Then
            value = cell.Value

            ' Trouver le bon intervalle
            For binIndex = 0 To UBound(bins) - 1
                If value >= bins(binIndex) And value < bins(binIndex + 1) Then
                    frequencies(binIndex) = frequencies(binIndex) + 1
                    Exit For
                ElseIf binIndex = UBound(bins) - 1 And value = bins(binIndex + 1) Then
                    ' Cas spécial pour la valeur maximale
                    frequencies(binIndex) = frequencies(binIndex) + 1
                    Exit For
                End If
            Next binIndex
        End If
    Next cell
End Sub

Sub CreateCategoryDistribution(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant, chartTop As Long)
    '=========================================
    ' Création d'un graphique de distribution des catégories
    '=========================================

    Dim textCol As Long
    Dim dataRange As Range
    Dim categories As Object
    Dim chartObject As ChartObject
    Dim chart As Chart
    Dim cell As Range
    Dim dataStartRow As Long
    Dim i As Long

    ' Trouver la première colonne textuelle
    For textCol = 1 To UBound(columnTypes)
        If columnTypes(textCol) = "TEXT" Then
            Exit For
        End If
    Next textCol

    If textCol > UBound(columnTypes) Then Exit Sub  ' Aucune colonne textuelle

    Set dataRange = sourceRange.Columns(textCol).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)
    Set categories = CreateObject("Scripting.Dictionary")

    ' Compter les occurrences de chaque catégorie
    For Each cell In dataRange
        If cell.Value <> "" Then
            If categories.Exists(cell.Value) Then
                categories(cell.Value) = categories(cell.Value) + 1
            Else
                categories.Add cell.Value, 1
            End If
        End If
    Next cell

    If categories.Count = 0 Then Exit Sub

    ' Créer les données dans la feuille
    dataStartRow = chartTop + 15

    With analysisSheet
        .Cells(dataStartRow, 1).Value = headers(textCol)
        .Cells(dataStartRow, 2).Value = "Nombre"

        i = 1
        Dim key As Variant
        For Each key In categories.Keys
            .Cells(dataStartRow + i, 1).Value = key
            .Cells(dataStartRow + i, 2).Value = categories(key)
            i = i + 1
        Next key

        ' Créer le graphique en secteurs
        Set chartObject = .ChartObjects.Add(Left:=500, Top:=chartTop, Width:=400, Height:=250)
        Set chart = chartObject.Chart

        With chart
            .ChartType = xlPie
            .SetSourceData .Parent.Parent.Range(.Parent.Parent.Cells(dataStartRow, 1), .Parent.Parent.Cells(dataStartRow + categories.Count, 2))
            .HasTitle = True
            .ChartTitle.Text = "Répartition par " & headers(textCol)
            .HasLegend = True
            .Legend.Position = xlLegendPositionRight

            ' Afficher les pourcentages
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.ShowPercentage = True
        End With

        ' Titre descriptif
        .Cells(chartTop - 2, 7).Value = "🥧 RÉPARTITION - " & UCase(headers(textCol))
        .Cells(chartTop - 2, 7).Font.Bold = True
        .Cells(chartTop - 2, 7).Font.Size = 12
    End With

    Set categories = Nothing
End Sub

Sub CreateTimeSeriesChart(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant, chartTop As Long)
    '=========================================
    ' Création d'un graphique temporel si données de dates
    '=========================================

    Dim dateCol As Long
    Dim numCol As Long
    Dim chartObject As ChartObject
    Dim chart As Chart
    Dim dateRange As Range
    Dim numRange As Range

    ' Trouver une colonne de dates et une colonne numérique
    dateCol = 0
    numCol = 0

    For dateCol = 1 To UBound(columnTypes)
        If columnTypes(dateCol) = "DATE" Then
            Exit For
        End If
    Next dateCol

    For numCol = 1 To UBound(columnTypes)
        If columnTypes(numCol) = "NUMERIC" Then
            Exit For
        End If
    Next numCol

    If dateCol > UBound(columnTypes) Or numCol > UBound(columnTypes) Then Exit Sub

    Set dateRange = sourceRange.Columns(dateCol).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)
    Set numRange = sourceRange.Columns(numCol).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

    ' Créer le graphique temporel
    With analysisSheet
        Set chartObject = .ChartObjects.Add(Left:=50, Top:=chartTop, Width:=600, Height:=300)
        Set chart = chartObject.Chart

        With chart
            .ChartType = xlLine
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = dateRange
            .SeriesCollection(1).Values = numRange
            .SeriesCollection(1).Name = headers(numCol)
            .HasTitle = True
            .ChartTitle.Text = "Évolution de " & headers(numCol) & " dans le temps"
            .HasLegend = False

            ' Formatage des axes
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = headers(dateCol)
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = headers(numCol)

            ' Formatage de la ligne
            .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .SeriesCollection(1).Format.Line.Weight = 2
            .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(1).MarkerSize = 5
        End With

        ' Titre descriptif
        .Cells(chartTop - 2, 1).Value = "📈 ÉVOLUTION TEMPORELLE"
        .Cells(chartTop - 2, 1).Font.Bold = True
        .Cells(chartTop - 2, 1).Font.Size = 12
    End With
End Sub

Sub CreateCorrelationChart(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant, chartTop As Long)
    '=========================================
    ' Création d'un graphique de corrélation (nuage de points)
    '=========================================

    Dim numericCols() As Long
    Dim numericCount As Long
    Dim col As Long
    Dim chartObject As ChartObject
    Dim chart As Chart

    ' Identifier les colonnes numériques
    numericCount = 0
    For col = 1 To UBound(columnTypes)
        If columnTypes(col) = "NUMERIC" Then
            numericCount = numericCount + 1
            ReDim Preserve numericCols(1 To numericCount)
            numericCols(numericCount) = col
        End If
    Next col

    If numericCount < 2 Then Exit Sub  ' Besoin d'au moins 2 variables numériques

    ' Utiliser les deux premières variables numériques
    Dim range1 As Range
    Dim range2 As Range

    Set range1 = sourceRange.Columns(numericCols(1)).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)
    Set range2 = sourceRange.Columns(numericCols(2)).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

    With analysisSheet
        Set chartObject = .ChartObjects.Add(Left:=700, Top:=chartTop, Width:=400, Height:=300)
        Set chart = chartObject.Chart

        With chart
            .ChartType = xlXYScatter
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = range1
            .SeriesCollection(1).Values = range2
            .SeriesCollection(1).Name = headers(numericCols(2)) & " vs " & headers(numericCols(1))
            .HasTitle = True
            .ChartTitle.Text = "Corrélation : " & headers(numericCols(1)) & " vs " & headers(numericCols(2))
            .HasLegend = False

            ' Formatage des axes
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = headers(numericCols(1))
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = headers(numericCols(2))

            ' Formatage des points
            .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(1).MarkerSize = 6
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)

            ' Ajouter une ligne de tendance si corrélation significative
            Dim correlation As Double
            correlation = CalculateCorrelation(sourceRange, numericCols(1), numericCols(2))

            If Abs(correlation) > 0.3 Then  ' Corrélation modérée ou forte
                .SeriesCollection(1).Trendlines.Add
                .SeriesCollection(1).Trendlines(1).Type = xlLinear
                .SeriesCollection(1).Trendlines(1).DisplayEquation = True
                .SeriesCollection(1).Trendlines(1).DisplayRSquared = True
            End If
        End With

        ' Titre descriptif avec coefficient de corrélation
        .Cells(chartTop - 2, 10).Value = "🔗 CORRÉLATION (r = " & Format(correlation, "0.000") & ")"
        .Cells(chartTop - 2, 10).Font.Bold = True
        .Cells(chartTop - 2, 10).Font.Size = 12

        ' Interprétation de la corrélation
        Dim interpretation As String
        If Abs(correlation) >= 0.8 Then
            interpretation = "Corrélation très forte"
        ElseIf Abs(correlation) >= 0.6 Then
            interpretation = "Corrélation forte"
        ElseIf Abs(correlation) >= 0.4 Then
            interpretation = "Corrélation modérée"
        ElseIf Abs(correlation) >= 0.2 Then
            interpretation = "Corrélation faible"
        Else
            interpretation = "Corrélation négligeable"
        End If

        .Cells(chartTop - 1, 10).Value = interpretation
        .Cells(chartTop - 1, 10).Font.Italic = True
    End With
End Sub
```

## 6. Module de détection d'anomalies

```vba
Sub DetectAnomalies(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Détection automatique d'anomalies dans les données
    '=========================================

    Dim startRow As Long
    Dim col As Long
    Dim anomalyRow As Long
    Dim totalAnomalies As Long

    startRow = 56  ' Position après le titre de section
    anomalyRow = startRow + 2
    totalAnomalies = 0

    With analysisSheet
        .Cells(startRow, 1).Value = "Cette section identifie automatiquement les valeurs suspectes ou incohérentes :"
        .Cells(startRow, 1).Font.Italic = True

        ' En-têtes du tableau d'anomalies
        .Cells(anomalyRow, 1).Value = "Type d'Anomalie"
        .Cells(anomalyRow, 2).Value = "Variable"
        .Cells(anomalyRow, 3).Value = "Description"
        .Cells(anomalyRow, 4).Value = "Valeur(s)"
        .Cells(anomalyRow, 5).Value = "Ligne(s)"
        .Cells(anomalyRow, 6).Value = "Sévérité"

        ' Formatage des en-têtes
        With .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6))
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(255, 0, 0)  ' Rouge pour les anomalies
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With

        anomalyRow = anomalyRow + 1
    End With

    ' Détecter les anomalies pour chaque colonne numérique
    For col = 1 To UBound(columnTypes)
        If columnTypes(col) = "NUMERIC" Then
            Call DetectOutliers(sourceRange, analysisSheet, col, headers(col), anomalyRow, totalAnomalies)
            Call DetectNegativeValues(sourceRange, analysisSheet, col, headers(col), anomalyRow, totalAnomalies)
        End If
    Next col

    ' Détecter les valeurs manquantes importantes
    Call DetectMissingData(sourceRange, analysisSheet, headers, columnTypes, anomalyRow, totalAnomalies)

    ' Détecter les doublons
    Call DetectDuplicates(sourceRange, analysisSheet, headers, anomalyRow, totalAnomalies)

    ' Détecter les incohérences temporelles
    Call DetectDateInconsistencies(sourceRange, analysisSheet, headers, columnTypes, anomalyRow, totalAnomalies)

    ' Résumé des anomalies
    Call SummarizeAnomalies(analysisSheet, startRow, totalAnomalies)
End Sub

Sub DetectOutliers(sourceRange As Range, analysisSheet As Worksheet, colIndex As Long, columnName As String, ByRef anomalyRow As Long, ByRef totalAnomalies As Long)
    '=========================================
    ' Détection des valeurs aberrantes (méthode IQR)
    '=========================================

    Dim dataRange As Range
    Dim q1 As Double, q3 As Double, iqr As Double
    Dim lowerBound As Double, upperBound As Double
    Dim cell As Range
    Dim outliers As String
    Dim outlierRows As String
    Dim outlierCount As Long

    Set dataRange = sourceRange.Columns(colIndex).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

    On Error Resume Next
    q1 = Application.Quartile(dataRange, 1)
    q3 = Application.Quartile(dataRange, 3)
    On Error GoTo 0

    If q1 = 0 And q3 = 0 Then Exit Sub  ' Pas assez de données

    iqr = q3 - q1
    lowerBound = q1 - 1.5 * iqr
    upperBound = q3 + 1.5 * iqr

    outlierCount = 0
    outliers = ""
    outlierRows = ""

    ' Parcourir les données pour identifier les outliers
    Dim rowIndex As Long
    For rowIndex = 1 To dataRange.Rows.Count
        If IsNumeric(dataRange.Cells(rowIndex, 1).Value) And dataRange.Cells(rowIndex, 1).Value <> "" Then
            Dim value As Double
            value = dataRange.Cells(rowIndex, 1).Value

            If value < lowerBound Or value > upperBound Then
                outlierCount = outlierCount + 1
                If outliers <> "" Then outliers = outliers & ", "
                If outlierRows <> "" Then outlierRows = outlierRows & ", "
                outliers = outliers & Format(value, "#,##0.00")
                outlierRows = outlierRows & (rowIndex + 1)  ' +1 car on exclut l'en-tête

                ' Limiter l'affichage à 5 valeurs
                If outlierCount >= 5 Then
                    outliers = outliers & "..."
                    outlierRows = outlierRows & "..."
                    Exit For
                End If
            End If
        End If
    Next rowIndex

    ' Ajouter à la liste des anomalies si outliers trouvés
    If outlierCount > 0 Then
        With analysisSheet
            .Cells(anomalyRow, 1).Value = "Valeurs aberrantes"
            .Cells(anomalyRow, 2).Value = columnName
            .Cells(anomalyRow, 3).Value = outlierCount & " valeur(s) hors limites [" & Format(lowerBound, "#,##0.00") & " ; " & Format(upperBound, "#,##0.00") & "]"
            .Cells(anomalyRow, 4).Value = outliers
            .Cells(anomalyRow, 5).Value = outlierRows
            .Cells(anomalyRow, 6).Value = IIf(outlierCount > 5, "ÉLEVÉE", IIf(outlierCount > 2, "MODÉRÉE", "FAIBLE"))

            ' Formatage selon la sévérité
            If outlierCount > 5 Then
                .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 200, 200)  ' Rouge clair
            ElseIf outlierCount > 2 Then
                .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 230, 200)  ' Orange clair
            Else
                .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 255, 200)  ' Jaune clair
            End If

            .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Borders.LineStyle = xlContinuous
            anomalyRow = anomalyRow + 1
            totalAnomalies = totalAnomalies + 1
        End With
    End If
End Sub

Sub DetectNegativeValues(sourceRange As Range, analysisSheet As Worksheet, colIndex As Long, columnName As String, ByRef anomalyRow As Long, ByRef totalAnomalies As Long)
    '=========================================
    ' Détection des valeurs négatives (selon contexte)
    '=========================================

    Dim dataRange As Range
    Dim cell As Range
    Dim negativeCount As Long
    Dim negativeValues As String
    Dim negativeRows As String

    Set dataRange = sourceRange.Columns(colIndex).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)

    ' Cette détection est contextuelle - on suppose que des valeurs comme "quantité", "prix", "age" ne devraient pas être négatives
    ' Pour simplifier, on détecte toutes les valeurs négatives et laisse l'utilisateur juger

    negativeCount = 0
    negativeValues = ""
    negativeRows = ""

    Dim rowIndex As Long
    For rowIndex = 1 To dataRange.Rows.Count
        If IsNumeric(dataRange.Cells(rowIndex, 1).Value) And dataRange.Cells(rowIndex, 1).Value <> "" Then
            Dim value As Double
            value = dataRange.Cells(rowIndex, 1).Value

            If value < 0 Then
                negativeCount = negativeCount + 1
                If negativeValues <> "" Then negativeValues = negativeValues & ", "
                If negativeRows <> "" Then negativeRows = negativeRows & ", "
                negativeValues = negativeValues & Format(value, "#,##0.00")
                negativeRows = negativeRows & (rowIndex + 1)

                If negativeCount >= 5 Then
                    negativeValues = negativeValues & "..."
                    negativeRows = negativeRows & "..."
                    Exit For
                End If
            End If
        End If
    Next rowIndex

    ' Ajouter à la liste si valeurs négatives trouvées
    If negativeCount > 0 Then
        With analysisSheet
            .Cells(anomalyRow, 1).Value = "Valeurs négatives"
            .Cells(anomalyRow, 2).Value = columnName
            .Cells(anomalyRow, 3).Value = negativeCount & " valeur(s) négative(s) détectée(s) - Vérifier la cohérence"
            .Cells(anomalyRow, 4).Value = negativeValues
            .Cells(anomalyRow, 5).Value = negativeRows
            .Cells(anomalyRow, 6).Value = "MODÉRÉE"

            .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 230, 200)  ' Orange clair
            .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Borders.LineStyle = xlContinuous

            anomalyRow = anomalyRow + 1
            totalAnomalies = totalAnomalies + 1
        End With
    End If
End Sub

Sub DetectMissingData(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant, ByRef anomalyRow As Long, ByRef totalAnomalies As Long)
    '=========================================
    ' Détection des colonnes avec beaucoup de données manquantes
    '=========================================

    Dim col As Long
    Dim dataRange As Range
    Dim missingCount As Long
    Dim totalCount As Long
    Dim missingRate As Double

    totalCount = sourceRange.Rows.Count - 1  ' Exclure l'en-tête

    For col = 1 To UBound(headers)
        Set dataRange = sourceRange.Columns(col).Offset(1, 0).Resize(totalCount)
        missingCount = Application.CountBlank(dataRange)
        missingRate = missingCount / totalCount

        ' Signaler si plus de 20% de données manquantes
        If missingRate > 0.2 Then
            With analysisSheet
                .Cells(anomalyRow, 1).Value = "Données manquantes"
                .Cells(anomalyRow, 2).Value = headers(col)
                .Cells(anomalyRow, 3).Value = Format(missingRate * 100, "0.0") & "% de données manquantes (" & missingCount & "/" & totalCount & ")"
                .Cells(anomalyRow, 4).Value = "N/A"
                .Cells(anomalyRow, 5).Value = "Multiple"
                .Cells(anomalyRow, 6).Value = IIf(missingRate > 0.5, "ÉLEVÉE", IIf(missingRate > 0.3, "MODÉRÉE", "FAIBLE"))

                ' Formatage selon le taux de données manquantes
                If missingRate > 0.5 Then
                    .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 200, 200)  ' Rouge clair
                ElseIf missingRate > 0.3 Then
                    .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 230, 200)  ' Orange clair
                Else
                    .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 255, 200)  ' Jaune clair
                End If

                .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Borders.LineStyle = xlContinuous
                anomalyRow = anomalyRow + 1
                totalAnomalies = totalAnomalies + 1
            End With
        End If
    Next col
End Sub

Sub DetectDuplicates(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, ByRef anomalyRow As Long, ByRef totalAnomalies As Long)
    '=========================================
    ' Détection des lignes complètement dupliquées
    '=========================================

    Dim rowDict As Object
    Dim row As Long
    Dim col As Long
    Dim rowSignature As String
    Dim duplicateCount As Long

    Set rowDict = CreateObject("Scripting.Dictionary")
    duplicateCount = 0

    ' Créer une signature pour chaque ligne
    For row = 2 To sourceRange.Rows.Count  ' Commencer après l'en-tête
        rowSignature = ""

        For col = 1 To sourceRange.Columns.Count
            rowSignature = rowSignature & "|" & CStr(sourceRange.Cells(row, col).Value)
        Next col

        If rowDict.Exists(rowSignature) Then
            duplicateCount = duplicateCount + 1
        Else
            rowDict.Add rowSignature, row
        End If
    Next row

    ' Signaler s'il y a des doublons
    If duplicateCount > 0 Then
        With analysisSheet
            .Cells(anomalyRow, 1).Value = "Lignes dupliquées"
            .Cells(anomalyRow, 2).Value = "Toutes colonnes"
            .Cells(anomalyRow, 3).Value = duplicateCount & " ligne(s) complètement identique(s) détectée(s)"
            .Cells(anomalyRow, 4).Value = "N/A"
            .Cells(anomalyRow, 5).Value = "Multiple"
            .Cells(anomalyRow, 6).Value = IIf(duplicateCount > 10, "ÉLEVÉE", IIf(duplicateCount > 5, "MODÉRÉE", "FAIBLE"))

            ' Formatage selon le nombre de doublons
            If duplicateCount > 10 Then
                .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 200, 200)  ' Rouge clair
            ElseIf duplicateCount > 5 Then
                .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 230, 200)  ' Orange clair
            Else
                .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 255, 200)  ' Jaune clair
            End If

            .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Borders.LineStyle = xlContinuous
            anomalyRow = anomalyRow + 1
            totalAnomalies = totalAnomalies + 1
        End With
    End If

    Set rowDict = Nothing
End Sub

Sub DetectDateInconsistencies(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant, ByRef anomalyRow As Long, ByRef totalAnomalies As Long)
    '=========================================
    ' Détection d'incohérences dans les dates
    '=========================================

    Dim col As Long
    Dim dataRange As Range
    Dim cell As Range
    Dim futureCount As Long
    Dim veryOldCount As Long
    Dim today As Date

    today = Date

    For col = 1 To UBound(columnTypes)
        If columnTypes(col) = "DATE" Then
            Set dataRange = sourceRange.Columns(col).Offset(1, 0).Resize(sourceRange.Rows.Count - 1)
            futureCount = 0
            veryOldCount = 0

            For Each cell In dataRange
                If IsDate(cell.Value) And cell.Value <> "" Then
                    Dim cellDate As Date
                    cellDate = CDate(cell.Value)

                    ' Détecter les dates futures (plus de 1 an dans le futur)
                    If cellDate > DateAdd("yyyy", 1, today) Then
                        futureCount = futureCount + 1
                    End If

                    ' Détecter les dates très anciennes (plus de 50 ans)
                    If cellDate < DateAdd("yyyy", -50, today) Then
                        veryOldCount = veryOldCount + 1
                    End If
                End If
            Next cell

            ' Signaler les anomalies de dates
            If futureCount > 0 Then
                With analysisSheet
                    .Cells(anomalyRow, 1).Value = "Dates futures suspectes"
                    .Cells(anomalyRow, 2).Value = headers(col)
                    .Cells(anomalyRow, 3).Value = futureCount & " date(s) dans un futur lointain (>1 an)"
                    .Cells(anomalyRow, 4).Value = "N/A"
                    .Cells(anomalyRow, 5).Value = "Multiple"
                    .Cells(anomalyRow, 6).Value = "MODÉRÉE"

                    .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 230, 200)  ' Orange clair
                    .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Borders.LineStyle = xlContinuous

                    anomalyRow = anomalyRow + 1
                    totalAnomalies = totalAnomalies + 1
                End With
            End If

            If veryOldCount > 0 Then
                With analysisSheet
                    .Cells(anomalyRow, 1).Value = "Dates très anciennes"
                    .Cells(anomalyRow, 2).Value = headers(col)
                    .Cells(anomalyRow, 3).Value = veryOldCount & " date(s) très anciennes (>50 ans)"
                    .Cells(anomalyRow, 4).Value = "N/A"
                    .Cells(anomalyRow, 5).Value = "Multiple"
                    .Cells(anomalyRow, 6).Value = "FAIBLE"

                    .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Interior.Color = RGB(255, 255, 200)  ' Jaune clair
                    .Range(.Cells(anomalyRow, 1), .Cells(anomalyRow, 6)).Borders.LineStyle = xlContinuous

                    anomalyRow = anomalyRow + 1
                    totalAnomalies = totalAnomalies + 1
                End With
            End If
        End If
    Next col
End Sub

Sub SummarizeAnomalies(analysisSheet As Worksheet, startRow As Long, totalAnomalies As Long)
    '=========================================
    ' Résumé des anomalies détectées
    '=========================================

    With analysisSheet
        If totalAnomalies = 0 Then
            .Cells(startRow + 1, 1).Value = "✅ AUCUNE ANOMALIE DÉTECTÉE"
            .Cells(startRow + 1, 1).Font.Bold = True
            .Cells(startRow + 1, 1).Font.Color = RGB(0, 128, 0)  ' Vert
            .Cells(startRow + 1, 1).Font.Size = 12

            .Cells(startRow + 3, 1).Value = "Les données semblent cohérentes et de bonne qualité. Aucune intervention nécessaire."
            .Cells(startRow + 3, 1).Font.Italic = True
        Else
            .Cells(startRow + 1, 1).Value = "⚠️ " & totalAnomalies & " TYPE(S) D'ANOMALIES DÉTECTÉES"
            .Cells(startRow + 1, 1).Font.Bold = True
            .Cells(startRow + 1, 1).Font.Color = RGB(255, 0, 0)  ' Rouge
            .Cells(startRow + 1, 1).Font.Size = 12

            .Cells(startRow + 3, 1).Value = "Recommandation : Examiner et corriger les anomalies avant l'analyse finale."
            .Cells(startRow + 3, 1).Font.Italic = True
            .Cells(startRow + 3, 1).Font.Color = RGB(255, 0, 0)
        End If
    End With
End Sub
```

## 7. Module de recommandations automatiques

```vba
Sub GenerateRecommendations(sourceRange As Range, analysisSheet As Worksheet, headers As Variant, columnTypes As Variant)
    '=========================================
    ' Génération automatique de recommandations
    '=========================================

    Dim startRow As Long
    Dim recRow As Long
    Dim recCount As Long

    startRow = 71  ' Position après le titre de section
    recRow = startRow + 2
    recCount = 0

    With analysisSheet
        .Cells(startRow, 1).Value = "Basées sur l'analyse de vos données, voici les recommandations automatiques :"
        .Cells(startRow, 1).Font.Italic = True

        ' Recommandations sur la qualité des données
        Call AddDataQualityRecommendations(sourceRange, analysisSheet, recRow, headers, columnTypes, recCount)

        ' Recommandations sur l'analyse statistique
        Call AddStatisticalRecommendations(sourceRange, analysisSheet, recRow, headers, columnTypes, recCount)

        ' Recommandations sur la visualisation
        Call AddVisualizationRecommendations(sourceRange, analysisSheet, recRow, headers, columnTypes, recCount)

        ' Recommandations pour les prochaines étapes
        Call AddNextStepsRecommendations(analysisSheet, recRow, recCount)

        ' Message de conclusion
        .Cells(recRow + 2, 1).Value = "💡 Ces recommandations sont générées automatiquement. Adaptez-les selon votre contexte métier spécifique."
        .Cells(recRow + 2, 1).Font.Italic = True
        .Cells(recRow + 2, 1).Font.Color = RGB(0, 102, 204)
    End With
End Sub

Sub AddDataQualityRecommendations(sourceRange As Range, analysisSheet As Worksheet, ByRef recRow As Long, headers As Variant, columnTypes As Variant, ByRef recCount As Long)
    '=========================================
    ' Recommandations sur la qualité des données
    '=========================================

    Dim totalCells As Long
    Dim emptyCells As Long
    Dim completenessRate As Double

    totalCells = (sourceRange.Rows.Count - 1) * sourceRange.Columns.Count
    emptyCells = Application.CountBlank(sourceRange.Offset(1, 0).Resize(sourceRange.Rows.Count - 1))
    completenessRate = (totalCells - emptyCells) / totalCells * 100

    With analysisSheet
        .Cells(recRow, 1).Value = "🔍 QUALITÉ DES DONNÉES"
        .Cells(recRow, 1).Font.Bold = True
        .Cells(recRow, 1).Font.Size = 12
        recRow = recRow + 1

        If completenessRate < 80 Then
            .Cells(recRow, 1).Value = "• Améliorer la complétude des données (actuellement " & Format(completenessRate, "0.0") & "%)"
            .Cells(recRow, 1).Font.Color = RGB(255, 0, 0)
            recRow = recRow + 1
            recCount = recCount + 1
        ElseIf completenessRate < 95 Then
            .Cells(recRow, 1).Value = "• Surveiller la qualité de saisie (complétude : " & Format(completenessRate, "0.0") & "%)"
            .Cells(recRow, 1).Font.Color = RGB(255, 165, 0)
            recRow = recRow + 1
            recCount = recCount + 1
        Else
            .Cells(recRow, 1).Value = "• Excellente qualité des données (complétude : " & Format(completenessRate, "0.0") & "%)"
            .Cells(recRow, 1).Font.Color = RGB(0, 128, 0)
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        ' Recommandations spécifiques selon le volume de données
        If sourceRange.Rows.Count - 1 < 30 Then
            .Cells(recRow, 1).Value = "• Collecter plus de données pour des analyses statistiques robustes (minimum 30 observations recommandé)"
            .Cells(recRow, 1).Font.Color = RGB(255, 165, 0)
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        recRow = recRow + 1
    End With
End Sub

Sub AddStatisticalRecommendations(sourceRange As Range, analysisSheet As Worksheet, ByRef recRow As Long, headers As Variant, columnTypes As Variant, ByRef recCount As Long)
    '=========================================
    ' Recommandations sur l'analyse statistique
    '=========================================

    Dim numericCount As Long
    Dim col As Long

    ' Compter les variables numériques
    For col = 1 To UBound(columnTypes)
        If columnTypes(col) = "NUMERIC" Then
            numericCount = numericCount + 1
        End If
    Next col

    With analysisSheet
        .Cells(recRow, 1).Value = "📊 ANALYSES STATISTIQUES"
        .Cells(recRow, 1).Font.Bold = True
        .Cells(recRow, 1).Font.Size = 12
        recRow = recRow + 1

        If numericCount >= 2 Then
            .Cells(recRow, 1).Value = "• Analyser les corrélations entre variables numériques pour identifier les relations"
            recRow = recRow + 1
            recCount = recCount + 1

            .Cells(recRow, 1).Value = "• Considérer une analyse de régression pour modéliser les relations causales"
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        If numericCount >= 1 Then
            .Cells(recRow, 1).Value = "• Vérifier la distribution des variables (normalité) avant tests statistiques avancés"
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        ' Recommandations selon la présence de données temporelles
        Dim hasDateColumn As Boolean
        For col = 1 To UBound(columnTypes)
            If columnTypes(col) = "DATE" Then
                hasDateColumn = True
                Exit For
            End If
        Next col

        If hasDateColumn And numericCount >= 1 Then
            .Cells(recRow, 1).Value = "• Analyser les tendances temporelles et la saisonnalité des données"
            recRow = recRow + 1
            recCount = recCount + 1

            .Cells(recRow, 1).Value = "• Considérer des prévisions basées sur les données historiques"
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        ' Recommandations sur les tests statistiques
        If sourceRange.Rows.Count - 1 >= 30 Then
            .Cells(recRow, 1).Value = "• Volume suffisant pour des tests d'hypothèses (t-test, ANOVA, etc.)"
            .Cells(recRow, 1).Font.Color = RGB(0, 128, 0)
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        recRow = recRow + 1
    End With
End Sub

Sub AddVisualizationRecommendations(sourceRange As Range, analysisSheet As Worksheet, ByRef recRow As Long, headers As Variant, columnTypes As Variant, ByRef recCount As Long)
    '=========================================
    ' Recommandations sur la visualisation
    '=========================================

    Dim numericCount As Long
    Dim textCount As Long
    Dim dateCount As Long
    Dim col As Long

    ' Compter les types de variables
    For col = 1 To UBound(columnTypes)
        Select Case columnTypes(col)
            Case "NUMERIC"
                numericCount = numericCount + 1
            Case "TEXT"
                textCount = textCount + 1
            Case "DATE"
                dateCount = dateCount + 1
        End Select
    Next col

    With analysisSheet
        .Cells(recRow, 1).Value = "📈 VISUALISATIONS AVANCÉES"
        .Cells(recRow, 1).Font.Bold = True
        .Cells(recRow, 1).Font.Size = 12
        recRow = recRow + 1

        If numericCount >= 2 Then
            .Cells(recRow, 1).Value = "• Créer une matrice de nuages de points pour explorer toutes les corrélations"
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        If textCount >= 1 And numericCount >= 1 Then
            .Cells(recRow, 1).Value = "• Utiliser des graphiques en boîtes (boxplots) pour comparer les distributions par catégorie"
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        If dateCount >= 1 And numericCount >= 1 Then
            .Cells(recRow, 1).Value = "• Développer un tableau de bord temporel interactif avec filtres par période"
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        .Cells(recRow, 1).Value = "• Ajouter des graphiques de performance (indicateurs clés, jauges, sparklines)"
        recRow = recRow + 1
        recCount = recCount + 1

        If sourceRange.Rows.Count - 1 > 100 Then
            .Cells(recRow, 1).Value = "• Considérer l'échantillonnage ou l'agrégation pour les visualisations avec beaucoup de données"
            recRow = recRow + 1
            recCount = recCount + 1
        End If

        recRow = recRow + 1
    End With
End Sub

Sub AddNextStepsRecommendations(analysisSheet As Worksheet, ByRef recRow As Long, ByRef recCount As Long)
    '=========================================
    ' Recommandations pour les prochaines étapes
    '=========================================

    With analysisSheet
        .Cells(recRow, 1).Value = "🚀 PROCHAINES ÉTAPES"
        .Cells(recRow, 1).Font.Bold = True
        .Cells(recRow, 1).Font.Size = 12
        recRow = recRow + 1

        .Cells(recRow, 1).Value = "• Automatiser cette analyse en créant une macro personnalisée pour vos données récurrentes"
        recRow = recRow + 1
        recCount = recCount + 1

        .Cells(recRow, 1).Value = "• Mettre en place un système de collecte de données plus structuré"
        recRow = recRow + 1
        recCount = recCount + 1

        .Cells(recRow, 1).Value = "• Former les équipes à l'interprétation de ces analyses statistiques"
        recRow = recRow + 1
        recCount = recCount + 1

        .Cells(recRow, 1).Value = "• Intégrer ces analyses dans un processus de décision régulier"
        recRow = recRow + 1
        recCount = recCount + 1

        .Cells(recRow, 1).Value = "• Considérer l'utilisation d'outils d'analyse plus avancés (Power BI, R, Python) pour des besoins complexes"
        recRow = recRow + 1
        recCount = recCount + 1

        recRow = recRow + 1
    End With
End Sub
```

## 8. Fonctions utilitaires et export

```vba
Sub ExportAnalysisReport()
    '=========================================
    ' Export du rapport d'analyse en PDF
    '=========================================

    Dim ws As Worksheet
    Dim fileName As String
    Dim filePath As String

    Set ws = ThisWorkbook.Sheets("ANALYSE_DONNEES")

    ' Générer le nom de fichier
    fileName = "Rapport_Analyse_" & Format(Now(), "yyyy-mm-dd_hh-mm") & ".pdf"
    filePath = ThisWorkbook.Path

    If filePath = "" Then
        filePath = Environ("USERPROFILE") & "\Desktop"  ' Sauvegarder sur le bureau si pas de chemin
    End If

    ' Configurer l'impression pour le PDF
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = "A1:L100"  ' Ajuster selon le contenu
        .CenterHorizontally = True
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
    End With

    ' Exporter en PDF
    On Error GoTo ExportError
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                          fileName:=filePath & "\" & fileName, _
                          Quality:=xlQualityStandard, _
                          IncludeDocProps:=True, _
                          IgnorePrintAreas:=False, _
                          OpenAfterPublish:=True

    MsgBox "Rapport exporté en PDF avec succès !" & vbNewLine & _
           "Fichier : " & fileName & vbNewLine & _
           "Emplacement : " & filePath, vbInformation
    Exit Sub

ExportError:
    MsgBox "Erreur lors de l'export PDF : " & Err.Description, vbExclamation
End Sub

Sub CreateAnalysisTemplate()
    '=========================================
    ' Création d'un modèle pour analyses futures
    '=========================================

    Dim templateSheet As Worksheet
    Dim templateName As String

    templateName = "MODELE_ANALYSE_" & Format(Now(), "ddmm")

    ' Créer une nouvelle feuille modèle
    Set templateSheet = ThisWorkbook.Sheets.Add
    templateSheet.Name = templateName

    With templateSheet
        ' Structure du modèle
        .Cells(1, 1).Value = "MODÈLE D'ANALYSE DE DONNÉES"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        .Range("A1:F1").Merge
        .Range("A1").HorizontalAlignment = xlCenter

        .Cells(3, 1).Value = "Instructions :"
        .Cells(3, 1).Font.Bold = True

        .Cells(4, 1).Value = "1. Copiez vos données dans cette feuille à partir de la ligne 8"
        .Cells(5, 1).Value = "2. Assurez-vous que la ligne 8 contient les en-têtes de colonnes"
        .Cells(6, 1).Value = "3. Exécutez la macro 'StartDataAnalysis' pour lancer l'analyse"

        ' Zone de données
        .Cells(8, 1).Value = "En-tête 1"
        .Cells(8, 2).Value = "En-tête 2"
        .Cells(8, 3).Value = "En-tête 3"
        .Cells(8, 4).Value = "En-tête 4"
        .Cells(8, 5).Value = "En-tête 5"

        ' Formatage des en-têtes
        With .Range("A8:E8")
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
        End With

        ' Exemples de données
        .Cells(9, 1).Value = "Exemple 1"
        .Cells(9, 2).Value = 100
        .Cells(9, 3).Value = Date
        .Cells(9, 4).Value = "Catégorie A"
        .Cells(9, 5).Value = 85.5

        .Cells(10, 1).Value = "Exemple 2"
        .Cells(10, 2).Value = 150
        .Cells(10, 3).Value = Date + 1
        .Cells(10, 4).Value = "Catégorie B"
        .Cells(10, 5).Value = 92.3

        ' Instructions de fin
        .Cells(13, 1).Value = "💡 Remplacez les données d'exemple par vos propres données"
        .Cells(13, 1).Font.Italic = True
        .Cells(13, 1).Font.Color = RGB(0, 102, 204)

        .Columns.AutoFit
    End With

    MsgBox "Modèle d'analyse créé : " & templateName & vbNewLine & _
           "Utilisez cette feuille pour vos prochaines analyses.", vbInformation
End Sub

Sub OptimizeAnalysisWorkbook()
    '=========================================
    ' Optimisation du classeur d'analyse
    '=========================================

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Supprimer les feuilles temporaires ou anciennes analyses
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 8) = "ANALYSE_" And ws.Name <> "ANALYSE_DONNEES" Then
            If MsgBox("Supprimer l'ancienne analyse : " & ws.Name & " ?", vbYesNo + vbQuestion) = vbYes Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next ws

    ' Nettoyer les objets graphiques inutiles
    If ThisWorkbook.Sheets("ANALYSE_DONNEES").ChartObjects.Count > 10 Then
        If MsgBox("Supprimer les anciens graphiques pour améliorer les performances ?", vbYesNo + vbQuestion) = vbYes Then
            Dim i As Long
            For i = ThisWorkbook.Sheets("ANALYSE_DONNEES").ChartObjects.Count To 6 Step -1
                ThisWorkbook.Sheets("ANALYSE_DONNEES").ChartObjects(i).Delete
            Next i
        End If
    End If

    ' Recalculer et optimiser
    Application.Calculate
    ThisWorkbook.Save

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Optimisation terminée. Classeur allégé et performances améliorées.", vbInformation
End Sub

Function GetDataSummary(sourceRange As Range) As String
    '=========================================
    ' Génération d'un résumé textuel des données
    '=========================================

    Dim summary As String
    Dim totalRows As Long
    Dim totalCols As Long
    Dim numericCols As Long
    Dim textCols As Long
    Dim dateCols As Long
    Dim col As Long

    totalRows = sourceRange.Rows.Count - 1  ' Exclure l'en-tête
    totalCols = sourceRange.Columns.Count

    ' Analyser les types de colonnes
    For col = 1 To totalCols
        Dim colRange As Range
        Set colRange = sourceRange.Columns(col).Offset(1, 0).Resize(totalRows)

        Dim numericCount As Long
        Dim dateCount As Long
        Dim textCount As Long
        Dim cell As Range

        For Each cell In colRange
            If cell.Value <> "" Then
                If IsNumeric(cell.Value) Then
                    numericCount = numericCount + 1
                ElseIf IsDate(cell.Value) Then
                    dateCount = dateCount + 1
                Else
                    textCount = textCount + 1
                End If
            End If
        Next cell

        ' Déterminer le type majoritaire
        If dateCount > numericCount And dateCount > textCount Then
            dateCols = dateCols + 1
        ElseIf numericCount > textCount Then
            numericCols = numericCols + 1
        Else
            textCols = textCols + 1
        End If
    Next col

    ' Construire le résumé
    summary = "Analyse de " & Format(totalRows, "#,##0") & " observations sur " & totalCols & " variables" & vbNewLine
    summary = summary & "• Variables numériques : " & numericCols & vbNewLine
    summary = summary & "• Variables textuelles : " & textCols & vbNewLine
    summary = summary & "• Variables temporelles : " & dateCols & vbNewLine

    ' Calculer la complétude
    Dim totalCells As Long
    Dim emptyCells As Long
    totalCells = totalRows * totalCols
    emptyCells = Application.CountBlank(sourceRange.Offset(1, 0).Resize(totalRows))

    summary = summary & "• Taux de complétude : " & Format((totalCells - emptyCells) / totalCells * 100, "0.0") & "%"

    GetDataSummary = summary
End Function
```

## Installation et utilisation de l'outil

### Guide d'installation

1. **Création du fichier**
   - Ouvrir un nouveau classeur Excel
   - Sauvegarder au format .xlsm (activé pour les macros)
   - Activer l'onglet Développeur si nécessaire

2. **Installation du code VBA**
   - Ouvrir l'éditeur VBA (Alt + F11)
   - Créer un nouveau module standard
   - Copier l'intégralité du code développé
   - Sauvegarder le classeur

3. **Préparation des données**
   - Organiser vos données avec des en-têtes en première ligne
   - S'assurer de la cohérence des types de données par colonne
   - Nettoyer les données aberrantes évidentes

### Guide d'utilisation

#### Analyse de base
1. **Ouvrir le fichier** contenant vos données
2. **Sélectionner la plage** de données à analyser (avec en-têtes)
3. **Exécuter la macro** `StartDataAnalysis` via Alt + F8
4. **Suivre les instructions** à l'écran
5. **Consulter les résultats** dans la feuille "ANALYSE_DONNEES"

#### Fonctionnalités avancées
- **Export PDF** : Utiliser `ExportAnalysisReport()` pour sauvegarder
- **Modèle réutilisable** : Créer avec `CreateAnalysisTemplate()`
- **Optimisation** : Nettoyer avec `OptimizeAnalysisWorkbook()`

### Types d'analyses supportées

#### Données de ventes
```
Date       | Vendeur | Produit    | Quantité | Prix_Unit | Total
01/01/2024 | Martin  | Laptop     | 2        | 800       | 1600
02/01/2024 | Durand  | Souris     | 10       | 25        | 250
```

#### Données d'enquête
```
Age | Sexe    | Satisfaction | Ville     | Revenus
25  | Homme   | 8           | Paris     | 35000
32  | Femme   | 9           | Lyon      | 42000
```

#### Données de production
```
Date_Prod  | Machine | Defauts | Production | Efficacite
15/01/2024 | A1      | 3       | 1000       | 97.5
16/01/2024 | A2      | 1       | 950        | 99.2
```

## Avantages de l'outil

### Pour les débutants
- **Automatisation complète** : Aucune connaissance statistique requise
- **Interface guidée** : Instructions claires à chaque étape
- **Interprétation automatique** : Recommandations générées automatiquement
- **Visualisations adaptées** : Graphiques choisis selon le type de données

### Pour les utilisateurs avancés
- **Code modulaire** : Facilement personnalisable et extensible
- **Analyses robustes** : Techniques statistiques éprouvées
- **Détection intelligente** : Identification automatique des types de données
- **Export professionnel** : Rapports prêts à présenter

### Pour l'entreprise
- **Gain de temps** : Heures d'analyse réduites à quelques minutes
- **Standardisation** : Même qualité d'analyse pour tous les utilisateurs
- **Reproductibilité** : Résultats cohérents et documentés
- **Formation réduite** : Utilisation immédiate sans formation préalable

## Limitations et perspectives d'évolution

### Limitations actuelles
- **Volume de données** : Optimisé pour moins de 10 000 lignes
- **Types de graphiques** : Limité aux graphiques Excel standard
- **Tests statistiques** : Analyses descriptives uniquement
- **Langues** : Interface en français uniquement

### Évolutions possibles
1. **Machine Learning** : Intégration d'algorithmes de classification/prédiction
2. **Big Data** : Support de sources de données externes (SQL, APIs)
3. **Interactivité** : Tableaux de bord dynamiques avec sélecteurs
4. **Collaboration** : Partage et commentaires en ligne
5. **IA Générative** : Interprétation automatique avancée des résultats

## Conclusion

Cet outil d'analyse de données démontre la capacité de VBA à créer des solutions analytiques sophistiquées accessibles à tous. Avec moins de 800 lignes de code, nous avons développé un analyseur capable de :

**Automatiser complètement** le processus d'analyse depuis la sélection des données jusqu'aux recommandations finales, **s'adapter intelligemment** aux différents types de données, **détecter automatiquement** les anomalies et problèmes de qualité, et **générer des rapports professionnels** prêts à présenter.

Ce projet illustre parfaitement comment VBA peut transformer des tâches analytiques complexes en processus simples et automatisés, démocratisant ainsi l'accès à l'analyse de données pour tous les utilisateurs, quelle que soit leur expertise technique.

### Impact pédagogique
- **Intégration des concepts** : Variables, boucles, fonctions, objets Excel
- **Algorithmes avancés** : Statistiques, détection d'anomalies, corrélations
- **Architecture logicielle** : Modularité, réutilisabilité, maintenabilité
- **Résolution de problèmes** : Approche méthodique et documentation

Ce projet constitue un excellent tremplin vers des solutions d'analyse de données plus avancées et prépare efficacement à l'utilisation d'outils spécialisés comme R, Python ou Power BI.

⏭️
