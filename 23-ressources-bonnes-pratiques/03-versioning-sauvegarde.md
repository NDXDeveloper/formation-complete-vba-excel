🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.3 Versioning et sauvegarde

## Introduction

Imaginez que vous travaillez sur un important projet VBA depuis plusieurs semaines. Tout fonctionne parfaitement, mais vous décidez d'ajouter une nouvelle fonctionnalité. Après quelques modifications, plus rien ne marche et vous ne vous souvenez plus exactement de ce qui fonctionnait avant. Sans système de versioning et de sauvegarde, vous pourriez perdre des heures, voire des jours de travail !

Le versioning (ou gestion des versions) consiste à garder une trace de toutes les modifications apportées à votre code au fil du temps. C'est comme avoir une machine à remonter le temps pour votre projet : vous pouvez revenir à n'importe quelle version antérieure qui fonctionnait.

La sauvegarde, elle, consiste à créer des copies de sécurité de votre travail pour le protéger contre les pannes, les suppressions accidentelles ou la corruption de fichiers.

## Pourquoi le versioning est-il crucial ?

### 1. Pouvoir revenir en arrière
```vba
' Version qui fonctionnait (v1.0)
Function CalculerRemise(montant As Double) As Double
    If montant > 1000 Then
        CalculerRemise = montant * 0.1
    Else
        CalculerRemise = 0
    End If
End Function

' Nouvelle version qui ne fonctionne plus (v1.1) - Bug introduit !
Function CalculerRemise(montant As Double) As Double
    If montant > 1000 Then
        CalculerRemise = montant * 0.1
    ElseIf montant > 500 Then
        CalculerRemise = montant * 0.05
    ' Oups ! J'ai oublié le cas Else - bug !
    End If
End Function
```

Avec un système de versioning, vous pouvez facilement revenir à la version 1.0 qui fonctionnait.

### 2. Tracer l'évolution du projet

Le versioning vous permet de voir :
- Quand une fonctionnalité a été ajoutée
- Qui a fait quelle modification (dans un contexte d'équipe)
- Pourquoi une modification a été faite
- L'impact de chaque changement

### 3. Tester en toute sécurité

Vous pouvez expérimenter de nouvelles idées sans crainte, sachant que vous pouvez toujours revenir à une version stable.

### 4. Collaborer efficacement

Si plusieurs personnes travaillent sur le même projet, le versioning permet de gérer les modifications de chacun sans conflits.

## Systèmes de versioning simples pour débutants

### 1. Versioning manuel par nom de fichier

C'est la méthode la plus simple pour commencer. Vous sauvegardez votre fichier Excel avec un nom différent à chaque étape importante :

```
Structure de noms suggérée :
GestionCommerciale_v1.0_InitialeFonctionnelle.xlsm
GestionCommerciale_v1.1_AjoutCalculTVA.xlsm
GestionCommerciale_v1.2_CorrectionBugRemise.xlsm
GestionCommerciale_v2.0_NouvelleFonctionRapport.xlsm
GestionCommerciale_v2.1_OptimisationPerformances.xlsm
```

**Avantages :**
- Très simple à mettre en place
- Aucun outil externe nécessaire
- Fonctionne immédiatement

**Inconvénients :**
- Prend beaucoup d'espace disque
- Difficile de comparer les versions
- Risque d'oublier de sauvegarder

### 2. Convention de nommage des versions

Utilisez un système de numérotation cohérent :

**Format recommandé : MAJEUR.MINEUR.CORRECTION**

- **MAJEUR** : Changements importants qui modifient fondamentalement le fonctionnement
- **MINEUR** : Nouvelles fonctionnalités ajoutées
- **CORRECTION** : Corrections de bugs sans nouvelles fonctionnalités

```
Exemples :
v1.0.0 - Version initiale fonctionnelle
v1.1.0 - Ajout de la gestion des remises
v1.1.1 - Correction du bug de calcul TVA
v1.2.0 - Ajout de l'export PDF
v2.0.0 - Refonte complète de l'interface
```

### 3. Journal des modifications

Tenez un journal détaillé dans un module dédié ou un fichier texte :

```vba
'**********************************************************************
' JOURNAL DES MODIFICATIONS - Projet Gestion Commerciale
'**********************************************************************
'
' v1.2.1 - 22/01/2025 - Jean Martin
' CORRECTIONS :
' - Correction du bug de division par zéro dans CalculerMoyenne()
' - Correction de l'affichage des dates au format français
'
' v1.2.0 - 20/01/2025 - Jean Martin
' NOUVELLES FONCTIONNALITÉS :
' - Ajout de la fonction d'export automatique des rapports
' - Nouvelle procédure de sauvegarde automatique
' AMÉLIORATIONS :
' - Optimisation des performances pour les gros volumes de données
' - Amélioration des messages d'erreur utilisateur
'
' v1.1.2 - 18/01/2025 - Marie Dubois
' CORRECTIONS :
' - Correction du plantage lors de la suppression d'un client
' - Résolution du problème d'encodage des caractères spéciaux
'
' v1.1.1 - 15/01/2025 - Jean Martin
' CORRECTIONS :
' - Correction du calcul des remises pour les clients VIP
' - Correction de l'arrondi des montants TTC
'**********************************************************************
```

## Stratégies de sauvegarde

### 1. La règle 3-2-1

C'est une règle d'or en matière de sauvegarde :
- **3** copies de vos données importantes
- Sur **2** supports différents
- Dont **1** copie externalisée (cloud, disque externe, etc.)

**Exemple pratique :**
1. **Copie de travail** : Sur votre ordinateur principal
2. **Copie locale** : Sur un disque dur externe ou un autre ordinateur
3. **Copie externe** : Sur OneDrive, Google Drive, Dropbox, etc.

### 2. Sauvegarde automatique d'Excel

Activez la sauvegarde automatique d'Excel pour éviter les pertes accidentelles :

```
Dans Excel :
1. Fichier → Options → Enregistrement
2. Cocher "Enregistrer les informations de récupération automatique toutes les X minutes"
3. Cocher "Conserver la dernière version récupérée automatiquement..."
```

### 3. Organisation des dossiers de sauvegarde

Créez une structure de dossiers claire :

```
📁 Projets VBA/
├── 📁 GestionCommerciale/
│   ├── 📁 Versions/
│   │   ├── 📄 v1.0.0_GestionCommerciale.xlsm
│   │   ├── 📄 v1.1.0_GestionCommerciale.xlsm
│   │   ├── 📄 v1.2.0_GestionCommerciale.xlsm
│   │   └── 📄 v1.2.1_GestionCommerciale.xlsm
│   ├── 📁 Sauvegardes/
│   │   ├── 📁 2025-01-22/
│   │   ├── 📁 2025-01-21/
│   │   └── 📁 2025-01-20/
│   ├── 📁 Documentation/
│   │   ├── 📄 Manuel_Utilisateur.docx
│   │   └── 📄 Cahier_des_charges.docx
│   └── 📄 GestionCommerciale_ACTUEL.xlsm
```

### 4. Fréquence de sauvegarde

**Sauvegarde quotidienne :**
- À la fin de chaque journée de travail
- Avant et après une modification importante
- Avant de tester une nouvelle fonctionnalité

**Sauvegarde de version :**
- Quand une fonctionnalité est terminée et testée
- Avant de commencer une nouvelle fonctionnalité majeure
- À la fin de chaque semaine
- Avant une démonstration ou une mise en production

## Export et sauvegarde du code VBA

### 1. Exporter les modules individuellement

Vous pouvez exporter chaque module VBA séparément :

```
Dans l'éditeur VBA :
1. Clic droit sur le module dans l'explorateur de projets
2. "Exporter un fichier..."
3. Choisir l'emplacement et le nom
4. Le fichier .bas contient tout le code du module
```

**Avantages :**
- Sauvegarde pure du code (sans les données Excel)
- Peut être versionnée avec des outils comme Git
- Facilite la comparaison entre versions

### 2. Procédure automatique d'export

```vba
'**********************************************************************
' Procédure : ExporterTousLesModules
' Description : Exporte automatiquement tous les modules VBA du projet
'              dans un dossier dédié pour sauvegarde
'**********************************************************************
Sub ExporterTousLesModules()
    Dim vbComp As Object
    Dim cheminExport As String
    Dim nomFichier As String
    Dim extension As String

    ' Définir le dossier d'export (à adapter selon votre structure)
    cheminExport = ThisWorkbook.Path & "\Export_Code_" & Format(Date, "yyyy-mm-dd") & "\"

    ' Créer le dossier s'il n'existe pas
    If Dir(cheminExport, vbDirectory) = "" Then
        MkDir cheminExport
    End If

    ' Parcourir tous les composants VBA du projet
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1  ' Module standard
                extension = ".bas"
            Case 2  ' Module de classe
                extension = ".cls"
            Case 3  ' Module de formulaire
                extension = ".frm"
            Case 100 ' Module de document (feuille, classeur)
                extension = ".cls"
        End Select

        ' Définir le nom du fichier d'export
        nomFichier = cheminExport & vbComp.Name & extension

        ' Exporter le module
        vbComp.Export nomFichier

        Debug.Print "Module exporté : " & nomFichier
    Next vbComp

    MsgBox "Export terminé dans : " & cheminExport
End Sub
```

### 3. Import de modules sauvegardés

```vba
'**********************************************************************
' Procédure : ImporterModule
' Description : Importe un module VBA depuis un fichier .bas/.cls
'**********************************************************************
Sub ImporterModule(cheminFichier As String)
    On Error GoTo GestionErreur

    ' Importer le module
    ThisWorkbook.VBProject.VBComponents.Import cheminFichier

    MsgBox "Module importé avec succès : " & cheminFichier
    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'import : " & Err.Description
End Sub
```

## Outils de versioning plus avancés

### 1. Git pour les développeurs

Si vous voulez aller plus loin, Git est l'outil de versioning professionnel le plus utilisé :

**Avantages :**
- Suivi précis de toutes les modifications
- Possibilité de créer des branches pour tester
- Collaboration facilitée en équipe
- Historique complet et détaillé

**Inconvénients pour VBA :**
- Courbe d'apprentissage importante
- Mieux adapté au code pur qu'aux fichiers Excel
- Nécessite d'exporter les modules VBA

### 2. OneDrive/SharePoint avec versioning

Microsoft Office propose un versioning automatique :

**Configuration :**
1. Sauvegarder votre fichier Excel sur OneDrive
2. Dans Excel Online, aller dans "Fichier" → "Informations" → "Historique des versions"
3. Possibilité de voir et restaurer les versions précédentes

## Bonnes pratiques de sauvegarde pour débutants

### 1. Sauvegarde avant modification

```vba
' Toujours faire une sauvegarde avant une modification importante
Sub ProcedureComplexe()
    ' Première ligne : commentaire de sauvegarde
    ' SAUVEGARDE : v1.2.3 → v1.2.4 - Ajout de la validation des données

    ' Votre code ici...
End Sub
```

### 2. Nommage explicite des sauvegardes

```
Noms explicites :
✅ GestionStock_v2.1_AvantModificationCalculs_20250122.xlsm
✅ Rapports_v1.5_FonctionExportPDFOK_20250122.xlsm

Noms peu clairs :
❌ Copie de GestionStock.xlsm
❌ GestionStock - Copie (2).xlsm
❌ Nouveau fichier.xlsm
```

### 3. Documentation des sauvegardes

Tenez un fichier texte simple listant vos sauvegardes :

```
📄 Journal_Sauvegardes.txt

=== JOURNAL DES SAUVEGARDES - Projet Gestion Commerciale ===

22/01/2025 - v1.2.4
- Sauvegarde avant ajout de la validation automatique des données
- Dernière version stable : toutes les fonctions de base OK
- Prochaine étape : ajouter contrôles de saisie

21/01/2025 - v1.2.3
- Sauvegarde après correction du bug de calcul des remises
- Tests réussis sur tous les cas de figure
- Version recommandée pour utilisation en production

20/01/2025 - v1.2.2
- Sauvegarde de développement - NE PAS UTILISER
- Bug connu dans le calcul des totaux
```

### 4. Checklist avant sauvegarde

Avant chaque sauvegarde importante :

```
□ Le code fonctionne sans erreur
□ Tous les tests essentiels sont passés
□ Les commentaires sont à jour
□ Le numéro de version est incrémenté
□ Le journal des modifications est mis à jour
□ Les fichiers temporaires sont supprimés
□ La sauvegarde précédente est conservée
```

## Récupération après problème

### 1. Si votre fichier Excel est corrompu

```
Méthodes de récupération :
1. Excel → Fichier → Ouvrir → Ouvrir et réparer
2. Restaurer depuis la sauvegarde automatique d'Excel
3. Revenir à la dernière version sauvegardée manuellement
4. Réimporter les modules VBA dans un nouveau classeur
```

### 2. Si vous avez perdu du code

```
Sources de récupération :
1. Fichiers .bas/.cls exportés automatiquement
2. Versions précédentes du fichier Excel
3. Copies dans la corbeille (si suppression récente)
4. Historique OneDrive/SharePoint
5. Sauvegarde sur disque externe
```

### 3. Procédure de récupération d'urgence

```vba
'**********************************************************************
' PROCÉDURE D'URGENCE : Récupération de code
' À utiliser uniquement en cas de perte majeure de données
'**********************************************************************
Sub RecuperationUrgence()
    Dim reponse As String

    ' Confirmer la procédure d'urgence
    reponse = MsgBox("ATTENTION : Procédure de récupération d'urgence" & vbCrLf & _
                    "Voulez-vous importer tous les modules de sauvegarde ?", _
                    vbYesNo + vbExclamation, "Récupération d'urgence")

    If reponse = vbYes Then
        ' Indiquer le dossier contenant les modules sauvegardés
        Dim cheminSauvegarde As String
        cheminSauvegarde = "C:\Sauvegardes\ModulesVBA\"  ' À adapter

        ' Ici, vous pourriez ajouter le code pour importer automatiquement
        ' tous les fichiers .bas et .cls du dossier

        MsgBox "Récupération terminée. Vérifiez le fonctionnement de votre application."
    End If
End Sub
```

## Planning de sauvegarde recommandé

### Pour un projet personnel

```
📅 Planning de sauvegarde :

QUOTIDIEN (fin de journée) :
- Sauvegarde automatique Excel activée
- Copie de travail sur disque dur principal

HEBDOMADAIRE (vendredi) :
- Sauvegarde versionnée (si modifications importantes)
- Copie sur support externe ou cloud
- Export des modules VBA modifiés

MENSUEL :
- Archivage des anciennes versions
- Nettoyage des fichiers temporaires
- Vérification de l'intégrité des sauvegardes
```

### Pour un projet professionnel

```
📅 Planning de sauvegarde professionnel :

AVANT CHAQUE MODIFICATION :
- Sauvegarde de la version courante
- Note dans le journal des modifications

QUOTIDIEN :
- Sauvegarde sur serveur d'entreprise
- Export automatique des modules

AVANT MISE EN PRODUCTION :
- Sauvegarde complète de la version stable
- Documentation de déploiement
- Plan de retour arrière (rollback)
```

## Résumé des bonnes pratiques

### Pour le versioning :
1. **Utilisez un système de numérotation cohérent** (MAJEUR.MINEUR.CORRECTION)
2. **Documentez chaque version** avec ses modifications et corrections
3. **Sauvegardez avant chaque modification importante**
4. **Gardez les versions stables** comme points de retour
5. **Exportez régulièrement vos modules VBA** pour sauvegarde pure du code

### Pour la sauvegarde :
1. **Appliquez la règle 3-2-1** (3 copies, 2 supports, 1 externe)
2. **Automatisez ce qui peut l'être** (sauvegarde Excel, scripts d'export)
3. **Organisez vos dossiers clairement** avec une structure logique
4. **Testez vos procédures de récupération** avant d'en avoir besoin
5. **Documentez vos sauvegardes** pour savoir quoi restaurer

### En cas de problème :
1. **Ne paniquez pas** - vos sauvegardes sont là pour ça
2. **Identifiez la dernière version stable** connue
3. **Restaurez progressivement** en testant à chaque étape
4. **Documentez l'incident** pour éviter qu'il se reproduise

Le versioning et la sauvegarde peuvent sembler fastidieux au début, mais ils deviennent rapidement des habitudes naturelles. Et le jour où ils vous sauveront des heures de travail, vous serez très content de les avoir mis en place !

⏭️
