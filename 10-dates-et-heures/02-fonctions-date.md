🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.2. Fonctions de date (Now, Date, Time)

## Introduction aux fonctions de date de base

VBA propose trois fonctions essentielles pour récupérer les informations de date et d'heure actuelles du système. Ces fonctions sont les plus utilisées dans la programmation quotidienne et constituent la base de nombreuses applications temporelles.

## La fonction Now()

### Qu'est-ce que Now() ?

La fonction `Now()` retourne la **date ET l'heure actuelles** du système. C'est la fonction la plus complète pour obtenir un timestamp précis du moment d'exécution.

### Syntaxe

```vba
Now()
```

**Remarque :** `Now` n'a pas de paramètres et les parenthèses sont optionnelles.

### Utilisation de base

```vba
Dim maintenant As Date
maintenant = Now()

Debug.Print maintenant
' Affiche par exemple : 15/03/2024 14:32:17
```

### Exemples pratiques

```vba
' Horodatage d'une action
Sub EnregistrerAction()
    Dim horodatage As Date
    horodatage = Now

    Range("A1").Value = "Action effectuée le : " & Format(Now, "dd/mm/yyyy à hh:nn:ss")
    ' Résultat : "Action effectuée le : 15/03/2024 à 14:32:17"
End Sub

' Mesure du temps d'exécution
Sub MesurerTemps()
    Dim debut As Date
    Dim fin As Date

    debut = Now

    ' Votre code à mesurer ici
    Application.Wait Now + TimeValue("00:00:02")  ' Attendre 2 secondes

    fin = Now

    Debug.Print "Temps d'exécution : " & Format(fin - debut, "hh:nn:ss")
End Sub
```

### Utilisation avec des cellules Excel

```vba
' Insérer la date/heure actuelle dans une cellule
Range("A1").Value = Now

' Formatter l'affichage dans la cellule
Range("A1").Value = Now
Range("A1").NumberFormat = "dd/mm/yyyy hh:mm:ss"
```

## La fonction Date()

### Qu'est-ce que Date() ?

La fonction `Date()` retourne uniquement la **date actuelle** du système, sans l'information d'heure. L'heure est automatiquement définie à 00:00:00 (minuit).

### Syntaxe

```vba
Date()
```

### Utilisation de base

```vba
Dim aujourdhui As Date
aujourdhui = Date()

Debug.Print aujourdhui
' Affiche par exemple : 15/03/2024 (sans l'heure)
```

### Exemples pratiques

```vba
' Vérifier si c'est un jour spécifique
Sub VerifierJour()
    If Date = #12/25/2024# Then
        MsgBox "C'est Noël !"
    End If
End Sub

' Calculer un âge en jours
Sub CalculerAge()
    Dim naissance As Date
    Dim agejours As Long

    naissance = #15/06/1990#
    agejours = Date - naissance

    Debug.Print "Age en jours : " & agejours
End Sub

' Créer un nom de fichier avec la date
Sub CreerNomFichier()
    Dim nomFichier As String
    nomFichier = "Rapport_" & Format(Date, "yyyy-mm-dd") & ".xlsx"

    Debug.Print nomFichier
    ' Résultat : "Rapport_2024-03-15.xlsx"
End Sub
```

### Différence avec Now()

```vba
' Comparaison des deux fonctions
Debug.Print "Now() : " & Now
Debug.Print "Date() : " & Date

' Résultat possible :
' Now() : 15/03/2024 14:32:17
' Date() : 15/03/2024 00:00:00
```

## La fonction Time()

### Qu'est-ce que Time() ?

La fonction `Time()` retourne uniquement l'**heure actuelle** du système, sans l'information de date. La date est automatiquement définie au 30 décembre 1899 (valeur 0 du système de dates VBA).

### Syntaxe

```vba
Time()
```

### Utilisation de base

```vba
Dim maintenant As Date
maintenant = Time()

Debug.Print maintenant
' Affiche par exemple : 14:32:17
```

### Exemples pratiques

```vba
' Vérifier les heures d'ouverture
Sub VerifierHeuresOuverture()
    Dim heureOuverture As Date
    Dim heureFermeture As Date

    heureOuverture = #8:00 AM#
    heureFermeture = #6:00 PM#

    If Time >= heureOuverture And Time <= heureFermeture Then
        MsgBox "Le magasin est ouvert"
    Else
        MsgBox "Le magasin est fermé"
    End If
End Sub

' Logger l'heure d'une action
Sub LoggerAction()
    Dim message As String
    message = Format(Time, "hh:nn:ss") & " - Action terminée"

    Debug.Print message
    ' Résultat : "14:32:17 - Action terminée"
End Sub

' Calculer la durée depuis le début de journée
Sub DureeDepuisMinuit()
    Dim duree As Date
    duree = Time  ' L'heure actuelle représente aussi la durée depuis minuit

    Debug.Print "Durée depuis minuit : " & Format(duree, "hh:nn:ss")
End Sub
```

## Comparaison des trois fonctions

### Tableau récapitulatif

| Fonction | Date | Heure | Usage typique |
|----------|------|-------|---------------|
| `Now()` | ✓ | ✓ | Horodatage complet, mesure de temps |
| `Date()` | ✓ | ✗ | Calculs sur les jours, comparaisons de dates |
| `Time()` | ✗ | ✓ | Contrôles horaires, mesures de durée |

### Exemple comparatif

```vba
Sub ComparerFonctions()
    Debug.Print "=== Comparaison des fonctions ==="
    Debug.Print "Now()  : " & Now
    Debug.Print "Date() : " & Date
    Debug.Print "Time() : " & Time
    Debug.Print ""

    ' Formatage pour plus de clarté
    Debug.Print "Now() formaté  : " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Debug.Print "Date() formaté : " & Format(Date, "dd/mm/yyyy")
    Debug.Print "Time() formaté : " & Format(Time, "hh:nn:ss")
End Sub
```

## Bonnes pratiques

### 1. Choisir la bonne fonction

```vba
' Bon : utiliser Date() pour des comparaisons de dates
If Date >= #1/1/2024# Then
    ' Logique pour l'année 2024 et après
End If

' Moins bon : utiliser Now() quand on n'a pas besoin de l'heure
If Now >= #1/1/2024# Then
    ' Fonctionne mais moins précis conceptuellement
End If
```

### 2. Stockage pour éviter les variations

```vba
' Bon : capturer une seule fois au début
Sub BonnePratique()
    Dim momentExecution As Date
    momentExecution = Now

    ' Utiliser momentExecution dans tout le code
    Range("A1").Value = "Début : " & momentExecution
    ' ... autres traitements ...
    Range("A2").Value = "Fin : " & momentExecution  ' Même valeur !
End Sub

' Moins bon : appeler Now() plusieurs fois
Sub MoinsBonnePratique()
    Range("A1").Value = "Début : " & Now
    ' ... autres traitements ...
    Range("A2").Value = "Fin : " & Now  ' Valeur différente !
End Sub
```

### 3. Formatage approprié

```vba
' Utiliser Format() pour un contrôle précis de l'affichage
Dim maintenant As Date
maintenant = Now

' Différents formats selon le besoin
Debug.Print Format(maintenant, "dd/mm/yyyy")         ' 15/03/2024
Debug.Print Format(maintenant, "dddd dd mmmm")       ' Vendredi 15 mars
Debug.Print Format(maintenant, "hh:nn")              ' 14:32
Debug.Print Format(maintenant, "yyyy-mm-dd hh:nn:ss") ' 2024-03-15 14:32:17
```

## Cas d'usage courants

### 1. Horodatage de documents

```vba
Sub HorodaterDocument()
    With ActiveSheet
        .Range("A1").Value = "Rapport généré le : " & Format(Now, "dd/mm/yyyy à hh:nn")
        .Range("A2").Value = "Données au : " & Format(Date, "dd/mm/yyyy")
    End With
End Sub
```

### 2. Contrôles temporels

```vba
Sub ControlerAcces()
    ' Accès autorisé seulement en semaine entre 8h et 18h
    If Weekday(Date, vbMonday) > 5 Then  ' Weekend
        MsgBox "Accès non autorisé le weekend"
        Exit Sub
    End If

    If Time < #8:00 AM# Or Time > #6:00 PM# Then
        MsgBox "Accès non autorisé en dehors des heures ouvrables"
        Exit Sub
    End If

    MsgBox "Accès autorisé"
End Sub
```

### 3. Calculs de durée

```vba
Sub CalculerDureeSession()
    Static debutSession As Date

    If debutSession = 0 Then
        ' Première exécution : mémoriser l'heure de début
        debutSession = Now
        MsgBox "Session démarrée à " & Format(debutSession, "hh:nn:ss")
    Else
        ' Exécutions suivantes : calculer la durée
        Dim duree As Date
        duree = Now - debutSession
        MsgBox "Durée de la session : " & Format(duree, "hh:nn:ss")
    End If
End Sub
```

## Points importants à retenir

**Simplicité** : Ces trois fonctions ne prennent aucun paramètre et sont très simples à utiliser.

**Temps système** : Elles retournent toutes l'heure du système où s'exécute le code, pas l'heure d'un serveur distant.

**Précision** : La précision dépend du système, généralement à la seconde près.

**Stabilité** : Pour des calculs précis, capturez la valeur une seule fois au début de votre procédure.

**Format d'affichage** : Le format d'affichage dépend des paramètres régionaux du système, utilisez `Format()` pour un contrôle précis.

---

*Ces trois fonctions sont les piliers de la gestion du temps en VBA. Leur maîtrise vous permettra de créer des applications temporellement intelligentes et robustes.*

⏭️
