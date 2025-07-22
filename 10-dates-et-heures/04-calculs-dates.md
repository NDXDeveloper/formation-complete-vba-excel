🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.4. Calculs sur les dates

## Introduction aux calculs de dates

Les calculs sur les dates sont l'une des fonctionnalités les plus puissantes de VBA. Grâce au stockage numérique des dates, vous pouvez facilement effectuer des opérations arithmétiques pour calculer des durées, ajouter ou soustraire du temps, et résoudre des problèmes temporels complexes.

## Principe de base : les dates sont des nombres

Rappelons que VBA stocke les dates sous forme de nombres décimaux :
- La partie entière = nombre de jours depuis le 1er janvier 1900
- La partie décimale = fraction de la journée (heure)

Cette représentation rend les calculs très simples et intuitifs.

## Opérations arithmétiques de base

### Addition et soustraction de jours

```vba
Dim dateDebut As Date
Dim dateFin As Date

dateDebut = #1/15/2024#

' Ajouter des jours
dateFin = dateDebut + 7           ' Ajouter 7 jours
Debug.Print dateFin               ' 22/01/2024

' Soustraire des jours
dateFin = dateDebut - 3           ' Soustraire 3 jours
Debug.Print dateFin               ' 12/01/2024

' Calculs avec des variables
Dim nbJours As Integer
nbJours = 30
dateFin = dateDebut + nbJours     ' Ajouter 30 jours
```

### Calcul de différences entre dates

```vba
Dim date1 As Date
Dim date2 As Date
Dim difference As Long

date1 = #1/15/2024#
date2 = #1/22/2024#

' Calculer la différence en jours
difference = date2 - date1
Debug.Print difference            ' 7 (jours)

' Calculer l'âge en jours
Dim naissance As Date
Dim agejours As Long

naissance = #6/15/1990#
agejours = Date - naissance
Debug.Print "Âge en jours : " & agejours
```

### Addition et soustraction d'heures

```vba
Dim dateHeure As Date
Dim nouvelleHeure As Date

dateHeure = #1/15/2024 2:30 PM#

' Ajouter des heures (1 heure = 1/24 de jour)
nouvelleHeure = dateHeure + (3 / 24)    ' Ajouter 3 heures
Debug.Print Format(nouvelleHeure, "dd/mm/yyyy hh:nn")  ' 15/01/2024 17:30

' Méthode plus lisible avec TimeValue
nouvelleHeure = dateHeure + TimeValue("03:00:00")      ' Ajouter 3 heures
nouvelleHeure = dateHeure + TimeValue("00:30:00")      ' Ajouter 30 minutes
nouvelleHeure = dateHeure + TimeValue("00:00:45")      ' Ajouter 45 secondes
```

## Fonctions spécialisées pour les calculs de dates

### DateAdd() - Ajouter des intervalles de temps

La fonction `DateAdd()` est l'outil le plus puissant pour ajouter ou soustraire des intervalles de temps spécifiques.

#### Syntaxe
```vba
DateAdd(interval, number, date)
```

- **interval** : type d'intervalle (chaîne de caractères)
- **number** : nombre d'unités à ajouter (peut être négatif)
- **date** : date de départ

#### Intervalles disponibles

| Intervalle | Description | Exemple |
|------------|-------------|---------|
| "yyyy" | Années | Ajouter 2 ans |
| "q" | Trimestres | Ajouter 1 trimestre |
| "m" | Mois | Ajouter 3 mois |
| "y" | Jour de l'année | Ajouter 100 jours |
| "d" | Jours | Ajouter 15 jours |
| "w" | Jours de la semaine | Ajouter 2 semaines |
| "ww" | Semaines | Ajouter 4 semaines |
| "h" | Heures | Ajouter 6 heures |
| "n" | Minutes | Ajouter 30 minutes |
| "s" | Secondes | Ajouter 45 secondes |

#### Exemples pratiques

```vba
Dim dateBase As Date
Dim resultat As Date

dateBase = #1/15/2024#

' Ajouter des années
resultat = DateAdd("yyyy", 2, dateBase)
Debug.Print Format(resultat, "dd/mm/yyyy")    ' 15/01/2026

' Ajouter des mois
resultat = DateAdd("m", 6, dateBase)
Debug.Print Format(resultat, "dd/mm/yyyy")    ' 15/07/2024

' Soustraire des jours (nombre négatif)
resultat = DateAdd("d", -10, dateBase)
Debug.Print Format(resultat, "dd/mm/yyyy")    ' 05/01/2024

' Ajouter des heures à une date avec heure
dateBase = #1/15/2024 2:30 PM#
resultat = DateAdd("h", 5, dateBase)
Debug.Print Format(resultat, "dd/mm/yyyy hh:nn")  ' 15/01/2024 19:30
```

### DateDiff() - Calculer la différence entre deux dates

La fonction `DateDiff()` calcule la différence entre deux dates dans l'unité de votre choix.

#### Syntaxe
```vba
DateDiff(interval, date1, date2, [firstdayofweek], [firstweekofyear])
```

#### Exemples d'utilisation

```vba
Dim date1 As Date
Dim date2 As Date
Dim difference As Long

date1 = #1/15/2024#
date2 = #6/20/2024#

' Différence en jours
difference = DateDiff("d", date1, date2)
Debug.Print difference & " jours"             ' 157 jours

' Différence en mois
difference = DateDiff("m", date1, date2)
Debug.Print difference & " mois"              ' 5 mois

' Différence en années
difference = DateDiff("yyyy", date1, date2)
Debug.Print difference & " années"            ' 0 années

' Différence en semaines
difference = DateDiff("ww", date1, date2)
Debug.Print difference & " semaines"          ' 22 semaines
```

### Calculs d'âge précis

```vba
Function CalculerAge(dateNaissance As Date, Optional dateReference As Date) As Integer
    ' Si pas de date de référence, utiliser aujourd'hui
    If dateReference = 0 Then dateReference = Date

    ' Calculer l'âge en années
    Dim age As Integer
    age = DateDiff("yyyy", dateNaissance, dateReference)

    ' Ajuster si l'anniversaire n'est pas encore passé cette année
    If DateAdd("yyyy", age, dateNaissance) > dateReference Then
        age = age - 1
    End If

    CalculerAge = age
End Function

' Utilisation
Dim naissance As Date
naissance = #6/15/1990#
Debug.Print "Âge : " & CalculerAge(naissance) & " ans"
```

## Calculs complexes avec les dates

### Calcul du prochain jour ouvrable

```vba
Function ProchainJourOuvrable(dateDepart As Date) As Date
    Dim resultat As Date
    resultat = dateDepart + 1

    ' Passer les weekends
    Do While Weekday(resultat, vbMonday) > 5  ' 6=Samedi, 7=Dimanche
        resultat = resultat + 1
    Loop

    ProchainJourOuvrable = resultat
End Function

' Utilisation
Dim vendredi As Date
vendredi = #1/19/2024#  ' Un vendredi
Debug.Print Format(ProchainJourOuvrable(vendredi), "dddd dd/mm/yyyy")
' Résultat : Lundi 22/01/2024
```

### Calcul du nombre de jours ouvrables

```vba
Function NombreJoursOuvrables(dateDebut As Date, dateFin As Date) As Integer
    Dim compteur As Integer
    Dim dateCourante As Date

    compteur = 0
    dateCourante = dateDebut

    Do While dateCourante <= dateFin
        ' Vérifier si c'est un jour ouvrable (Lundi=1 à Vendredi=5)
        If Weekday(dateCourante, vbMonday) <= 5 Then
            compteur = compteur + 1
        End If
        dateCourante = dateCourante + 1
    Loop

    NombreJoursOuvrables = compteur
End Function

' Utilisation
Dim debut As Date
Dim fin As Date
debut = #1/15/2024#  ' Lundi
fin = #1/21/2024#    ' Dimanche

Debug.Print "Jours ouvrables : " & NombreJoursOuvrables(debut, fin)
' Résultat : 5 jours ouvrables
```

### Calcul de la fin de mois

```vba
Function FinDuMois(dateReference As Date) As Date
    ' Aller au premier jour du mois suivant, puis reculer d'un jour
    Dim premierJourMoisSuivant As Date
    premierJourMoisSuivant = DateSerial(Year(dateReference), Month(dateReference) + 1, 1)
    FinDuMois = premierJourMoisSuivant - 1
End Function

' Utilisation
Dim uneDate As Date
uneDate = #1/15/2024#
Debug.Print Format(FinDuMois(uneDate), "dd/mm/yyyy")
' Résultat : 31/01/2024
```

## Calculs avec les heures et minutes

### Addition de temps précis

```vba
Sub CalculsTemps()
    Dim heureDebut As Date
    Dim duree As Date
    Dim heureFin As Date

    heureDebut = #9:00 AM#

    ' Ajouter 2 heures et 30 minutes
    duree = TimeValue("02:30:00")
    heureFin = heureDebut + duree

    Debug.Print "Début : " & Format(heureDebut, "hh:nn")
    Debug.Print "Fin : " & Format(heureFin, "hh:nn")
    ' Début : 09:00
    ' Fin : 11:30
End Sub
```

### Calcul de durée entre deux heures

```vba
Function CalculerDuree(heureDebut As Date, heureFin As Date) As Date
    ' Gérer le passage de minuit
    If heureFin < heureDebut Then
        ' La fin est le lendemain
        CalculerDuree = (heureFin + 1) - heureDebut
    Else
        CalculerDuree = heureFin - heureDebut
    End If
End Function

' Utilisation
Dim debut As Date
Dim fin As Date
Dim duree As Date

debut = #9:00 AM#
fin = #5:30 PM#

duree = CalculerDuree(debut, fin)
Debug.Print "Durée : " & Format(duree, "h:nn")
' Résultat : 8:30
```

## Gestion des cas particuliers

### Années bissextiles

```vba
Function EstBissextile(annee As Integer) As Boolean
    ' Une année est bissextile si :
    ' - Elle est divisible par 4 ET
    ' - (Elle n'est pas divisible par 100 OU elle est divisible par 400)
    EstBissextile = (annee Mod 4 = 0) And ((annee Mod 100 <> 0) Or (annee Mod 400 = 0))
End Function

' Utilisation
Debug.Print "2024 est bissextile : " & EstBissextile(2024)  ' True
Debug.Print "2023 est bissextile : " & EstBissextile(2023)  ' False
```

### Gestion des fins de mois variables

```vba
Function AjouterMoisSecurise(dateDepart As Date, nbMois As Integer) As Date
    ' Utiliser DateAdd qui gère automatiquement les fins de mois
    AjouterMoisSecurise = DateAdd("m", nbMois, dateDepart)

    ' Exemple : 31/01/2024 + 1 mois = 29/02/2024 (année bissextile)
    '          31/01/2023 + 1 mois = 28/02/2023 (année normale)
End Function
```

## Fonctions utilitaires pour les calculs

### Début et fin de période

```vba
Function DebutSemaine(dateReference As Date) As Date
    ' Trouver le lundi de la semaine
    Dim jourSemaine As Integer
    jourSemaine = Weekday(dateReference, vbMonday)  ' 1=Lundi, 7=Dimanche
    DebutSemaine = dateReference - (jourSemaine - 1)
End Function

Function DebutMois(dateReference As Date) As Date
    DebutMois = DateSerial(Year(dateReference), Month(dateReference), 1)
End Function

Function DebutAnnee(dateReference As Date) As Date
    DebutAnnee = DateSerial(Year(dateReference), 1, 1)
End Function

' Utilisation
Dim uneDate As Date
uneDate = #1/17/2024#  ' Mercredi

Debug.Print "Date : " & Format(uneDate, "dddd dd/mm/yyyy")
Debug.Print "Début semaine : " & Format(DebutSemaine(uneDate), "dddd dd/mm/yyyy")
Debug.Print "Début mois : " & Format(DebutMois(uneDate), "dddd dd/mm/yyyy")
Debug.Print "Début année : " & Format(DebutAnnee(uneDate), "dddd dd/mm/yyyy")
```

## Calculs avancés pour applications métier

### Calcul d'échéances

```vba
Function CalculerEcheances(dateDebut As Date, nbEcheances As Integer, _
                          intervalleEcheances As String) As Date()

    Dim echeances() As Date
    ReDim echeances(1 To nbEcheances)

    Dim i As Integer
    For i = 1 To nbEcheances
        echeances(i) = DateAdd(intervalleEcheances, i, dateDebut)
    Next i

    CalculerEcheances = echeances
End Function

' Utilisation
Sub AfficherEcheances()
    Dim dateContrat As Date
    Dim echeances() As Date
    Dim i As Integer

    dateContrat = #1/15/2024#
    echeances = CalculerEcheances(dateContrat, 6, "m")  ' 6 échéances mensuelles

    Debug.Print "Échéances du contrat :"
    For i = 1 To UBound(echeances)
        Debug.Print "Échéance " & i & " : " & Format(echeances(i), "dd/mm/yyyy")
    Next i
End Sub
```

### Calcul de performance temporelle

```vba
Function MesurerPerformance() As String
    Static tempsDebut As Date

    If tempsDebut = 0 Then
        ' Premier appel : démarrer le chrono
        tempsDebut = Timer / 86400 + Date  ' Timer en secondes, convertir en fraction de jour
        MesurerPerformance = "Chronomètre démarré"
    Else
        ' Appel suivant : calculer la durée
        Dim tempsFin As Date
        Dim duree As Date

        tempsFin = Timer / 86400 + Date
        duree = tempsFin - tempsDebut

        tempsDebut = 0  ' Réinitialiser
        MesurerPerformance = "Durée : " & Format(duree, "hh:nn:ss")
    End If
End Function

' Utilisation
Sub TestPerformance()
    Debug.Print MesurerPerformance()  ' Démarrer

    ' Code à mesurer
    Application.Wait Now + TimeValue("00:00:02")  ' Attendre 2 secondes

    Debug.Print MesurerPerformance()  ' Arrêter et afficher
End Sub
```

## Conseils et bonnes pratiques

### 1. Attention aux calculs avec les heures

```vba
' Bon : utiliser des fonctions spécialisées
Dim resultat As Date
resultat = DateAdd("h", 2, Now)  ' Ajouter 2 heures

' Moins bon : calcul manuel (risque d'erreur)
resultat = Now + (2 / 24)  ' Peut créer des imprécisions
```

### 2. Gérer les cas limites

```vba
Function AjouterJoursOuvrables(dateDepart As Date, nbJours As Integer) As Date
    Dim resultat As Date
    Dim joursAjoutes As Integer

    resultat = dateDepart
    joursAjoutes = 0

    Do While joursAjoutes < nbJours
        resultat = resultat + 1
        ' Compter seulement les jours ouvrables
        If Weekday(resultat, vbMonday) <= 5 Then
            joursAjoutes = joursAjoutes + 1
        End If
    Loop

    AjouterJoursOuvrables = resultat
End Function
```

### 3. Valider les paramètres

```vba
Function CalculerDifferenceSecurise(date1 As Date, date2 As Date, unite As String) As Long
    ' Vérifier que les dates sont valides
    If date1 = 0 Or date2 = 0 Then
        CalculerDifferenceSecurise = -1  ' Erreur
        Exit Function
    End If

    ' Vérifier l'unité
    Select Case LCase(unite)
        Case "d", "m", "yyyy", "h", "n", "s"
            CalculerDifferenceSecurise = DateDiff(unite, date1, date2)
        Case Else
            CalculerDifferenceSecurise = -1  ' Unité invalide
    End Select
End Function
```

## Points importants à retenir

**Simplicité arithmétique** : Les dates étant des nombres, les calculs de base (+ et -) sont très simples.

**Fonctions spécialisées** : `DateAdd()` et `DateDiff()` gèrent automatiquement les complexités (mois variables, années bissextiles).

**Précision** : Attention aux calculs avec les heures - préférer les fonctions VBA aux calculs manuels.

**Validation** : Toujours vérifier la validité des dates avant les calculs.

**Cas particuliers** : Penser aux weekends, jours fériés, fins de mois variables selon vos besoins métier.

**Performance** : Pour des calculs répétitifs, optimiser avec des boucles et des fonctions appropriées.

---

*Les calculs de dates sont au cœur de nombreuses applications métier. Une maîtrise de ces techniques vous permettra de résoudre facilement des problèmes temporels complexes et de créer des fonctionnalités sophistiquées.*

⏭️
