🔝 Retour au [Sommaire](/SOMMAIRE.md)

# E. Glossaire des termes techniques

## Introduction

Ce glossaire définit tous les termes techniques VBA que vous rencontrerez dans cette formation. Chaque définition est volontairement simple et accompagnée d'exemples concrets pour faciliter votre compréhension. Les termes sont classés par ordre alphabétique pour une recherche rapide.

**Comment utiliser ce glossaire :**
- **Définition simple** : Explication en langage courant
- **Exemple** : Illustration concrète quand c'est utile
- **Voir aussi** : Renvois vers des termes liés
- **★** : Termes essentiels à connaître absolument

---

## A

### ActiveCell ★
**Définition :** La cellule actuellement sélectionnée dans Excel  
**Exemple :** `ActiveCell.Value = "Bonjour"` écrit "Bonjour" dans la cellule sélectionnée  
**Voir aussi :** Range, Selection  

### ActiveSheet ★
**Définition :** La feuille de calcul actuellement affichée et active  
**Exemple :** `ActiveSheet.Name` vous donne le nom de la feuille active  
**Voir aussi :** Worksheet, Workbook  

### ActiveWorkbook ★
**Définition :** Le classeur Excel actuellement ouvert et actif  
**Exemple :** `ActiveWorkbook.Save` sauvegarde le classeur actuel  
**Voir aussi :** Workbook, ThisWorkbook  

### ADO (ActiveX Data Objects)
**Définition :** Technologie Microsoft pour accéder aux bases de données depuis VBA  
**Exemple :** Permet de connecter Excel à une base SQL Server  
**Voir aussi :** Base de données, SQL  

### API (Application Programming Interface)
**Définition :** Interface qui permet à VBA de communiquer avec Windows ou d'autres logiciels  
**Exemple :** API pour connaître le nom de l'utilisateur Windows  
**Voir aussi :** DLL, Declare  

### Application ★
**Définition :** L'objet qui représente Excel lui-même dans VBA  
**Exemple :** `Application.Quit` ferme Excel complètement  
**Voir aussi :** Object, Excel  

### Argument ★
**Définition :** Information que vous donnez à une fonction ou procédure  
**Exemple :** Dans `Left("Bonjour", 3)`, "Bonjour" et 3 sont des arguments  
**Voir aussi :** Paramètre, Function, Sub  

### Array
**Définition :** Variable qui peut contenir plusieurs valeurs à la fois  
**Exemple :** `Dim nombres(1 To 10) As Integer` crée un tableau de 10 nombres  
**Voir aussi :** Tableau, Variable  

---

## B

### Boolean ★
**Définition :** Type de variable qui ne peut être que True (vrai) ou False (faux)  
**Exemple :** `Dim estValide As Boolean` puis `estValide = True`  
**Voir aussi :** Variable, True, False  

### Breakpoint
**Définition :** Point d'arrêt dans le code pour déboguer  
**Exemple :** F9 sur une ligne pour que l'exécution s'arrête à cette ligne  
**Voir aussi :** Debug, F9  

### ByRef
**Définition :** Façon de passer une variable à une procédure en permettant sa modification  
**Exemple :** `Sub Test(ByRef nombre As Integer)` - la procédure peut changer 'nombre'  
**Voir aussi :** ByVal, Paramètre  

### ByVal ★
**Définition :** Façon de passer une variable à une procédure sans permettre sa modification  
**Exemple :** `Sub Test(ByVal nombre As Integer)` - 'nombre' ne sera pas modifié  
**Voir aussi :** ByRef, Paramètre  

---

## C

### Call ★
**Définition :** Mot-clé pour exécuter une procédure depuis une autre  
**Exemple :** `Call MaProcedure()` exécute la procédure nommée MaProcedure  
**Voir aussi :** Sub, Procedure  

### Cells ★
**Définition :** Façon de référencer une cellule par numéro de ligne et colonne  
**Exemple :** `Cells(1, 1)` désigne la cellule A1 (ligne 1, colonne 1)  
**Voir aussi :** Range, ActiveCell  

### Class
**Définition :** Modèle pour créer des objets personnalisés  
**Exemple :** Créer une classe "Produit" avec propriétés nom et prix  
**Voir aussi :** Object, Module de classe  

### Collection ★
**Définition :** Groupe d'objets du même type  
**Exemple :** `Worksheets` est une collection de toutes les feuilles  
**Voir aussi :** Object, For Each  

### Commentaire ★
**Définition :** Texte dans le code qui explique mais n'est pas exécuté  
**Exemple :** `' Ceci est un commentaire` (commence par une apostrophe)  
**Voir aussi :** Documentation  

### Compilation
**Définition :** Vérification de la syntaxe du code par VBA  
**Exemple :** Debug → Compiler le projet vérifie tout votre code  
**Voir aussi :** Erreur de compilation, Debug  

### Constante ★
**Définition :** Variable dont la valeur ne change jamais  
**Exemple :** `Const TVA As Double = 0.2` définit un taux de TVA fixe  
**Voir aussi :** Variable, Const  

---

## D

### Date ★
**Définition :** Type de variable pour stocker des dates et heures  
**Exemple :** `Dim aujourdhui As Date` puis `aujourdhui = Date()`  
**Voir aussi :** Now, Time, Variable  

### Debug ★
**Définition :** Processus de recherche et correction des erreurs  
**Exemple :** F8 pour exécuter ligne par ligne et déboguer  
**Voir aussi :** Breakpoint, F8, Error  

### Declare
**Définition :** Mot-clé pour utiliser des fonctions Windows (API)  
**Exemple :** `Declare Function GetUserName Lib "advapi32.dll"`  
**Voir aussi :** API, DLL  

### Dim ★
**Définition :** Mot-clé pour déclarer (créer) une variable  
**Exemple :** `Dim monNom As String` crée une variable texte  
**Voir aussi :** Variable, As  

### Do Loop ★
**Définition :** Boucle qui répète du code tant qu'une condition est vraie  
**Exemple :** `Do While i < 10 ... Loop` répète tant que i est inférieur à 10  
**Voir aussi :** While, Loop, For  

### Double ★
**Définition :** Type de variable pour les nombres décimaux  
**Exemple :** `Dim prix As Double` puis `prix = 19.99`  
**Voir aussi :** Integer, Single, Variable  

---

## E

### End ★
**Définition :** Mot-clé qui termine une structure (If, Sub, Function, etc.)  
**Exemple :** `End If` termine un bloc If, `End Sub` termine une procédure  
**Voir aussi :** If, Sub, Function  

### Error ★
**Définition :** Problème qui empêche le code de fonctionner normalement  
**Exemple :** "Type incompatible" est une erreur fréquente  
**Voir aussi :** Debug, On Error, Err  

### Event ★
**Définition :** Action qui déclenche automatiquement du code  
**Exemple :** L'ouverture d'un classeur déclenche l'événement Workbook_Open  
**Voir aussi :** Workbook_Open, Worksheet_Change  

### Exit ★
**Définition :** Mot-clé pour sortir d'une procédure ou boucle  
**Exemple :** `Exit Sub` sort immédiatement de la procédure  
**Voir aussi :** End, Return  

---

## F

### False ★
**Définition :** Valeur booléenne signifiant "faux"  
**Exemple :** `If 5 > 10 Then` donne False car 5 n'est pas supérieur à 10  
**Voir aussi :** True, Boolean, If  

### For ★
**Définition :** Boucle qui répète du code un nombre déterminé de fois  
**Exemple :** `For i = 1 To 10` répète 10 fois  
**Voir aussi :** Next, Loop, Do  

### Function ★
**Définition :** Bloc de code qui calcule et retourne une valeur  
**Exemple :** `Function Ajouter(a, b)` peut retourner a + b  
**Voir aussi :** Sub, Return, End Function  

---

## G

### Global
**Définition :** Variable accessible depuis tout le projet VBA  
**Exemple :** `Global monNom As String` dans un module standard  
**Voir aussi :** Public, Private, Scope  

### GoTo
**Définition :** Instruction qui fait sauter l'exécution à une autre ligne  
**Exemple :** `GoTo FinProcedure` va à la ligne marquée FinProcedure:  
**Voir aussi :** Label, On Error GoTo  

---

## I

### If ★
**Définition :** Structure pour exécuter du code seulement si une condition est vraie  
**Exemple :** `If age >= 18 Then ... End If` s'exécute si age ≥ 18  
**Voir aussi :** Then, Else, End If  

### InputBox ★
**Définition :** Boîte de dialogue qui demande une saisie à l'utilisateur  
**Exemple :** `nom = InputBox("Votre nom?")` demande le nom  
**Voir aussi :** MsgBox, UserForm  

### Integer ★
**Définition :** Type de variable pour les nombres entiers  
**Exemple :** `Dim age As Integer` puis `age = 25`  
**Voir aussi :** Long, Double, Variable  

---

## L

### Long ★
**Définition :** Type de variable pour les grands nombres entiers  
**Exemple :** `Dim population As Long` puis `population = 67000000`  
**Voir aussi :** Integer, Double, Variable  

### Loop ★
**Définition :** Fin d'une boucle Do  
**Exemple :** `Do ... Loop While condition` répète tant que condition est vraie  
**Voir aussi :** Do, While, Until  

---

## M

### Macro ★
**Définition :** Nom courant pour une procédure VBA  
**Exemple :** "J'ai créé une macro pour automatiser ce calcul"  
**Voir aussi :** Sub, Procedure  

### Méthode ★
**Définition :** Action qu'un objet peut effectuer  
**Exemple :** `Range("A1").Copy` - Copy est une méthode de l'objet Range  
**Voir aussi :** Object, Propriété  

### Module ★
**Définition :** Fichier qui contient du code VBA  
**Exemple :** Module1, Module2 dans l'explorateur de projets  
**Voir aussi :** Sub, Function, Class  

### MsgBox ★
**Définition :** Boîte de dialogue qui affiche un message à l'utilisateur  
**Exemple :** `MsgBox "Bonjour!"` affiche "Bonjour!" à l'écran  
**Voir aussi :** InputBox, vbOK, vbYesNo  

---

## N

### Next ★
**Définition :** Fin d'une boucle For  
**Exemple :** `For i = 1 To 10 ... Next i` - Next termine la boucle  
**Voir aussi :** For, Step  

### Nothing ★
**Définition :** Valeur spéciale qui signifie "aucun objet"  
**Exemple :** `Set monObjet = Nothing` libère la référence à l'objet  
**Voir aussi :** Set, Object, Is Nothing  

### Now ★
**Définition :** Fonction qui retourne la date et l'heure actuelles  
**Exemple :** `maintenant = Now()` donne la date/heure actuelle  
**Voir aussi :** Date, Time  

---

## O

### Object ★
**Définition :** "Chose" dans Excel que VBA peut manipuler  
**Exemple :** Une cellule, une feuille, un classeur sont des objets  
**Voir aussi :** Propriété, Méthode, Class  

### On Error ★
**Définition :** Instruction pour gérer les erreurs  
**Exemple :** `On Error Resume Next` continue malgré les erreurs  
**Voir aussi :** Error, Resume, GoTo  

### Optional
**Définition :** Paramètre qu'on peut omettre lors de l'appel d'une fonction  
**Exemple :** `Function Test(nom As String, Optional age As Integer = 0)`  
**Voir aussi :** Parameter, Default  

---

## P

### Paramètre ★
**Définition :** Information qu'on donne à une procédure ou fonction  
**Exemple :** Dans `Left("Bonjour", 3)`, "Bonjour" et 3 sont des paramètres  
**Voir aussi :** Argument, Function, Sub  

### Private ★
**Définition :** Variable ou procédure accessible seulement dans le module actuel  
**Exemple :** `Private Sub MaProcedure()` n'est visible que dans ce module  
**Voir aussi :** Public, Scope  

### Procédure ★
**Définition :** Bloc de code qui effectue une tâche  
**Exemple :** Une Sub ou Function est une procédure  
**Voir aussi :** Sub, Function, Macro  

### Propriété ★
**Définition :** Caractéristique d'un objet qu'on peut lire ou modifier  
**Exemple :** `Range("A1").Value` - Value est une propriété de Range  
**Voir aussi :** Object, Méthode  

### Public ★
**Définition :** Variable ou procédure accessible depuis tout le projet  
**Exemple :** `Public Const TVA = 0.2` est accessible partout  
**Voir aussi :** Private, Global, Scope  

---

## R

### Range ★
**Définition :** Objet qui représente une ou plusieurs cellules  
**Exemple :** `Range("A1:C3")` représente un bloc de 9 cellules  
**Voir aussi :** Cells, ActiveCell  

### ReDim
**Définition :** Redimensionner un tableau dynamique  
**Exemple :** `ReDim monTableau(1 To 20)` change la taille du tableau  
**Voir aussi :** Array, Dim  

### Resume
**Définition :** Continue l'exécution après une erreur  
**Exemple :** `Resume Next` continue à la ligne suivante  
**Voir aussi :** On Error, Error  

### Return
**Définition :** Valeur qu'une fonction donne en résultat  
**Exemple :** `Function Doubler(x) ... Doubler = x * 2` retourne x*2  
**Voir aussi :** Function, End Function  

---

## S

### Scope
**Définition :** Portée d'une variable (où elle est accessible)  
**Exemple :** Variable Private = scope du module, Public = scope du projet  
**Voir aussi :** Private, Public, Dim  

### Select Case ★
**Définition :** Structure pour choisir entre plusieurs options  
**Exemple :** `Select Case age` puis `Case 0 To 17`, `Case 18 To 65`, etc.  
**Voir aussi :** Case, If, End Select  

### Set ★
**Définition :** Mot-clé pour assigner un objet à une variable  
**Exemple :** `Set maFeuille = Worksheets("Feuil1")`  
**Voir aussi :** Object, Nothing  

### Single
**Définition :** Type de variable pour nombres décimaux (moins précis que Double)  
**Exemple :** `Dim temperature As Single`  
**Voir aussi :** Double, Integer, Variable  

### String ★
**Définition :** Type de variable pour stocker du texte  
**Exemple :** `Dim message As String` puis `message = "Bonjour"`  
**Voir aussi :** Variable, Text  

### Sub ★
**Définition :** Procédure qui effectue des actions mais ne retourne pas de valeur  
**Exemple :** `Sub DireBonjour()` exécute des actions  
**Voir aussi :** Function, Procedure, End Sub  

---

## T

### Then ★
**Définition :** Partie d'une instruction If qui s'exécute si la condition est vraie  
**Exemple :** `If age >= 18 Then MsgBox "Majeur"`  
**Voir aussi :** If, Else, End If  

### ThisWorkbook ★
**Définition :** Le classeur qui contient le code VBA  
**Exemple :** `ThisWorkbook.Save` sauvegarde le classeur avec les macros  
**Voir aussi :** ActiveWorkbook, Workbook  

### True ★
**Définition :** Valeur booléenne signifiant "vrai"  
**Exemple :** `If 5 > 3 Then` donne True car 5 est supérieur à 3  
**Voir aussi :** False, Boolean, If  

### Type
**Définition :** Catégorie de données (String, Integer, Date, etc.)  
**Exemple :** `Dim nom As String` - String est le type de la variable nom  
**Voir aussi :** Variable, String, Integer  

---

## U

### UDF (User Defined Function)
**Définition :** Fonction personnalisée créée par l'utilisateur  
**Exemple :** Créer une fonction pour calculer la TVA  
**Voir aussi :** Function, Custom  

### UserForm
**Définition :** Fenêtre personnalisée avec boutons, zones de texte, etc.  
**Exemple :** Formulaire de saisie des données client  
**Voir aussi :** InputBox, MsgBox, Interface  

### Until
**Définition :** Condition de fin dans une boucle Do  
**Exemple :** `Do ... Loop Until i = 10` répète jusqu'à ce que i égale 10  
**Voir aussi :** Do, Loop, While  

---

## V

### Variable ★
**Définition :** "Boîte" qui stocke une valeur qu'on peut utiliser et modifier  
**Exemple :** `Dim age As Integer` crée une variable pour stocker un âge  
**Voir aussi :** Dim, Type, String, Integer  

### Variant ★
**Définition :** Type de variable qui peut contenir n'importe quoi  
**Exemple :** `Dim donnee As Variant` peut contenir texte, nombre, date, etc.  
**Voir aussi :** Variable, Type  

### VBA ★
**Définition :** Visual Basic for Applications - le langage de programmation d'Excel  
**Exemple :** "J'apprends VBA pour automatiser Excel"  

### VBE (Visual Basic Editor) ★
**Définition :** L'environnement où on écrit le code VBA  
**Exemple :** Alt+F11 ouvre le VBE depuis Excel  
**Voir aussi :** Editor, Alt+F11  

---

## W

### While ★
**Définition :** Condition de continuation dans une boucle  
**Exemple :** `Do While i < 10` répète tant que i est inférieur à 10  
**Voir aussi :** Do, Loop, Until  

### With ★
**Définition :** Structure pour simplifier les références à un objet  
**Exemple :** `With Range("A1") ... .Value = "Test" ... End With`  
**Voir aussi :** End With, Object  

### Workbook ★
**Définition :** Classeur Excel (fichier .xlsx)  
**Exemple :** `Workbooks("MonFichier.xlsx")` référence un classeur spécifique  
**Voir aussi :** ActiveWorkbook, ThisWorkbook  

### Worksheet ★
**Définition :** Feuille de calcul dans un classeur  
**Exemple :** `Worksheets("Feuil1")` référence la première feuille  
**Voir aussi :** ActiveSheet, Sheet  

---

## Conseils d'utilisation du glossaire

### Pour les débutants :
1. **Commencez par les termes ★** : Ce sont les plus importants
2. **Lisez les exemples** : Ils rendent les définitions concrètes
3. **Suivez les "Voir aussi"** : Pour comprendre les liens entre concepts
4. **Revenez souvent** : Normal de ne pas tout retenir d'un coup

### Progression suggérée :
**Semaine 1** : Variable, Dim, String, Integer, Sub, Range  
**Semaine 2** : If, For, Object, Propriété, Méthode  
**Semaine 3** : Error, Debug, Function, Boolean  
**Mois suivant** : Concepts plus avancés selon vos besoins  

### Comment bien utiliser ce glossaire :
- **Marquez vos favoris** : Surlignez les termes que vous utilisez souvent
- **Créez vos exemples** : Adaptez les exemples à vos cas d'usage
- **Testez dans VBA** : Expérimentez avec les exemples
- **Partagez avec d'autres** : Enseigner aide à mémoriser

**Rappelez-vous :** Il est normal de ne pas connaître tous ces termes au début. Ce glossaire est là pour vous accompagner tout au long de votre apprentissage VBA !

⏭️
