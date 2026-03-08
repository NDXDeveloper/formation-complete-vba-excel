🔝 Retour au [Sommaire](/SOMMAIRE.md)

# C. Raccourcis clavier de l'éditeur VBA

## Introduction

Connaître les raccourcis clavier de l'éditeur VBA vous fera gagner beaucoup de temps et rendra votre programmation plus fluide. Cette annexe présente les raccourcis les plus utiles, organisés par catégorie et par niveau de priorité pour les débutants.

**Comment utiliser cette référence :**
- **★★★** : Raccourcis essentiels à apprendre en premier
- **★★☆** : Raccourcis très utiles pour améliorer votre efficacité
- **★☆☆** : Raccourcis avancés pour les utilisateurs expérimentés
- **Ctrl + Touche** : Maintenez Ctrl enfoncé puis appuyez sur la touche
- **Alt + Touche** : Maintenez Alt enfoncé puis appuyez sur la touche

---

## 1. Raccourcis de base et navigation

### Accès et ouverture ★★★

**Alt + F11** - Ouvrir l'éditeur VBA depuis Excel
*Le raccourci le plus important ! Mémorisez-le en premier*

**Alt + Tab** - Basculer entre l'éditeur VBA et Excel
*Pratique pour tester votre code et voir les résultats*

**Ctrl + R** - Afficher/Masquer l'Explorateur de projets
*Indispensable pour naviguer entre vos modules et feuilles*

**F4** - Afficher/Masquer la fenêtre Propriétés
*Utile quand vous créez des UserForms*

### Navigation dans le code ★★★

**Ctrl + Début** - Aller au début du module
*Retour rapide au sommet de votre code*

**Ctrl + Fin** - Aller à la fin du module
*Aller directement à la fin de votre code*

**Ctrl + G** - Afficher la fenêtre Exécution (Immediate Window)
*Permet de tester des instructions et afficher les résultats de Debug.Print*

**Ctrl + Flèche droite/gauche** - Naviguer mot par mot
*Plus rapide que caractère par caractère*

---

## 2. Exécution et débogage

### Exécution du code ★★★

**F5** - Exécuter la procédure ou continuer l'exécution
*Le raccourci le plus utilisé pour tester votre code !*

**F8** - Exécuter ligne par ligne (pas à pas)
*Essentiel pour déboguer et comprendre ce que fait votre code*

**Shift + F8** - Exécuter procédure par procédure
*Exécute une procédure entière d'un coup lors du débogage*

**Ctrl + Break** - Arrêter l'exécution
*Arrête un code qui tourne en boucle infinie*

### Débogage ★★☆

**F9** - Ajouter/Supprimer un point d'arrêt
*Marque une ligne où l'exécution s'arrêtera automatiquement*

**Ctrl + Shift + F9** - Supprimer tous les points d'arrêt
*Nettoie tous vos points d'arrêt en une fois*

**Ctrl + L** - Afficher la boîte de dialogue "Appels"
*Voir la pile des procédures appelées*

**Ctrl + F8** - Exécuter jusqu'au curseur
*Continue l'exécution jusqu'à la ligne où se trouve votre curseur*

---

## 3. Édition et modification du code

### Édition de base ★★★

**Ctrl + Z** - Annuler la dernière action
*Comme dans tous les logiciels, indispensable !*

**Ctrl + Y** - Refaire l'action annulée
*Annule l'annulation*

**Ctrl + X** - Couper
*Coupe le texte sélectionné*

**Ctrl + C** - Copier
*Copie le texte sélectionné*

**Ctrl + V** - Coller
*Colle le texte dans le presse-papiers*

**Ctrl + A** - Sélectionner tout
*Sélectionne tout le code du module*

### Édition avancée ★★☆

**Tab** - Indenter vers la droite
*Décale le code vers la droite pour l'organiser*

**Shift + Tab** - Indenter vers la gauche
*Décale le code vers la gauche*

---

## 4. Recherche et remplacement

### Recherche ★★★

**Ctrl + F** - Ouvrir la boîte de dialogue Rechercher
*Trouve du texte dans votre code*

**F3** - Rechercher l'occurrence suivante
*Continue la recherche vers le bas*

**Shift + F3** - Rechercher l'occurrence précédente
*Continue la recherche vers le haut*

**Ctrl + H** - Ouvrir la boîte de dialogue Remplacer
*Remplace du texte par autre chose*

### Recherche avancée ★★☆

Pour rechercher ou remplacer dans **tout le projet**, utilisez Ctrl + F ou Ctrl + H puis sélectionnez l'option **« Projet en cours »** dans la boîte de dialogue.

---

## 5. Aide et informations

### Aide contextuelle ★★★

**F1** - Afficher l'aide VBA
*Ouvre l'aide pour l'élément sélectionné*

**Ctrl + I** - Informations rapides
*Affiche des infos sur la fonction sous le curseur*

**Ctrl + J** - Liste des membres
*Affiche les propriétés et méthodes disponibles*

**Ctrl + Espace** - Complétion automatique
*Complète automatiquement ce que vous tapez*

### Inspection du code ★★☆

**F2** - Explorateur d'objets
*Parcourt tous les objets, propriétés et méthodes disponibles*

**Shift + F2** - Aller à la définition
*Va à la définition de la procédure ou variable*

**Ctrl + Shift + F2** - Retour de la définition
*Retourne d'où vous veniez*

---

## 6. Fenêtres et affichage

### Gestion des fenêtres ★★☆

**Ctrl + F6** - Fenêtre suivante
*Passe d'une fenêtre de code à l'autre*

**Ctrl + Shift + F6** - Fenêtre précédente
*Sens inverse du raccourci précédent*

**Ctrl + F4** - Fermer la fenêtre courante
*Ferme le module de code actuel*

### Import/Export ★☆☆

L'import et l'export de modules se font via le menu **Fichier** :
- **Fichier** → **Importer un fichier...** pour ajouter un module (.bas, .cls, .frm)
- **Fichier** → **Exporter un fichier...** pour sauvegarder un module

---

## 7. Raccourcis spéciaux et astuces

### Commentaires ★★☆

Pour commenter ou décommenter plusieurs lignes, utilisez les boutons de la **barre d'outils Édition** :
- Affichez-la via **Affichage** → **Barres d'outils** → **Édition**
- Cliquez sur le bouton **Commenter bloc** ou **Décommenter bloc**
*Il n'existe pas de raccourci clavier natif pour cette action*

### Formatage ★★☆

L'éditeur VBA ne dispose pas de raccourci d'indentation automatique. Utilisez **Tab** et **Shift + Tab** pour ajuster manuellement l'indentation de vos lignes sélectionnées. La barre d'outils Édition propose aussi les boutons **Retrait** et **Retrait négatif**.

---

## 8. Conseils pour bien utiliser les raccourcis

### Pour les débutants - Commencez par ces 5 raccourcis :
1. **Alt + F11** - Ouvrir l'éditeur VBA
2. **F5** - Exécuter le code
3. **Ctrl + S** - Sauvegarder (fonctionne aussi dans VBA)
4. **Ctrl + Z** - Annuler
5. **Ctrl + F** - Rechercher

### Progression suggérée :
**Semaine 1** : Maîtrisez les 5 raccourcis de base  
**Semaine 2** : Ajoutez F8 (pas à pas) et F9 (point d'arrêt)  
**Semaine 3** : Apprenez Ctrl + G (fenêtre Exécution) et Ctrl + R (explorateur)  
**Mois suivant** : Intégrez progressivement les autres selon vos besoins  

### Astuces pour mémoriser :
- **Pratiquez régulièrement** : Utilisez consciemment les raccourcis au lieu de la souris
- **Affichez cette liste** : Gardez-la visible pendant que vous codez
- **Un par jour** : Apprenez un nouveau raccourci chaque jour
- **Associez à vos actions** : Pensez "F5 = test" quand vous testez votre code

### Raccourcis universels qui fonctionnent aussi :
- **Ctrl + S** : Sauvegarder
- **Ctrl + N** : Nouveau
- **Ctrl + O** : Ouvrir
- **Ctrl + P** : Imprimer
- **Alt + F4** : Fermer l'application

---

## 9. Personnalisation des raccourcis

### Modifier les raccourcis ★☆☆
Vous pouvez personnaliser certains raccourcis via :
1. **Outils** → **Options**
2. Onglet **Éditeur**
3. **Format du code** pour l'indentation
4. **Complétion automatique** pour les suggestions

### Barres d'outils personnalisées
- **Clic droit** sur une barre d'outils
- **Personnaliser** pour ajouter vos boutons favoris
- Glissez-déposez les commandes que vous utilisez souvent

---

## Aide-mémoire rapide

### Les 10 raccourcis indispensables :
1. **Alt + F11** - Ouvrir VBA
2. **F5** - Exécuter
3. **F8** - Pas à pas
4. **F9** - Point d'arrêt
5. **Ctrl + R** - Explorateur de projets
6. **Ctrl + F** - Rechercher
7. **Ctrl + S** - Sauvegarder
8. **Ctrl + Z** - Annuler
9. **F1** - Aide
10. **Ctrl + Break** - Arrêter l'exécution

### Raccourcis pour gagner du temps :
- **Ctrl + Espace** : Complétion automatique
- **Ctrl + I** : Info-bulle
- **Ctrl + G** : Fenêtre Exécution
- **F3** : Recherche suivante

**Conseil final :** Ne tentez pas d'apprendre tous les raccourcis d'un coup. Commencez par les essentiels et ajoutez-en de nouveaux au fur et à mesure que vous vous sentez à l'aise. Les raccourcis deviennent vraiment utiles quand ils deviennent automatiques !

⏭️ [D. Exemples de code réutilisables](/annexes/d-exemples-code-reutilisables.md)
