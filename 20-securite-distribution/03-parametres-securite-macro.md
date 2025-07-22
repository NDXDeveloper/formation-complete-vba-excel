🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.3. Paramètres de sécurité macro

## Qu'est-ce que les paramètres de sécurité macro ?

Les paramètres de sécurité macro sont comme les **règles de sécurité d'un aéroport** : ils déterminent quelles macros peuvent "embarquer" (s'exécuter) dans Excel et lesquelles doivent être bloquées ou inspectées. Ces paramètres protègent votre ordinateur contre les macros potentiellement dangereuses tout en vous permettant d'utiliser celles qui sont légitimes.

Excel propose plusieurs niveaux de sécurité, depuis "laisser passer tout le monde" (très risqué) jusqu'à "ne laisser passer personne" (très sécurisé mais contraignant). Comprendre ces paramètres vous aide à trouver le bon équilibre entre sécurité et fonctionnalité.

## Pourquoi Excel bloque-t-il les macros ?

**Protection contre les malwares** : Les cybercriminels utilisent souvent les macros pour infecter les ordinateurs avec des virus, ransomwares ou autres logiciels malveillants.

**Prévention des accidents** : Une macro mal écrite peut endommager vos données, supprimer des fichiers, ou perturber le fonctionnement de votre ordinateur.

**Conformité sécuritaire** : Les entreprises ont besoin de contrôler quels codes peuvent s'exécuter sur leurs systèmes pour respecter leurs politiques de sécurité.

**Protection des utilisateurs non techniques** : Beaucoup d'utilisateurs ne savent pas distinguer une macro sûre d'une macro dangereuse.

## Accéder aux paramètres de sécurité macro

### Via le menu Excel

1. Ouvrez Excel
2. Cliquez sur **Fichier** (onglet en haut à gauche)
3. Cliquez sur **Options** (en bas du menu)
4. Dans la fenêtre qui s'ouvre, sélectionnez **Centre de gestion de la confidentialité** (dans la liste de gauche)
5. Cliquez sur **Paramètres du Centre de gestion de la confidentialité...**
6. Sélectionnez **Paramètres des macros** dans la liste de gauche

### Via le ruban Excel

1. Allez dans l'onglet **Développeur** du ruban
2. Cliquez sur **Sécurité des macros** dans le groupe "Code"
3. Vous arrivez directement aux paramètres des macros

**Note** : Si vous ne voyez pas l'onglet Développeur, activez-le dans Fichier > Options > Personnaliser le ruban.

## Les quatre niveaux de sécurité

### 1. Désactiver toutes les macros sans notification

**Ce que ça fait** : Bloque toutes les macros, même les vôtres, sans vous prévenir.

**Niveau de sécurité** : Maximum

**Avantages** :
- Protection totale contre les macros malveillantes
- Aucun risque d'exécution accidentelle
- Idéal pour les ordinateurs qui n'ont jamais besoin de macros

**Inconvénients** :
- Aucune macro ne fonctionne, même les légitimes
- Peut casser des fichiers Excel qui dépendent des macros
- Aucune notification, donc difficile de diagnostiquer les problèmes

**Quand l'utiliser** :
- Ordinateurs utilisés uniquement pour consulter des données
- Environnements à très haute sécurité
- Utilisateurs qui n'ont jamais besoin de macros

### 2. Désactiver toutes les macros avec notification (RECOMMANDÉ)

**Ce que ça fait** : Bloque les macros par défaut mais affiche une **barre d'avertissement jaune** vous permettant de les activer manuellement.

**Niveau de sécurité** : Élevé

**Avantages** :
- Protection par défaut contre les macros dangereuses
- Possibilité d'activer les macros quand vous en avez besoin
- Vous garde conscient de la présence de macros
- Bon équilibre sécurité/fonctionnalité

**Inconvénients** :
- Nécessite une action manuelle pour chaque fichier
- Peut être agaçant si vous travaillez souvent avec des macros
- Risque d'habituation aux clics d'activation

**Quand l'utiliser** :
- **Paramètre recommandé pour la plupart des utilisateurs**
- Environnements où on utilise occasionnellement des macros
- Quand vous voulez contrôler l'exécution des macros

**Comment ça marche** :
- Quand vous ouvrez un fichier avec macros, une barre jaune apparaît
- Cliquez sur **"Activer le contenu"** pour autoriser les macros
- L'autorisation ne vaut que pour cette session

### 3. Désactiver toutes les macros sauf celles signées numériquement

**Ce que ça fait** : Autorise automatiquement les macros signées par des développeurs approuvés, bloque toutes les autres.

**Niveau de sécurité** : Moyennement élevé

**Avantages** :
- Autorisation automatique des macros de confiance
- Protection contre les macros non signées
- Bon pour les environnements professionnels avec des développeurs certifiés

**Inconvénients** :
- Bloque toutes vos propres macros non signées
- Nécessite que les développeurs signent leurs macros
- Plus complexe à gérer

**Quand l'utiliser** :
- Entreprises avec des politiques de signature strictes
- Quand vous travaillez principalement avec des macros commerciales signées
- Environnements où la signature numérique est standard

### 4. Activer toutes les macros (non recommandé)

**Ce que ça fait** : Exécute automatiquement toutes les macros, sans aucune vérification.

**Niveau de sécurité** : Très faible

**Avantages** :
- Aucune interruption dans le travail
- Toutes les macros fonctionnent immédiatement
- Simple pour les développeurs

**Inconvénients** :
- **TRÈS DANGEREUX** : aucune protection contre les macros malveillantes
- Peut exécuter du code nuisible sans vous prévenir
- Non recommandé par Microsoft et les experts en sécurité

**Quand l'utiliser** :
- **JAMAIS en utilisation normale**
- Uniquement pour des tests de développement sur un ordinateur isolé
- Environnements complètement contrôlés et sécurisés

## Emplacements approuvés

### Qu'est-ce que c'est ?

Les emplacements approuvés sont des **dossiers spéciaux** sur votre ordinateur où Excel fait automatiquement confiance à tous les fichiers, y compris leurs macros. C'est comme avoir une "zone VIP" où les contrôles de sécurité sont allégés.

### Comment ça fonctionne

**Autorisation automatique** : Tous les fichiers dans ces dossiers peuvent exécuter leurs macros sans restriction, quel que soit votre niveau de sécurité.

**Sous-dossiers inclus** : Par défaut, les sous-dossiers sont aussi considérés comme approuvés.

**Paramètre par utilisateur** : Chaque utilisateur Windows a ses propres emplacements approuvés.

### Emplacements par défaut

Excel configure automatiquement certains dossiers comme approuvés :

**Dossier de démarrage Excel** :
- `C:\Users\[VotreNom]\AppData\Roaming\Microsoft\Excel\XLSTART\`
- Fichiers qui s'ouvrent automatiquement au démarrage d'Excel

**Dossier des modèles** :
- `C:\Users\[VotreNom]\AppData\Roaming\Microsoft\Templates\`
- Modèles de documents Excel

**Dossier des compléments** :
- Divers dossiers selon votre installation Office

### Ajouter un emplacement approuvé

1. Allez dans **Centre de gestion de la confidentialité** > **Emplacements approuvés**
2. Cliquez sur **Ajouter un nouvel emplacement...**
3. **Parcourez** pour sélectionner le dossier
4. **Cochez** "Les sous-dossiers de cet emplacement sont également approuvés" si souhaité
5. Ajoutez une **description** pour vous souvenir pourquoi ce dossier est approuvé
6. Cliquez sur **OK**

### Bonnes pratiques pour les emplacements approuvés

**Soyez restrictif** : N'ajoutez que les dossiers que vous contrôlez complètement

**Évitez les dossiers partagés** : Ne mettez pas de dossiers réseau accessibles à d'autres

**Surveillance** : Surveillez régulièrement ce qui se trouve dans ces dossiers

**Documentation** : Notez pourquoi chaque emplacement a été ajouté

**Exemples appropriés** :
- Votre dossier de développement personnel : `C:\MesMacros\`
- Dossier d'outils internes : `C:\OutilsEntreprise\`

**Exemples à éviter** :
- Dossier Téléchargements : `C:\Users\[VotreNom]\Downloads\`
- Bureau : `C:\Users\[VotreNom]\Desktop\`
- Dossiers réseau partagés

## Éditeurs approuvés

### Qu'est-ce que c'est ?

Les éditeurs approuvés sont des **développeurs en qui vous avez confiance**. Une fois qu'un développeur est dans cette liste, toutes ses macros signées s'exécutent automatiquement.

### Comment ajouter un éditeur approuvé

**Méthode automatique** (recommandée) :
1. Ouvrez un fichier avec des macros signées par ce développeur
2. Quand la barre d'avertissement apparaît, cliquez sur **Options...**
3. Sélectionnez **"Approuver tous les documents de cet éditeur"**
4. Cliquez sur **OK**

**Méthode manuelle** :
1. Allez dans **Centre de gestion de la confidentialité** > **Éditeurs approuvés**
2. Les éditeurs de confiance apparaissent dans la liste
3. Vous pouvez les supprimer en les sélectionnant et cliquant sur **Supprimer**

### Gérer les éditeurs approuvés

**Révision périodique** : Vérifiez régulièrement votre liste et supprimez les éditeurs dont vous n'avez plus besoin

**Principe de moindre privilège** : N'approuvez que les éditeurs strictement nécessaires

**Attention aux certificats expirés** : Les éditeurs avec des certificats expirés peuvent poser des problèmes

## Paramètres avancés

### Désactiver tous les compléments d'application

**Ce que ça fait** : Empêche le chargement des compléments (add-ins) d'application.

**Quand l'utiliser** : Environnements très sécurisés où les compléments sont interdits.

### Appliquer les paramètres de macro aux compléments installés

**Ce que ça fait** : Étend vos paramètres de sécurité macro aux compléments Excel.

**Recommandation** : Généralement laissé coché pour une sécurité cohérente.

### Faire confiance à l'accès au modèle objet du projet VBA

**Ce que ça fait** : Permet à d'autres applications (comme Word, PowerPoint) d'accéder au code VBA d'Excel.

**Attention** : Ne cochez cette option que si vous utilisez des solutions qui nécessitent cette interaction entre applications Office.

## Situations courantes et solutions

### "Mes macros ne fonctionnent plus"

**Cause probable** : Paramètres de sécurité trop restrictifs

**Solutions** :
1. Vérifiez le paramètre : assurez-vous d'être en mode "avec notification"
2. Cherchez la barre d'avertissement jaune et cliquez sur "Activer le contenu"
3. Considérez ajouter votre dossier de travail aux emplacements approuvés

### "Excel demande toujours d'activer les macros"

**Cause** : Comportement normal avec le paramètre "avec notification"

**Solutions** :
1. Ajoutez le dossier aux emplacements approuvés
2. Signez vos macros et ajoutez-vous aux éditeurs approuvés
3. Acceptez ce comportement comme une mesure de sécurité

### "Les macros de mes collègues ne fonctionnent pas"

**Causes possibles** :
- Macros non signées avec paramètres restrictifs
- Emplacement non approuvé
- Éditeur non approuvé

**Solutions** :
1. Demandez à vos collègues de signer leurs macros
2. Créez un dossier partagé approuvé (avec précaution)
3. Établissez une politique de signature d'équipe

### "Erreur de sécurité macro"

**Diagnostic** :
1. Vérifiez l'état de la signature (valide/invalide/expirée)
2. Contrôlez les paramètres de sécurité actuels
3. Examinez les emplacements et éditeurs approuvés

## Paramètres recommandés par contexte

### Utilisateur débutant

**Paramètre macro** : "Désactiver toutes les macros avec notification"
**Emplacements approuvés** : Aucun ajout initial
**Éditeurs approuvés** : Ajouter au cas par cas
**Avantage** : Protection maximale avec possibilité d'apprentissage

### Développeur VBA

**Paramètre macro** : "Désactiver toutes les macros avec notification"
**Emplacements approuvés** : Dossier de développement personnel
**Éditeurs approuvés** : Votre propre certificat
**Avantage** : Productivité pour le développement, sécurité pour les autres fichiers

### Entreprise avec politique stricte

**Paramètre macro** : "Désactiver toutes les macros sauf celles signées numériquement"
**Emplacements approuvés** : Dossiers d'applications métier
**Éditeurs approuvés** : Développeurs internes certifiés
**Avantage** : Sécurité maximale avec fonctionnalité contrôlée

### Environnement de formation

**Paramètre macro** : "Désactiver toutes les macros avec notification"
**Emplacements approuvés** : Dossier des exercices de formation
**Éditeurs approuvés** : Instructeur/formateur
**Avantage** : Apprentissage facilité tout en maintenant la sécurité

## Dépannage des problèmes de sécurité

### Vérifier les paramètres actuels

1. **Ouvrez le Centre de gestion** comme décrit précédemment
2. **Notez le paramètre** actuel des macros
3. **Examinez la liste** des emplacements approuvés
4. **Vérifiez les éditeurs** approuvés

### Messages d'erreur courants

**"Les macros ont été désactivées"** :
- Solution : Activez les macros via la barre d'avertissement ou modifiez les paramètres

**"Signature non valide"** :
- Solution : Vérifiez si la macro a été modifiée après signature, ou si le certificat a expiré

**"Éditeur non approuvé"** :
- Solution : Ajoutez l'éditeur aux éditeurs approuvés ou modifiez les paramètres

### Test de fonctionnement

Pour tester vos paramètres :

1. **Créez une macro simple** dans un nouveau fichier
2. **Sauvegardez et fermez** le fichier
3. **Rouvrez le fichier** et observez le comportement
4. **Ajustez les paramètres** si nécessaire

## Sécurité et politique d'entreprise

### Considérations organisationnelles

**Politiques IT** : Respectez les politiques de sécurité de votre organisation

**Formation des utilisateurs** : Sensibilisez les équipes aux risques liés aux macros

**Audit régulier** : Vérifiez périodiquement les paramètres sur tous les postes

**Documentation** : Maintenez une documentation claire des paramètres approuvés

### Gestion centralisée

Dans les grandes organisations :

**Stratégies de groupe** : Les administrateurs peuvent imposer des paramètres via Active Directory

**Déploiement automatisé** : Les paramètres peuvent être configurés automatiquement sur tous les postes

**Surveillance** : Les équipes IT peuvent monitorer l'utilisation des macros

Les paramètres de sécurité macro sont un élément crucial pour trouver l'équilibre entre productivité et sécurité. Une configuration réfléchie, adaptée à votre contexte d'usage, vous permettra de bénéficier de la puissance des macros VBA tout en protégeant votre système contre les menaces potentielles.

⏭️
