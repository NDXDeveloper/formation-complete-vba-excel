🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.2. Signature numérique des macros

## Qu'est-ce qu'une signature numérique ?

Une signature numérique est comme une **carte d'identité électronique** pour votre code VBA. Elle prouve que le code provient bien de vous et qu'il n'a pas été modifié depuis que vous l'avez signé.

Imaginez que vous envoyez une lettre importante par la poste. Pour prouver que c'est bien vous qui l'avez écrite et qu'elle n'a pas été modifiée en chemin, vous pourriez la signer avec votre signature manuscrite et utiliser un sceau officiel. La signature numérique fait exactement la même chose pour votre code VBA.

## Pourquoi signer numériquement vos macros ?

**Établir la confiance** : Les utilisateurs peuvent vérifier que votre code provient bien de vous et non d'un inconnu potentiellement malveillant.

**Prouver l'intégrité** : La signature garantit que le code n'a pas été modifié depuis que vous l'avez signé. Si quelqu'un change ne serait-ce qu'un caractère, la signature devient invalide.

**Contourner les restrictions de sécurité** : Excel fait plus facilement confiance aux macros signées par des développeurs reconnus, permettant leur exécution même avec des paramètres de sécurité stricts.

**Conformité professionnelle** : Dans les environnements d'entreprise, la signature numérique est souvent obligatoire pour distribuer du code VBA.

**Responsabilité légale** : La signature engage votre responsabilité, ce qui rassure les utilisateurs sur la qualité et la sécurité du code.

## Comment fonctionne la signature numérique ?

### Le principe de base

Une signature numérique utilise des **mathématiques complexes** (cryptographie) pour créer une "empreinte" unique de votre code. Cette empreinte est liée à votre identité de développeur grâce à un **certificat numérique**.

### Les éléments clés

**Certificat numérique** : C'est votre "carte d'identité" électronique, qui contient vos informations (nom, organisation, email) et est validé par une autorité de certification.

**Clé privée** : Un code secret que vous seul possédez et qui sert à créer vos signatures.

**Clé publique** : Un code public qui permet aux autres de vérifier vos signatures.

**Empreinte (hash)** : Une représentation mathématique unique de votre code.

### Le processus

1. **Création de l'empreinte** : Excel calcule une empreinte mathématique de votre code
2. **Chiffrement** : Cette empreinte est chiffrée avec votre clé privée
3. **Attachement** : La signature chiffrée est attachée à votre fichier
4. **Vérification** : Quand quelqu'un ouvre le fichier, Excel vérifie la signature avec votre clé publique

## Types de certificats numériques

### Certificats auto-signés (SelfCert)

**Ce que c'est** : Un certificat que vous créez vous-même sur votre ordinateur, gratuit mais limité.

**Avantages** :
- Gratuit et immédiat
- Facile à créer
- Suffisant pour un usage personnel ou des tests

**Inconvénients** :
- Non reconnu par d'autres ordinateurs
- Pas de validation d'identité
- Déclenche des avertissements de sécurité

**Quand l'utiliser** : Pour des tests, du développement personnel, ou dans un environnement contrôlé où vous pouvez ajouter le certificat aux ordinateurs de destination.

### Certificats d'autorité de certification

**Ce que c'est** : Un certificat émis par une autorité reconnue (comme VeriSign, DigiCert, etc.) qui valide votre identité.

**Avantages** :
- Reconnu universellement
- Validation officielle de votre identité
- Aucun avertissement de sécurité pour les utilisateurs
- Confiance immédiate

**Inconvénients** :
- Coût (généralement 100-500€ par an)
- Processus de validation parfois long
- Renouvellement périodique nécessaire

**Quand l'utiliser** : Pour la distribution commerciale, les environnements d'entreprise, ou quand la confiance maximale est requise.

## Créer un certificat auto-signé avec SelfCert

### Localiser l'outil SelfCert

1. **Recherchez** "SelfCert" dans le menu Démarrer de Windows
2. Ou naviguez vers le dossier d'installation d'Office :
   - `C:\Program Files\Microsoft Office\OfficeXX\` (remplacez XX par votre version)
   - Cherchez le fichier `SELFCERT.EXE`

### Créer le certificat

1. **Lancez SelfCert.exe**
2. **Saisissez un nom** pour votre certificat (exemple : "Mon Nom - Développeur VBA")
3. Cliquez sur **OK**
4. Le certificat est automatiquement créé et installé sur votre ordinateur

### Recommandations pour le nom

**Nom descriptif** : Utilisez votre vrai nom ou le nom de votre organisation
**Objectif clair** : Ajoutez "Développeur VBA" ou "Macros Excel" pour clarifier l'usage
**Exemples** :
- "Jean Dupont - Développeur VBA"
- "Entreprise ABC - Macros Internes"
- "Service Informatique - Solutions Excel"

## Signer vos macros

### Étape 1 : Ouvrir l'éditeur VBA

1. Ouvrez votre fichier Excel contenant les macros
2. Appuyez sur **Alt + F11** pour ouvrir l'éditeur VBA
3. Assurez-vous que votre code est finalisé (la signature sera invalidée si vous modifiez le code après)

### Étape 2 : Accéder au menu de signature

1. Dans l'éditeur VBA, allez dans le menu **Outils**
2. Cliquez sur **Signature numérique...**
3. Une boîte de dialogue s'ouvre

### Étape 3 : Sélectionner votre certificat

1. Dans la boîte de dialogue, cliquez sur **Choisir...**
2. **Sélectionnez votre certificat** dans la liste
3. Cliquez sur **OK**
4. Votre certificat apparaît maintenant dans la boîte de dialogue
5. Cliquez sur **OK** pour signer

### Étape 4 : Sauvegarder

1. **Sauvegardez votre fichier** Excel
2. La signature est maintenant attachée au fichier

## Vérifier une signature

### Ouvrir un fichier signé

Quand vous ouvrez un fichier avec des macros signées :

1. **Excel affiche une barre d'information** indiquant que le fichier contient des macros signées
2. **Cliquez sur la barre** pour voir les détails de la signature
3. Vous pouvez voir **qui a signé** le fichier et **quand**

### États de signature possibles

**Signature valide** :
- ✅ Icône verte ou message de confirmation
- Le code n'a pas été modifié depuis la signature
- Le certificat est reconnu et valide

**Signature invalide** :
- ❌ Icône rouge ou message d'avertissement
- Le code a été modifié après la signature
- Possible tentative de falsification

**Signature non reconnue** :
- ⚠️ Icône jaune ou message d'avertissement
- Certificat auto-signé ou autorité non reconnue
- Pas forcément dangereux, mais moins fiable

### Vérification détaillée

1. Dans l'éditeur VBA, allez dans **Outils** > **Signature numérique**
2. La boîte de dialogue montre l'état actuel de la signature
3. Cliquez sur **Détails** pour voir les informations complètes du certificat

## Gérer les certificats

### Voir vos certificats installés

1. Appuyez sur **Windows + R**
2. Tapez `certmgr.msc` et appuyez sur **Entrée**
3. Naviguez vers **Personnel** > **Certificats**
4. Vous voyez tous vos certificats personnels

### Exporter un certificat

Si vous voulez utiliser votre certificat sur un autre ordinateur :

1. Dans le gestionnaire de certificats, **clic droit** sur votre certificat
2. Sélectionnez **Toutes les tâches** > **Exporter**
3. Suivez l'assistant d'exportation
4. **Important** : N'exportez jamais la clé privée sur un ordinateur non sécurisé

### Importer un certificat

Pour faire confiance à un certificat sur un autre ordinateur :

1. **Double-cliquez** sur le fichier de certificat (.cer)
2. Cliquez sur **Installer le certificat**
3. Choisissez le magasin approprié (généralement "Personnes de confiance")

## Problèmes courants et solutions

### "Aucun certificat n'est disponible"

**Cause** : Aucun certificat n'est installé sur votre ordinateur

**Solution** :
1. Créez un certificat avec SelfCert
2. Ou installez un certificat d'autorité de certification
3. Vérifiez que le certificat est dans le bon magasin

### "La signature n'est pas valide"

**Causes possibles** :
- Le code a été modifié après signature
- L'horloge système est incorrecte
- Le certificat a expiré

**Solutions** :
1. Re-signez le code après modifications
2. Vérifiez la date et l'heure de votre ordinateur
3. Renouvelez le certificat si nécessaire

### "Éditeur non approuvé"

**Cause** : Le certificat n'est pas reconnu par Excel

**Solutions** :
1. Ajoutez le certificat aux "Éditeurs approuvés"
2. Utilisez un certificat d'autorité reconnue
3. Modifiez les paramètres de sécurité macro

## Impact sur la sécurité Excel

### Paramètres de sécurité

Les macros signées sont traitées différemment selon les paramètres de sécurité d'Excel :

**Désactiver toutes les macros avec notification** :
- Macros non signées : Bloquées avec notification
- Macros signées par éditeur non approuvé : Demande d'autorisation
- Macros signées par éditeur approuvé : Exécution automatique

**Désactiver toutes les macros sauf celles signées numériquement** :
- Seules les macros signées peuvent s'exécuter
- Très sécurisé mais peut bloquer des macros légitimes

### Éditeurs approuvés

Quand vous faites confiance à un développeur :

1. **Lors de l'ouverture** d'un fichier signé, cliquez sur "Approuver tous les documents de cet éditeur"
2. Le certificat est ajouté aux **éditeurs approuvés**
3. **Tous les futurs fichiers** signés par ce développeur s'exécuteront automatiquement

## Bonnes pratiques

### Pour la signature

**Signez en dernier** : Ne signez qu'une fois votre code complètement terminé et testé

**Gardez vos clés sécurisées** : Protégez votre certificat et votre clé privée

**Documentez vos signatures** : Tenez un registre de ce que vous avez signé et quand

**Testez sur d'autres ordinateurs** : Vérifiez que vos signatures fonctionnent ailleurs

### Pour la distribution

**Informez les utilisateurs** : Expliquez aux utilisateurs ce qu'est la signature numérique et pourquoi elle est importante

**Fournissez des instructions** : Donnez des instructions claires sur l'installation de votre certificat si nécessaire

**Maintenez la validité** : Renouvelez vos certificats avant expiration

**Support technique** : Préparez-vous à aider les utilisateurs avec les problèmes de signature

### Pour la sécurité

**Environnement sécurisé** : Signez vos macros sur un ordinateur sécurisé

**Sauvegarde des certificats** : Sauvegardez vos certificats dans un endroit sûr

**Révocation si nécessaire** : Sachez comment révoquer un certificat en cas de compromission

**Audit régulier** : Vérifiez régulièrement l'état de vos signatures

## Limitations et considérations

### Limitations techniques

**Modification du code** : Toute modification du code invalide la signature

**Dépendance au certificat** : Si le certificat expire ou est révoqué, la signature devient invalide

**Compatibilité** : Les signatures peuvent ne pas fonctionner sur toutes les versions d'Office

### Considérations organisationnelles

**Politique d'entreprise** : Respectez les politiques de signature de votre organisation

**Coût** : Les certificats officiels représentent un coût récurrent

**Complexité** : La gestion des certificats ajoute de la complexité au processus de développement

### Aspects légaux

**Responsabilité** : Signer du code engage votre responsabilité légale

**Non-répudiation** : Une signature valide prouve que vous avez bien signé le code

**Conformité** : Certaines réglementations exigent des signatures numériques

La signature numérique des macros est un outil puissant pour établir la confiance et la sécurité dans la distribution de code VBA. Bien qu'elle ajoute une étape au processus de développement, elle est essentielle pour toute distribution professionnelle ou commerciale de solutions VBA.

⏭️
