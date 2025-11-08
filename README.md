# Serveur MCP Office 365

Implémentation complète du serveur MCP (Model Context Protocol) qui connecte Claude aux services Microsoft 365 via l'API Microsoft Graph.

> **🚀 Fonctionnement Autonome !** Fonctionne sans authentification par navigateur après la configuration initiale. Actualisation automatique des jetons et support du planificateur de tâches Windows pour un fonctionnement invisible en arrière-plan. Voir [TASK_SCHEDULER_SETUP.md](TASK_SCHEDULER_SETUP.md) pour le guide de configuration Windows.

## Fonctionnalités

- **Intégration Complète Microsoft 365** : Email, Calendrier, Teams, OneDrive/SharePoint, Contacts et Planner
- **Fonctionnement Autonome** : S'exécute sans navigateur après l'authentification initiale
- **Gestion Automatique des Jetons** : Stockage persistant des jetons avec actualisation automatique
- **Gestion des Pièces Jointes** : Téléchargement des pièces jointes intégrées et mappage des URL SharePoint vers les chemins locaux
- **Recherche Email Avancée** : Recherche unifiée avec support KQL et optimisation automatique
- **Gestion des Réunions Teams** : Accès aux transcriptions, enregistrements et insights IA
- **Gestion de Fichiers** : Opérations complètes sur OneDrive et SharePoint
- **Gestion des Contacts** : Opérations CRUD complètes pour les contacts Outlook avec recherche avancée
- **Gestion des Tâches** : Intégration complète avec Microsoft Planner
- **Chemins Configurables** : Variables d'environnement pour tous les chemins de synchronisation locaux

## Démarrage Rapide

### Prérequis
- Node.js 16 ou supérieur
- Compte Microsoft 365 (personnel ou professionnel/scolaire)
- Enregistrement d'application Azure (voir ci-dessous)

### Installation

1. Cloner le dépôt :
```bash
git clone https://github.com/yourusername/office-mcp.git
cd office-mcp
```

2. Installer les dépendances :
```bash
npm install
```

3. Copier le modèle d'environnement :
```bash
cp .env.example .env
```

4. Configurer votre fichier `.env` avec :
   - Identifiants de l'application Azure (voir Configuration Azure ci-dessous)
   - Chemins de fichiers locaux pour la synchronisation SharePoint/OneDrive
   - Paramètres optionnels

5. Exécuter l'authentification initiale :
```bash
npm run auth-server
# Visitez http://localhost:3000/auth et connectez-vous
```

6. Configurer Claude Desktop (voir Configuration Claude Desktop ci-dessous)

## Capacités Principales

### Opérations Email
- **Recherche Unifiée** : Outil `email_search` unique avec optimisation automatique
- **Extraction de Contacts** : Extrait les contacts des emails avec URLs LinkedIn, numéros de téléphone et informations d'entreprise
- **Gestion des Pièces Jointes** : Téléchargement des pièces jointes intégrées, mappage des URL SharePoint vers les chemins locaux
- **Fonctionnalités Avancées** : Catégories, règles, boîte de réception prioritaire, gestion des dossiers
- **Opérations par Lots** : Déplacement efficace de plusieurs emails

### Gestion du Calendrier
- **Opérations CRUD Complètes** : Créer, lire, mettre à jour, supprimer des événements
- **Intégration Teams** : Créer des réunions avec liens Teams
- **Support de la Récurrence** : Modèles d'événements récurrents complexes
- **Gestion du Fuseau Horaire UTC** : Gestion appropriée des fuseaux horaires

### Fonctionnalités Teams
- **Gestion des Réunions** : Créer, mettre à jour, annuler des réunions
- **Accès aux Transcriptions** : Récupérer les transcriptions des réunions
- **Accès aux Enregistrements** : Accéder aux enregistrements des réunions
- **Opérations sur les Canaux** : Messages, membres, onglets
- **Gestion des Chats** : Créer, envoyer, gérer les messages de chat

### Gestion de Fichiers
- **Intégration SharePoint** : Mappage des chemins de synchronisation locaux
- **Prise en charge de OneDrive** : Opérations de fichiers complètes
- **Opérations par Lots** : Télécharger/téléverser plusieurs fichiers
- **Recherche** : Recherche de contenu et de métadonnées

### Gestion des Contacts
- **Opérations CRUD Complètes** : Créer, lire, mettre à jour, supprimer des contacts
- **Recherche Avancée** : Rechercher par nom, email, entreprise ou tout champ de contact
- **Champs de Contact Complets** : Support pour emails, téléphones, adresses, anniversaires, notes
- **Gestion des Dossiers** : Organiser les contacts dans des dossiers
- **Opérations en Masse** : Gérer plusieurs contacts efficacement

### Gestion des Tâches (Planner)
- **Opérations sur les Plans** : Créer et gérer des plans
- **Affectation des Tâches** : Recherche et affectation d'utilisateurs
- **Organisation en Compartiments** : Grouper les tâches efficacement
- **Opérations en Masse** : Mettre à jour/supprimer plusieurs tâches

### Extraction de Contacts depuis les Emails
- **Analyse Intelligente** : Extraction automatique des informations de contact depuis les métadonnées et le contenu des emails
- **Extraction de Données Riches** : Capture des noms, emails, numéros de téléphone, URLs LinkedIn, noms d'entreprise et titres de poste
- **Détection de Signature** : Détection et analyse intelligente des signatures d'email (français et anglais)
- **Déduplication** : Fusion automatique des contacts en double avec scoring de confiance
- **Filtrage des Newsletters** : Détection et exclusion automatique des emails marketing/bulk avec analyse multi-facteurs (en-têtes, expéditeur, contenu)
- **Référence Croisée Outlook** : Identifier les nouveaux contacts non présents dans vos contacts Outlook
- **Export CSV** : Exporter tous les contacts extraits vers un fichier CSV pour un import facile
- **Filtrage Flexible** : Filtrer par plage de dates, requête de recherche ou dossiers spécifiques
- **Traitement par Lots** : Traiter des milliers d'emails efficacement avec suivi de progression
- **Support Multilingue** : Reconnaissance des signatures en français et anglais, numéros de téléphone français (+33, 01 XX XX XX XX)

**Exemple d'Utilisation :**
```
Extraire tous les contacts de mes emails de la boîte de réception des 30 derniers jours et sauvegarder en CSV
```

**Format de Sortie CSV :**
- email, displayName, firstName, lastName
- phoneNumbers, linkedInUrls
- companyName, jobTitle
- source (metadata/signature/body), isInOutlook, firstSeenDate, extractionConfidence

**Formats Français Supportés :**
- Signatures : Cordialement, Bien cordialement, Salutations, Amitiés, etc.
- Téléphones : +33 X XX XX XX XX, 01.XX.XX.XX.XX, 06 XX XX XX XX
- Entreprises : SA, SARL, SAS, SASU, SNC, Société
- Titres : PDG, DG, Directeur, Responsable, Ingénieur, etc.

**Filtrage Intelligent des Newsletters :**

Le système inclut un détecteur de newsletters sophistiqué qui filtre automatiquement les emails marketing et bulk pour extraire uniquement les contacts professionnels pertinents.

**Signaux de Détection (Bilingue Français/Anglais) :**
- **En-têtes** : List-Unsubscribe, Precedence: bulk, ESP (Mailchimp, Sendgrid, etc.)
- **Expéditeur** : noreply@, newsletter@, ne-pas-repondre@, marketing@, bulletin@
- **Contenu** : Liens de désabonnement, "Afficher dans le navigateur", préférences email
- **Structure** : Ratio images/texte élevé, pixels de tracking, mises en page en tableaux
- **Destinataires** : BCC, noms génériques ("Cher Client", "Valued Customer")

**Configuration Personnalisée :**

Modifiez `/config/newsletter-rules.json` pour ajuster le filtrage :
```json
{
  "whitelist": {
    "domains": ["partenaire-important.com"],
    "senders": ["newsletter-utile@entreprise.com"]
  },
  "blacklist": {
    "domains": ["spam-company.com"],
    "senders": ["promo@publicite.com"]
  },
  "settings": {
    "defaultThreshold": 60
  }
}
```

**Paramètres de l'Outil :**
- `excludeNewsletters` : Activer/désactiver le filtrage (défaut: true)
- `newsletterThreshold` : Seuil de confiance 0-100 (défaut: 60, plus élevé = filtrage plus strict)
- `saveNewsletterReport` : Enregistrer un rapport JSON des newsletters filtrées (défaut: false)

**Exemple d'Utilisation :**
```
Extraire les contacts des 90 derniers jours, exclure les newsletters avec un seuil de 70, et sauvegarder le rapport
```

## Enregistrement et Configuration de l'Application Azure

Pour utiliser ce serveur MCP, vous devez d'abord enregistrer et configurer une application dans le Portail Azure. Les étapes suivantes vous guideront dans le processus d'enregistrement d'une nouvelle application, de configuration de ses permissions et de génération d'un secret client.

### Enregistrement de l'Application

1. Ouvrir le [Portail Azure](https://portal.azure.com/) dans votre navigateur
2. Se connecter avec un compte Microsoft professionnel ou personnel
3. Rechercher ou cliquer sur "Inscriptions d'applications"
4. Cliquer sur "Nouvelle inscription"
5. Entrer un nom pour l'application, par exemple "Office MCP Server"
6. Sélectionner l'option "Comptes dans un annuaire organisationnel et comptes Microsoft personnels"
7. Dans la section "URI de redirection", sélectionner "Web" dans le menu déroulant et entrer "http://localhost:3000/auth/callback" dans la zone de texte
8. Cliquer sur "Inscrire"
9. Depuis la section Vue d'ensemble de la page des paramètres de l'application, copier "l'ID d'application (client)" et le saisir comme OFFICE_CLIENT_ID dans le fichier .env ainsi que dans le fichier claude-config-sample.json

### Permissions de l'Application

1. Depuis la page des paramètres de l'application dans le Portail Azure, sélectionner l'option "Autorisations API" dans la section Gérer
2. Cliquer sur "Ajouter une autorisation"
3. Cliquer sur "Microsoft Graph"
4. Sélectionner "Autorisations déléguées"
5. Rechercher et sélectionner la case à cocher à côté de chacune de ces permissions :
    - offline_access
    - User.Read
    - User.ReadWrite
    - User.ReadBasic.All
    - Mail.Read
    - Mail.ReadWrite
    - Mail.Send
    - Calendars.Read
    - Calendars.ReadWrite
    - Contacts.ReadWrite
    - Files.Read
    - Files.ReadWrite
    - Files.ReadWrite.All
    - Team.ReadBasic.All
    - Team.Create
    - Chat.Read
    - Chat.ReadWrite
    - ChannelMessage.Read.All
    - ChannelMessage.Send
    - OnlineMeetingTranscript.Read.All
    - OnlineMeetings.ReadWrite
    - Tasks.Read
    - Tasks.ReadWrite
    - Group.Read.All
    - Directory.Read.All
    - Presence.Read
    - Presence.ReadWrite
6. Cliquer sur "Ajouter des autorisations"

### Secret Client

1. Depuis la page des paramètres de l'application dans le Portail Azure, sélectionner l'option "Certificats et secrets" dans la section Gérer
2. Passer à l'onglet "Secrets clients"
3. Cliquer sur "Nouveau secret client"
4. Entrer une description, par exemple "Secret Client"
5. Sélectionner la durée d'expiration la plus longue possible
6. Cliquer sur "Ajouter"
7. Copier la valeur du secret et la saisir comme OFFICE_CLIENT_SECRET dans le fichier .env ainsi que dans le fichier claude-config-sample.json

## Configuration de l'Environnement

### Variables Requises
```bash
# Enregistrement de l'Application Azure
OFFICE_CLIENT_ID=your-azure-app-client-id
OFFICE_CLIENT_SECRET=your-azure-app-client-secret
OFFICE_TENANT_ID=common

# Authentification
OFFICE_REDIRECT_URI=http://localhost:3000/auth/callback
```

### Variables Optionnelles
```bash
# Chemins de fichiers locaux (personnaliser selon votre système)
SHAREPOINT_SYNC_PATH=/path/to/your/sharepoint/sync
ONEDRIVE_SYNC_PATH=/path/to/your/onedrive/sync
TEMP_ATTACHMENTS_PATH=/path/to/temp/attachments
SHAREPOINT_SYMLINK_PATH=/path/to/sharepoint/symlink

# Paramètres du serveur
USE_TEST_MODE=false
TRANSPORT_TYPE=stdio  # ou 'http' pour le mode autonome
HTTP_PORT=3333
HTTP_HOST=127.0.0.1
```

## Configuration de Claude Desktop

1. Localiser votre fichier de configuration Claude Desktop :
   - Windows : `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS : `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Linux : `~/.config/Claude/claude_desktop_config.json`

2. Ajouter la configuration du serveur MCP :
```json
{
  "mcpServers": {
    "office-mcp": {
      "command": "node",
      "args": ["/path/to/office-mcp/index.js"],
      "env": {
        "OFFICE_CLIENT_ID": "votre-client-id",
        "OFFICE_CLIENT_SECRET": "votre-client-secret",
        "SHAREPOINT_SYNC_PATH": "/path/to/sharepoint",
        "ONEDRIVE_SYNC_PATH": "/path/to/onedrive"
      }
    }
  }
}
```

3. Redémarrer Claude Desktop

4. Dans Claude, utiliser l'outil `authenticate` pour se connecter à Microsoft 365

## Tests

### Inspecteur MCP
Tester le serveur directement en utilisant l'Inspecteur MCP :
```bash
npx @modelcontextprotocol/inspector node index.js
```

### Mode Test
Activer le mode test pour utiliser des données simulées sans appels API :
```bash
USE_TEST_MODE=true node index.js
```

## Flux d'Authentification

1. Démarrer le serveur d'authentification :
   - Windows : Exécuter `start-auth-server.bat` ou `run-office-mcp.bat`
   - Unix/Linux/macOS : Exécuter `./start-auth-server.sh`
2. Le serveur d'authentification s'exécute sur le port 3000 et gère les callbacks OAuth
3. Dans Claude, utiliser l'outil `authenticate` pour obtenir une URL d'authentification
4. Compléter l'authentification dans votre navigateur
5. Les jetons sont stockés dans `~/.office-mcp-tokens.json`

## Fonctionnement Autonome

### Actualisation Automatique des Jetons
Après l'authentification initiale, le serveur actualise automatiquement les jetons sans interaction utilisateur.

### Mode Transport HTTP
Pour les environnements autonomes, utiliser le transport HTTP :
```bash
TRANSPORT_TYPE=http HTTP_PORT=3333 node index.js
```

### Service Windows (Optionnel)
Pour un fonctionnement en arrière-plan sur Windows :
1. Compléter l'authentification initiale
2. Configurer comme tâche du Planificateur de tâches Windows
3. S'exécute de manière invisible au démarrage du système

## Dépannage

### Problèmes Courants

1. **Erreurs d'Authentification**
   - Assurez-vous que l'application Azure dispose des bonnes permissions
   - Vérifiez que le fichier de jeton existe : `~/.office-mcp-tokens.json`
   - Vérifiez que l'URI de redirection correspond à la configuration Azure

2. **Recherche d'Email avec Filtres de Date**
   - Les recherches filtrées par date sont maintenant routées directement vers l'API $filter pour plus de fiabilité
   - Utilisez le caractère générique `*` pour tous les emails dans une plage de dates
   - `startDate` et `endDate` supportent le format ISO (2025-08-27) ou relatif (7d/1w/1m/1y)

3. **Problèmes de Pièces Jointes d'Email**
   - Configurez les chemins de synchronisation locaux dans `.env`
   - Assurez-vous que le répertoire temporaire dispose des permissions d'écriture
   - Vérifiez que la synchronisation SharePoint est active

4. **Limites de Taux API**
   - Le serveur inclut une nouvelle tentative automatique avec backoff exponentiel
   - Réduisez la fréquence des requêtes si le problème persiste

5. **Erreurs de Permission**
   - Vérifiez que toutes les permissions Graph API requises sont accordées
   - Le consentement administrateur peut être requis pour certaines permissions

## Considérations de Sécurité

- **Stockage des Jetons** : Les jetons sont chiffrés et stockés localement
- **Variables d'Environnement** : Ne jamais committer les fichiers `.env`
- **Secrets Clients** : Rotation régulière et utilisation d'Azure Key Vault en production
- **Chemins Locaux** : Utiliser des variables d'environnement au lieu de chemins codés en dur
- **Journalisation d'Audit** : Tous les appels API sont journalisés pour la surveillance de sécurité

## Contribution

Les contributions sont les bienvenues ! Veuillez :
1. Fork le dépôt
2. Créer une branche de fonctionnalité
3. Soumettre une pull request

## Licence

Licence MIT - Voir le fichier LICENSE pour plus de détails
