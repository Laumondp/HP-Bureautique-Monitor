# HP Bureautique Monitor

Application Google Apps Script pour surveiller les imprimantes HP Bureautique sur un plan interactif avec zoom.

**Version actuelle : v1.31**

## Fonctionnalités

- **Plan interactif** : Vue d'ensemble avec zoom et déplacement
- **Image de plan personnalisée** : Intégration d'un plan PNG en arrière-plan (transparence 50%)
- **Statut en temps réel** : Indicateurs vert (en ligne) / rouge (hors ligne) / orange (inconnu)
- **Images des modeles** : Photos des imprimantes HP affichees dans la liste et la barre laterale
- **Badge WiFi** : Icone WiFi pour les imprimantes sans fil avec couleur selon statut
- **Barre laterale droite** : Totaux par modele avec badges vert (online) et rouge (offline)
- **Mode Edition** : Bouton pour verrouiller/deverrouiller le deplacement des imprimantes
- **Grille magnetique** : Alignement automatique sur une grille de 10 pixels en mode edition
- **Design moderne** : Header avec effet RGB anime
- **Gestion des imprimantes** : Ajouter, modifier, supprimer facilement
- **Stockage Google Sheets** : Configuration centralisee et editable
- **Validation des donnees** : Unicite (ID, IP, MAC, Serie) et format (IP, MAC)
- **Alertes email** : Notifications automatiques lors des changements de statut
- **Controle d'acces** : Restriction par liste d'emails autorises
- **Securite API** : Cle API requise pour les requetes du script PowerShell
- **Auto-refresh** : Actualisation automatique toutes les 60 secondes
- **Force Refresh** : Bouton pour vider le cache et recharger les donnees

## Structure du projet

```
hp-bureautique-monitor/
├── Code.gs              # Code principal Google Apps Script
├── Index.html           # Interface utilisateur
├── appsscript.json      # Manifeste Google Apps Script
├── PingPrinters.ps1     # Script PowerShell pour ping
├── secrets.ps1          # Clé API secrète (non versionné)
├── secrets.template.ps1 # Template pour créer secrets.ps1
├── .claspignore         # Fichiers ignorés par clasp
├── .gitignore           # Fichiers ignorés par git
└── README.md            # Documentation
```

## Installation

### 1. Créer le projet Google Apps Script

1. Allez sur [script.google.com](https://script.google.com)
2. Créez un nouveau projet nommé "HP Bureautique Monitor"
3. Copiez le contenu des fichiers :
   - `Code.gs` → Code.gs
   - `Index.html` → Index.html (Fichier → Nouveau → Fichier HTML)

### 2. Initialiser la feuille Google Sheets

Dans l'éditeur Google Apps Script :
1. Sélectionnez la fonction `initializeSpreadsheet`
2. Cliquez sur "Exécuter"
3. Accordez les autorisations demandées
4. Notez l'ID du Google Sheets créé et mettez-le dans `HARDCODED_SPREADSHEET_ID` dans Code.gs

### 3. Déployer l'application web

1. Cliquez sur "Déployer" → "Nouveau déploiement"
2. Sélectionnez "Application Web"
3. Configuration :
   - Description : "HP Bureautique Monitor"
   - Exécuter en tant que : "Moi"
   - Qui a accès : "Tout le monde" (Anyone)
4. Cliquez sur "Déployer"
5. Copiez l'URL de l'application web dans `PingPrinters.ps1`

### 4. Configurer la clé API

1. Modifiez `secrets.ps1` avec votre clé API
2. Assurez-vous que la même clé est définie dans `Code.gs` (constante `API_SECRET_KEY`)

### 5. Exécuter le script de ping

```powershell
cd C:\chemin\vers\hp-bureautique-monitor
.\PingPrinters.ps1
```

## Configuration clasp (optionnel)

Pour synchroniser automatiquement avec Google Apps Script :

```powershell
npm install -g @google/clasp
clasp login
clasp clone <SCRIPT_ID>
```

## Licence

MIT License - Libre d'utilisation et de modification.
