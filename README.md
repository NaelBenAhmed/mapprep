# Guide Préparateur

Appli mobile pour les préparateurs de commandes : chrono de mission selon le type et le nombre de colis, suivi du quota et de la prime (calculs repris de l'outil Quota colis), et plan de l'entrepôt avec itinéraire.

En ligne : https://naelbenahmed.github.io/mapprep/

## Installer sur le téléphone
- iPhone (Safari) : Partager › Sur l'écran d'accueil
- Android (Chrome) : ⋮ › Installer l'application

## Mises à jour
Lancer `./bump.sh` avant de publier : il change le numéro de version dans `version.json`, `sw.js` et `index.html`. Les téléphones se mettent à jour tout seuls à l'ouverture de l'appli.

## Fichiers
- `index.html` : l'appli
- `entrepot.json` : plan (Google Sheet « Plan Bis » : onglets Plan Epaux + Mapping_IDs)
- `sw.js`, `manifest.webmanifest`, icônes : installation et fonctionnement hors ligne
