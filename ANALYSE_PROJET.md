# Analyse du projet BOTBYES360

## Vue d'ensemble
Le projet est une automatisation **desktop** pour Salesforce (BYES 360) composée de :
- une interface graphique Tkinter (`interface_bot.py`) pour piloter le bot et afficher les logs,
- un moteur Selenium (`bot_byes360.py`) pour naviguer dans Salesforce, télécharger un PDF PO, extraire un compte et mettre à jour le devis,
- un script batch Windows (`LANCER_BOT.bat`) pour installer les dépendances et lancer l'interface.

## Architecture
- **UI** : classe `BotInterface` avec thread de fond pour le bot, file `queue.Queue` pour les logs, `threading.Event` pour l'arrêt propre.
- **Automatisation** : classe `BotBYES360` avec étapes séquentielles : ouverture devis → recherche PO → téléchargement PDF → extraction compte → mise à jour du champ Compte.
- **Extraction PDF** : priorité à `pdfplumber`, avec fallback regex si structure du texte variable.

## Points forts
- Séparation claire UI / logique bot.
- Mécanisme d'arrêt utilisateur (`stop_event`) bien intégré.
- Logs horodatés par niveau (`INFO`, `SUCCESS`, `ERROR`, etc.).
- Tolérance aux variations UI Salesforce via plusieurs sélecteurs XPath fallback.

## Risques et limites
- Forte dépendance aux sélecteurs XPath Salesforce (fragile aux changements UI).
- Nombreux `time.sleep()` (stabilité potentiellement variable selon latence/réseau).
- Installation de dépendances au runtime (peut échouer hors connexion).
- Couplage Windows/Edge (profil utilisateur Edge requis).

## Correctifs effectués pendant l'analyse
1. Renommage du fichier bot vers `bot_byes360.py` pour alignement avec les imports (`from bot_byes360 import run_bot`) et la documentation.
2. Correction d'une erreur d'indentation dans la méthode `_type_account`, qui bloquait l'exécution Python.

## Recommandations prioritaires
1. Remplacer progressivement les `sleep` par des `WebDriverWait` ciblés.
2. Externaliser les sélecteurs XPath dans une configuration versionnée.
3. Ajouter un mode "dry-run" (sans clic final de sauvegarde) pour validation.
4. Ajouter des tests unitaires pour les fonctions d'extraction PDF.
5. Prévoir un fichier de config (`.env`/JSON) pour navigateur, timeouts, dossier de téléchargement.
