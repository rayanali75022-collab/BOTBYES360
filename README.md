# 🤖 Bot BYES 360 — Prise de Commande

## Installation (une seule fois)

1. Assure-toi que Python est installé (`py --version` dans cmd)
2. Double-clique sur **LANCER_BOT.bat** → il installe tout automatiquement

---

## Utilisation

### Étape 1 — Lancer le bot
Double-clique sur **LANCER_BOT.bat**

### Étape 2 — Coller l'URL du devis
Dans BYES 360, ouvre le devis que tu veux traiter et copie l'URL depuis la barre d'adresse.
Colle-la dans le champ de l'interface.

Format attendu :
```
https://equans.lightning.force.com/lightning/r/SBQQ__Quote_c/a2XIV000000eleX2AQ/view
```

### Étape 3 — Lancer
Clique sur **▶ LANCER LE BOT** et surveille les logs en temps réel.

### Étape 4 — Stop si problème
Si quelque chose va pas, clique sur **⛔ STOP** — le bot s'arrête proprement.

---

## Ce que fait le bot

1. Ouvre le devis dans Salesforce
2. Détecte le fichier PO (commence par "PO0") dans la section Fichiers
3. Télécharge le PDF
4. Extrait le nom du compte depuis **"Adresse de facturation"**
5. Efface le compte existant dans Salesforce
6. Tape le nouveau compte, sélectionne le résultat
7. Enregistre

---

## Alertes et erreurs

| Couleur log | Signification |
|-------------|---------------|
| 🟢 Vert | Succès |
| 🔴 Rouge | Erreur — bot stoppé sur cette étape |
| 🟡 Orange | Avertissement — bot continue mais attention |
| ⚫ Gris | Info standard |

En cas d'erreur critique, une popup apparaît automatiquement.

---

## Prérequis navigateur

Le bot utilise **Microsoft Edge** par défaut (déjà installé sur Windows 11).
Tu dois être **déjà connecté** à BYES 360 dans Edge avant de lancer le bot.

---

## Fichiers

| Fichier | Rôle |
|---------|------|
| `LANCER_BOT.bat` | Lance tout automatiquement |
| `interface_bot.py` | Interface graphique |
| `bot_byes360.py` | Logique du bot Selenium |

---

## Problèmes connus

**"Le champ Compte introuvable"** → Salesforce Lightning change parfois ses sélecteurs.
Me contacter pour mettre à jour le script.

**"PDF non lisible"** → Certains PDF sont des scans images. Dans ce cas l'extraction échoue.
