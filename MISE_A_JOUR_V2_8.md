# QHSE Manager Pro v2.8

## Nouveautés

- Nom de l’application simplifié en **QHSE Manager Pro**.
- Navigation regroupée en menus compacts : Administration, Produits chimiques, Rapports journaliers, Pilotage et Messagerie.
- Suppression du doublon « Rapport journalier » dans la navigation.
- Retour au jeu de couleurs bleu historique de l’application.
- En-tête du tableau du rapport journalier figé pendant le défilement.
- Nouveau **plan d’action mensuel** avec création, modification, suppression, filtres, pagination de 20 lignes et export Excel.
- Couleurs automatiques :
  - Priorité : Critique, Élevée, Moyenne, Faible.
  - Statut : Ouvert, En cours, Fermé, Suspendu.
- Nouveau **centre de messages** pour les administrateurs : email individuel ou groupé, lien WhatsApp direct et lien du groupe WhatsApp du projet.
- Historique des messages, 20 lignes par page.
- Endpoints protégés pour les rappels automatiques :
  - `/tasks/monthly-chemical-reminder` : rappel des produits chimiques le 26.
  - `/tasks/workforce-reminder` : rappel des effectifs après 10h lorsque la saisie manque.

## Limite WhatsApp importante

Une application web ordinaire ne peut pas publier silencieusement dans un groupe WhatsApp. La version v2.8 peut :

1. ouvrir une conversation individuelle avec un message prérempli ;
2. ouvrir le groupe du projet à partir de son lien ;
3. conserver le message à copier-coller.

Un envoi totalement automatique nécessiterait l’API officielle WhatsApp Business et un fournisseur approuvé.

## Variables Render à conserver

```text
DATA_DIR=/var/data
CLOUD_MODE=true
SESSION_TIMEOUT_MINUTES=30
```

## Variables pour les emails

```text
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_USERNAME=...
SMTP_PASSWORD=...
SMTP_FROM=qhse@example.com
SMTP_USE_SSL=false
TASK_TOKEN=une-valeur-secrete-longue
```

Utiliser un mot de passe d’application lorsque le fournisseur de messagerie l’exige.

## Planification des rappels

Le Web Service doit recevoir des requêtes planifiées avec l’en-tête :

```text
X-Task-Token: valeur-de-TASK_TOKEN
```

Appels à programmer :

- Tous les jours à 10h05, heure de Lomé : `/tasks/workforce-reminder`
- Chaque 26 du mois, par exemple à 08h00 : `/tasks/monthly-chemical-reminder`

Ces appels peuvent être exécutés par Render Cron Job, cron-job.org, UptimeRobot ou un service similaire. Les endpoints accèdent ensuite à la base persistante via le Web Service.

## Déploiement

Remplacer proprement les fichiers du dépôt local, puis :

```bash
git add .
git commit -m "Mise à jour QHSE Manager Pro v2.8"
git push origin main
```

Les nouvelles tables sont créées automatiquement sans effacer les données existantes.
