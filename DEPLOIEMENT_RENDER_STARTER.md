# Déploiement Render Starter - QHSE Manager Pro

## Paramètres Render

- Type : Web Service
- Plan : Starter
- Build Command : `pip install -r requirements.txt`
- Start Command : `gunicorn app:app`
- Health Check Path : `/health`

## Variables d'environnement

- `CLOUD_MODE=true`
- `DATA_DIR=/var/data`
- `SECRET_KEY` générée automatiquement par Render

## Disque persistant

- Nom : `qhse-manager-data`
- Mount path : `/var/data`
- Taille : `1 GB`

La base SQLite, les exports, les uploads et les sauvegardes sont stockés dans `/var/data`.

## Vérification après déploiement

1. Ouvrir `/health` : le statut doit être `ok`.
2. Se connecter avec le compte super admin.
3. Créer un test : projet, utilisateur, produit.
4. Redémarrer le service Render.
5. Vérifier que les données sont toujours présentes.
6. Tester une sauvegarde manuelle depuis le module Sauvegardes.
