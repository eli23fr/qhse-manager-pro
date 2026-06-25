# Déploiement Cloud - QHSE Manager Pro v2.0

## Objectif

Mettre l'application en ligne pour permettre aux sous-traitants et administrateurs de se connecter via un lien web.

## Fichiers ajoutés

- `Procfile`
- `render.yaml`
- `gunicorn`
- route `/health`
- variable `CLOUD_MODE=true`

## Test local

```bash
pip install -r requirements.txt
python app.py
```

Ouvrir :

```text
http://127.0.0.1:5000
```

## Déploiement Render

1. Créer un dépôt GitHub.
2. Envoyer tous les fichiers du dossier dans ce dépôt.
3. Aller sur Render.
4. Créer un nouveau Web Service.
5. Connecter le dépôt GitHub.
6. Laisser Render utiliser `render.yaml`.
7. Vérifier :
   - `CLOUD_MODE=true`
   - `SECRET_KEY` générée automatiquement.
8. Déployer.

## Connexion

Identifiants par défaut :

```text
admin
admin123
```

## Important

Cette version est faite pour le premier test Cloud.  
Elle utilise encore SQLite. Pour un usage durable multi-utilisateurs, la version v2.1 migrera vers PostgreSQL.
