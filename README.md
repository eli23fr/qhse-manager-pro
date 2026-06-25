# QHSE Manager Pro - Chemical Register v2.0 CLOUD READY

Copyright © 2026 Kodjotse Eli ADIGBLI. Tous droits réservés.

## Version stable

Cette version repart sur une base propre et inclut :

- Super Admin, Admin Projet, Sous-traitant, Lecteur ;
- gestion des projets enrichie ;
- gestion des utilisateurs ;
- fiches sous-traitants enrichies ;
- fiches produits chimiques enrichies ;
- saisie mensuelle par produit validé ou produit non listé ;
- consolidation générale ;
- alertes QHSE ;
- journal des actions ;
- export Excel avec onglets Rapport Général, Alertes et À propos.

## Lancement Windows

1. Décompresser le ZIP dans un nouveau dossier.
2. Double-cliquer sur `INSTALLER_DEPENDANCES.bat`.
3. Double-cliquer sur `LANCER_APPLICATION.bat`.
4. Le navigateur s'ouvre automatiquement sur http://127.0.0.1:5000

## Identifiants Super Admin

Utilisateur : admin
Mot de passe : admin123

## Important

Les données sont conservées dans le fichier `qhse_chemical_register.db`.
Ne supprime pas ce fichier si tu veux conserver les données.


## Nouveautés v1.1

- Tableau de bord enrichi :
  - entrées totales ;
  - quantités utilisées ;
  - alertes critiques ;
  - alertes attention ;
  - top 10 produits stockés ;
  - top 10 produits utilisés ;
  - stock par sous-traitant ;
  - répartition par famille ;
  - évolution mensuelle.


## Nouveautés v2.0 CLOUD READY

- Préparation au déploiement Cloud.
- Ajout de `gunicorn`.
- Ajout de `Procfile`.
- Ajout de `render.yaml`.
- Ajout de `/health`.
- Ajout de `CLOUD_MODE`.
- Guide `DEPLOIEMENT_CLOUD.md`.

## Note

Cette version est destinée au premier test Cloud.  
Elle utilise encore SQLite. La version suivante préparera PostgreSQL.
