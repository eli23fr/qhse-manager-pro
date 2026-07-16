# QHSE Manager Pro — Mise à jour v2.6

## Fonctionnalités ajoutées

- Modification et suppression des saisies d'effectifs, selon le rôle et le périmètre autorisé.
- Bouton Suspendre / Réactiver dans la gestion des utilisateurs.
- Déconnexion automatique après 30 minutes d'inactivité, configurable avec `SESSION_TIMEOUT_MINUTES`.
- Colonne Utilisation ajoutée dans la liste des saisies produits et dans les deux feuilles du rapport Excel.
- Affichage mobile amélioré : viewport, navigation horizontale, tableaux défilables et formulaires adaptés.
- Pagination ajoutée aux listes à fort volume : utilisateurs, produits, saisies produits, effectifs et journal d'audit.
- Version affichée : v2.6 MOBILITÉ + CONTRÔLE + PAGINATION.

## Variable Render facultative

La durée de session est de 30 minutes par défaut. Pour la modifier dans Render :

- Key : `SESSION_TIMEOUT_MINUTES`
- Value : nombre de minutes, par exemple `45`

## Déploiement

Remplacer les fichiers du dépôt par ceux de cette version, puis exécuter :

```bash
git add .
git commit -m "Mise à jour v2.6 mobilité contrôle pagination"
git push origin main
```

Le disque persistant et les variables `DATA_DIR=/var/data` et `CLOUD_MODE=true` restent inchangés.
