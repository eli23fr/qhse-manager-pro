# QHSE Manager Pro v2.7

## Nouveautés

- Pagination fixe à 20 lignes par page.
- Champ **Utilisation** saisi directement par les sous-traitants dans la saisie produits.
- La valeur d'utilisation alimente la saisie et met à jour la fiche produit associée.
- Les administrateurs conservent les droits de modification et de suppression selon leur périmètre.
- Nouveau module **Rapport HSE journalier sous-traitant**.
- Modèle initial construit à partir du rapport fourni : indicateurs proactifs/réactifs, main-d'œuvre, urgence, circulation, travaux dangereux, matières dangereuses, bruit, biodiversité, patrimoine, eau, déchets.
- Les administrateurs peuvent ajouter, modifier ou désactiver des lignes et des colonnes du rapport.
- Export Excel filtrable par projet, sous-traitant et période.
- Statistiques par jour, semaine ou mois.
- Trois palettes Excel : Bleu professionnel, Vert HSE et Sable chantier.
- Trois thèmes d'interface : Bleu, Vert HSE et Sable.
- Compatibilité mobile conservée et renforcée.

## Déploiement

Remplacer le contenu du dépôt local par cette version, sans supprimer le dossier `.git`, puis :

```bash
git add .
git commit -m "Mise à jour QHSE Manager Pro v2.7"
git push origin main
```

Render redéploiera automatiquement. Les nouvelles tables sont créées au démarrage dans la base SQLite persistante existante.

## Configuration Render inchangée

```text
DATA_DIR=/var/data
CLOUD_MODE=true
```
