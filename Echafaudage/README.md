# Plan d'échafaudage MULTISECU+ 2 — GAILLON

Outil de modélisation d'un échafaudage multidirectionnel Altrad MULTISECU+ 2 :
plan 2D coté et étiqueté, vue 3D, quantitatif estimatif.

Ouvrir `plan-echafaudage-multisecu.html` dans un navigateur (Chrome, Edge, Firefox).
La vue 3D charge la bibliothèque three.js depuis internet ; le plan 2D fonctionne hors ligne.
Le plan en cours est conservé automatiquement dans le navigateur ; pour le transmettre,
utiliser **Exporter › Fichier projet (.json)** puis **Plans › Importer** chez le destinataire.

## Utilisation

1. **Façade / mur** : tracer les façades du bâtiment (segments enchaînés, Échap pour terminer).
2. **File de travées** : cliquer le départ puis l'arrivée le long de la façade.
   L'outil s'accroche à 0,20 m de la façade et découpe la longueur en travées de 3,00 / 2,50 / 2,00 / 1,50 m.
   Les travées seules (dont la travée d'angle 0,80) s'accrochent aux poteaux existants.
3. **Accessoires** : plancher à trappe + échelle, garde-corps côté façade, filet / bâche,
   diagonale, amarrage, consoles 0,40 / 0,70 — cliquer sur une travée.
   Garde-corps extérieurs, d'extrémité et plinthes sont posés automatiquement.
4. **Annotations** : zones, cotes libres, étiquettes.
5. Sélectionner un élément pour le modifier ; **Supprimer** (ou touche Suppr) pour l'effacer,
   Ctrl+Z pour annuler.
6. Bouton **Plan 2D / Volume 3D** pour basculer de vue.
7. **Exporter** : plan 2D coté (.png avec cartouche), vue 3D (.png), quantitatif (.csv), fichier projet (.json).

Raccourcis : V sélection · R pivoter 90° · F changer de côté façade · flèches pour décaler ·
Alt pour désactiver l'accrochage · double-clic sur une travée pour sélectionner toute la file.

## Catalogue

Les dimensions (largeur 0,80, longueurs de travées, hauteurs de niveau, consoles) sont regroupées
en tête du script, dans l'objet `CAT`. Elles reprennent la notice MULTISECU+ 2
(Altrad Plettac Mefran, réf. 111146) et sont à vérifier sur le parc réellement disponible.

Le quantitatif et les « points de vigilance » sont des aides à la discussion : ils ne remplacent
ni la notice du fabricant, ni le plan de montage, ni la note de calcul.
