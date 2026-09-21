# Macro « Ajouter une réponse » — GAILLON, Suivi des VISAS

Deux éléments à coller dans le classeur `GAILLON - Suivi des VISAS.xlsm`
(Alt + F11 pour ouvrir l'éditeur VBA) :

| Fichier | Où le coller |
|---|---|
| `Module_Reponses.bas` | Un **nouveau module standard** (Insertion > Module). Ne pas recopier la 1re ligne `Attribute VB_Name = ...`. |
| `frmReponseVisa.txt` | Un **UserForm vide** (Insertion > UserForm) renommé `frmReponseVisa` dans la fenêtre Propriétés (F4), puis clic droit > Code. |

Aucun contrôle n'est à dessiner : listes déroulantes, zone de texte et boutons
sont créés automatiquement à l'ouverture du formulaire.

Bouton : Insertion > Formes (ou Développeur > Insérer > Bouton), clic droit >
**Affecter une macro** > `AjouterReponseVisa`.

## Ce que fait la macro

1. Choix du **numéro de fiche** (colonne D) puis de son **indice** (colonne E) —
   seuls les indices existants pour la fiche choisie sont proposés.
2. Choix de **l'entité qui répond** — liste construite à partir des en-têtes de
   la ligne 4 (blocs AB, AE, AH, AK, AN, AQ, AT, AW, AZ, BC, BF).
3. Choix du **visa** (liste lue dans `Données!E47:E52` : F, FAO, S, R, ATT, EXP)
   et de la **date** (par défaut la date du jour).
4. Saisie ou **copier-coller** de la réponse dans une zone de texte multiligne.
5. Écriture dans les 3 cellules du bloc de l'entité : `Visa`, `Date`, `Commentaire`.

La réponse déjà enregistrée pour l'entité sélectionnée est rechargée dans le
formulaire et un avertissement s'affiche avant tout remplacement.

## Adaptation

Toutes les positions (onglets, lignes d'en-tête, colonnes) sont regroupées dans
les constantes en haut de `Module_Reponses.bas`.
