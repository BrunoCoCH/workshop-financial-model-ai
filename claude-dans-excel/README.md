# Atelier — Utiliser le plugin Claude dans Excel

Exercice pratique pour financiers : faire travailler **Claude for Excel** (add-in Microsoft 365) sur un vrai classeur — formules, synthèse, commentaires — **sans quitter Excel**.

## Contenu

| Fichier | Rôle |
|---|---|
| `Pilotage_Mensuel_AlpesConseil.xlsx` | Classeur d’exercice (données + missions) |
| `build_exercice_excel.py` | Régénère le classeur |

## Prérequis

- Excel Microsoft 365 (desktop ou web)
- Add-in **Claude** installé (Organisation → Compléments, ou [claude.ai / Excel](https://claude.ai))
- Compte Anthropic avec accès à Claude in Excel

## Déroulé (20–25 min)

1. Ouvrez `Pilotage_Mensuel_AlpesConseil.xlsx` dans Excel.
2. Ouvrez le panneau **Claude** (onglet Accueil / Compléments).
3. Suivez les **4 missions** de la feuille `Missions` — une à la fois, dans l’ordre.
4. Vérifiez chaque résultat avant de passer à la suivante.
5. En fin d’atelier : comparez à la feuille `Corrige_formateur` (formateur uniquement).

## Ce que l’exercice entraîne

- Demander des **formules Excel** (pas des valeurs figées) et les faire écrire dans la grille
- Faire lire une plage nommée / un tableau
- Obtenir un **commentaire de gestion** dans une cellule ou une zone dédiée
- Corriger une erreur de logique (signe d’écart, total)

## Prompt d’ouverture type (Mission 1)

Copiez-collez depuis la feuille `Missions`, ou :

> Sur la feuille Donnees, colonnes Budget et Realise sont remplies.  
> Dans la colonne Ecart (D), écris des **formules** Realise − Budget.  
> Dans la colonne Ecart_pct (E), écris des formules d’écart en %.  
> Ne remplace pas les montants existants.
