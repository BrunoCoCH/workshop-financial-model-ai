# Atelier — Créer une skill (`budget-variance-note`)

Créez une skill hyper simple qui transforme un tableau budget / réalisé en note COMEX d'une demi-page.

## Contenu

| Fichier | Rôle |
|---|---|
| `SKILL.md` | Modèle de skill à créer (ou à coller via Create-Skill) |
| `Budget_vs_Realise_AlpesConseil.xlsx` | Chiffres fictifs (Alpes Conseil Sàrl) |
| `build_exercice_excel.py` | Régénère le classeur (`pip install openpyxl` puis `python build_exercice_excel.py`) |

## Déroulé (15–20 min)

1. Dans Cursor : *« Utilise Create-Skill pour créer la skill budget-variance-note »* — partez du `SKILL.md` de ce dossier.
2. Ouvrez le classeur ; travaillez d'abord la feuille **Mars** (ignorez **Corrige_formateur**).
3. Lancez : *« Rédige la note d'écarts à partir de ce fichier »*.
4. Comparez à **Corrige_formateur** (ordre de grandeur, pas mot à mot).
5. Variante : feuille **Trimestre**.

## Prompt type pour Create-Skill

> Crée une skill `budget-variance-note` pour des financiers suisses. Elle rédige une note courte d'écarts budget vs réalisé. Une question à la fois si des chiffres manquent. Sortie : synthèse 2 phrases, tableau max 3 postes, 3 points, questions métier, 1 prochaine action. CHF avec apostrophes. N'invente pas les causes. Inclus un mini-exemple CA 100k vs 82k.
