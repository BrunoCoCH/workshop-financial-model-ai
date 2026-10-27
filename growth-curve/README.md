# Atelier — Courbes de croissance (`growth-curve`)

Exercice pour préparer les **hypothèses multi-années** d’un plan financier au format de la skill [`financial-model-platinn-romandee`](https://github.com/BrunoCoCH/financial-model-platinn-romandee) (mode services).

Vous ne construisez pas tout le modèle Excel ici : vous travaillez les **courbes** (année 1 → année 5) que la skill interpolera / utilisera dans l’interview et le JSON d’inputs.

## Contenu

| Fichier | Rôle |
|---|---|
| `growth-curve.xlsx` | Cas Alpes Conseil — courbes à compléter |
| `build_exercice_excel.py` | Régénère le classeur |

## Prérequis

- Excel
- Option A : skill **financial-model-platinn-romandee** (Cursor / Claude)
- Option B : Claude in Excel ou Cursor sans la skill (interpolation manuelle guidée)

## Déroulé (25–30 min)

1. Ouvrez `growth-curve.xlsx`, lisez **Consignes**.
2. Feuille **Courbes** : Y1 et Y5 sont donnés ; complétez Y2–Y4 (linéaire ou avec rupture justifiée).
3. Feuille **Recap** : validez cohérence (clients, capacité, marge indicative).
4. Passez le résultat à la skill : *« Voici mes courbes — lance l’interview / remplis le JSON services à partir de ça »*, ou exportez vers un `*_inputs.json`.

## Lien avec la skill

Les lignes du classeur correspondent aux drivers services : prix, volume (clients, churn, demande projets), capacité (FTE, util.), cash. Voir le `SKILL.md` de financial-model-platinn-romandee.
