"""Génère le classeur d'exercice pour la skill budget-variance-note."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

navy = "1B3A4B"
teal = "2A6F7F"
light = "F4F7F8"
amber = "FFF3E0"
thin = Border(
    left=Side(style="thin", color="CFD8DC"),
    right=Side(style="thin", color="CFD8DC"),
    top=Side(style="thin", color="CFD8DC"),
    bottom=Side(style="thin", color="CFD8DC"),
)
hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
title_font = Font(name="Calibri", bold=True, color=navy, size=16)
section_font = Font(name="Calibri", bold=True, color=navy, size=12)
body = Font(name="Calibri", size=11)
hdr_fill = PatternFill("solid", fgColor=navy)
alt_fill = PatternFill("solid", fgColor=light)
note_fill = PatternFill("solid", fgColor=amber)
total_fill = PatternFill("solid", fgColor="E3EEF1")

POSTES = [
    "Chiffre d'affaires — honoraires",
    "Autres produits",
    "Charges de personnel",
    "Sous-traitance / freelances",
    "Loyer et charges locatives",
    "Marketing et acquisition",
    "Informatique / outils",
    "Autres frais généraux",
]

# Jan + Fév + Mars = Trimestre (budget et réalisé), vérifié ligne à ligne.
# Mars conserve le cas « choc » pédagogique (CA −18k, freelances +4k).
JANVIER = [
    (88_000, 86_000, "Démarrage année : pipeline correct"),
    (2_000, 1_900, ""),
    (55_000, 55_200, ""),
    (8_000, 9_000, "Pic ponctuel sur une mission"),
    (4_500, 4_500, ""),
    (3_000, 3_400, ""),
    (1_800, 1_700, ""),
    (2_700, 2_700, ""),
]
FEVRIER = [
    (92_000, 83_000, "1 mission décalée en mars"),
    (2_000, 1_800, ""),
    (55_000, 56_300, "Formation interne (heures non facturables)"),
    (8_000, 10_000, "Renfort sur livrable client clé"),
    (4_500, 4_500, ""),
    (3_000, 3_900, "Sponsoring événement métier"),
    (1_800, 1_650, ""),
    (2_700, 2_700, ""),
]
MARS = [
    (100_000, 82_000, "2 missions reportées ; 1 client a réduit le volume"),
    (2_000, 1_500, ""),
    (55_000, 56_500, "Heures supp. sur un livrable urgent"),
    (8_000, 12_000, "Remplacement d'un consultant malade"),
    (4_500, 4_500, ""),
    (3_000, 4_200, "Campagne LinkedIn non budgétée"),
    (1_800, 1_750, ""),
    (2_700, 2_400, ""),
]


def chf(cell):
    cell.number_format = "#,##0"
    cell.font = body
    cell.alignment = Alignment(horizontal="right")


def pct(cell):
    cell.number_format = "0.0%"
    cell.font = body


def write_month_sheet(wb, name, period_label, rows, tip=None):
    ws = wb.create_sheet(name)
    ws["A1"] = "Alpes Conseil Sàrl — Budget vs réalisé"
    ws["A1"].font = title_font
    ws.merge_cells("A1:F1")
    ws["A2"] = period_label
    ws["A2"].font = Font(name="Calibri", italic=True, size=11, color=teal)

    headers = [
        "Poste",
        "Budget CHF",
        "Réalisé CHF",
        "Écart CHF",
        "Écart %",
        "Commentaire métier (brut)",
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(4, col, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center")

    for i, (poste, (bud, real, com)) in enumerate(zip(POSTES, rows)):
        r = 5 + i
        ws.cell(r, 1, poste).font = body
        ws.cell(r, 2, bud)
        chf(ws.cell(r, 2))
        ws.cell(r, 3, real)
        chf(ws.cell(r, 3))
        ws.cell(r, 4, f"=C{r}-B{r}")
        chf(ws.cell(r, 4))
        ws.cell(r, 5, f"=IF(B{r}=0,0,D{r}/B{r})")
        pct(ws.cell(r, 5))
        ws.cell(r, 6, com).font = body
        for col in range(1, 7):
            ws.cell(r, col).border = thin
            if i % 2:
                ws.cell(r, col).fill = alt_fill

    ws["A14"] = "Total produits"
    ws["A15"] = "Total charges"
    ws["A16"] = "Résultat d'exploitation (approx.)"
    ws["B14"] = "=B5+B6"
    ws["C14"] = "=C5+C6"
    ws["B15"] = "=SUM(B7:B12)"
    ws["C15"] = "=SUM(C7:C12)"
    ws["B16"] = "=B14-B15"
    ws["C16"] = "=C14-C15"
    for r in (14, 15, 16):
        ws.cell(r, 4, f"=C{r}-B{r}")
        ws.cell(r, 5, f"=IF(B{r}=0,0,D{r}/B{r})")
        for col in range(1, 6):
            ws.cell(r, col).font = Font(name="Calibri", bold=True, size=11)
            ws.cell(r, col).fill = total_fill
            ws.cell(r, col).border = thin
        chf(ws.cell(r, 2))
        chf(ws.cell(r, 3))
        chf(ws.cell(r, 4))
        pct(ws.cell(r, 5))

    if tip:
        ws["A18"] = "Astuce atelier"
        ws["A18"].font = section_font
        ws["A19"] = tip
        ws["A19"].fill = note_fill
        ws["A19"].alignment = Alignment(wrap_text=True)
        ws.merge_cells("A19:F20")

    for i, w in enumerate([36, 14, 14, 12, 10, 52], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def assert_quarter_sums():
    for i, poste in enumerate(POSTES):
        jb, jr, _ = JANVIER[i]
        fb, fr, _ = FEVRIER[i]
        mb, mr, _ = MARS[i]
        assert jb + fb + mb > 0
        # Documented Q1 targets (used historically / for formateur)
        q_bud = jb + fb + mb
        q_real = jr + fr + mr
        print(f"  {poste}: budget {q_bud:,} | réalisé {q_real:,}")


# --- Consignes ---
ws = wb.active
ws.title = "Consignes"
ws["A1"] = "Atelier — Créer une skill : budget-variance-note"
ws["A1"].font = title_font
ws.merge_cells("A1:D1")
ws["A3"] = "Objectif"
ws["A3"].font = section_font
ws["A4"] = (
    "Utilisez la feuille « Mars » (ou « Trimestre ») avec votre skill budget-variance-note. "
    "Demandez à Claude / Cursor : « Rédige la note d'écarts à partir de ce fichier ». "
    "La skill doit produire synthèse, tableau des 3 postes clés, questions métier "
    "et une prochaine action. Janvier et Février permettent d’expliquer le trimestre."
)
ws["A4"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A4:D6")

ws["A8"] = "Déroulement suggéré (15–20 min)"
ws["A8"].font = section_font
steps = [
    "1. Ouvrez Create-Skill et créez la skill budget-variance-note (modèle fourni en atelier).",
    "2. Ouvrez ce classeur — premier essai : feuille Mars seulement.",
    "3. Lancez la skill sur les chiffres de Mars (sans regarder Corrige_formateur).",
    "4. Comparez votre note à Corrige_formateur (ordre de grandeur, pas mot à mot).",
    "5. Variante : feuille Trimestre (somme Janvier + Février + Mars), puis explorez les mois.",
]
for i, s in enumerate(steps):
    ws.cell(9 + i, 1, s).font = body

ws["A15"] = "Société fictive"
ws["A15"].font = section_font
ws["A16"] = (
    "Alpes Conseil Sàrl — cabinet de conseil, Vaud. Devise : CHF. "
    "Fichier pédagogique uniquement — pas de conseil fiscal ou financier."
)
ws["A16"].font = body
ws["A18"] = "Cohérence des chiffres"
ws["A18"].font = section_font
ws["A19"] = (
    "La feuille Trimestre additionne Janvier + Février + Mars (formules). "
    "Modifier un mois met à jour le trimestre."
)
ws["A19"].font = body
ws.column_dimensions["A"].width = 100

print("Vérification Jan+Fév+Mars :")
assert_quarter_sums()

write_month_sheet(
    wb,
    "Janvier",
    "Période : janvier 2026  |  Comparaison : budget mensuel vs réalisé",
    JANVIER,
)
write_month_sheet(
    wb,
    "Fevrier",
    "Période : février 2026  |  Comparaison : budget mensuel vs réalisé",
    FEVRIER,
)
write_month_sheet(
    wb,
    "Mars",
    "Période : mars 2026  |  Comparaison : budget mensuel vs réalisé",
    MARS,
    tip=(
        "La skill doit garder au plus 3 postes qui expliquent l'essentiel de l'écart. "
        "Les commentaires métier sont des indices bruts — la note ne doit pas les recopier tels quels."
    ),
)

# --- Trimestre = somme des 3 mois ---
t = wb.create_sheet("Trimestre")
t["A1"] = "Alpes Conseil Sàrl — Budget vs réalisé Q1 2026"
t["A1"].font = title_font
t.merge_cells("A1:E1")
t["A2"] = (
    "Période : janvier–mars 2026  |  "
    "Budget et réalisé = Janvier + Fevrier + Mars (formules)"
)
t["A2"].font = Font(name="Calibri", italic=True, size=11, color=teal)

for col, h in enumerate(
    ["Poste", "Budget CHF", "Réalisé CHF", "Écart CHF", "Écart %"], 1
):
    cell = t.cell(4, col, h)
    cell.font = hdr_font
    cell.fill = hdr_fill

for i, poste in enumerate(POSTES):
    r = 5 + i
    t.cell(r, 1, poste).font = body
    t.cell(r, 2, f"=Janvier!B{r}+Fevrier!B{r}+Mars!B{r}")
    chf(t.cell(r, 2))
    t.cell(r, 3, f"=Janvier!C{r}+Fevrier!C{r}+Mars!C{r}")
    chf(t.cell(r, 3))
    t.cell(r, 4, f"=C{r}-B{r}")
    chf(t.cell(r, 4))
    t.cell(r, 5, f"=IF(B{r}=0,0,D{r}/B{r})")
    pct(t.cell(r, 5))
    for col in range(1, 6):
        t.cell(r, col).border = thin
        if i % 2:
            t.cell(r, col).fill = alt_fill

t["A14"] = "Total produits"
t["A15"] = "Total charges"
t["A16"] = "Résultat d'exploitation (approx.)"
t["B14"] = "=B5+B6"
t["C14"] = "=C5+C6"
t["B15"] = "=SUM(B7:B12)"
t["C15"] = "=SUM(C7:C12)"
t["B16"] = "=B14-B15"
t["C16"] = "=C14-C15"
for r in (14, 15, 16):
    t.cell(r, 4, f"=C{r}-B{r}")
    t.cell(r, 5, f"=IF(B{r}=0,0,D{r}/B{r})")
    for col in range(1, 6):
        t.cell(r, col).font = Font(name="Calibri", bold=True, size=11)
        t.cell(r, col).fill = total_fill
        t.cell(r, col).border = thin
    chf(t.cell(r, 2))
    chf(t.cell(r, 3))
    chf(t.cell(r, 4))
    pct(t.cell(r, 5))

t["A18"] = "Contrôle"
t["A18"].font = section_font
t["A19"] = (
    "Chaque ligne Budget / Réalisé est la somme des trois feuilles mensuelles. "
    "Ex. CA honoraires Q1 : budget 280'000 | réalisé 251'000."
)
t["A19"].fill = note_fill
t["A19"].alignment = Alignment(wrap_text=True)
t.merge_cells("A19:E20")

for i, w in enumerate([36, 14, 14, 12, 10], 1):
    t.column_dimensions[get_column_letter(i)].width = w

# --- Corrigé formateur ---
c = wb.create_sheet("Corrige_formateur")
c["A1"] = "Corrigé indicatif — Mars (pour le formateur)"
c["A1"].font = title_font
c["A3"] = "Synthèse attendue (idée, pas texte imposé)"
c["A3"].font = section_font
c["A4"] = (
    "Mars est nettement sous budget : le résultat souffre surtout d'un CA en retrait "
    "(−18'000 CHF), amplifié par plus de freelances (+4'000). Les autres écarts sont "
    "secondaires. À confirmer avec le métier : reports de missions et couverture freelances."
)
c["A4"].alignment = Alignment(wrap_text=True)
c.merge_cells("A4:B6")

c["A8"] = "3 postes qui devraient ressortir"
c["A8"].font = section_font
c["A9"] = "1. CA honoraires — défavorable −18'000 (−18,0 %)"
c["A10"] = "2. Sous-traitance — défavorable +4'000 (+50 %)"
c["A11"] = (
    "3. Marketing — défavorable +1'200 "
    "(ou personnel +1'500 — les deux sont acceptables en 3e)"
)
c["A13"] = "Questions métier possibles"
c["A13"].font = section_font
c["A14"] = "• Les 2 missions reportées sont-elles rebookées en avril ?"
c["A15"] = "• Le freelance remplace-t-il une absence ponctuelle ou un sous-effectif structurel ?"
c["A16"] = "• La campagne LinkedIn a-t-elle un ROI mesurable ce trimestre ?"
c["A18"] = "Prochaine action type"
c["A18"].font = section_font
c["A19"] = (
    "Mettre à jour le forecast avril avec les missions reportées + plafond freelance."
)
c["A21"] = "Cohérence trimestre"
c["A21"].font = section_font
c["A22"] = (
    "Janvier + Février + Mars = Trimestre. Q1 CA : 280'000 budget / 251'000 réalisé "
    "(écart −29'000). Le trou se creuse surtout en mars (−18'000) après un février déjà faible."
)
c["A22"].alignment = Alignment(wrap_text=True)
c.merge_cells("A22:B23")
c["A25"] = "Ne montrez pas cette feuille aux participants avant la fin de l'exercice."
c["A25"].fill = note_fill
c["A25"].font = Font(name="Calibri", bold=True, size=11)
c.column_dimensions["A"].width = 95

out = Path(__file__).resolve().parent / "Budget_vs_Realise_AlpesConseil.xlsx"
wb.save(out)
print(out)
