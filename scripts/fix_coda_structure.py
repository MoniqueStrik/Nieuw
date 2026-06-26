"""
Fix Coda tab structure per PROMPT:
- Section 6.1: exploitatie (IG rows 17-109, E>0), col A = grootboek
- Section 6.2: balans (IG rows 121-272, E>0), col A = "+"
- Section 6.3: sluitrekening hardcoded (A=2999999, B=86100)
Uses MATCH squeeze to eliminate blank rows.
"""
import openpyxl
from openpyxl.styles import Alignment

wb = openpyxl.load_workbook('Format_2026_begrotingswijziging_raad.xlsx', data_only=False)
wb.calculation.calcMode = 'auto'
wb.calculation.fullCalcOnLoad = True

ig = wb['Invullen groen']

# ── 1. IG col P (16): sequential counter for exploitatie rows 17-109 (E>0) ──────
# Already covers 17-272 - that's fine, we'll MATCH in P17:P109 range only
# Verify/update P to be correct for rows 17-109
# Current: =IF(E{r}>0,COUNTIF($E$17:E{r},">0"),0) - already correct

# ── 2. IG col S (19): sequential counter for balans rows 121-272 (E>0) ──────────
ig.column_dimensions['S'].hidden = True
ig.column_dimensions['S'].width = 5

for r in range(121, 273):
    ig.cell(r, 19).value = f'=IF(E{r}>0,COUNTIF($E$121:E{r},">0"),0)'

print('IG: added balans sequential counter in col S (rows 121-272)')

# ── 3. Rebuild Coda ──────────────────────────────────────────────────────────────
coda = wb['Coda']

# Hide helper cols
coda.column_dimensions['J'].hidden = True  # full omschrijving helper
coda.column_dimensions['L'].hidden = True  # row lookup helper

# K1 = prefix formula (keep existing if present, else set it)
prefix_formula = (
    "=IFERROR(LEFT('Invullen groen'!H4,"
    "FIND(\"/\",'Invullen groen'!H4)-1)&\"-\"&RIGHT('Invullen groen'!H4,2),\"\")"
)
if not coda.cell(1, 11).value:
    coda.cell(1, 11).value = prefix_formula

def idx(col_letter, h_ref):
    return f"INDEX('Invullen groen'!${col_letter}:${col_letter},{h_ref})"

def omschr_full(h):
    """Full omschrijving formula with prefix, substitutions, lowercase."""
    e_i = idx('E', h)
    o_i = idx('O', h)
    n_i = idx('N', h)
    g_i = idx('G', h)
    return (
        f'=IF({h}>0,IFERROR($K$1&" "&TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE('
        f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE('
        f'SUBSTITUTE(LOWER(IF(LEFT(TEXT({e_i},"0"),1)="6",{o_i},'
        f'IF({n_i}>0,{g_i},{o_i}))),'
        f'"2024",""),"2025",""),"2026",""),"2027",""),"2028",""),"2029",""),'
        f'"taakveld ",""),"installaties","inst"),"ruwbouw","ruwb"),'
        f'"afbouw","afb"),"bijstelling","bijst"),"investeringen","invest")),""),")'
    )

def omschr_short(r):
    """Word-truncated omschrijving from J col same row (max 36 chars)."""
    j = f'J{r}'
    return (
        f'=IFERROR(IF(LEN({j})<=36,{j},'
        f'IFERROR(LEFT({j},FIND(CHAR(1),SUBSTITUTE(LEFT({j},36)," ",CHAR(1),'
        f'LEN(LEFT({j},36))-LEN(SUBSTITUTE(LEFT({j},36)," ",""))))-1),'
        f'LEFT({j},36))),"")'
    )

# ── Section 6.1: exploitatie rows (Coda rows 5-104) ─────────────────────────────
# MATCH finds Nth row with E>0 in IG P17:P109 (sections A+B)
EXPL_START = 5
EXPL_END = 104  # 100 slots (max 68+32=100 for sections A+B)

for r in range(EXPL_START, EXPL_END + 1):
    h = f'$L{r}'
    n = r - EXPL_START + 1

    # L col: IG row number for Nth exploitatie entry
    coda.cell(r, 12).value = (
        f"=IFERROR(MATCH(ROW()-{EXPL_START-1},'Invullen groen'!$P$17:$P$109,0)+16,0)"
    )

    # c1 Grootboek (IG col E)
    coda.cell(r, 1).value = f"=IFERROR(IF({h}>0,{idx('E',h)},\"\"),\"\")"

    # c2 Econ.Categorie (IG col F)
    coda.cell(r, 2).value = (
        f"=IFERROR(IF({h}>0,IF({idx('F',h)}>0,{idx('F',h)},\"\"),\"\"),\"\")"
    )

    # c3 Verbijzondering (IG col N)
    coda.cell(r, 3).value = (
        f"=IFERROR(IF({h}>0,IF({idx('N',h)}>0,{idx('N',h)},\"\"),\"\"),\"\")"
    )

    # c4 Verdere verbijz.: empty per PROMPT
    coda.cell(r, 4).value = None

    # c5 Omschrijving (word-truncated from col J same row)
    coda.cell(r, 5).value = omschr_short(r)

    # c6 Bedrag 2026 (IG col H)
    coda.cell(r, 6).value = (
        f"=IFERROR(IF({h}>0,IF({idx('H',h)}=0,\"\",INT({idx('H',h)})),\"\"),\"\")"
    )

    # c7 MJR 2027 (IG col I)
    coda.cell(r, 7).value = (
        f"=IFERROR(IF({h}>0,IF({idx('I',h)}=0,\"\",INT({idx('I',h)})),\"\"),\"\")"
    )

    # c8 MJR 2028 (IG col J)
    coda.cell(r, 8).value = (
        f"=IFERROR(IF({h}>0,IF({idx('J',h)}=0,\"\",INT({idx('J',h)})),\"\"),\"\")"
    )

    # c9 MJR 2029 (IG col K)
    coda.cell(r, 9).value = (
        f"=IFERROR(IF({h}>0,IF({idx('K',h)}=0,\"\",INT({idx('K',h)})),\"\"),\"\")"
    )

    # c10 Full omschrijving helper
    coda.cell(r, 10).value = omschr_full(h)

print(f'Coda: section 6.1 written (rows {EXPL_START}-{EXPL_END})')

# ── Section 6.2: balans rows (Coda rows 105-250) ─────────────────────────────────
# MATCH finds Nth row with E>0 in IG S121:S272 (sections C+D+E)
BAL_START = 105
BAL_END = 250  # 146 slots (max 47+46+53=146 for sections C+D+E)

for r in range(BAL_START, BAL_END + 1):
    h = f'$L{r}'
    n = r - BAL_START + 1

    # L col: IG row number for Nth balans entry
    coda.cell(r, 12).value = (
        f"=IFERROR(MATCH(ROW()-{BAL_START-1},'Invullen groen'!$S$121:$S$272,0)+120,0)"
    )

    # c1 Kolom A = "+" (gecentreerd) for balans
    cell_a = coda.cell(r, 1)
    cell_a.value = f'=IF({h}>0,"+","")'
    cell_a.alignment = Alignment(horizontal='center')

    # c2 Econ.Categorie (IG col F)
    coda.cell(r, 2).value = (
        f"=IFERROR(IF({h}>0,IF({idx('F',h)}>0,{idx('F',h)},\"\"),\"\"),\"\")"
    )

    # c3 Verbijzondering (IG col N)
    coda.cell(r, 3).value = (
        f"=IFERROR(IF({h}>0,IF({idx('N',h)}>0,{idx('N',h)},\"\"),\"\"),\"\")"
    )

    # c4 empty
    coda.cell(r, 4).value = None

    # c5 Omschrijving (word-truncated from col J same row)
    coda.cell(r, 5).value = omschr_short(r)

    # c6-c9 Bedragen
    for ci, ig_col in [(6,'H'), (7,'I'), (8,'J'), (9,'K')]:
        coda.cell(r, ci).value = (
            f"=IFERROR(IF({h}>0,IF({idx(ig_col,h)}=0,\"\",INT({idx(ig_col,h)})),\"\"),\"\")"
        )

    # c10 Full omschrijving helper
    coda.cell(r, 10).value = omschr_full(h)

print(f'Coda: section 6.2 written (rows {BAL_START}-{BAL_END})')

# ── Section 6.3: sluitrekening (Coda row 251) ────────────────────────────────────
SLUIT_ROW = 251

coda.cell(SLUIT_ROW, 1).value = 2999999
coda.cell(SLUIT_ROW, 2).value = 86100
coda.cell(SLUIT_ROW, 3).value = None
coda.cell(SLUIT_ROW, 4).value = None
coda.cell(SLUIT_ROW, 5).value = f'=$K$1&" sluitrekening balans"'
coda.cell(SLUIT_ROW, 6).value = f'=IFERROR(-SUM(F{BAL_START}:F{BAL_END}),0)'
coda.cell(SLUIT_ROW, 7).value = f'=IFERROR(-SUM(G{BAL_START}:G{BAL_END}),0)'
coda.cell(SLUIT_ROW, 8).value = f'=IFERROR(-SUM(H{BAL_START}:H{BAL_END}),0)'
coda.cell(SLUIT_ROW, 9).value = f'=IFERROR(-SUM(I{BAL_START}:I{BAL_END}),0)'
coda.cell(SLUIT_ROW, 10).value = None
coda.cell(SLUIT_ROW, 12).value = None

print(f'Coda: section 6.3 sluitrekening written (row {SLUIT_ROW})')

# ── Clear rows 252+ (leftover from old structure) ────────────────────────────────
for r in range(SLUIT_ROW + 1, 280):
    for c in range(1, 13):
        coda.cell(r, c).value = None

print('Coda: cleared trailing rows 252-279')

wb.save('Format_2026_begrotingswijziging_raad.xlsx')
print('\nSaved.')
