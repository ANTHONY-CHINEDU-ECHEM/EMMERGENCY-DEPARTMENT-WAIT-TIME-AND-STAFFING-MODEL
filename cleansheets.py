"""Sheet by sheet data cleaning procedure writers (C1 to C11)."""
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.styles import Alignment
from .xl import *

PASS_RULES = lambda ws, rng: (
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"PASS"'], fill=fill(GREEN_L), font=F(9, True, GREEN))),
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"FAIL"'], fill=fill(RED_L), font=F(9, True, RED))),
)


class Ctx:
    def __init__(self, wb, cl, raw, clean, n_raw):
        self.wb, self.cl, self.raw, self.clean = wb, cl, raw, clean
        self.n_raw = n_raw                    # data rows in Raw Data
        self.n_clean = len(clean)
        self.rawcol = {c: L(i + 1) for i, c in enumerate(raw.columns)}
        self.cleancol = {c: L(i + 1) for i, c in enumerate(clean.columns)}

    def rr(self, c):
        x = self.rawcol[c]
        return f"'Raw Data'!${x}$2:${x}${self.n_raw + 1}"

    def cr(self, c):
        x = self.cleancol[c]
        return f"'Clean Data'!${x}$2:${x}${self.n_clean + 1}"


def proc_header(ws, row, purpose, method, excel_equiv):
    for label, text in (("Objective", purpose), ("Procedure", method), ("Excel / Power Query equivalent", excel_equiv)):
        ws.cell(row, 2, label).font = F(9, True, TEAL)
        ws.cell(row, 2).alignment = Alignment(vertical="top")
        row = note(ws, row, 3, text, span=11, size=9, color=INK)
    return row + 1


def C1(ctx):
    ws = ctx.wb.create_sheet("C1 Data Profile")
    setup(ws, "C1  Data Profile of the Raw Export",
          "Step 1 of the cleaning procedure. Baseline measurement of every column before any change is made.", cols=13)
    for k, w in zip("BCDEFGHIJKLMN", [30, 8, 11, 11, 11, 10, 11, 12, 12, 12, 12, 50, 4]):
        ws.column_dimensions[k].width = w
    r = proc_header(ws, 5,
                    "Quantify completeness, cardinality and value ranges for all columns so that every later cleaning rule is justified by evidence rather than assumption.",
                    "Each column in Raw Data is measured for non blank count, blank count, distinct raw values and distinct values after trimming, removing trailing periods, replacing underscores and ignoring case. "
                    "A large gap between the two distinct counts proves the column has formatting variants. Numeric columns receive a raw minimum and maximum to expose impossible values.",
                    "Power Query: View > Column profile, Column quality and Column distribution. Live counts on this sheet use COUNTA and ROWS against the Raw Data sheet and recalculate on open.")
    cl = ctx.cl
    issues = {}
    tm = cl.text_map
    for c, g in tm[tm.IssueType != "Canonical form"].groupby("Column"):
        issues.setdefault(c, []).append(f"{len(g)} formatting variants ({int(g.Records.sum())} records)")
    for c in cl.cfg["bool"]:
        issues.setdefault(c, []).append("12 different TRUE/FALSE encodings")
    for c in cl.cfg["dates"]:
        issues.setdefault(c, []).append("Stored as text in 4 date formats")
    nr = cl.numeric_rules
    for _, x in nr.iterrows():
        if x.BelowMin + x.AboveMax:
            issues.setdefault(x.Column, []).append(f"{x.BelowMin} below minimum, {x.AboveMax} above maximum")
    kc = cl.cfg["key"]
    issues.setdefault(kc, []).append(f"{int(ctx.raw[kc].duplicated().sum() - ctx.raw[kc].isna().sum() + 1)} repeated values")
    prof = cl.profile_df.copy()
    headers = ["Column", "Raw col", "Observed type", "Non blank (live)", "Blank (live)", "Blank %", "Distinct raw",
               "Distinct normalised", "Variant inflation", "Raw minimum", "Raw maximum", "Issues detected"]
    header_row(ws, r, 2, headers)
    top = r + 1
    for i, p in enumerate(prof.itertuples(), start=0):
        rr = top + i
        z = i % 2 == 1
        vals = [p.Column, ctx.rawcol[p.Column], p.ObservedType,
                f"=COUNTA({ctx.rr(p.Column)})", f"=ROWS({ctx.rr(p.Column)})-E{rr}", f"=F{rr}/ROWS({ctx.rr(p.Column)})",
                p.Distinct, p.DistinctAfterNormalising, f"=IFERROR(H{rr}/I{rr},1)",
                p.Min, p.Max, "; ".join(issues.get(p.Column, [])) or "No issue detected"]
        fm = [None, None, None, "#,##0", "#,##0", "0.0%", "#,##0", "#,##0", "0.00x", "#,##0.00", "#,##0.00", None]
        for j, v in enumerate(vals):
            c = ws.cell(rr, 2 + j, v)
            body_cell(c, fm[j], z, align="wrap" if j == 11 else None)
    end = top + len(prof) - 1
    ws.conditional_formatting.add(f"G{top}:G{end}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.1, color="E29578"))
    ws.conditional_formatting.add(f"J{top}:J{end}", CellIsRule(operator="greaterThan", formula=["1.01"], fill=fill(AMBER_L)))
    r = end + 2
    r = section(ws, r, 2, "Profile headline", 12)
    ws.cell(r, 2, "Raw data rows").font = F(9)
    ws.cell(r, 4, f"=ROWS({ctx.rr(kc)})").number_format = "#,##0"
    ws.cell(r + 1, 2, "Columns").font = F(9)
    ws.cell(r + 1, 4, len(prof))
    ws.cell(r + 2, 2, "Total blank cells").font = F(9)
    ws.cell(r + 2, 4, f"=SUM(F{top}:F{end})").number_format = "#,##0"
    ws.cell(r + 3, 2, "Overall completeness").font = F(9)
    ws.cell(r + 3, 4, f"=1-D{r + 2}/(D{r}*D{r + 1})").number_format = "0.00%"
    ws.cell(r + 4, 2, "Columns with formatting variants").font = F(9)
    ws.cell(r + 4, 4, f'=COUNTIF(J{top}:J{end},">1.01")')
    for k in range(r, r + 5):
        ws.cell(k, 4).font = F(10, True, NAVY)
    ws.freeze_panes = ws.cell(top, 3)


def C2(ctx):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C2 Structural Integrity")
    setup(ws, "C2  Structural Integrity: Blank Rows and Exact Duplicates",
          "Step 2 of the cleaning procedure. Remove rows that carry no information or repeat an earlier row exactly.", cols=12)
    for k, w in zip("BCDEFGHIJKLM", [14, 16, 20, 4, 14, 16, 22, 4, 14, 14, 14, 14]):
        ws.column_dimensions[k].width = w
    r = proc_header(ws, 5,
                    "Guarantee that each row represents one real event, so counts, rates and averages are not inflated.",
                    "(1) A row is blank when all 45 source fields are empty; blank rows are deleted. (2) A row is an exact duplicate when every field matches an earlier row character for character; "
                    "the first occurrence is kept and later copies are deleted. Each removed row is listed below with its original Excel row number in Raw Data so the decision can be audited.",
                    "Power Query: Home > Remove Rows > Remove Blank Rows, then Home > Remove Rows > Remove Duplicates across all columns. Worksheet check: COUNTIFS across all columns, or Data > Remove Duplicates.")
    n_raw = ctx.n_raw
    nb, nd = len(cl.blank_rows), len(cl.exact_dups)
    kpi(ws, r, 2, "RAW ROWS", f"=ROWS({ctx.rr(cl.cfg['key'])})", "#,##0", "Rows in Raw Data", 2)
    kpi(ws, r, 5, "BLANK ROWS REMOVED", nb, "#,##0", "All fields empty", 2, RED)
    kpi(ws, r, 8, "EXACT DUPLICATES REMOVED", nd, "#,##0", "Identical to an earlier row", 2, AMBER)
    kpi(ws, r, 11, "ROWS CARRIED FORWARD", f"=B{r + 1}-E{r + 1}-H{r + 1}", "#,##0", "Input to step C3", 2, GREEN)
    r += 4
    ws.cell(r, 2, "Live check: blank primary keys in Raw Data").font = F(9)
    ws.cell(r, 7, f"=COUNTBLANK({ctx.rr(cl.cfg['key'])})").font = F(10, True, NAVY)
    ws.cell(r, 8, f'=IF(G{r}=E{r - 3},"PASS","FAIL")')
    PASS_RULES(ws, f"H{r}")
    r += 2
    top = r
    section(ws, r, 2, "Blank rows removed", 3)
    section(ws, r, 6, "Exact duplicate rows removed", 3)
    b = cl.blank_rows.rename(columns={"source_row": "Raw Data row"})
    write_table(ws, r + 1, 2, b.assign(Action="Deleted"))
    d = cl.exact_dups.rename(columns={"source_row": "Raw Data row", "duplicate_of_row": "Duplicate of row", cl.cfg["key"]: "Record key"})
    write_table(ws, r + 1, 6, d)


def C3(ctx):
    cl = ctx.cl
    k = cl.cfg["key"]
    ws = ctx.wb.create_sheet("C3 Key Collision Repair")
    setup(ws, f"C3  Primary Key Collision Repair ({k})",
          "Step 3 of the cleaning procedure. Make the primary key unique without deleting genuine records.", cols=12)
    for kk, w in zip("BCDEFGHIJKLM", [14, 18, 18, 36, 18, 4, 26, 16, 60, 4, 4, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    f"After exact duplicates were removed, some different records still shared the same {k}. Deleting them would lose real events, so the key is repaired instead.",
                    f"Every record also carries {cl.cfg['anchor']}, which profiling proved to be unique and whose numeric suffix matches the record sequence. Within each group of records sharing a {k}, "
                    f"the record whose key suffix matches its {cl.cfg['anchor']} suffix keeps the key. Every other record is re keyed as {cl.cfg['key_prefix']} plus its anchor suffix. "
                    "A key_repaired_flag column in Clean Data marks every re keyed record.",
                    f"Power Query: Group By {k} with Count Rows, filter Count > 1, then Add Column > Conditional Column rebuilding the key from {cl.cfg['anchor']}. Worksheet check: COUNTIF({k} range, key) > 1.")
    col = cl.collisions
    nfix = int((col.resolution.str.startswith("Re-keyed")).sum()) if len(col) else 0
    kpi(ws, r, 2, "KEY GROUPS IN COLLISION", int(col.original_key.nunique()) if len(col) else 0, "#,##0", "Keys used by more than one record", 2, AMBER)
    kpi(ws, r, 5, "RECORDS RE KEYED", nfix, "#,##0", "Retained, new unique key", 2, TEAL)
    reg_top = r + 4 + (len(cl.identity_flags) + 3 if cl.identity_flags else 0) + 1
    kpi(ws, r, 8, "REGISTER KEYS STILL DUPLICATED (LIVE)", f'=COUNTIF($G${reg_top + 1}:$G${reg_top + max(1, len(col))},">1")', "#,##0", "Must equal zero", 3, GREEN)
    r += 4
    if cl.identity_flags:
        r = section(ws, r, 2, "Related identifiers flagged but deliberately not changed", 9)
        r = write_table(ws, r, 2, pd.DataFrame(cl.identity_flags), wrap_cols=("Treatment",)) + 1
    r = section(ws, r, 2, "Collision resolution register", 9)
    reg = col.rename(columns={"source_row": "Raw Data row", "original_key": "Original key", "anchor_id": "Anchor identifier",
                              "resolution": "Resolution", "new_key": "Key in Clean Data"})
    reg["Occurrences in Clean Data (live)"] = [f"=COUNTIF({ctx.cr(k)},\"{x}\")" for x in reg["Key in Clean Data"]]
    assert r == reg_top, (r, reg_top)
    write_table(ws, r, 2, reg[["Raw Data row", "Original key", "Anchor identifier", "Resolution", "Key in Clean Data", "Occurrences in Clean Data (live)"]])
    ws.column_dimensions["G"].width = 16


def C4(ctx):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C4 Text Standardization")
    setup(ws, "C4  Categorical Text Standardization",
          "Step 4 of the cleaning procedure. Collapse whitespace, case, punctuation and separator variants into one governed label.", cols=12)
    for kk, w in zip("BCDEFGHIJKLM", [26, 36, 34, 10, 40, 4, 26, 34, 12, 12, 10, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    "Ensure that grouping, filtering and lookups treat 'ICU', ' icu ', 'ICU.' and 'I_C_U' style variants as the same category.",
                    "For each categorical column a normalisation key is built: trim, remove a trailing period, replace underscores with spaces, collapse repeated spaces and ignore case. "
                    "All raw values sharing a key are mapped to the most frequent clean spelling (no padding, no trailing period, no underscore). Identifier columns are trimmed and upper cased; ZIP codes are stored as five character text.",
                    "Power Query: Transform > Format > Trim, Clean, Capitalize Each Word, then Replace Values for '_' and '.', then Merge with this mapping table. Worksheet: =PROPER(TRIM(SUBSTITUTE(SUBSTITUTE(x,\"_\",\" \"),\".\",\"\"))) with an INDEX/MATCH to the governed list.")
    tm = cl.text_map
    nv = int((tm.IssueType != "Canonical form").sum())
    kpi(ws, r, 2, "COLUMNS STANDARDISED", len(cl.cfg["categorical"]), "#,##0", "Categorical fields", 1)
    kpi(ws, r, 3, "RAW VARIANTS FOUND", nv, "#,##0", "Non standard spellings", 1, AMBER)
    kpi(ws, r, 4, "RECORDS CORRECTED", int(tm.loc[tm.IssueType != "Canonical form", "Records"].sum()), "#,##0", "Cells rewritten", 2, TEAL)
    kpi(ws, r, 8, "ROWS DUPLICATED AFTER STANDARDISING", cl.normalised_dups, "#,##0", "Hidden duplicates exposed by cleaning", 3, GREEN)
    r += 4
    # issue type breakdown
    r0 = r
    section(ws, r, 8, "Variants by issue type", 4)
    it = tm[tm.IssueType != "Canonical form"].groupby("IssueType").agg(Variants=("RawValue", "count"), Records=("Records", "sum")).reset_index().sort_values("Records", ascending=False)
    write_table(ws, r + 1, 8, it.rename(columns={"IssueType": "Issue type"}), fmts={"Records": "#,##0"}, wrap_cols=("Issue type",))
    # reconciliation per standard value
    r2 = r + len(it) + 3
    section(ws, r2, 8, "Reconciliation: standard label counts (live)", 4)
    std = tm.groupby(["Column", "StandardValue"]).Records.sum().reset_index()
    header_row(ws, r2 + 1, 8, ["Column / label", "Sum of raw variants", "Count in Clean Data", "Check"])
    top_map = r + 1
    end_map = top_map + len(tm)
    for i, x in enumerate(std.itertuples()):
        rr = r2 + 2 + i
        z = i % 2 == 1
        body_cell(ws.cell(rr, 8, f"{x.Column}: {x.StandardValue}"), None, z)
        body_cell(ws.cell(rr, 9, f'=SUMIFS($E${top_map + 1}:$E${end_map},$B${top_map + 1}:$B${end_map},"{x.Column}",$D${top_map + 1}:$D${end_map},"{x.StandardValue}")'), "#,##0", z)
        crit = str(x.StandardValue).replace('"', '""')
        body_cell(ws.cell(rr, 10, f'=COUNTIF({ctx.cr(x.Column)},"{crit}")'), "#,##0", z)
        body_cell(ws.cell(rr, 11, f'=IF(J{rr}<=I{rr},"PASS","FAIL")'), None, z)
    PASS_RULES(ws, f"K{r2 + 2}:K{r2 + 1 + len(std)}")
    note(ws, r2 + 2 + len(std), 8, "Clean Data counts can be lower than the raw sum only where exact duplicates or blank rows were removed in step C2; a higher count would indicate a leak and returns FAIL.", span=4)
    section(ws, r, 2, "Full variant to standard label mapping", 5)
    m = tm.rename(columns={"RawValue": "Raw value (as stored)", "StandardValue": "Standard label", "IssueType": "Issue type"})
    m = m.sort_values(["Column", "Standard label", "Records"], ascending=[True, True, False])
    write_table(ws, r + 1, 2, m[["Column", "Raw value (as stored)", "Standard label", "Records", "Issue type"]], fmts={"Records": "#,##0"})
    ws.conditional_formatting.add(f"F{r + 2}:F{r + 1 + len(m)}", FormulaRule(formula=[f'F{r + 2}<>"Canonical form"'], font=F(9, False, AMBER)))


def C5(ctx):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C5 Boolean Normalization")
    setup(ws, "C5  Boolean Flag Normalization",
          "Step 5 of the cleaning procedure. Twelve different encodings of yes and no are converted to native TRUE and FALSE.", cols=12)
    for kk, w in zip("BCDEFGHIJKLM", [30, 12, 12, 12, 12, 12, 4, 14, 14, 12, 4, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    "Allow flags to be summed, averaged and used in COUNTIFS without every formula having to test twelve spellings.",
                    "Tokens are trimmed and compared case insensitively. TRUE tokens: true, yes, y, 1. FALSE tokens: false, no, n, 0. Any token outside both lists would be set to blank and reported; none were found. "
                    "Blank flags in the raw file belonged only to fully blank rows removed in C2.",
                    "Power Query: Replace Values using this dictionary, then Change Type > True/False. Worksheet: =IF(OR(LOWER(TRIM(x))={\"true\",\"yes\",\"y\",\"1\"}),TRUE,IF(OR(LOWER(TRIM(x))={\"false\",\"no\",\"n\",\"0\"}),FALSE,\"\")).")
    section(ws, r, 2, "Token dictionary", 3)
    tok = pd.DataFrame({"Raw token": ["true", "TRUE", "yes", "Yes", "Y", "1", "false", "FALSE", "no", "No", "N", "0"],
                        "Standard value": [True] * 6 + [False] * 6})
    write_table(ws, r + 1, 2, tok)
    section(ws, r, 8, "Why this matters", 3)
    note(ws, r + 1, 8, "Before normalisation a simple COUNTIF(flag,\"Yes\") would capture only about one sixth of the true positives, understating every rate in this workbook by roughly 83 percent.", span=3, height=80)
    r = r + len(tok) + 3
    r = section(ws, r, 2, "Per column result with live verification", 6)
    bm = cl.bool_map
    header_row(ws, r, 2, ["Flag column", "Raw distinct tokens", "TRUE (converted)", "FALSE (converted)", "TRUE rate", "TRUE in Clean Data (live)", ""])
    for i, c in enumerate(cl.cfg["bool"]):
        rr = r + 1 + i
        g = bm[bm.Column == c]
        t = int(g[g.Standard == True].Records.sum())
        f_ = int(g[g.Standard == False].Records.sum())
        z = i % 2 == 1
        vals = [c, len(g), t, f_, f"=D{rr}/(D{rr}+E{rr})", f"=COUNTIF({ctx.cr(c)},TRUE)"]
        fm = [None, "0", "#,##0", "#,##0", "0.0%", "#,##0"]
        for j, v in enumerate(vals):
            body_cell(ws.cell(rr, 2 + j, v), fm[j], z)
    end = r + len(cl.cfg["bool"])
    ws.conditional_formatting.add(f"F{r + 1}:F{end}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5EA8A0"))
    note(ws, end + 1, 2, "The live column can be slightly below the converted count because duplicate rows removed in C2 are excluded from Clean Data.", span=7)


def C6(ctx):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C6 Date Standardization")
    setup(ws, "C6  Date Standardization",
          "Step 6 of the cleaning procedure. Four text date formats are parsed into true Excel dates.", cols=12)
    for kk, w in zip("BCDEFGHIJKLM", [26, 44, 12, 12, 16, 16, 4, 4, 4, 4, 4, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    "Produce real date values so time series, week numbers, day of week and elapsed time calculations are possible.",
                    "Each value is matched against explicit patterns in priority order: YYYY-MM-DD, YYYY/MM/DD and DD-Mon-YYYY are unambiguous. For NN/NN/YYYY values: if the first number exceeds 12 it must be the day (DD/MM); "
                    "if the second exceeds 12 it must be the day (MM/DD). When both are 12 or below the value is genuinely ambiguous. The data dictionary documents US style examples (03/11/2024), so ambiguous values are resolved as MM/DD/YYYY "
                    "and marked with an _ambiguous_flag column so any analysis can be rerun excluding them.",
                    "Power Query: Add Column > Conditional Column on Text.Length and Text.PositionOf, then Date.FromText with Culture \"en-US\" or \"en-GB\" per branch. Worksheet: DATE(RIGHT(x,4),LEFT(x,2),MID(x,4,2)) style formulas per pattern.")
    ds = cl.date_summary.rename(columns={"DetectedPattern": "Detected pattern", "ParsedExample": "Parsed as", "Example": "Raw example"})
    ds = ds[["Column", "Detected pattern", "Records", "Parsed", "Raw example", "Parsed as"]]
    r = section(ws, r, 2, "Pattern inventory", 6)
    r = write_table(ws, r, 2, ds, fmts={"Records": "#,##0", "Parsed": "#,##0"}) + 1
    r = section(ws, r, 2, "Live verification against Clean Data", 6)
    header_row(ws, r, 2, ["Date column", "Parsed dates (live)", "Ambiguous resolved", "Earliest", "Latest", ""])
    for i, c in enumerate(cl.cfg["dates"]):
        rr = r + 1 + i
        z = i % 2 == 1
        vals = [c, f"=COUNT({ctx.cr(c)})", f"=COUNTIF({ctx.cr(c + '_ambiguous_flag')},TRUE)", f"=MIN({ctx.cr(c)})", f"=MAX({ctx.cr(c)})"]
        fm = [None, "#,##0", "#,##0", "dd mmm yyyy", "dd mmm yyyy"]
        for j, v in enumerate(vals):
            body_cell(ws.cell(rr, 2 + j, v), fm[j], z)


def C7(ctx):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C7 Numeric Validation")
    setup(ws, "C7  Numeric Validation and Outlier Treatment",
          "Step 7 of the cleaning procedure. Impossible and extreme values are blanked so they cannot distort averages.", cols=13)
    for kk, w in zip("BCDEFGHIJKLMN", [30, 9, 10, 10, 11, 10, 9, 9, 50, 12, 12, 10, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    "Remove sign flipped values, repeated sentinel values and magnitude errors without deleting the rest of the record.",
                    "Two tests per column. (1) Domain limits: physical or definitional bounds such as no negative durations and no percentages above their definitional maximum. "
                    "(2) Tukey outer fence: any value above Q3 + 3 x IQR is an extreme outlier; the fence is applied only where the variable has no hard physical maximum. "
                    "The applied maximum is the tighter of the two. Failing cells are set to blank (never the whole row) and are then imputed or left blank in C9 according to documented rules. "
                    "The outlier_fields column in Clean Data names every field that was blanked on each record.",
                    "Power Query: Add Column > Conditional Column (if [x] < min or [x] > max then null else [x]). Worksheet: =IF(OR(x<min,x>max),\"\",x) with QUARTILE.INC for the fence.")
    nr = cl.numeric_rules.copy()
    header_row(ws, r, 2, ["Column", "Unit", "Domain min", "Domain max", "Tukey fence (Q3+3xIQR)", "Applied max", "Below min", "Above max", "Rationale",
                          "Clean min (live)", "Clean max (live)", "Check"])
    for i, x in enumerate(nr.itertuples()):
        rr = r + 1 + i
        z = i % 2 == 1
        vals = [x.Column, x.Unit, x.DomainMin, x.DomainMax, x.TukeyFence, x.AppliedMax, x.BelowMin, x.AboveMax, x.Rationale,
                f"=MIN({ctx.cr(x.Column)})", f"=MAX({ctx.cr(x.Column)})",
                f'=IF(AND(OR(D{rr}="",K{rr}>=D{rr}),OR(G{rr}="",L{rr}<=G{rr}+0.000001)),"PASS","FAIL")']
        fm = [None, None, "#,##0.##", "#,##0.##", "#,##0.00", "#,##0.00", "#,##0", "#,##0", None, "#,##0.00", "#,##0.00", None]
        for j, v in enumerate(vals):
            body_cell(ws.cell(rr, 2 + j, v), fm[j], z, align="wrap" if j == 8 else None)
    end = r + len(nr)
    PASS_RULES(ws, f"M{r + 1}:M{end}")
    ws.conditional_formatting.add(f"H{r + 1}:I{end}", CellIsRule(operator="greaterThan", formula=["0"], font=F(9, True, RED)))
    r = end + 2
    r = section(ws, r, 2, "Examples of blanked values (first six per column)", 8)
    ex = cl.numeric_examples.rename(columns={"source_row": "Raw Data row", "RawValue": "Raw value"})
    write_table(ws, r, 2, ex[["Column", "Raw Data row", "Raw value", "Reason", "Action"]], fmts={"Raw value": "#,##0.00"})


def C8(ctx, context_text):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C8 Cross Field Consistency")
    setup(ws, "C8  Cross Field Consistency Rules",
          "Step 8 of the cleaning procedure. Tests whether related fields agree, and records which field is treated as authoritative.", cols=10)
    for kk, w in zip("BCDEFGHIJK", [30, 52, 12, 12, 10, 62, 4, 4, 4, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    "Detect records where two fields that should agree by definition contradict each other, and make an explicit, documented decision rather than a silent one.",
                    context_text,
                    "Worksheet: helper columns with the rule as a Boolean formula, then COUNTIF(helper,TRUE). Power Query: Add Column > Custom Column with the rule, then Group By.")
    cr = cl.cross.rename(columns={"RecordsTested": "Records tested", "RecordsFailing": "Records failing", "FailRate": "Fail rate", "Rule": "Rule tested"})
    write_table(ws, r, 2, cr[["Check", "Rule tested", "Records tested", "Records failing", "Fail rate", "Decision"]],
                fmts={"Records tested": "#,##0", "Records failing": "#,##0", "Fail rate": "0.0%"}, wrap_cols=("Rule tested", "Decision", "Check"))
    for i in range(len(cr)):
        ws.row_dimensions[r + 1 + i].height = 64


def C9(ctx):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C9 Missing Value Treatment")
    setup(ws, "C9  Missing Value Treatment",
          "Step 9 of the cleaning procedure. Every blank cell is either filled by a documented rule or deliberately retained.", cols=10)
    for kk, w in zip("BCDEFGHIJK", [32, 12, 56, 12, 12, 14, 10, 4, 4, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    "Preserve sample size while avoiding bias: impute only where a defensible estimate exists and leave blanks where filling would invent information.",
                    "Numeric operational measures are filled with the median of the same operational group (unit, line or corridor), because the median is robust to the skew seen in C1 and the group captures real structural differences. "
                    "Categorical gaps receive an explicit 'Unknown' or 'Not Recorded' label so they remain visible in every pivot. Dates, survey scores and identifiers are never imputed. "
                    "Imputed cells are listed per record in the imputed_fields column of Clean Data.",
                    "Power Query: Group By the segment with Median, Merge back, then Replace Values null with the merged median. Worksheet: =IF(x=\"\",AGGREGATE(17,6,range/(group=g),2),x) or MEDIAN(IF()) array.")
    md = cl.missing_df.rename(columns={"NullsBefore": "Blanks before", "Filled": "Cells filled", "NullsAfter": "Blanks after"})
    header_row(ws, r, 2, ["Column", "Blanks before", "Strategy", "Cells filled", "Blanks after", "Blanks in Clean Data (live)", "Check"])
    for i, x in enumerate(md.itertuples()):
        rr = r + 1 + i
        z = i % 2 == 1
        vals = [x.Column, x._2, x.Strategy, x._4, x._5, f"=COUNTBLANK({ctx.cr(x.Column)})", f'=IF(G{rr}=F{rr},"PASS","FAIL")']
        fm = [None, "#,##0", None, "#,##0", "#,##0", "#,##0", None]
        for j, v in enumerate(vals):
            body_cell(ws.cell(rr, 2 + j, v), fm[j], z, align="wrap" if j == 2 else None)
    end = r + len(md)
    PASS_RULES(ws, f"H{r + 1}:H{end}")
    rr = end + 2
    ws.cell(rr, 2, "Records with at least one imputed value (live)").font = F(9)
    c = ws.cell(rr, 5, f'=COUNTIF({ctx.cr("imputed_count")},">0")')
    c.font = F(10, True, NAVY)
    c.number_format = "#,##0"
    ws.cell(rr + 1, 2, "Share of records touched by imputation").font = F(9)
    c = ws.cell(rr + 1, 5, f"=E{rr}/ROWS({ctx.cr('imputed_count')})")
    c.font = F(10, True, NAVY)
    c.number_format = "0.0%"


def C10(ctx, features):
    ws = ctx.wb.create_sheet("C10 Feature Engineering")
    setup(ws, "C10  Feature Engineering",
          "Step 10 of the cleaning procedure. Analysis ready fields derived from the cleaned source fields.", cols=8)
    for kk, w in zip("BCDEFGHI", [30, 60, 50, 16, 4, 4, 4, 4]):
        ws.column_dimensions[kk].width = w
    r = proc_header(ws, 5,
                    "Create the time keys, segment keys and governed metrics that every model and dashboard sheet depends on, once, in one place.",
                    "Each feature is computed from cleaned fields only. Definitions below are the single source of truth; model sheets reference these columns through named ranges.",
                    "Power Query: Add Column > Custom Column. Worksheet equivalents are shown per feature.")
    f = pd.DataFrame(features, columns=["Feature", "Definition", "Worksheet equivalent"])
    f["Worksheet equivalent"] = f["Worksheet equivalent"].map(lambda x: "Formula:  " + x.lstrip("=") if str(x).startswith("=") else x)
    f["Populated (live)"] = [f"=COUNTA({ctx.cr(x)})" for x in f.Feature]
    write_table(ws, r, 2, f, fmts={"Populated (live)": "#,##0"}, wrap_cols=("Definition", "Worksheet equivalent"))
    for i in range(len(f)):
        ws.row_dimensions[r + 1 + i].height = 40


def C11(ctx):
    cl = ctx.cl
    ws = ctx.wb.create_sheet("C11 Cleaning Audit Log")
    setup(ws, "C11  Cleaning Audit Log and Row Reconciliation",
          "Final step. One line per transformation, with row and cell counts, reconciled live against Raw Data and Clean Data.", cols=9)
    for kk, w in zip("BCDEFGHIJ", [8, 28, 44, 11, 11, 13, 62, 4, 4]):
        ws.column_dimensions[kk].width = w
    a = pd.DataFrame(cl.audit)
    a = a.rename(columns={"RowsIn": "Rows in", "RowsOut": "Rows out", "CellsChanged": "Cells changed"})
    r = write_table(ws, 5, 2, a[["Step", "Sheet", "Action", "Rows in", "Rows out", "Cells changed", "Note"]],
                    fmts={"Rows in": "#,##0", "Rows out": "#,##0", "Cells changed": "#,##0"}, wrap_cols=("Note", "Action"))
    for i in range(len(a)):
        ws.row_dimensions[6 + i].height = 30
    r += 1
    r = section(ws, r, 2, "Row reconciliation (live formulas)", 6)
    k = cl.cfg["key"]
    lines = [("Rows in Raw Data", f"=ROWS({ctx.rr(k)})"),
             ("Less: blank rows (C2)", -len(cl.blank_rows)),
             ("Less: exact duplicates (C2)", -len(cl.exact_dups)),
             ("Expected rows in Clean Data", f"=SUM(F{r}:F{r + 2})"),
             ("Actual rows in Clean Data", f"=COUNTA({ctx.cr(k)})"),
             ("Difference", f"=F{r + 4}-F{r + 3}")]
    for i, (lab, v) in enumerate(lines):
        ws.cell(r + i, 3, lab).font = F(10, i >= 3)
        c = ws.cell(r + i, 6, v)
        c.number_format = "#,##0;(#,##0);0"
        c.font = F(10, i >= 3, NAVY)
        c.border = BORDER
    ws.cell(r + 5, 7, f'=IF(F{r + 5}=0,"PASS","FAIL")')
    PASS_RULES(ws, f"G{r + 5}")
    r += 7
    note(ws, r, 2, "Principle applied throughout: no row is deleted unless it is empty or an exact copy; no value is changed without a flag column that allows the change to be traced, reversed or excluded from analysis.", span=7, color=TEAL, italic=True)
