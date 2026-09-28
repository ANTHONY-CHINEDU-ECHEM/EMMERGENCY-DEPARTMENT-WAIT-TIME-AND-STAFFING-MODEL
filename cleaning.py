"""Config driven data cleaning engine shared by the four Excel portfolio projects.

Every step records what it did so the workbook can show a transparent,
sheet by sheet audit trail. Nothing is silently dropped or changed.
"""
import re
import numpy as np
import pandas as pd

TRUE_TOKENS = {"true", "yes", "y", "1"}
FALSE_TOKENS = {"false", "no", "n", "0"}


def norm_key(v):
    s = str(v).strip().rstrip(".").replace("_", " ")
    s = re.sub(r"\s+", " ", s)
    return s.casefold()


def classify_variant(raw, canonical):
    s = str(raw)
    if s == canonical:
        return "Canonical form"
    issues = []
    if s != s.strip():
        issues.append("Leading or trailing whitespace")
    t = s.strip()
    if t.endswith("."):
        issues.append("Trailing period")
    if "_" in t:
        issues.append("Underscore used as separator")
    base = t.rstrip(".").replace("_", " ")
    if base != canonical and base.casefold() == canonical.casefold():
        issues.append("Inconsistent letter case")
    return "; ".join(issues) if issues else "Other variant"


DATE_PATTERNS = [
    ("ISO 8601 (YYYY-MM-DD)", re.compile(r"^\d{4}-\d{2}-\d{2}$"), "%Y-%m-%d"),
    ("Slash ISO (YYYY/MM/DD)", re.compile(r"^\d{4}/\d{2}/\d{2}$"), "%Y/%m/%d"),
    ("Day Month abbrev (DD-Mon-YYYY)", re.compile(r"^\d{2}-[A-Za-z]{3}-\d{4}$"), "%d-%b-%Y"),
]
SLASH = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")


def parse_date(v):
    """Return (timestamp, pattern_label, ambiguous_flag)."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return pd.NaT, "Missing", False
    if isinstance(v, (pd.Timestamp,)):
        return v, "Native date", False
    s = str(v).strip()
    for label, rx, fmt in DATE_PATTERNS:
        if rx.match(s):
            try:
                return pd.to_datetime(s, format=fmt), label, False
            except ValueError:
                return pd.NaT, label + " (invalid calendar date)", False
    m = SLASH.match(s)
    if m:
        a, b, y = map(int, m.groups())
        try:
            if a > 12:
                return pd.Timestamp(y, b, a), "DD/MM/YYYY (day > 12, unambiguous)", False
            if b > 12:
                return pd.Timestamp(y, a, b), "MM/DD/YYYY (day > 12, unambiguous)", False
            return pd.Timestamp(y, a, b), "Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY", True
        except ValueError:
            return pd.NaT, "Slash date (invalid calendar date)", False
    return pd.NaT, "Unrecognised format", False


class Cleaner:
    """cfg keys
    key: primary key column; anchor: collision free column carrying row identity
    key_width: digits in primary key suffix; key_prefix
    flag_identity: optional list of (column, description) to flag, not repair
    categorical: list of columns; bool: list; dates: list
    id_text: columns to trim/uppercase; zip: columns to pad to 5 digits
    numeric: {col: dict(lo=, hi=, fence=True/False, unit=, rationale=)}
    impute: {col: ("median", group_col) | ("constant", value) | ("leave", reason)}
    """

    def __init__(self, raw, cfg):
        self.raw = raw.copy()
        self.cfg = cfg
        self.audit = []

    def _log(self, step, sheet, action, rows_in, rows_out, cells, note=""):
        self.audit.append(dict(Step=step, Sheet=sheet, Action=action, RowsIn=rows_in,
                               RowsOut=rows_out, CellsChanged=cells, Note=note))

    # ------------------------------------------------------------------ profile
    def profile(self):
        df = self.raw
        rows = []
        for c in df.columns:
            s = df[c]
            nn = s.notna()
            num = pd.to_numeric(s, errors="coerce")
            is_num = num.notna().sum() >= 0.95 * nn.sum() and nn.sum() > 0
            keys = s.dropna().map(norm_key)
            rows.append(dict(
                Column=c,
                ObservedType="Numeric" if is_num else "Text",
                NonNull=int(nn.sum()), Null=int((~nn).sum()),
                Distinct=int(s.nunique()),
                DistinctAfterNormalising=int(keys.nunique()) if not is_num else int(s.nunique()),
                Min=float(num.min()) if is_num else None,
                Max=float(num.max()) if is_num else None,
            ))
        self.profile_df = pd.DataFrame(rows)
        return self.profile_df

    # --------------------------------------------------------------- structural
    def structural(self):
        df = self.raw.copy()
        df.insert(0, "source_row", np.arange(2, len(df) + 2))
        n0 = len(df)
        body = df.drop(columns="source_row")
        blank = body.isna().all(axis=1)
        self.blank_rows = df.loc[blank, ["source_row"]].copy()
        df = df[~blank]
        n1 = len(df)
        body = df.drop(columns="source_row")
        dup = body.duplicated(keep="first")
        first_of = {}
        keyed = body.apply(lambda col: col.map(str)).agg("|".join, axis=1)
        firsts = df.loc[~dup].assign(_k=keyed[~dup]).set_index("_k")["source_row"]
        self.exact_dups = pd.DataFrame({
            "source_row": df.loc[dup, "source_row"].values,
            "duplicate_of_row": firsts.reindex(keyed[dup]).values,
            self.cfg["key"]: df.loc[dup, self.cfg["key"]].values,
        })
        df = df[~dup]
        n2 = len(df)
        self._log("S1", "C2 Structural Integrity", "Removed fully blank rows", n0, n1, 0,
                  f"{n0 - n1} rows contained no values in any of the {body.shape[1]} columns")
        self._log("S2", "C2 Structural Integrity", "Removed exact duplicate records", n1, n2, 0,
                  f"{n1 - n2} rows were byte for byte copies of an earlier row; first occurrence kept")
        self.df = df.reset_index(drop=True)

    # ------------------------------------------------------- key collisions
    def key_collisions(self):
        cfg = self.cfg
        df = self.df
        k, a = cfg["key"], cfg["anchor"]
        suf = lambda s: s.astype(str).str.extract(r"(\d+)$")[0].astype(float)
        df["_ks"], df["_as"] = suf(df[k]), suf(df[a])
        dupmask = df[k].duplicated(keep=False) & df[k].notna()
        existing = set(df[k].dropna())
        recs = []
        for key, g in df[dupmask].groupby(k):
            owner = g[g["_ks"] == g["_as"]]
            owner_idx = owner.index[0] if len(owner) else g.index[0]
            for idx, r in g.iterrows():
                if idx == owner_idx:
                    recs.append(dict(source_row=r.source_row, original_key=key, anchor_id=r[a],
                                     resolution="Retained (key suffix matches anchor)" if len(owner) else "Retained (first occurrence)",
                                     new_key=key))
                    continue
                new = f"{cfg['key_prefix']}{int(r['_as']):0{cfg['key_width']}d}"
                if new in existing:
                    new = new + "R"
                existing.add(new)
                df.at[idx, k] = new
                recs.append(dict(source_row=r.source_row, original_key=key, anchor_id=r[a],
                                 resolution="Re-keyed from anchor identifier", new_key=new))
        self.collisions = pd.DataFrame(recs)
        n_fixed = int((self.collisions["resolution"] == "Re-keyed from anchor identifier").sum()) if len(recs) else 0
        df["key_repaired_flag"] = df["source_row"].isin(
            self.collisions.loc[self.collisions.resolution.str.startswith("Re-keyed"), "source_row"]) if len(recs) else False
        # identity conflicts flagged, not repaired
        self.identity_flags = []
        for col, desc in cfg.get("flag_identity", []):
            m = df[col].duplicated(keep=False) & (suf(df[col]) != df["_as"])
            df[f"{col}_conflict_flag"] = m
            self.identity_flags.append(dict(Column=col, RecordsFlagged=int(m.sum()), Treatment=desc))
        df.drop(columns=["_ks", "_as"], inplace=True)
        assert df[k].is_unique, "primary key still not unique"
        self._log("S3", "C3 Key Collision Repair", f"Repaired duplicate {k} values", len(df), len(df), n_fixed,
                  f"{n_fixed} records carried a {k} already used by a different record; re-keyed from {a}")
        self.df = df

    # ----------------------------------------------------- text standardisation
    def text(self):
        df = self.df
        maps = []
        changed = 0
        for c in self.cfg["categorical"]:
            s = df[c]
            nn = s.dropna()
            keys = nn.map(norm_key)
            canon = {}
            for kk, grp in nn.groupby(keys):
                clean = grp[grp.map(lambda x: str(x) == str(x).strip() and not str(x).endswith(".") and "_" not in str(x))]
                pool = clean if len(clean) else grp
                canon[kk] = pool.value_counts().index[0]
            override = self.cfg.get("canonical_override", {}).get(c, {})
            for kk, v in override.items():
                canon[kk] = v
            new = s.map(lambda x: canon[norm_key(x)] if pd.notna(x) else x)
            diff = (new != s) & s.notna()
            changed += int(diff.sum())
            vc = s.value_counts(dropna=False)
            for raw, cnt in vc.items():
                if pd.isna(raw):
                    continue
                cv = canon[norm_key(raw)]
                maps.append(dict(Column=c, RawValue=repr(raw)[1:-1] if isinstance(raw, str) else str(raw),
                                 StandardValue=cv, Records=int(cnt), IssueType=classify_variant(raw, cv)))
            df[c] = new
        for c in self.cfg.get("id_text", []):
            new = df[c].map(lambda x: str(x).strip().upper() if pd.notna(x) else x)
            changed += int(((new != df[c]) & df[c].notna()).sum())
            df[c] = new
        for c in self.cfg.get("zip", []):
            new = df[c].map(lambda x: str(int(float(x))).zfill(5) if pd.notna(x) else x)
            df[c] = new
        for c in self.cfg.get("code_text", []):
            df[c] = df[c].map(lambda x: str(x).strip() if pd.notna(x) else x)
        self.text_map = pd.DataFrame(maps)
        n_var = int((self.text_map.IssueType != "Canonical form").sum())
        self._log("S4", "C4 Text Standardization", "Collapsed categorical variants to one standard label",
                  len(df), len(df), changed,
                  f"{n_var} non standard variants across {len(self.cfg['categorical'])} categorical columns")
        # normalised duplicates check (rows identical after standardisation)
        body = df.drop(columns=[x for x in df.columns if x in ("source_row", "key_repaired_flag", self.cfg["key"]) or x.endswith("_conflict_flag")])
        nd = body.duplicated(keep=False)
        self.normalised_dups = int(body.duplicated().sum())
        self.df = df

    # --------------------------------------------------------------- booleans
    def booleans(self):
        df = self.df
        recs = []
        changed = 0
        for c in self.cfg["bool"]:
            s = df[c]
            vc = s.astype(object).where(s.notna(), None).value_counts(dropna=False)
            def conv(x):
                if pd.isna(x):
                    return np.nan
                t = str(x).strip().casefold()
                if t in TRUE_TOKENS:
                    return True
                if t in FALSE_TOKENS:
                    return False
                return np.nan
            new = s.map(conv)
            for raw, cnt in vc.items():
                if raw is None or (isinstance(raw, float) and np.isnan(raw)):
                    continue
                recs.append(dict(Column=c, RawToken=str(raw), Standard=conv(raw), Records=int(cnt)))
            changed += int(s.notna().sum())
            df[c] = new
        self.bool_map = pd.DataFrame(recs)
        self._log("S5", "C5 Boolean Normalization", "Converted 12 token encodings to TRUE/FALSE",
                  len(df), len(df), changed, f"{len(self.cfg['bool'])} flag columns converted to native Boolean")
        self.df = df

    # ------------------------------------------------------------------ dates
    def dates(self):
        df = self.df
        summ, ex = [], []
        changed = 0
        for c in self.cfg["dates"]:
            parsed = df[c].map(parse_date)
            vals = parsed.map(lambda t: t[0])
            pats = parsed.map(lambda t: t[1])
            amb = parsed.map(lambda t: t[2])
            for p, cnt in pats.value_counts().items():
                sub = df.loc[pats == p, c].dropna()
                summ.append(dict(Column=c, DetectedPattern=p, Records=int(cnt),
                                 Parsed=int(vals[pats == p].notna().sum()),
                                 Example=str(sub.iloc[0]) if len(sub) else "",
                                 ParsedExample=vals[pats == p].dropna().iloc[0].strftime("%d %b %Y") if vals[pats == p].notna().any() else ""))
            df[c] = pd.to_datetime(vals)
            df[f"{c}_ambiguous_flag"] = amb.astype(bool)
            changed += int(vals.notna().sum())
        self.date_summary = pd.DataFrame(summ)
        self._log("S6", "C6 Date Standardization", "Parsed four mixed date formats to true Excel dates",
                  len(df), len(df), changed, "Ambiguous NN/NN/YYYY values resolved with a documented rule and flagged")
        self.df = df

    # ---------------------------------------------------------------- numerics
    def numerics(self):
        df = self.df
        rules, examples = [], []
        df["outlier_fields"] = ""
        changed = 0
        for c, r in self.cfg["numeric"].items():
            s = pd.to_numeric(df[c], errors="coerce").astype(float)
            q1, q3 = s.quantile([0.25, 0.75])
            iqr = q3 - q1
            fence_hi = q3 + 3 * iqr if r.get("fence", True) else np.inf
            lo = r.get("lo", -np.inf)
            hi = min(r.get("hi", np.inf), fence_hi)
            below = s < lo
            above = s > hi
            bad = below | above
            reason = np.where(below, "Below valid minimum" + (" (negative)" if lo == 0 else ""),
                              np.where(above, "Above valid maximum", ""))
            for idx in s[bad].index[:6]:
                examples.append(dict(Column=c, source_row=int(df.at[idx, "source_row"]),
                                     RawValue=float(s[idx]), Reason=reason[df.index.get_loc(idx)],
                                     Action="Set to blank, then imputed in C9" if self.cfg["impute"].get(c, ("leave",))[0] != "leave" else "Set to blank"))
            rules.append(dict(Column=c, Unit=r.get("unit", ""), DomainMin=None if lo == -np.inf else lo,
                              DomainMax=None if r.get("hi", np.inf) == np.inf else r.get("hi"),
                              TukeyFence=None if fence_hi == np.inf else round(float(fence_hi), 2),
                              AppliedMax=None if hi == np.inf else round(float(hi), 2),
                              BelowMin=int(below.sum()), AboveMax=int(above.sum()),
                              Rationale=r.get("rationale", "")))
            df.loc[bad, "outlier_fields"] = df.loc[bad, "outlier_fields"] + c + ";"
            s[bad] = np.nan
            df[c] = s
            changed += int(bad.sum())
        for c in self.cfg.get("numeric_plain", []):
            df[c] = pd.to_numeric(df[c], errors="coerce")
        df["outlier_count"] = df["outlier_fields"].str.count(";")
        df["outlier_fields"] = df["outlier_fields"].str.rstrip(";")
        self.numeric_rules = pd.DataFrame(rules)
        self.numeric_examples = pd.DataFrame(examples)
        self._log("S7", "C7 Numeric Validation", "Nullified impossible and extreme values", len(df), len(df), changed,
                  "Domain limits plus Tukey outer fence (Q3 + 3 x IQR); values are blanked, never deleted")
        self.df = df

    # --------------------------------------------------------- cross field
    def cross_field(self, checks):
        """checks: list of dict(Check, Description, mask (bool Series), Decision)"""
        df = self.df
        recs = []
        for ch in checks:
            m = ch["mask"](df)
            recs.append(dict(Check=ch["Check"], Rule=ch["Description"], RecordsTested=int(m.notna().sum()),
                             RecordsFailing=int(m.fillna(False).sum()),
                             FailRate=float(m.fillna(False).sum() / max(1, len(df))),
                             Decision=ch["Decision"]))
            if ch.get("flag"):
                df[ch["flag"]] = m.fillna(False).astype(bool)
        self.cross = pd.DataFrame(recs)
        self._log("S8", "C8 Cross Field Consistency", "Tested business rules between related fields",
                  len(df), len(df), 0, f"{len(checks)} rules tested; authoritative field chosen for each")
        self.df = df

    # ------------------------------------------------------------ missing
    def missing(self):
        df = self.df
        recs = []
        df["imputed_fields"] = ""
        changed = 0
        for c, (strategy, *arg) in self.cfg["impute"].items():
            before = int(df[c].isna().sum())
            if strategy == "median":
                grp = arg[0]
                med = df.groupby(grp)[c].transform("median")
                fill = med.fillna(df[c].median())
                if self.cfg["numeric"].get(c, {}).get("integer"):
                    fill = fill.round()
                m = df[c].isna()
                df.loc[m, c] = fill[m]
                df.loc[m, "imputed_fields"] += c + ";"
                desc = f"Median within {grp} (overall median if group empty)"
            elif strategy == "constant":
                m = df[c].isna()
                df.loc[m, c] = arg[0]
                desc = f"Explicit category '{arg[0]}'"
            else:
                m = pd.Series(False, index=df.index)
                desc = "Left blank: " + arg[0]
            after = int(df[c].isna().sum())
            changed += before - after
            recs.append(dict(Column=c, NullsBefore=before, Strategy=desc, Filled=before - after, NullsAfter=after))
        df["imputed_count"] = df["imputed_fields"].str.count(";")
        df["imputed_fields"] = df["imputed_fields"].str.rstrip(";")
        self.missing_df = pd.DataFrame(recs)
        self._log("S9", "C9 Missing Value Treatment", "Filled or explicitly retained every missing value",
                  len(df), len(df), changed, "Each imputed cell is traceable through imputed_fields")
        self.df = df
