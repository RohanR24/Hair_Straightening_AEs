"""
Supporting Information pipeline (single file)

"Comparative Adverse Event Reporting Profiles of Relaxer/Alkali and Keratin/Blowout
Hair-Straightening Products: An FDA Disproportionality Analysis"
Journal of Cosmetic Dermatology, major revision.

What this script does, in order:
  1. Reads the report-level analysis workbook (Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx).
  2. Runs every analysis requested by the reviewer:
       - full results for all tested reactions (primary and SA3 co-primary), with Woolf and
         exact conditional 95% CIs and explicit handling of zero-cell (boundary) estimates
       - effect estimates across SA1-SA4
       - possible-duplicate screen and de-duplicated analysis
       - temporal restrictions (excluding 2010-2012, 2013-2023 only, 2004-2023 only) and female-only
       - calendar-period-stratified and age-stratified Mantel-Haenszel estimates, and logistic
         regression adjusted for age and sex
       - exposure-classification detail (criteria, SA3 exclusion reasons, report-level list)
       - brand concentration and leave-top-brand-out analyses
       - reaction inclusion-threshold sensitivity analysis
       - co-occurrence of signals and symptom-cluster analyses (incl. temporal, brand, period-stratified)
       - descriptive listing of serious-outcome reports
       - annual and period-level denominators with Wilson 95% CIs
  3. Writes the supplementary tables to Excel (journal file + author-only working file).
  4. Draws the manuscript figures (600-dpi PNG, PDF, and EPS):
       Figure 1  study flow, with the possible-duplicate screen and SA3 exclusions by reason
       Figure 2  forest plot of the primary signals (no point estimate for zero-cell reactions)
       Figure 3  robustness heatmap with the q value (or P value) printed in each cell
       Figure 4  annual report counts and period-level proportions with Wilson 95% CIs
  5. Draws the supplementary figures S1-S3 (PNG and PDF).
  6. Builds the Word Supporting Information document (tables S1-S15 and figures S1-S3).
     Writes manuscript Table 2 (primary and SA3 columns) as a ready-to-paste Word table.
  7. Runs the quality-control checks and writes QC_report.md.

HOW TO RUN
  1. Edit WORKBOOK_PATH below so it points at your spreadsheet.
  2. (Optional) edit OUTPUT_DIR. If left blank, a folder called "Supplement_outputs" is
     created next to the spreadsheet.
  3. Run:   python hair_straightening_supplement.py

Requirements: Python 3.9+, pandas, numpy, scipy, statsmodels, openpyxl, matplotlib, python-docx
  pip install pandas numpy scipy statsmodels openpyxl matplotlib python-docx
"""

# ======================================================================================
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  USER SETTINGS  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# ======================================================================================

# Full path to the report-level workbook. Keep the r in front of the quotes (it matters on Windows).
#   Mac example:     r"/Users/yourname/Documents/Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx"
#   Windows example: r"C:\Users\yourname\Documents\Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx"
WORKBOOK_PATH = r"/PATH/TO/Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx"

# (If this path does not exist but the script sits in the repository's analysis/ folder, the copy in
#  ../data/ is used automatically.)

# Folder for all outputs. Leave as "" to use a "Supplement_outputs" folder next to the workbook.
OUTPUT_DIR = r""

MAKE_MAIN_FIGURES = True   # manuscript Figures 1-4 (600-dpi PNG + PDF + EPS)
MAKE_FIGURES = True        # supplementary figures S1-S3 (PNG + PDF)
MAKE_WORD = True           # Supporting_Information.docx
MAKE_TABLE2 = True         # manuscript Table 2 as a ready-to-paste Word table (Table2.docx)
RUN_QC = True              # QC_report.md (agreement and plausibility checks)

# ======================================================================================

import math
import os
import re
import sys
import warnings
from collections import Counter

import numpy as np
import pandas as pd
from scipy.stats import fisher_exact, hypergeom
from scipy.stats.contingency import odds_ratio
from statsmodels.stats.multitest import multipletests
from statsmodels.stats.contingency_tables import StratifiedTable
from statsmodels.stats.proportion import proportion_confint
import statsmodels.api as sm

warnings.filterwarnings("ignore")

RELAXER, KERATIN = "Relaxer/alkali", "Keratin/blowout"
MIN_REPORTS = 5          # primary inclusion rule: >=5 reports in either cohort
PEAK_YEARS = {2010, 2011, 2012}
MOCRA_DATE = pd.Timestamp("2023-12-29")
MANUSCRIPT_TITLE = ("Comparative Adverse Event Reporting Profiles of Relaxer/Alkali and Keratin/Blowout "
                    "Hair-Straightening Products: An FDA Disproportionality Analysis")
COLOR = {RELAXER: "#B23A48", KERATIN: "#2C6E7F"}     # cohort identity in all figures (as in the submitted Figure 4)
INK, MUTED, GRID = "#1f1f1f", "#6b6b6b", "#d9d9d9"


# ======================================================================================
# 1. LOADING
# ======================================================================================

def load(path):
    xl = pd.ExcelFile(path)
    qc = xl.parse("Report_Level_QC")
    rc = xl.parse("Report_Characteristics")
    raw = xl.parse("Temporal_Trend", header=None)
    hi = hj = None
    for i in range(raw.shape[0]):
        for j in range(raw.shape[1]):
            if str(raw.iat[i, j]).strip() == "Report number":
                hi, hj = i, j
                break
        if hi is not None:
            break
    rl = raw.iloc[hi + 1:, hj:hj + 10].copy()
    rl.columns = list(raw.iloc[hi, hj:hj + 10])
    rl = rl.dropna(how="all")

    for t in (qc, rc, rl):
        t["rn"] = t["Report number"].astype(str).str.strip()
    rl["date"] = pd.to_datetime(rl["Date created"])
    df = (qc.merge(rc.drop(columns=["Arm", "Report number"]), on="rn", how="left")
            .merge(rl[["rn", "date"]], on="rn", how="left"))
    df["cohort"] = np.where(df["Arm"].str.startswith("Alkali"), RELAXER, KERATIN)
    df["relaxer"] = (df["cohort"] == RELAXER).astype(int)
    df["year"] = df["date"].dt.year
    # one entry per distinct (case-normalized) reaction term per report
    df["rx"] = df["Reactions"].fillna("").apply(
        lambda s: frozenset(t.strip().upper() for t in str(s).split(";") if t.strip()))
    df["female"] = (df["Gender"] == "Female").astype(int)
    df["age"] = pd.to_numeric(df["Age"], errors="coerce").where(df["Age available"] == "Yes")
    excl = xl.parse("Strict_Excluded_Reports")
    excl["rn"] = excl["Report number"].astype(str).str.strip()
    df = df.merge(excl[["rn", "Reason for conservative exclusion"]], on="rn", how="left")
    return df


# ======================================================================================
# 2. CORE 2x2 STATISTICS
# ======================================================================================

def two_by_two(d, term):
    A = d[d.relaxer == 1]
    K = d[d.relaxer == 0]
    a = int(A["rx"].apply(lambda s: term in s).sum())
    c = int(K["rx"].apply(lambda s: term in s).sum())
    return a, len(A) - a, c, len(K) - c


def ror_stats(a, b, c, d):
    """ROR (relaxer/alkali vs keratin/blowout) with Woolf CI when no cell is zero and the
    exact conditional CI otherwise; the exact conditional CI is also returned for all."""
    tbl = [[a, b], [c, d]]
    p = fisher_exact(tbl).pvalue
    ex = odds_ratio(tbl, kind="conditional").confidence_interval(0.95)
    ex_lo, ex_hi = float(ex.low), float(ex.high)
    if min(a, b, c, d) > 0:
        ror = a * d / (b * c)
        se = math.sqrt(1 / a + 1 / b + 1 / c + 1 / d)
        lo, hi = math.exp(math.log(ror) - 1.96 * se), math.exp(math.log(ror) + 1.96 * se)
        method, boundary = "Woolf", ""
    else:
        lo, hi, method = ex_lo, ex_hi, "Exact conditional"
        if a == 0:
            ror, boundary = 0.0, "lower (0 relaxer/alkali reports)"
        elif c == 0:
            ror, boundary = math.inf, "upper (0 keratin/blowout reports)"
        else:
            ror, boundary = np.nan, "other zero cell"
    return dict(ror=ror, lo=lo, hi=hi, ci_method=method, boundary=boundary,
                exact_lo=ex_lo, exact_hi=ex_hi, p=p)


def analyse(d, terms, label):
    rows = []
    nA, nK = int(d.relaxer.sum()), int((d.relaxer == 0).sum())
    for t in terms:
        a, b, c, dd = two_by_two(d, t)
        s = ror_stats(a, b, c, dd)
        rows.append(dict(analysis=label, reaction=t, relaxer_n=a, relaxer_N=nA,
                         keratin_n=c, keratin_N=nK, **s))
    out = pd.DataFrame(rows)
    out["q"] = multipletests(out["p"], method="fdr_bh")[1]
    out["fdr_sig"] = out["q"] < 0.05
    out["direction"] = np.where(out["relaxer_n"] / nA > out["keratin_n"] / nK,
                                "Higher in relaxer/alkali", "Higher in keratin/blowout")
    return out


def tested_terms(d, min_reports=MIN_REPORTS, mode="either"):
    A = Counter(t for s in d[d.relaxer == 1]["rx"] for t in s)
    K = Counter(t for s in d[d.relaxer == 0]["rx"] for t in s)
    terms = set(A) | set(K)
    if mode == "either":
        keep = [t for t in terms if max(A[t], K[t]) >= min_reports]
    else:
        keep = [t for t in terms if A[t] + K[t] >= min_reports]
    return sorted(keep)


def mh_stratified(d, strat, t):
    """Mantel-Haenszel OR across strata with CMH test. Returns (or, lo, hi, p); or/lo/hi are NaN
    when not estimable (no events in one cohort across all strata)."""
    tabs = []
    for _, g in d.groupby(strat, observed=True):
        a, b, c, dd = two_by_two(g, t)
        if (a + b) > 0 and (c + dd) > 0:
            tabs.append(np.array([[a, b], [c, dd]], dtype=float))
    st = StratifiedTable(tabs)
    try:
        p = st.test_null_odds(correction=True).pvalue
    except Exception:
        p = np.nan
    num = sum(t_[0, 0] * t_[1, 1] / t_.sum() for t_ in tabs)
    den = sum(t_[0, 1] * t_[1, 0] / t_.sum() for t_ in tabs)
    if num == 0 or den == 0:
        return np.nan, np.nan, np.nan, p
    lo, hi = st.oddsratio_pooled_confint()
    return st.oddsratio_pooled, lo, hi, p


def flag_terms(d, terms):
    """Collapse a set of terms into one pseudo-term 'X' (report counted once if any term present)."""
    ts = set(terms)
    return d.assign(rx=d.rx.map(lambda s: frozenset(["X"]) if s & ts else frozenset()))


# ======================================================================================
# 3. FORMATTING HELPERS
# ======================================================================================

def disp(term):
    s = term.replace("DYSPNOEA", "DYSPNEA").replace("RHINORRHOEA", "RHINORRHEA") \
            .replace("OEDEMA", "EDEMA").replace("HAEM", "HEM").replace("DIARRHOEA", "DIARRHEA") \
            .replace("AESTHESIA", "ESTHESIA").replace("COLOUR", "COLOR").replace("LEUKAEMIA", "LEUKEMIA") \
            .replace("HYPERAEMIA", "HYPEREMIA").replace("OESOPH", "ESOPH")
    return s.capitalize()


def fmt_p(p):
    if pd.isna(p):
        return "NA"
    if p < 0.001:
        return "<.001"
    return (f"{p:.3f}" if p < 0.1 else f"{p:.2f}").lstrip("0")


def lab_p(name, x):
    """'q=.025' but 'q<.001' (no '=<')."""
    v = fmt_p(x)
    return f"{name}{v}" if v.startswith("<") else f"{name}={v}"


def fmt_ror(r):
    if r["boundary"].startswith("lower"):
        return "NE*"
    if r["boundary"].startswith("upper"):
        return "NE†"
    return f"{r['ror']:.2f}"


def _b(x):
    return "<0.01" if x < 0.005 else f"{x:.2f}"


def fmt_ci(r, exact=False):
    lo, hi = (r["exact_lo"], r["exact_hi"]) if exact else (r["lo"], r["hi"])
    if r["boundary"].startswith("lower") or (exact and r["relaxer_n"] == 0):
        return f"0 to {hi:.2f}"
    if r["boundary"].startswith("upper") or (exact and r["keratin_n"] == 0):
        return f"{_b(lo)} to ∞"
    return f"{_b(lo)}-{hi:.2f}"


def fmt_n(n, N):
    return f"{n}/{N} ({100 * n / N:.1f})"


def fmt_mh(o, lo, hi):
    return "NE (zero events in one cohort)" if pd.isna(o) else f"{o:.2f} ({lo:.2f}-{hi:.2f})"


def pretty(res):
    return pd.DataFrame({
        "Reaction (MedDRA PT)": res["reaction"].map(disp),
        "Relaxer/alkali, n/N (%)": [fmt_n(a, b) for a, b in zip(res.relaxer_n, res.relaxer_N)],
        "Keratin/blowout, n/N (%)": [fmt_n(a, b) for a, b in zip(res.keratin_n, res.keratin_N)],
        "ROR": [fmt_ror(r) for _, r in res.iterrows()],
        "95% CI": [fmt_ci(r) for _, r in res.iterrows()],
        "CI method": res["ci_method"],
        "Exact conditional 95% CI": [fmt_ci(r, exact=True) for _, r in res.iterrows()],
        "P (Fisher)": res["p"].map(fmt_p),
        "q (BH-FDR)": res["q"].map(fmt_p),
        "FDR significant": np.where(res["fdr_sig"], "Yes", "No"),
        "Direction": res["direction"],
    })


# ======================================================================================
# 4. BRAND HARMONIZATION (rule based; full mapping exported for manual review)
# ======================================================================================

BRAND_RULES = [
    # keratin/blowout (order matters: specific brands before the generic "Brazilian Blowout")
    (r"CADIVEU", "Cadiveu"),
    (r"SIMPLY SMOOTH", "Simply Smooth"),
    (r"COPP?OL+A|COPOLLA|KERATIN COMPLEX|^CK HAIR", "Keratin Complex/Coppola"),
    (r"S[AU]+VE\b|SUVE\b", "Suave Keratin Infusion"),
    (r"\bOGX\b|ORGANIX", "OGX/Organix"),
    (r"TRESEMME", "TRESemmé"),
    (r"GLOBAL KERATIN|^GK HAIR", "GK Hair/Global Keratin"),
    (r"PRAVANA", "Pravana"),
    (r"CEZANNE", "Cezanne"),
    (r"GOLDWELL", "Goldwell"),
    (r"REJUVENOL", "Rejuvenol"),
    (r"COCOCHOCO", "Cocochoco"),
    (r"KERATIN CURE", "Keratin Cure"),
    (r"KERATIN EXPRESS", "Keratin Express"),
    (r"LIQUID KERATIN", "Liquid Keratin"),
    (r"EQUATE", "Equate"),
    (r"SCHWARZKOP", "Schwarzkopf"),
    (r"BRA[SZ]+I?L+[AI]*[AI]?N.*BLOW ?OU|BRAZLIAN BLOWOUT|BRAZILAIN|^THE BRAZILIAN BLOWOUT|GIB BRAZILIAN",
     "Brazilian Blowout (brand or procedure name)"),
    # relaxer/alkali
    (r"SOFT AND BEAUTIFUL|JUST FOR ME", "Soft & Beautiful/Just for Me"),
    (r"DARK (AND|N|&) LOVELY", "Dark and Lovely"),
    (r"OPTIMUM", "Optimum (SoftSheen-Carson)"),
    (r"ULTRA SHEEN", "Ultra Sheen"),
    (r"PRECISE", "Ultra Precise (SoftSheen-Carson)"),
    (r"MOTIONS", "Motions"),
    (r"\bTCB\b", "TCB"),
    (r"CREME OF NATURE", "Creme of Nature"),
    (r"REVLON", "Revlon/Realistic"),
    (r"AFRICA'?S BEST|KIDS ORGANICS", "Africa's Best"),
    (r"AFRICAN PRIDE", "African Pride"),
    (r"GENTLE ?TREATMENT", "Gentle Treatment"),
    (r"STRICKLAND", "J Strickland African Gold"),
    (r"MIZANI", "Mizani"),
    (r"\bORS\b|OLIVE MIRACLE", "ORS/Olive Miracle"),
    (r"PROFECTIV", "Profectiv"),
    (r"SILK ELEMENTS", "Silk Elements"),
    (r"DR\.? MIRACLE", "Dr. Miracle's"),
    (r"HAWAIIAN SILK", "Hawaiian Silk"),
    (r"ISOPLUS", "Isoplus"),
    (r"PRO-?LINE", "Pro-Line"),
    (r"LUSTRASILK", "Lustrasilk"),
    (r"LUSTRE|LUSTER|\bPINK\b", "Luster's Pink"),
]
GENERIC = re.compile(r"^(HAIR )?(PERM|RELAXER|PERM SOLUTION \d|HAIR PERM|KERATIN|KERATIN (HAIR )?TREATMENT|"
                     r"BRAZILIAN KERATIN TREATMENT|KERATIN HAIR STRAIGHTENING ?TREATMENT|ORIGINAL KERATIN|"
                     r"NO-LYE CONDITIONING CREME RELAXER|HAIR RELAXER|KERATIN SMOOTHING (SHAMPOO|CONDITIONER|COMPLEX)|"
                     r"KERATIN TREATMENT REJUVENOL)$")


def brand_of(name):
    n = re.sub(r"\s+", " ", str(name).upper()).strip()
    for pat, b in BRAND_RULES:
        if re.search(pat, n):
            return b
    if GENERIC.match(n):
        return "Unbranded/generic name"
    return "Other: " + n.title()[:45]


def keyword_products(row):
    """Product names that placed the report in its cohort: qualifying suspect product(s)
    when present, otherwise the keyword-matching product(s) among all listed products."""
    q = row["Qualifying suspect product(s)"]
    if pd.notna(q):
        return [x.strip() for x in str(q).split("|") if x.strip()], "Suspect"
    kws = ["RELAXER", "LYE", "PERM"] if row["cohort"] == RELAXER else ["KERATIN", "BLOWOUT"]
    out = []
    for part in str(row["All products"]).split("|"):
        nm = re.sub(r"^\s*(SUSPECT|CONCOMITANT)\s*—\s*", "", part).split("[code")[0].strip()
        if any(re.search(r"\b" + k + r"\b", nm.upper()) for k in kws):
            out.append(nm)
    return out, "Concomitant only"


# ======================================================================================
# 5. SYMPTOM CLUSTERS (term lists reviewed by a dermatologist on the author team)
# ======================================================================================

CLUSTERS = {
    "Ocular irritation": [
        "EYE IRRITATION", "LACRIMATION INCREASED", "OCULAR HYPERAEMIA", "EYE PRURITUS", "EYE PAIN",
        "EYE SWELLING", "EYELID IRRITATION", "BLEPHARITIS", "FOREIGN BODY SENSATION IN EYES",
        "EYE BURNS", "CHEMICAL EYE INJURY", "DRY EYE", "VISION BLURRED", "PERIORBITAL OEDEMA",
        "CORNEAL ABRASION"],
    "Respiratory/upper-airway irritation": [
        "DYSPNOEA", "DYSPNOEA EXERTIONAL", "COUGH", "PRODUCTIVE COUGH", "NASAL DISCOMFORT",
        "NASAL DRYNESS", "NASAL CONGESTION", "RHINORRHOEA", "SNEEZING", "EPISTAXIS",
        "THROAT IRRITATION", "THROAT TIGHTNESS", "DRY THROAT", "PHARYNGITIS", "OROPHARYNGEAL PAIN",
        "OROPHARYNGEAL SWELLING", "PHARYNGEAL OEDEMA", "DYSPHONIA", "APHONIA", "CHOKING",
        "RESPIRATORY DISTRESS", "RESPIRATORY DISORDER", "RESPIRATORY TRACT CONGESTION",
        "RESPIRATORY TRACT IRRITATION", "SINUS CONGESTION", "ASTHMA", "WHEEZING", "BRONCHOSPASM",
        "BRONCHIAL HYPERREACTIVITY", "BRONCHIAL IRRITATION", "INCREASED BRONCHIAL SECRETION",
        "INCREASED UPPER AIRWAY SECRETION", "PULMONARY CONGESTION", "AIRWAY BURNS",
        "PNEUMONITIS CHEMICAL"],
    "Hair loss/hair-shaft injury": [
        "ALOPECIA", "ALOPECIA EFFLUVIUM", "HYPOTRICHOSIS", "TRICHORRHEXIS", "HAIR TEXTURE ABNORMAL",
        "HAIR DISORDER"],
    "Cutaneous/scalp chemical injury": [
        "CHEMICAL INJURY", "CHEMICAL BURN OF SKIN", "CHEMICAL BURN", "CAUSTIC INJURY",
        "BURNS SECOND DEGREE", "THERMAL BURN", "SKIN BURNING SENSATION", "BLISTER", "SCAB",
        "SKIN INJURY", "SKIN ULCER", "SKIN EXFOLIATION", "PAIN OF SKIN", "SCAR", "WOUND"],
    "Neurologic/gastrointestinal symptoms": [
        "HEADACHE", "MIGRAINE", "CLUSTER HEADACHE", "DIZZINESS", "VERTIGO", "SYNCOPE", "PRESYNCOPE",
        "NAUSEA", "VOMITING"],
}

# openFDA queries used to build the cohorts (base URL https://api.fda.gov/food/event.json).
# Counts are from a re-run on the data version last updated 6 November 2026; the analysis used the version last
# updated 7 July 2026 (extracted 7 August 2026), and the re-run returned the same 352 report numbers and reaction terms.
OPENFDA_QUERIES = [
    ("Relaxer/alkali cohort (full records)",
     'search=products.industry_code:53+AND+(products.name_brand:"relaxer"+OR+products.name_brand:"lye"+OR+products.name_brand:"perm")&limit=1000', "76"),
    ("Keratin/blowout cohort (full records)",
     'search=products.industry_code:53+AND+(products.name_brand:"keratin"+OR+products.name_brand:"blowout")&limit=1000', "276"),
    ("Both cohorts combined (overlap check: equals 76 + 276, so no report is in both)",
     'search=products.industry_code:53+AND+(products.name_brand:"relaxer"+OR+products.name_brand:"lye"+OR+products.name_brand:"perm"+OR+products.name_brand:"keratin"+OR+products.name_brand:"blowout")&limit=1', "352"),
    ("All cosmetic reports in the endpoint (context)", "search=products.industry_code:53&limit=1",
     "52,223 (52,219 in the version used for analysis)"),
    ("Reaction frequencies, relaxer/alkali (check only; case-sensitive)",
     'search=products.industry_code:53+AND+(products.name_brand:"relaxer"+OR+products.name_brand:"lye"+OR+products.name_brand:"perm")&count=reactions.exact', "-"),
    ("Reaction frequencies, keratin/blowout (check only; case-sensitive)",
     'search=products.industry_code:53+AND+(products.name_brand:"keratin"+OR+products.name_brand:"blowout")&count=reactions.exact', "-"),
]

CHRONIC_OR_UNRELATED = {"LUNG NEOPLASM MALIGNANT", "ACUTE MYELOID LEUKAEMIA", "CHRONIC LYMPHOCYTIC LEUKAEMIA",
                        "THYROID CANCER", "THYROID NEOPLASM", "OVARIAN CANCER", "BREAST MASS",
                        "AUTOIMMUNE DISORDER", "INTRA-UTERINE DEATH", "ABORTION SPONTANEOUS",
                        "UTERINE LEIOMYOMA", "HEPATITIS", "DRUG-INDUCED LIVER INJURY", "LIVER INJURY",
                        "ANEURYSM", "RETINAL DETACHMENT", "VITREOUS DETACHMENT"}


# ======================================================================================
# 6. ALL ANALYSES -> TABLES
# ======================================================================================

def run_analyses(wb_path):
    df = load(wb_path)
    terms = tested_terms(df)
    sheets = {}
    key = {}   # machine-readable results used by the figures and QC

    # ---- primary and SA3 (co-primary), full results -------------------------------------
    primary = analyse(df, terms, "Primary (keyword-defined)")
    sa = {
        "SA1 (single cosmetic product)": df[df["SA1 keep"] == "Yes"],
        "SA2 (single suspect product)": df[df["SA2 keep"] == "Yes"],
        "SA3 (product-name validation)": df[df["SA3 keep"] == "Yes"],
        "SA4 (SA3 + single suspect)": df[df["SA4 keep"] == "Yes"],
    }
    sa_res = {k: analyse(v, terms, k) for k, v in sa.items()}
    sig10 = primary.loc[primary.fdr_sig].sort_values(["p", "reaction"], kind="mergesort")["reaction"].tolist()
    key.update(primary=primary, sa_res=sa_res, sig10=sig10, terms=terms)

    sheets["TableS1_Primary_all"] = pretty(primary.sort_values(["p", "reaction"], kind="mergesort"))
    sheets["TableS2_SA3_all"] = pretty(sa_res["SA3 (product-name validation)"].sort_values(["p", "reaction"], kind="mergesort"))

    # ---- exposure classification ----------------------------------------------------------
    kp = df.apply(keyword_products, axis=1)
    df["kw_products"] = kp.map(lambda x: x[0])
    df["kw_role"] = kp.map(lambda x: x[1])
    df["brands"] = df["kw_products"].map(lambda ps: sorted({brand_of(p) for p in ps}) or ["(none)"])
    df["brand"] = df["brands"].map(lambda b: b[0])

    def n_rk(d):
        return f"{(d.relaxer == 1).sum()} / {(d.relaxer == 0).sum()}"
    crit = pd.DataFrame([
        ["Primary: relaxer/alkali", 'Report lists a product with industry code 53 (Cosmetics) and a product whose brand-name field (products.name_brand) contains the word "relaxer", "lye", or "perm" (openFDA search; case-insensitive whole-word match, so "perm" does not match "permanent"); queries in Table S15', int(df.relaxer.sum())],
        ["Primary: keratin/blowout", 'Report lists a product with industry code 53 and a product whose brand-name field contains the word "keratin" or "blowout" (case-insensitive whole-word match); no report met both definitions', int((df.relaxer == 0).sum())],
        ["Keyword product was the suspect product", "Report-level role of the keyword-matching product",
         f"{(df.kw_role == 'Suspect').sum()} of {len(df)} (relaxer {((df.kw_role == 'Suspect') & (df.relaxer == 1)).sum()}/{df.relaxer.sum()}; "
         f"keratin {((df.kw_role == 'Suspect') & (df.relaxer == 0)).sum()}/{(df.relaxer == 0).sum()})"],
        ["SA1", "Exactly one cosmetic (industry code 53) product, and it matches the cohort keywords", n_rk(sa["SA1 (single cosmetic product)"])],
        ["SA2", "Exactly one suspect product; it is a cosmetic and matches the cohort keywords", n_rk(sa["SA2 (single suspect product)"])],
        ["SA3", "Conservative product-name review: suspect cosmetic product clearly identifiable as a chemical straightening product (relaxer, or keratin/blowout smoothing treatment); excludes perm-only, non-straightening keratin formats (shampoo, conditioner, color, fibers), ambiguous keratin-labeled products, Brazilian Blowout ancillary products, non-hair products, and reports without a qualifying suspect cosmetic", n_rk(sa["SA3 (product-name validation)"])],
        ["SA4", "SA3 criteria plus exactly one suspect product (the single-suspect restriction of SA2). The suspect product must satisfy SA3 but need not contain the cohort keyword, so two relaxer/alkali reports with a generically named suspect straightener (hair texturizer; hair straightener) and a concomitant keyword product are in SA4 but not SA2", n_rk(sa["SA4 (SA3 + single suspect)"])],
    ], columns=["Definition", "Criteria", "Reports (relaxer/alkali / keratin/blowout)"])

    ex = df[df["SA3 keep"] == "No"].copy()
    ex["reason_first"] = ex["Reason for conservative exclusion"].fillna("Not recorded").map(lambda s: s.split(" | ")[0])
    reasons = (ex.groupby(["reason_first", "cohort"]).size().unstack(fill_value=0)
                 .reindex(columns=[RELAXER, KERATIN], fill_value=0))
    reasons["Total"] = reasons.sum(axis=1)
    reasons = reasons.sort_values("Total", ascending=False).reset_index().rename(columns={"reason_first": "Primary reason for SA3 exclusion"})
    tot = pd.DataFrame([["Total excluded", reasons[RELAXER].sum(), reasons[KERATIN].sum(), reasons["Total"].sum()]], columns=reasons.columns)
    reasons = pd.concat([reasons, tot], ignore_index=True)

    def examples(sub, n=6):
        c = Counter(p.upper() for ps in sub["kw_products"] for p in ps)
        return "; ".join(f"{k.title()} ({v})" if v > 1 else k.title() for k, v in c.most_common(n))
    ex_rows = []
    for coh in (RELAXER, KERATIN):
        inc = df[(df.cohort == coh) & (df["SA3 keep"] == "Yes")]
        ex_rows.append([coh, "Included in SA3", examples(inc)])
        for rsn, g in ex[ex.cohort == coh].groupby("reason_first"):
            ex_rows.append([coh, "Excluded: " + rsn, examples(g, 4)])
    exmpl = pd.DataFrame(ex_rows, columns=["Cohort", "SA3 status", "Representative product names (n reports)"])

    classif = pd.DataFrame({
        "Report number": df["rn"], "Cohort": df["cohort"], "Date created": df["date"].dt.date,
        "Keyword-matching product(s)": df["kw_products"].map(" | ".join),
        "Role of keyword product": df["kw_role"], "Harmonized brand": df["brands"].map("; ".join),
        "SA1": df["SA1 keep"], "SA2": df["SA2 keep"], "SA3": df["SA3 keep"], "SA4": df["SA4 keep"],
        "SA3 exclusion reason": df["Reason for conservative exclusion"].fillna("").map(
            lambda x: " | ".join(dict.fromkeys(x.split(" | "))) if x else x),
    }).sort_values(["Cohort", "Date created"])
    sheets["TableS3a_Classification_criteria"] = crit
    sheets["TableS3b_SA3_exclusion_reasons"] = reasons
    sheets["TableS3c_Representative_names"] = exmpl
    sheets["TableS3d_Report_level_classif"] = classif
    key.update(df=df, reasons=reasons)

    # ---- effect estimates across exposure definitions -------------------------------------
    def cell(res, t):
        r = res.set_index("reaction").loc[t]
        return f"{fmt_ror(r)} ({fmt_ci(r)}); {lab_p('q', r['q'])}"

    def across(res_map):
        extra = sorted(t for t in terms if t not in sig10 and
                       any(bool(r.set_index("reaction").loc[t, "fdr_sig"]) for r in res_map.values()))
        rows = [[disp(t)] + [cell(res_map[k], t) for k in res_map] for t in sig10 + extra]
        cols = ["Reaction"] + [f"{k}\nN={int(v.relaxer_N.iloc[0])}/{int(v.keratin_N.iloc[0])}" for k, v in res_map.items()]
        return pd.DataFrame(rows, columns=cols)
    sheets["TableS4_Estimates_SA1_SA4"] = across({"Primary (keyword-defined)": primary, **sa_res})

    # ---- possible duplicates ---------------------------------------------------------------
    d2 = df.copy()
    d2["rxkey"] = d2["rx"].map(lambda s: "|".join(sorted(s)))
    clusters, cid = [], 0
    for _, g in d2.groupby(["cohort", "date", "rxkey", "brand"]):
        if len(g) < 2:
            continue
        ages = set(g["age"].dropna())
        sexes = set(g["Gender"].dropna()) - {"Unknown"}
        if len(ages) > 1 or len(sexes) > 1:
            continue
        cid += 1
        keep = g.sort_values(["age", "rn"], na_position="last").iloc[0]["rn"]
        for _, r in g.iterrows():
            clusters.append(dict(cluster=cid, rn=r.rn, cohort=r.cohort, date=r.date.date(),
                                 product=" | ".join(r.kw_products), sex=r.Gender,
                                 age=r.age, reactions="; ".join(sorted(r.rx)),
                                 retained="Yes" if r.rn == keep else "No (removed in de-duplicated analysis)"))
    dup = pd.DataFrame(clusters, columns=["cluster", "rn", "cohort", "date", "product", "sex", "age", "reactions", "retained"])
    drop_ids = set(dup.loc[dup.retained != "Yes", "rn"])
    sheets["TableS5_Possible_duplicates"] = dup.rename(columns={
        "cluster": "Cluster", "rn": "Report number", "cohort": "Cohort", "date": "Date created", "product": "Product",
        "sex": "Sex", "age": "Age (y)", "reactions": "Reactions", "retained": "Retained"})
    key.update(dup=dup, drop_ids=drop_ids)

    # ---- additional sensitivity analyses ---------------------------------------------------
    extra = {
        "Possible duplicates removed": df[~df.rn.isin(drop_ids)],
        "Excluding 2010-2012": df[~df.year.isin(PEAK_YEARS)],
        "2013-2023 only": df[df.year >= 2013],
        "2004-2023 only": df[df.year >= 2004],
        "Female reports only": df[df.Gender == "Female"],
    }
    extra_res = {k: analyse(v, terms, k) for k, v in extra.items()}
    sheets["TableS6_Additional_sensitivity"] = across({"Primary (keyword-defined)": primary, **extra_res})
    key.update(extra_res=extra_res)

    # ---- period-stratified and age-stratified/adjusted -------------------------------------
    df["period"] = pd.cut(df.year, [2002, 2009, 2012, 2023], labels=["2003-2009", "2010-2012", "2013-2023"])
    ag = df[df.age.notna()].copy()
    ag["agegrp"] = pd.cut(ag.age, [0, 34, 49, 200], labels=["<35", "35-49", ">=50"])
    strat_rows, period_mh = [], {}
    for t in sig10:
        crude = primary.set_index("reaction").loc[t]
        o, lo, hi, p_p = mh_stratified(df, "period", t)
        period_mh[t] = (o, lo, hi, p_p)
        oa, loa, hia, p_a = mh_stratified(ag, "agegrp", t)
        ag["y"] = ag["rx"].map(lambda s: int(t in s))
        a_n, k_n = int(ag[ag.relaxer == 1].y.sum()), int(ag[ag.relaxer == 0].y.sum())
        if a_n > 0 and k_n > 0:
            try:
                m = sm.Logit(ag["y"], sm.add_constant(ag[["relaxer", "age", "female"]].astype(float))).fit(disp=0)
                ci = np.exp(m.conf_int().loc["relaxer"])
                logit = f"{np.exp(m.params['relaxer']):.2f} ({ci.iloc[0]:.2f}-{ci.iloc[1]:.2f}); {lab_p('P', m.pvalues['relaxer'])}"
            except Exception:
                logit = "Not estimable"
        else:
            logit = "Not estimable (0 events in one cohort)"
        strat_rows.append([disp(t), f"{fmt_ror(crude)} ({fmt_ci(crude)})", fmt_mh(o, lo, hi), fmt_p(p_p),
                           f"{a_n}/{int(ag.relaxer.sum())} vs {k_n}/{int((ag.relaxer == 0).sum())}",
                           fmt_mh(oa, loa, hia), fmt_p(p_a), logit])
    sheets["TableS7_Stratified_adjusted"] = pd.DataFrame(strat_rows, columns=[
        "Reaction", "Crude ROR (95% CI), all reports",
        "Period-stratified MH OR (95% CI)", "Period-stratified CMH P",
        "Events among reports with age (relaxer vs keratin)", "Age-group-stratified MH OR (95% CI)",
        "Age-stratified CMH P", "Logistic OR adjusted for age and sex (95% CI); P"])
    key.update(period_mh=period_mh)

    # ---- product-level concentration and leave-top-brand-out -------------------------------
    conc, loo, loo_num = [], [], {}
    for t in sig10:
        r = primary.set_index("reaction").loc[t]
        hi_rel = r["direction"].startswith("Higher in relaxer")
        arm = df[df.relaxer == (1 if hi_rel else 0)]
        s = arm[arm.rx.map(lambda x: t in x)]
        bc = Counter(b for bs in s.brands for b in set(bs))
        named = {k: v for k, v in bc.items() if not k.startswith("Unbranded")}
        top, topn = max(named.items(), key=lambda kv: kv[1]) if named else ("-", 0)
        top3 = sum(v for _, v in sorted(named.items(), key=lambda kv: -kv[1])[:3])
        conc.append([disp(t), RELAXER if hi_rel else KERATIN, len(s), len(named),
                     bc.get("Unbranded/generic name", 0), top, f"{topn} ({100 * topn / len(s):.0f}%)",
                     f"{top3} ({100 * top3 / len(s):.0f}%)",
                     f"{s.year.isin(PEAK_YEARS).sum()} ({100 * s.year.isin(PEAK_YEARS).mean():.0f}%)",
                     f"{int(s.year.min())}-{int(s.year.max())}"])
        sub = df[~((df.relaxer == (1 if hi_rel else 0)) & df.brands.map(lambda bs: top in bs))]
        a, b, c, dd = two_by_two(sub, t)
        st = ror_stats(a, b, c, dd)
        row = dict(relaxer_n=a, keratin_n=c, **st)
        loo_num[t] = dict(row, top=top)
        loo.append([disp(t), top, f"{a}/{a + b}", f"{c}/{c + dd}",
                    fmt_ror(row), fmt_ci(row), fmt_p(st["p"]),
                    "Yes" if (st["p"] < 0.05) else "No",
                    "Same" if (a / (a + b) > c / (c + dd)) == hi_rel else "Reversed"])
    sheets["TableS8a_Product_concentration"] = pd.DataFrame(conc, columns=[
        "Reaction", "Cohort with higher reporting", "Reports with reaction", "Distinct named brands",
        "Reports with unbranded/generic name", "Top contributing brand", "Top brand, n (%)",
        "Top 3 brands, n (%)", "Reports from 2010-2012, n (%)", "Years spanned"])
    sheets["TableS8b_Leave_top_brand_out"] = pd.DataFrame(loo, columns=[
        "Reaction", "Brand removed (all its reports)", "Relaxer/alkali n/N", "Keratin/blowout n/N",
        "ROR", "95% CI", "P (Fisher, unadjusted)", "P<.05", "Direction vs primary"])
    brand_tab = (df.explode("brands").groupby(["cohort", "brands"]).size().rename("Reports")
                   .reset_index().sort_values(["cohort", "Reports"], ascending=[True, False]))
    sheets["TableS8c_Brand_composition"] = brand_tab.rename(columns={"cohort": "Cohort", "brands": "Harmonized brand"})
    mapping = sorted({(p.upper(), brand_of(p)) for ps in df.kw_products for p in ps})
    sheets["REVIEW_Brand_mapping"] = pd.DataFrame(mapping, columns=["Product name as reported", "Assigned brand (verify)"])
    key.update(loo=loo_num)

    # ---- inclusion-threshold sensitivity ---------------------------------------------------
    rules = [(3, "either"), (5, "either"), (10, "either"), (10, "combined"), (20, "combined")]
    thr_rows, thr_detail = [], {}
    for n, mode in rules:
        tt = tested_terms(df, n, mode)
        res = analyse(df, tt, f">={n} {mode}")
        label = (f">={n} reports in {'either cohort' if mode == 'either' else 'total (both cohorts combined)'}"
                 + (" (primary)" if (n, mode) == (5, "either") else ""))
        sig = res[res.fdr_sig].reaction.tolist()
        thr_rows.append([label, len(tt), len(sig), ", ".join(disp(s) for s in sorted(sig))])
        thr_detail[label] = res.set_index("reaction")["q"]
    sheets["TableS9a_Threshold_summary"] = pd.DataFrame(thr_rows, columns=["Inclusion rule", "Reactions tested", "FDR-significant", "Significant reactions"])
    sheets["TableS9b_Threshold_q_values"] = pd.DataFrame(
        {lab: [fmt_p(s.get(t, np.nan)) if t in s.index else "not tested" for t in sig10] for lab, s in thr_detail.items()},
        index=[disp(t) for t in sig10]).reset_index().rename(columns={"index": "Reaction (q value)"})

    # ---- co-occurrence -----------------------------------------------------------------------
    K = df[df.relaxer == 0]
    A = df[df.relaxer == 1]
    ker_sig = [t for t in sig10 if primary.set_index("reaction").loc[t, "direction"].startswith("Higher in keratin")]
    rel_sig = [t for t in sig10 if t not in ker_sig]

    def comat(d, ts):
        m = pd.DataFrame(0, index=[disp(t) for t in ts], columns=[disp(t) for t in ts])
        for i in ts:
            for j in ts:
                m.loc[disp(i), disp(j)] = int(d.rx.map(lambda s: i in s and j in s).sum())
        return m.reset_index().rename(columns={"index": f"Reports with both (diagonal = total), N={len(d)}"})
    sheets["TableS10a_Cooccur_keratin"] = comat(K, ker_sig)
    sheets["TableS10b_Cooccur_relaxer"] = comat(A, rel_sig + [t for t in ker_sig if t == "HEADACHE"])
    nsig = df.assign(k=df.rx.map(lambda s: len(s & set(sig10))))
    sheets["TableS10c_Signals_per_report"] = (nsig.groupby(["cohort", "k"]).size().unstack(0, fill_value=0)
                                              .reindex(columns=[RELAXER, KERATIN], fill_value=0).reset_index()
                                              .rename(columns={"k": "Number of the 10 primary signals in report"}))
    key.update(three_way=int(K.rx.map(lambda s: {"COUGH", "NASAL DISCOMFORT", "NAUSEA"} <= s).sum()))

    # ---- symptom clusters --------------------------------------------------------------------
    sheets["TableS11a_Cluster_definitions"] = pd.DataFrame(
        [[c, len(v), "; ".join(disp(t) for t in v)] for c, v in CLUSTERS.items()],
        columns=["Cluster", "Terms", "MedDRA preferred terms included"])
    is_bb = df.brands.map(lambda bs: any(b.startswith("Brazilian Blowout") for b in bs)) & (df.relaxer == 0)
    cl_rows, cl_num = [], []
    for lab, d in [("Primary", df), ("SA1", sa["SA1 (single cosmetic product)"]), ("SA2", sa["SA2 (single suspect product)"]),
                   ("SA3", sa["SA3 (product-name validation)"]), ("SA4", sa["SA4 (SA3 + single suspect)"]),
                   ("Excluding 2010-2012", extra["Excluding 2010-2012"]),
                   ("2013-2023 only", extra["2013-2023 only"]),
                   ("Brazilian Blowout reports removed", df[~is_bb]),
                   ("Brazilian Blowout reports and 2010-2012 removed", df[~is_bb & ~df.year.isin(PEAK_YEARS)])]:
        for c, ts in CLUSTERS.items():
            a, b, cc, d4 = two_by_two(flag_terms(d, ts), "X")
            st = ror_stats(a, b, cc, d4)
            row = dict(relaxer_n=a, keratin_n=cc, **st)
            cl_num.append(dict(row, analysis=lab, cluster=c))
            cl_rows.append([lab, c, fmt_n(a, a + b), fmt_n(cc, cc + d4), fmt_ror(row), fmt_ci(row), fmt_p(st["p"])])
    cl_res = pd.DataFrame(cl_rows, columns=["Analysis", "Cluster", "Relaxer/alkali, n/N (%)",
                                            "Keratin/blowout, n/N (%)", "ROR", "95% CI", "P (Fisher)"])
    C = CLUSTERS
    variants = [
        ("Cutaneous/scalp chemical injury", "As defined", C["Cutaneous/scalp chemical injury"]),
        ("Cutaneous/scalp chemical injury", "Adding burning sensation (site unspecified)",
         C["Cutaneous/scalp chemical injury"] + ["BURNING SENSATION"]),
        ("Ocular irritation", "As defined", C["Ocular irritation"]),
        ("Ocular irritation", "Removing vision blurred", [t for t in C["Ocular irritation"] if t != "VISION BLURRED"]),
        ("Ocular irritation", "Adding eye disorder, eyelid disorder, visual impairment",
         C["Ocular irritation"] + ["EYE DISORDER", "EYELID DISORDER", "VISUAL IMPAIRMENT"]),
        ("Respiratory/upper-airway irritation", "As defined", C["Respiratory/upper-airway irritation"]),
        ("Respiratory/upper-airway irritation", "Adding sinusitis, nasopharyngitis, bronchitis (infection terms)",
         C["Respiratory/upper-airway irritation"] + ["SINUSITIS", "NASOPHARYNGITIS", "BRONCHITIS"]),
        ("Hair loss/hair-shaft injury", "As defined", C["Hair loss/hair-shaft injury"]),
        ("Hair loss/hair-shaft injury", "Alopecia and trichorrhexis only", ["ALOPECIA", "TRICHORRHEXIS"]),
        ("Not clustered (reference)", "Dermatitis-type terms: pruritus, rash (any), erythema, skin irritation, dermatitis (any), skin reaction",
         ["PRURITUS", "RASH", "RASH PAPULAR", "RASH PRURITIC", "RASH SCALY", "ERYTHEMA", "SKIN IRRITATION",
          "DERMATITIS", "DERMATITIS CONTACT", "SKIN REACTION"]),
        ("Not clustered (reference)", "Burning sensation alone", ["BURNING SENSATION"]),
    ]
    vrows = []
    for c, lab, ts in variants:
        a, b, cc, d4 = two_by_two(flag_terms(df, ts), "X")
        st = ror_stats(a, b, cc, d4)
        row = dict(relaxer_n=a, keratin_n=cc, **st)
        vrows.append([c, lab, fmt_n(a, a + b), fmt_n(cc, cc + d4), fmt_ror(row), fmt_ci(row), fmt_p(st["p"])])
    mh_rows = []
    for c, ts in CLUSTERS.items():
        o, lo, hi, p_ = mh_stratified(flag_terms(df, ts), "period", "X")
        cl_num.append(dict(analysis="Period-stratified (MH)", cluster=c, ror=o, lo=lo, hi=hi, p=p_,
                           boundary="" if not pd.isna(o) else "lower", relaxer_n=np.nan, keratin_n=np.nan))
        crude = cl_res[(cl_res.Analysis == "Primary") & (cl_res.Cluster == c)].iloc[0]
        mh_rows.append([c, f"{crude.ROR} ({crude['95% CI']})", fmt_mh(o, lo, hi), fmt_p(p_)])
    sheets["TableS11b_Cluster_results"] = cl_res
    sheets["TableS11c_Cluster_term_sensitivity"] = pd.DataFrame(vrows, columns=[
        "Cluster", "Term-list variant (primary cohorts)", "Relaxer/alkali, n/N (%)", "Keratin/blowout, n/N (%)",
        "ROR", "95% CI", "P (Fisher)"])
    sheets["TableS11d_Cluster_period_stratified"] = pd.DataFrame(mh_rows, columns=[
        "Cluster", "Crude ROR (95% CI)", "Period-stratified MH OR (95% CI)", "Period-stratified CMH P"])
    key.update(cl_res=cl_res, cl_num=pd.DataFrame(cl_num))

    # ---- serious outcomes ---------------------------------------------------------------------
    ser = df[(df["Death"] == "Yes") | (df["Life Threatening"] == "Yes") | (df["Disability"] == "Yes")].copy()
    ser["Outcome(s)"] = ser.apply(lambda r: ", ".join(o for o in ["Death", "Life Threatening", "Disability", "Hospitalization"] if r[o] == "Yes"), axis=1)
    ser["flag"] = ser.rx.map(lambda s: "; ".join(disp(t) for t in sorted(s & CHRONIC_OR_UNRELATED)))
    sheets["TableS12_Serious_outcomes"] = pd.DataFrame({
        "Report number": ser.rn, "Cohort": ser.cohort, "Date created": ser.date.dt.date,
        "Harmonized brand": ser.brands.map("; ".join), "Product(s)": ser.kw_products.map(" | ".join),
        "Outcome(s)": ser["Outcome(s)"], "Reported reactions": ser.rx.map(lambda s: "; ".join(disp(t) for t in sorted(s))),
        "Terms unlikely to reflect acute product exposure (author review)": ser.flag,
    }).sort_values(["Cohort", "Date created"])
    ser_sum = []
    out_label = {"Life Threatening": "Life-threatening", "Visited Emergency Room": "Emergency department visit"}
    for o in ["Death", "Life Threatening", "Disability", "Hospitalization", "Visited Emergency Room"]:
        g = df[df[o] == "Yes"]
        ser_sum.append([out_label.get(o, o), (g.relaxer == 1).sum(), (g.relaxer == 0).sum(),
                        ", ".join(f"{k}: {v}" for k, v in g.year.value_counts().sort_index().items()),
                        ", ".join(f"{k.replace('Other: ', '')} ({v})"
                                  for k, v in Counter(b for bs in g.brands for b in bs).most_common(3))])
    sheets["TableS12b_Serious_summary"] = pd.DataFrame(ser_sum, columns=["Outcome", "Relaxer/alkali", "Keratin/blowout", "Reports by year", "Most frequent brands"])

    # ---- denominators and period estimates ----------------------------------------------------
    focus = ["ALOPECIA", "EYE IRRITATION", "DYSPNOEA", "DIZZINESS", "NASAL DISCOMFORT"]
    ann = []
    for y in range(int(df.year.min()), int(df.year.max()) + 1):
        g = df[df.year == y]
        row = {"Year": y, "Relaxer/alkali reports": int((g.relaxer == 1).sum()), "Keratin/blowout reports": int((g.relaxer == 0).sum())}
        for t in focus:
            for coh, flag in ((RELAXER, 1), (KERATIN, 0)):
                row[f"{disp(t)}, {coh} n"] = int(g[g.relaxer == flag].rx.map(lambda s: t in s).sum())
        ann.append(row)
    sheets["TableS13a_Annual_counts"] = pd.DataFrame(ann)
    per = []
    periods = {"2003-2009": (2003, 2009), "2010-2012": (2010, 2012), "2013-2017": (2013, 2017), "2018-2023": (2018, 2023)}
    for lab, (y0, y1) in periods.items():
        g = df[(df.year >= y0) & (df.year <= y1)]
        for coh, flag in ((RELAXER, 1), (KERATIN, 0)):
            gg = g[g.relaxer == flag]
            for t in focus:
                n = int(gg.rx.map(lambda s: t in s).sum())
                N = len(gg)
                lo, hi = proportion_confint(n, N, method="wilson") if N else (np.nan, np.nan)
                per.append([lab, coh, disp(t), n, N, f"{100 * n / N:.1f}" if N else "NA",
                            f"{100 * lo:.1f}-{100 * hi:.1f}" if N else "NA", lo, hi])
    per = pd.DataFrame(per, columns=["Period", "Cohort", "Reaction", "n", "N", "%", "Wilson 95% CI (%)", "_lo", "_hi"])
    sheets["TableS13b_Period_proportions"] = per.drop(columns=["_lo", "_hi"])
    key.update(annual=pd.DataFrame(ann), period_prop=per)

    # ---- study period checks ------------------------------------------------------------------
    sheets["TableS14_Study_period_checks"] = pd.DataFrame([
        ["Earliest report (date_created)", str(df.date.min().date()), df.loc[df.date.idxmin(), "rn"]],
        ["Latest report (date_created)", str(df.date.max().date()), df.loc[df.date.idxmax(), "rn"]],
        ["Reports created in 2003", int((df.year == 2003).sum()), ", ".join(df.loc[df.year == 2003, "rn"])],
        ["Reports created on/after 2023-12-29 (MoCRA)", int((df.date >= MOCRA_DATE).sum()), ""],
        ["Distinct report numbers", df.rn.nunique(), f"of {len(df)} reports"],
    ], columns=["Check", "Value", "Detail"])

    sheets["TableS15_openFDA_queries"] = pd.DataFrame(
        [[i + 1, purpose, q, n] for i, (purpose, q, n) in enumerate(OPENFDA_QUERIES)],
        columns=["Query", "Purpose", "API call (https://api.fda.gov/food/event.json?...)", "Reports returned (re-run, November 6, 2026)"])

    # ---- draft revised Table 2 with SA3 co-primary columns (author-only) ----------------------
    s3 = sa_res["SA3 (product-name validation)"].set_index("reaction")
    t2 = []
    for t in sig10:
        r = primary.set_index("reaction").loc[t]
        q = s3.loc[t]
        t2.append([disp(t), fmt_n(r.relaxer_n, r.relaxer_N), fmt_n(r.keratin_n, r.keratin_N),
                   fmt_ror(r), fmt_ci(r), fmt_p(r.p), fmt_p(r.q),
                   fmt_n(q.relaxer_n, q.relaxer_N), fmt_n(q.keratin_n, q.keratin_N),
                   fmt_ror(q), fmt_ci(q), fmt_p(q.q)])
    sheets["DRAFT_Table2_revised"] = pd.DataFrame(t2, columns=[
        "Reaction", "Primary: relaxer/alkali n/N (%)", "Primary: keratin/blowout n/N (%)", "Primary ROR",
        "Primary 95% CI", "Primary P", "Primary q", "SA3: relaxer/alkali n/N (%)", "SA3: keratin/blowout n/N (%)",
        "SA3 ROR", "SA3 95% CI", "SA3 q"])

    key.update(sa=sa, extra=extra, is_bb=is_bb)
    return key, sheets


def _order(name):
    m = re.match(r"TableS(\d+)([a-z]?)_", name)
    return (0, int(m.group(1)), m.group(2)) if m else (1, 0, name)


def write_excel(sheets, n_terms, out_path, working_path):
    notes = pd.DataFrame([
        ["Source", "FDA CAERS cosmetic reports (industry code 53) via openFDA /food/event (data version last updated July 7, 2026; extracted August 7, 2026). Queries in Table S15."],
        ["Unit of analysis", "Report. Each distinct case-normalized MedDRA preferred term counted once per report; missing age and sex were not imputed."],
        ["Comparison", "ROR = odds of a reaction among relaxer/alkali reports vs keratin/blowout reports (ROR >1 = higher in relaxer/alkali)."],
        ["CIs", "Woolf (log) 95% CI when all cells > 0; exact conditional 95% CI when any cell = 0. Exact conditional CIs are also shown for every reaction."],
        ["Zero cells", "NE* = not estimable because no relaxer/alkali report listed the reaction (conditional MLE on the lower boundary, 0); the exact 95% CI upper limit is shown. In Tables S7 and S11d, NE marks stratified or adjusted estimates that are not estimable because no relaxer/alkali report listed the reaction."],
        ["Testing family", f"Reactions with >=5 reports in either cohort of the primary analysis ({n_terms} reactions); the same family and BH-FDR correction are used in every analysis that retests all reactions; stratified, adjusted, brand-removal, and symptom-cluster comparisons use unadjusted P<.05."],
        ["Possible duplicates", "Same cohort, same date_created, identical reaction set, same harmonized brand, and no conflicting non-missing age or sex; one report retained per cluster (age recorded preferred, then lowest report number)."],
        ["Brands", "Rule-based harmonization of keyword-matching product names, verified by the authors. 'Brazilian Blowout' may denote the brand of that name or the procedure generically."],
        ["Symptom clusters", "Term lists are given in Table S11a; Table S11c shows results under alternative term lists. Symptom clusters were defined post hoc and are a secondary analysis."],
        ["Code", "hair_straightening_supplement.py regenerates every table and figure from the report-level workbook."],
    ], columns=["Item", "Note"])
    journal = sorted([n for n in sheets if n.startswith("TableS")], key=_order)
    working = [n for n in sheets if not n.startswith("TableS")]
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        notes.to_excel(xw, sheet_name="README", index=False)
        for name in journal:
            sheets[name].to_excel(xw, sheet_name=name[:31], index=False)
    with pd.ExcelWriter(working_path, engine="openpyxl") as xw:   # author-only: do not upload
        for name in working:
            sheets[name].to_excel(xw, sheet_name=name[:31], index=False)


# ======================================================================================
# 7. SUPPLEMENTARY FIGURES
# ======================================================================================

def make_figures(key, out_dir):
    plt = _mpl(8)                      # same font and settings as the main figures (section 7b)
    from matplotlib.lines import Line2D

    plt.rcParams.update({"axes.edgecolor": MUTED, "axes.labelcolor": INK, "xtick.color": MUTED, "ytick.color": INK})
    primary, sa_res, extra_res = key["primary"], key["sa_res"], key["extra_res"]
    sig10 = key["sig10"]
    paths = {}

    def forest_panel(ax, rows, color, xlim=(0.004, 60)):
        """rows: list of (label, ror, lo, hi, significant, zero_relaxer). Log x axis."""
        ys = np.arange(len(rows))[::-1]
        for y, (lab, ror, lo, hi, sig, zero) in zip(ys, rows):
            if zero:                                  # boundary estimate: no marker, bar to upper limit
                ax.annotate("", xy=(xlim[0] * 1.05, y), xytext=(min(hi, xlim[1]), y),
                            arrowprops=dict(arrowstyle="-|>", color=color, lw=1.2, mutation_scale=7,
                                            alpha=1 if sig else 0.4))
                ax.text(min(hi, xlim[1]) * 1.15, y, "NE" if sig else "NE, ns", va="center", fontsize=6.5, color=MUTED)
                continue
            if pd.isna(ror):
                continue
            ax.plot([max(lo, xlim[0]), min(hi, xlim[1])], [y, y], color=color, lw=1.2, solid_capstyle="round")
            ax.plot(ror, y, "o", ms=4.5, color=color if sig else "white", mec=color, mew=1.2, zorder=3)
        ax.axvline(1, color=MUTED, lw=0.8, ls="--", zorder=0)
        ax.set_xscale("log")
        ax.set_xlim(*xlim)
        ax.set_ylim(-0.7, len(rows) - 0.3)
        ax.set_yticks(ys)
        ax.set_yticklabels([r[0] for r in rows], fontsize=7)
        ax.set_xticks([0.01, 0.1, 1, 10])
        ax.set_xticklabels(["0.01", "0.1", "1", "10"])
        ax.grid(axis="x", color=GRID, lw=0.5)
        ax.tick_params(axis="y", length=0)

    def res_row(lab, res, t, sig_col="fdr_sig"):
        r = res.set_index("reaction").loc[t]
        zero = r["relaxer_n"] == 0
        return (lab, r["ror"], r["lo"], r["hi"], bool(r[sig_col]), zero)

    # ---- Figure S1: the 10 primary signals across analyses ------------------------------------
    labels = ["Primary", "SA1", "SA2", "SA3", "SA4", "Duplicates removed", "Excl. 2010-2012",
              "2013-2023 only", "Period-stratified (MH)", "Top brand removed"]
    fig, axes = plt.subplots(2, 5, figsize=(13, 6.2), sharex=True)
    for ax, t in zip(axes.ravel(), sig10):
        hi_rel = primary.set_index("reaction").loc[t, "direction"].startswith("Higher in relaxer")
        col = COLOR[RELAXER] if hi_rel else COLOR[KERATIN]
        rows = [res_row("Primary", primary, t)]
        rows += [res_row(lab, sa_res[k], t) for lab, k in zip(["SA1", "SA2", "SA3", "SA4"], sa_res)]
        rows += [res_row("Duplicates removed", extra_res["Possible duplicates removed"], t),
                 res_row("Excl. 2010-2012", extra_res["Excluding 2010-2012"], t),
                 res_row("2013-2023 only", extra_res["2013-2023 only"], t)]
        o, lo, hi, p = key["period_mh"][t]
        rows.append(("Period-stratified (MH)", o, lo, hi, p < 0.05, pd.isna(o)))
        L = key["loo"][t]
        rows.append(("Top brand removed", L["ror"], L["lo"], L["hi"], L["p"] < 0.05, L["relaxer_n"] == 0))
        # period-stratified NE: draw nothing but a label (no CI available)
        forest_panel(ax, [r if not (r[0] == "Period-stratified (MH)" and r[5]) else (r[0], np.nan, np.nan, np.nan, r[4], False)
                          for r in rows], col)
        if pd.isna(o):
            ax.text(0.006, 1, "NE" if p < 0.05 else "NE, ns", va="center", fontsize=6.5, color=MUTED)
        ax.set_title(disp(t), fontsize=8.5, color=INK, loc="left", fontweight="bold")
    for ax in axes[1]:
        ax.set_xlabel("ROR (log scale)")
    for ax in axes[:, 1:].ravel():
        ax.set_yticklabels([])
    handles = [Line2D([], [], color=COLOR[RELAXER], marker="o", lw=1.2, label="Higher reporting with relaxer/alkali"),
               Line2D([], [], color=COLOR[KERATIN], marker="o", lw=1.2, label="Higher reporting with keratin/blowout"),
               Line2D([], [], color=MUTED, marker="o", mfc="white", lw=0, label="Open marker or faded arrow (ns): not significant")]
    fig.legend(handles=handles, loc="lower center", ncol=3, frameon=False, fontsize=8, bbox_to_anchor=(0.5, -0.01))
    fig.tight_layout(rect=(0, 0.05, 1, 1))
    paths["S1"] = os.path.join(out_dir, "FigureS1_signal_estimates_across_analyses")
    for ext in ("png", "pdf"):
        fig.savefig(f"{paths['S1']}.{ext}", dpi=300, bbox_inches="tight")
    plt.close(fig)

    # ---- Figure S2: period-level proportions with Wilson CIs ----------------------------------
    pp = key["period_prop"]
    rx_order = ["Alopecia", "Eye irritation", "Dyspnea", "Dizziness", "Nasal discomfort"]
    per_order = ["2003-2009", "2010-2012", "2013-2017", "2018-2023"]
    fig, axes = plt.subplots(1, 5, figsize=(13, 3.4), sharey=True)
    for ax, rx_ in zip(axes, rx_order):
        for k, (coh, off, mk) in enumerate(((RELAXER, -0.12, "o"), (KERATIN, 0.12, "s"))):
            g = pp[(pp.Reaction == rx_) & (pp.Cohort == coh)].set_index("Period").reindex(per_order)
            x = np.arange(len(per_order)) + off
            y = g.n / g.N * 100
            ax.vlines(x, g._lo * 100, g._hi * 100, color=COLOR[coh], lw=1.2)
            ax.plot(x, y, mk, ms=4.5, color=COLOR[coh], label=coh)
        ax.set_xticks(range(len(per_order)))
        Ns = [f"{p}\nR {int(pp[(pp.Period == p) & (pp.Cohort == RELAXER)].N.iloc[0])} / "
              f"K {int(pp[(pp.Period == p) & (pp.Cohort == KERATIN)].N.iloc[0])}" for p in per_order]
        ax.set_xticklabels(Ns, fontsize=6.5)
        ax.set_title(rx_, fontsize=8.5, loc="left", fontweight="bold", color=INK)
        ax.set_ylim(0, 100)
        ax.grid(axis="y", color=GRID, lw=0.5)
    axes[0].set_ylabel("Reports listing the reaction (%)")
    h, l_ = axes[0].get_legend_handles_labels()
    fig.legend(h, l_, loc="lower center", ncol=2, frameon=False, fontsize=8, bbox_to_anchor=(0.5, -0.04))
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    paths["S2"] = os.path.join(out_dir, "FigureS2_period_proportions")
    for ext in ("png", "pdf"):
        fig.savefig(f"{paths['S2']}.{ext}", dpi=300, bbox_inches="tight")
    plt.close(fig)

    # ---- Figure S3: symptom clusters across analyses -------------------------------------------
    cn = key["cl_num"]
    an_order = ["Primary", "SA1", "SA2", "SA3", "SA4", "Excluding 2010-2012", "2013-2023 only",
                "Brazilian Blowout reports removed", "Brazilian Blowout reports and 2010-2012 removed", "Period-stratified (MH)"]
    short = {"Excluding 2010-2012": "Excl. 2010-2012", "Brazilian Blowout reports removed": "Brazilian Blowout removed",
             "Brazilian Blowout reports and 2010-2012 removed": "BB and 2010-2012 removed"}
    fig, axes = plt.subplots(1, 5, figsize=(13, 3.8), sharex=True)
    for i, (ax, c) in enumerate(zip(axes, CLUSTERS)):
        g = cn[cn.cluster == c].set_index("analysis")
        prim = g.loc["Primary"]
        col = COLOR[RELAXER] if prim.relaxer_n / 76 > prim.keratin_n / 276 and prim.ror > 1 else COLOR[KERATIN]
        rows = []
        for a in an_order:
            r = g.loc[a]
            zero = str(r.boundary).startswith("lower") and a != "Period-stratified (MH)"
            rows.append((short.get(a, a), r.ror, r.lo, r.hi, r.p < 0.05, zero))
        forest_panel(ax, rows, col, xlim=(0.002, 60))
        ax.set_title(c, fontsize=8, loc="left", fontweight="bold", color=INK)
        ax.set_xlabel("ROR (log scale)")
        if i:
            ax.set_yticklabels([])
    fig.legend(handles=handles, loc="lower center", ncol=3, frameon=False, fontsize=8, bbox_to_anchor=(0.5, -0.04))
    fig.tight_layout(rect=(0, 0.07, 1, 1))
    paths["S3"] = os.path.join(out_dir, "FigureS3_cluster_estimates_across_analyses")
    for ext in ("png", "pdf"):
        fig.savefig(f"{paths['S3']}.{ext}", dpi=300, bbox_inches="tight")
    plt.close(fig)
    return {k: v + ".png" for k, v in paths.items()}


FIGURE_LEGENDS = {
    "S1": ("Effect estimates for the 10 primary signals across sensitivity analyses",
           "Each panel shows one reaction. Markers are RORs (Mantel-Haenszel OR for the period-stratified row) and lines are "
           "95% CIs on a log scale; the dashed line marks ROR=1. Filled markers indicate q<.05 (Benjamini-Hochberg, "
           "same 45-reaction family) or, for the period-stratified and top-brand-removed rows, P<.05. Color shows the "
           "cohort with higher reporting in the primary analysis. When no relaxer/alkali report listed the reaction, no "
           "point estimate exists (NE); an arrow extends from the exact upper 95% confidence limit toward 0, faded and marked "
           "ns when not significant. Values are "
           "given in Tables S1, S2, S4, S6, S7, and S8b. The sensitivity cohorts are nested subsets of the primary cohorts "
           "and are not independent replications."),
    "S2": ("Proportion of reports listing selected reactions by calendar period",
           "Points are the proportion of reports in each cohort and period that listed the reaction; vertical lines are "
           "Wilson 95% CIs. Numbers under each period give the number of relaxer/alkali (R) and keratin/blowout (K) reports. "
           "Proportions describe reporting, not incidence. Values are given in Table S13b."),
    "S3": ("Symptom-cluster estimates across sensitivity analyses",
           "Markers are RORs (Mantel-Haenszel OR for the period-stratified row) and lines are 95% CIs on a log scale. "
           "Filled markers indicate P<.05 (unadjusted Fisher exact or Cochran-Mantel-Haenszel test). NE, not estimable "
           "because no relaxer/alkali report listed a cluster term; the arrow extends from the exact upper 95% confidence "
           "limit (faded and marked ns when not significant). Cluster definitions are in Table S11a; values are in Tables S11b and S11d. Symptom clusters are a "
           "secondary analysis."),
}


# ======================================================================================
# 7b. MAIN FIGURES 1-4 (manuscript)
# ======================================================================================
# Style follows the submitted figures: Arial (or the first available metric-compatible substitute),
# 6.5-inch width, 600-dpi PNG plus PDF and EPS with embedded TrueType text. No transparency is used,
# so the EPS files look the same as the PNG and PDF files.

MAIN_DPI = 600
MAIN_WIDTH_IN = 6.5
MAIN_FORMATS = ("png", "pdf", "eps")
INK_MAIN, REF_LINE = "#1A1A2E", "#888888"
SIG_FILL, NS_FILL = "#2C6E7F", "#EDEDED"         # Figure 3 cells; SIG_FILL is also the Figure 1 header color
RX_PALETTE = ["#3B5BA5", "#B07A12", "#7B4FA0", "#2F8F5B"]   # Figure 4C reactions (checked for color-vision deficiency)
RX_MARKERS = ["o", "s", "D", "^"]
ALL_COSMETIC_REPORTS = 52219      # openFDA query 4 (Table S15) in the data version used for the analysis
FIG4_PERIODS = ["2003-2009", "2010-2012", "2013-2017", "2018-2023"]
FIG4C_REACTIONS = ["Eye irritation", "Dyspnea", "Dizziness", "Nasal discomfort"]

# Short labels for the SA3 exclusion reasons shown in Figure 1 (full wording in Table S3b).
SA3_REASON_SHORT = {
    "Keratin-only/ambiguous name; no explicit straightening context": "Keratin-only or ambiguous product name",
    "Keratin term in non-straightening product format": "Keratin term in a non-straightening product",
    "Perm-only product; not clearly a straightener/relaxer": "Perm-only product",
    "Brazilian Blowout ancillary product": "Brazilian Blowout ancillary product",
    "No suspect code-53 product met the conservative straightening definition.": "No qualifying suspect cosmetic product",
    "No clear keratin/blowout straightening terminology": "No clear straightening terminology",
    "No clear alkali-straightening terminology": "No clear straightening terminology",
    "Lye soap; not a hair straightener": "Not a hair product (lye soap, permanent makeup)",
    "Permanent makeup/tattoo; not a hair straightener": "Not a hair product (lye soap, permanent makeup)",
}


def _pick_font():
    import matplotlib.font_manager as fm
    available = {f.name for f in fm.fontManager.ttflist}
    for name in ("Arial", "Helvetica", "Liberation Sans", "Nimbus Sans", "DejaVu Sans"):
        if name in available:
            return name
    return "DejaVu Sans"


def _mpl(font_size=9):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({
        "font.family": _pick_font(), "font.size": font_size,
        "axes.edgecolor": "#333333", "axes.linewidth": 0.8, "axes.labelcolor": INK_MAIN,
        "xtick.color": "#333333", "ytick.color": INK_MAIN,
        "axes.spines.top": False, "axes.spines.right": False,
        "figure.facecolor": "white", "savefig.facecolor": "white",
        "pdf.fonttype": 42, "ps.fonttype": 42,
        "lines.solid_capstyle": "round", "lines.solid_joinstyle": "round",
    })
    return plt


def _save_main(fig, base):
    for ext in MAIN_FORMATS:
        kw = dict(bbox_inches="tight", pad_inches=0.08)
        if ext == "png":
            kw["dpi"] = MAIN_DPI
        fig.savefig(f"{base}.{ext}", **kw)


def tdisp(term):
    """Title-case display name used in the main figures (as in the submitted figures)."""
    return " ".join(w.capitalize() for w in disp(term).split())


def _log_ticks(vmin, vmax, max_ticks=6):
    """1-2-5 ticks on a log axis, thinned so that 1 is always kept."""
    cand = sorted(b * 10.0 ** k for k in range(math.floor(math.log10(vmin)) - 1, math.ceil(math.log10(vmax)) + 2)
                  for b in (1, 2, 5))
    cand = [t for t in cand if vmin * 0.98 <= t <= vmax * 1.02]
    step = max(1, math.ceil(len(cand) / max_ticks))
    i1 = min(range(len(cand)), key=lambda i: abs(math.log10(cand[i])))
    return cand[i1 % step::step]


def _tick_label(v):
    return f"{int(round(v))}" if abs(v - round(v)) < 1e-9 else f"{v:.3f}".rstrip("0").rstrip(".")


# ---- numbers behind Figures 1 and 3 (also used by the QC checks) -----------------------------

def figure1_counts(key):
    df, sa, dup = key["df"], key["sa"], key["dup"]
    out = dict(all=ALL_COSMETIC_REPORTS, N=len(df), R=int(df.relaxer.sum()), K=int((df.relaxer == 0).sum()),
               first=df.date.min(), last=df.date.max())
    for tag, k in zip(["SA1", "SA2", "SA3", "SA4"], sa):
        out[tag] = (int(sa[k].relaxer.sum()), int((sa[k].relaxer == 0).sum()))
    rem = dup[dup.retained != "Yes"]
    out["dup"] = dict(clusters=int(dup.cluster.nunique()), reports=len(dup), removed=len(rem),
                      R=int((rem.cohort == RELAXER).sum()), K=int((rem.cohort == KERATIN).sum()))
    rs = key["reasons"]
    rs = rs[rs.iloc[:, 0] != "Total excluded"].copy()
    rs["short"] = rs.iloc[:, 0].map(lambda s: SA3_REASON_SHORT.get(s, s))
    g = rs.groupby("short", sort=False)[[RELAXER, KERATIN]].sum()
    g = g.loc[(g[RELAXER] + g[KERATIN]).sort_values(ascending=False, kind="stable").index]
    out["sa3_reasons"] = [(lab, int(r[RELAXER]), int(r[KERATIN])) for lab, r in g.iterrows()]
    out["sa3_excluded"] = (int(g[RELAXER].sum()), int(g[KERATIN].sum()))
    return out


def figure3_matrix(key):
    """Rows: every reaction with q<.05 in at least one FDR-corrected analysis, grouped by the direction of the
    primary estimate. Columns: FDR-corrected analyses (q values, same 45-reaction family) followed by the
    period-stratified and top-brand-removed analyses (unadjusted P values; 10 primary signals only)."""
    P = key["primary"].set_index("reaction")
    sa, ex = key["sa_res"], key["extra_res"]
    qspec = [("Primary", "Keyword\ndefinition", key["primary"]),
             ("SA1", "1 cosmetic\nproduct", sa["SA1 (single cosmetic product)"]),
             ("SA2", "1 suspect\nproduct", sa["SA2 (single suspect product)"]),
             ("SA3", "Strict\nname", sa["SA3 (product-name validation)"]),
             ("SA4", "Strict +\n1 suspect", sa["SA4 (SA3 + single suspect)"]),
             ("Duplicates", "removed", ex["Possible duplicates removed"]),
             ("Temporal", "Excluding\n2010-2012", ex["Excluding 2010-2012"]),     # adjacent columns with the same
             ("Temporal", "2013-2023\nonly", ex["2013-2023 only"])]              # name share one header label
    rel_hi = lambda t: P.loc[t, "direction"].startswith("Higher in relaxer")
    cols = []
    for name, desc, res in qspec:
        R = res.set_index("reaction")
        cells = {t: dict(v=float(R.loc[t, "q"]), sig=bool(R.loc[t, "fdr_sig"]),
                         same=(R.loc[t, "direction"] == P.loc[t, "direction"])) for t in R.index}
        cols.append(dict(name=name, desc=desc, kind="q", cells=cells,
                         N=f"{int(res.relaxer_N.iloc[0])}/{int(res.keratin_N.iloc[0])}"))
    mh, loo = key["period_mh"], key["loo"]
    cells = {}
    for t, (o, lo, hi, p) in mh.items():
        same = (rel_hi(t) == (o > 1)) if not pd.isna(o) else not rel_hi(t)
        cells[t] = dict(v=float(p), sig=bool(p < 0.05), same=same)
    cols.append(dict(name="Stratified", desc="by period", kind="P", cells=cells,
                     N=f"{int(key['df'].relaxer.sum())}/{int((key['df'].relaxer == 0).sum())}"))
    cells = {t: dict(v=float(L["p"]), sig=bool(L["p"] < 0.05), same=(rel_hi(t) == (L["ror"] > 1)))
             for t, L in loo.items()}
    cols.append(dict(name="Top brand", desc="removed", kind="P", cells=cells, N="Varies"))

    qcols = [c for c in cols if c["kind"] == "q"]
    rows = [t for t in P.index if any(c["cells"][t]["sig"] for c in qcols)]
    nsig = {t: sum(c["cells"][t]["sig"] for c in qcols) for t in rows}
    order = lambda ts: sorted(ts, key=lambda t: (-nsig[t], P.loc[t, "q"], t))
    groups = [("Higher reporting with relaxer/alkali products (primary analysis)", order([t for t in rows if rel_hi(t)])),
              ("Higher reporting with keratin/blowout products (primary analysis)", order([t for t in rows if not rel_hi(t)]))]
    groups = [(g, ts) for g, ts in groups if ts]
    flipped = [(c["name"], t) for c in cols for t in rows if t in c["cells"] and c["cells"][t]["sig"] and not c["cells"][t]["same"]]
    return dict(cols=cols, groups=groups, flipped=flipped)


# ---- Figure 1: study flow ---------------------------------------------------------------------

def _figure1(key, base):
    plt = _mpl(8)
    from matplotlib.patches import FancyBboxPatch, Rectangle
    c = figure1_counts(key)
    W, EDGE, LINE_H, HEAD_H, R_ = MAIN_WIDTH_IN, INK_MAIN, 0.175, 0.3, 0.09
    fig = plt.figure(figsize=(W, 10))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")

    def box(x, top, w, header, lines, align="left", dashed=False, head_fs=8.5):
        """Draw a box whose top edge is at `top`; return its bottom edge. A line '---' draws a divider."""
        n_div = sum(ln == "---" for ln in lines)
        h = (HEAD_H if header else 0) + 0.12 + (len(lines) - n_div) * LINE_H + n_div * 0.12 + 0.07
        shape = dict(boxstyle=f"round,pad=0,rounding_size={R_}")
        outer = FancyBboxPatch((x, top - h), w, h, fc="white", ec="none", **shape)
        ax.add_patch(outer)
        if header:
            band = Rectangle((x, top - HEAD_H), w, HEAD_H, fc=SIG_FILL, ec="none")
            ax.add_patch(band)
            band.set_clip_path(outer)
            ax.plot([x, x + w], [top - HEAD_H] * 2, color=EDGE, lw=1.0)
            ax.text(x + w / 2, top - HEAD_H / 2, header, ha="center", va="center", color="white",
                    fontweight="bold", fontsize=head_fs)
        ax.add_patch(FancyBboxPatch((x, top - h), w, h, fc="none", ec=EDGE, lw=1.2,
                                    ls=(0, (4, 2.5)) if dashed else "-", **shape))
        y = top - (HEAD_H if header else 0) - 0.12
        for ln in lines:
            if ln == "---":
                ax.plot([x + 0.1, x + w - 0.1], [y - 0.02] * 2, color=EDGE, lw=0.8)
                y -= 0.12
                continue
            bold = ln.startswith("**")
            ax.text(x + (w / 2 if align == "center" else 0.14), y - LINE_H / 2, ln.strip("*"),
                    ha="center" if align == "center" else "left", va="center", fontsize=8,
                    fontweight="bold" if bold else "normal", color=INK_MAIN)
            y -= LINE_H
        return top - h

    def arrow(x0, y0, x1, y1, head=True):
        ax.annotate("", xy=(x1, y1), xytext=(x0, y0),
                    arrowprops=dict(arrowstyle="-|>" if head else "-", color=EDGE, lw=1.2, shrinkA=0, shrinkB=0,
                                    mutation_scale=9))

    cx, bw = W / 2, 3.3
    bx = cx - bw / 2
    y = box(bx, 0, bw, "FDA openFDA Cosmetic Adverse Event Reports",
            ["FDA industry code 53 (Cosmetics)", f"N = {c['all']:,}"])
    arrow(cx, y, cx, y - 0.24)
    y = box(bx, y - 0.24, bw, "Product-Name Exposure Screening",
            ["Relaxer/Alkali: relaxer, lye, perm", "Keratin/Blowout: keratin, blowout",
             "Whole-word match in the product brand name"])
    arrow(cx, y, cx, y - 0.24)
    y = box(bx, y - 0.24, bw, "Primary Keyword-Defined Cohort",
            [f"N = {c['N']}", f"Relaxer/Alkali: n = {c['R']}", f"Keratin/Blowout: n = {c['K']}",
             f"Reports created {c['first']:%B %Y} to {c['last']:%B %Y}"])

    # possible-duplicate screen (sensitivity analysis), CONSORT-style side box
    d = c["dup"]
    stem_top = y
    side_top = y - 0.18
    side_x = cx + 0.4
    side_bottom = box(side_x, side_top, W - 0.06 - side_x, None,
                      ["**Possible-duplicate screen (sensitivity analysis)",
                       f"{d['clusters']} clusters comprising {d['reports']} reports",
                       f"{d['removed']} reports removed (R/A {d['R']}, K/B {d['K']}); Table S5"], dashed=True)
    ax.plot([cx, side_x], [(side_top + side_bottom) / 2] * 2, color=EDGE, lw=1.0, ls=(0, (3, 2)))
    lab_y = side_bottom - 0.24
    arrow(cx, stem_top, cx, lab_y + 0.02)
    ax.text(cx, lab_y - 0.03, "Alternative exposure definitions applied\nindependently to the primary cohort",
            ha="center", va="top", fontsize=8, fontweight="bold", color=INK_MAIN, linespacing=1.3)

    # SA1-SA4
    gap, n_sa = 0.12, 4
    sw = (W - 0.1 - gap * (n_sa - 1)) / n_sa
    xs = [0.05 + i * (sw + gap) for i in range(n_sa)]
    bus_y = lab_y - 0.44
    ax.plot([xs[0] + sw / 2, xs[-1] + sw / 2], [bus_y, bus_y], color=EDGE, lw=1.2)
    ax.plot([cx, cx], [lab_y - 0.4, bus_y], color=EDGE, lw=1.2)
    desc = {"SA1": ["Single", "Cosmetic", "Product"], "SA2": ["Single", "Suspect", "Product"],
            "SA3": ["Conservative", "Product-Name", "Validation"],
            "SA4": ["Conservative", "Validation + Single", "Suspect Product"]}
    sa_top = bus_y - 0.24
    bottoms = []
    for x, tag in zip(xs, ["SA1", "SA2", "SA3", "SA4"]):
        arrow(x + sw / 2, bus_y, x + sw / 2, sa_top)
        r_, k_ = c[tag]
        bottoms.append(box(x, sa_top, sw, tag, desc[tag] + ["---", f"R/A: n = {r_}", f"K/B: n = {k_}", "---",
                                                             f"N = {r_ + k_}"], align="center"))

    # SA3 exclusions by reason
    reasons = c["sa3_reasons"]
    def entry(lab, r_, k_):
        return f"{lab}: " + ", ".join(f"{a} {n}" for a, n in (("R/A", r_), ("K/B", k_)) if n)
    items = [entry(*t) for t in reasons]
    ex_r, ex_k = c["sa3_excluded"]
    items.append(f"**Total excluded: R/A {ex_r}, K/B {ex_k}")
    half = math.ceil(len(items) / 2)
    left, right = items[:half], items[half:]
    ex_top = min(bottoms) - 0.24
    sa3_x = xs[2] + sw / 2
    ax.plot([sa3_x, sa3_x], [bottoms[2], ex_top], color=EDGE, lw=1.0, ls=(0, (3, 2)))
    ex_bottom = box(0.05, ex_top, W - 0.1, f"Reports Excluded From SA3 by Primary Reason (Table S3b)", [""] * half)
    row0 = ex_top - HEAD_H - 0.12
    for col_x, col in ((0.2, left), (W / 2 + 0.05, right)):
        for i, ln in enumerate(col):
            ax.text(col_x, row0 - (i + 0.5) * LINE_H, ln.strip("*"), ha="left", va="center", fontsize=7.6,
                    fontweight="bold" if ln.startswith("**") else "normal", color=INK_MAIN)

    ax.set_xlim(0, W)
    ax.set_ylim(ex_bottom - 0.05, 0.05)
    fig.set_size_inches(W, 0.1 - ex_bottom)
    _save_main(fig, base)
    plt.close(fig)


# ---- Figure 2: forest plot of the primary signals -----------------------------------------------

def _figure2(key, base):
    plt = _mpl(9)
    import matplotlib
    from matplotlib.lines import Line2D
    from matplotlib.legend_handler import HandlerBase

    P = key["primary"].set_index("reaction")
    S = [v.set_index("reaction") for v in key["sa_res"].values()]
    rows = []
    for t in key["sig10"]:
        r = P.loc[t]
        b = r["boundary"]
        rows.append(dict(
            t=t, ror=r["ror"], lo=r["lo"], hi=r["hi"], p=r["p"],
            robust=all(bool(s.loc[t, "fdr_sig"]) for s in S),
            rel=r["direction"].startswith("Higher in relaxer"),
            zero="lower" if b.startswith("lower") else ("upper" if b.startswith("upper") else ""),
            label=(f"NE (0 to {r['hi']:.2f})" if b.startswith("lower") else
                   f"NE ({_b(r['lo'])} to ∞)" if b.startswith("upper") else
                   f"{r['ror']:.2f} ({_b(r['lo'])}-{r['hi']:.2f})")))
    # robust signals first; within each group relaxer/alkali-higher first, then by P
    rows.sort(key=lambda d: (not d["robust"], not d["rel"], d["p"], d["t"]))
    rows = rows[::-1]
    n = len(rows)

    est = [r for r in rows if not r["zero"]]
    vmin = min([r["lo"] for r in est] + [r["hi"] for r in rows if r["zero"] == "lower"]) * 0.6
    vmax = max([r["hi"] for r in est] + [r["lo"] for r in rows if r["zero"] == "upper"]) * 1.6
    ROW, TOP, BOT, LEFT, AXW = 0.42, 0.75, 2.1, 1.65, 2.95
    H = TOP + ROW * n + BOT
    fig = plt.figure(figsize=(MAIN_WIDTH_IN, H))
    ax = fig.add_axes([LEFT / MAIN_WIDTH_IN, BOT / H, AXW / MAIN_WIDTH_IN, ROW * n / H])
    for y, r in enumerate(rows):
        if r["zero"] == "lower":          # no point estimate: bar from the axis to the exact upper limit
            ax.annotate("", xy=(vmin, y), xytext=(r["hi"], y), annotation_clip=False,
                        arrowprops=dict(arrowstyle="-|>", color=INK_MAIN, lw=1.3, shrinkA=0, shrinkB=0, mutation_scale=9))
            ax.plot([r["hi"]] * 2, [y - 0.13, y + 0.13], color=INK_MAIN, lw=1.3)
        elif r["zero"] == "upper":
            ax.annotate("", xy=(vmax, y), xytext=(r["lo"], y), annotation_clip=False,
                        arrowprops=dict(arrowstyle="-|>", color=INK_MAIN, lw=1.3, shrinkA=0, shrinkB=0, mutation_scale=9))
            ax.plot([r["lo"]] * 2, [y - 0.13, y + 0.13], color=INK_MAIN, lw=1.3)
        else:
            ax.plot([r["lo"], r["hi"]], [y, y], color=INK_MAIN, lw=1.3, zorder=2)
            ax.plot(r["ror"], y, marker="o", ms=6.5, mec=INK_MAIN, mew=1.4,
                    mfc=INK_MAIN if r["robust"] else "white", zorder=3)
        ax.text(1.04, y, r["label"], transform=ax.get_yaxis_transform(), va="center", fontsize=7.6, color="#222222")
    ax.text(1.04, n - 0.35, "ROR (95% CI)", transform=ax.get_yaxis_transform(), va="bottom", fontsize=7.6,
            fontweight="bold", color="#222222")
    ax.axvline(1, color=REF_LINE, lw=1, ls="--", zorder=1)
    ax.set_xscale("log")
    ax.set_xlim(vmin, vmax)
    ticks = _log_ticks(vmin, vmax)
    ax.set_xticks(ticks)
    ax.set_xticklabels([_tick_label(t) for t in ticks])
    ax.xaxis.set_minor_locator(matplotlib.ticker.NullLocator())
    ax.set_ylim(-0.6, n - 0.4)
    ax.set_yticks(range(n))
    ax.set_yticklabels([tdisp(r["t"]) for r in rows])
    for tick, r in zip(ax.get_yticklabels(), rows):
        tick.set_fontweight("bold" if r["robust"] else "normal")
    ax.tick_params(axis="y", length=0)
    ax.set_xlabel("Reporting Odds Ratio (log scale)", fontsize=9.5)
    ax.set_title("Reporting Odds Ratios for Primary-Significant Adverse Events", fontsize=10, fontweight="bold",
                 loc="left", x=-LEFT / AXW + 0.02, pad=12)
    for xa, ha, txt in ((0, "left", "← Higher reporting odds,\nkeratin/blowout products"),
                        (1, "right", "Higher reporting odds,\nrelaxer/alkali products →")):
        ax.annotate(txt, (xa, 0), xycoords="axes fraction", xytext=(0, -40), textcoords="offset points",
                    ha=ha, va="top", fontsize=7.5)

    class _NEHandler(HandlerBase):
        def create_artists(self, legend, orig, xd, yd, w, h, fs, trans):
            y = yd + h / 2
            return [Line2D([xd + w * 0.2, xd + w], [y, y], color=INK_MAIN, lw=1.3, transform=trans),
                    Line2D([xd + w] * 2, [y - h * 0.35, y + h * 0.35], color=INK_MAIN, lw=1.3, transform=trans),
                    Line2D([xd + w * 0.16], [y], marker="<", ms=5, color=INK_MAIN, lw=0, transform=trans)]
    ne_proxy = Line2D([], [])
    handles = [Line2D([0], [0], color=INK_MAIN, marker="o", mfc=INK_MAIN, lw=1.3, ms=6.5),
               Line2D([0], [0], color=INK_MAIN, marker="o", mfc="white", mew=1.4, lw=1.3, ms=6.5)]
    labels = ["Significant in the primary analysis and all four exposure-definition sensitivity analyses (bold label)",
              "Significant in the primary analysis; not in all four sensitivity analyses"]
    if any(r["zero"] for r in rows):
        handles.append(ne_proxy)
        labels.append("No relaxer/alkali reports: ROR not estimable (NE); bar spans 0 to the exact upper 95% CI limit")
    fig.legend(handles, labels, handler_map={ne_proxy: _NEHandler()}, loc="upper center",
               bbox_to_anchor=(0.5, (BOT - 1.05) / H), fontsize=7.2, frameon=False, ncol=1, handlelength=2.6)
    fig.text(0.5, 0.12 / H, "Dashed line, ROR = 1. ROR, reporting odds ratio; CI, confidence interval; NE, not estimable.",
             ha="center", va="bottom", fontsize=6.6, color="#555555")
    _save_main(fig, base)
    plt.close(fig)


# ---- Figure 3: robustness heatmap with q (or P) values -------------------------------------------

def _figure3(key, base):
    plt = _mpl(8)
    import matplotlib.transforms as mtrans
    from matplotlib.patches import Rectangle, Patch
    m = figure3_matrix(key)
    cols, groups = m["cols"], m["groups"]
    GAP, CW, RH, HEADER_ROW, GROUP_GAP = 0.4, 0.47, 0.29, 0.85, 0.25
    xs, seen_p = [], False
    for i, c in enumerate(cols):
        seen_p = seen_p or c["kind"] == "P"
        xs.append(i + (GAP if seen_p else 0))
    ylay, y = [], 0.0
    for g, ts in groups:
        ylay.append(("group", g, y))
        y += HEADER_ROW
        for t in ts:
            ylay.append(("row", t, y + 0.5))
            y += 1
        y += GROUP_GAP
    y_total = y - GROUP_GAP
    width_units = xs[-1] + 1
    LEFT, TOP, BOT = 1.55, 1.35, 0.55
    H = TOP + y_total * RH + BOT
    fig = plt.figure(figsize=(MAIN_WIDTH_IN, H))
    ax = fig.add_axes([LEFT / MAIN_WIDTH_IN, BOT / H, width_units * CW / MAIN_WIDTH_IN, y_total * RH / H])
    ax.set_xlim(-0.5, width_units - 0.5)
    ax.set_ylim(y_total, 0)
    ax.axis("off")

    for kind, lab, yy in ylay:
        if kind == "group":
            ax.text(-0.5, yy + HEADER_ROW * 0.55, lab, ha="left", va="center", fontsize=7.4, fontstyle="italic",
                    color="#333333")
            continue
        ax.text(-0.62, yy, tdisp(lab), ha="right", va="center", fontsize=8, color=INK_MAIN)
        for x, c in zip(xs, cols):
            cell = c["cells"].get(lab)
            if cell is None:          # analysis not run for this reaction
                ax.add_patch(Rectangle((x - 0.46, yy - 0.42), 0.92, 0.84, fc="white", ec="#C8C8C8", lw=0.6))
                ax.text(x, yy, "–", ha="center", va="center", fontsize=7, color="#777777")
                continue
            ax.add_patch(Rectangle((x - 0.5, yy - 0.5), 1, 1, fc=SIG_FILL if cell["sig"] else NS_FILL,
                                   ec="white", lw=1.6))
            ax.text(x, yy, fmt_p(cell["v"]), ha="center", va="center", fontsize=7,
                    color="white" if cell["sig"] else INK_MAIN, fontweight="bold" if cell["sig"] else "normal")

    # column headers (anchored to the top of the grid in points, so spacing does not depend on the data)
    for x, c in zip(xs, cols):
        ax.annotate(c["N"], (x, 0), xytext=(0, 3), textcoords="offset points", ha="center", va="bottom",
                    fontsize=6.5, color="#666666", annotation_clip=False)
        ax.annotate(c["desc"], (x, 0), xytext=(0, 30), textcoords="offset points", ha="center", va="top",
                    fontsize=6.5, color=INK_MAIN, linespacing=1.15, annotation_clip=False)
    i = 0
    while i < len(cols):                      # bold names; adjacent columns with the same name share one label
        j = i
        while j + 1 < len(cols) and cols[j + 1]["name"] == cols[i]["name"]:
            j += 1
        ax.annotate(cols[i]["name"], ((xs[i] + xs[j]) / 2, 0), xytext=(0, 32), textcoords="offset points",
                    ha="center", va="bottom", fontsize=7, fontweight="bold", color=INK_MAIN, annotation_clip=False)
        i = j + 1
    for kind_, lab in (("q", "FDR-adjusted q value"), ("P", "Unadjusted P value")):
        xx = [x for x, c in zip(xs, cols) if c["kind"] == kind_]
        if not xx:
            continue
        x0, x1 = min(xx) - 0.45, max(xx) + 0.45
        lift = ax.transData + mtrans.ScaledTranslation(0, 45 / 72, fig.dpi_scale_trans)
        ax.plot([x0, x1], [0, 0], color="#555555", lw=0.8, transform=lift, clip_on=False)
        ax.annotate(lab, ((x0 + x1) / 2, 0), xytext=(0, 47), textcoords="offset points", ha="center",
                    va="bottom", fontsize=7.4, fontweight="bold", color="#333333", annotation_clip=False)
    ax.set_title("Robustness of Adverse Event Signals Across Sensitivity Analyses", fontsize=9.5,
                 fontweight="bold", pad=66)

    handles = [Patch(fc=SIG_FILL, ec="none"), Patch(fc=NS_FILL, ec="none"), Patch(fc="white", ec="#C8C8C8", lw=0.6)]
    labels = ["Significant (q<.05; P<.05 in the last two columns)", "Not significant",
              "Not assessed (–)"]
    fig.legend(handles, labels, loc="lower center", bbox_to_anchor=(0.5, 0.02 / H), ncol=3, fontsize=7,
               frameon=False, handlelength=1.4, handleheight=1.0, columnspacing=1.6)
    _save_main(fig, base)
    plt.close(fig)


# ---- Figure 4: temporal patterns ---------------------------------------------------------------

def _figure4(key, base):
    plt = _mpl(9)
    import matplotlib
    from matplotlib.transforms import blended_transform_factory as blend
    ann, pp = key["annual"], key["period_prop"]
    # vertical layout in inches, top to bottom: (space above panel, panel height, space below panel)
    PLOT_H, LEFT, RIGHT = 1.9, 1.3, 0.12
    layout = [(0.5, PLOT_H, 0.5), (0.42, PLOT_H, 0.82), (0.42, PLOT_H, 0.62)]
    H = sum(sum(t) for t in layout)
    fig = plt.figure(figsize=(MAIN_WIDTH_IN, H))
    axes, y = [], H
    for above, h, below in layout:
        y -= above + h
        axes.append(fig.add_axes([LEFT / MAIN_WIDTH_IN, y / H, (MAIN_WIDTH_IN - LEFT - RIGHT) / MAIN_WIDTH_IN, h / H]))
        y -= below
    axA, axB, axC = axes
    marker = {RELAXER: "o", KERATIN: "s"}

    # A. annual counts
    for coh in (RELAXER, KERATIN):
        axA.plot(ann.Year, ann[f"{coh} reports"], color=COLOR[coh], marker=marker[coh], ms=3.4, lw=1.3, label=coh)
    spans = [tuple(int(v) for v in p.split("-")) for p in FIG4_PERIODS]
    for a, _ in spans[1:]:
        axA.axvline(a - 0.5, color=REF_LINE, lw=0.8, ls=":", zorder=0)
    for (a, b), lab in zip(spans, FIG4_PERIODS):
        axA.text((a + b) / 2, 1.01, lab, transform=blend(axA.transData, axA.transAxes), ha="center", va="bottom",
                 fontsize=7, color="#555555")
    axA.set_xlim(spans[0][0] - 0.6, spans[-1][1] + 0.6)
    axA.set_ylim(0, None)
    axA.xaxis.set_major_locator(matplotlib.ticker.MultipleLocator(2))
    axA.set_ylabel("Reports, No.", fontsize=8.5)
    axA.set_xlabel("Year", fontsize=8.5)
    axA.legend(fontsize=7.2, frameon=False, loc="upper right")
    axA.set_title("A. Annual Report Counts by Cohort", fontsize=9, fontweight="bold", loc="left", pad=15)
    for ax in (axA,):
        ax.tick_params(axis="x", labelsize=8)

    def period_panel(ax, series, ylabel, title, n_rows):
        x = np.arange(len(FIG4_PERIODS))
        top = 0
        for lab, coh, rx_, col, mk, off in series:
            g = pp[(pp.Reaction == rx_) & (pp.Cohort == coh)].set_index("Period").reindex(FIG4_PERIODS)
            p = 100 * g.n / g.N
            lo, hi = 100 * g._lo, 100 * g._hi
            eb = ax.errorbar(x + off, p, yerr=[p - lo, hi - p], fmt=mk, color=col, ms=5, lw=1.2, capsize=2.5,
                             capthick=1.2, label=lab)
            for art in eb.get_children():
                art.set_clip_on(False)          # a 0% marker sits on the axis instead of being cut in half
            top = max(top, hi.max())
        ax.set_xticks(x)
        ax.set_xticklabels(FIG4_PERIODS, fontsize=8)
        ax.set_xlim(-0.5, len(x) - 0.5)
        ax.set_ylim(0, min(100, 10 * math.ceil(top / 10)))
        ax.set_ylabel(ylabel, fontsize=8.5)
        ax.set_title(title, fontsize=9, fontweight="bold", loc="left")
        ax.tick_params(axis="x", length=0, pad=5)
        tr = blend(ax.transData, ax.transAxes)
        # number-of-reports table under the axis, positioned in points so it does not depend on panel height
        def put(xd, txt, dy, **kw):
            ax.annotate(txt, (xd, 0), xycoords=tr, xytext=(0, dy), textcoords="offset points", va="center",
                        fontsize=7, annotation_clip=False, **kw)
        put(-0.5, "No. of reports", -27, ha="right", fontstyle="italic", color="#555555")
        for j, (lab, coh) in enumerate(n_rows):
            dy = -27 - 11 * (j + 1)
            put(-0.5, lab, dy, ha="right", color=INK_MAIN)
            Ns = pp[(pp.Cohort == coh)].drop_duplicates("Period").set_index("Period").reindex(FIG4_PERIODS).N
            for xi, N in zip(x, Ns):
                put(xi, f"{int(N)}", dy, ha="center", color=INK_MAIN)

    period_panel(axB, [(RELAXER, RELAXER, "Alopecia", COLOR[RELAXER], "o", -0.09),
                       (KERATIN, KERATIN, "Alopecia", COLOR[KERATIN], "s", 0.09)],
                 "Reports Listing\nAlopecia, %", "B. Reports Listing Alopecia by Cohort and Period",
                 [(RELAXER, RELAXER), (KERATIN, KERATIN)])
    axB.legend(fontsize=7.2, frameon=False, loc="upper right")
    offs = np.linspace(-0.24, 0.24, len(FIG4C_REACTIONS))
    period_panel(axC, [(rx_, KERATIN, rx_, col, mk, off) for rx_, col, mk, off in
                       zip(FIG4C_REACTIONS, RX_PALETTE, RX_MARKERS, offs)],
                 "Keratin/Blowout Reports\nListing Reaction, %",
                 "C. Other Keratin/Blowout-Associated Reactions by Period", [(KERATIN, KERATIN)])
    axC.legend(fontsize=7.2, frameon=False, loc="upper right", ncol=2, columnspacing=1.2)
    _save_main(fig, base)
    plt.close(fig)


def make_main_figures(key, out_dir):
    """Manuscript Figures 1-4. Returns {figure number: path to the PNG}."""
    specs = [("1", "Figure1_study_flow", _figure1), ("2", "Figure2_forest_plot", _figure2),
             ("3", "Figure3_robustness_heatmap", _figure3), ("4", "Figure4_temporal_patterns", _figure4)]
    paths = {}
    for num, name, fn in specs:
        base = os.path.join(out_dir, name)
        fn(key, base)
        paths[num] = base + ".png"
    return paths


# ======================================================================================
# 8. WORD SUPPORTING INFORMATION
# ======================================================================================

def _num_word(n):
    """Numbers one to nine as words (journal style for prose); larger numbers as numerals."""
    return ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"][n] if 0 <= n <= 9 else str(n)


def word_spec(key, sheets):
    """Titles, column headers, rows, and footnotes for every Word table."""
    T = {k: v.copy() for k, v in sheets.items()}
    df = key["df"]
    R, K = RELAXER, KERATIN
    nR, nK = int(df.relaxer.sum()), int((df.relaxer == 0).sum())
    n_terms = len(key["terms"])

    def cell(x):
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return ""
        if isinstance(x, (float, np.floating)) and float(x).is_integer():
            x = int(x)
        s = str(x)
        return (s.replace(">=", "≥").replace("<=", "≤").replace("\u00a0", " ")
                 .replace("Life Threatening", "Life-threatening").replace("Visited Emergency Room", "Emergency department visit"))

    def tab(frame, cols=None, rename=None):
        f = frame if cols is None else frame[cols]
        if rename:
            f = f.rename(columns=rename)
        return {"columns": [cell(c) for c in f.columns], "rows": [[cell(v) for v in r] for r in f.itertuples(index=False)]}

    blocks = []

    def add(num, title, frame_spec, notes, weights=None, font=8):
        frame_spec.update(id=num, title=title, notes=notes, weights=weights, font=font)
        blocks.append(frame_spec)

    full_cols = ["Reaction (MedDRA PT)", f"{R}, n/N (%)", f"{K}, n/N (%)", "ROR", "95% CI",
                 "Exact conditional 95% CI", "P (Fisher)", "q (BH-FDR)", "FDR significant"]
    ren_full = {"Reaction (MedDRA PT)": "Reaction (MedDRA preferred term)", "FDR significant": "q<.05"}
    common_full = (f"Reactions are ordered by P value. The tested family comprises the {n_terms} reactions reported in ≥5 reports "
                   f"in either cohort of the primary analysis; the same {n_terms} reactions and Benjamini-Hochberg correction are used "
                   "in every analysis. ROR >1 indicates relatively higher reporting among relaxer/alkali reports; ROR <1, among "
                   "keratin/blowout reports. 95% CI is the Woolf (log) interval when all cells are non-zero and the exact conditional "
                   "interval otherwise; the exact conditional interval is also shown for every reaction. NE*, not estimable because "
                   "no relaxer/alkali report listed the reaction (conditional maximum-likelihood estimate on the lower boundary, 0); "
                   "the exact 95% upper confidence limit is shown.")
    W_FULL = [2.3, 1.45, 1.55, 0.7, 1.15, 1.35, 0.85, 0.85, 0.7]
    add("S1", f"Primary disproportionality analysis: all {n_terms} tested reactions (relaxer/alkali, n = {nR}; keratin/blowout, n = {nK})",
        tab(T["TableS1_Primary_all"], full_cols, ren_full), [common_full], W_FULL)
    sa3 = df[df["SA3 keep"] == "Yes"]
    add("S2", f"Co-primary analysis (SA3, product-name-validated cohorts): all {n_terms} tested reactions "
              f"(relaxer/alkali, n = {int(sa3.relaxer.sum())}; keratin/blowout, n = {int((sa3.relaxer == 0).sum())})",
        tab(T["TableS2_SA3_all"], full_cols, ren_full),
        ["SA3 retains reports whose suspect cosmetic product was clearly identifiable by name as a chemical hair-straightening "
         "product (relaxer, or keratin/blowout smoothing treatment); criteria and exclusions are given in Tables S3a-S3c. " + common_full], W_FULL)

    add("S3a", "Exposure classification criteria and cohort sizes", tab(T["TableS3a_Classification_criteria"]),
        ["Primary cohorts were defined by case-insensitive keyword matching of the product brand-name field. SA, sensitivity analysis. "
         f"Report-level classification of all {len(df)} reports (keyword-matching product, its role, harmonized brand, and SA1-SA4 "
         "membership) is provided as Table S3d in the Excel file (Data S2; sheet TableS3d_Report_level_classif). Product names were "
         "classified independently by two authors, with disagreements resolved by consensus."], [1.6, 7.2, 2.2])
    add("S3b", "Primary reason for exclusion from SA3", tab(T["TableS3b_SA3_exclusion_reasons"]),
        ["One primary reason is listed per excluded report."], [6, 1.5, 1.5, 1])
    add("S3c", "Representative product names by SA3 status", tab(T["TableS3c_Representative_names"]),
        ["Product names are shown as reported (title case, spelling unchanged), with the number of reports in parentheses when > 1. "
         "Up to six names are shown for included products and up to four for each exclusion reason."], [1.4, 3.6, 7])

    def extra_note(frame, what):
        names = [str(v).lower() for v in frame.iloc[len(key["sig10"]):, 0]]
        if not names:
            return ""
        one = len(names) == 1
        lead = "The last row shows" if one else f"The last {_num_word(len(names))} rows show"
        listed = names[0] if one else ", ".join(names[:-1]) + (", and " if len(names) > 2 else " and ") + names[-1]
        return (f" {lead} {listed}, which {'was' if one else 'were'} not significant in the primary analysis but "
                f"reached q<.05 in at least one {what}.")
    t4 = T["TableS4_Estimates_SA1_SA4"]
    add("S4", "Effect estimates for the 10 primary signals, and for other reactions significant under any definition, across "
              "exposure definitions", tab(t4),
        ["Cells show ROR (95% CI); Benjamini-Hochberg q value. N in column headings, relaxer/alkali reports/keratin/blowout reports. "
         "SA1, exactly one cosmetic product, matching the cohort keywords; SA2, exactly one suspect product, a cosmetic matching the "
         "cohort keywords; SA3, product-name validation (Tables S3a-S3c); SA4, SA3 plus exactly one suspect product. The sensitivity "
         "cohorts are nested subsets of the primary cohorts and are not independent replications; they test robustness to exposure "
         "misclassification." + extra_note(t4, "exposure definition") + " NE*, as in Table S1."], [1.6, 2, 2, 2, 2, 2])

    s5 = T["TableS5_Possible_duplicates"].copy()
    s5["Reactions"] = s5["Reactions"].map(lambda x: "; ".join(disp(t) for t in str(x).split("; ")))
    s5["Retained"] = s5["Retained"].str.replace(" (removed in de-duplicated analysis)", "", regex=False)
    s5["Product"] = s5["Product"].str.title()
    add("S5", "Possible duplicate reports", tab(s5),
        [f"Possible duplicates were defined as reports in the same cohort with the same creation date, an identical set of "
         f"reaction terms, the same harmonized brand, and no conflicting non-missing age or sex. {_num_word(s5.Cluster.nunique()).capitalize()} clusters "
         f"comprising {_num_word(len(s5))} reports were identified; one report per cluster was retained (age recorded preferred, then "
         f"lowest report number), removing {_num_word(int((s5.Retained == 'No').sum()))} reports in the de-duplicated analysis "
         f"(Table S6). CAERS records "
         "carry no identifiers that would allow duplicates to be confirmed, so these are possible, not confirmed, duplicates."],
        [0.8, 1.1, 1.3, 1.1, 2.6, 0.8, 0.7, 4.4, 0.9])

    n_peak_k = int(((df.relaxer == 0) & df.year.isin(PEAK_YEARS)).sum())
    late = df[df.year >= 2013]
    t6 = T["TableS6_Additional_sensitivity"]
    n_2003 = int((df.year == 2003).sum())
    add("S6", "Additional sensitivity analyses: possible duplicates removed, calendar-period restrictions, and female reports only",
        tab(t6),
        [f"Cells show ROR (95% CI); Benjamini-Hochberg q value across the same {n_terms}-reaction family. N in column headings, "
         f"relaxer/alkali reports/keratin/blowout reports. Excluding 2010-2012 removes the peak in keratin/blowout reporting "
         f"({n_peak_k} of {nK} keratin/blowout reports). The 2013-2023 analysis retains {int(late.relaxer.sum())} relaxer/alkali and "
         f"{int((late.relaxer == 0).sum())} keratin/blowout reports and has limited power. The 2004-2023 analysis removes the "
         f"{_num_word(n_2003)} report{'s' if n_2003 != 1 else ''} created in 2003." + extra_note(t6, "of these analyses") +
         " NE*, as in Table S1."], [1.6, 2, 2, 2, 2, 2, 2])

    s7 = T["TableS7_Stratified_adjusted"].replace({"NE (zero events in one cohort)": "NE†", "Not estimable (0 events in one cohort)": "NE†"})
    nage_R = int(df[df.age.notna()].relaxer.sum())
    nage_K = int((df[df.age.notna()].relaxer == 0).sum())
    add("S7", "Calendar-period-stratified, age-stratified, and age- and sex-adjusted estimates for the 10 primary signals", tab(s7),
        ["Period strata: 2003-2009, 2010-2012, and 2013-2023. Age strata: <35, 35-49, and ≥50 years. Mantel-Haenszel (MH) odds ratios "
         "with Cochran-Mantel-Haenszel (CMH) tests (continuity corrected). Age-stratified and logistic analyses are restricted to reports "
         f"with age recorded ({nage_R} relaxer/alkali and {nage_K} keratin/blowout reports). Logistic regression models the reaction on cohort, "
         "age (continuous, years), and female sex (vs male or not recorded). NE†, not estimable because no relaxer/alkali report with "
         "the reaction was available; the CMH test is still reported. NE*, as in Table S1."],
        [1.5, 1.4, 1.5, 1.0, 1.6, 1.5, 1.0, 1.9])

    add("S8a", "Product (brand) concentration of the 10 primary signals",
        tab(T["TableS8a_Product_concentration"], rename={"Reports with unbranded/generic name": "Unbranded or generic name, n",
                                                         "Distinct named brands": "Distinct named brands, n"}),
        ["Counts refer to reports in the cohort with higher reporting of the reaction. Brands were harmonized from reported product "
         "names using rule-based mapping verified by the authors; a report can name more than one brand. 'Brazilian Blowout' may denote "
         "the brand of that name or be used generically for the procedure, so the two cannot be separated."],
        [1.4, 1.2, 0.9, 0.9, 1.0, 2.2, 0.9, 0.9, 1.1, 0.9])
    add("S8b", "Leave-top-brand-out analysis for the 10 primary signals", tab(T["TableS8b_Leave_top_brand_out"]),
        ["All reports naming the top contributing brand (Table S8a) were removed from the cohort with higher reporting and the ROR was "
         "re-estimated. P values are unadjusted Fisher exact tests. NE*, as in Table S1."],
        [1.5, 2.6, 1.1, 1.1, 0.7, 1.0, 1.1, 0.7, 1.0])
    bc = sheets["TableS8c_Brand_composition"]
    rows = []
    for coh in (K, R):
        g = bc[bc.Cohort == coh]
        for _, r in g[g.Reports >= 2].iterrows():
            rows.append([coh, r["Harmonized brand"], int(r.Reports)])
        small = g[g.Reports < 2]
        if len(small):
            rows.append([coh, f"Other names reported once ({len(small)} names)", int(small.Reports.sum())])
    add("S8c", "Brand composition of the primary cohorts", tab(pd.DataFrame(rows, columns=["Cohort", "Harmonized brand", "Reports"])),
        ["Reports naming more than one brand are counted under each. The full list of names reported once is provided in the "
         "Excel file (Data S2; sheet TableS8c_Brand_composition)."], [2, 6, 1.2])

    add("S9a", "Sensitivity of results to the reaction inclusion threshold", tab(T["TableS9a_Threshold_summary"]),
        ["Each rule defines a different testing family; Benjamini-Hochberg correction is applied within that family. 'Either cohort' "
         "requires the threshold to be met in at least one cohort; 'total' counts reports from both cohorts combined."], [3, 1.2, 1.2, 7])
    add("S9b", "Benjamini-Hochberg q values for the 10 primary signals under alternative inclusion thresholds",
        tab(T["TableS9b_Threshold_q_values"]), ["Column headings give the inclusion rule."], [2, 1.6, 1.6, 1.6, 1.6, 1.6])

    co_k = T["TableS10a_Cooccur_keratin"]
    add("S10a", f"Co-occurrence of the keratin/blowout-associated signals within keratin/blowout reports (N = {nK})",
        tab(co_k, rename={co_k.columns[0]: "Reaction"}),
        ["Diagonal cells give the number of reports listing the reaction; off-diagonal cells give the number listing both reactions."],
        [2] + [1] * (len(co_k.columns) - 1))
    co_r = T["TableS10b_Cooccur_relaxer"]
    add("S10b", f"Co-occurrence of the relaxer/alkali-associated signals and headache within relaxer/alkali reports (N = {nR})",
        tab(co_r, rename={co_r.columns[0]: "Reaction"}),
        ["Diagonal cells give the number of reports listing the reaction; off-diagonal cells give the number listing both reactions."],
        [2] + [1] * (len(co_r.columns) - 1))
    add("S10c", "Number of the 10 primary signal reactions listed per report", tab(T["TableS10c_Signals_per_report"]),
        ["Values are numbers of reports."], [3, 1.3, 1.3])

    add("S11a", "Symptom cluster definitions", tab(T["TableS11a_Cluster_definitions"], rename={"Terms": "No. of terms"}),
        ["Clusters were assigned from MedDRA preferred terms observed in the data. A report counts toward a cluster if it lists at least "
         "one of the cluster's terms. Non-specific terms without a stated body site (for example, burning sensation, pain) and "
         "dermatitis-type terms were not assigned to a cluster; Table S11c shows results when borderline terms are added or removed."],
        [2.3, 0.8, 9])
    add("S11b", "Cluster-level disproportionality in the primary analysis and in the exposure-definition, temporal, and "
                "product-concentration sensitivity analyses", tab(T["TableS11b_Cluster_results"]),
        ["A report counts toward a cluster if it lists at least one of the cluster's terms (Table S11a). 'Brazilian Blowout reports "
         "removed' excludes every keratin/blowout report naming Brazilian Blowout (as a brand or as the procedure). P values are "
         "unadjusted Fisher exact tests. Symptom clusters were defined post hoc and are a secondary analysis. NE*, as in Table S1."], [2.6, 2.6, 1.4, 1.4, 0.7, 1.0, 0.8])
    add("S11c", "Sensitivity of cluster results to alternative term lists (primary cohorts)", tab(T["TableS11c_Cluster_term_sensitivity"]),
        ["P values are unadjusted Fisher exact tests. The last two rows show terms not assigned to any cluster, for reference."],
        [2.2, 4, 1.4, 1.4, 0.7, 1.1, 0.9])
    add("S11d", "Calendar-period-stratified cluster-level estimates (primary cohorts)", tab(T["TableS11d_Cluster_period_stratified"]),
        ["Period strata: 2003-2009, 2010-2012, and 2013-2023. Mantel-Haenszel (MH) odds ratios with Cochran-Mantel-Haenszel (CMH) "
         "tests (continuity corrected). Symptom clusters were defined post hoc and are a secondary analysis."], [3, 2.2, 2.4, 1.4])

    s12 = T["TableS12_Serious_outcomes"].copy()
    s12["Product(s)"] = s12["Product(s)"].str.title()
    s12 = s12.drop(columns=["Harmonized brand"]).rename(columns={
        "Terms unlikely to reflect acute product exposure (author review)": "Terms unlikely to reflect acute product exposure†"})
    add("S12", "Reports recording death, a life-threatening event, or disability", tab(s12),
        ["Outcomes and reaction terms are as recorded by the reporter and were not adjudicated; CAERS reports do not establish that the "
         "product caused the event. †Terms identified on author review as chronic, neoplastic, obstetric, or otherwise unlikely to "
         "reflect an acute effect of product exposure."], [1.1, 1.1, 1.0, 2.2, 1.4, 5.2, 1.6], font=7.5)
    s12b = T["TableS12b_Serious_summary"]
    s12b["Most frequent brands"] = s12b["Most frequent brands"].str.replace("Other: ", "", regex=False)
    add("S12b", "Serious and health care outcomes by cohort, year, and brand", tab(s12b),
        ["Values are numbers of reports; a report can record more than one outcome. Brand counts are among reports with the outcome."],
        [1.6, 1.0, 1.0, 4.2, 4.2])

    a = T["TableS13a_Annual_counts"].copy()
    short = {"Relaxer/alkali reports": "All reports, R", "Keratin/blowout reports": "All reports, K"}
    for c in a.columns[3:]:
        rx_, coh = c.rsplit(", ", 1)
        short[c] = f"{rx_}, {'R' if coh.startswith('Relaxer') else 'K'}"
    a = a.rename(columns=short)
    a = pd.concat([a, pd.DataFrame([{c: ("Total" if c == "Year" else int(a[c].sum())) for c in a.columns}])], ignore_index=True)
    add("S13a", "Annual numbers of reports and of reports listing selected reactions (denominators for Fig. 4)", tab(a),
        ["R, relaxer/alkali; K, keratin/blowout. Year is the year the report was created in CAERS (date_created)."],
        [0.8] + [1] * (len(a.columns) - 1), font=7.5)
    pp = sheets["TableS13b_Period_proportions"].copy()
    pp["val"] = pp.apply(lambda r: f"{r.n}/{r.N} ({r['%']}; {r['Wilson 95% CI (%)']})", axis=1)
    wide = pp.pivot_table(index=["Period", "Reaction"], columns="Cohort", values="val", aggfunc="first", sort=False).reset_index()
    wide = wide[["Period", "Reaction", R, K]].rename(columns={R: f"{R}, n/N (%; 95% CI)", K: f"{K}, n/N (%; 95% CI)"})
    n_early_k = int(((df.relaxer == 0) & (df.year <= 2009)).sum())
    add("S13b", "Proportion of reports listing selected reactions by calendar period", tab(wide),
        ["Proportions are reporting proportions within each cohort and period, not incidence. 95% CIs are Wilson score "
         f"intervals. The 2003-2009 keratin/blowout estimate rests on {_num_word(n_early_k)} reports."], [1.2, 1.8, 2.6, 2.6])
    add("S14", "Study-period and record checks", tab(T["TableS14_Study_period_checks"]),
        ["The Modernization of Cosmetics Regulation Act of 2022 (MoCRA) adverse event reporting requirements took effect on "
         "December 29, 2023."], [4, 2, 3])
    add("S15", "openFDA application programming interface (API) queries", tab(T["TableS15_openFDA_queries"]),
        ["Base URL: https://api.fda.gov/food/event.json. Brand-name searches match whole words case-insensitively. The analysis "
         "used the data version last updated July 7, 2026 (extracted August 7, 2026). Re-running queries 1 and 2 on the version "
         "last updated November 6, 2026, returned the same 352 report numbers with identical reaction terms. Count queries (5 and 6) "
         "tabulate reaction terms case-sensitively and were used only as a check; the analysis merged case variants within each "
         "report. Code that re-runs these queries is provided in the code repository."], [0.6, 3.2, 6.6, 1.8])

    d0, d1 = df.date.min(), df.date.max()
    conventions = [
        f"Data source. US Food and Drug Administration (FDA) CAERS reports involving cosmetic products (industry code 53), obtained "
        f"through the openFDA food/event endpoint (data version last updated July 7, 2026; extracted August 7, 2026) using whole-word "
        f"searches of the product brand-name field (Table S15). The primary cohorts comprise {nR} relaxer/alkali and {nK} "
        f"keratin/blowout reports created between {d0:%B} {d0.day}, {d0.year}, and {d1:%B} {d1.day}, {d1.year}.",
        "Unit of analysis. The report. Each distinct MedDRA preferred term (case-normalized) is counted once per report; missing age "
        "and sex were not imputed.",
        "Reporting odds ratio (ROR). Odds of a reaction among relaxer/alkali reports divided by the odds among keratin/blowout reports. "
        "ROR >1 indicates relatively higher reporting with relaxer/alkali products and ROR <1 relatively higher reporting with "
        "keratin/blowout products. RORs describe relative reporting, not incidence or risk.",
        "Confidence intervals. Woolf (log) 95% CIs when all cells are non-zero; exact conditional 95% CIs otherwise. NE*, not "
        "estimable because no relaxer/alkali report listed the reaction; the exact 95% upper confidence limit is shown.",
        f"Testing. Two-sided Fisher exact tests with Benjamini-Hochberg false discovery rate (FDR) correction across the same "
        f"{n_terms}-reaction family in every analysis that retests all reactions; q<.05 was considered significant. Stratified, "
        "adjusted, brand-removal, and symptom-cluster comparisons, which were limited to the primary signals or the five clusters, "
        "used unadjusted P<.05.",
        "Abbreviations. BH, Benjamini-Hochberg; CAERS, Center for Food Safety and Applied Nutrition (CFSAN) Adverse Event Reporting "
        "System; CI, confidence interval; CMH, Cochran-Mantel-Haenszel; FDR, false discovery rate; MedDRA, Medical Dictionary for "
        "Regulatory Activities; MH, Mantel-Haenszel; NE, not estimable; ROR, reporting odds ratio; SA, sensitivity analysis.",
        "Reproducibility. All tables and figures were generated from the report-level data set by the analysis code in the code "
        "repository (see the Data Availability Statement); the Excel file (Data S2) contains every table, including the "
        "report-level classification (Table S3d).",
    ]
    return blocks, conventions


def build_word(blocks, conventions, figures, out_path):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, Twips

    FONT = "Arial"
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.space_after = Pt(0)
    for z in doc.settings.element.findall(qn("w:zoom")):
        z.set(qn("w:percent"), "100")

    sec = doc.sections[0]                                  # US Letter, landscape, 0.6" margins
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11), Inches(8.5)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(0.6))
    content_w = Inches(11 - 1.2)

    # footer with page number
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = fp.add_run("Supporting Information, page ")
    r.font.size = Pt(8)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    rr = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "16")
    rpr.append(sz)
    rr.append(rpr)
    t = OxmlElement("w:t")
    t.text = "1"
    rr.append(t)
    fld.append(rr)
    fp._p.append(fld)

    def para(text="", size=10, bold=False, italic=False, before=0, after=0, keep=False, page_break=False, align=None):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        pf.keep_with_next = keep
        pf.page_break_before = page_break
        if align:
            p.alignment = align
        if text:
            run = p.add_run(text)
            run.font.size, run.bold, run.italic = Pt(size), bold, italic
        return p

    def set_border(cell, **edges):
        tcPr = cell._tc.get_or_add_tcPr()
        b = OxmlElement("w:tcBorders")
        for edge in ("top", "left", "bottom", "right"):
            spec = edges.get(edge)
            el = OxmlElement(f"w:{edge}")
            if spec:
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(spec[0]))
                el.set(qn("w:color"), spec[1])
            else:
                el.set(qn("w:val"), "nil")
            b.append(el)
        tcPr.append(b)

    def shade(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        s = OxmlElement("w:shd")
        s.set(qn("w:val"), "clear")
        s.set(qn("w:color"), "auto")
        s.set(qn("w:fill"), fill)
        tcPr.append(s)

    def row_flags(row, header=False):
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        trPr.append(cs)
        if header:
            h = OxmlElement("w:tblHeader")
            trPr.append(h)

    def table(block):
        cols, rows = block["columns"], block["rows"]
        n = len(cols)
        w = block["weights"] if block["weights"] and len(block["weights"]) == n else [1] * n
        widths = [int(content_w * x / sum(w)) for x in w]
        size = Pt(block["font"])
        keep = len(rows) <= 25
        aligns = [WD_ALIGN_PARAGRAPH.LEFT] + [
            WD_ALIGN_PARAGRAPH.CENTER if max((len(r[j]) for r in rows), default=0) <= 26 else WD_ALIGN_PARAGRAPH.LEFT
            for j in range(1, n)]
        tb = doc.add_table(rows=1 + len(rows), cols=n)
        tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        tb.autofit = False
        tblPr = tb._tbl.tblPr
        lay = OxmlElement("w:tblLayout")
        lay.set(qn("w:type"), "fixed")
        mar = OxmlElement("w:tblCellMar")
        for side, v in (("top", 30), ("left", 70), ("bottom", 30), ("right", 70)):    # schema order
            e = OxmlElement(f"w:{side}")
            e.set(qn("w:w"), str(v))
            e.set(qn("w:type"), "dxa")
            mar.append(e)
        look = tblPr.find(qn("w:tblLook"))
        if tblPr.find(qn("w:tblLayout")) is None:    # tb.autofit = False has usually added it already
            (look.addprevious if look is not None else tblPr.append)(lay)
        (look.addprevious if look is not None else tblPr.append)(mar)    # tblCellMar must precede tblLook
        RULE, THIN = (8, "000000"), (4, "BFBFBF")
        for i, values in enumerate([cols] + rows):
            row = tb.rows[i]
            head = i == 0
            last = i == len(rows)
            is_total = (not head) and str(values[0]).startswith("Total")
            row_flags(row, header=head)
            for j, v in enumerate(values):
                c = row.cells[j]
                c.width = widths[j]
                if head:
                    set_border(c, top=RULE, bottom=RULE)        # tcBorders must precede shd
                    shade(c, "EDEDED")
                else:
                    set_border(c, top=RULE if is_total else None, bottom=RULE if last else THIN)
                lines = str(v).split("\n")
                c.paragraphs[0].text = ""
                for k, line in enumerate(lines):
                    p = c.paragraphs[0] if k == 0 else c.add_paragraph()
                    p.alignment = aligns[j]
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.keep_with_next = head or (keep and not last)
                    run = p.add_run(line)
                    run.font.size = size
                    run.bold = head or is_total
        for j, col in enumerate(tb.columns):
            col.width = widths[j]

    # ---- front matter
    para("Data S1. Supporting Information", size=16, bold=True, after=6, keep=True)
    para(MANUSCRIPT_TITLE, size=11, italic=True, after=12)
    para("General conventions", size=12, bold=True, before=6, after=6, keep=True)
    for c in conventions:
        lead, rest = c.split(". ", 1)
        p = para(after=5)
        r1 = p.add_run(lead + ". ")
        r1.bold, r1.font.size = True, Pt(9.5)
        r2 = p.add_run(rest)
        r2.font.size = Pt(9.5)
    para("List of tables and figures", size=12, bold=True, before=12, after=6, keep=True)
    items = [(f"Table {b['id']}.", b["title"]) for b in blocks] + \
            [(f"Figure {k}.", FIGURE_LEGENDS[k][0]) for k in figures]
    at = next((i for i, (lab, _) in enumerate(items) if lab == "Table S3c."), None)
    if at is not None:                         # Table S3d is in the Excel file only
        items.insert(at + 1, ("Table S3d.", "Report-level exposure classification of all reports (Excel file only; Data S2)"))
    for lab, title in items:
        p = para(after=2)
        p.paragraph_format.left_indent = Twips(1300)
        p.paragraph_format.first_line_indent = Twips(-1300)
        p.paragraph_format.tab_stops.add_tab_stop(Twips(1300))
        r1 = p.add_run(lab + "\t")
        r1.bold, r1.font.size = True, Pt(9)
        r2 = p.add_run(title)
        r2.font.size = Pt(9)

    # ---- tables
    for b in blocks:
        new_group = re.fullmatch(r"S\d+a?", b["id"]) is not None
        p = para(before=0 if new_group else 18, after=6, keep=True, page_break=new_group)
        r1 = p.add_run(f"Table {b['id']}. {b['title']}")
        r1.bold, r1.font.size = True, Pt(10.5)
        table(b)
        for note in b["notes"]:
            para(note, size=8, before=4)

    # ---- figures
    for k, path in figures.items():
        title, legend = FIGURE_LEGENDS[k]
        p = para(after=6, keep=True, page_break=True)
        r1 = p.add_run(f"Figure {k}. {title}")
        r1.bold, r1.font.size = True, Pt(10.5)
        pic = para(keep=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        pic.add_run().add_picture(path, width=Inches(9.6) if k != "S2" else Inches(9.8))
        para(legend, size=8, before=4)

    doc.save(out_path)


# ======================================================================================
# 8b. MANUSCRIPT TABLE 2 (Word table, ready to paste into the manuscript)
# ======================================================================================
# Formatting copies the manuscript's existing tables: Helvetica Neue 7.5 pt, single black grid lines,
# gray header rows and first column, and a fixed 6.5-inch width (the full text width, like Table 1).

TABLE2_FONT, TABLE2_PT = "Helvetica Neue", 7.5
TABLE2_HEAD_FILL, TABLE2_LABEL_FILL = "B0B3B2", "D4D4D4"
TABLE2_COL_DXA = [1200, 930, 930, 1300, 640, 600, 930, 930, 1300, 600]   # column widths; sum 9360 = 6.5 in
SA3_KEY = "SA3 (product-name validation)"


def table2_content(key):
    """Group headings, column headings, and body rows of manuscript Table 2: the 10 primary signals in the
    primary cohorts and in the co-primary SA3 cohorts, ordered by primary Fisher exact P value."""
    P = key["primary"].set_index("reaction")
    S3 = key["sa_res"][SA3_KEY].set_index("reaction")

    def n_pct(n, N):
        return f"{int(n)} ({100 * n / N:.1f}%)"

    def ror_ci(r):
        return f"{fmt_ror(r)} ({fmt_ci(r)})"

    rows = []
    for t in key["sig10"]:
        p, s = P.loc[t], S3.loc[t]
        rows.append([tdisp(t), n_pct(p.relaxer_n, p.relaxer_N), n_pct(p.keratin_n, p.keratin_N), ror_ci(p),
                     fmt_p(p.p), fmt_p(p.q), n_pct(s.relaxer_n, s.relaxer_N), n_pct(s.keratin_n, s.keratin_N),
                     ror_ci(s), fmt_p(s.q)])
    nR, nK = int(P.relaxer_N.iloc[0]), int(P.keratin_N.iloc[0])
    sR, sK = int(S3.relaxer_N.iloc[0]), int(S3.keratin_N.iloc[0])
    groups = ["Primary analysis (keyword-defined cohorts)",
              "Co-primary analysis (SA3, product-name-validated cohorts)"]
    n_eq = lambda n: f"n\u00a0=\u00a0{n}"                 # non-breaking spaces keep "n = 76" on one line
    cols = ["Reaction",
            f"Relaxer/alkali ({n_eq(nR)}), No. (%)", f"Keratin/blowout ({n_eq(nK)}), No. (%)", "ROR (95% CI)",
            "P Value", "FDR q",
            f"Relaxer/alkali ({n_eq(sR)}), No. (%)", f"Keratin/blowout ({n_eq(sK)}), No. (%)", "ROR (95% CI)", "FDR q"]
    return groups, cols, rows


def build_table2(key, out_path):
    """Write manuscript Table 2 as the only content of a Word file, so it can be copied and pasted in place
    of the old table. The title and legend stay in the manuscript."""
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, Twips

    groups, cols, rows = table2_content(key)
    W = TABLE2_COL_DXA
    if len(W) != len(cols):
        raise ValueError("TABLE2_COL_DXA needs one width per column")

    def el(tag, **attrs):
        e = OxmlElement(tag)
        for k, v in attrs.items():
            e.set(qn(k), str(v))
        return e

    doc = Document()
    sec = doc.sections[0]                                   # US Letter, portrait, 1-inch margins
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(1))
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = TABLE2_FONT, Pt(TABLE2_PT)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), TABLE2_FONT)
    normal.paragraph_format.space_before = normal.paragraph_format.space_after = Pt(0)
    for z in doc.settings.element.findall(qn("w:zoom")):     # the default template omits the required zoom percent
        z.set(qn("w:percent"), "100")

    tb = doc.add_table(rows=2 + len(rows), cols=len(cols))
    tb.autofit = False
    tblPr = tb._tbl.tblPr                                   # rebuilt in schema order
    for child in list(tblPr):
        tblPr.remove(child)
    tblPr.append(el("w:tblW", **{"w:w": sum(W), "w:type": "dxa"}))
    borders = el("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(el(f"w:{edge}", **{"w:val": "single", "w:sz": 6, "w:space": 0, "w:color": "000000"}))
    tblPr.append(borders)
    tblPr.append(el("w:tblLayout", **{"w:type": "fixed"}))
    mar = el("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        mar.append(el(f"w:{side}", **{"w:w": 60, "w:type": "dxa"}))
    tblPr.append(mar)
    tblPr.append(el("w:tblLook", **{"w:val": "0600", "w:firstRow": 0, "w:lastRow": 0, "w:firstColumn": 0,
                                    "w:lastColumn": 0, "w:noHBand": 1, "w:noVBand": 1}))
    for gc, w in zip(tb._tbl.tblGrid.findall(qn("w:gridCol")), W):
        gc.set(qn("w:w"), str(w))

    # two header rows: "Reaction" spans both; the group headings span their columns
    n_p = 5                                                 # primary columns after "Reaction"
    tb.cell(0, 0).merge(tb.cell(1, 0))
    tb.cell(0, 1).merge(tb.cell(0, n_p))
    tb.cell(0, n_p + 1).merge(tb.cell(0, len(cols) - 1))

    def put(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text.replace("/", "/​"))      # zero-width space lets "Relaxer/alkali" wrap at the slash
        run.bold = bold
        run.font.name, run.font.size = TABLE2_FONT, Pt(TABLE2_PT)
        run._element.rPr.rFonts.set(qn("w:eastAsia"), TABLE2_FONT)

    put(tb.cell(0, 0), cols[0], bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    put(tb.cell(0, 1), groups[0], bold=True)
    put(tb.cell(0, n_p + 1), groups[1], bold=True)
    for j in range(1, len(cols)):
        put(tb.cell(1, j), cols[j], bold=True)
    for i, values in enumerate(rows, start=2):
        put(tb.cell(i, 0), values[0], bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        for j in range(1, len(cols)):
            put(tb.cell(i, j), values[j])

    # widths (merged cells get the sum of the columns they span), shading, alignment, row flags
    for i, row in enumerate(tb.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(el("w:cantSplit"))
        if i < 2:
            trPr.append(el("w:tblHeader"))
        col = 0
        for tc in row._tr.tc_lst:
            span = tc.grid_span
            tcPr = tc.get_or_add_tcPr()
            tcPr.get_or_add_tcW().set(qn("w:w"), str(sum(W[col:col + span])))
            tcPr.get_or_add_tcW().set(qn("w:type"), "dxa")
            fill = TABLE2_HEAD_FILL if i < 2 else (TABLE2_LABEL_FILL if col == 0 else None)
            if fill:
                tcPr.append(el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": fill}))
            tcPr.append(el("w:vAlign", **{"w:val": "center"}))
            col += span
    doc.save(out_path)


# ======================================================================================
# 9. QUALITY-CONTROL CHECKS
# ======================================================================================

def run_qc(wb_path, xlsx_path, working_path, key, out_md):
    xl = pd.ExcelFile(wb_path)
    st = pd.ExcelFile(xlsx_path)
    S = {n: st.parse(n) for n in st.sheet_names}
    if os.path.exists(working_path):
        wx = pd.ExcelFile(working_path)
        S.update({n: wx.parse(n) for n in wx.sheet_names})

    def sheet(prefix):
        return S[[n for n in S if n.startswith(prefix)][0]]

    results = []

    def check(group, name, ok, detail=""):
        results.append((group, name, "PASS" if ok else "FAIL", detail))

    df, primary, sa_res, extra = key["df"], key["primary"], key["sa_res"], key["extra_res"]
    sig10 = key["sig10"]
    P = primary.set_index("reaction")
    nR, nK = int(df.relaxer.sum()), int((df.relaxer == 0).sum())

    # A. agreement with the pre-computed sheets in the source workbook
    wbp = xl.parse("Primary_ROR").dropna(subset=["Alkali n"])
    wbp["R"] = wbp["Reaction"].str.upper().str.strip()
    check("A. Workbook", "Tested family equals the workbook's Primary_ROR reactions", set(wbp.R) == set(key["terms"]),
          f"{len(key['terms'])} vs {len(wbp)}")
    mism = []
    for _, w in wbp.iterrows():
        if w.R not in P.index:
            mism.append((w.R, "missing"))
            continue
        r = P.loc[w.R]
        if (r.relaxer_n, r.keratin_n) != (int(w["Alkali n"]), int(w["Keratin/blowout n"])): mism.append((w.R, "counts"))
        if not math.isclose(r.p, w["Fisher p"], rel_tol=1e-6, abs_tol=1e-12): mism.append((w.R, "p"))
        if not math.isclose(r.q, w["BH-FDR q"], rel_tol=1e-6, abs_tol=1e-12): mism.append((w.R, "q"))
        wr = w["ROR (alkali vs keratin/blowout)"]
        if not (math.isclose(r.ror, wr, rel_tol=1e-6) or (r.ror == 0 and wr == 0)): mism.append((w.R, "ROR"))
        if not math.isclose(r.hi, w["95% CI high"], rel_tol=1e-3): mism.append((w.R, "CI high"))
        if not math.isclose(r.lo, w["95% CI low"], rel_tol=1e-3, abs_tol=1e-6): mism.append((w.R, "CI low"))
        if (w["FDR significant"] == "Yes") != bool(r.fdr_sig): mism.append((w.R, "sig"))
    check("A. Workbook", "Primary: counts, ROR, CI, P, q, significance match workbook", not mism, f"{len(mism)} mismatches {mism[:5]}")
    coh = xl.parse("Sensitivity_Cohorts").set_index("Analysis")
    nm = {"SA1 (single cosmetic product)": "SA1_OneCosmetic", "SA2 (single suspect product)": "SA2_OneSuspect",
          "SA3 (product-name validation)": "SA3_StrictName", "SA4 (SA3 + single suspect)": "SA4_StrictOneSuspect"}
    bad = [k for k, v in nm.items() if (sa_res[k].relaxer_N.iloc[0], sa_res[k].keratin_N.iloc[0]) !=
           (coh.loc[v, "Alkali N"], coh.loc[v, "Keratin/blowout N"])]
    check("A. Workbook", "SA1-SA4 cohort sizes match workbook", not bad, str(bad))
    rob = xl.parse("Robustness_Summary")
    rob["R"] = rob.Reaction.str.upper().str.strip()
    mm = []
    for k, tag in zip(nm, ["SA1", "SA2", "SA3", "SA4"]):
        Rr = sa_res[k].set_index("reaction")
        for _, w in rob.iterrows():
            r = Rr.loc[w.R]
            if not math.isclose(r.q, w[f"{tag} q"], rel_tol=1e-5, abs_tol=1e-9): mm.append((tag, w.R, "q"))
            if not (math.isclose(r.ror, w[f"{tag} ROR"], rel_tol=1e-5) or (r.ror == 0 and w[f"{tag} ROR"] == 0)): mm.append((tag, w.R, "ROR"))
            if (w[f"{tag} sig"] == "Yes") != bool(r.fdr_sig): mm.append((tag, w.R, "sig"))
    check("A. Workbook", "SA1-SA4: ROR, q and significance match workbook's Robustness_Summary", not mm, f"{len(mm)} mismatches {mm[:5]}")
    raw = xl.parse("Temporal_Trend", header=None)
    hr = next(i for i in range(raw.shape[0]) if str(raw.iat[i, 0]).strip() == "Year")
    end = hr + 1
    while end < raw.shape[0] and pd.notna(raw.iat[end, 0]) and str(raw.iat[end, 0]).strip().isdigit():
        end += 1
    ann_wb = raw.iloc[hr + 1:end, :11]
    ann_wb.columns = list(raw.iloc[hr, :11])
    ann = key["annual"].set_index("Year")
    bad = []
    for _, w in ann_wb.iterrows():
        y = int(w.Year)
        if (ann.loc[y, "Relaxer/alkali reports"], ann.loc[y, "Keratin/blowout reports"]) != (int(w["Relaxer/alkali reports"]), int(w["Keratin/blowout reports"])): bad.append((y, "reports"))
        if (ann.loc[y, "Alopecia, Relaxer/alkali n"], ann.loc[y, "Alopecia, Keratin/blowout n"]) != (int(w["Relaxer/alkali alopecia, n"]), int(w["Keratin/blowout alopecia, n"])): bad.append((y, "alopecia"))
    check("A. Workbook", "Annual report and alopecia counts match workbook's Temporal_Trend (Fig 4)", not bad, str(bad))

    # B. values hard-coded from the submitted manuscript (expected to fail if the data change)
    ms_t2 = {"EYE IRRITATION": (1, 70, "0.04", "<.001", "<.001"), "ALOPECIA": (36, 50, "4.07", "<.001", "<.001"),
             "DYSPNOEA": (1, 41, "0.08", "<.001", ".007"), "DIZZINESS": (0, 30, "≤0.44", "<.001", ".009"),
             "COUGH": (0, 25, "≤0.54", ".004", ".025"), "NASAL DISCOMFORT": (0, 25, "≤0.54", ".004", ".025"),
             "NAUSEA": (0, 25, "≤0.54", ".004", ".025"), "LACRIMATION INCREASED": (0, 23, "≤0.60", ".006", ".036"),
             "TRICHORRHEXIS": (20, 36, "2.38", ".008", ".038"), "HEADACHE": (8, 66, "0.37", ".011", ".048")}
    bad = []
    for t, exp in ms_t2.items():
        if t not in P.index:
            bad.append(t)
            continue
        r = P.loc[t]
        rtxt = f"≤{r.hi:.2f}" if r.relaxer_n == 0 else f"{r.ror:.2f}"
        if (r.relaxer_n, r.keratin_n, rtxt, fmt_p(r.p), fmt_p(r.q)) != exp: bad.append(t)
    check("B. Manuscript", "Submitted Table 2 (counts, ROR/upper limit, P, q) for all 10 reactions", not bad and set(ms_t2) == set(sig10), str(bad))
    check("B. Manuscript", "Woolf CIs quoted in the submitted text",
          all(t in P.index and f"{P.loc[t].lo:.2f}-{P.loc[t].hi:.2f}" == s for t, s in
              [("ALOPECIA", "2.36-7.01"), ("TRICHORRHEXIS", "1.28-4.42"), ("HEADACHE", "0.17-0.82"),
               ("EYE IRRITATION", "0.01-0.29"), ("DYSPNOEA", "0.01-0.57")]))

    def pct(n, N): return f"{n} ({100 * n / N:.1f}%)"
    t1 = {}
    for c_, flag in (("R", 1), ("K", 0)):
        g = df[df.relaxer == flag]
        N = len(g)
        ages = g.age.dropna()
        t1[c_] = dict(female=pct((g.Gender == "Female").sum(), N), male=pct((g.Gender == "Male").sum(), N),
                      age_av=pct(len(ages), N), age=f"{ages.median():.1f} ({ages.quantile(.25):.1f}-{ages.quantile(.75):.1f})",
                      hosp=pct((g.Hospitalization == "Yes").sum(), N), dis=pct((g.Disability == "Yes").sum(), N),
                      lt=pct((g["Life Threatening"] == "Yes").sum(), N), death=pct((g.Death == "Yes").sum(), N))
    ms_t1 = {"R": dict(female="71 (93.4%)", male="2 (2.6%)", age_av="36 (47.4%)", age="47.0 (36.8-53.8)",
                       hosp="3 (3.9%)", dis="1 (1.3%)", lt="0 (0.0%)", death="0 (0.0%)"),
             "K": dict(female="246 (89.1%)", male="11 (4.0%)", age_av="197 (71.4%)", age="38.0 (31.0-47.0)",
                       hosp="18 (6.5%)", dis="18 (6.5%)", lt="16 (5.8%)", death="2 (0.7%)")}
    bad = [(c_, k) for c_ in ms_t1 for k in ms_t1[c_] if ms_t1[c_][k] != t1[c_][k]]
    check("B. Manuscript", "Submitted Table 1 (sex, age availability, median age/IQR, serious outcomes)", not bad, str(bad))

    # C. internal consistency of the supplementary workbook
    e1s = sheet("TableS1").set_index("Reaction (MedDRA PT)")
    e2s = sheet("TableS2").set_index("Reaction (MedDRA PT)")
    t2 = sheet("DRAFT_Table2")
    bad = []
    for _, r in t2.iterrows():
        a_, b_ = e1s.loc[r.Reaction], e2s.loc[r.Reaction]
        if (r["Primary ROR"], r["Primary 95% CI"], r["Primary q"]) != (a_.ROR, a_["95% CI"], a_["q (BH-FDR)"]): bad.append(("P", r.Reaction))
        if (r["SA3 ROR"], r["SA3 95% CI"], r["SA3 q"]) != (b_.ROR, b_["95% CI"], b_["q (BH-FDR)"]): bad.append(("SA3", r.Reaction))
    check("C. Internal", "Draft revised Table 2 equals Table S1 (primary) and Table S2 (SA3)", not bad, str(bad))
    t2_docx = os.path.join(os.path.dirname(xlsx_path), "Table2.docx")
    if os.path.exists(t2_docx):
        from docx import Document
        tbl = Document(t2_docx).tables[0]
        got = [[c.text.replace("​", "") for c in row.cells] for row in tbl.rows[2:]]

        def n_pct(s):                                       # "1/76 (1.3)" -> "1 (1.3%)"
            n, pct = re.match(r"(\d+)/\d+ \(([\d.]+)\)", s).groups()
            return f"{n} ({pct}%)"
        exp = []
        for rx_ in t2.Reaction:
            a_, b_ = e1s.loc[rx_], e2s.loc[rx_]
            exp.append([" ".join(w.capitalize() for w in rx_.split()),
                        n_pct(a_[f"{RELAXER}, n/N (%)"]), n_pct(a_[f"{KERATIN}, n/N (%)"]),
                        f"{a_.ROR} ({a_['95% CI']})", a_["P (Fisher)"], a_["q (BH-FDR)"],
                        n_pct(b_[f"{RELAXER}, n/N (%)"]), n_pct(b_[f"{KERATIN}, n/N (%)"]),
                        f"{b_.ROR} ({b_['95% CI']})", b_["q (BH-FDR)"]])
        diff = [e[0] for g, e in zip(got, exp) if g != e]
        check("C. Internal", "Manuscript Table 2 (Table2.docx) equals Tables S1 (primary) and S2 (SA3)",
              len(got) == len(exp) and not diff, f"{len(got)} rows; differing: {diff}")
    e4 = sheet("TableS4")
    c_p = [c for c in e4.columns if c.startswith("Primary")][0]
    c_3 = [c for c in e4.columns if c.startswith("SA3")][0]
    bad = [r.Reaction for _, r in e4.iterrows()
           if not str(r[c_p]).startswith(f"{e1s.loc[r.Reaction].ROR} ({e1s.loc[r.Reaction]['95% CI']})")
           or not str(r[c_3]).startswith(f"{e2s.loc[r.Reaction].ROR} ({e2s.loc[r.Reaction]['95% CI']})")]
    check("C. Internal", "Table S4 primary and SA3 columns equal Tables S1 and S2", not bad, str(bad))
    co = sheet("TableS10a").set_index(sheet("TableS10a").columns[0])
    bad = [r for r in co.index if co.loc[r, r] != int(e1s.loc[r]["Keratin/blowout, n/N (%)"].split("/")[0])]
    check("C. Internal", "Co-occurrence diagonal equals keratin counts in Table S1", not bad, str(bad))
    cl = sheet("TableS11b")
    cl = cl[cl.Analysis == "Primary"].set_index("Cluster")
    bad = []
    for c_, ts in CLUSTERS.items():
        members = [disp(t) for t in ts if disp(t) in e1s.index]
        for col in ("Relaxer/alkali, n/N (%)", "Keratin/blowout, n/N (%)"):
            if members and int(cl.loc[c_, col].split("/")[0]) < max(int(e1s.loc[m][col].split("/")[0]) for m in members):
                bad.append((c_, col))
    check("C. Internal", "Each cluster count >= its largest component reaction count", not bad, str(bad))
    pc = sheet("TableS8a").set_index("Reaction")
    bad = [t for t in pc.index if int(pc.loc[t, "Reports with reaction"]) != int(
        e1s.loc[t]["Relaxer/alkali, n/N (%)" if pc.loc[t, "Cohort with higher reporting"] == RELAXER else "Keratin/blowout, n/N (%)"].split("/")[0])]
    check("C. Internal", "Product-concentration report counts equal Table S1 counts", not bad, str(bad))
    dup = key["dup"]
    d_res = extra["Possible duplicates removed"]
    rem = dup[dup.retained != "Yes"]
    check("C. Internal", "De-duplicated cohort sizes = primary minus removed reports",
          (d_res.relaxer_N.iloc[0], d_res.keratin_N.iloc[0]) == (nR - (rem.cohort == RELAXER).sum(), nK - (rem.cohort == KERATIN).sum()))
    check("C. Internal", "Every possible-duplicate cluster retains exactly one report",
          len(dup) == 0 or (dup.groupby("cluster").retained.apply(lambda s: (s == "Yes").sum()) == 1).all())
    ser = sheet("TableS12_")
    check("C. Internal", "Serious-outcome listing contains every death/life-threatening/disability report",
          len(ser) == int(((df.Death == "Yes") | (df["Life Threatening"] == "Yes") | (df.Disability == "Yes")).sum()), f"{len(ser)} rows")
    pp = sheet("TableS13b")
    Ns = pp.drop_duplicates(["Period", "Cohort"]).groupby("Cohort").N.sum().to_dict()
    check("C. Internal", "Period denominators sum to the cohort sizes", Ns == {KERATIN: nK, RELAXER: nR}, str(Ns))
    th = sheet("TableS9a")
    prim = th[th["Inclusion rule"].str.contains("primary")].iloc[0]
    check("C. Internal", "Threshold table primary rule matches the primary analysis",
          (int(prim["Reactions tested"]), int(prim["FDR-significant"])) == (len(key["terms"]), len(sig10)))
    ann_s = sheet("TableS13a")
    check("C. Internal", "Annual counts sum to the cohort sizes",
          ann_s["Relaxer/alkali reports"].sum() == nR and ann_s["Keratin/blowout reports"].sum() == nK)
    s3d = sheet("TableS3d")
    bad = []
    for sa_, k in zip(["SA1", "SA2", "SA3", "SA4"], sa_res):
        n = s3d[s3d[sa_] == "Yes"].groupby("Cohort").size()
        if (n.get(RELAXER, 0), n.get(KERATIN, 0)) != (int(sa_res[k].relaxer_N.iloc[0]), int(sa_res[k].keratin_N.iloc[0])):
            bad.append(sa_)
    check("C. Internal", "Report-level classification flags (Table S3d) reproduce SA1-SA4 cohort sizes", not bad, str(bad))
    sp = sheet("TableS14").set_index("Check")
    check("C. Internal", "No report created on/after 29 Dec 2023 (MoCRA); report numbers unique",
          int(sp.loc["Reports created on/after 2023-12-29 (MoCRA)", "Value"]) == 0
          and int(sp.loc["Distinct report numbers", "Value"]) == len(df))

    # D. statistics re-derived by hand
    allres = [primary] + list(sa_res.values()) + list(extra.values())
    check("D. Statistics", "q >= P for every reaction in every analysis", all((r.q >= r.p - 1e-15).all() for r in allres))

    def bh(p):
        p = np.asarray(p)
        n = len(p)
        o = np.argsort(p)
        q = np.empty(n)
        q[o] = np.minimum.accumulate((p[o] * n / np.arange(1, n + 1))[::-1])[::-1]
        return np.minimum(q, 1)
    check("D. Statistics", "Benjamini-Hochberg q re-derived by hand matches in every analysis",
          all(np.allclose(bh(r.p.values), r.q.values) for r in allres))
    w = primary[primary.ci_method == "Woolf"]
    check("D. Statistics", "Every Woolf CI contains its ROR; every zero-cell row uses an exact CI and is flagged",
          ((w.lo <= w.ror) & (w.ror <= w.hi)).all() and (primary[primary.ci_method != "Woolf"].boundary != "").all())
    t0 = sig10[0]
    a, c = int(P.loc[t0].relaxer_n), int(P.loc[t0].keratin_n)
    pmf = hypergeom(len(df), nR, a + c).pmf(np.arange(0, a + c + 1))
    p_hand = pmf[pmf <= pmf[a] * (1 + 1e-7)].sum()
    check("D. Statistics", f"Fisher P for {disp(t0).lower()} re-derived from the hypergeometric distribution",
          math.isclose(p_hand, P.loc[t0].p, rel_tol=1e-6), f"{p_hand:.6g} vs {P.loc[t0].p:.6g}")

    # E. direction consistency
    rev = []
    for r in allres:
        Rr = r.set_index("reaction")
        for t in sig10:
            x = Rr.loc[t]
            if x.relaxer_n / x.relaxer_N == x.keratin_n / x.keratin_N:
                continue
            if x.direction != P.loc[t].direction:
                rev.append((r.analysis.iloc[0], t))
    check("E. Direction", "No sensitivity analysis reverses the direction of any primary signal", not rev, str(rev))
    f3 = figure3_matrix(key)
    check("E. Direction", "Figure 3: no significant cell has a direction opposite to the primary estimate",
          not f3["flipped"], str(f3["flipped"]))
    f1 = figure1_counts(key)
    check("B. Manuscript", "Figure 1 counts (N, cohorts, SA1-SA4) match the submitted Figure 1",
          (f1["all"], f1["N"], f1["R"], f1["K"], f1["SA1"], f1["SA2"], f1["SA3"], f1["SA4"]) ==
          (52219, 352, 76, 276, (60, 238), (65, 242), (65, 204), (59, 190)))
    check("C. Internal", "Figure 1 SA3 exclusions by reason sum to the primary cohort minus SA3",
          f1["sa3_excluded"] == (f1["R"] - f1["SA3"][0], f1["K"] - f1["SA3"][1]), str(f1["sa3_excluded"]))

    # F. plausibility
    obs = set().union(*df.rx)
    miss = sorted(t for v in CLUSTERS.values() for t in v if t not in obs)
    check("F. Plausibility", "Every cluster term occurs in the data (no British/American spelling mismatches)", not miss, str(miss))
    dupl = [t for t in obs if sum(t in v for v in CLUSTERS.values()) > 1]
    check("F. Plausibility", "No MedDRA term is assigned to more than one cluster", not dupl, str(dupl))
    s4 = df[df["SA4 keep"] == "Yes"]
    check("F. Plausibility", "SA4 = SA3 reports with exactly one suspect product",
          (s4["SA3 keep"] == "Yes").all() and (s4["N SUSPECT"] == 1).all())
    txt = " ".join(" ".join(map(str, S[n].astype(str).values.ravel())) for n in S)
    check("F. Plausibility", "No malformed labels ('=<', 'nan%', '0.00-', 'inf') anywhere in the supplementary workbook",
          not any(x in txt for x in ["=<", "nan%", " 0.00-", "inf)", "(0.00"]))

    # robustness matrix
    analyses = {"Primary": primary, **{k.split(" (")[0]: v for k, v in sa_res.items()}, **extra}
    mat = pd.DataFrame({k: [("Y" if v.set_index("reaction").loc[t].fdr_sig else ".") for t in sig10] for k, v in analyses.items()},
                       index=[disp(t) for t in sig10])
    mat["Period-stratified (P<.05)"] = ["Y" if key["period_mh"][t][3] < .05 else "." for t in sig10]
    mat["Top brand removed (P<.05)"] = ["Y" if key["loo"][t]["p"] < .05 else "." for t in sig10]

    n_fail = sum(1 for r in results if r[2] == "FAIL")
    lines = [f"# QC report: {len(results)} checks, {len(results) - n_fail} passed, {n_fail} failed\n",
             "Group B compares against values hard-coded from the submitted manuscript; those checks are expected to "
             "fail if the input data change.\n"]
    for g in sorted({r[0] for r in results}):
        lines.append(f"\n## {g}\n")
        for _, nme, res, det in [r for r in results if r[0] == g]:
            lines.append(f"- [{res}] {nme}" + (f" -- {det}" if res == "FAIL" and det else ""))
    lines.append("\n## Robustness matrix (Y = FDR q<.05, or P<.05 where noted)\n")
    header = "| Reaction | " + " | ".join(mat.columns) + " |"
    lines += [header, "|" + "---|" * (len(mat.columns) + 1)]
    lines += ["| " + i + " | " + " | ".join(mat.loc[i]) + " |" for i in mat.index]
    with open(out_md, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    return len(results), n_fail


# ======================================================================================
# 10. MAIN
# ======================================================================================

def main():
    wb_path = os.path.expanduser(WORKBOOK_PATH)
    repo_default = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "data",
                                "Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx")
    in_repo = False
    if not os.path.isfile(wb_path) and os.path.isfile(repo_default):
        wb_path = os.path.normpath(repo_default)       # running inside the GitHub repository layout
        in_repo = True
    if not os.path.isfile(wb_path):
        sys.exit(f"Spreadsheet not found:\n  {wb_path}\nEdit WORKBOOK_PATH at the top of this script.")
    if OUTPUT_DIR:
        out_dir = os.path.expanduser(OUTPUT_DIR)
    elif in_repo:
        out_dir = os.path.normpath(os.path.join(os.path.dirname(wb_path), os.pardir, "outputs"))
    else:
        out_dir = os.path.join(os.path.dirname(os.path.abspath(wb_path)), "Supplement_outputs")
    os.makedirs(out_dir, exist_ok=True)
    xlsx = os.path.join(out_dir, "Supplementary_Tables.xlsx")
    working = os.path.join(out_dir, "Supplementary_Tables_AUTHOR_WORKING.xlsx")

    print(f"Reading {wb_path}")
    key, sheets = run_analyses(wb_path)
    print(f"  {len(key['df'])} reports; {len(key['terms'])} reactions tested; "
          f"{len(key['sig10'])} FDR-significant: {', '.join(disp(t) for t in key['sig10'])}")
    print(f"  possible-duplicate clusters: {key['dup']['cluster'].nunique()} ({len(key['drop_ids'])} reports removed)")

    write_excel(sheets, len(key["terms"]), xlsx, working)
    print(f"Wrote {xlsx}\nWrote {working}  (author-only; do not upload)")

    if MAKE_MAIN_FIGURES:
        main_figs = make_main_figures(key, out_dir)
        print("Wrote main figures: " + ", ".join(os.path.basename(p) for p in main_figs.values())
              + " (+ PDF and EPS versions)")

    figures = {}
    if MAKE_FIGURES:
        figures = make_figures(key, out_dir)
        print("Wrote figures: " + ", ".join(os.path.basename(p) for p in figures.values()) + " (+ PDF versions)")

    if MAKE_WORD:
        blocks, conventions = word_spec(key, sheets)
        docx_path = os.path.join(out_dir, "Supporting_Information.docx")
        build_word(blocks, conventions, figures, docx_path)
        print(f"Wrote {docx_path}  ({len(blocks)} tables, {len(figures)} figures)")

    if MAKE_TABLE2:
        t2_path = os.path.join(out_dir, "Table2.docx")
        build_table2(key, t2_path)
        print(f"Wrote {t2_path}  (manuscript Table 2, ready to paste)")

    if RUN_QC:
        qc_path = os.path.join(out_dir, "QC_report.md")
        n, n_fail = run_qc(wb_path, xlsx, working, key, qc_path)
        print(f"Wrote {qc_path}  ({n} checks, {n - n_fail} passed, {n_fail} failed)")

    print("Done.")


if __name__ == "__main__":
    main()
