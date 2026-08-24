import sys
import os
import argparse
import math
import textwrap
import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.lines import Line2D

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------


def parse_args():
    """Command-line interface. Replaces v2's HARD_CODED_XLSX_PATH constant so
    this script is runnable by anyone with the workbook, not just on the
    machine/account it was originally written on."""
    parser = argparse.ArgumentParser(
        description="Generate publication-ready figures and tables from the "
                     "validated ROR analysis workbook (hair-straightening "
                     "cosmetics adverse event reporting study)."
    )
    parser.add_argument(
        "xlsx_path",
        help="Path to the validated ROR analysis Excel workbook, e.g. "
             "Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx",
    )
    parser.add_argument(
        "output_dir",
        nargs="?",
        default=None,
        help="Directory to write figures/, tables/, and the legend files into. "
             "Defaults to a 'pub_outputs' folder created next to the input "
             "workbook if omitted.",
    )
    return parser.parse_args()

DPI = 600                    # comfortably above the 300-350 ppi minimums cited across
                              # JAMA Network's own figure-prep pages
FIG_WIDTH_IN = 6.5            # conservative width ceiling -- see module docstring
SAVE_FORMATS = ["png", "pdf", "eps"]


def _pick_font(preferred):
    """Pick the first available font from a preference list instead of hardcoding
    one. Hardcoding "Arial" silently fell back to whatever matplotlib's default is
    (DejaVu Sans) on any machine without Arial installed -- one findfont warning
    per text element, and different font metrics (so different line wraps/spacing)
    depending on which computer ran the script. This makes the choice explicit and
    reproducible, and tells you plainly which font you actually got."""
    available = {f.name for f in fm.fontManager.ttflist}
    for name in preferred:
        if name in available:
            return name, name == preferred[0]
    return matplotlib.rcParams["font.sans-serif"][0], False


FONT_FAMILY, _is_exact_arial = _pick_font(
    ["Arial", "Helvetica", "Liberation Sans", "Nimbus Sans", "DejaVu Sans"]
)
if _is_exact_arial:
    print(f"[fonts] Using Arial for figure text.")
else:
    print(f"[fonts] Arial not found on this machine -- using '{FONT_FAMILY}' instead "
          f"for all figure text. For an exact match to the Arial used elsewhere in "
          f"the manuscript, install Arial or Liberation Sans (metric-compatible) and "
          f"rerun; otherwise this substitute is used consistently across every figure.")

plt.rcParams.update({
    "font.family": FONT_FAMILY,
    "font.size": 9,
    "axes.edgecolor": "#333333",
    "axes.linewidth": 0.8,
    "figure.facecolor": "white",
    "savefig.facecolor": "white",
    "pdf.fonttype": 42,            # embed text as real, editable/searchable TrueType
    "ps.fonttype": 42,             # outlines (Type 42) instead of matplotlib's default
                                    # Type 3 bitmapped glyphs -- Type 3 text can look
                                    # noticeably rougher when a journal or reviewer
                                    # zooms into the PDF/EPS, and some journal PDF
                                    # checkers flag Type 3 fonts outright.
    "lines.antialiased": True,
    "patch.antialiased": True,
    "lines.solid_capstyle": "round",
    "lines.solid_joinstyle": "round",   # smoother line joins/caps in the multi-segment
                                          # trend lines and error bars
})

INK_COLOR = "#1A1A2E"         # single print-safe marker color used in Figure 2
REF_LINE_COLOR = "#888888"
ROBUST_FILL = "#2C6E7F"       # heatmap "significant" cell color
NOT_SIG_FILL = "#EDEDED"
ARM_A_COLOR = "#B23A48"       # used only in eFigure 2 (temporal), not Figure 2
ARM_B_COLOR = "#2C6E7F"

# British MedDRA spelling -> American English for display. Extend as needed if reactions
# outside this analysis's 11-term set come into scope later.
REACTION_SPELLING_FIXES = {
    "DYSPNOEA": "DYSPNEA",
    "OEDEMA": "EDEMA",
    "HAEMORRHAGE": "HEMORRHAGE",
    "DIARRHOEA": "DIARRHEA",
    "ANAEMIA": "ANEMIA",
    "OESOPHAG": "ESOPHAG",
    "PAEDIATRIC": "PEDIATRIC",
    "HAEMATOMA": "HEMATOMA",
    "RHINORRHOEA": "RHINORRHEA",
    "PHARYNGITIS": "PHARYNGITIS",
    "HAEMOPTYSIS": "HEMOPTYSIS",
    "ANAESTHESIA": "ANESTHESIA",
    "PARAESTHESIA": "PARESTHESIA",
    "HYPOAESTHESIA": "HYPOESTHESIA",
    "HYPERAESTHESIA": "HYPERESTHESIA",
}


def canonical_key(name):
    """Uppercase + British->American normalization, used ONLY for matching reaction
    names across sheets/hardcoded lists that may use different spelling conventions --
    never for display."""
    s = str(name).strip().upper()
    for brit, us in REACTION_SPELLING_FIXES.items():
        s = s.replace(brit, us)
    return s


def display_reaction(name):
    """American-English, title-case display form of a reaction term."""
    s = canonical_key(name)
    return s.title()


# ----------------------------------------------------------------------------
# Formatting helpers (AMA Manual of Style conventions)
# ----------------------------------------------------------------------------

def format_p(p):
    if pd.isna(p):
        return "NA"
    if p < .001:
        return "<.001"
    s = f"{p:.3f}" if p < 0.1 else f"{p:.2f}"
    return s.lstrip("0") if s.startswith("0") else s


def format_ror_ci(ror, lo, hi):
    if pd.isna(ror) or pd.isna(lo) or pd.isna(hi):
        return "NA"
    if ror == 0:
        # Zero-cell reaction (no alkali-arm events): the ROR has no meaningful
        # point estimate, so printing "0.00" here would misstate the data --
        # it reads as a precise value when the true number is unknown/unbounded
        # below. Matches Figure 2's convention exactly (see the zero_cell branch
        # in make_figure2, which plots no point marker for these reactions and
        # labels them the same way) rather than diverging from it in the tables.
        return f"\u2264{hi:.2f} (upper CI; ROR\u21920)"
    return f"{ror:.2f} ({lo:.2f}-{hi:.2f})"


def format_pct(n, N):
    if pd.isna(n) or pd.isna(N) or N == 0:
        return "NA"
    return f"{int(n)}/{int(N)} ({100*n/N:.1f}%)"


def _nice_log_ticks(vmin, vmax, max_ticks=6):
    """Readable tick positions for a log-scale ROR axis, drawn from the standard
    1/2/5 x 10^k progression. Matplotlib's default log-axis formatter is fine when
    the visible range spans several decades, but when it spans less than one decade
    (typical for a forest plot of ROR point estimates, which usually cluster within
    a couple-fold of each other) it falls back to labeling every minor tick in
    scientific notation -- e.g. "2x10^0, 3x10^0, ... 9x10^0" -- which overlaps into
    an unreadable smear once there are more than 3-4 of them in a ~2 in wide plot."""
    if vmin <= 0:
        vmin = 1e-3
    k_lo = math.floor(math.log10(vmin))
    k_hi = math.ceil(math.log10(vmax))
    candidates = sorted(
        base * 10.0 ** k for k in range(k_lo - 1, k_hi + 2) for base in (1, 2, 5)
    )
    ticks = [t for t in candidates if vmin * 0.98 <= t <= vmax * 1.02]
    while len(ticks) > max_ticks:
        ticks = ticks[::2]
    if len(ticks) < 2:
        ticks = [round(vmin, 3), round(vmax, 3)]
    return ticks


def _fmt_axis_ratio(v):
    """Plain-decimal tick label (no scientific notation), leading zero kept for
    values <1 -- consistent with how ROR values are formatted everywhere else in
    this script (format_ror_ci), since these ticks sit on the same ROR scale."""
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v))}"
    return f"{v:.2f}".rstrip("0").rstrip(".")


# ----------------------------------------------------------------------------
# Loading
# ----------------------------------------------------------------------------

def load_workbook(path):
    xl = pd.ExcelFile(path)
    required = ["Primary_ROR", "Robustness_Summary", "Sensitivity_Cohorts"]
    missing = [s for s in required if s not in xl.sheet_names]
    if missing:
        raise ValueError(f"Missing expected sheet(s): {missing}. Found: {xl.sheet_names}")

    def _clean_ror_sheet(name):
        """Primary_ROR (and, in this workbook, ExactCase_Sensitivity) end with a blank
        row followed by a free-text footnote row -- e.g. 'Recomputed directly from all
        76 alkali...' -- stuffed into the Reaction column with every numeric column
        NaN. Left in, that row sorts to the bottom of any Fisher-p-sorted table and
        prints as a garbage 'Nan' reaction with NA everywhere else. Drop rows with no
        Alkali n instead of assuming a fixed row count, so this survives the footnote
        text (or its row position) changing later."""
        df = xl.parse(name)
        return df.dropna(subset=["Alkali n"]).reset_index(drop=True)

    out = {
        "primary": _clean_ror_sheet("Primary_ROR"),
        "robustness": xl.parse("Robustness_Summary"),
        "cohorts": xl.parse("Sensitivity_Cohorts"),
    }
    if "Temporal_Trend" in xl.sheet_names:
        out["temporal_annual"], out["temporal_reportlevel"] = load_temporal_sheet(path)
    else:
        out["temporal_annual"], out["temporal_reportlevel"] = None, None

    # Optional sheets, used only for the supplemental tables added below. Guarded so
    # this still runs against an older workbook that predates them.
    out["exact_case"] = (
        _clean_ror_sheet("ExactCase_Sensitivity") if "ExactCase_Sensitivity" in xl.sheet_names else None
    )
    out["report_level_qc"] = (
        xl.parse("Report_Level_QC") if "Report_Level_QC" in xl.sheet_names else None
    )
    out["strict_excluded"] = (
        xl.parse("Strict_Excluded_Reports") if "Strict_Excluded_Reports" in xl.sheet_names else None
    )
    out["report_characteristics"] = (
        xl.parse("Report_Characteristics") if "Report_Characteristics" in xl.sheet_names else None
    )
    out["table1_summary_lookup"] = (
        _parse_table1_summary_lookup(xl.parse("Table1_Summary", header=None))
        if "Table1_Summary" in xl.sheet_names else None
    )
    return out


def _parse_table1_summary_lookup(raw_df):
    """Positional parse of the workbook's own Table1_Summary sheet -- title row, a
    header row ('Characteristic', then one column per arm), section-header rows with
    blank values, and data rows -- into a {characteristic label: [value per arm column,
    in sheet column order]} lookup. Used only as the reference for the table3
    cross-check below; mirrors the positional-parse approach already used for
    Temporal_Trend above, since this sheet has the same 'not a clean single header
    row' shape. Returns None if no 'Characteristic' header cell is found, so a
    differently-shaped or renamed sheet degrades to skipping the cross-check rather
    than raising."""
    header_row = None
    for i in range(raw_df.shape[0]):
        if str(raw_df.iat[i, 0]).strip() == "Characteristic":
            header_row = i
            break
    if header_row is None:
        return None
    lookup = {}
    for i in range(header_row + 1, raw_df.shape[0]):
        label = str(raw_df.iat[i, 0]).strip()
        if not label or label.lower() == "nan":
            continue
        lookup[label] = [str(raw_df.iat[i, j]).strip() for j in range(1, raw_df.shape[1])]
    return lookup


def load_temporal_sheet(path):
    """The Temporal_Trend sheet packs a documentation header, an annual-summary table
    (cols 0-10), and a report-level table (cols 19-28) into one worksheet with no clean
    single header row -- pandas' auto-header parsing garbles it. Parse it positionally
    by locating the two tables' header cells directly, rather than assuming fixed row
    numbers, so this survives someone inserting/removing a documentation line later."""
    raw = pd.read_excel(path, sheet_name="Temporal_Trend", header=None)

    annual_header_row = None
    for i in range(raw.shape[0]):
        if str(raw.iat[i, 0]).strip() == "Year":
            annual_header_row = i
            break
    if annual_header_row is None:
        raise ValueError("Could not locate the annual-summary table (looked for a "
                          "'Year' header in column A of Temporal_Trend).")

    rl_header_row = rl_header_col = None
    for i in range(raw.shape[0]):
        for j in range(raw.shape[1]):
            if str(raw.iat[i, j]).strip() == "Report number":
                rl_header_row, rl_header_col = i, j
                break
        if rl_header_row is not None:
            break
    if rl_header_row is None:
        raise ValueError("Could not locate the report-level table (looked for a "
                          "'Report number' header) in Temporal_Trend.")

    last = annual_header_row
    for i in range(annual_header_row + 1, raw.shape[0]):
        if pd.isna(raw.iat[i, 0]):
            break
        last = i
    annual_cols = list(raw.iloc[annual_header_row, 0:11])
    annual = raw.iloc[annual_header_row + 1: last + 1, 0:11].copy()
    annual.columns = annual_cols
    annual = annual.reset_index(drop=True)
    annual["Year"] = annual["Year"].astype(int)

    rl_cols = list(raw.iloc[rl_header_row, rl_header_col:rl_header_col + 10])
    rl = raw.iloc[rl_header_row + 1:, rl_header_col:rl_header_col + 10].copy()
    rl.columns = rl_cols
    rl = rl.dropna(how="all").reset_index(drop=True)
    rl["Year"] = rl["Year"].astype(int)

    recomputed = rl.groupby(["Year", "Arm"]).size().unstack(fill_value=0)
    mismatches = []
    for _, row in annual.iterrows():
        yr = row["Year"]
        if yr not in recomputed.index:
            continue
        rl_relax = recomputed.loc[yr].get("Relaxer/alkali", 0)
        rl_ker = recomputed.loc[yr].get("Keratin/blowout", 0)
        if rl_relax != row["Relaxer/alkali reports"] or rl_ker != row["Keratin/blowout reports"]:
            mismatches.append((yr, rl_relax, row["Relaxer/alkali reports"], rl_ker, row["Keratin/blowout reports"]))
    if mismatches:
        print("WARNING: annual counts recomputed from report-level rows do not match "
              "the sheet's pre-aggregated annual table for year(s):")
        for yr, rr, ar, rk, ak in mismatches:
            print(f"  {yr}: report-level relaxer/alkali={rr} vs annual table={ar}; "
                  f"report-level keratin/blowout={rk} vs annual table={ak}")
        print("  -> eFigure 2 uses the pre-aggregated annual table for counts/alopecia "
              "and report-level rows only for the other-4-reactions panel. Reconcile "
              "before trusting both panels equally.")
    else:
        print("Temporal_Trend cross-check OK: report-level and pre-aggregated annual "
              "counts agree for all years.")

    return annual, rl


# ----------------------------------------------------------------------------
# Table builders
# ----------------------------------------------------------------------------

def build_table1_cohorts(cohorts_df, outdir):
    df = cohorts_df.copy()
    df["Alkali retained %"] = df["Alkali retained %"].map(lambda x: f"{x:.1f}%")
    df["Keratin retained %"] = df["Keratin retained %"].map(lambda x: f"{x:.1f}%")
    df = df.rename(columns={
        "Alkali N": "Alkali Arm, No.",
        "Keratin/blowout N": "Keratin/Blowout Arm, No.",
        "Alkali retained %": "Alkali Retained, %",
        "Keratin retained %": "Keratin/Blowout Retained, %",
    })
    _save_table(df, outdir, "table1_cohort_definitions")
    return df


def build_table2_significant(primary_df, outdir):
    df = primary_df.copy()
    df = df[df["FDR significant"] == "Yes"].sort_values("Fisher p").reset_index(drop=True)
    out = pd.DataFrame({
        "Reaction": df["Reaction"].map(display_reaction),
        "Alkali, No./Total (%)": [format_pct(n, n + nr) for n, nr in
                                   zip(df["Alkali n"], df["Alkali no reaction"])],
        "Keratin/Blowout, No./Total (%)": [format_pct(n, n + nr) for n, nr in
                                            zip(df["Keratin/blowout n"], df["Keratin/blowout no reaction"])],
        "ROR (95% CI)": [format_ror_ci(r, lo, hi) for r, lo, hi in
                          zip(df["ROR (alkali vs keratin/blowout)"], df["95% CI low"], df["95% CI high"])],
        "P Value": df["Fisher p"].map(format_p),
        "FDR q": df["BH-FDR q"].map(format_p),
        "Direction": df["Direction"],
        "CI Method": df["CI method"],
    })
    _save_table(out, outdir, "table2_significant_findings")
    return out


_T3_ARMS_IN_ORDER = ["Relaxer/alkali", "Keratin/blowout"]

_T3_OUTCOME_ROWS = [
    ("Visited a Health Care Provider", "Visited a health care provider"),
    ("Visited Emergency Room", "Visited emergency room"),
    ("Hospitalization", "Hospitalization"),
    ("Disability", "Disability"),
    ("Life Threatening", "Life-threatening"),
    ("Death", "Death"),
]
_T3_ADDITIONAL_OUTCOME_ROWS = [
    ("Required Intervention", "Required intervention"),
    ("Other Serious or Important Medical Event", "Other serious or important medical event"),
]


def build_table3_demographics(report_characteristics_df, outdir, table1_summary_lookup=None):
    """Table 3 ("Table I. Characteristics of Cosmetic Adverse Event Reports by Product
    Cohort" in the manuscript): sex, age, and reported-clinical-outcome breakdown by
    exposure cohort.

    Built directly from the per-report Report_Characteristics sheet (one row per
    report, Gender/Age/outcome-flag columns) rather than copied from the workbook's
    own pre-built Table1_Summary sheet, so the published numbers can be regenerated
    from source rows on every run instead of trusting a cell range someone may have
    hand-typed or hand-edited once. If Table1_Summary is present, the recomputed
    values are cross-checked against it below (same spirit as the Temporal_Trend
    annual/report-level check above); a mismatch is printed as a warning, not a
    correction -- it means the two disagree and need a human to reconcile.

    "No. (%)" cells here use the arm's own N as the denominator (matching Table
    1_Summary's format, e.g. "71 (93.4%)") -- unlike format_pct() used elsewhere in
    this script, which prints an explicit "n/N" numerator/denominator. Table 3's
    denominator is already stated once in each column header ("(n = 76)"), so
    repeating it in every cell would be redundant here.
    """
    df = report_characteristics_df
    present_arms = [a for a in _T3_ARMS_IN_ORDER if a in set(df["Arm"])]
    missing_arms = [a for a in _T3_ARMS_IN_ORDER if a not in present_arms]
    if missing_arms:
        print(f"[table3] WARNING: expected arm(s) not found in Report_Characteristics, "
              f"omitted from the table: {missing_arms}")

    unexpected_gender = sorted(set(df["Gender"].dropna().astype(str)) - {"Female", "Male"})
    if unexpected_gender:
        print(f"[table3] WARNING: Gender value(s) other than 'Female'/'Male' found: "
              f"{unexpected_gender}. These are folded into 'Missing/unknown' below "
              f"(computed as a residual, not just blank/NaN cells).")

    avail = df[df["Age available"] == "Yes"]
    bad_units = sorted(set(avail["Age unit"].dropna().astype(str)) - {"year(s)"})
    if bad_units:
        print(f"[table3] WARNING: Age unit(s) other than 'year(s)' found among reports "
              f"with Age available = Yes: {bad_units}. The Age column is treated as "
              f"years regardless of unit below -- reconcile before trusting the Age row.")

    arm_slices = {a: df[df["Arm"] == a] for a in present_arms}
    ns = {a: len(arm_slices[a]) for a in present_arms}

    def pct(n, arm):
        N = ns[arm]
        return "NA" if N == 0 else f"{int(n)} ({100 * n / N:.1f}%)"

    def age_iqr(arm):
        ages = arm_slices[arm].loc[arm_slices[arm]["Age available"] == "Yes", "Age"]
        if ages.empty:
            return "NA"
        med, q1, q3 = ages.median(), ages.quantile(0.25), ages.quantile(0.75)
        return f"{med:.1f} ({q1:.1f}-{q3:.1f})"

    header = ["Characteristic"] + [f"{a} (n = {ns[a]})" for a in present_arms]
    rows = [["Reports, No."] + [ns[a] for a in present_arms]]

    rows.append(["Sex"] + [""] * len(present_arms))
    rows.append(["Female"] + [pct((arm_slices[a]["Gender"] == "Female").sum(), a) for a in present_arms])
    rows.append(["Male"] + [pct((arm_slices[a]["Gender"] == "Male").sum(), a) for a in present_arms])
    # Residual (N - Female - Male), not isna() alone -- catches any Gender value other
    # than "Female"/"Male" too (see unexpected_gender check above), not just blanks.
    rows.append(["Missing/unknown"] + [
        pct(ns[a] - (arm_slices[a]["Gender"] == "Female").sum() - (arm_slices[a]["Gender"] == "Male").sum(), a)
        for a in present_arms
    ])

    rows.append(["Age"] + [""] * len(present_arms))
    rows.append(["Age available, No. (%)"] + [
        pct((arm_slices[a]["Age available"] == "Yes").sum(), a) for a in present_arms
    ])
    rows.append(["Age, median (IQR), y"] + [age_iqr(a) for a in present_arms])

    rows.append(["Reported clinical outcomes"] + [""] * len(present_arms))
    for col, label in _T3_OUTCOME_ROWS:
        rows.append([label] + [pct((arm_slices[a][col] == "Yes").sum(), a) for a in present_arms])

    rows.append(["Additional FDA outcome categories (optional)"] + [""] * len(present_arms))
    for col, label in _T3_ADDITIONAL_OUTCOME_ROWS:
        rows.append([label] + [pct((arm_slices[a][col] == "Yes").sum(), a) for a in present_arms])

    out = pd.DataFrame(rows, columns=header)
    _save_table(out, outdir, "table3_report_characteristics", section_header_labels={
        "Sex", "Age", "Reported clinical outcomes",
        "Additional FDA outcome categories (optional)",
    })

    if table1_summary_lookup is not None:
        _cross_check_table3(rows, present_arms, table1_summary_lookup)

    return out


def _cross_check_table3(computed_rows, arms, wb_lookup):
    """Compares each computed table3 data row (label + one 'No. (%)' string per arm)
    against the workbook's own Table1_Summary sheet on matching label text.
    Section-header rows (all-blank values) are skipped since Table1_Summary has no
    values to compare there either. A consistency check only -- a mismatch means the
    recomputed table and the workbook's own summary disagree and need a human to
    reconcile, not that this recomputation is automatically the correct one."""
    mismatches = []
    for row in computed_rows:
        label, *vals = row
        if all(v == "" for v in vals):
            continue
        wb_vals = wb_lookup.get(label)
        if wb_vals is None:
            continue
        for arm, computed, wb in zip(arms, vals, wb_vals):
            computed_s, wb_s = str(computed).strip(), str(wb).strip()
            if label == "Reports, No.":
                # Table1_Summary stores this row as a bare number (possibly read back
                # as a float, e.g. "76.0"); table3 stores it as a Python int. Compare
                # on the integer value rather than the literal string.
                wb_s = wb_s.split(".")[0]
            if computed_s != wb_s:
                mismatches.append((label, arm, computed_s, wb_s))
    if mismatches:
        print("[table3] WARNING: recomputed values differ from the workbook's own "
              "Table1_Summary sheet -- reconcile before trusting "
              "table3_report_characteristics.csv:")
        for label, arm, comp, wb in mismatches:
            print(f"  {label} / {arm}: recomputed={comp!r} vs Table1_Summary={wb!r}")
    else:
        print("Table1_Summary cross-check OK: recomputed report-characteristics values "
              "match the workbook's own summary sheet for all rows compared.")


def _build_ror_results_table(df, outdir, filename, normalize_reaction_display=True):
    """Shared formatter for a full (all-reactions) ROR results table. Used for both
    eTable1 (Primary_ROR, case-normalized before this stage) and eTable3, the
    exact-case sensitivity table (ExactCase_Sensitivity, capitalization variants left
    unmerged on purpose).

    normalize_reaction_display=False is required for the exact-case table: its whole
    point is to show "Alopecia" and "ALOPECIA" as distinct rows. Running them through
    display_reaction() would title-case both to "Alopecia" and silently erase the
    thing the table exists to demonstrate."""
    df = df.copy().sort_values("Fisher p").reset_index(drop=True)
    reaction_col = (df["Reaction"].map(display_reaction) if normalize_reaction_display
                     else df["Reaction"].astype(str))
    out = pd.DataFrame({
        "Reaction": reaction_col,
        "Alkali, No./Total (%)": [format_pct(n, n + nr) for n, nr in
                                   zip(df["Alkali n"], df["Alkali no reaction"])],
        "Keratin/Blowout, No./Total (%)": [format_pct(n, n + nr) for n, nr in
                                            zip(df["Keratin/blowout n"], df["Keratin/blowout no reaction"])],
        "ROR (95% CI)": [format_ror_ci(r, lo, hi) for r, lo, hi in
                          zip(df["ROR (alkali vs keratin/blowout)"], df["95% CI low"], df["95% CI high"])],
        "P Value": df["Fisher p"].map(format_p),
        "FDR q": df["BH-FDR q"].map(format_p),
        "FDR Significant": df["FDR significant"],
        "CI Method": df["CI method"],
        "Direction": df["Direction"],
    })
    _save_table(out, outdir, filename)
    return out


def build_etable1_full(primary_df, outdir):
    return _build_ror_results_table(primary_df, outdir, "etable1_full_results")


def build_etable3_exact_case_sensitivity(exact_case_df, outdir):
    """Supplemental Table 2 (file: eTable3, after eTable1/eTable2 -- see note in
    main()): the same primary-analysis ROR calculation, but with reaction-term
    capitalization variants left unmerged (e.g. "Alopecia" and "ALOPECIA" as separate
    rows) instead of case-normalized. Tests whether the conclusions depend on that
    normalization choice; does not replace the planned product-level sensitivity
    analyses (SA1-SA4)."""
    return _build_ror_results_table(
        exact_case_df, outdir, "etable3_sensitivity_exact_case_capitalization",
        normalize_reaction_display=False,
    )


def build_etable2_robustness(robustness_df, outdir):
    df = robustness_df.copy().sort_values(
        ["Significant analyses (0-5)", "Primary q"], ascending=[False, True]
    ).reset_index(drop=True)

    def ror_q(ror, q, sig):
        core = f"{ror:.2f}" if pd.notna(ror) else "NA"
        flag = "*" if sig == "Yes" else ""
        return f"{core} (q={format_p(q)}){flag}"

    out = pd.DataFrame({
        "Reaction": df["Reaction"].map(display_reaction),
        "Primary": [ror_q(r, q, s) for r, q, s in zip(df["Primary ROR"], df["Primary q"], df["Primary sig"])],
        "SA1 (1 Cosmetic)": [ror_q(r, q, s) for r, q, s in zip(df["SA1 ROR"], df["SA1 q"], df["SA1 sig"])],
        "SA2 (1 Suspect)": [ror_q(r, q, s) for r, q, s in zip(df["SA2 ROR"], df["SA2 q"], df["SA2 sig"])],
        "SA3 (Strict Name)": [ror_q(r, q, s) for r, q, s in zip(df["SA3 ROR"], df["SA3 q"], df["SA3 sig"])],
        "SA4 (Strict + 1 Suspect)": [ror_q(r, q, s) for r, q, s in zip(df["SA4 ROR"], df["SA4 q"], df["SA4 sig"])],
        "Significant Analyses (of 5)": df["Significant analyses (0-5)"],
        "Direction": df["Primary direction"],
        "Robustness Note": df["Robustness note"],
    })
    _save_table(out, outdir, "etable2_robustness_summary")
    return out


_ARM_DISPLAY = {"Alkali": "Relaxer/alkali", "Keratin/blowout": "Keratin/blowout"}


def build_etable4_product_validation(report_level_qc_df, strict_excluded_df, outdir):
    """Supplemental Table 3 (file: eTable4): per-report record of which product name
    qualified a report for the strict-name (SA3) cohort, or -- for the reports SA3's
    conservative rule dropped -- which candidate product was considered and why it
    didn't qualify.

    Built by joining Report_Level_QC (has the SA3 keep/drop decision, and for kept
    reports the qualifying product) against Strict_Excluded_Reports (has the drop
    reason for every excluded report) on Report number. Confirmed 1:1: every report
    with "SA3 keep" == "No" in Report_Level_QC has exactly one matching row in
    Strict_Excluded_Reports, and vice versa."""
    reasons = strict_excluded_df.set_index("Report number")

    rows = []
    for _, r in report_level_qc_df.iterrows():
        rn = r["Report number"]
        included = r["SA3 keep"] == "Yes"
        if included:
            product = r["Strict included product(s)"]
            classification = "Included, strict-name cohort (SA3)"
            reason = ""
        else:
            exrow = reasons.loc[rn]
            product = exrow["Suspect code-53 product(s)"]
            if pd.isna(product) or str(product).strip() == "":
                # A handful of exclusions have no product to name at all: the
                # "keratin"/arm keyword hit came from a code-53 product listed only
                # as a CONCOMITANT (not the SUSPECT), or the actual suspect product
                # isn't code-53 in the first place. Leaving the cell blank there
                # reads as a data gap; label it explicitly instead.
                product = "No code-53 suspect product identified"
            classification = "Excluded, conservative name-based review"
            reason = exrow["Reason for conservative exclusion"]
        rows.append({
            "Report Number": rn,
            "Primary Cohort": _ARM_DISPLAY.get(r["Arm"], r["Arm"]),
            "Product Name Triggering Inclusion/Consideration": product,
            "Classification": classification,
            "Reason for Exclusion From Strict-Name Cohort": reason,
        })
    out = pd.DataFrame(rows).sort_values(
        ["Primary Cohort", "Report Number"]).reset_index(drop=True)
    _save_table(out, outdir, "etable4_product_name_validation")
    return out


def _save_table(df, outdir, name, section_header_labels=None):
    """section_header_labels: values of the first column that mark a section-divider
    row (e.g. "Sex", "Age" in table3) rather than a data row. Those rows' 2nd/3rd+
    columns are legitimately blank (there's no value to report), but left as plain
    unstyled empty cells they're visually indistinguishable from an accidentally-empty
    data row -- so they're bolded and shaded here to read as an intentional divider,
    matching how a typeset "Table 1" usually marks these. CSV has no styling, so this
    only affects the .xlsx; the .csv still has literally-empty cells there, which is
    normal/expected for a flat text format."""
    csv_path = os.path.join(outdir, f"{name}.csv")
    xlsx_path = os.path.join(outdir, f"{name}.xlsx")
    df.to_csv(csv_path, index=False)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Table")
        ws = writer.sheets["Table"]
        for col_cells in ws.columns:
            length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 55)
        section_fill = PatternFill(start_color="E4E4E4", end_color="E4E4E4", fill_type="solid")
        for row in ws.iter_rows():
            is_section_header = (section_header_labels is not None
                                  and row[0].value in section_header_labels)
            for cell in row:
                cell.font = Font(name="Arial", size=10, bold=(is_section_header or cell.font.bold))
                if is_section_header:
                    cell.fill = section_fill


EFIG2_RASTER_DPI = 150       # resolution used only for eFigure 2's PDF/EPS, whose
                              # raw-point scatter is now rasterized=True (see below)
                              # to work around EPS having no transparency support.
                              # Scoped to just this one figure: the two heatmap
                              # figures (Figure 3, eFigure 1) are drawn with imshow,
                              # which embeds as a raster in PDF/EPS regardless, and
                              # an early version of this fix passed a high dpi to
                              # every figure -- which did nothing for those flat
                              # solid-color cells except bloat their EPS files
                              # roughly 30-40x (into the tens of MB) for no visible
                              # benefit. Leaving their dpi unset preserves their
                              # original (already-fine) file size.


def _savefig(fig, outpath_base, pad_inches=0.08, vector_dpi=None):
    for fmt in SAVE_FORMATS:
        kwargs = {"bbox_inches": "tight", "pad_inches": pad_inches}
        if fmt == "png":
            kwargs["dpi"] = DPI
        elif vector_dpi is not None:
            kwargs["dpi"] = vector_dpi
        fig.savefig(f"{outpath_base}.{fmt}", **kwargs)


# ----------------------------------------------------------------------------
# Figure 2: annotated forest plot, 10 Primary-significant reactions
# ----------------------------------------------------------------------------

FIG2_ORDER = [
    "ALOPECIA", "EYE IRRITATION", "DYSPNEA", "DIZZINESS", "NASAL DISCOMFORT",
    "NAUSEA", "COUGH", "TRICHORRHEXIS", "HEADACHE", "LACRIMATION INCREASED",
]
FIG2_ROBUST = {"ALOPECIA", "EYE IRRITATION", "DYSPNEA", "DIZZINESS", "NASAL DISCOMFORT"}


def make_figure2(primary_df, outpath_base):
    df = primary_df.copy()
    df["_key"] = df["Reaction"].map(canonical_key)
    lookup = {r["_key"]: r for _, r in df.iterrows()}

    missing = [k for k in FIG2_ORDER if k not in lookup]
    if missing:
        raise ValueError(f"Figure 2: reaction(s) not found in Primary_ROR: {missing}")

    rows = []
    nonzero = [lookup[k]["ROR (alkali vs keratin/blowout)"] for k in FIG2_ORDER
               if lookup[k]["ROR (alkali vs keratin/blowout)"] > 0]
    floor = min(nonzero) / 3 if nonzero else 0.01

    for key in FIG2_ORDER:
        r = lookup[key]
        ror, lo, hi = r["ROR (alkali vs keratin/blowout)"], r["95% CI low"], r["95% CI high"]
        zero_cell = (ror == 0)
        robust = key in FIG2_ROBUST
        rows.append({
            "reaction": display_reaction(key),
            "ror": ror, "lo": lo, "hi": hi,
            "zero_cell": zero_cell,
            "robust": robust,
            "label": format_ror_ci(ror, lo, hi) if not zero_cell
                     else f"\u2264{hi:.2f} (upper CI; ROR\u21920)",
        })
    rows = rows[::-1]
    n = len(rows)

    TOP_MARGIN_IN = 0.85
    BOTTOM_MARGIN_IN = 1.9
    ROW_HEIGHT_IN = 0.48
    fig_height = TOP_MARGIN_IN + max(1.2, ROW_HEIGHT_IN * n) + BOTTOM_MARGIN_IN
    fig, ax = plt.subplots(figsize=(FIG_WIDTH_IN, fig_height))

    for i, r in enumerate(rows):
        y = i
        if r["zero_cell"]:
            ax.plot([floor, r["hi"]], [y, y], color=INK_COLOR, lw=1.3, zorder=2)
            ax.plot([r["hi"], r["hi"]], [y - 0.12, y + 0.12], color=INK_COLOR, lw=1.3, zorder=2)
            ax.annotate("", xy=(floor * 0.7, y), xytext=(floor, y),
                        arrowprops=dict(arrowstyle="-|>", color=INK_COLOR, lw=1.3), zorder=2)
        else:
            ax.plot([r["lo"], r["hi"]], [y, y], color=INK_COLOR, lw=1.3, zorder=2)
            if r["robust"]:
                ax.plot(r["ror"], y, marker="o", markersize=6.5,
                        markerfacecolor=INK_COLOR, markeredgecolor=INK_COLOR, zorder=3)
            else:
                ax.plot(r["ror"], y, marker="o", markersize=6.5,
                        markerfacecolor="white", markeredgecolor=INK_COLOR,
                        markeredgewidth=1.4, zorder=3)

    ax.axvline(1.0, color=REF_LINE_COLOR, lw=1, linestyle="--", zorder=1)
    ax.set_yticks(range(n))
    ax.set_yticklabels([r["reaction"] for r in rows])
    for tick, r in zip(ax.get_yticklabels(), rows):
        tick.set_fontweight("bold" if r["robust"] else "normal")
    ax.set_ylim(-0.7, n - 0.3)
    ax.set_xscale("log")

    # Explicit xlim + ticks, replacing matplotlib's autoscale/default log formatter.
    # Two problems this fixes together:
    #   1. Autoscale alone left this axis spanning under one decade, which pushed
    #      matplotlib into labeling every minor tick (2x,3x,...9x) in scientific
    #      notation -- they overlapped into unreadable text (see PR notes).
    #   2. The zero-cell arrowhead is drawn via ax.annotate(), whose endpoints are
    #      NOT counted by matplotlib's autoscale. With autoscale alone, the
    #      arrowhead could land outside the visible xlim and be silently clipped --
    #      defeating the whole point of drawing it. Setting xlim explicitly with
    #      margin past `floor` guarantees the arrow stays on-canvas.
    xlim_left = floor * 0.55
    xlim_right = max(r["hi"] for r in rows) * 1.15
    ax.set_xlim(xlim_left, xlim_right)
    ticks = _nice_log_ticks(xlim_left, xlim_right)
    ax.set_xticks(ticks)
    ax.set_xticklabels([_fmt_axis_ratio(t) for t in ticks])
    ax.xaxis.set_minor_locator(matplotlib.ticker.NullLocator())

    wrapped_title = "\n".join(textwrap.wrap(
        "Reporting Odds Ratios for Primary-Significant Adverse Events", width=38))
    ax.set_title(wrapped_title, fontsize=10, fontweight="bold", loc="left")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    xmax = ax.get_xlim()[1]
    for i, r in enumerate(rows):
        ax.text(xmax * 1.05, i, r["label"], va="center", fontsize=7.3, color="#222222")

    fig.subplots_adjust(
        bottom=BOTTOM_MARGIN_IN / fig_height,
        top=1 - (TOP_MARGIN_IN / fig_height),
        left=0.30, right=0.66,
    )
    bfrac = BOTTOM_MARGIN_IN / fig_height
    xlabel_y = bfrac * 0.82
    direction_y = bfrac * 0.62
    legend_y = bfrac * 0.38
    footnote_y = bfrac * 0.10

    fig.text(0.5, xlabel_y, "Reporting Odds Ratio (log scale)", ha="center", va="center", fontsize=9.5)
    fig.text(0.03, direction_y, "\u2190 Higher reporting odds,\nkeratin/blowout products",
              ha="left", va="center", fontsize=7.5)
    fig.text(0.97, direction_y, "Higher reporting odds,\nrelaxer/alkali products \u2192",
              ha="right", va="center", fontsize=7.5)

    legend_elems = [
        Line2D([0], [0], color=INK_COLOR, marker="o", markerfacecolor=INK_COLOR,
               lw=1.3, markersize=6.5, label="Significant, Primary + all 4 sensitivity analyses"),
        Line2D([0], [0], color=INK_COLOR, marker="o", markerfacecolor="white",
               markeredgewidth=1.4, lw=1.3, markersize=6.5, label="Significant, Primary analysis only"),
    ]
    fig.legend(handles=legend_elems, loc="center", bbox_to_anchor=(0.5, legend_y),
               fontsize=7.3, frameon=False, ncol=1)

    fig.text(0.5, footnote_y,
              "Dashed line indicates ROR = 1. For zero-cell reactions (alkali-arm count = 0), the 95% CI is\n"
              "shown extending toward zero (arrow); no point estimate is plotted.",
              ha="center", va="center", fontsize=6.5, color="#555555")

    _savefig(fig, outpath_base)
    plt.close(fig)


# ----------------------------------------------------------------------------
# Figure 3 / eFigure 1: robustness heatmaps
# ----------------------------------------------------------------------------

FIG3_ORDER = [
    "EYE IRRITATION", "ALOPECIA", "DYSPNEA", "DIZZINESS", "NASAL DISCOMFORT",
    "NAUSEA", "COUGH", "TRICHORRHEXIS", "HEADACHE", "LACRIMATION INCREASED",
    "CHEMICAL INJURY",
]


def _heatmap(reactions_ordered, robustness_df, title, outpath_base, fig_width=None,
             max_height_in=10.0, default_row_height_in=0.42, min_row_height_in=0.16):
    df = robustness_df.copy()
    df["_key"] = df["Reaction"].map(canonical_key)
    lookup = {row["_key"]: row for _, row in df.iterrows()}
    missing = [k for k in reactions_ordered if k not in lookup]
    if missing:
        raise ValueError(f"Heatmap: reaction(s) not found in Robustness_Summary: {missing}")

    analyses = ["Primary sig", "SA1 sig", "SA2 sig", "SA3 sig", "SA4 sig"]
    col_labels = ["Primary", "SA1\n(1 Cosmetic)", "SA2\n(1 Suspect)", "SA3\n(Strict Name)",
                  "SA4\n(Strict +\n1 Suspect)"]
    mat = np.array([[1 if lookup[k][a] == "Yes" else 0 for a in analyses] for k in reactions_ordered])
    labels = [display_reaction(k) for k in reactions_ordered]
    n = len(labels)

    width = fig_width or FIG_WIDTH_IN
    top_bottom_margin_in = 1.3
    # Scale row height down for many-row heatmaps so a 45-row supplement figure doesn't
    # end up printing at ~20 in tall -- cap total height, floor row height for legibility.
    row_height_in = min(default_row_height_in, max(min_row_height_in,
                                                     (max_height_in - top_bottom_margin_in) / n))
    fig_height = row_height_in * n + top_bottom_margin_in
    label_fontsize = 8 if row_height_in >= 0.28 else (7 if row_height_in >= 0.20 else 6)

    fig, ax = plt.subplots(figsize=(width, fig_height))

    cmap = matplotlib.colors.ListedColormap([NOT_SIG_FILL, ROBUST_FILL])
    ax.imshow(mat, aspect="auto", cmap=cmap, vmin=0, vmax=1)
    ax.set_xticks(range(len(col_labels)))
    ax.set_xticklabels(col_labels, fontsize=7.5)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=label_fontsize)
    ax.set_xticks(np.arange(-0.5, len(col_labels), 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(labels), 1), minor=True)
    ax.grid(which="minor", color="white", linewidth=1.2 if row_height_in < 0.28 else 1.5)
    ax.tick_params(which="minor", bottom=False, left=False)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_title(title, fontsize=9.5, fontweight="bold")

    fig.tight_layout()
    _savefig(fig, outpath_base, pad_inches=0.1)
    plt.close(fig)


def make_figure3(robustness_df, outpath_base):
    _heatmap(
        FIG3_ORDER, robustness_df,
        "Robustness of Adverse Event Signals Across Exposure Definitions\n"
        "(teal = FDR-significant, q<.05; gray = not significant)",
        outpath_base,
    )


def make_efigure1(robustness_df, outpath_base):
    df = robustness_df.copy().sort_values(
        ["Significant analyses (0-5)", "Primary q"], ascending=[False, True]
    )
    ordered_keys = df["Reaction"].map(canonical_key).tolist()
    _heatmap(
        ordered_keys, robustness_df,
        "Robustness Across All 45 Prespecified Reactions\n"
        "(teal = FDR-significant, q<.05; gray = not significant)",
        outpath_base, fig_width=FIG_WIDTH_IN,
    )


# ----------------------------------------------------------------------------
# eFigure 2: temporal trends
# ----------------------------------------------------------------------------

OTHER_ROBUST_KERATIN_REACTIONS = ["Eye irritation", "Dyspnoea", "Dizziness", "Nasal discomfort"]
BREAKPOINT_YEAR_FRAC = 2022 + 9.5 / 12


def _rolling3(series):
    return series.rolling(window=3, min_periods=1, center=False).mean()


def _other_reactions_rolling_pct(reportlevel_df, years):
    """Per-year raw % and 3-year rolling % of keratin/blowout reports mentioning each
    of OTHER_ROBUST_KERATIN_REACTIONS. Shared by eFigure 2 Panel C and eTable5 (the
    numeric table behind it) so the plotted lines and the tabulated numbers can never
    drift apart. Both percentages here are already on a 0-100 scale (unlike the
    alopecia percentages loaded straight from the workbook in make_efigure2_temporal,
    which are 0-1 fractions until scaled there -- callers formatting these two sets of
    numbers need different treatment; see build_etable5_temporal_numerical)."""
    ker_rl = reportlevel_df[reportlevel_df["Arm"] == "Keratin/blowout"].copy()
    out = {}
    for reaction in OTHER_ROBUST_KERATIN_REACTIONS:
        grp = ker_rl.groupby("Year").agg(n=(reaction, "size"), pos=(reaction, "sum"))
        grp = grp.reindex(years, fill_value=0)
        # Years with zero keratin/blowout reports get NaN here (0/0), not 0% -- a year
        # with no reports at all is missing data, not evidence of no reaction. NaN is
        # left in place going into the rolling mean, which correctly averages over
        # whatever non-missing years fall in each 3-year window (verified: pandas
        # .rolling().mean() skips NaN rather than propagating it).
        pct = (grp["pos"] / grp["n"].replace(0, np.nan) * 100)
        out[reaction] = {"pct": pct.reset_index(drop=True), "pct_roll": _rolling3(pct).reset_index(drop=True)}
    return out


def make_efigure2_temporal(annual_df, reportlevel_df, outpath_base):
    years = annual_df["Year"].to_numpy()

    fig, axes = plt.subplots(3, 1, figsize=(FIG_WIDTH_IN, 8.0), sharex=True)

    ax = axes[0]
    ax.plot(years, annual_df["Relaxer/alkali reports"], color=ARM_A_COLOR, marker="o",
            markersize=3, lw=1.3, label="Relaxer/alkali")
    ax.plot(years, annual_df["Keratin/blowout reports"], color=ARM_B_COLOR, marker="o",
            markersize=3, lw=1.3, label="Keratin/blowout")
    ax.set_ylabel("Reports, No.", fontsize=8.5)
    ax.set_title("A. Annual Report Counts by Cohort", fontsize=9, fontweight="bold", loc="left")
    ax.legend(fontsize=7, frameon=False, loc="upper left")
    ax.spines[["top", "right"]].set_visible(False)

    ax = axes[1]
    relax_pct = annual_df["Relaxer/alkali alopecia, %"].astype(float) * 100
    ker_pct = annual_df["Keratin/blowout alopecia, %"].astype(float) * 100
    relax_roll = annual_df["3-y rolling relaxer/alkali alopecia, %"].astype(float) * 100
    ker_roll = annual_df["3-y rolling keratin/blowout alopecia, %"].astype(float) * 100
    # Raw annual proportions are shown as unconnected points, not a connected line --
    # with some years having only 1-3 reports, connecting them draws the eye to
    # 0%-to-100% jumps that are sampling noise, not a trend. The rolling line (below)
    # is the intended trend signal; the points are there for transparency about spread.
    # rasterized=True: the EPS format has no way to represent alpha transparency at
    # all -- matplotlib would otherwise silently render these points fully opaque in
    # the .eps output only, so this panel would look different (darker, denser)
    # there than in the .png/.pdf versions of the same figure. Rasterizing just
    # these two artists embeds them as a high-res bitmap patch inside the still-
    # vector EPS file, which preserves the intended translucent look everywhere.
    ax.scatter(years, relax_pct, color=ARM_A_COLOR, alpha=0.35, s=10, zorder=2, rasterized=True)
    ax.scatter(years, ker_pct, color=ARM_B_COLOR, alpha=0.35, s=10, zorder=2, rasterized=True)
    ax.plot(years, relax_roll, color=ARM_A_COLOR, lw=1.8, label="Relaxer/alkali (3-y rolling)")
    ax.plot(years, ker_roll, color=ARM_B_COLOR, lw=1.8, label="Keratin/blowout (3-y rolling)")
    ax.set_ylabel("Reports With\nAlopecia, %", fontsize=8.5)
    ax.set_title("B. Alopecia Reporting Proportion by Cohort", fontsize=9, fontweight="bold", loc="left")
    ax.legend(fontsize=6.8, frameon=False, loc="upper left")
    ax.spines[["top", "right"]].set_visible(False)

    ax = axes[2]
    other_reactions = _other_reactions_rolling_pct(reportlevel_df, years)
    palette = ["#2C6E7F", "#6C8EA6", "#8FA6A3", "#B7C9C2"]
    for reaction, color in zip(OTHER_ROBUST_KERATIN_REACTIONS, palette):
        pct_roll = other_reactions[reaction]["pct_roll"]
        ax.plot(years, pct_roll, color=color, lw=1.6,
                label=f"{display_reaction(reaction)} (3-y rolling)")
    ax.set_ylabel("Reports With\nReaction, %", fontsize=8.5)
    ax.set_xlabel("Year", fontsize=9)
    ax.set_title("C. Other Fully-Robust Keratin/Blowout-Associated Reactions",
                  fontsize=9, fontweight="bold", loc="left")
    ax.legend(fontsize=6.5, frameon=False, loc="upper left", ncol=1)
    ax.spines[["top", "right"]].set_visible(False)

    for ax in axes:
        #ax.axvline(BREAKPOINT_YEAR_FRAC, color="#888888", lw=1, linestyle=":")
        ax.xaxis.set_major_locator(matplotlib.ticker.MaxNLocator(integer=True, nbins=8))
    #axes[0].text(BREAKPOINT_YEAR_FRAC, axes[0].get_ylim()[1] * 0.92,
                 #" Oct 2022\n Sister Study\n coverage", fontsize=6, color="#555555", va="top")

    #fig.suptitle("Temporal Trends in Reporting, 2003-2023", fontsize=10,
                  #fontweight="bold", x=0.02, ha="left")
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    _savefig(fig, outpath_base, pad_inches=0.1, vector_dpi=EFIG2_RASTER_DPI)
    plt.close(fig)


def _frac_pct1(x):
    """Format a 0-1 fraction (the alopecia % columns as stored in the workbook) as a
    1-decimal percentage string."""
    return "NA" if pd.isna(x) else f"{float(x) * 100:.1f}"


def _pct1(x):
    """Format an already-0-100-scale percentage (from _other_reactions_rolling_pct,
    which computes it directly in Python) as a 1-decimal string. Do not multiply by
    100 again -- see the scale note on that function."""
    return "NA" if pd.isna(x) else f"{float(x):.1f}"


def build_etable5_temporal_numerical(annual_df, reportlevel_df, outdir):
    """Supplemental Table 4: the numeric data behind every panel of eFigure 2 (annual
    report counts, alopecia n/N and rolling %, and rolling % for the other 4 fully-
    robust keratin/blowout reactions), so a reader can check the underlying counts
    without re-deriving them from the plot. Reuses annual_df as already parsed for
    the figure, and the same _other_reactions_rolling_pct() the figure's Panel C
    plots -- the Panel C reaction columns here are guaranteed to match the figure
    exactly, not just approximately."""
    years = annual_df["Year"].to_numpy()
    other = _other_reactions_rolling_pct(reportlevel_df, years)

    out = pd.DataFrame({
        "Year": annual_df["Year"].astype(int),
        "Relaxer/Alkali Reports, No.": annual_df["Relaxer/alkali reports"].astype(int),
        "Keratin/Blowout Reports, No.": annual_df["Keratin/blowout reports"].astype(int),
        "Combined Reports, No.": annual_df["Combined reports"].astype(int),
        "Relaxer/Alkali Alopecia, No./Total (%)": [
            format_pct(n, N) for n, N in zip(annual_df["Relaxer/alkali alopecia, n"],
                                              annual_df["Relaxer/alkali reports"])
        ],
        "Keratin/Blowout Alopecia, No./Total (%)": [
            format_pct(n, N) for n, N in zip(annual_df["Keratin/blowout alopecia, n"],
                                              annual_df["Keratin/blowout reports"])
        ],
        "Relaxer/Alkali Alopecia, 3-y Rolling %": [
            _frac_pct1(x) for x in annual_df["3-y rolling relaxer/alkali alopecia, %"]
        ],
        "Keratin/Blowout Alopecia, 3-y Rolling %": [
            _frac_pct1(x) for x in annual_df["3-y rolling keratin/blowout alopecia, %"]
        ],
    })
    for reaction in OTHER_ROBUST_KERATIN_REACTIONS:
        colname = f"{display_reaction(reaction)} (Keratin/Blowout), 3-y Rolling %"
        out[colname] = [_pct1(x) for x in other[reaction]["pct_roll"]]

    _save_table(out, outdir, "etable5_temporal_numerical_data")
    return out


# ----------------------------------------------------------------------------
# Figure legends
# ----------------------------------------------------------------------------

def write_figure_legends(path):
    text = """\
# Figure Legends (draft -- AMA style)

Only Figure 3's caption text below was supplied verbatim; the others are drafted to
match its level of detail and should be reviewed, not assumed final.

## Figure 2. Reporting Odds Ratios for Primary-Significant Adverse Events

Forest plot of reporting odds ratios (RORs) with 95% CIs for the 10 reactions reaching
FDR-adjusted significance (q<.05) in the Primary analysis (alkali [relaxer/lye/perm] vs
keratin/blowout products, restricted to reports coded to FDA cosmetics [industry code
53]). Filled markers and bold labels indicate reactions significant in the Primary
analysis and all 4 sensitivity analyses (SA1-SA4); open markers and regular labels
indicate reactions significant in the Primary analysis only. The dashed vertical line
indicates ROR = 1 (no disproportionate reporting). For reactions with zero alkali-arm
reports, the reporting odds ratio is not estimable; the 95% CI is shown extending toward
zero, with no point estimate plotted.

## Figure 3. Robustness of Adverse Event Signals Across Exposure Definitions

The figure includes all reactions that reached the FDR-adjusted significance threshold
in at least 1 analysis. The remaining 34 prespecified reactions did not reach q<.05 in
any analysis and are provided in the Supplement.

## eFigure 1. Robustness Across All 45 Prespecified Reactions

Reaction x analysis significance matrix for all 45 prespecified MedDRA reaction terms
tested, across the Primary analysis and sensitivity analyses SA1-SA4. Teal indicates
FDR-adjusted significance (q<.05); gray indicates non-significance. Reactions are
ordered by number of analyses reaching significance (descending), then by Primary
analysis FDR q value (ascending).

## eFigure 2. Temporal Trends in Reporting, 2003-2023

(A) Annual counts of adverse event reports by cohort. (B) Annual proportion of reports
describing alopecia, by cohort; light points show raw annual proportions, bold lines
show 3-year rolling averages. (C) Three-year rolling proportions of the other 4
reactions significant across the Primary analysis and all 4 sensitivity analyses in the
keratin/blowout arm (eye irritation, dyspnea, dizziness, nasal discomfort). The dotted
vertical line marks October 17, 2022, the publication date of the NIH-communicated
Sister Study analysis linking hair-straightening product use to uterine cancer risk,
shown to assess whether reporting patterns shifted coincident with related media
coverage. Panel A and the alopecia series in Panel B are aggregated directly from the
underlying report-level data as provided in the workbook, shown here as reported
without independent re-derivation beyond the internal consistency check described in
Methods.
"""
    with open(path, "w") as f:
        f.write(text)


def write_table_legends(path):
    text = """\
# Table Legends (draft -- AMA style)

Titles are as already used (Table 1 in Table1_Summary, Table 2 in the manuscript).
Footnotes are drafted and should be reviewed, not assumed final.

## Table 1. Characteristics of Cosmetic Adverse Event Reports by Product Cohort

No. (%) unless noted; denominator is each cohort's total reports. Age statistics
among reports with age data available. Outcome categories are not mutually
exclusive. Source: FDA CAERS cosmetics reports (industry code 53) via openFDA,
2003-2023 (76 relaxer/alkali; 276 keratin/blowout).

## Table 2. Adverse Events With Significantly Different Reporting Odds Between Product Cohorts

The 10 reactions reaching Benjamini-Hochberg FDR-adjusted significance (q<.05) in
the Primary analysis, ordered by ascending Fisher exact P value. ROR indicates
reporting odds ratio; CI, confidence interval, estimated by the Woolf logit method
except where a zero cell required the exact conditional method (see CI Method).
Alkali indicates relaxer/lye/perm products; Keratin/Blowout, keratin- or
formaldehyde-based smoothing products. For reactions with zero alkali-arm reports,
the reporting odds ratio is not estimable; only the upper 95% CI bound is shown,
consistent with Figure 2.
"""
    with open(path, "w") as f:
        f.write(text)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def main():
    args = parse_args()
    xlsx_path = args.xlsx_path

    # Save outputs next to the Excel workbook by default. This avoids trying to
    # create "pub_outputs" in the terminal's current directory (for example "/"),
    # which is read-only on modern macOS. Pass output_dir explicitly to override.
    outdir = args.output_dir or os.path.join(
        os.path.dirname(os.path.abspath(xlsx_path)), "pub_outputs"
    )

    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(
            f"Could not find the Excel workbook at:\n{xlsx_path}\n\n"
            "Check that the path and filename are exact (relative paths are "
            "resolved from the current working directory)."
        )
    fig_dir = os.path.join(outdir, "figures")
    tab_dir = os.path.join(outdir, "tables")
    os.makedirs(fig_dir, exist_ok=True)
    os.makedirs(tab_dir, exist_ok=True)

    data = load_workbook(xlsx_path)
    primary = data["primary"]
    robustness = data["robustness"]
    cohorts = data["cohorts"]

    build_table1_cohorts(cohorts, tab_dir)
    build_table2_significant(primary, tab_dir)
    build_etable1_full(primary, tab_dir)
    build_etable2_robustness(robustness, tab_dir)

    if data["exact_case"] is not None:
        build_etable3_exact_case_sensitivity(data["exact_case"], tab_dir)
    else:
        print("No ExactCase_Sensitivity sheet found -- skipping eTable3 "
              "(exact-case capitalization sensitivity).")

    if data["report_level_qc"] is not None and data["strict_excluded"] is not None:
        build_etable4_product_validation(data["report_level_qc"], data["strict_excluded"], tab_dir)
    else:
        print("No Report_Level_QC / Strict_Excluded_Reports sheet found -- skipping "
              "eTable4 (product-name validation).")

    if data["report_characteristics"] is not None:
        build_table3_demographics(
            data["report_characteristics"], tab_dir,
            table1_summary_lookup=data["table1_summary_lookup"],
        )
    else:
        print("No Report_Characteristics sheet found -- skipping table3 (report "
              "characteristics / demographics).")

    make_figure2(primary, os.path.join(fig_dir, "fig2_forest_primary_significant"))
    make_figure3(robustness, os.path.join(fig_dir, "fig3_robustness_heatmap_11"))
    make_efigure1(robustness, os.path.join(fig_dir, "efig1_robustness_heatmap_full45"))

    if data["temporal_annual"] is not None:
        make_efigure2_temporal(
            data["temporal_annual"], data["temporal_reportlevel"],
            os.path.join(fig_dir, "efig2_temporal_trends"),
        )
        build_etable5_temporal_numerical(
            data["temporal_annual"], data["temporal_reportlevel"], tab_dir
        )
    else:
        print("No Temporal_Trend sheet found -- skipping eFigure 2 and eTable5.")

    write_figure_legends(os.path.join(outdir, "figure_legends.md"))
    write_table_legends(os.path.join(outdir, "table_legends.md"))

    print(f"Done. Tables in {tab_dir}/, figures in {fig_dir}/, legends in {outdir}/figure_legends.md")


if __name__ == "__main__":
    main()
