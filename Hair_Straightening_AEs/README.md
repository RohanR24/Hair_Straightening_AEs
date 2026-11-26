# Adverse event reporting for relaxer/alkali and keratin/blowout hair-straightening products

Code, queries, and derived data for the study **"Comparative Adverse Event Reporting Profiles of Relaxer/Alkali and Keratin/Blowout Hair-Straightening Products: An FDA Disproportionality Analysis."**

The study compares adverse event reports involving relaxer/alkali products (76 reports) with reports involving keratin/blowout smoothing products (276 reports) in the US Food and Drug Administration's legacy cosmetic adverse event data (CAERS), from 2003 through 2023, before mandatory reporting under the Modernization of Cosmetics Regulation Act of 2022 (MoCRA). Reporting odds ratios are estimated for 45 reactions, with Fisher exact tests, Benjamini-Hochberg false discovery rate correction, and a set of exposure-definition, duplicate, temporal, product-level, and symptom-cluster sensitivity analyses.

## Repository contents

| Path | What it is |
|---|---|
| `api_queries/QUERIES.md` | Every openFDA query used, with its purpose and the number of reports it returns |
| `api_queries/fetch_openfda.py` | Re-runs the queries, saves the raw records, and checks them against the study data set |
| `data/Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx` | Report-level study data set (one row per report; see `data/README.md`) |
| `data/openfda_pull_2026-11-20/` | Raw records and query log from a re-run on 20 November 2026 |
| `analysis/hair_straightening_supplement.py` | Runs every analysis and writes the manuscript Figures 1-4 and Table 2, the supplementary tables (Excel and Word), the supplementary figures, and a quality-control report |
| `requirements.txt` | Python packages needed |

## Data source and version

- **Endpoint:** openFDA food adverse event reports, `https://api.fda.gov/food/event.json`, which distributes the legacy CAERS records. Cosmetic reports are identified by industry code 53 (Cosmetics).
- **Version used in the study:** dataset last updated **7 July 2026** (`meta.last_updated`), extracted on **7 August 2026**.
- **Re-run check:** on 20 November 2026 (dataset last updated 6 November 2026), the cohort queries returned the same 352 report numbers with identical reaction terms (`data/openfda_pull_2026-11-20/`).
- **Cohort definitions:** a report enters the relaxer/alkali cohort if it lists a cosmetic product and a product whose brand name contains the word "relaxer", "lye", or "perm", and the keratin/blowout cohort if the brand name contains "keratin" or "blowout". openFDA matches whole words without regard to case, so "perm" does not match "permanent". No report met both definitions.
- openFDA updates its data periodically, so later versions may differ slightly from the study extraction. `fetch_openfda.py` reports any difference.

The source data are public. openFDA notes that adverse event reports are not verified and do not establish that a product caused an event, and they cannot be used to estimate incidence or risk.

## How to reproduce

Install the requirements (Python 3.9 or later):

```bash
pip install -r requirements.txt
```

**1. Re-run the openFDA queries (optional).**

```bash
cd api_queries
python fetch_openfda.py
```

This saves the raw records to `data/openfda_pull_<date>/` and prints whether the report numbers and reaction terms still match the study data set. No API key is needed; if you have one, set the environment variable `OPENFDA_API_KEY`.

**2. Run the analyses.**

```bash
cd analysis
python hair_straightening_supplement.py
```

The script uses `data/Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx` automatically and writes everything to `outputs/`:

- `Figure1_study_flow`, `Figure2_forest_plot`, `Figure3_robustness_heatmap`, `Figure4_temporal_patterns`: the manuscript figures (600-dpi PNG, PDF, and EPS), with every number computed from the data set
- `Supplementary_Tables.xlsx`: Tables S1-S15, including the report-level exposure classification (Table S3d) and the openFDA queries (Table S15)
- `Supporting_Information.docx`: the formatted supplementary tables and Figures S1-S3
- `Table2.docx`: manuscript Table 2 (primary and SA3 co-primary results for the 10 primary signals) as a Word table, formatted like the manuscript's tables
- `FigureS1-S3` (PNG and PDF)
- `QC_report.md`: agreement checks against the data set and the submitted manuscript, hand re-derivations of the test statistics, and a robustness matrix

Figures use Arial when it is installed and otherwise the first available of Helvetica, Liberation Sans, Nimbus Sans, or DejaVu Sans.

To use a copy of the workbook stored elsewhere, edit `WORKBOOK_PATH` at the top of the script.

## Methods in brief

- **Unit of analysis:** the report. Reaction terms (MedDRA preferred terms) are case-normalized and counted once per report.
- **Tested reactions:** the 45 reactions reported at least five times in either cohort. The same 45 reactions and correction are used in every analysis.
- **Statistics:** reporting odds ratio (relaxer/alkali vs keratin/blowout) with Woolf 95% CIs, or exact conditional 95% CIs when a cell is zero; two-sided Fisher exact tests; Benjamini-Hochberg FDR correction (q < .05).
- **Sensitivity analyses:** single cosmetic product (SA1), single suspect product (SA2), product-name validation (SA3, co-primary), SA3 with a single suspect product (SA4), removal of possible duplicates, exclusion of 2010-2012, restriction to 2013-2023 and to 2004-2023, female consumers only, calendar-period- and age-stratified Mantel-Haenszel estimates, logistic regression adjusted for age and sex, brand concentration and leave-top-brand-out analyses, alternative inclusion thresholds, co-occurrence, and symptom clusters.

## Citation

The citation will be added when the article is published.
