# Study data set

`Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx` is the report-level data set used for every analysis. It is derived entirely from public openFDA records (see `../api_queries/QUERIES.md`) and contains no information beyond what openFDA publishes.

Sheets used by `analysis/hair_straightening_supplement.py`:

| Sheet | Contents |
|---|---|
| `Report_Level_QC` | One row per report: cohort, report number, reaction terms, all listed products with role and industry code, number of cosmetic and suspect products, SA1-SA4 membership, and the product(s) that qualified the report for SA3 |
| `Report_Characteristics` | Consumer sex and age (with unit and availability) and reported outcomes as recorded by openFDA (for example, death, life-threatening event, disability, hospitalization, emergency room visit) |
| `Temporal_Trend` | Creation date of each report and annual counts |
| `Strict_Excluded_Reports` | Reports excluded from SA3 and the reason for each exclusion |
| `Primary_ROR`, `Robustness_Summary`, `Sensitivity_Cohorts` | Pre-computed results that the quality-control step uses to confirm the script reproduces them |

`openfda_pull_2026-11-20/` holds a re-run of the cohort queries: the raw JSON records for each cohort, a one-row-per-report CSV (`reports_flat.csv`), and the query log with the dataset version (`query_log.json`).
