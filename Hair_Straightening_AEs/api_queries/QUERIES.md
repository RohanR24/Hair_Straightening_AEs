# openFDA queries

Base URL: `https://api.fda.gov/food/event.json`

All queries were run against the openFDA food adverse event endpoint, which distributes the legacy CAERS records. The study used the dataset version last updated **7 July 2026** (extracted **7 August 2026**). The counts below are from a re-run on 20 November 2026 (dataset last updated 6 November 2026), which returned the same 352 report numbers and reaction terms as the study data set.

| # | Purpose | Query (append to the base URL after `?`) | Reports returned |
|---|---|---|---|
| 1 | Relaxer/alkali cohort, full records | `search=products.industry_code:53+AND+(products.name_brand:"relaxer"+OR+products.name_brand:"lye"+OR+products.name_brand:"perm")&limit=1000` | 76 |
| 2 | Keratin/blowout cohort, full records | `search=products.industry_code:53+AND+(products.name_brand:"keratin"+OR+products.name_brand:"blowout")&limit=1000` | 276 |
| 3 | Both cohorts combined (overlap check) | `search=products.industry_code:53+AND+(products.name_brand:"relaxer"+OR+products.name_brand:"lye"+OR+products.name_brand:"perm"+OR+products.name_brand:"keratin"+OR+products.name_brand:"blowout")&limit=1` | 352 |
| 4 | All cosmetic reports in the endpoint (context) | `search=products.industry_code:53&limit=1` | 52,223 (52,219 in the version used for the study) |
| 5 | Reaction frequencies, relaxer/alkali (check only) | `search=products.industry_code:53+AND+(products.name_brand:"relaxer"+OR+products.name_brand:"lye"+OR+products.name_brand:"perm")&count=reactions.exact` | - |
| 6 | Reaction frequencies, keratin/blowout (check only) | `search=products.industry_code:53+AND+(products.name_brand:"keratin"+OR+products.name_brand:"blowout")&count=reactions.exact` | - |
| 7 | Confirm industry code 53 returns cosmetic products | `search=products.industry_code:53&limit=5` (each product's `industry_name` reads "Cosmetics") | - |

## Notes

- **Whole-word matching.** openFDA searches of `products.name_brand` match whole words without regard to case: `"perm"` matches "HAIR PERM" but not "PERMANENT HAIR COLOR".
- **Mutually exclusive cohorts.** Query 3 returns exactly the sum of queries 1 and 2 (76 + 276 = 352), so no report met both definitions.
- **Report and product structure.** A report enters a cohort if it lists any product with industry code 53 and any product whose brand name contains a keyword. The role of each product (suspect or concomitant) and its industry code are kept in the records and used in sensitivity analyses SA1-SA4.
- **Case variants.** Count queries (5 and 6) tabulate reaction terms exactly as stored, so "ALOPECIA" and "Alopecia" appear separately. The analysis merges case variants within each report and counts each reaction once per report, so its counts are taken from the full records (queries 1 and 2), not from the count queries.
- **Pagination.** Both cohorts are smaller than openFDA's 1,000-record page limit; `fetch_openfda.py` pages with `skip` if a later version exceeds it.
- **Study period.** Included reports were created between 11 August 2003 and 20 December 2023. None was created on or after 29 December 2023, when MoCRA's mandatory serious adverse event reporting took effect.
