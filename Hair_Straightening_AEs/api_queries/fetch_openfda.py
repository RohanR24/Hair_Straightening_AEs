"""
Re-run the openFDA queries that define the two study cohorts and save the raw records.

  python fetch_openfda.py                 # fetch, save, and compare with the study data set in ../data/
  python fetch_openfda.py --no-compare    # fetch and save only

What it does
  1. Queries the openFDA food adverse event endpoint (https://api.fda.gov/food/event.json), which holds the
     legacy CAERS records, for reports listing a cosmetic product (industry code 53) whose brand name contains
     the word "relaxer", "lye", or "perm" (relaxer/alkali cohort) or "keratin" or "blowout" (keratin/blowout cohort).
  2. Records the dataset version (meta.last_updated) and the number of reports returned by every query.
  3. Saves the full JSON records and a flat, one-row-per-report CSV in ../data/openfda_pull_<date>/.
  4. Checks that no report falls in both cohorts.
  5. Compares the report numbers and reaction terms with the study data set
     (../data/Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx) and prints any differences.

Notes
  - openFDA brand-name searches match whole words, case-insensitively ("perm" does not match "permanent").
  - The study used the data version last updated 2026-07-07 (extracted 2026-08-07). openFDA updates its data
    periodically, so later versions can differ slightly; step 5 reports any difference.
  - No API key is needed for these few calls. If you have one, set the environment variable OPENFDA_API_KEY.
  - Requires only the Python standard library; the comparison in step 5 also needs pandas and openpyxl.
"""

import csv
import datetime as dt
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

BASE = "https://api.fda.gov/food/event.json"
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.normpath(os.path.join(HERE, os.pardir, "data"))
STUDY_WORKBOOK = os.path.join(DATA_DIR, "Cosmetics_AE_ROR_Analysis_WITH_DEMOGRAPHICS.xlsx")

RELAXER_TERMS = ["relaxer", "lye", "perm"]
KERATIN_TERMS = ["keratin", "blowout"]


def brand_clause(terms):
    return "(" + " OR ".join(f'products.name_brand:"{t}"' for t in terms) + ")"


COHORTS = {
    "Relaxer/alkali": f"products.industry_code:53 AND {brand_clause(RELAXER_TERMS)}",
    "Keratin/blowout": f"products.industry_code:53 AND {brand_clause(KERATIN_TERMS)}",
}
CHECKS = {
    "Both cohorts combined": f"products.industry_code:53 AND {brand_clause(RELAXER_TERMS + KERATIN_TERMS)}",
    "All cosmetic reports (industry code 53)": "products.industry_code:53",
}


def url_for(search, **params):
    """openFDA expects '+' between search terms; quotes and colons stay readable."""
    q = urllib.parse.quote(search, safe=':"()[]*').replace("%20", "+")
    extra = "".join(f"&{k}={v}" for k, v in params.items())
    key = os.environ.get("OPENFDA_API_KEY")
    return f"{BASE}?search={q}{extra}" + (f"&api_key={key}" if key else "")


def get(url, retries=3):
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(url, timeout=90) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code == 404:                       # openFDA returns 404 when nothing matches
                return {"meta": {"results": {"total": 0}}, "results": []}
            if attempt == retries - 1:
                raise
        except urllib.error.URLError:
            if attempt == retries - 1:
                raise
        time.sleep(2 * (attempt + 1))


def fetch_all(search, page=1000):
    """Download every record matching `search` (paginated with skip)."""
    first = get(url_for(search, limit=page))
    total = first["meta"]["results"]["total"]
    records = list(first.get("results", []))
    while len(records) < total:
        nxt = get(url_for(search, limit=page, skip=len(records)))
        batch = nxt.get("results", [])
        if not batch:
            break
        records += batch
        time.sleep(0.3)
    return first["meta"], total, records


def total_only(search):
    meta = get(url_for(search, limit=1))["meta"]
    return meta, meta["results"]["total"]


def flatten(cohort, r):
    c = r.get("consumer", {}) or {}
    reactions = sorted({t.strip().upper() for t in r.get("reactions", []) if t.strip()})
    prods = r.get("products", []) or []
    return {
        "cohort": cohort,
        "report_number": r.get("report_number", ""),
        "date_created": r.get("date_created", ""),
        "date_started": r.get("date_started", ""),
        "consumer_age": c.get("age", ""),
        "consumer_age_unit": c.get("age_unit", ""),
        "consumer_gender": c.get("gender", ""),
        "outcomes": "; ".join(r.get("outcomes", []) or []),
        "reactions_case_normalized": "; ".join(reactions),
        "n_products": len(prods),
        "products (role | industry code | brand name)": " || ".join(
            f"{p.get('role', '')} | {p.get('industry_code', '')} | {p.get('name_brand', '')}" for p in prods),
    }


def main():
    compare = "--no-compare" not in sys.argv
    stamp = dt.date.today().isoformat()
    out = os.path.join(DATA_DIR, f"openfda_pull_{stamp}")
    os.makedirs(out, exist_ok=True)

    log, rows, by_cohort = [], [], {}
    for cohort, search in COHORTS.items():
        meta, total, recs = fetch_all(search)
        by_cohort[cohort] = recs
        log.append({"query": cohort, "url": url_for(search, limit=1000).split("&api_key=")[0],
                    "last_updated": meta.get("last_updated"), "reports_returned": total})
        with open(os.path.join(out, f"{cohort.split('/')[0].lower()}_records.json"), "w", encoding="utf-8") as fh:
            json.dump(recs, fh, indent=1)
        rows += [flatten(cohort, r) for r in recs]
        print(f"{cohort:40s} {total:6d} reports   (data last updated {meta.get('last_updated')})")
    for name, search in CHECKS.items():
        meta, total = total_only(search)
        log.append({"query": name, "url": url_for(search, limit=1).split("&api_key=")[0],
                    "last_updated": meta.get("last_updated"), "reports_returned": total})
        print(f"{name:40s} {total:6d} reports")

    with open(os.path.join(out, "reports_flat.csv"), "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    with open(os.path.join(out, "query_log.json"), "w", encoding="utf-8") as fh:
        json.dump({"run_date": stamp, "queries": log}, fh, indent=1)

    ids = {c: {r["report_number"] for r in recs} for c, recs in by_cohort.items()}
    overlap = ids["Relaxer/alkali"] & ids["Keratin/blowout"]
    print(f"\nReports in both cohorts: {len(overlap)}" + (f" -> {sorted(overlap)}" if overlap else " (cohorts are mutually exclusive)"))
    print(f"Saved raw records, reports_flat.csv and query_log.json to {out}")

    if compare and os.path.isfile(STUDY_WORKBOOK):
        try:
            import pandas as pd
        except ImportError:
            print("pandas not installed; skipping comparison with the study data set.")
            return
        wb = pd.read_excel(STUDY_WORKBOOK, sheet_name="Report_Level_QC")
        wb["rn"] = wb["Report number"].astype(str).str.strip()
        wb["cohort"] = wb["Arm"].map(lambda a: "Relaxer/alkali" if str(a).startswith("Alkali") else "Keratin/blowout")
        study_rx = dict(zip(wb.rn, wb.Reactions.fillna("").map(
            lambda s: frozenset(t.strip().upper() for t in str(s).split(";") if t.strip()))))
        print("\nComparison with the study data set:")
        all_same = True
        for cohort, recs in by_cohort.items():
            study = set(wb.loc[wb.cohort == cohort, "rn"])
            now = ids[cohort]
            changed = [r["report_number"] for r in recs if r["report_number"] in study_rx and
                       frozenset(t.strip().upper() for t in r.get("reactions", []) if t.strip()) != study_rx[r["report_number"]]]
            print(f"  {cohort}: study {len(study)}, now {len(now)}; only in study {sorted(study - now)}; "
                  f"only now {sorted(now - study)}; reaction sets changed {changed}")
            all_same &= study == now and not changed
        print("  RESULT: identical report numbers and reaction terms." if all_same else
              "  RESULT: differences found (see above); openFDA data may have been updated since the study extraction.")


if __name__ == "__main__":
    main()
