"""
CCCM / Site Management Cluster — Gaza 4W Data Validation & Aggregation Pipeline
================================================================================
Phase 1: Clean consolidated partner submissions, detect data-quality issues,
         and emit a row-level error log.
Phase 2: Produce aggregation tables used by the Streamlit dashboard.

Inputs (in --uploads directory):
  - 4Ws_Consolidated.xlsx                                 (sheet: activities)
  - SM_Cluster_MonthlyActivityReporting_FINAL_TEMPLATE_260505.xlsx
        (sheets: 'Activity Index', 'COD-Sites')

Outputs (in --out directory):
  - cleaned_activities.csv
  - validation_error_log.csv
  - agg_framework_split.csv
  - agg_partner_footprint.csv
  - agg_population_reach.csv
"""
from __future__ import annotations

import argparse
import re
from difflib import SequenceMatcher
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Column-name constants (the submission columns contain embedded newlines —
# isolating them here keeps the rest of the code readable).
# ---------------------------------------------------------------------------
COL_MONTH       = "Reporting Month"
COL_ORG         = "Organization Name"
COL_PROG_ORG    = "Program / Coordinating / Donor Organization"
COL_GOV         = "Governorate\n(Select from Dropdown)"
COL_NBHD        = "Neighborhood\n(Select from Dropdown)"
COL_SITE        = "Site Name\n(Select from Dropdown, or type name and code if not found)"
COL_STATUS      = "Activity\nStatus\n(Select from Dropdown)"
COL_PRIMARY     = "Primary Activities\n(Select From Dropdown)"
COL_SUB         = "Sub Activity\n(Select from Dropdown)"
COL_INDICATOR   = "Indicator\n(Select from Dropdown)"
COL_SPEC        = "Specification\n(Select from Dropdown)"
COL_DETAILS     = "Activity Details\n(Provide additional brief remarks on the activity, if any)"
COL_UNITS       = "Units\n(Auto Calculated, Do not modify)"
COL_TOTAL       = "Total\n Planned/Reached\nCount\n"

DEMO_COLS = [
    "Boys (<18)", "Girls (<18)",
    "Men (18-59)", "Women (18-59)",
    "Elderly Male (≥60)", "Elderly Female (≥60)",
]
PWD_COLS = ["Male-PWD", "Female-PWD"]

# Indicator units that represent counts of people and should disaggregate
PEOPLE_UNITS = {"# of individuals", "# of participants", "# of people"}

# Site code pattern: 3-letter prefix + 4 digits (e.g. GZA6013, DEB0038, NGZ4916, KHY1234, RAF0099)
SITE_CODE_RE = re.compile(r"\b([A-Z]{3}\d{4})\b")

# Governorate normalisation. Master COD-Sites uses "Khan Younis" and
# "Deir Al-Balah" — we align both sides to those spellings so joins are clean.
GOV_CANON = {
    "khan yunis":     "Khan Younis",
    "khan younis":    "Khan Younis",
    "deir al balah":  "Deir Al-Balah",
    "deir al-balah":  "Deir Al-Balah",
    "deir albalah":   "Deir Al-Balah",
    "north gaza":     "North Gaza",
    "gaza":           "Gaza",
    "rafah":          "Rafah",
}


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------
def load_inputs(uploads: Path):
    sub_path = uploads / "4Ws_Consolidated.xlsx"
    tpl_path = uploads / "SM_Cluster_MonthlyActivityReporting_FINAL_TEMPLATE_260505.xlsx"

    activities = pd.read_excel(sub_path, sheet_name="activities")
    index_df   = pd.read_excel(tpl_path, sheet_name="Activity Index")
    sites_df   = pd.read_excel(tpl_path, sheet_name="COD-Sites")

    # Strip whitespace from master column values we'll join on
    index_df["Indicators"] = index_df["Indicators"].astype(str).str.strip()
    index_df["Unit"]       = index_df["Unit"].astype(str).str.strip()
    sites_df["Site ID"]    = sites_df["Site ID"].astype(str).str.strip()
    sites_df["Governorate"] = sites_df["Governorate"].astype(str).str.strip()

    return activities, index_df, sites_df


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def normalise_governorate(value) -> str | None:
    if pd.isna(value):
        return None
    key = re.sub(r"\s+", " ", str(value).strip().lower())
    key = key.replace("-", " ")
    return GOV_CANON.get(key, str(value).strip())


def extract_site_code(site_text) -> str | None:
    if pd.isna(site_text):
        return None
    m = SITE_CODE_RE.search(str(site_text).upper())
    return m.group(1) if m else None


def build_indicator_alias_map(submissions: pd.Series, master: pd.Series,
                              cutoff: float = 0.85) -> dict[str, str]:
    """Fuzzy-match every submitted indicator to its closest master indicator.

    Returns a dict {submitted_string: canonical_master_string}. Any submitted
    string already present verbatim in the master maps to itself.
    """
    master_list = master.dropna().astype(str).str.strip().unique().tolist()
    master_set  = set(master_list)
    alias: dict[str, str] = {}

    for raw in submissions.dropna().astype(str).str.strip().unique():
        if raw in master_set:
            alias[raw] = raw
            continue
        best_score, best_match = 0.0, None
        for m in master_list:
            s = SequenceMatcher(None, raw.lower(), m.lower()).ratio()
            if s > best_score:
                best_score, best_match = s, m
        if best_match is not None and best_score >= cutoff:
            alias[raw] = best_match
        else:
            alias[raw] = raw  # unresolved – will be flagged separately
    return alias


# ---------------------------------------------------------------------------
# Phase 1 — Cleaning + Validation
# ---------------------------------------------------------------------------
def clean_and_validate(activities: pd.DataFrame,
                       index_df: pd.DataFrame,
                       sites_df: pd.DataFrame):
    errors: list[dict] = []
    df = activities.copy()

    # --- 1a. Full-row duplicates (BEFORE adding any synthetic columns) --
    dup_mask = df.duplicated(keep="first")
    removed_per_org = (
        df.loc[dup_mask, COL_ORG].value_counts().to_dict()
        if dup_mask.any() else {}
    )
    for org, n in removed_per_org.items():
        errors.append({
            "row_id": "—", "organization": org, "month": "—", "site": "—",
            "error_type": "Exact Duplicate Removed",
            "description": f"{n} fully-duplicated row(s) deduplicated"
        })
    df = df.loc[~dup_mask].copy()
    df["_row_id"] = np.arange(len(df))  # stable id for traceability (post-dedup)

    # --- 1b. Governorate normalisation ---------------------------------
    df["governorate_clean"] = df[COL_GOV].apply(normalise_governorate)

    # --- 1c. Indicator alignment via fuzzy matching --------------------
    alias = build_indicator_alias_map(df[COL_INDICATOR], index_df["Indicators"])
    df["indicator_clean"] = df[COL_INDICATOR].astype(str).str.strip().map(alias)

    master_idx = index_df.set_index("Indicators")[["Activity_Code", "Unit",
                                                   "Indicator Purpose",
                                                   "Primary Activity",
                                                   "Sub Activity"]]
    unresolved_mask = ~df["indicator_clean"].isin(master_idx.index)
    for _, r in df.loc[unresolved_mask].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Indicator Not in Master",
            "description": f"Indicator '{r[COL_INDICATOR]}' did not match the Activity Index"
        })

    # Attach canonical metadata (NaN for unresolved rows)
    df = df.merge(master_idx, left_on="indicator_clean",
                  right_index=True, how="left", suffixes=("", "_master"))
    df = df.rename(columns={
        "Activity_Code":     "activity_code",
        "Unit":              "unit_master",
        "Indicator Purpose": "indicator_purpose",
        "Primary Activity":  "primary_activity_master",
        "Sub Activity":      "sub_activity_master",
    })

    # --- 1d. Partial duplicates (same org / month / site / indicator / total) ---
    partial_keys = [COL_ORG, COL_MONTH, COL_SITE, "indicator_clean", COL_TOTAL]
    partial_mask = df.duplicated(subset=partial_keys, keep=False)
    # Only flag the SECOND+ occurrences (keep="first") for review
    flag_mask = partial_mask & df.duplicated(subset=partial_keys, keep="first")
    for _, r in df.loc[flag_mask].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Potential Partial Duplicate",
            "description": (f"Same org/month/site/indicator/total as another row "
                            f"(indicator='{r['indicator_clean']}', total={r[COL_TOTAL]})")
        })

    # --- 1e. Neighborhood null ------------------------------------------
    nbhd_missing = df[COL_NBHD].isna() | (df[COL_NBHD].astype(str).str.strip() == "")
    for _, r in df.loc[nbhd_missing].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Missing Neighborhood",
            "description": "Neighborhood field is blank"
        })

    # --- 1f. Site code reference integrity ------------------------------
    valid_site_ids = set(sites_df["Site ID"].dropna().astype(str).str.upper())
    df["site_code"] = df[COL_SITE].apply(extract_site_code)

    no_code_mask = df["site_code"].isna()
    for _, r in df.loc[no_code_mask].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "No Site Code Detected",
            "description": "Could not extract a SiteID (e.g. GZA1234) from the site name"
        })

    bad_code_mask = (~no_code_mask) & (~df["site_code"].isin(valid_site_ids))
    for _, r in df.loc[bad_code_mask].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Unknown Site Code",
            "description": f"Site code '{r['site_code']}' not found in COD-Sites master"
        })

    # --- 1g. Demographic / caseload coherence ---------------------------
    df[COL_TOTAL] = pd.to_numeric(df[COL_TOTAL], errors="coerce")
    for c in DEMO_COLS + PWD_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["demo_sum"] = df[DEMO_COLS].sum(axis=1, min_count=1)
    df["demo_all_missing"] = df[DEMO_COLS].isna().all(axis=1)

    is_people = df["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)
    has_total = df[COL_TOTAL].notna() & (df[COL_TOTAL] > 0)

    # Missing disaggregation (informational, not an error)
    miss_disagg = is_people & has_total & df["demo_all_missing"]
    for _, r in df.loc[miss_disagg].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Missing Disaggregation",
            "description": (f"People-indicator with total={int(r[COL_TOTAL])} "
                            f"has no age/sex breakdown")
        })

    # Demographic sum mismatch
    mismatch = is_people & has_total & (~df["demo_all_missing"]) & \
               (df["demo_sum"].round() != df[COL_TOTAL].round())
    for _, r in df.loc[mismatch].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Demographic Sum Mismatch",
            "description": (f"Sum of age/sex disaggregation = {int(r['demo_sum'])} ≠ "
                            f"reported total = {int(r[COL_TOTAL])}")
        })

    # Missing total entirely
    no_total = df[COL_TOTAL].isna()
    for _, r in df.loc[no_total].iterrows():
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Missing Total Count",
            "description": "Total Planned/Reached Count is blank or non-numeric"
        })

    # --- 1h. Outliers ----------------------------------------------------
    # Per-indicator z-score above 3, AND/OR absolute > 5,000 for a single
    # site-month combo on a people indicator.
    df["zscore"] = (
        df.groupby("indicator_clean")[COL_TOTAL]
          .transform(lambda s: (s - s.mean()) / s.std(ddof=0)
                     if s.std(ddof=0) and not np.isnan(s.std(ddof=0)) else 0)
    )
    high_z = df["zscore"].abs() > 3
    big_abs = is_people & (df[COL_TOTAL] > 5000)
    outlier_mask = high_z | big_abs
    for _, r in df.loc[outlier_mask].iterrows():
        reason_bits = []
        if abs(r["zscore"]) > 3:
            reason_bits.append(f"z={r['zscore']:.1f}")
        if pd.notna(r[COL_TOTAL]) and r[COL_TOTAL] > 5000:
            reason_bits.append(f"total={int(r[COL_TOTAL])}>5000")
        errors.append({
            "row_id": r["_row_id"], "organization": r[COL_ORG],
            "month": r[COL_MONTH], "site": r[COL_SITE],
            "error_type": "Potential Outlier",
            "description": f"Unusually high value ({', '.join(reason_bits)}) for indicator '{r['indicator_clean']}'"
        })

    # --- Final cleaned frame --------------------------------------------
    df["is_clean"] = True  # default; flip below for any row that produced an error
    error_df = pd.DataFrame(errors)
    if not error_df.empty:
        dirty_ids = set(error_df.loc[error_df["row_id"] != "—", "row_id"])
        df.loc[df["_row_id"].isin(dirty_ids), "is_clean"] = False

    # Add tidy helper columns for the dashboard
    df["framework_category"] = np.where(
        df["indicator_purpose"].fillna("").str.contains("Flash Appeal"),
        "Flash Appeal",
        "Minimum Package / SMC Coordination"
    )
    # Where indicator_purpose is missing (unresolved indicator), still tag NA
    df.loc[df["indicator_purpose"].isna(), "framework_category"] = "Unresolved"

    return df, error_df, removed_per_org


# ---------------------------------------------------------------------------
# Phase 2 — Aggregations
# ---------------------------------------------------------------------------
def build_aggregations(clean: pd.DataFrame):
    # a) Flash Appeal vs Minimum Package split
    agg_split = (
        clean.dropna(subset=[COL_TOTAL])
             .groupby(["framework_category", COL_MONTH], as_index=False)
             .agg(rows=("_row_id", "count"),
                  total_count=(COL_TOTAL, "sum"),
                  unique_sites=("site_code", pd.Series.nunique),
                  unique_partners=(COL_ORG, pd.Series.nunique))
    )

    # b) Partner footprint matrix: unique sites per partner per month per governorate
    agg_partners = (
        clean.dropna(subset=["site_code"])
             .groupby([COL_ORG, COL_MONTH, "governorate_clean"], as_index=False)
             .agg(unique_sites=("site_code", pd.Series.nunique),
                  activities=("_row_id", "count"))
             .rename(columns={COL_ORG: "organization",
                              COL_MONTH: "month",
                              "governorate_clean": "governorate"})
    )

    # c) Safe population reach — only people-indicators, grouped to prevent
    # double-counting at the same site / month / primary activity.
    people_mask = clean["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)
    people_df = clean.loc[people_mask].copy()
    # Deduplicate at site x month x primary_activity (take MAX so the highest
    # reported reach wins when partners overlap)
    deduped = (
        people_df.groupby(["governorate_clean", COL_MONTH,
                           "primary_activity_master", "site_code"],
                          as_index=False)[COL_TOTAL].max()
    )
    agg_reach = (
        deduped.groupby(["governorate_clean", COL_MONTH,
                         "primary_activity_master"], as_index=False)
               .agg(reached=(COL_TOTAL, "sum"),
                    sites=("site_code", pd.Series.nunique))
               .rename(columns={"governorate_clean": "governorate",
                                COL_MONTH: "month",
                                "primary_activity_master": "primary_activity"})
    )

    return agg_split, agg_partners, agg_reach


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser()
    p.add_argument("--uploads", default="/mnt/user-data/uploads")
    p.add_argument("--out",     default="/mnt/user-data/outputs")
    args = p.parse_args()

    uploads = Path(args.uploads)
    out     = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    activities, index_df, sites_df = load_inputs(uploads)
    print(f"Loaded {len(activities)} submission rows, "
          f"{len(index_df)} indicators, {len(sites_df)} master sites.")

    clean, error_log, removed = clean_and_validate(activities, index_df, sites_df)
    print(f"After dedup: {len(clean)} rows. "
          f"Removed {sum(removed.values())} exact duplicates across "
          f"{len(removed)} partner(s).")
    print(f"Validation log: {len(error_log)} entries.")

    agg_split, agg_partners, agg_reach = build_aggregations(clean)

    # Persist
    clean.to_csv(out / "cleaned_activities.csv", index=False)
    error_log.to_csv(out / "validation_error_log.csv", index=False)
    agg_split.to_csv(out / "agg_framework_split.csv", index=False)
    agg_partners.to_csv(out / "agg_partner_footprint.csv", index=False)
    agg_reach.to_csv(out / "agg_population_reach.csv", index=False)

    # Headline numbers
    total_clean = int(clean["is_clean"].sum())
    pct = 100 * total_clean / len(clean) if len(clean) else 0
    print(f"Data Quality Score: {pct:.1f}% ({total_clean}/{len(clean)} rows clean)")
    print("Outputs written to", out)


if __name__ == "__main__":
    main()
