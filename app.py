"""
Gaza CCCM/Site Management Cluster — 4W Activity Dashboard
=========================================================
Self-contained Streamlit app for deployment on Streamlit Community Cloud.

Expects this repo layout:
    app.py
    validate_pipeline.py
    data/
        4Ws_Consolidated.xlsx
        SM_Cluster_MonthlyActivityReporting_FINAL_TEMPLATE_260505.xlsx

Run locally:    streamlit run app.py
Deploy:         push to GitHub → share.streamlit.io
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# Make the validator importable from the same directory
sys.path.insert(0, str(Path(__file__).parent))
from validate_pipeline import load_inputs, clean_and_validate

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
DATA_DIR = Path(__file__).parent / "data"

COL_MONTH     = "Reporting Month"
COL_ORG       = "Organization Name"
COL_SITE      = "Site Name\n(Select from Dropdown, or type name and code if not found)"
COL_TOTAL     = "Total\n Planned/Reached\nCount\n"
DEMO_COLS = [
    "Boys (<18)", "Girls (<18)",
    "Men (18-59)", "Women (18-59)",
    "Elderly Male (≥60)", "Elderly Female (≥60)",
]
PWD_COLS = ["Male-PWD", "Female-PWD"]
PEOPLE_UNITS = {"# of individuals", "# of participants", "# of people"}

# Accessible palette tuned for humanitarian reporting (colour-blind safe)
PALETTE = ["#0072B2", "#E69F00", "#009E73", "#D55E00", "#56B4E9", "#CC79A7", "#F0E442"]
GOV_ORDER = ["North Gaza", "Gaza", "Deir Al-Balah", "Khan Younis", "Rafah"]

st.set_page_config(
    page_title="Gaza SMC 4W Dashboard",
    page_icon="🏚️",
    layout="wide",
)


# ---------------------------------------------------------------------------
# Pipeline runner (cached so it only re-runs when the xlsx files change)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Validating partner submissions…")
def load_data():
    """Run the full validation pipeline in-memory and return cleaned data + errors.

    Cached: Streamlit re-runs this only when the underlying xlsx files change.
    """
    activities, index_df, sites_df = load_inputs(DATA_DIR)
    clean, errors, _removed = clean_and_validate(activities, index_df, sites_df)

    # Coerce numerics (validate_pipeline already does this internally, but the
    # in-memory frame may carry object dtypes through the cache boundary)
    clean[COL_TOTAL] = pd.to_numeric(clean[COL_TOTAL], errors="coerce")
    for c in DEMO_COLS + PWD_COLS:
        if c in clean.columns:
            clean[c] = pd.to_numeric(clean[c], errors="coerce")
    return clean, errors


# ---------------------------------------------------------------------------
# Sidebar filters
# ---------------------------------------------------------------------------
def sidebar_filters(clean: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filters")

    months = sorted(clean[COL_MONTH].dropna().unique())
    sel_month = st.sidebar.multiselect("Reporting Month",
                                       options=months, default=months)

    govs = [g for g in GOV_ORDER if g in clean["governorate_clean"].unique()]
    govs += [g for g in clean["governorate_clean"].dropna().unique() if g not in govs]
    sel_gov = st.sidebar.multiselect("Governorate", options=govs, default=govs)

    orgs = sorted(clean[COL_ORG].dropna().unique())
    sel_org = st.sidebar.multiselect("Organization", options=orgs, default=orgs)

    fw_opts = sorted(clean["framework_category"].dropna().unique())
    sel_fw = st.sidebar.multiselect("Framework Category",
                                    options=fw_opts, default=fw_opts)

    mask = (
        clean[COL_MONTH].isin(sel_month)
        & clean["governorate_clean"].isin(sel_gov)
        & clean[COL_ORG].isin(sel_org)
        & clean["framework_category"].isin(sel_fw)
    )
    st.sidebar.markdown(f"**Rows shown:** {mask.sum():,} of {len(clean):,}")
    st.sidebar.caption(
        "Filters apply to all KPIs and charts. "
        "The Data Quality tab shows the *full* validation log."
    )
    return clean.loc[mask].copy()


# ---------------------------------------------------------------------------
# KPI helpers
# ---------------------------------------------------------------------------
def compute_verified_reach(df: pd.DataFrame) -> int:
    """Sum of people-indicator totals, deduplicated at the
    site x month x primary-activity level to avoid double counting."""
    if df.empty:
        return 0
    people = df[df["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)]
    if people.empty:
        return 0
    deduped = (
        people.groupby(["site_code", COL_MONTH, "primary_activity_master"],
                       dropna=False)[COL_TOTAL].max()
    )
    return int(deduped.sum(skipna=True))


def quality_score(df: pd.DataFrame) -> float:
    if df.empty:
        return 0.0
    return 100.0 * df["is_clean"].sum() / len(df)


# ---------------------------------------------------------------------------
# Visual sections
# ---------------------------------------------------------------------------
def framework_chart(df: pd.DataFrame):
    if df.empty:
        st.info("No data for current filters.")
        return
    agg = (df.dropna(subset=[COL_TOTAL])
             .groupby(["framework_category", COL_MONTH], as_index=False)[COL_TOTAL]
             .sum()
             .rename(columns={COL_TOTAL: "Total Reach / Count",
                              COL_MONTH: "Month"}))
    fig = px.bar(agg, x="Total Reach / Count", y="framework_category",
                 color="Month", orientation="h", barmode="group",
                 color_discrete_sequence=PALETTE,
                 labels={"framework_category": "Framework"})
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10),
                      legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)


def geographic_chart(df: pd.DataFrame):
    if df.empty:
        st.info("No data for current filters.")
        return
    # Use deduplicated reach so the picture is honest
    people = df[df["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)]
    deduped = (
        people.groupby(["governorate_clean", "Neighborhood\n(Select from Dropdown)",
                        "site_code", COL_MONTH, "primary_activity_master"],
                       dropna=False)[COL_TOTAL].max().reset_index()
    )
    if deduped.empty:
        st.info("No people-indicator rows for the current filters.")
        return
    deduped = deduped.rename(columns={
        "governorate_clean": "Governorate",
        "Neighborhood\n(Select from Dropdown)": "Neighborhood",
        COL_TOTAL: "Reach",
    })
    deduped["Neighborhood"] = deduped["Neighborhood"].fillna("(unspecified)")
    fig = px.sunburst(deduped, path=["Governorate", "Neighborhood"],
                      values="Reach",
                      color="Governorate",
                      color_discrete_sequence=PALETTE)
    fig.update_layout(height=480, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)


def partner_matrix(df: pd.DataFrame):
    if df.empty:
        st.info("No data for current filters.")
        return
    pa_col = "primary_activity_master"
    pivot = (df.dropna(subset=[pa_col])
               .groupby([COL_ORG, pa_col])["_row_id"].count()
               .unstack(fill_value=0))
    if pivot.empty:
        st.info("Nothing to show.")
        return
    # Shorten long activity labels for the axis
    pivot.columns = [c if len(c) <= 60 else c[:57] + "…" for c in pivot.columns]
    fig = px.imshow(pivot.values,
                    x=pivot.columns, y=pivot.index,
                    color_continuous_scale="Blues",
                    aspect="auto", text_auto=True,
                    labels=dict(color="Activity count"))
    fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=120),
                      xaxis=dict(tickangle=-30))
    st.plotly_chart(fig, use_container_width=True)


def demographic_chart(df: pd.DataFrame):
    cols_present = [c for c in DEMO_COLS if c in df.columns]
    if df.empty or not cols_present:
        st.info("No demographic data for current filters.")
        return
    totals = df[cols_present].sum(numeric_only=True)
    totals = totals[totals > 0]
    if totals.empty:
        st.info("No age/sex disaggregation reported for the selected rows.")
        return
    dem_df = totals.reset_index()
    dem_df.columns = ["Group", "Count"]
    fig = px.bar(dem_df, x="Group", y="Count",
                 color="Group", color_discrete_sequence=PALETTE,
                 text="Count")
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
    fig.update_layout(height=360, showlegend=False,
                      margin=dict(l=10, r=10, t=10, b=10),
                      yaxis_title="Individuals reached")
    st.plotly_chart(fig, use_container_width=True)

    male_pwd   = int(df.get("Male-PWD",   pd.Series(dtype=float)).sum(skipna=True))
    female_pwd = int(df.get("Female-PWD", pd.Series(dtype=float)).sum(skipna=True))
    c1, c2, c3 = st.columns(3)
    c1.metric("Male PWD",   f"{male_pwd:,}")
    c2.metric("Female PWD", f"{female_pwd:,}")
    c3.metric("Total PWD",  f"{male_pwd + female_pwd:,}")


def quality_explorer(error_log: pd.DataFrame):
    if error_log.empty:
        st.success("No validation issues recorded. 🎉")
        return
    c1, c2 = st.columns([1, 1])
    with c1:
        sel_org = st.multiselect(
            "Partner", sorted(error_log["organization"].dropna().unique()),
            default=[])
    with c2:
        sel_type = st.multiselect(
            "Error type", sorted(error_log["error_type"].dropna().unique()),
            default=[])

    view = error_log.copy()
    if sel_org:
        view = view[view["organization"].isin(sel_org)]
    if sel_type:
        view = view[view["error_type"].isin(sel_type)]

    st.markdown(f"**{len(view):,}** issue(s) shown.")
    st.dataframe(view, use_container_width=True, hide_index=True)

    # Download button for the filtered log — useful for emailing partners
    csv_bytes = view.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered log (CSV)",
        data=csv_bytes,
        file_name="validation_log_filtered.csv",
        mime="text/csv",
    )

    st.markdown("##### Issue distribution by partner")
    chart_df = (error_log.groupby(["organization", "error_type"])
                         .size().reset_index(name="count"))
    fig = px.bar(chart_df, x="organization", y="count", color="error_type",
                 color_discrete_sequence=PALETTE)
    fig.update_layout(height=360, xaxis_title="", yaxis_title="Issues",
                      legend_title_text="", margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
def main():
    st.title("Gaza Site Management Cluster — 4W Activity Dashboard")
    st.caption("Partner submissions are validated and deduplicated on every "
               "deploy. Filters apply to all sections except the Data Quality log.")

    try:
        clean, errors = load_data()
    except FileNotFoundError as e:
        st.error(
            f"Could not load source files from `{DATA_DIR}`. "
            f"Make sure both xlsx files are committed to the `data/` folder."
        )
        st.exception(e)
        return

    filt = sidebar_filters(clean)

    # --- KPI row -------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Verified Reach (individuals)",
              f"{compute_verified_reach(filt):,}")
    k2.metric("Active Partners",
              f"{filt[COL_ORG].nunique():,}")
    k3.metric("Sites Covered",
              f"{filt['site_code'].dropna().nunique():,}")
    k4.metric("Data Quality Score",
              f"{quality_score(filt):.1f}%")

    st.divider()

    tabs = st.tabs([
        "Framework Performance",
        "Geographic Footprint",
        "Partner × Activity Matrix",
        "Demographics",
        "Data Quality Log",
    ])

    with tabs[0]:
        st.markdown("Total reach by framework category (Flash Appeal vs. "
                    "Minimum Package / SMC Coordination), split by month.")
        framework_chart(filt)

    with tabs[1]:
        st.markdown("Reach broken down by Governorate → Neighborhood. "
                    "Reach figures are deduplicated at site × month × primary "
                    "activity to avoid double-counting.")
        geographic_chart(filt)

    with tabs[2]:
        st.markdown("Count of activity entries by partner and primary activity. "
                    "Empty cells reveal coverage gaps.")
        partner_matrix(filt)

    with tabs[3]:
        st.markdown("Age / sex breakdown of people reached, plus PWD totals "
                    "(disability disaggregation is reported separately, not "
                    "added to the demographic total).")
        demographic_chart(filt)

    with tabs[4]:
        st.markdown("Row-level validation findings from the pipeline. "
                    "Filter by partner to share the action list with each org.")
        quality_explorer(errors)


if __name__ == "__main__":
    main()
