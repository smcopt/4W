"""
Gaza Site Management Cluster — 4W Activity Dashboard
=====================================================
Power-BI-style public dashboard, branded per the CCCM Cluster Design Guide.

Repo layout expected on Streamlit Cloud:
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
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from validate_pipeline import load_inputs, clean_and_validate

DATA_DIR = Path(__file__).parent / "data"

# ---------------------------------------------------------------------------
# CCCM Cluster brand palette (per Design Guide)
# ---------------------------------------------------------------------------
BLUE_SAPPHIRE   = "#1B657C"   # primary
BURNT_SIENNA    = "#EC6B4D"   # primary accent
BALTIC_SEA      = "#2C2C2C"   # text (warmer than black)
ECRU_WHITE      = "#F5F3E8"   # canvas
MOSS_GREEN      = "#BBDFBB"   # secondary
MOONSTONE_BLUE  = "#6FC5BC"   # secondary
STEEL_BLUE      = "#4595AD"   # secondary

# Ordered sequence for categorical color cycling
CCCM_SEQUENCE = [BLUE_SAPPHIRE, BURNT_SIENNA, STEEL_BLUE, MOONSTONE_BLUE,
                 MOSS_GREEN, "#A6C7D3", "#F4A88A"]
# Sequential scale for choropleths / heatmaps — Ecru → Blue Sapphire
CCCM_SEQUENTIAL = [
    [0.00, ECRU_WHITE],
    [0.25, "#BFD7DE"],
    [0.50, "#6FA8B8"],
    [0.75, STEEL_BLUE],
    [1.00, BLUE_SAPPHIRE],
]

# ---------------------------------------------------------------------------
# Column constants
# ---------------------------------------------------------------------------
COL_MONTH = "Reporting Month"
COL_ORG   = "Organization Name"
COL_SITE  = "Site Name\n(Select from Dropdown, or type name and code if not found)"
COL_TOTAL = "Total\n Planned/Reached\nCount\n"
COL_NBHD  = "Neighborhood\n(Select from Dropdown)"
DEMO_COLS = [
    "Boys (<18)", "Girls (<18)",
    "Men (18-59)", "Women (18-59)",
    "Elderly Male (≥60)", "Elderly Female (≥60)",
]
PWD_COLS = ["Male-PWD", "Female-PWD"]
PEOPLE_UNITS = {"# of individuals", "# of participants", "# of people"}
GOV_ORDER = ["North Gaza", "Gaza", "Deir Al-Balah", "Khan Younis", "Rafah"]

# ---------------------------------------------------------------------------
# Page config + global CSS (Inter font, brand colors, BI-style chrome)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Gaza SMC 4W Dashboard",
    page_icon="🏚️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">

<style>
/* Global font + canvas */
html, body, [class*="css"]  {{
    font-family: 'Inter', 'Arial', sans-serif !important;
    color: {BALTIC_SEA};
}}
.stApp {{
    background-color: {ECRU_WHITE};
}}

/* Hide default Streamlit chrome for a cleaner BI look */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{visibility: hidden;}}
.block-container {{padding-top: 0.5rem; padding-bottom: 1rem; max-width: 100%;}}

/* Sidebar styling — Blue Sapphire panel like the Ukraine dashboard */
section[data-testid="stSidebar"] {{
    background-color: {BLUE_SAPPHIRE};
    border-right: 0;
}}
section[data-testid="stSidebar"] * {{
    color: {ECRU_WHITE} !important;
}}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
    color: {ECRU_WHITE} !important;
    font-weight: 600;
}}
section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {{
    background-color: {BURNT_SIENNA} !important;
    color: white !important;
}}
section[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background-color: rgba(245,243,232,0.12);
    border-color: rgba(245,243,232,0.3);
}}

/* Header bar */
.header-bar {{
    background-color: {BLUE_SAPPHIRE};
    color: {ECRU_WHITE};
    padding: 14px 24px;
    border-radius: 6px;
    margin-bottom: 14px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-left: 8px solid {BURNT_SIENNA};
}}
.header-bar h1 {{
    color: {ECRU_WHITE};
    margin: 0;
    font-size: 22px;
    font-weight: 700;
    letter-spacing: 0.2px;
}}
.header-bar .subtitle {{
    color: {ECRU_WHITE};
    opacity: 0.85;
    font-size: 13px;
    margin-top: 2px;
}}
.header-bar .timestamp {{
    color: {ECRU_WHITE};
    opacity: 0.9;
    font-size: 12px;
    text-align: right;
}}

/* KPI cards */
.kpi-card {{
    background-color: white;
    border-radius: 6px;
    padding: 16px 18px;
    border-top: 4px solid {BLUE_SAPPHIRE};
    box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    height: 100%;
}}
.kpi-card.accent {{ border-top-color: {BURNT_SIENNA}; }}
.kpi-label {{
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.7px;
    color: {BALTIC_SEA};
    opacity: 0.7;
    font-weight: 500;
    margin: 0 0 4px 0;
}}
.kpi-value {{
    font-size: 28px;
    font-weight: 700;
    color: {BLUE_SAPPHIRE};
    margin: 0;
    line-height: 1.1;
}}
.kpi-card.accent .kpi-value {{ color: {BURNT_SIENNA}; }}
.kpi-sub {{
    font-size: 11px;
    color: {BALTIC_SEA};
    opacity: 0.6;
    margin-top: 2px;
}}

/* Chart panels — BI-style framed cards with header strip */
.panel {{
    background-color: white;
    border-radius: 6px;
    padding: 0;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    margin-bottom: 10px;
    overflow: hidden;
}}
.panel-header {{
    background-color: #E8E2D3;
    color: {BALTIC_SEA};
    padding: 8px 14px;
    font-weight: 600;
    font-size: 13px;
    letter-spacing: 0.3px;
    border-bottom: 2px solid {BLUE_SAPPHIRE};
}}
.panel-body {{
    padding: 8px 14px 12px 14px;
}}

/* Tables */
.stDataFrame {{
    border-radius: 4px;
    overflow: hidden;
}}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
    background: transparent;
}}
.stTabs [data-baseweb="tab"] {{
    background-color: rgba(27,101,124,0.08);
    color: {BLUE_SAPPHIRE};
    border-radius: 4px 4px 0 0;
    padding: 6px 16px;
    font-weight: 500;
}}
.stTabs [aria-selected="true"] {{
    background-color: {BLUE_SAPPHIRE} !important;
    color: {ECRU_WHITE} !important;
}}

/* Caption / footer note */
.brand-footer {{
    font-size: 11px;
    color: {BALTIC_SEA};
    opacity: 0.6;
    padding: 8px 4px 0 4px;
    border-top: 1px solid rgba(27,101,124,0.15);
    margin-top: 8px;
}}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Data load (validation pipeline runs in-memory, cached)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Validating partner submissions…")
def load_data():
    activities, index_df, sites_df = load_inputs(DATA_DIR)
    clean, errors, _removed = clean_and_validate(activities, index_df, sites_df)
    clean[COL_TOTAL] = pd.to_numeric(clean[COL_TOTAL], errors="coerce")
    for c in DEMO_COLS + PWD_COLS:
        if c in clean.columns:
            clean[c] = pd.to_numeric(clean[c], errors="coerce")
    return clean, errors


# ---------------------------------------------------------------------------
# Plotly theme helper
# ---------------------------------------------------------------------------
def brand_layout(fig: go.Figure, height: int = 320, showlegend: bool = True) -> go.Figure:
    # Preserve any custom margin already set by the caller
    existing_margin = fig.layout.margin
    margin = dict(l=10, r=10, t=10, b=10)
    if existing_margin:
        for k in ("l", "r", "t", "b"):
            v = getattr(existing_margin, k, None)
            if v is not None:
                margin[k] = v
    fig.update_layout(
        font=dict(family="Inter, Arial, sans-serif", color=BALTIC_SEA, size=12),
        paper_bgcolor="white",
        plot_bgcolor="white",
        height=height,
        margin=margin,
        showlegend=showlegend,
        xaxis=dict(gridcolor="rgba(44,44,44,0.08)", zerolinecolor="rgba(44,44,44,0.15)"),
        yaxis=dict(gridcolor="rgba(44,44,44,0.08)", zerolinecolor="rgba(44,44,44,0.15)"),
    )
    # Only set default legend orientation if caller hasn't customized
    if not fig.layout.legend.orientation:
        fig.update_layout(legend=dict(
            font=dict(size=11), orientation="h",
            yanchor="bottom", y=-0.25, xanchor="left", x=0,
        ))
    return fig


def panel(title: str):
    """Context manager-ish helper to wrap a chart in a branded panel."""
    st.markdown(
        f'<div class="panel"><div class="panel-header">{title}</div>'
        f'<div class="panel-body">',
        unsafe_allow_html=True,
    )


def end_panel():
    st.markdown("</div></div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Brand-coloured cell gradient (matplotlib-free replacement for
# Styler.background_gradient, which requires matplotlib on Streamlit Cloud)
# ---------------------------------------------------------------------------
def _hex_to_rgb(hx: str) -> tuple[int, int, int]:
    hx = hx.lstrip("#")
    return int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)


def brand_gradient(series: pd.Series,
                   start_hex: str = ECRU_WHITE,
                   end_hex: str = BURNT_SIENNA) -> list[str]:
    """Per-cell CSS `background-color` strings, linearly interpolated between
    two brand colours based on each value's position in `series`."""
    s = pd.to_numeric(series, errors="coerce").fillna(0)
    vmin, vmax = float(s.min()), float(s.max())
    rng = vmax - vmin if vmax > vmin else 1.0
    r1, g1, b1 = _hex_to_rgb(start_hex)
    r2, g2, b2 = _hex_to_rgb(end_hex)
    styles = []
    for v in s:
        t = (v - vmin) / rng
        t = t ** 0.75  # gentle curve so small values still get a hint of colour
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        text = BALTIC_SEA if t < 0.55 else "white"
        styles.append(f"background-color: rgb({r},{g},{b}); color: {text};")
    return styles


# ---------------------------------------------------------------------------
# Sidebar (filters + branding)
# ---------------------------------------------------------------------------
def sidebar_filters(clean: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar:
        st.markdown(f"""
        <div style="text-align:center; padding: 12px 0 18px 0;
                    border-bottom: 1px solid rgba(245,243,232,0.25); margin-bottom: 14px;">
            <div style="font-size:22px; font-weight:700; letter-spacing: 1px;">CCCM CLUSTER</div>
            <div style="font-size:11px; opacity:0.85; margin-top:2px;">
                Site Management — Gaza
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### Filters")

        months = sorted(clean[COL_MONTH].dropna().unique())
        sel_month = st.multiselect("Reporting Month", options=months, default=months)

        govs = [g for g in GOV_ORDER if g in clean["governorate_clean"].unique()]
        govs += [g for g in clean["governorate_clean"].dropna().unique() if g not in govs]
        sel_gov = st.multiselect("Governorate", options=govs, default=govs)

        orgs = sorted(clean[COL_ORG].dropna().unique())
        sel_org = st.multiselect("Organization", options=orgs, default=orgs)

        fw_opts = sorted(clean["framework_category"].dropna().unique())
        sel_fw = st.multiselect("Framework Category", options=fw_opts, default=fw_opts)

        prim_opts = sorted(clean["primary_activity_master"].dropna().unique())
        sel_prim = st.multiselect("Primary Activity", options=prim_opts, default=[])
        if not sel_prim:
            sel_prim = prim_opts

        mask = (
            clean[COL_MONTH].isin(sel_month)
            & clean["governorate_clean"].isin(sel_gov)
            & clean[COL_ORG].isin(sel_org)
            & clean["framework_category"].isin(sel_fw)
            & clean["primary_activity_master"].isin(sel_prim)
        )
        filtered = clean.loc[mask].copy()

        st.markdown("---")
        st.markdown(
            f"<div style='font-size:11px; opacity:0.85;'>"
            f"<b>{mask.sum():,}</b> activities shown<br>"
            f"of {len(clean):,} validated entries"
            f"</div>", unsafe_allow_html=True
        )
        st.markdown(
            f"<div style='font-size:10px; opacity:0.6; margin-top:18px;'>"
            f"Source: SMC Partner 4W Submissions<br>"
            f"Validated via in-memory pipeline<br>"
            f"Sites cross-checked against COD-Sites"
            f"</div>", unsafe_allow_html=True
        )
    return filtered


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
def header(clean: pd.DataFrame):
    months_in_data = clean[COL_MONTH].dropna().unique().tolist()
    period = " – ".join(months_in_data) if months_in_data else "—"
    st.markdown(f"""
    <div class="header-bar">
      <div>
        <h1>Site Management Cluster — 4W Operational Presence</h1>
        <div class="subtitle">Who is doing What, Where and for Whom · Gaza Strip</div>
      </div>
      <div class="timestamp">
        Reporting period: <b>{period}</b><br>
        Last validated: <b>{pd.Timestamp.today().strftime('%d %b %Y')}</b>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
def compute_verified_reach(df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    people = df[df["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)]
    if people.empty:
        return 0
    return int(
        people.groupby(["site_code", COL_MONTH, "primary_activity_master"],
                       dropna=False)[COL_TOTAL].max().sum(skipna=True)
    )


def kpi_card(col, label: str, value: str, sublabel: str = "", accent: bool = False):
    css_class = "kpi-card accent" if accent else "kpi-card"
    col.markdown(f"""
    <div class="{css_class}">
        <p class="kpi-label">{label}</p>
        <p class="kpi-value">{value}</p>
        <p class="kpi-sub">{sublabel}</p>
    </div>
    """, unsafe_allow_html=True)


def kpi_ribbon(filt: pd.DataFrame):
    c1, c2, c3, c4, c5 = st.columns(5)
    reach = compute_verified_reach(filt)
    activities = len(filt)
    partners = filt[COL_ORG].nunique()
    sites = filt["site_code"].dropna().nunique()
    govs = filt["governorate_clean"].dropna().nunique()

    kpi_card(c1, "People Reached", f"{reach:,}",
             "deduplicated across overlapping activities", accent=True)
    kpi_card(c2, "Activities Reported", f"{activities:,}",
             "validated partner entries")
    kpi_card(c3, "Implementing Partners", f"{partners}",
             "organizations active")
    kpi_card(c4, "Sites Covered", f"{sites}",
             "matched to COD-Sites master")
    kpi_card(c5, "Governorates", f"{govs}",
             "of 5 in the Gaza Strip")


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
def chart_reach_by_governorate(df: pd.DataFrame):
    people = df[df["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)]
    if people.empty:
        st.info("No people-indicator data for current filters.")
        return
    deduped = (people.groupby(["governorate_clean", "site_code",
                               COL_MONTH, "primary_activity_master"],
                              dropna=False)[COL_TOTAL].max().reset_index())
    agg = (deduped.groupby("governorate_clean", as_index=False)[COL_TOTAL].sum()
                  .sort_values(COL_TOTAL, ascending=True))
    agg = agg[agg["governorate_clean"].notna()]
    fig = go.Figure(go.Bar(
        x=agg[COL_TOTAL], y=agg["governorate_clean"],
        orientation="h",
        marker=dict(color=BLUE_SAPPHIRE),
        text=[f"{int(v):,}" for v in agg[COL_TOTAL]],
        textposition="outside",
        textfont=dict(color=BALTIC_SEA, size=11),
        hovertemplate="<b>%{y}</b><br>People reached: %{x:,}<extra></extra>",
    ))
    xmax = float(agg[COL_TOTAL].max() or 0) * 1.18
    fig.update_layout(
        xaxis=dict(showticklabels=False, range=[0, xmax]),
        yaxis=dict(title=None, tickfont=dict(size=11)),
    )
    st.plotly_chart(brand_layout(fig, height=280, showlegend=False),
                    use_container_width=True)


def chart_framework_split(df: pd.DataFrame):
    if df.empty:
        st.info("No data."); return
    agg = (df.dropna(subset=[COL_TOTAL])
             .groupby(["framework_category", COL_MONTH], as_index=False)[COL_TOTAL].sum())
    color_map = {"Flash Appeal": BURNT_SIENNA,
                 "Minimum Package / SMC Coordination": BLUE_SAPPHIRE,
                 "Unresolved": "#A0A0A0"}
    fig = px.bar(
        agg, x=COL_TOTAL, y="framework_category", color=COL_MONTH,
        orientation="h", barmode="group",
        color_discrete_sequence=[BLUE_SAPPHIRE, BURNT_SIENNA, STEEL_BLUE],
        labels={COL_TOTAL: "Reach / count", "framework_category": ""},
    )
    fig.update_traces(hovertemplate="<b>%{y}</b><br>%{x:,}<extra></extra>")
    fig.update_layout(
        legend=dict(orientation="h", yanchor="top", y=1.18,
                    xanchor="right", x=1.0, font=dict(size=11),
                    title_text=""),
        yaxis=dict(tickfont=dict(size=11)),
    )
    st.plotly_chart(brand_layout(fig, height=280, showlegend=True),
                    use_container_width=True)


def chart_demographics_donut(df: pd.DataFrame):
    cols_present = [c for c in DEMO_COLS if c in df.columns]
    if not cols_present:
        st.info("No demographic columns."); return
    totals = df[cols_present].sum(numeric_only=True)
    totals = totals[totals > 0]
    if totals.empty:
        st.info("No age/sex disaggregation reported for current filters."); return

    # Friendlier labels matching the Somalia dashboard's style
    label_map = {
        "Boys (<18)": "Boys", "Girls (<18)": "Girls",
        "Men (18-59)": "Men", "Women (18-59)": "Women",
        "Elderly Male (≥60)": "Elderly ♂", "Elderly Female (≥60)": "Elderly ♀",
    }
    colors = {
        "Boys": STEEL_BLUE, "Girls": BURNT_SIENNA,
        "Men": BLUE_SAPPHIRE, "Women": "#9B4A36",
        "Elderly ♂": MOONSTONE_BLUE, "Elderly ♀": MOSS_GREEN,
    }
    labels = [label_map.get(k, k) for k in totals.index]
    fig = go.Figure(go.Pie(
        labels=labels, values=totals.values, hole=0.55,
        marker=dict(colors=[colors[l] for l in labels],
                    line=dict(color="white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=11, color=BALTIC_SEA),
        hovertemplate="<b>%{label}</b><br>%{value:,} (%{percent})<extra></extra>",
    ))
    total = int(totals.sum())
    fig.add_annotation(
        text=f"<b>{total:,}</b><br>"
             f"<span style='font-size:9px;'>with disaggregation</span>",
        showarrow=False, font=dict(size=18, color=BLUE_SAPPHIRE),
    )
    st.plotly_chart(brand_layout(fig, height=300, showlegend=False),
                    use_container_width=True)

    male_pwd   = int(df.get("Male-PWD",   pd.Series(dtype=float)).sum(skipna=True))
    female_pwd = int(df.get("Female-PWD", pd.Series(dtype=float)).sum(skipna=True))
    st.markdown(f"""
    <div style="display:flex; justify-content:space-around; padding-top:4px;
                border-top: 1px solid rgba(27,101,124,0.15); margin-top:4px;">
      <div style="text-align:center;">
        <div style="font-size:10px; opacity:0.7; text-transform:uppercase;">PWD ♂</div>
        <div style="font-size:18px; font-weight:700; color:{BLUE_SAPPHIRE};">{male_pwd:,}</div>
      </div>
      <div style="text-align:center;">
        <div style="font-size:10px; opacity:0.7; text-transform:uppercase;">PWD ♀</div>
        <div style="font-size:18px; font-weight:700; color:{BURNT_SIENNA};">{female_pwd:,}</div>
      </div>
      <div style="text-align:center;">
        <div style="font-size:10px; opacity:0.7; text-transform:uppercase;">Total PWD</div>
        <div style="font-size:18px; font-weight:700; color:{BALTIC_SEA};">{male_pwd+female_pwd:,}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)


def chart_partner_matrix(df: pd.DataFrame):
    pa_col = "primary_activity_master"
    pivot = (df.dropna(subset=[pa_col])
               .groupby([COL_ORG, pa_col])["_row_id"].count()
               .unstack(fill_value=0))
    if pivot.empty:
        st.info("No data."); return

    # Shorten long activity labels
    short_labels = {c: (c if len(c) <= 45 else c[:42] + "…") for c in pivot.columns}
    pivot.columns = [short_labels[c] for c in pivot.columns]

    fig = go.Figure(go.Heatmap(
        z=pivot.values,
        x=pivot.columns, y=pivot.index,
        colorscale=CCCM_SEQUENTIAL,
        showscale=True,
        text=pivot.values,
        texttemplate="%{text}",
        textfont=dict(size=10, color=BALTIC_SEA),
        hovertemplate="<b>%{y}</b><br>%{x}<br>Activities: %{z}<extra></extra>",
        colorbar=dict(title=dict(text="Activities", font=dict(size=10)),
                      thickness=10, len=0.7, x=1.02),
    ))
    fig.update_layout(
        xaxis=dict(tickangle=-30, tickfont=dict(size=10), automargin=True),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        margin=dict(l=10, r=10, t=10, b=140),
    )
    st.plotly_chart(brand_layout(fig, height=380, showlegend=False),
                    use_container_width=True)


def chart_neighborhood_treemap(df: pd.DataFrame):
    people = df[df["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)]
    if people.empty:
        st.info("No people-indicator data."); return
    deduped = (people.groupby(["governorate_clean", COL_NBHD, "site_code",
                               COL_MONTH, "primary_activity_master"],
                              dropna=False)[COL_TOTAL].max().reset_index())
    deduped = deduped.rename(columns={
        "governorate_clean": "Governorate", COL_NBHD: "Neighborhood",
        COL_TOTAL: "Reach",
    })
    deduped["Neighborhood"] = deduped["Neighborhood"].fillna("(unspecified)")
    agg = (deduped.groupby(["Governorate", "Neighborhood"], as_index=False)["Reach"].sum())
    agg = agg[agg["Governorate"].notna() & (agg["Reach"] > 0)]
    if agg.empty:
        st.info("No reach to display."); return

    fig = px.treemap(
        agg, path=["Governorate", "Neighborhood"], values="Reach",
        color="Reach", color_continuous_scale=CCCM_SEQUENTIAL,
    )
    fig.update_traces(
        marker=dict(line=dict(color="white", width=2)),
        hovertemplate="<b>%{label}</b><br>Reach: %{value:,}<extra></extra>",
        textfont=dict(family="Inter, Arial", size=12),
    )
    fig.update_layout(coloraxis_showscale=False)
    st.plotly_chart(brand_layout(fig, height=380, showlegend=False),
                    use_container_width=True)


def chart_indicator_topn(df: pd.DataFrame):
    if df.empty:
        st.info("No data."); return
    agg = (df.dropna(subset=[COL_TOTAL])
             .groupby("indicator_clean", as_index=False)
             .agg(reach=(COL_TOTAL, "sum"), entries=("_row_id", "count")))
    agg = agg.sort_values("reach", ascending=True).tail(8)
    if agg.empty:
        st.info("No data."); return
    agg["label"] = agg["indicator_clean"].apply(
        lambda s: s if len(s) <= 55 else s[:52] + "…"
    )
    fig = go.Figure(go.Bar(
        x=agg["reach"], y=agg["label"], orientation="h",
        marker=dict(color=BLUE_SAPPHIRE),
        text=[f"{int(v):,}" for v in agg["reach"]],
        textposition="outside",
        textfont=dict(size=10, color=BALTIC_SEA),
        hovertemplate="<b>%{y}</b><br>Total: %{x:,}<extra></extra>",
    ))
    fig.update_layout(xaxis=dict(showticklabels=False),
                      yaxis=dict(title=None, tickfont=dict(size=10)))
    st.plotly_chart(brand_layout(fig, height=320, showlegend=False),
                    use_container_width=True)


def table_partner_presence(df: pd.DataFrame):
    if df.empty:
        st.info("No data."); return
    tbl = (df.dropna(subset=["site_code"])
             .groupby(COL_ORG)
             .agg(Activities=("_row_id", "count"),
                  Sites=("site_code", pd.Series.nunique),
                  Governorates=("governorate_clean", pd.Series.nunique),
                  Reach=(COL_TOTAL, "sum"))
             .sort_values("Reach", ascending=False)
             .reset_index()
             .rename(columns={COL_ORG: "Partner"}))
    tbl["Reach"] = tbl["Reach"].fillna(0).astype(int)

    # Style with brand colors (matplotlib-free)
    styled = (tbl.style
                .format({"Activities": "{:,}", "Sites": "{:,}",
                         "Governorates": "{:,}", "Reach": "{:,}"})
                .apply(lambda s: brand_gradient(s), subset=["Reach"])
                .set_properties(**{
                    "font-family": "Inter, Arial, sans-serif",
                    "font-size": "12px",
                })
                .set_table_styles([
                    {"selector": "th", "props": [
                        ("background-color", BLUE_SAPPHIRE),
                        ("color", ECRU_WHITE),
                        ("font-weight", "600"),
                        ("text-align", "left"),
                        ("padding", "8px"),
                        ("font-size", "11px"),
                        ("text-transform", "uppercase"),
                        ("letter-spacing", "0.4px"),
                    ]},
                    {"selector": "td", "props": [("padding", "6px 8px")]},
                ]))
    st.dataframe(styled, use_container_width=True, hide_index=True, height=320)


def table_top_sites(df: pd.DataFrame):
    people = df[df["unit_master"].fillna("").str.strip().isin(PEOPLE_UNITS)]
    if people.empty:
        st.info("No people-indicator data."); return
    deduped = (people.groupby(["site_code", COL_SITE, "governorate_clean",
                               COL_MONTH, "primary_activity_master"],
                              dropna=False)[COL_TOTAL].max().reset_index())
    tbl = (deduped.groupby(["site_code", COL_SITE, "governorate_clean"],
                           as_index=False)[COL_TOTAL].sum()
                  .sort_values(COL_TOTAL, ascending=False).head(15))
    tbl = tbl.rename(columns={
        "site_code": "Site ID",
        COL_SITE: "Site Name",
        "governorate_clean": "Governorate",
        COL_TOTAL: "Reach",
    })
    tbl["Reach"] = tbl["Reach"].astype(int)
    # Truncate site name
    tbl["Site Name"] = tbl["Site Name"].apply(
        lambda s: (s[:40] + "…") if isinstance(s, str) and len(s) > 40 else s
    )
    styled = (tbl.style
                .format({"Reach": "{:,}"})
                .apply(lambda s: brand_gradient(s), subset=["Reach"])
                .set_properties(**{"font-family": "Inter, Arial",
                                   "font-size": "11px"})
                .set_table_styles([
                    {"selector": "th", "props": [
                        ("background-color", BLUE_SAPPHIRE),
                        ("color", ECRU_WHITE),
                        ("font-weight", "600"),
                        ("font-size", "10px"),
                        ("padding", "6px"),
                        ("text-transform", "uppercase"),
                    ]},
                    {"selector": "td", "props": [("padding", "5px 6px")]},
                ]))
    st.dataframe(styled, use_container_width=True, hide_index=True, height=380)


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
def main():
    try:
        clean, _errors = load_data()
    except FileNotFoundError as e:
        st.error(
            f"Could not load source files from `{DATA_DIR}`. "
            f"Make sure both xlsx files are in the `data/` folder."
        )
        st.exception(e)
        return

    filt = sidebar_filters(clean)
    header(filt if not filt.empty else clean)
    kpi_ribbon(filt)

    if filt.empty:
        st.warning("No activities match the current filter selection.")
        return

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    # ---- Row 1: Operational footprint ---------------------------------
    r1c1, r1c2, r1c3 = st.columns([1.0, 1.3, 1.0])
    with r1c1:
        panel("People Reached by Governorate")
        chart_reach_by_governorate(filt)
        end_panel()
    with r1c2:
        panel("Framework Performance — Flash Appeal vs Minimum Package")
        chart_framework_split(filt)
        end_panel()
    with r1c3:
        panel("Age & Sex Breakdown")
        chart_demographics_donut(filt)
        end_panel()

    # ---- Row 2: Geographic & partner detail ---------------------------
    r2c1, r2c2 = st.columns([1.4, 1.0])
    with r2c1:
        panel("Reach by Governorate → Neighborhood")
        chart_neighborhood_treemap(filt)
        end_panel()
    with r2c2:
        panel("Partner Presence Summary")
        table_partner_presence(filt)
        end_panel()

    # ---- Row 3: Activity & site detail --------------------------------
    r3c1, r3c2 = st.columns([1.0, 1.0])
    with r3c1:
        panel("Top Indicators by Reach / Count")
        chart_indicator_topn(filt)
        end_panel()
    with r3c2:
        panel("Top Sites by People Reached")
        table_top_sites(filt)
        end_panel()

    # ---- Row 4: Partner × Activity matrix -----------------------------
    panel("Partner × Primary Activity Coverage Matrix")
    chart_partner_matrix(filt)
    end_panel()

    # ---- Footer -------------------------------------------------------
    st.markdown(f"""
    <div class="brand-footer">
      <b>Site Management Cluster — Gaza</b> · 4W Operational Presence Dashboard ·
      Data deduplicated and validated against the SMC Activity Index and OCHA COD-Sites master.
      For methodology questions contact the Cluster IM team.
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
