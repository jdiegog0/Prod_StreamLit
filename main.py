"""
=============================================================
  Productivity Dashboard — Unified Model
  Modelos:
    · Galderma  → More is Best (Points)
    · AMS       → Less is Best (Effort)
=============================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(layout="wide", page_title="Productivity Analysis")

# ─── WINDOW SIZES ────────────────────────────────────────────
CURRENT_SIZE  = 3
BASELINE_SIZE = 3
GAP_SIZE      = 3

COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
]

XAXIS_STYLE = dict(title="Period", tickformat="%b %Y", dtick="M1", tickangle=45)

st.title("📊 Productivity Analysis")


# ════════════════════════════════════════════════════════════
#  DATA LOAD FOR GALDERMA MODEL (Points / More is Best)
# ════════════════════════════════════════════════════════════
def load_galderma(uploaded_file):

    xls  = pd.ExcelFile(uploaded_file)
    sname = "RawData" if "RawData" in xls.sheet_names else xls.sheet_names[0]
    df   = pd.read_excel(uploaded_file, sheet_name=sname)
    df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)

    df["Points"] = pd.to_numeric(df["Points"], errors="coerce")
    df["Period"] = pd.to_datetime(df["Period"], errors="coerce").dt.to_period("M").dt.to_timestamp()
    df = df[df["Status"].isin(["Ready to Deploy", "Closed"])].copy()
    df["Grupo"] = "Grupo"

    # Split Developer por "/"
    if "Developer" in df.columns:
        df["Developer"] = df["Developer"].astype("string").str.strip()
        df = df[df["Developer"].notna() & (df["Developer"] != "")].copy()
        df = df.assign(Developer=df["Developer"].str.split("/")).explode("Developer")
        df["Developer"] = df["Developer"].astype("string").str.strip()
        df = df[
            df["Developer"].notna()
            & (df["Developer"] != "")
            & (df["Developer"].str.lower() != "nan")
        ].copy()

    ## df = df[df["Period"].notna() & df["Points"].notna()].copy() ## Blank or NAN filter for Modelo Galderma
    # ════════════════════════════════════════════════════════════
    config = {
        "metric_col":   "Points",
        "more_is_best": True,
        "dimensions":   [c for c in ["Developer", "QA Tester", "Grupo"] if c in df.columns],
        "label_real":   "Real Points",
        "label_exp":    "Expected Points",
    }
    return df, config


# ════════════════════════════════════════════════════════════
#  DATA LOAD FOR AMS MODEL (Effort / Less is Best)
# ════════════════════════════════════════════════════════════
def detect_columns(df: pd.DataFrame) -> dict:
    
    col_map = {
        "Assigned To": None, "Group": None, "WBS": None,
        "Category": None, "Service Type": None, "EndDate": None,
        "Effort": None, "Points": None, "Developer": None,
        "Status": None, "Period": None, "QA Tester": None,
        "Issue Type": None, "Priority": None,
    }
    for col in df.columns:
        c = col.lower().strip()
        if   c in ["assigned to", "assignee", "resource", "is"]:  col_map["Assigned To"]  = col
        elif c == "group":                                          col_map["Group"]        = col
        elif c == "wbs":                                            col_map["WBS"]          = col
        elif c == "category":                                       col_map["Category"]     = col
        elif c in ["service type", "servicetype"]:                  col_map["Service Type"] = col
        elif c in ["enddate", "end date"]:                         col_map["EndDate"]      = col
        elif c == "effort":                                         col_map["Effort"]       = col
        elif c == "points":                                         col_map["Points"]       = col
        elif c == "developer":                                      col_map["Developer"]    = col
        elif c == "status":                                         col_map["Status"]       = col
        elif c == "period":                                         col_map["Period"]       = col
        elif c in ["qa tester", "qatester"]:                       col_map["QA Tester"]    = col
        elif c in ["issue type", "issuetype"]:                     col_map["Issue Type"]   = col
        elif c == "priority":                                       col_map["Priority"]     = col
    return col_map


def load_ams(uploaded_file):
   
    xls   = pd.ExcelFile(uploaded_file)
    sname = "RawData" if "RawData" in xls.sheet_names else xls.sheet_names[0]
    df    = pd.read_excel(uploaded_file, sheet_name=sname)
    df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)

    col_map = detect_columns(df)

    required = ["Assigned To", "Group", "WBS", "EndDate", "Effort"]
    missing  = [k for k in required if col_map[k] is None]
    if missing:
        return None, f"Missing required columns for AMS model: {missing}"

    df = df.rename(columns={v: k for k, v in col_map.items() if v is not None}).copy()

    df["EndDate"] = pd.to_datetime(df["EndDate"], errors="coerce")
    df["Period"]  = df["EndDate"].dt.to_period("M").dt.to_timestamp()
    df["Effort"]  = pd.to_numeric(df["Effort"], errors="coerce")
    ## df = df[df["Period"].notna() & df["Effort"].notna()].copy()  ## Blank or NAN filter for AMS Model
    # ════════════════════════════════════════════════════════════

    config = {
        "metric_col":   "Effort",
        "more_is_best": False,
        "dimensions":   [c for c in ["Assigned To", "Group", "WBS", "Category", "Service Type"]
                         if c in df.columns],
        "label_real":   "Real Effort",
        "label_exp":    "Expected Effort",
    }
    return df, config


# ════════════════════════════════════════════════════════════
#  MONTHLY AGGREGATION
# ════════════════════════════════════════════════════════════
def aggregate_monthly(df: pd.DataFrame, dimension: str, metric_col: str) -> pd.DataFrame:

    agg = (
        df.groupby(["Period", dimension], dropna=False)
        .agg(n=(metric_col, "size"), # Change from "count" to "size" to include the blank records
             Sum=(metric_col, "sum"), 
             Mean=(metric_col, "mean")) 
        .reset_index()
        .sort_values(["Period", dimension])
        .reset_index(drop=True)
    )
    return agg


# ════════════════════════════════════════════════════════════
#  PRODUCTIVITY CALCULATION
# ════════════════════════════════════════════════════════════
def fx_productivity_v3(db_agg, dimension, more_is_best, selected_values=None):

    signo = 1 if more_is_best else -1
    if selected_values is not None:
        db_agg = db_agg[db_agg[dimension].isin(selected_values)].copy()

    fechas = sorted(db_agg["Period"].unique())
    rows   = []

    for current_period in fechas:
        subset_all = db_agg[db_agg["Period"] <= current_period].copy()
        services   = subset_all[dimension].unique()

        period_effort_data = 0.0
        period_base_equiv  = 0.0
        any_calc = False

        for svc in services:
            svc_data = (
                subset_all[subset_all[dimension] == svc]
                .sort_values("Period", ascending=False)
                .reset_index(drop=True)
            )
            n         = len(svc_data)
            max_fecha = svc_data["Period"].max()

            if n < CURRENT_SIZE or current_period > max_fecha:
                continue

            has_baseline_full = n >= (CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE)
            cur_s, cur_e = 0, CURRENT_SIZE

            if has_baseline_full:
                bl_s = CURRENT_SIZE + GAP_SIZE
                bl_e = CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE
            else:
                bl_s = max(0, n - BASELINE_SIZE)
                bl_e = n

            cw = svc_data.iloc[cur_s:cur_e]
            bw = svc_data.iloc[bl_s:bl_e]

            effort_data     = cw["Sum"].sum()
            units_data      = cw["n"].sum()
            effort_baseline = bw["Sum"].sum()
            units_baseline  = bw["n"].sum()

            if units_baseline == 0 or units_data == 0:
                continue

            epu_bl     = effort_baseline / units_baseline
            base_equiv = epu_bl * units_data

            period_effort_data += effort_data
            period_base_equiv  += base_equiv
            any_calc = True

        if not any_calc or period_base_equiv == 0:
            continue

        productivity = ((period_effort_data - period_base_equiv) / period_base_equiv) * signo
        rows.append({
            "ActualPeriod":   current_period,
            "EffortData":     period_effort_data,
            "BaseEfforEquiv": period_base_equiv,
            "Value":          productivity,
        })

    return pd.DataFrame(rows)


# ════════════════════════════════════════════════════════════
#  PRODUCTIVITY CALCULATION #2
# ════════════════════════════════════════════════════════════
def calc_individual_productivity(db_agg, dimension, more_is_best, selected_values):
    
    results = []
    for val in selected_values:
        res = fx_productivity_v3(db_agg, dimension, more_is_best, [val])
        if not res.empty:
            res[dimension] = str(val)
            results.append(res)
    return pd.concat(results, ignore_index=True) if results else pd.DataFrame()


def calc_global_productivity(db_agg, dimension, more_is_best, selected_values):
    
    res = fx_productivity_v3(db_agg, dimension, more_is_best, selected_values)
    if not res.empty:
        res[dimension] = "Group Total"
    return res


# ════════════════════════════════════════════════════════════
#  GRAFICS
# ════════════════════════════════════════════════════════════
def make_count_chart(db_agg, dimension, selected_values):
    fig = go.Figure()
    for i, val in enumerate(selected_values):
        sub = db_agg[db_agg[dimension] == str(val)].sort_values("Period")
        if sub.empty:
            continue
        fig.add_trace(go.Scatter(
            x=sub["Period"], y=sub["n"],
            mode="lines+markers+text",
            name=str(val),
            text=[f"{v:.0f}" for v in sub["n"]],
            textposition="top center",
            line=dict(color=COLORS[i % len(COLORS)], width=2),
            marker=dict(size=6),
        ))
    fig.update_layout(
        title=f"{dimension} — Ticket Count Over Time",
        xaxis=XAXIS_STYLE, yaxis_title="Count (n)",
        height=420, hovermode="x unified",
    )
    return fig


def make_mean_chart(db_agg, dimension, metric_col, selected_values):
    fig = go.Figure()
    for i, val in enumerate(selected_values):
        sub = db_agg[db_agg[dimension] == str(val)].sort_values("Period")
        if sub.empty:
            continue
        fig.add_trace(go.Scatter(
            x=sub["Period"], y=sub["Mean"],
            mode="lines+markers+text",
            name=str(val),
            text=[f"{v:.2f}" for v in sub["Mean"]],
            textposition="top center",
            line=dict(color=COLORS[i % len(COLORS)], width=2),
            marker=dict(size=6),
        ))
    fig.update_layout(
        title=f"{dimension} — Mean {metric_col} Over Time",
        xaxis=XAXIS_STYLE, yaxis_title=f"Mean {metric_col}",
        height=420, hovermode="x unified",
    )
    return fig


def make_productivity_chart(prod_df, dimension, label_real, label_exp):
    fig = go.Figure()
    vals = prod_df[dimension].unique() if dimension in prod_df.columns else ["Group Total"]
    for i, val in enumerate(vals):
        sub = (prod_df[prod_df[dimension] == str(val)]
               if dimension in prod_df.columns else prod_df).sort_values("ActualPeriod")
        prod_pct = sub["Value"] * 100
        fig.add_trace(go.Scatter(
            x=sub["ActualPeriod"], y=prod_pct,
            mode="lines+markers+text",
            name=str(val),
            text=[f"{v:.1f}%" if pd.notna(v) else "" for v in prod_pct],
            textposition="top center",
            line=dict(color=COLORS[i % len(COLORS)], width=2),
            marker=dict(size=6),
        ))
    fig.add_hline(y=0, line_dash="dash", line_color="black", line_width=1.5)
    fig.update_layout(
        title=f"Productivity Over Time by {dimension}",
        xaxis=XAXIS_STYLE, yaxis_title="Productivity (%)",
        height=420, hovermode="x unified",
    )
    return fig


def make_velocity_chart(prod_df, dimension, metric_col, label_real, label_exp):
    fig = go.Figure()
    vals   = prod_df[dimension].unique() if dimension in prod_df.columns else ["Group Total"]
    styles = [("EffortData", label_real, "solid"), ("BaseEfforEquiv", label_exp, "dash")]
    for i, val in enumerate(vals):
        sub = (prod_df[prod_df[dimension] == str(val)]
               if dimension in prod_df.columns else prod_df).sort_values("ActualPeriod")
        base_color = COLORS[i % len(COLORS)]
        for col, lbl, dash in styles:
            fig.add_trace(go.Scatter(
                x=sub["ActualPeriod"], y=sub[col],
                mode="lines+markers",
                name=f"{val} — {lbl}",
                line=dict(color=base_color, width=2, dash=dash),
                marker=dict(size=5),
            ))
    fig.add_hline(y=0, line_dash="dash", line_color="black", line_width=1.5)
    fig.update_layout(
        title=f"Velocity: {label_real} vs {label_exp}",
        xaxis=XAXIS_STYLE, yaxis_title=metric_col,
        height=420, hovermode="x unified",
    )
    return fig


# ════════════════════════════════════════════════════════════
#  UI PRINCIPAL
# ════════════════════════════════════════════════════════════


st.sidebar.header("⚙️ Model")
model_choice = st.sidebar.radio(
    "Select productivity model",
    ["🟢  Galderma · Points  (More is Best)", "🔵  AMS · Effort  (Less is Best)"],
    help=(
        "Galderma: Tracks story points delivered by developers.\n"
        "AMS: Tracks effort (hours) consumed per ticket — lower is better."
    ),
)
is_ams = model_choice.startswith("🔵")

st.sidebar.markdown("---")
st.sidebar.header("📂 Data")
uploaded_file = st.file_uploader(
    f"Upload {'AMS' if is_ams else 'Galderma'} Excel file (.xlsx)",
    type=["xlsx"],
)

# ── Data Load ──────────────────────────────
df, config = None, None

if uploaded_file:
    if is_ams:
        result = load_ams(uploaded_file)
        if result[0] is None:
            st.error(result[1])
            st.stop()
        df, config = result
    else:
        df, config = load_galderma(uploaded_file)
else:
    # Local Files
    galderma_path = Path("Galderma_12-01-24_to_03-31-26.xlsx")
    ams_path      = Path("DataForPythonAMS.xlsx")

    if not is_ams and galderma_path.exists():
        df, config = load_galderma(str(galderma_path))
        st.info(f"📄 Usando archivo local: {galderma_path.name}  —  {len(df):,} filas")
    elif is_ams and ams_path.exists():
        result = load_ams(str(ams_path))
        if result[0] is None:
            st.error(result[1])
            st.stop()
        df, config = result
        st.info(f"📄 Local File: {ams_path.name}  —  {len(df):,} filas")
    else:
        st.info("⬆️  Upload an Excel file to begin.")
        st.stop()

# ── Badge de modelo activo ───────────────────────────────────
if is_ams:
    st.markdown(
        "🔵 **AMS MODEL** — LESS IS BEST (EFFORT) &nbsp;|&nbsp; "
        f"Metric: `{config['metric_col']}`"
    )
else:
    st.markdown(
        "🟢 **GALDERMA MODEL** — MORE IS BEST (POINTS) &nbsp;|&nbsp; "
        f"Metric: `{config['metric_col']}`"
    )

# ── Controls  ────────────────────────────────────────
st.sidebar.header("Controls")

dimension = st.sidebar.selectbox("Analyze by", config["dimensions"])

df[dimension] = df[dimension].astype(str)
values = sorted(df[dimension].dropna().unique().tolist())

default_sel = values[:3] if len(values) >= 3 else values
selected_values = st.sidebar.multiselect("Select values", values, default=default_sel)

analysis_mode = st.sidebar.radio(
    "Analysis mode",
    ["Individual (one series per value)", "Global (combined into one series)"],
    help="Individual = R Recursive mode.  Global = R Single mode.",
)

show_charts = st.sidebar.multiselect(
    "Charts to show",
    ["Productivity", "Velocity (Real vs Expected)", "Count over Time", "Mean over Time"],
    default=["Productivity", "Velocity (Real vs Expected)"],
)

if not selected_values:
    st.warning("Select at least one value.")
    st.stop()

# ── Agregación ────────────────────────────────────────────────
df_filtered = df[df[dimension].isin(selected_values)].copy()
db_agg = aggregate_monthly(df_filtered, dimension, config["metric_col"])
db_agg[dimension] = db_agg[dimension].astype(str)

# ── Productividad ─────────────────────────────────────────────
if "Individual" in analysis_mode:
    prod_df = calc_individual_productivity(
        db_agg, dimension, config["more_is_best"], selected_values
    )
else:
    prod_df = calc_global_productivity(
        db_agg, dimension, config["more_is_best"], selected_values
    )

# ── Gráficas ──────────────────────────────────────────────────
if prod_df.empty:
    st.warning(
        f"⚠️ Not enough historical data to calculate productivity. "
        f"Each value needs at least {CURRENT_SIZE} periods."
    )
else:
    if "Productivity" in show_charts:
        st.subheader("📈 Productivity Over Time")
        st.caption(
            "Positive = better than baseline  |  Negative = worse than baseline  |  "
            "Zero line = baseline level"
        )
        st.plotly_chart(
            make_productivity_chart(prod_df, dimension, config["label_real"], config["label_exp"]),
            use_container_width=True,
        )

    if "Velocity (Real vs Expected)" in show_charts:
        st.subheader("⚡ Velocity: Real vs Expected")
        st.caption(
            f"{config['label_real']} = sum of {config['metric_col']} in current window  |  "
            f"{config['label_exp']} = what baseline EpU predicts for current volume"
        )
        st.plotly_chart(
            make_velocity_chart(
                prod_df, dimension, config["metric_col"],
                config["label_real"], config["label_exp"]
            ),
            use_container_width=True,
        )

if "Count over Time" in show_charts:
    st.subheader("🔢 Ticket Count Over Time")
    st.plotly_chart(
        make_count_chart(db_agg, dimension, selected_values),
        use_container_width=True,
    )

if "Mean over Time" in show_charts:
    st.subheader(f"📊 Mean {config['metric_col']} Over Time")
    st.plotly_chart(
        make_mean_chart(db_agg, dimension, config["metric_col"], selected_values),
        use_container_width=True,
    )

# ── Tablas (expanders) ────────────────────────────────────────
with st.expander("📋 Aggregated Monthly Data (db_agg)", expanded=False):
    st.dataframe(db_agg.sort_values(["Period", dimension]), use_container_width=True)

if not prod_df.empty:
    with st.expander("📋 Productivity Results", expanded=False):
        display = prod_df.copy()
        display["Productivity %"] = (display["Value"] * 100).map(
            lambda x: f"{x:.4f}%" if pd.notna(x) else ""
        )
        st.dataframe(display.sort_values("ActualPeriod"), use_container_width=True)
