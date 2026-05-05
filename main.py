"""
=============================================================
  Galderma · Productivity Dashboard
=============================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(layout="wide", page_title="Productivity Analysis")

# ─── WINDOW SIZES (idénticos al R) ──────────────────────────
CURRENT_SIZE  = 3
BASELINE_SIZE = 3
GAP_SIZE      = 3

STATUS_VALIDOS = ["Ready to Deploy", "Closed"]

COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
]

st.title("📊 Productivity Analysis")


# ════════════════════════════════════════════════════════════
#  CARGA DE DATOS
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data(source):
    df = pd.read_excel(source)
    df["Period"] = pd.to_datetime(df["Period"])
    df["Points"] = pd.to_numeric(df["Points"], errors="coerce")
    df = df[df["Status"].isin(STATUS_VALIDOS)].copy()
    df["Grupo"] = "Grupo"
    return df


# ════════════════════════════════════════════════════════════
#  AGREGACIÓN MENSUAL  (réplica del db_agg de R)
# ════════════════════════════════════════════════════════════
def aggregate_monthly(df: pd.DataFrame, dimension: str, metric_col: str) -> pd.DataFrame:
    """
    Mirrors R:
        group_by(Period, var) %>%
        summarise(n=n(), Sum=sum(Target), Mean=mean(Target))
    """
    agg = (
        df.groupby(["Period", dimension], dropna=False)
        .agg(
            n   =(metric_col, "count"),
            Sum =(metric_col, "sum"),
            Mean=(metric_col, "mean"),
        )
        .reset_index()
        .sort_values(["Period", dimension])
        .reset_index(drop=True)
    )
    return agg


# ════════════════════════════════════════════════════════════
#  CORE: fx.PRODUCTIVITY.v3 
# ════════════════════════════════════════════════════════════
def fx_productivity_v3(
    db_agg: pd.DataFrame,
    dimension: str,
    more_is_best: bool,
    selected_values: list = None,
) -> pd.DataFrame:
    """
    Ventanas Python (0-based) ↔ R (1-based):
      current   [0 : CURRENT_SIZE]
      baseline  [CURRENT+GAP : CURRENT+GAP+BASELINE]  (o fallback)
    signo = +1 si more_is_best, -1 si no.
    """
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
            cur_start, cur_end = 0, CURRENT_SIZE

            if has_baseline_full:
                base_start = CURRENT_SIZE + GAP_SIZE
                base_end   = CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE
            else:
                base_start = max(0, n - BASELINE_SIZE)
                base_end   = n

            current_window  = svc_data.iloc[cur_start:cur_end]
            baseline_window = svc_data.iloc[base_start:base_end]

            effort_data     = current_window["Sum"].sum()
            units_data      = current_window["n"].sum()
            effort_baseline = baseline_window["Sum"].sum()
            units_baseline  = baseline_window["n"].sum()

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
#  MODOS DE ANÁLISIS
# ════════════════════════════════════════════════════════════
def calc_individual_productivity(db_agg, dimension, more_is_best, selected_values):
    """R: Recursive mode — un cálculo independiente por valor."""
    results = []
    for val in selected_values:
        res = fx_productivity_v3(db_agg, dimension, more_is_best, [val])
        if not res.empty:
            res[dimension] = str(val)
            results.append(res)
    return pd.concat(results, ignore_index=True) if results else pd.DataFrame()


def calc_global_productivity(db_agg, dimension, more_is_best, selected_values):
    """R: Single/Global mode — todos los valores combinados."""
    res = fx_productivity_v3(db_agg, dimension, more_is_best, selected_values)
    if not res.empty:
        res[dimension] = "Group Total"
    return res


# ════════════════════════════════════════════════════════════
#  GRÁFICAS 
# ════════════════════════════════════════════════════════════
XAXIS_STYLE = dict(title="Period", tickformat="%b %Y", dtick="M1", tickangle=45)


def make_count_chart(db_agg, dimension, selected_values):
    """p1: Count (n) por período con etiquetas en puntos."""
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
        xaxis=XAXIS_STYLE,
        yaxis_title="Count (n)",
        height=420,
        hovermode="x unified",
    )
    return fig


def make_mean_chart(db_agg, dimension, metric_col, selected_values):
    """p2: Mean del metric por período con etiquetas en puntos."""
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
        xaxis=XAXIS_STYLE,
        yaxis_title=f"Mean {metric_col}",
        height=420,
        hovermode="x unified",
    )
    return fig


def make_productivity_chart(prod_df, dimension):
    """p3: Productivity (%) por período + línea cero de referencia."""
    fig = go.Figure()
    values_in_df = prod_df[dimension].unique() if dimension in prod_df.columns else ["Group Total"]
    for i, val in enumerate(values_in_df):
        sub = (prod_df[prod_df[dimension] == str(val)] if dimension in prod_df.columns
               else prod_df).sort_values("ActualPeriod")
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
        xaxis=XAXIS_STYLE,
        yaxis_title="Productivity (%)",
        height=420,
        hovermode="x unified",
    )
    return fig


def make_velocity_chart(prod_df, dimension, metric_col):
    """p4: Real vs Expected (línea sólida vs punteada) por período."""
    fig = go.Figure()
    values_in_df = prod_df[dimension].unique() if dimension in prod_df.columns else ["Group Total"]
    styles = [
        ("EffortData",     "Real",     "solid"),
        ("BaseEfforEquiv", "Expected", "dash"),
    ]
    for i, val in enumerate(values_in_df):
        sub = (prod_df[prod_df[dimension] == str(val)] if dimension in prod_df.columns
               else prod_df).sort_values("ActualPeriod")
        base_color = COLORS[i % len(COLORS)]
        for col, label, dash in styles:
            fig.add_trace(go.Scatter(
                x=sub["ActualPeriod"], y=sub[col],
                mode="lines+markers",
                name=f"{val} — {label} {metric_col}",
                line=dict(color=base_color, width=2, dash=dash),
                marker=dict(size=5),
            ))
    fig.add_hline(y=0, line_dash="dash", line_color="black", line_width=1.5)
    fig.update_layout(
        title=f"Velocity: Real vs Expected {metric_col}",
        xaxis=XAXIS_STYLE,
        yaxis_title=metric_col,
        height=420,
        hovermode="x unified",
    )
    return fig


# ════════════════════════════════════════════════════════════
#  UI PRINCIPAL
# ════════════════════════════════════════════════════════════
uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

DEFAULT_PATH = Path("Galderma_12-01-24_to_03-31-26.xlsx")

if uploaded_file:
    df = load_data(uploaded_file)
elif DEFAULT_PATH.exists():
    df = load_data(str(DEFAULT_PATH))
    st.info(f"📄 Usando archivo local: {DEFAULT_PATH.name}  —  {len(df):,} filas tras filtro")
else:
    st.info("⬆️ Upload an Excel file to begin.")
    st.stop()

# ── Sidebar ──────────────────────────────────────────────────
st.sidebar.header("Controls")

DIMENSIONS = ["Developer", "QA Tester", "Grupo"]
available  = [d for d in DIMENSIONS if d in df.columns]
dimension  = st.sidebar.selectbox("Analyze by", available)

df[dimension] = df[dimension].astype(str)
values = sorted(df[dimension].dropna().unique().tolist())

selected_values = st.sidebar.multiselect(
    "Select values",
    values,
    default=values[:3] if len(values) >= 3 else values,
)

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

# ── Agregación (paso 1) ────────────────────────────────
df_filtered = df[df[dimension].isin(selected_values)].copy()
db_agg = aggregate_monthly(df_filtered, dimension, "Points")
db_agg[dimension] = db_agg[dimension].astype(str)

# ── Productividad (paso 2) ─────────────────────────────
if "Individual" in analysis_mode:
    prod_df = calc_individual_productivity(db_agg, dimension, True, selected_values)
else:
    prod_df = calc_global_productivity(db_agg, dimension, True, selected_values)

# ── Gráficas ─────────────────────────────────────────────────
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
        st.plotly_chart(make_productivity_chart(prod_df, dimension), use_container_width=True)

    if "Velocity (Real vs Expected)" in show_charts:
        st.subheader("⚡ Velocity: Real vs Expected")
        st.caption(
            "Real = sum of Points in current window  |  "
            "Expected = what baseline EpU predicts for current volume"
        )
        st.plotly_chart(make_velocity_chart(prod_df, dimension, "Points"), use_container_width=True)

if "Count over Time" in show_charts:
    st.subheader("🔢 Ticket Count Over Time")
    st.plotly_chart(make_count_chart(db_agg, dimension, selected_values), use_container_width=True)

if "Mean over Time" in show_charts:
    st.subheader("📊 Mean Points Over Time")
    st.plotly_chart(make_mean_chart(db_agg, dimension, "Points", selected_values), use_container_width=True)

# ── Tablas (expanders) ───────────────────────────────
with st.expander("📋 Aggregated Monthly Data (db_agg)", expanded=False):
    st.dataframe(db_agg.sort_values(["Period", dimension]), use_container_width=True)

if not prod_df.empty:
    with st.expander("📋 Productivity Results", expanded=False):
        display = prod_df.copy()
        display["Productivity %"] = (display["Value"] * 100).map(
            lambda x: f"{x:.4f}%" if pd.notna(x) else ""
        )
        st.dataframe(display.sort_values("ActualPeriod"), use_container_width=True)
