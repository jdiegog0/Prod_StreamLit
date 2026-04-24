# ============================================================
# PRODUCTIVITY ANALYSIS - Streamlit App
# ============================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(layout="wide", page_title="Productivity Analysis")


CURRENT_SIZE  = 3   
GAP_SIZE      = 3   
BASELINE_SIZE = 3   

st.title("📊 Productivity Analysis")


# ============================================================
# FILE READING
# ============================================================

def load_excel_first_or_rawdata(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = "RawData" if "RawData" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


# ============================================================
# COLUMN DETECTION
# ============================================================

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)
    return df


def detect_columns(df: pd.DataFrame) -> dict:
    column_map = {
        "Assigned To": None, "Group": None, "WBS": None,
        "Category": None, "Service Type": None, "EndDate": None,
        "Effort": None, "Points": None, "Developer": None,
        "Status": None, "Period": None, "QA Tester": None,
        "Issue Type": None, "Priority": None
    }
    for col in df.columns:
        c = col.lower().strip()
        if c in ["assigned to", "assignee", "resource", "is"]:
            column_map["Assigned To"] = col
        elif c == "group":
            column_map["Group"] = col
        elif c == "wbs":
            column_map["WBS"] = col
        elif c == "category":
            column_map["Category"] = col
        elif c in ["service type", "servicetype"]:
            column_map["Service Type"] = col
        elif c in ["enddate", "end date"]:
            column_map["EndDate"] = col
        elif c == "effort":
            column_map["Effort"] = col
        elif c == "points":
            column_map["Points"] = col
        elif c == "developer":
            column_map["Developer"] = col
        elif c == "status":
            column_map["Status"] = col
        elif c == "period":
            column_map["Period"] = col
        elif c in ["qa tester", "qatester"]:
            column_map["QA Tester"] = col
        elif c in ["issue type", "issuetype"]:
            column_map["Issue Type"] = col
        elif c == "priority":
            column_map["Priority"] = col
    return column_map


# ============================================================
# MODEL: "LESS IS BEST" (Effort)
# ============================================================

def prepare_effort_model(df: pd.DataFrame, column_map: dict):
    required = ["Assigned To", "Group", "WBS", "EndDate", "Effort"]
    missing = [k for k in required if column_map[k] is None]
    if missing:
        return None, f"Missing required columns: {missing}"

    df = df.rename(columns={v: k for k, v in column_map.items() if v is not None}).copy()
    df["EndDate"] = pd.to_datetime(df["EndDate"], errors="coerce")
    df["Period"]  = df["EndDate"].dt.to_period("M").dt.to_timestamp()
    df["Effort"]  = pd.to_numeric(df["Effort"], errors="coerce")
    df = df[df["Period"].notna() & df["Effort"].notna()].copy()

    config = {
        "metric_col":   "Effort",
        "more_is_best": False,
        "dimensions":   [c for c in ["Assigned To", "Group", "WBS", "Category", "Service Type"] if c in df.columns],
    }
    return df, config


# ============================================================
# MODEL: "MORE IS BEST" (Points)
# ============================================================

def prepare_points_model(df: pd.DataFrame, column_map: dict):
    required = ["Points", "Developer", "Status", "Period"]
    missing = [k for k in required if column_map[k] is None]
    if missing:
        return None, f"Missing required columns: {missing}"

    df = df.rename(columns={v: k for k, v in column_map.items() if v is not None}).copy()
    df["Points"] = pd.to_numeric(df["Points"], errors="coerce")
    df["Period"] = pd.to_datetime(df["Period"], errors="coerce").dt.to_period("M").dt.to_timestamp()

  
    df = df[df["Status"].isin(["Ready to Deploy", "Closed"])].copy()

   
    df["Group"] = "Group"

    
    df["Developer"] = df["Developer"].astype("string").str.strip()
    df = df[df["Developer"].notna() & (df["Developer"] != "")].copy()
    df = df.assign(Developer=df["Developer"].str.split("/")).explode("Developer")
    df["Developer"] = df["Developer"].astype("string").str.strip()
    df = df[
        df["Developer"].notna()
        & (df["Developer"] != "")
        & (df["Developer"].str.lower() != "nan")
    ].copy()

    df = df[df["Period"].notna() & df["Points"].notna()].copy()

    config = {
        "metric_col":   "Points",
        "more_is_best": True,
        "dimensions":   [c for c in ["Developer", "Group", "QA Tester", "Priority", "Issue Type"] if c in df.columns],
    }
    return df, config


# ============================================================
# STEP 1 — MONTHLY AGGREGATION
# ============================================================

def aggregate_monthly(df: pd.DataFrame, dimension: str, metric_col: str) -> pd.DataFrame:
    """
    Mirrors R's db_agg:
        group_by(Period, var) %>%
        summarise(n=n(), Sum=sum(Target), Mean=mean(Target))

    Returns columns: Period, <dimension>, n, Sum, Mean
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


# ============================================================
# STEP 2 — PRODUCTIVITY CALCULATION
# ============================================================

def fx_productivity_v3(
    db_agg: pd.DataFrame,
    dimension: str,
    more_is_best: bool,
    selected_values: list = None,
) -> pd.DataFrame:
    """
    Equivalent to R's fx.PRODUCTIVITY.v3.

    Parameters
    ----------
    db_agg : aggregated DataFrame with columns [Period, dimension, n, Sum, Mean].
             One row per (Period, dimension-value). Equivalent to R's db_agg
             passed after group_by(Period, var) → summarise(n, Sum, Mean).
    dimension : grouping column (e.g. "Developer", "Assigned To")
    more_is_best : True → higher metric is better; False → lower is better
    selected_values : if provided, restricts which services are iterated.
                      Caller is responsible for the nrow >= 5 guard (R's
                      if(nrow(db_agg) >= 5)) before calling this function.

    Returns
    -------
    DataFrame: ActualPeriod, EffortData, BaseEfforEquiv, Value
    """
    signo = 1 if more_is_best else -1

    if selected_values is not None:
        db_agg = db_agg[db_agg[dimension].isin(selected_values)].copy()

    fechas = sorted(db_agg["Period"].unique())
    rows = []

    for current_period in fechas:

        
        subset_all = db_agg[db_agg["Period"] <= current_period].copy()
        services = subset_all[dimension].unique()

        period_effort_data    = 0.0
        period_base_equiv     = 0.0
        any_calc = False

        for svc in services:
            svc_data = (
                subset_all[subset_all[dimension] == svc]
                .sort_values("Period", ascending=False)
                .reset_index(drop=True)
            )
            n = len(svc_data)

            
            max_fecha = svc_data["Period"].max()
            if n < CURRENT_SIZE or current_period > max_fecha:
                continue

            
            has_baseline_full = n >= (CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE)

            cur_start  = 0
            cur_end    = CURRENT_SIZE 

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

            
            epu_bl = effort_baseline / units_baseline

            
            base_equiv = epu_bl * units_data

            period_effort_data += effort_data
            period_base_equiv  += base_equiv
            any_calc = True

        if not any_calc or period_base_equiv == 0:
            continue

        
        productivity = ((period_effort_data - period_base_equiv) / period_base_equiv) * signo

        rows.append({
            "ActualPeriod":   current_period,
            "EffortData":     period_effort_data,   # Real Effort/Points 
            "BaseEfforEquiv": period_base_equiv,    # Expected based on baseline EpU
            "Value":          productivity,          # Productivity index
        })

    return pd.DataFrame(rows)


# ============================================================
# INDIVIDUAL MODE
# ============================================================

MIN_PERIODS = 5   


def calc_individual_productivity(
    db_agg: pd.DataFrame,
    dimension: str,
    more_is_best: bool,
    selected_values: list,
) -> pd.DataFrame:
    """
    Runs fx_productivity_v3 separately per selected dimension value.
    Equivalent to R's individual loop (Recursive mode).

    Applies R's nrow(db_agg) >= 5 guard per value before calculating.
    """
    all_results = []
    for val in selected_values:
        sub_agg = db_agg[db_agg[dimension] == str(val)]

        
        if len(sub_agg) < MIN_PERIODS:
            continue

        res = fx_productivity_v3(
            db_agg=sub_agg,
            dimension=dimension,
            more_is_best=more_is_best,
            selected_values=[val],
        )
        if not res.empty:
            res[dimension] = str(val)
            all_results.append(res)

    return pd.concat(all_results, ignore_index=True) if all_results else pd.DataFrame()


# ============================================================
# GLOBAL MODE
# ============================================================

def calc_global_productivity(
    db_agg: pd.DataFrame,
    dimension: str,
    more_is_best: bool,
    selected_values: list,
) -> pd.DataFrame:
    """
    Runs fx_productivity_v3 on the combined selected values.
    Equivalent to R's Single/Global mode.

    Uses the full db_agg (all selected values together) so that
    multiple services aggregate into one productivity series —
    identical to R passing db_i (unfiltered) to generate_report.
    """
    
    if len(db_agg) < MIN_PERIODS:
        return pd.DataFrame()

    res = fx_productivity_v3(
        db_agg=db_agg,
        dimension=dimension,
        more_is_best=more_is_best,
        selected_values=selected_values,
    )
    if not res.empty:
        res[dimension] = "Group Total"
    return res


# ============================================================
# CHART HELPERS
# R charts:
#   p1 → Count over Time (n)       ← db_agg
#   p2 → Mean over Time (Mean)     ← db_agg
#   p3 → Productivity (Value)      ← Output_Caso
#   p4 → Velocity: EffortData vs BaseEfforEquiv ← Output_Caso (pivoted)
# ============================================================

COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
]


def make_count_chart(db_agg: pd.DataFrame, dimension: str, selected_values: list) -> go.Figure:
    """p1 equivalent: Count (n) over time per dimension value."""
    fig = go.Figure()
    for i, val in enumerate(selected_values):
        sub = db_agg[db_agg[dimension] == str(val)].sort_values("Period")
        if sub.empty:
            continue
        color = COLORS[i % len(COLORS)]
        fig.add_trace(go.Scatter(
            x=sub["Period"], y=sub["n"],
            mode="lines+markers+text",
            name=str(val),
            text=[f"{v:.0f}" for v in sub["n"]],
            textposition="top center",
            line=dict(color=color, width=2),
            marker=dict(size=6),
        ))
    fig.update_layout(
        title=f"{dimension} — Ticket Count Over Time",
        xaxis=dict(title="Period", tickformat="%b %Y", dtick="M1", tickangle=45),
        yaxis_title="Count (n)",
        height=420,
        hovermode="x unified",
    )
    return fig


def make_mean_chart(db_agg: pd.DataFrame, dimension: str, metric_col: str, selected_values: list) -> go.Figure:
    """p2 equivalent: Mean metric over time per dimension value."""
    fig = go.Figure()
    for i, val in enumerate(selected_values):
        sub = db_agg[db_agg[dimension] == str(val)].sort_values("Period")
        if sub.empty:
            continue
        color = COLORS[i % len(COLORS)]
        fig.add_trace(go.Scatter(
            x=sub["Period"], y=sub["Mean"],
            mode="lines+markers+text",
            name=str(val),
            text=[f"{v:.2f}" for v in sub["Mean"]],
            textposition="top center",
            line=dict(color=color, width=2),
            marker=dict(size=6),
        ))
    fig.update_layout(
        title=f"{dimension} — Mean {metric_col} Over Time",
        xaxis=dict(title="Period", tickformat="%b %Y", dtick="M1", tickangle=45),
        yaxis_title=f"Mean {metric_col}",
        height=420,
        hovermode="x unified",
    )
    return fig


def make_productivity_chart(prod_df: pd.DataFrame, dimension: str, y_dtick: int = 25) -> go.Figure:
    """p3 equivalent: Productivity (Value) over time + zero reference line."""
    fig = go.Figure()
    values_in_df = prod_df[dimension].unique() if dimension in prod_df.columns else ["Group Total"]

    for i, val in enumerate(values_in_df):
        if dimension in prod_df.columns:
            sub = prod_df[prod_df[dimension] == str(val)].sort_values("ActualPeriod")
        else:
            sub = prod_df.sort_values("ActualPeriod")
        color = COLORS[i % len(COLORS)]
        prod_pct = sub["Value"] * 100
        fig.add_trace(go.Scatter(
            x=sub["ActualPeriod"], y=prod_pct,
            mode="lines+markers+text",
            name=str(val),
            text=[f"{v:.1f}%" if pd.notna(v) else "" for v in prod_pct],
            textposition="top center",
            line=dict(color=color, width=2),
            marker=dict(size=6),
        ))

    
    fig.add_hline(y=0, line_dash="dash", line_color="black", line_width=1.5)

    fig.update_layout(
        title=f"Productivity Over Time by {dimension}",
        xaxis=dict(title="Period", tickformat="%b %Y", dtick="M1", tickangle=45),
        yaxis=dict(title="Productivity (%)", dtick=y_dtick, tickmode="linear"),
        #yaxis_title="Productivity (%)",
        height=420,
        hovermode="x unified",
    )
    return fig


def make_velocity_chart(prod_df: pd.DataFrame, dimension: str, metric_col: str) -> go.Figure:
    """
    p4 equivalent: EffortData vs BaseEfforEquiv over time.
    R: pivot_longer(c(EffortData, BaseEfforEquiv)) → geom_line by Metric
    """
    fig = go.Figure()
    values_in_df = prod_df[dimension].unique() if dimension in prod_df.columns else ["Group Total"]

    styles = [
        ("EffortData",     "Real",     "solid"),
        ("BaseEfforEquiv", "Expected", "dash"),
    ]

    for i, val in enumerate(values_in_df):
        if dimension in prod_df.columns:
            sub = prod_df[prod_df[dimension] == str(val)].sort_values("ActualPeriod")
        else:
            sub = prod_df.sort_values("ActualPeriod")
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
        xaxis=dict(title="Period", tickformat="%b %Y", dtick="M1", tickangle=45),
        yaxis_title=metric_col,
        height=420,
        hovermode="x unified",
    )
    return fig


# ============================================================
# STREAMLIT UI
# ============================================================

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df_raw = load_excel_first_or_rawdata(uploaded_file)
    df_raw = clean_columns(df_raw)
    column_map = detect_columns(df_raw)

    # ── Sidebar ──────────────────────────────────────────────
    st.sidebar.header("Controls")

    model = st.sidebar.selectbox(
        "Productivity Model",
        ["More is Best (Points)", "Less is Best (Effort)"]
    )

    if "Less is Best" in model:
        prepared, result = prepare_effort_model(df_raw, column_map)
    else:
        prepared, result = prepare_points_model(df_raw, column_map)

    if prepared is None:
        st.error(result)
        st.stop()

    df     = prepared
    config = result

    if not config["dimensions"]:
        st.error("No valid dimensions available for this file.")
        st.stop()

    dimension = st.sidebar.selectbox("Analyze by", config["dimensions"])

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
        help="Individual = R Recursive mode. Global = R Single mode.",
    )

    show_charts = st.sidebar.multiselect(
        "Charts to show",
        ["Productivity", "Velocity (Real vs Expected)", "Count over Time", "Mean over Time"],
        default=["Productivity", "Velocity (Real vs Expected)"],
    )

    if not selected_values:
        st.warning("Select at least one value.")
        st.stop()

    # ── Aggregation ────────────────────────────────
    df_filtered = df[df[dimension].isin(selected_values)].copy()
    db_agg = aggregate_monthly(df_filtered, dimension, config["metric_col"])
    db_agg[dimension] = db_agg[dimension].astype(str)

    # ── Productivity Calculation  ───────────────────
    if "Individual" in analysis_mode:
        prod_df = calc_individual_productivity(
            db_agg, dimension, config["more_is_best"], selected_values
        )
    else:
        prod_df = calc_global_productivity(
            db_agg, dimension, config["more_is_best"], selected_values
        )

    # ── Charts ────────────────────────────────────────────────
    if prod_df.empty:
        st.warning(
            "⚠️ Not enough historical data to calculate productivity. "
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
                "Real = sum of metric in current window  |  "
                "Expected = what baseline EpU predicts for current volume"
            )
            st.plotly_chart(
                make_velocity_chart(prod_df, dimension, config["metric_col"]),
                use_container_width=True
            )

    if "Count over Time" in show_charts:
        st.subheader("🔢 Ticket Count Over Time")
        st.plotly_chart(make_count_chart(db_agg, dimension, selected_values), use_container_width=True)

    if "Mean over Time" in show_charts:
        st.subheader(f"📊 Mean {config['metric_col']} Over Time")
        st.plotly_chart(
            make_mean_chart(db_agg, dimension, config["metric_col"], selected_values),
            use_container_width=True
        )

    # ── Data tables ───────────────────────────────────────────
    with st.expander("📋 Aggregated Monthly Data (db_agg)", expanded=False):
        st.caption(
            "R equivalent: group_by(Period, var) → summarise(n=n(), Sum=sum(Target), Mean=mean(Target))"
        )
        st.dataframe(db_agg.sort_values(["Period", dimension]), use_container_width=True)

    if not prod_df.empty:
        with st.expander("📋 Productivity Results (fx.PRODUCTIVITY.v3 output)", expanded=False):
            display = prod_df.copy()
            display["Productivity %"] = (display["Value"] * 100).map(
                lambda x: f"{x:.4f}%" if pd.notna(x) else ""
            )
            display["EpU_Real"]     = display["EffortData"]     / display["EffortData"].replace(0, float("nan"))
            st.caption(
                "R columns: ActualPeriod → EffortData (real sum) → BaseEfforEquiv (EpU_BL × UnitsData) → Value"
            )
            st.dataframe(display.sort_values("ActualPeriod"), use_container_width=True)

    with st.expander("🔬 Calculation Trace (window detail per period)", expanded=False):
        st.caption(
            f"Windows: Current=[0:{CURRENT_SIZE}] · Gap=[{CURRENT_SIZE}:{CURRENT_SIZE+GAP_SIZE}] (skipped) · "
            f"Baseline=[{CURRENT_SIZE+GAP_SIZE}:{CURRENT_SIZE+GAP_SIZE+BASELINE_SIZE}] (or fallback last {BASELINE_SIZE}). "
            f"R guard: min {MIN_PERIODS} periods required per value."
        )
        trace_rows = []
        for val in selected_values:
            sub = db_agg[db_agg[dimension] == str(val)].sort_values("Period", ascending=False).reset_index(drop=True)
            n_periods = len(sub)
            for current_period in sorted(sub["Period"].unique()):
                hist = sub[sub["Period"] <= current_period].sort_values("Period", ascending=False).reset_index(drop=True)
                n = len(hist)
                max_f = hist["Period"].max()
                if n < CURRENT_SIZE or current_period > max_f:
                    continue
                has_full = n >= (CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE)
                cs, ce = 0, CURRENT_SIZE
                bs = (CURRENT_SIZE + GAP_SIZE) if has_full else max(0, n - BASELINE_SIZE)
                be = (CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE) if has_full else n
                cw = hist.iloc[cs:ce]
                bw = hist.iloc[bs:be]
                ed = cw["Sum"].sum(); ud = cw["n"].sum()
                eb = bw["Sum"].sum(); ub = bw["n"].sum()
                epu = eb / ub if ub > 0 else None
                bequiv = epu * ud if epu is not None else None
                guard_ok = n_periods >= MIN_PERIODS
                trace_rows.append({
                    dimension:          val,
                    "ActualPeriod":     current_period,
                    "n_hist":           n,
                    "has_baseline_full":has_full,
                    "guard_ok (≥5)":    guard_ok,
                    "cur_window":       f"[{cs}:{ce}] → {hist['Period'].iloc[cs:ce].dt.strftime('%b%y').tolist()}",
                    "base_window":      f"[{bs}:{be}] → {hist['Period'].iloc[bs:be].dt.strftime('%b%y').tolist()}",
                    "EffortData":       ed,
                    "UnitsData":        ud,
                    "EffortBaseline":   eb,
                    "UnitsBaseline":    ub,
                    "EpU_BL":           round(epu, 4) if epu else None,
                    "BaseEfforEquiv":   round(bequiv, 4) if bequiv else None,
                })
        if trace_rows:
            st.dataframe(pd.DataFrame(trace_rows), use_container_width=True)

else:
    st.info("⬆️ Upload an Excel file to begin.")
