# ============================================================
# PRODUCTIVITY ANALYSIS - Streamlit App
# ============================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(layout="wide")

CURRENT_SIZE  = 3
GAP_SIZE      = 3
BASELINE_SIZE = 3

st.title("📊 Productivity Analysis (Unified Model)")



# ------------------------------------------------------------
# FILE READ
# ------------------------------------------------------------

def load_excel_first_or_rawdata(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = "RawData" if "RawData" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


# ------------------------------------------------------------
# COLUMN SETUP NAME
# ------------------------------------------------------------
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


# ------------------------------------------------------------
# "LESS IS BEST" MODEL 
# ------------------------------------------------------------
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
        "metric_col":       "Effort",
        "real_label":       "Real Effort",
        "expected_label":   "Expected Effort",
        "more_is_best":     False,
        "dimensions":       [c for c in ["Assigned To", "Group", "WBS", "Category", "Service Type"] if c in df.columns],
    }
    return df, config


# ------------------------------------------------------------
# "MORE IS BEST" MODEL 
# ------------------------------------------------------------
def prepare_points_model(df: pd.DataFrame, column_map: dict):
    required = ["Points", "Developer", "Status", "Period"]
    missing = [k for k in required if column_map[k] is None]
    if missing:
        return None, f"Missing required columns: {missing}"

    df = df.rename(columns={v: k for k, v in column_map.items() if v is not None}).copy()
    df["Points"] = pd.to_numeric(df["Points"], errors="coerce")
    df["Period"] = pd.to_datetime(df["Period"], errors="coerce").dt.to_period("M").dt.to_timestamp()

    # - Status filter only
    df = df[df["Status"].isin(["Ready to Deploy", "Closed"])].copy()

    # Synthetic group to mimic R's db$Grupo <- "Grupo"
    df["Group"] = "Group"

    # Clean Developer values
    df["Developer"] = df["Developer"].astype("string").str.strip()
    df = df[df["Developer"].notna() & (df["Developer"] != "")].copy()

    # Split only if "/" is present, trim spaces afterward
    df = df.assign(Developer=df["Developer"].str.split("/")).explode("Developer")
    df["Developer"] = df["Developer"].astype("string").str.strip()

        # Remove blanks and accidental literal nan text
    df = df[
        df["Developer"].notna()
        & (df["Developer"] != "")
        & (df["Developer"].str.lower() != "nan")
    ].copy()

    df = df[df["Period"].notna() & df["Points"].notna()].copy()

    config = {
        "metric_col":     "Points",
        "real_label":     "Real Points",
        "expected_label": "Expected Points",
        "more_is_best":   True,
        "dimensions":     [c for c in ["Developer", "Group", "QA Tester", "Priority", "Issue Type"] if c in df.columns],
    }
    return df, config


# ------------------------------------------------------------
# MONTHLY AGG
# ------------------------------------------------------------
def aggregate_monthly(df: pd.DataFrame, period_col: str, dimension: str, metric_col: str) -> pd.DataFrame:
    return (
        df.groupby([period_col, dimension], dropna=False)
        .agg(
            Value=  (metric_col, "sum"),   # value (float): Total Effort/Points for the month
            Units=  (metric_col, "size"),  # units (int): Monthly ticket count
        )
        .reset_index()
        .rename(columns={period_col: "Period"})
        .sort_values(["Period", dimension])
        .reset_index(drop=True)
    )


# ------------------------------------------------------------
# CÁLCULO PRINCIPAL DE PRODUCTIVIDAD
# Replica exactamente la lógica de fx.PRODUCTIVITY.v3 en R.
#
# Para cada periodo evaluado (current_period):
#   1. Filtra datos históricos: Period <= current_period
#   2. Ordena descendente (más reciente primero)
#   3. Define ventanas por posición (índices):
#        Current  [0 : CURRENT_SIZE]
#        Gap      [CURRENT_SIZE : CURRENT_SIZE + GAP_SIZE]   (se ignora)
#        Baseline [CURRENT_SIZE+GAP_SIZE : CURRENT_SIZE+GAP_SIZE+BASELINE_SIZE]
#        Fallback si no hay suficiente historial:
#          Baseline = últimos BASELINE_SIZE periodos disponibles
#   4. NUEVO CHEQUEO (igual que R): si current_period > max fecha real
#      del grupo, NO se calcula (evita proyectar hacia el futuro).
#   5. Calcula EpU_BL = EffortBaseline / UnitsBaseline
#      BaseEffortEquiv = EpU_BL * UnitsData  (lo que se esperaría hoy
#      si el equipo rindiera igual que en baseline)
#   6. Productivity = (EffortData - BaseEffortEquiv) / BaseEffortEquiv * signo
# ------------------------------------------------------------
def calc_r_compatible(
    group: pd.DataFrame,
    real_label: str,
    expected_label: str,
    more_is_best: bool
) -> pd.DataFrame:

    group = group.sort_values("Period").reset_index(drop=True)
    fechas = sorted(group["Period"].unique())

    # 1: higher is better, -1: lower is better
    signo = 1 if more_is_best else -1

    # Max available date in group data
    max_fecha_real = group["Period"].max()

    rows = []

    for current_period in fechas:

        
        # Skip calculation if period exceeds max available date.
        if current_period > max_fecha_real:
            continue

        # Available history up to this period (descending)
        subset = (
            group[group["Period"] <= current_period]
            .sort_values("Period", ascending=False)
            .reset_index(drop=True)
        )
        n = len(subset)

       
        if n < CURRENT_SIZE:
            continue

       
        has_baseline_full = n >= (CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE)

        # Current window: positions 0..CURRENT_SIZE-1 (most recent)
        current_window = subset.iloc[0 : CURRENT_SIZE]

        # Base line window:

        if has_baseline_full:
            base_start = CURRENT_SIZE + GAP_SIZE
            base_end   = CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE
        else:
            base_start = max(0, n - BASELINE_SIZE)   
            base_end   = n

        baseline_window = subset.iloc[base_start : base_end]

        # SUM
        effort_data     = current_window["Value"].sum() 
        units_data      = current_window["Units"].sum()  

        effort_baseline = baseline_window["Value"].sum()  
        units_baseline  = baseline_window["Units"].sum()  

        # Avoid division by zero.
        if units_baseline == 0 or units_data == 0:
            continue

        # EpU_BL = Effort/points per ticket in baseline period
        epu_bl = effort_baseline / units_baseline

        
        # BaseEffortEquiv: Expected effort for current volume based on baseline performance
        base_effort_equiv = epu_bl * units_data

        if base_effort_equiv == 0:
            continue

       
        # Productivity calc
        productivity = ((effort_data - base_effort_equiv) / base_effort_equiv) * signo

        rows.append({
            "Period":              current_period,
            real_label:            effort_data,        # Real Effort / Real Points
            expected_label:        base_effort_equiv,  # Expected Effort / Expected Points
            "Baseline (Moving)":   base_effort_equiv,  # Base Line
            "Productivity":        productivity,
            "Tickets (Window)":    units_data,         # Current Tickets
        })

    return pd.DataFrame(rows)


# ============================================================
# STREAMLIT
# ============================================================
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = load_excel_first_or_rawdata(uploaded_file)
    df = clean_columns(df)
    column_map = detect_columns(df)

    st.sidebar.header("Controls")
    model = st.sidebar.selectbox(
        "Productivity Model",
        ["Less is Best (Effort)", "More is Best (Points)"]
    )

    if "Less is Best" in model:
        prepared, result = prepare_effort_model(df, column_map)
    else:
        prepared, result = prepare_points_model(df, column_map)

    if prepared is None:
        st.error(result)
        st.info("Please change to the corresponding file for this analysis.") 
        st.stop()

    df     = prepared
    config = result

    if not config["dimensions"]:
        st.error("No valid dimensions are available for this file.")
        st.stop()

    dimension = st.sidebar.selectbox("Analyze by", config["dimensions"])

    values = sorted(df[dimension].dropna().astype(str).unique().tolist())
    selected_values = st.sidebar.multiselect(
        "Select values",
        values,
        default=values[:1] if values else []
    )

   
    metrics = st.sidebar.multiselect(
        "Metrics",
        [
            "Productivity",
            config["real_label"],
            config["expected_label"],
            "Baseline (Moving)",     # NEW
            "Tickets (Window)",      # Tickets in current window (3 months)
            "Tickets (Monthly)",     # Tickets for the specific month
        ],
        default=["Productivity"],
    )

    if not selected_values or not metrics:
        st.warning("Select at least one value and one metric.")
        st.stop()

    df[dimension] = df[dimension].astype(str)
    df = df[df[dimension].isin(selected_values)].copy()

    
    agg = aggregate_monthly(df, "Period", dimension, config["metric_col"])

    final_df = pd.DataFrame()

    for val in selected_values:
        sub = agg[agg[dimension] == str(val)].copy()

        res = calc_r_compatible(
            sub,
            real_label=     config["real_label"],
            expected_label= config["expected_label"],
            more_is_best=   config["more_is_best"],
        )

        if not res.empty:
            res[dimension] = str(val)


            monthly = sub[["Period", "Units"]].rename(columns={"Units": "Tickets (Monthly)"})
            res = res.merge(monthly, on="Period", how="left")

            # Fill date range for continuous visualization
            full_range = pd.date_range(res["Period"].min(), res["Period"].max(), freq="MS")
            res = (
                res.set_index("Period")
                .reindex(full_range)
                .rename_axis("Period")
                .reset_index()
            )
            res[dimension] = str(val)
            final_df = pd.concat([final_df, res], ignore_index=True)

    if final_df.empty:
        st.warning("No data available.")
        st.stop()

    # Convert productivity to percentage for visualization
    final_df["Productivity %"] = final_df["Productivity"] * 100

    # TREND CHART
    st.subheader("📈 Trend")
    fig = go.Figure()
    timeline = pd.date_range(final_df["Period"].min(), final_df["Period"].max(), freq="MS")

    for val in selected_values:
        sub = final_df[final_df[dimension] == str(val)].copy()
        sub = (
            sub.set_index("Period")
            .reindex(timeline)
            .rename_axis("Period")
            .reset_index()
        )

        for metric in metrics:
            
            col = "Productivity %" if metric == "Productivity" else metric
            if col not in sub.columns:
                continue

            text_vals = [
                f"{v:.2f}%" if metric == "Productivity" and pd.notna(v)
                else f"{v:.2f}" if pd.notna(v)
                else ""
                for v in sub[col]
            ]

            # Baseline with distinct style (dotted)
            line_style = dict(dash="dash") if metric == "Baseline (Moving)" else {}

            fig.add_trace(
                go.Scatter(
                    x=    sub["Period"],
                    y=    sub[col],
                    mode= "lines+markers+text",
                    name= f"{val} - {metric}",
                    text= text_vals,
                    textposition="top center",
                    line= line_style,
                )
            )

    fig.update_layout(
        height=550,
        xaxis=dict(
            title="Month",
            tickformat="%b %Y",
            dtick="M1",
            tickangle=45,
        ),
        yaxis_title="Value",
    )
    st.plotly_chart(fig, use_container_width=True)

    
    st.subheader("📋 Monthly Values")
    display_df = final_df.copy()
    display_df["Productivity %"] = display_df["Productivity %"].map(
        lambda x: f"{x:.4f}%" if pd.notna(x) else ""
    )
    st.dataframe(display_df)

else:
    st.info("Upload an Excel file to begin.")
