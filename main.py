#This script uses Streamlit.
#To install it:
#     pip install streamlit pandas openpyxl
#     python -m pip install plotly
#To run it:
#     python -m streamlit run app5.py


import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(layout="wide")

CURRENT_SIZE = 3
GAP_SIZE = 3
BASELINE_SIZE = 3

st.title("📊 Productivity Analysis (Unified Model)")


def load_excel_first_or_rawdata(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = "RawData" if "RawData" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)
    return df


def detect_columns(df: pd.DataFrame) -> dict:
    column_map = {
        "Assigned To": None,
        "Group": None,
        "WBS": None,
        "Category": None,
        "Service Type": None,
        "EndDate": None,
        "Effort": None,
        "Points": None,
        "Developer": None,
        "Status": None,
        "Period": None,
        "QA Tester": None,
        "Issue Type": None,
        "Priority": None
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


def prepare_effort_model(df: pd.DataFrame, column_map: dict):
    required = ["Assigned To", "Group", "WBS", "EndDate", "Effort"]
    missing = [k for k in required if column_map[k] is None]
    if missing:
        return None, f"Missing required columns: {missing}"

    df = df.rename(columns={v: k for k, v in column_map.items() if v is not None}).copy()

    df["EndDate"] = pd.to_datetime(df["EndDate"], errors="coerce")
    df["Period"] = df["EndDate"].dt.to_period("M").dt.to_timestamp()
    df["Effort"] = pd.to_numeric(df["Effort"], errors="coerce")

    df = df[df["Period"].notna() & df["Effort"].notna()].copy()

    config = {
        "metric_col": "Effort",
        "real_label": "Real Effort",
        "expected_label": "Expected Effort",
        "more_is_best": False,
        "dimensions": [c for c in ["Assigned To", "Group", "WBS", "Category", "Service Type"] if c in df.columns],
    }
    return df, config


def prepare_points_model(df: pd.DataFrame, column_map: dict):
    required = ["Points", "Developer", "Status", "Period"]
    missing = [k for k in required if column_map[k] is None]
    if missing:
        return None, f"Missing required columns: {missing}"

    df = df.rename(columns={v: k for k, v in column_map.items() if v is not None}).copy()

    df["Points"] = pd.to_numeric(df["Points"], errors="coerce")
    df["Period"] = pd.to_datetime(df["Period"], errors="coerce").dt.to_period("M").dt.to_timestamp()

    # Match R behavior exactly:
    # - Status filter only
    # - Keep zero-point rows
    df = df[df["Status"].isin(["Ready to Deploy", "Closed"])].copy()

    # Synthetic group to mimic R's db$Grupo <- "Grupo"
    df["Group"] = "Group"

    # Clean Developer values
    df["Developer"] = df["Developer"].astype("string")
    df = df[df["Developer"].notna()].copy()

    # Split only if "/" is present, trim spaces afterward
    df = df.assign(Developer=df["Developer"].str.split("/")).explode("Developer")
    df["Developer"] = df["Developer"].astype("string").str.strip()

    # Remove blanks and accidental literal nan text
    df = df[
        df["Developer"].notna()
        & (df["Developer"] != "")
        & (df["Developer"].str.lower() != "nan")
    ].copy()

    # Keep valid rows
    df = df[df["Period"].notna() & df["Points"].notna()].copy()

    config = {
        "metric_col": "Points",
        "real_label": "Real Points",
        "expected_label": "Expected Points",
        "more_is_best": True,
        "dimensions": [c for c in ["Developer", "Group", "QA Tester", "Priority", "Issue Type"] if c in df.columns],
    }
    return df, config


def aggregate_monthly(df: pd.DataFrame, period_col: str, dimension: str, metric_col: str) -> pd.DataFrame:
    return (
        df.groupby([period_col, dimension], dropna=False)
        .agg(
            Value=(metric_col, "sum"),
            Units=(metric_col, "size"),
        )
        .reset_index()
        .rename(columns={period_col: "Period"})
        .sort_values(["Period", dimension])
        .reset_index(drop=True)
    )


def calc_r_compatible(group: pd.DataFrame, real_label: str, expected_label: str, more_is_best: bool) -> pd.DataFrame:
    """
    Exact mirror of the R logic for one selected service/dimension value.
    Input must already be monthly-aggregated with columns: Period, Value, Units
    """
    group = group.sort_values("Period").reset_index(drop=True)
    fechas = sorted(group["Period"].unique())
    rows = []

    for current_period in fechas:
        subset = (
            group[group["Period"] <= current_period]
            .sort_values("Period", ascending=False)
            .reset_index(drop=True)
        )

        n = len(subset)

        # Same as R: need at least current_size rows
        if n < CURRENT_SIZE:
            continue

        has_baseline_full = n >= (CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE)

        # current = 1:3 in R -> 0:3 in pandas
        current = subset.iloc[0:CURRENT_SIZE]

        # baseline:
        # full  = 7:9 in R -> 6:9 in pandas
        # fallback = last 3 rows
        if has_baseline_full:
            baseline = subset.iloc[
                CURRENT_SIZE + GAP_SIZE : CURRENT_SIZE + GAP_SIZE + BASELINE_SIZE
            ]
        else:
            baseline = subset.iloc[max(0, n - BASELINE_SIZE) : n]

        value_data = current["Value"].sum()
        units_data = current["Units"].sum()

        value_base = baseline["Value"].sum()
        units_base = baseline["Units"].sum()

        if units_base == 0 or units_data == 0:
            continue

        expected = (value_base / units_base) * units_data
        if expected == 0:
            continue

        if more_is_best:
            productivity = (value_data - expected) / expected
        else:
            productivity = (expected - value_data) / expected

        rows.append(
            {
                "Period": current_period,
                real_label: value_data,
                expected_label: expected,
                "Productivity": productivity,
                "Tickets (Window)": units_data,
            }
        )

    return pd.DataFrame(rows)


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
        st.stop()

    df = prepared
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
            "Tickets (Window)",
            "Tickets (Monthly)",
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
            real_label=config["real_label"],
            expected_label=config["expected_label"],
            more_is_best=config["more_is_best"],
        )

        if not res.empty:
            res[dimension] = str(val)

            monthly = sub[["Period", "Units"]].rename(columns={"Units": "Tickets (Monthly)"})
            res = res.merge(monthly, on="Period", how="left")

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

    final_df["Productivity %"] = final_df["Productivity"] * 100

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

            fig.add_trace(
                go.Scatter(
                    x=sub["Period"],
                    y=sub[col],
                    mode="lines+markers+text",
                    name=f"{val} - {metric}",
                    text=text_vals,
                    textposition="top center",
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