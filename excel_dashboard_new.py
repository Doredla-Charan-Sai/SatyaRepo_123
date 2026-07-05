import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, date

st.set_page_config(page_title="Excel Analytics Dashboard", layout="wide")

# --- Configuration / Required Columns ---
REQUIRED_COLUMNS = [
    "Client",
    "Source System",
    "NSCC Dealer Number",
    "Branch Code",
    "Rep Code",
    "Rep Code Created Date",
    "Firm Name",
    "Firm CRD",
    "Office Address Line 1",
    "Office City",
    "Office Region",
    "Person ID",
    "Person First Name",
    "Person Last Name",
    "YTD Trade Amount",
    "Asset Balance",
]

COLOR_BANDS = [
    (0, 90, "#d4edda", "Recent"),
    (91, 180, "#fff3cd", "Moderate"),
    (181, 270, "#ffe5d0", "Aging"),
    (271, 360, "#f8d7da", "Old"),
    (361, 10000000, "#e2d9f3", "Very Old"),
]

# --- Helpers ---

def normalize(s):
    if s is None:
        return ""
    return str(s).strip().lower()


def normalize_list(lst):
    return [normalize(x) for x in lst]


def read_excel_file(uploaded):
    try:
        # read entire sheet into DataFrame
        df = pd.read_excel(uploaded, engine="openpyxl")
    except Exception:
        df = pd.read_excel(uploaded)
    return df


def drop_leading_empty_column(df):
    # Data always starts from column B per spec — drop first column if empty or unnamed
    if df.shape[1] == 0:
        return df
    first_col = df.columns[0]
    if normalize(first_col) in ("", "unnamed: 0") or df[first_col].isna().all():
        return df.drop(columns=[first_col])
    return df


def validate_columns(df_cols, required_cols):
    df_norm = normalize_list(df_cols)
    req_norm = normalize_list(required_cols)
    matched = []
    missing = []
    for req, req_n in zip(required_cols, req_norm):
        if req_n in df_norm:
            matched.append(req)
        else:
            missing.append(req)
    # also detect extra
    extra = [c for c in df_cols if normalize(c) not in req_norm]
    return matched, missing, extra


def format_currency(val):
    try:
        if pd.isna(val):
            return ""
        return f"${val:,.2f}"
    except Exception:
        return str(val)


def fmt_date(v):
    try:
        if pd.isna(v):
            return ""
        if isinstance(v, (pd.Timestamp, datetime, date)):
            return v.strftime("%b %d, %Y")
        parsed = pd.to_datetime(v, errors="coerce")
        if pd.isna(parsed):
            return str(v)
        return parsed.strftime("%b %d, %Y")
    except Exception:
        return str(v)


def compute_age_days(dt_series):
    today = pd.to_datetime(date.today())
    dt_parsed = pd.to_datetime(dt_series, errors="coerce")
    age = (today - dt_parsed).dt.days
    return age


def age_to_color_label(age):
    if pd.isna(age):
        return ("", "Invalid Date")
    for lo, hi, color, label in COLOR_BANDS:
        if lo <= age <= hi:
            return (color, label)
    return ("", "")


def style_rows(df, age_series):
    # returns a Styler with row background colors applied
    def _row_color(i):
        age = age_series.iat[i]
        color, _ = age_to_color_label(age)
        if color:
            return [f'background-color: {color}'] * len(df.columns)
        # invalid or no color
        return [""] * len(df.columns)

    return df.style.apply(lambda _: ["background-color: transparent"] * len(df.columns), axis=1).apply(lambda row: _row_color(row.name), axis=1)


# --- UI ---

st.title("Excel Analytics Dashboard")
st.write("Upload an Excel file (.xlsx/.xls) and explore interactive insights.")

with st.sidebar:
    st.header("Upload Excel")
    uploaded_file = st.file_uploader("Drag & drop or browse file", type=["xlsx", "xls"], accept_multiple_files=False)
    if uploaded_file is not None:
        size = len(uploaded_file.getvalue())
        if size > 10 * 1024 * 1024:
            st.warning("File is larger than 10MB. Check that you want to continue.")
            confirm_large = st.checkbox("I confirm I want to upload this large file")
        else:
            confirm_large = True
    else:
        confirm_large = False

# use session state to hold parsed data
if "df_raw" not in st.session_state:
    st.session_state.df_raw = None
    st.session_state.validated = False
    st.session_state.validation = {}

if uploaded_file is not None and confirm_large:
    try:
        df = read_excel_file(uploaded_file)
    except Exception as e:
        st.error("Failed to read Excel file: " + str(e))
        st.stop()

    if df.empty:
        st.error("No data found in the uploaded file")
        st.stop()

    df = drop_leading_empty_column(df)

    cols = list(df.columns)
    matched, missing, extra = validate_columns(cols, REQUIRED_COLUMNS)

    st.session_state.validation = {"matched": matched, "missing": missing, "extra": extra}

    if missing:
        st.error("Column validation failed. Required columns are missing or misnamed.")
        with st.expander("Validation details"):
            st.write("Expected columns:", REQUIRED_COLUMNS)
            st.write("Matched columns:", matched)
            st.write("Missing or misnamed columns:", missing)
            if extra:
                st.write("Extra columns detected:", extra)
        st.stop()

    # validation passed
    st.success("File validated successfully — loading dashboard...")
    # coerce date column
    date_col = [c for c in df.columns if normalize(c) == normalize("Rep Code Created Date")][0]
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    # store
    st.session_state.df_raw = df
    st.session_state.validated = True

# If validated load dashboard
if st.session_state.validated and st.session_state.df_raw is not None:
    df = st.session_state.df_raw.copy()

    # Format for display and operations
    # identify currency and date columns from REQUIRED_COLUMNS
    date_col = [c for c in df.columns if normalize(c) == normalize("Rep Code Created Date")][0]
    currency_cols = [c for c in df.columns if normalize(c) in normalize_list(["YTD Trade Amount", "Asset Balance"])]

    st.subheader("Dashboard")

    # KPI cards
    col1, col2, col3, col4 = st.columns(4)
    filtered_df = df.copy()

    # Month filter
    months = df[date_col].dt.to_period("M").dropna().unique()
    months = sorted([m.to_timestamp() for m in months])
    month_labels = [d.strftime("%b %Y") for d in months]

    st.sidebar.header("Filters")
    select_all = st.sidebar.checkbox("Select All Months", value=True)
    if select_all:
        selected_months = st.sidebar.multiselect("Filter by Month (Rep Code Created Date)", month_labels, default=month_labels)
    else:
        selected_months = st.sidebar.multiselect("Filter by Month (Rep Code Created Date)", month_labels, default=[])

    # apply month filter
    if selected_months:
        # map labels back to timestamps
        sel_ts = [datetime.strptime(m, "%b %Y") for m in selected_months]
        sel_periods = [pd.Period(year=t.year, month=t.month, freq="M") for t in sel_ts]
        mask = df[date_col].dt.to_period("M").isin(sel_periods)
        filtered_df = df[mask]
    else:
        filtered_df = df.iloc[0:0]

    # KPIs
    total_records = len(filtered_df)
    total_ytd = filtered_df[currency_cols[0]].sum() if currency_cols else 0
    total_asset = filtered_df[currency_cols[1]].sum() if len(currency_cols) > 1 else 0
    min_date = filtered_df[date_col].min()
    max_date = filtered_df[date_col].max()

    col1.metric("Total Records", f"{total_records}")
    col2.metric("Total YTD Trade Amount", f"${total_ytd:,.2f}")
    col3.metric("Total Asset Balance", f"${total_asset:,.2f}")
    date_range_text = "—"
    if pd.notna(min_date) and pd.notna(max_date):
        date_range_text = f"{min_date.strftime('%b %d, %Y')} → {max_date.strftime('%b %d, %Y')}"
    col4.metric("Date Range", date_range_text)

    st.markdown("---")

    # Table view selection
    view = st.radio("Choose Table View", ["Table View 1 — Filtered Table", "Table View 2 — Sorted & Color Coded"], horizontal=True)

    # common table controls
    page_size = st.selectbox("Rows per page", [10, 25, 50, 100], index=1)
    total_pages = max(1, (len(filtered_df) + page_size - 1) // page_size)
    page = st.number_input("Page", min_value=1, max_value=total_pages, value=1)

    start = (page - 1) * page_size
    end = start + page_size

    if view.startswith("Table View 1"):
        st.write(f"Showing {len(filtered_df)} record(s)")
        if filtered_df.empty:
            st.info("No records match the current filters.")
        else:
            display_df = filtered_df.copy()
            # formatting
            for c in currency_cols:
                if c in display_df.columns:
                    display_df[c] = display_df[c].apply(format_currency)
            if date_col in display_df.columns:
                display_df[date_col] = display_df[date_col].apply(fmt_date)

            st.dataframe(display_df.iloc[start:end], use_container_width=True)

    else:
        # Table View 2 — sorting + color coding
        st.write("Sorting and color-coded age bands based on Rep Code Created Date")
        # choose sortable columns
        sortable_candidates = []
        # numeric columns
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]) or pd.api.types.is_datetime64_any_dtype(df[c]):
                sortable_candidates.append(c)
        sort_by = st.selectbox("Sort By", options=["(No Sort)"] + sortable_candidates, index=0)
        asc = st.radio("Order", ["Descending ↓", "Ascending ↑"], index=0, horizontal=True)
        if sort_by != "(No Sort)":
            ascending = asc.endswith("↑")
            filtered_df = filtered_df.sort_values(by=sort_by, ascending=ascending)

        # compute ages
        ages = compute_age_days(filtered_df[date_col])
        # count invalid dates
        invalid_dates = ages.isna().sum()
        if invalid_dates:
            st.warning(f"{invalid_dates} row(s) have invalid dates; color coding skipped for those rows.")

        # prepare display df with formatted currency and dates for readability
        display_df = filtered_df.copy()
        for c in currency_cols:
            if c in display_df.columns:
                display_df[c] = display_df[c].apply(lambda v: float(v) if pd.notna(v) else np.nan)
        # format dates for display but keep original for age computation
        if date_col in display_df.columns:
            display_df[date_col] = display_df[date_col].apply(lambda v: v if pd.isna(v) else v)

        # style rows
        sty = style_rows(display_df.iloc[start:end].reset_index(drop=True), ages.iloc[start:end].reset_index(drop=True))
        # show legend
        legend_cols = st.columns(len(COLOR_BANDS))
        for i, (_, _, color, label) in enumerate(COLOR_BANDS):
            with legend_cols[i]:
                st.markdown(f"<div style='background:{color};padding:8px;border-radius:4px;text-align:center'>{label}</div>", unsafe_allow_html=True)

        st.write(f"Showing {len(filtered_df)} record(s)")
        st.dataframe(sty, use_container_width=True)

    st.markdown("---")

    # Export button (export filtered + sorted)
    def to_excel_bytes(df_to_export, ages_series):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_to_export.to_excel(writer, index=False, sheet_name="Export")
            workbook = writer.book
            worksheet = writer.sheets["Export"]

            # apply currency formats
            money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
            date_fmt = workbook.add_format({'num_format': 'mmm dd, yyyy'})

            # set column widths and formats
            for i, col in enumerate(df_to_export.columns):
                # width
                width = max(12, min(40, max(df_to_export[col].astype(str).map(len).max(), len(col) + 2)))
                worksheet.set_column(i, i, width)
                if normalize(col) in normalize_list(["ytd trade amount", "asset balance"]):
                    worksheet.set_column(i, i, width, money_fmt)
                if normalize(col) == normalize("Rep Code Created Date"):
                    worksheet.set_column(i, i, width, date_fmt)

            # apply row background colors based on ages
            for row_idx, age in enumerate(ages_series):
                color, _ = age_to_color_label(age)
                if color:
                    fmt = workbook.add_format({'bg_color': color})
                    worksheet.set_row(row_idx + 1, None, fmt)
        output.seek(0)
        return output.getvalue()

    export_df = filtered_df.copy()
    export_ages = compute_age_days(export_df[date_col])
    if not export_df.empty:
        excel_bytes = to_excel_bytes(export_df, export_ages)
        st.download_button("Export current view to Excel", data=excel_bytes, file_name="export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Upload a validated Excel file using the sidebar to begin.")

# Footer / small help
st.markdown("---")
st.caption("Notes: Header row is expected at Row 1; data should start from Column B. Date parsing uses MM/DD/YYYY where possible.")
