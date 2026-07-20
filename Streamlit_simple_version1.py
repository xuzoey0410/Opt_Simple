from io import BytesIO
from functools import lru_cache
import importlib
import math
from pathlib import Path
import re
import sys

import pandas as pd
import plotly.express as px
import streamlit as st

APP_DIR = Path(__file__).parent
SAMPLE_INPUT = APP_DIR / "input_sample.xlsx"
LOCAL_INPUT = APP_DIR / "All_Input1.xlsx"
PROJECT_DIR = Path.home() / "Desktop" / "Demand Project"
PROJECT_SAMPLE_INPUT = PROJECT_DIR / "input_sample.xlsx"
PROJECT_LOCAL_INPUT = PROJECT_DIR / "All_Input1.xlsx"
USG_WSPW_SHEET = "POM27_030526_USG"

for candidate_dir in [APP_DIR, PROJECT_DIR]:
    candidate_path = str(candidate_dir)
    sys.path = [path for path in sys.path if path != candidate_path]

for candidate_dir in [APP_DIR, PROJECT_DIR]:
    if candidate_dir.exists():
        sys.path.insert(0, str(candidate_dir))

if "Output_Simple" in sys.modules:
    del sys.modules["Output_Simple"]
planner = importlib.import_module("Output_Simple")

st.set_page_config(page_title="Supply Chain Planner", layout="wide")
st.title("Supply Chain Planner")


def sheet_name(excel_file, target):
    lookup = {s.strip().lower(): s for s in excel_file.sheet_names}
    return lookup.get(target.strip().lower())


def input_source(file_bytes):
    if file_bytes:
        return BytesIO(file_bytes)
    if SAMPLE_INPUT.exists():
        return SAMPLE_INPUT
    if LOCAL_INPUT.exists():
        return LOCAL_INPUT
    if PROJECT_SAMPLE_INPUT.exists():
        return PROJECT_SAMPLE_INPUT
    if PROJECT_LOCAL_INPUT.exists():
        return PROJECT_LOCAL_INPUT
    return None


@st.cache_data(show_spinner=False)
def load_tables(file_bytes):
    source = input_source(file_bytes)
    if source is None:
        return None

    xl = pd.ExcelFile(source)

    def read(target, required=True):
        name = sheet_name(xl, target)
        if name is None:
            if required:
                raise ValueError(f"Missing sheet: {target}")
            return pd.DataFrame()
        return pd.read_excel(xl, sheet_name=name)

    return {
        "flow": read(planner.FLOW_SHEET),
        "product": read(planner.PRODUCT_SHEET),
        "demand": read(planner.DEMAND_SHEET),
        "inventory": read(planner.INVENTORY_SHEET),
        "target": read(planner.TARGET_SHEET, required=False),
    }


def empty_wspw_wafer_start_table():
    return pd.DataFrame(columns=["Basic Type", "RFP"])


def parse_wspw_month_header(value):
    if pd.isna(value):
        return None

    text = str(value).strip().replace("Sept", "Sep")
    match = re.fullmatch(r"([A-Za-z]{3,})'?([0-9]{2,4})", text)
    if match:
        month_lookup = {
            "jan": 1,
            "feb": 2,
            "mar": 3,
            "apr": 4,
            "may": 5,
            "jun": 6,
            "jul": 7,
            "aug": 8,
            "sep": 9,
            "oct": 10,
            "nov": 11,
            "dec": 12,
        }
        month = month_lookup.get(match.group(1)[:3].lower())
        year = int(match.group(2))
        if month is not None:
            if year < 100:
                year += 2000
            return f"{year:04d}-{month:02d}"

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m")


def build_week_label(month_label, week_value):
    week_number = pd.to_numeric(week_value, errors="coerce")
    if pd.isna(week_number):
        return None

    parsed_month = pd.to_datetime(month_label, errors="coerce")
    if pd.isna(parsed_month):
        return None

    label_year, month_start_week = month_start_year_week(parsed_month.strftime("%Y-%m"))
    week_number = int(week_number)
    if week_number < month_start_week:
        label_year += 1
    return f"{label_year}-CW{week_number:02d}"


@lru_cache(maxsize=None)
def add_calendar_weeks(year, week, delta):
    year = int(year)
    week = int(week) + int(delta)
    while week > 52:
        week -= 52
        year += 1
    while week <= 0:
        week += 52
        year -= 1
    return year, week


@lru_cache(maxsize=None)
def month_start_year_week(month_label):
    parsed_month = pd.to_datetime(month_label, errors="coerce")
    if pd.isna(parsed_month):
        raise ValueError(f"Month format is not recognized: {month_label}")

    year = int(parsed_month.year)
    week = 2
    for month_number in range(1, int(parsed_month.month)):
        year, week = add_calendar_weeks(year, week, planner.month_to_weeks(f"{parsed_month.year}-{month_number:02d}"))
    return year, week


@lru_cache(maxsize=None)
def month_week_labels(month_label):
    year, start_week = month_start_year_week(month_label)
    return [
        f"{week_year}-CW{week_number:02d}"
        for week_year, week_number in [add_calendar_weeks(year, start_week, offset) for offset in range(planner.month_to_weeks(month_label))]
    ]


@lru_cache(maxsize=None)
def shift_week_label(week_label, delta):
    match = re.fullmatch(r"(\d{4})-CW(\d{1,2})", str(week_label).strip())
    if not match:
        return str(week_label)
    year, week = add_calendar_weeks(int(match.group(1)), int(match.group(2)), delta)
    return f"{year}-CW{week:02d}"


@lru_cache(maxsize=None)
def week_sort_key(week_label):
    match = re.fullmatch(r"(\d{4})-CW(\d{1,2})", str(week_label).strip())
    if not match:
        return (9999, 99, str(week_label))
    return (int(match.group(1)), int(match.group(2)), "")


def is_internal_week_label(value):
    return re.fullmatch(r"\d{4}-CW\d{2}", str(value)) is not None


def is_display_week_label(value):
    return re.fullmatch(r"\d{2}[A-Za-z]{3} CW\d{2}", str(value)) is not None


@lru_cache(maxsize=None)
def week_display_label(week_label):
    text = str(week_label)
    if not is_internal_week_label(text):
        return text

    year, week_number, _ = week_sort_key(text)
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    candidate_years = [year, year - 1, year + 1]
    for candidate_year in candidate_years:
        for month_number, month_name in enumerate(month_names, start=1):
            month_label = f"{candidate_year:04d}-{month_number:02d}"
            labels = month_week_labels(month_label)
            if text in labels:
                return f"{str(candidate_year)[2:]}{month_name} CW{week_number:02d}"
    return text


def week_columns(output_df):
    return [col for col in output_df.columns if is_internal_week_label(col) or is_display_week_label(col)]


def display_week_table(output_df):
    display_df = output_df.copy()
    display_df = repair_week_sequence_columns(display_df)
    display_df = display_df.rename(columns={col: week_display_label(col) for col in display_df.columns if is_internal_week_label(col)})
    if "Week" in display_df.columns:
        display_df["Week"] = repair_week_sequence_values(display_df["Week"].tolist())
        display_df["Week"] = display_df["Week"].map(week_display_label)
    return display_df


def repair_week_sequence_values(values):
    repaired = []
    previous_year = None
    previous_week = None
    for value in values:
        text = str(value)
        if not is_internal_week_label(text):
            repaired.append(value)
            continue

        year, week_number, _ = week_sort_key(text)
        if previous_week == 52 and week_number == 1 and previous_year is not None and year <= previous_year:
            year = previous_year + 1
            text = f"{year}-CW01"
        repaired.append(text)
        previous_year = year
        previous_week = week_number
    return repaired


def repair_week_sequence_columns(output_df):
    if output_df is None or output_df.empty:
        return output_df

    repaired_columns = repair_week_sequence_values(output_df.columns.tolist())
    if repaired_columns == output_df.columns.tolist():
        return output_df

    repaired_df = output_df.copy()
    repaired_df.columns = repaired_columns
    return repaired_df


def table_week_labels(output_df):
    if output_df is None or output_df.empty:
        return []
    if "Week" in output_df.columns:
        return [str(label) for label in output_df["Week"].dropna().tolist() if is_internal_week_label(label)]
    return [col for col in output_df.columns if is_internal_week_label(col)]


def filter_week_range(output_df, start_week, end_week):
    if output_df is None or output_df.empty or not start_week or not end_week:
        return output_df

    if week_sort_key(start_week) > week_sort_key(end_week):
        start_week, end_week = end_week, start_week

    if "Week" in output_df.columns:
        output_df = output_df.copy()
        keep_rows = output_df["Week"].map(lambda week: is_internal_week_label(week) and week_sort_key(start_week) <= week_sort_key(week) <= week_sort_key(end_week))
        return output_df[keep_rows].copy()

    week_cols = [col for col in output_df.columns if is_internal_week_label(col)]
    kept_week_cols = [col for col in week_cols if week_sort_key(start_week) <= week_sort_key(col) <= week_sort_key(end_week)]
    id_cols = [col for col in output_df.columns if col not in week_cols]
    return output_df[id_cols + kept_week_cols].copy()


def week_range_selector(title, output_df, key_prefix):
    labels = sorted(set(table_week_labels(output_df)), key=week_sort_key)
    if not labels:
        return None, None

    options = [week_display_label(label) for label in labels]
    label_to_week = dict(zip(options, labels))
    start_key = f"{key_prefix}_start_week"
    end_key = f"{key_prefix}_end_week"

    if st.session_state.get(start_key) not in options:
        st.session_state[start_key] = options[0]
    if st.session_state.get(end_key) not in options:
        st.session_state[end_key] = options[-1]

    reset_col, start_col, end_col = st.columns([0.7, 1, 1])
    with reset_col:
        st.write("")
        st.write("")
        if st.button("Full range", key=f"{key_prefix}_full_range"):
            st.session_state[start_key] = options[0]
            st.session_state[end_key] = options[-1]
    with start_col:
        start_label = st.selectbox(f"{title} start", options, key=start_key)
    with end_col:
        end_label = st.selectbox(f"{title} end", options, key=end_key)

    return label_to_week[start_label], label_to_week[end_label]


def find_wspw_sheet(excel_file):
    exact_sheet = sheet_name(excel_file, USG_WSPW_SHEET)
    if exact_sheet is not None:
        return exact_sheet

    for candidate in excel_file.sheet_names:
        try:
            preview = pd.read_excel(excel_file, sheet_name=candidate, header=None, nrows=30)
        except Exception:
            continue
        has_wspw = preview.apply(lambda col: col.astype(str).str.contains("WSPW", case=False, na=False)).any().any()
        if has_wspw:
            return candidate
    return None


def parse_wspw_wafer_start(raw_df):
    wspw_mask = raw_df.apply(lambda col: col.astype(str).str.contains("WSPW", case=False, na=False))
    positions = list(zip(*wspw_mask.to_numpy().nonzero()))
    if not positions:
        return empty_wspw_wafer_start_table()

    wspw_row = int(positions[0][0])
    month_header_row = max(0, wspw_row - 5)
    week_header_row = max(0, wspw_row - 4)

    week_labels = {}
    current_month = None
    for col_index in range(raw_df.shape[1]):
        month_label = parse_wspw_month_header(raw_df.iat[month_header_row, col_index])
        if month_label is not None:
            current_month = month_label

        label = build_week_label(current_month, raw_df.iat[week_header_row, col_index]) if current_month else None
        if label is not None:
            week_labels[col_index] = label

    if not week_labels:
        return empty_wspw_wafer_start_table()

    rows = []
    for row_index in range(wspw_row + 1, raw_df.shape[0]):
        section_value = raw_df.iat[row_index, 0]
        if pd.notna(section_value) and str(section_value).strip():
            break

        basic_type = raw_df.iat[row_index, 1] if raw_df.shape[1] > 1 else None
        rfp = raw_df.iat[row_index, 2] if raw_df.shape[1] > 2 else None
        has_product = pd.notna(basic_type) or pd.notna(rfp)

        values = {
            label: pd.to_numeric(raw_df.iat[row_index, col_index], errors="coerce")
            for col_index, label in week_labels.items()
        }
        has_values = any(not pd.isna(value) for value in values.values())
        if not has_product and not has_values:
            continue

        row = {
            "Basic Type": "Grand Total" if not has_product else str(basic_type).strip(),
            "RFP": "Grand Total" if not has_product else str(rfp).strip(),
        }
        row.update({label: (0 if pd.isna(value) else value) for label, value in values.items()})
        rows.append(row)

    if not rows:
        return empty_wspw_wafer_start_table()

    return repair_week_sequence_columns(pd.DataFrame(rows, columns=["Basic Type", "RFP", *week_labels.values()]))


@st.cache_data(show_spinner=False)
def load_wspw_wafer_start(file_bytes):
    source = input_source(file_bytes)
    if source is None:
        return empty_wspw_wafer_start_table()

    try:
        xl = pd.ExcelFile(source)
        target_sheet = find_wspw_sheet(xl)
        if target_sheet is None:
            return empty_wspw_wafer_start_table()

        raw_df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        return parse_wspw_wafer_start(raw_df)
    except Exception:
        return empty_wspw_wafer_start_table()


def target_value(target_df, key, default):
    if target_df.empty or not {"Columns", "Value"}.issubset(target_df.columns):
        return default
    matches = target_df[target_df["Columns"].astype(str).str.strip().str.lower() == key.lower()]
    if matches.empty:
        return default
    value = pd.to_numeric(matches.iloc[0]["Value"], errors="coerce")
    return default if pd.isna(value) else value


def ensure_priority_column(product_df):
    updated_df = product_df.copy()
    default_columns = {
        "Priority": 0,
        "Total Yield": "",
        "Min Reach Level": "",
        "Target Reach Level": "",
        "Earliest Wafer Start Time": "",
    }
    existing_lower = {str(col).strip().lower() for col in updated_df.columns}
    for column, default_value in default_columns.items():
        if column.lower() not in existing_lower:
            updated_df.insert(len(updated_df.columns), column, default_value)
    return updated_df


def demand_month_columns(demand_df):
    return [col for col in demand_df.columns if planner.is_month_col(col)]


def normalize_month_column(value):
    raw_value = str(value).strip()
    if not raw_value:
        raise ValueError("Please enter a demand month, for example 2028-10.")

    parsed = pd.to_datetime(raw_value, errors="coerce")
    if pd.isna(parsed):
        raise ValueError("Month format is not recognized. Please use a format like 2028-10 or Oct-2028.")
    return parsed.strftime("%Y-%m")


def parse_earliest_wafer_start_time(value):
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    if is_internal_week_label(text):
        return text

    display_match = re.fullmatch(r"(\d{2})([A-Za-z]{3})\s+CW(\d{1,2})", text)
    if display_match:
        year = 2000 + int(display_match.group(1))
        month_lookup = {name.lower(): index for index, name in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}
        month = month_lookup.get(display_match.group(2).lower())
        week = int(display_match.group(3))
        if month is not None:
            month_label = f"{year:04d}-{month:02d}"
            for label in month_week_labels(month_label):
                if week_sort_key(label)[1] == week:
                    return label

    try:
        month_label = normalize_month_column(text)
        labels = month_week_labels(month_label)
        return labels[0] if labels else None
    except ValueError:
        pass

    raise ValueError(f"Earliest Wafer Start Time is not recognized: {text}. Use YYYY-CWNN, YYMon CWNN, or YYYY-MM.")


def demand_date_range_defaults(demand_df):
    month_cols = demand_month_columns(demand_df)
    if not month_cols:
        today = pd.Timestamp.today().replace(day=1)
        return today.date(), (today + pd.DateOffset(months=5)).date()

    parsed_months = pd.to_datetime([normalize_month_column(col) for col in month_cols], errors="coerce")
    parsed_months = [month for month in parsed_months if not pd.isna(month)]
    if not parsed_months:
        today = pd.Timestamp.today().replace(day=1)
        return today.date(), (today + pd.DateOffset(months=5)).date()

    return min(parsed_months).date(), max(parsed_months).date()


def apply_demand_date_range(demand_df, start_value, end_value):
    start_month = pd.to_datetime(start_value, errors="coerce")
    end_month = pd.to_datetime(end_value, errors="coerce")
    if pd.isna(start_month) or pd.isna(end_month):
        raise ValueError("Please select both start and end month.")

    start_month = start_month.replace(day=1)
    end_month = end_month.replace(day=1)
    if start_month > end_month:
        raise ValueError("Start month must be before or equal to end month.")

    updated_df = demand_df.copy()
    old_month_cols = demand_month_columns(updated_df)
    old_month_lookup = {normalize_month_column(col): col for col in old_month_cols}
    id_cols = [col for col in updated_df.columns if col not in old_month_cols]
    new_month_cols = [month.strftime("%Y-%m") for month in pd.date_range(start_month, end_month, freq="MS")]

    output_df = updated_df[id_cols].copy()
    for month_col in new_month_cols:
        if month_col in old_month_lookup:
            output_df[month_col] = updated_df[old_month_lookup[month_col]]
        else:
            output_df[month_col] = 0

    return output_df


def sync_header_labels(table_key, columns, source_key):
    labels_key = f"{table_key}_header_labels"
    source_state_key = f"{table_key}_header_source"
    columns = [str(col) for col in columns]

    if st.session_state.get(source_state_key) != source_key or labels_key not in st.session_state:
        st.session_state[source_state_key] = source_key
        st.session_state[labels_key] = {col: col for col in columns}
    else:
        labels = st.session_state[labels_key]
        st.session_state[labels_key] = {col: labels.get(col, col) for col in columns}

    return st.session_state[labels_key]


def render_header_editor(table_key, columns, source_key):
    columns = [str(col) for col in columns]
    labels = sync_header_labels(table_key, columns, source_key)
    editor_df = pd.DataFrame({
        "Column": columns,
        "Display Header": [labels.get(col, col) for col in columns],
    })

    with st.expander("Edit table headers"):
        edited_df = st.data_editor(
            editor_df,
            width="stretch",
            hide_index=True,
            key=f"{table_key}_header_editor",
            column_config={
                "Column": st.column_config.TextColumn("Column", disabled=True),
                "Display Header": st.column_config.TextColumn("Display Header"),
            },
        )

        if st.button("Apply headers", key=f"{table_key}_apply_headers"):
            new_labels = {}
            for _, row in edited_df.iterrows():
                column = str(row.get("Column", "")).strip()
                label = str(row.get("Display Header", "")).strip()
                if column:
                    new_labels[column] = label or column
            st.session_state[f"{table_key}_header_labels"] = {col: new_labels.get(col, col) for col in columns}
            st.success("Headers updated")

    return st.session_state[f"{table_key}_header_labels"]


def generic_column_config(columns, labels):
    return {col: st.column_config.Column(label=labels.get(str(col), str(col))) for col in columns}


def flow_to_editor_table(flow_df):
    df = planner.clean_columns(flow_df).dropna(how="all")
    if df.empty:
        return pd.DataFrame(columns=["Stage", "Cycle Time Week", "Transit Week", "Yield"])

    if {"Stage", "Cycle Time Week", "Transit Week"}.issubset(df.columns):
        output_df = df.copy()
        for column in ["Cycle Time Week", "Transit Week", "Yield"]:
            if column in output_df.columns:
                output_df[column] = pd.to_numeric(output_df[column], errors="coerce").fillna(0)
        if "Yield" not in output_df.columns:
            output_df["Yield"] = 1.0
        return output_df[["Stage", "Cycle Time Week", "Transit Week", "Yield"]]

    stage_col = planner.find_col(df, ["Stages", "Stage", "Process", "Flow"])
    week_col = planner.find_col(df, ["Time/Week", "Time Week", "Cycle_Time_Week", "Cycle Time Week"])
    yield_col = planner.find_col(df, ["Yield", "Yield Rate"], required=False)

    rows = []
    for _, row in df.iterrows():
        stage_value = row.get(stage_col)
        if pd.isna(stage_value):
            continue

        stage_name = str(stage_value).strip()
        stage_lower = stage_name.lower()
        week_value = pd.to_numeric(row.get(week_col), errors="coerce")
        week_value = 0 if pd.isna(week_value) else float(week_value)

        if "transit" in stage_lower or "transist" in stage_lower:
            if rows:
                rows[-1]["Transit Week"] += week_value
            continue

        yield_value = pd.to_numeric(row.get(yield_col), errors="coerce") if yield_col else 1
        rows.append({
            "Stage": stage_name,
            "Cycle Time Week": week_value,
            "Transit Week": 0.0,
            "Yield": 1.0 if pd.isna(yield_value) else float(yield_value),
        })

    return pd.DataFrame(rows, columns=["Stage", "Cycle Time Week", "Transit Week", "Yield"])


def flow_editor_to_workbook_table(flow_df):
    if not {"Stage", "Cycle Time Week", "Transit Week"}.issubset(flow_df.columns):
        return flow_df

    rows = []
    for _, row in flow_df.dropna(how="all").iterrows():
        stage_name = str(row.get("Stage", "")).strip()
        if not stage_name:
            continue

        cycle_week = pd.to_numeric(row.get("Cycle Time Week"), errors="coerce")
        transit_week = pd.to_numeric(row.get("Transit Week"), errors="coerce")
        yield_value = pd.to_numeric(row.get("Yield"), errors="coerce")

        rows.append({
            "Stages": stage_name,
            "Time/Week": 0 if pd.isna(cycle_week) else cycle_week,
            "Yield": 1 if pd.isna(yield_value) else yield_value,
        })

        if not pd.isna(transit_week) and transit_week > 0:
            rows.append({
                "Stages": "Transist Time",
                "Time/Week": transit_week,
                "Yield": None,
            })

    return pd.DataFrame(rows, columns=["Stages", "Time/Week", "Yield"])


def make_input_bytes(flow_df, product_df, demand_df, inventory_df, min_reach, target_reach, tester_number, tester_smoothing_weeks, wafer_start_before):
    target_df = pd.DataFrame({
        "Columns": ["Min Reach Level", "Target Reach Level", "Tester Number", "Tester Smoothing Weeks", "Wafer Start Time"],
        "Value": [min_reach, target_reach, tester_number, tester_smoothing_weeks, wafer_start_before],
    })

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        flow_editor_to_workbook_table(flow_df).to_excel(writer, sheet_name=planner.FLOW_SHEET, index=False)
        product_df.to_excel(writer, sheet_name=planner.PRODUCT_SHEET, index=False)
        demand_df.to_excel(writer, sheet_name=planner.DEMAND_SHEET, index=False)
        inventory_df.to_excel(writer, sheet_name=planner.INVENTORY_SHEET, index=False)
        target_df.to_excel(writer, sheet_name=planner.TARGET_SHEET, index=False)
    output.seek(0)
    return output.getvalue()


def parse_wafer_start_before(value):
    if not str(value).strip():
        return None

    wafer_start_before = int(float(str(value).strip()))
    if wafer_start_before < 0:
        raise ValueError("Wafer Start Time must be blank or a non-negative number of weeks.")
    return wafer_start_before


def build_product_lookup(product_df):
    df = planner.clean_columns(product_df).dropna(how="all")
    if df.empty:
        return {}

    product_col = planner.find_col(df, ["Product"])
    basic_col = planner.find_col(df, ["Basic Type", "Basic_Type", "Type"], required=False)
    cpw_col = planner.find_col(df, ["CPW"])
    weekly_out_col = planner.find_col(df, ["Wafer / tester / week", "Wafer/tester/week", "Weekly Output", "weekly_output", "Tester Output", "Output Per Week", "UPH"])
    priority_col = planner.find_col(df, ["Priority", "Priority Score", "Priority_Score"], required=False)
    total_yield_col = planner.find_col(df, ["Total Yield", "Total_Yield", "Yield", "Product Yield"], required=False)
    min_reach_col = planner.find_col(df, ["Min Reach Level", "Min Reach", "Min_REACH"], required=False)
    target_reach_col = planner.find_col(df, ["Target Reach Level", "Target Reach", "Target_REACH"], required=False)
    earliest_start_col = planner.find_col(df, ["Earliest Wafer Start Time", "Earliest Wafer Start", "Earliest Start Time"], required=False)

    products = {}
    for _, row in df.iterrows():
        product_name = row.get(product_col)
        product_key = planner.normalize_key(product_name)
        if not product_key or product_key in products:
            continue

        cpw = pd.to_numeric(row.get(cpw_col), errors="coerce")
        weekly_out = pd.to_numeric(row.get(weekly_out_col), errors="coerce")
        if pd.isna(cpw) or pd.isna(weekly_out) or float(cpw) <= 0 or float(weekly_out) <= 0:
            continue

        total_yield = None
        if total_yield_col is not None and not pd.isna(row.get(total_yield_col)):
            total_yield = planner.normalize_yield(row.get(total_yield_col))
        priority = pd.to_numeric(row.get(priority_col), errors="coerce") if priority_col is not None else 0
        min_reach = pd.to_numeric(row.get(min_reach_col), errors="coerce") if min_reach_col is not None else None
        target_reach = pd.to_numeric(row.get(target_reach_col), errors="coerce") if target_reach_col is not None else None
        earliest_start = parse_earliest_wafer_start_time(row.get(earliest_start_col)) if earliest_start_col is not None else None

        products[product_key] = {
            "Product_Key": product_key,
            "Basic_Type": str(row.get(basic_col)).strip() if basic_col and not pd.isna(row.get(basic_col)) else "",
            "Product": str(product_name).strip(),
            "CPW": float(cpw),
            "weekly_output": float(weekly_out),
            "Priority": 0.0 if pd.isna(priority) else float(priority),
            "Total_Yield": total_yield,
            "Min_Reach": None if min_reach is None or pd.isna(min_reach) else float(min_reach),
            "Target_Reach": None if target_reach is None or pd.isna(target_reach) else float(target_reach),
            "Earliest_Wafer_Start": earliest_start,
        }
    return products


def validate_earliest_wafer_start(wafer_by_product, product_lookup, context):
    violations = []
    for product_key, values in wafer_by_product.items():
        params = product_lookup.get(product_key, {})
        earliest_week = params.get("Earliest_Wafer_Start")
        if not earliest_week:
            continue
        for week, wafer in values.items():
            if float(wafer or 0) > 1e-9 and week_sort_key(week) < week_sort_key(earliest_week):
                violations.append({
                    "Product": params.get("Product", product_key),
                    "Wafer Start Week": week,
                    "Earliest Wafer Start Time": earliest_week,
                    "Wafer": float(wafer),
                })

    if violations:
        preview = "; ".join(
            f"{row['Product']} {week_display_label(row['Wafer Start Week'])} < {week_display_label(row['Earliest Wafer Start Time'])}"
            for row in violations[:8]
        )
        more = " ..." if len(violations) > 8 else ""
        raise ValueError(f"{context} violates product Earliest Wafer Start Time. Cannot calculate: {preview}{more}")


def cap_wafer_start_to_target(source_wafer_by_product, target_wafer_by_product, product_lookup):
    capped = {}
    reduction_rows = []
    for product_key in sorted(set(source_wafer_by_product) | set(target_wafer_by_product)):
        params = product_lookup.get(product_key, {})
        product_name = params.get("Product", product_key)
        capped_weeks = {}
        all_weeks = sorted(
            set(source_wafer_by_product.get(product_key, {})) | set(target_wafer_by_product.get(product_key, {})),
            key=week_sort_key,
        )
        for week in all_weeks:
            source_wafer = max(0.0, float(source_wafer_by_product.get(product_key, {}).get(week, 0.0)))
            target_wafer = max(0.0, float(target_wafer_by_product.get(product_key, {}).get(week, 0.0)))
            capped_wafer = min(source_wafer, target_wafer) if source_wafer > 0 else 0.0
            if capped_wafer > 1e-9:
                capped_weeks[week] = capped_wafer
            if source_wafer > capped_wafer + 1e-9:
                reduction_rows.append({
                    "Product": product_name,
                    "Wafer Start Week": week,
                    "Input Wafer Start": source_wafer,
                    "Target Reach Max Wafer Start": target_wafer,
                    "Used Wafer Start": capped_wafer,
                    "Reduced Wafer Start": source_wafer - capped_wafer,
                    "Reason": "WSPW input exceeds Target Reach ideal production",
                })
        capped[product_key] = capped_weeks
    return capped, pd.DataFrame(reduction_rows)


def build_min_reach_floor_wafer_start(demand_by_product, demand_week_labels, product_lookup, inventory_lookup, flow_df, min_reach_default):
    offsets = build_flow_offsets(flow_df)
    flow_yield = total_yield_from_flow(flow_df)
    reach_window = int(planner.DEFAULT_TESTER_CONFIG.get("reach_window", 4))
    wafer_by_product = {}

    for product_key, params in product_lookup.items():
        demand_map = demand_by_product.get(product_key, {})
        if not demand_map:
            continue

        output_per_wafer = params["CPW"] * (params["Total_Yield"] if params["Total_Yield"] is not None else flow_yield)
        if output_per_wafer <= 0:
            continue

        min_reach = product_reach_setting(params, "Min_Reach", min_reach_default)
        stock = float(inventory_lookup.get(product_key, 0))
        wafer_by_week = {}
        for label in demand_week_labels:
            demand = float(demand_map.get(label, 0.0))
            future_labels = [shift_week_label(label, offset) for offset in range(1, reach_window + 1)]
            future_avg = sum(float(demand_map.get(future_label, 0.0)) for future_label in future_labels) / len(future_labels)
            target_stock = max(0.0, float(min_reach)) * future_avg
            required_output = max(0.0, demand + target_stock - stock)
            wafer = int(math.ceil(required_output / output_per_wafer)) if required_output > 0 else 0
            if wafer > 0:
                source_week = shift_week_label(label, -offsets["dc"])
                wafer_by_week[source_week] = wafer_by_week.get(source_week, 0.0) + wafer
            stock = max(0.0, stock + wafer * output_per_wafer - demand)
        wafer_by_product[product_key] = wafer_by_week

    return wafer_by_product


def build_inventory_lookup(inventory_df):
    df = planner.clean_columns(inventory_df).dropna(how="all")
    if df.empty:
        return {}

    product_col = planner.find_col(df, ["Product"], required=False)
    inventory_col = planner.find_col(df, ["Existing DC inventory", "Existing DC inventories", "Exisitng DC inventories", "Inventory", "Stock"], required=False)
    if product_col is None or inventory_col is None:
        return {}

    df["_Product_Key"] = df[product_col].apply(planner.normalize_key)
    df["_Inventory"] = pd.to_numeric(df[inventory_col], errors="coerce").fillna(0)
    return df[df["_Product_Key"] != ""].groupby("_Product_Key")["_Inventory"].sum().to_dict()


def build_flow_offsets(flow_df):
    editor_flow_df = flow_to_editor_table(flow_df)
    offsets = {"bump": 0, "sort": 0, "dps": 0, "dc": 0, "total": 0}
    found = set()
    cumulative_week = 0

    for _, row in editor_flow_df.dropna(how="all").iterrows():
        stage_name = str(row.get("Stage", "")).strip().lower()
        if not stage_name:
            continue

        cycle_week = pd.to_numeric(row.get("Cycle Time Week"), errors="coerce")
        transit_week = pd.to_numeric(row.get("Transit Week"), errors="coerce")
        stage_weeks = 0 if pd.isna(cycle_week) else int(round(float(cycle_week)))
        stage_weeks += 0 if pd.isna(transit_week) else int(round(float(transit_week)))
        output_week = cumulative_week + stage_weeks

        if "bump" in stage_name and "bump" not in found:
            offsets["bump"] = cumulative_week
            found.add("bump")
        if ("sort" in stage_name or "test" in stage_name) and "sort" not in found:
            offsets["sort"] = cumulative_week
            found.add("sort")
        if "dps" in stage_name and "dps" not in found:
            offsets["dps"] = cumulative_week
            found.add("dps")
        if stage_name == "dc" and "dc" not in found:
            offsets["dc"] = cumulative_week
            found.add("dc")

        cumulative_week = output_week

    offsets["total"] = cumulative_week
    if "dc" not in found:
        offsets["dc"] = offsets["dps"]
    return offsets


def reach_level_by_week(stock_by_week, demand_by_week, ordered_week_labels, reach_window=None):
    reach_window = int(reach_window or planner.DEFAULT_TESTER_CONFIG.get("reach_window", 4))
    reach_window = max(1, reach_window)
    reach = {}
    for label in ordered_week_labels:
        future_labels = [shift_week_label(label, offset) for offset in range(1, reach_window + 1)]
        future_avg_demand = sum(float(demand_by_week.get(future_label, 0.0)) for future_label in future_labels) / len(future_labels)
        reach[label] = float(stock_by_week.get(label, 0.0)) / future_avg_demand if future_avg_demand > 0 else 0.0
    return reach


def total_yield_from_flow(flow_df):
    editor_flow_df = flow_to_editor_table(flow_df)
    total_yield = 1.0
    for _, row in editor_flow_df.iterrows():
        yield_value = pd.to_numeric(row.get("Yield"), errors="coerce")
        if not pd.isna(yield_value):
            total_yield *= planner.normalize_yield(yield_value)
    return total_yield


def build_weekly_demand_by_product(demand_df):
    df = planner.clean_columns(demand_df).dropna(how="all")
    if df.empty:
        return {}, []

    product_col = planner.find_col(df, ["Product"])
    month_cols = [col for col in df.columns if col != product_col and planner.is_month_col(col)]
    demand_by_product = {}
    week_labels = []

    for _, row in df.iterrows():
        product_key = planner.normalize_key(row.get(product_col))
        if not product_key:
            continue

        product_demand = demand_by_product.setdefault(product_key, {})
        for month_col in month_cols:
            month_label = normalize_month_column(month_col)
            labels = month_week_labels(month_label)
            month_demand = pd.to_numeric(row.get(month_col), errors="coerce")
            weekly_demand = 0.0 if pd.isna(month_demand) else float(month_demand) / len(labels)
            for label in labels:
                product_demand[label] = product_demand.get(label, 0.0) + weekly_demand
                week_labels.append(label)

    ordered_week_labels = sorted(set(week_labels), key=week_sort_key)
    return demand_by_product, ordered_week_labels


def week_range(start_label, end_label):
    labels = []
    current = str(start_label)
    while week_sort_key(current) <= week_sort_key(end_label):
        labels.append(current)
        current = shift_week_label(current, 1)
        if len(labels) > 2000:
            break
    return labels


def expand_week_horizon(base_labels, pre_weeks=0, post_weeks=0):
    if not base_labels:
        return []
    start_label = min(base_labels, key=week_sort_key)
    end_label = max(base_labels, key=week_sort_key)
    start_label = shift_week_label(start_label, -int(pre_weeks))
    end_label = shift_week_label(end_label, int(post_weeks))
    return week_range(start_label, end_label)


def table_to_week_map(table_df, product_lookup):
    week_cols = [col for col in table_df.columns if re.fullmatch(r"\d{4}-CW\d{2}", str(col))]
    wafer_by_product = {}
    for _, row in table_df.iterrows():
        product_key = planner.normalize_key(row.get("RFP"))
        if not product_key or product_key not in product_lookup:
            continue
        values = {
            str(col): float(pd.to_numeric(row.get(col), errors="coerce"))
            for col in week_cols
            if not pd.isna(pd.to_numeric(row.get(col), errors="coerce"))
        }
        if any(value != 0 for value in values.values()):
            wafer_by_product[product_key] = values
    return wafer_by_product


def build_initial_target_reach_wafer_start(demand_by_product, demand_week_labels, product_lookup, inventory_lookup, flow_df, target_reach):
    offsets = build_flow_offsets(flow_df)
    flow_yield = total_yield_from_flow(flow_df)
    reach_window = int(planner.DEFAULT_TESTER_CONFIG.get("reach_window", 4))
    wafer_by_product = {}

    for product_key, params in product_lookup.items():
        demand_map = demand_by_product.get(product_key, {})
        if not demand_map:
            continue

        output_per_wafer = params["CPW"] * (params["Total_Yield"] if params["Total_Yield"] is not None else flow_yield)
        if output_per_wafer <= 0:
            continue

        stock = float(inventory_lookup.get(product_key, 0))
        wafer_by_week = {}
        product_target_reach = product_reach_setting(params, "Target_Reach", target_reach)
        for label in demand_week_labels:
            demand = float(demand_map.get(label, 0.0))
            future_labels = [shift_week_label(label, offset) for offset in range(1, reach_window + 1)]
            future_avg = sum(float(demand_map.get(future_label, 0.0)) for future_label in future_labels) / len(future_labels)
            target_stock = float(product_target_reach) * future_avg
            required_output = max(0.0, demand + target_stock - stock)
            wafer = int(math.ceil(required_output / output_per_wafer)) if required_output > 0 else 0
            if wafer > 0:
                source_week = shift_week_label(label, -offsets["dc"])
                wafer_by_week[source_week] = wafer_by_week.get(source_week, 0.0) + wafer
            stock = max(0.0, stock + wafer * output_per_wafer - demand)
        wafer_by_product[product_key] = wafer_by_week

    return wafer_by_product


def shift_week_values(values_by_week, offset):
    shifted = {}
    for label, value in values_by_week.items():
        shifted_label = shift_week_label(label, offset)
        shifted[shifted_label] = shifted.get(shifted_label, 0.0) + float(value)
    return shifted


def build_stage_maps_from_wafer_start(wafer_by_product, product_lookup, demand_by_product, flow_df):
    offsets = build_flow_offsets(flow_df)
    flow_yield = total_yield_from_flow(flow_df)
    all_week_labels = set()
    product_stage_maps = []

    for product_key, wafer_start in wafer_by_product.items():
        params = product_lookup.get(product_key)
        if params is None:
            continue

        product_yield = params["Total_Yield"] if params["Total_Yield"] is not None else flow_yield
        gross_chip_per_wafer = params["CPW"]
        good_chip_per_wafer = gross_chip_per_wafer * product_yield
        bump = shift_week_values(wafer_start, offsets["bump"])
        sort = shift_week_values(wafer_start, offsets["sort"])
        tester = {label: value / params["weekly_output"] for label, value in sort.items()}
        dps = {label: value * gross_chip_per_wafer for label, value in shift_week_values(wafer_start, offsets["dps"]).items()}
        dc = {label: value * good_chip_per_wafer for label, value in shift_week_values(wafer_start, offsets["dc"]).items()}
        demand = demand_by_product.get(product_key, {})

        for values in [wafer_start, bump, sort, tester, dps, dc]:
            all_week_labels.update(values.keys())

        product_stage_maps.append({
            "params": params,
            "product_key": product_key,
            "Wafer Start": wafer_start,
            "Bump Wafer": bump,
            "Sort Wafer": sort,
            "Tester Required": tester,
            "Tester Used": tester,
            "DPS Chip": dps,
            "DC Chip": dc,
            "Demand": demand,
        })

    return product_stage_maps, all_week_labels


def build_initial_wafer_start_map(wspw_wafer_start_df, flow_df, product_df, demand_df, inventory_df, target_reach):
    product_lookup = build_product_lookup(product_df)
    inventory_lookup = build_inventory_lookup(inventory_df)
    demand_by_product, demand_week_labels = build_weekly_demand_by_product(demand_df)
    wspw_wafer_start_df = repair_week_sequence_columns(wspw_wafer_start_df)
    wafer_by_product = table_to_week_map(wspw_wafer_start_df, product_lookup)
    source_type = "WSPW Wafer Start" if wafer_by_product else "Target Reach Backward Calculation"

    if not wafer_by_product:
        wafer_by_product = build_initial_target_reach_wafer_start(demand_by_product, demand_week_labels, product_lookup, inventory_lookup, flow_df, target_reach)

    return wafer_by_product, product_lookup, inventory_lookup, demand_by_product, source_type


def build_loading_view_from_wafer_start(wafer_by_product, product_lookup, inventory_lookup, demand_by_product, flow_df, tester_capacity, extra_week_labels=None):
    product_stage_values, all_week_labels = build_stage_maps_from_wafer_start(wafer_by_product, product_lookup, demand_by_product, flow_df)
    tester_by_week = {}

    def loading_metric_value(metric, value):
        numeric_value = float(value or 0.0)
        if metric == "Reach Level":
            return numeric_value
        return int(round(numeric_value))

    for product_values in product_stage_values:
        tester = product_values["Tester Required"]
        for label, value in tester.items():
            tester_by_week[label] = tester_by_week.get(label, 0.0) + value

    if extra_week_labels:
        all_week_labels.update(extra_week_labels)

    if all_week_labels:
        loading_start_week = min(all_week_labels, key=week_sort_key)
        for product_values in product_stage_values:
            all_week_labels.update(
                label for label in product_values["Demand"].keys()
                if week_sort_key(label) >= week_sort_key(loading_start_week)
            )

    ordered_week_labels = expand_week_horizon(list(all_week_labels), pre_weeks=0, post_weeks=0)
    for product_values in product_stage_values:
        product_key = product_values["product_key"]
        stock = {}
        stock_value = float(inventory_lookup.get(product_key, 0))
        dc = product_values["DC Chip"]
        demand = product_values["Demand"]
        for label in ordered_week_labels:
            stock_value = max(0.0, stock_value + dc.get(label, 0.0) - demand.get(label, 0.0))
            stock[label] = stock_value
        product_values["Stock Chip"] = stock
        product_values["Reach Level"] = reach_level_by_week(stock, demand, ordered_week_labels)

    metric_order = ["Wafer Start", "Bump Wafer", "Sort Wafer", "DPS Chip", "DC Chip", "Demand", "Stock Chip", "Reach Level"]
    loading_rows = []
    for metric in metric_order:
        metric_entries = [(product_values["params"], product_values[metric]) for product_values in product_stage_values]
        if not metric_entries:
            continue

        loading_rows.append({"Stage": metric, "Basic Type": "", "RFP": "", **{label: "" for label in ordered_week_labels}})
        totals = {label: 0.0 for label in ordered_week_labels}
        for params, values in metric_entries:
            row = {
                "Stage": "",
                "Basic Type": params["Basic_Type"],
                "RFP": params["Product"],
            }
            for label in ordered_week_labels:
                value = values.get(label, 0.0)
                row_value = loading_metric_value(metric, value)
                row[label] = row_value
                if metric != "Reach Level":
                    totals[label] += float(row_value)
            loading_rows.append(row)

        total_row = {"Stage": "", "Basic Type": "", "RFP": "Grand Total"}
        if metric == "Reach Level":
            totals = {label: "" for label in ordered_week_labels}
        total_row.update(totals)
        loading_rows.append(total_row)

    tester_total = {label: tester_by_week.get(label, 0.0) for label in ordered_week_labels}
    tester_capacity_row = {label: float(tester_capacity) for label in ordered_week_labels}
    tester_remaining = {label: tester_capacity_row[label] - tester_total[label] for label in ordered_week_labels}
    tester_over_limit = {label: tester_total[label] > tester_capacity_row[label] + 1e-9 for label in ordered_week_labels}
    loading_rows.append({"Stage": "Tester Summary", "Basic Type": "", "RFP": "", **{label: "" for label in ordered_week_labels}})
    loading_rows.append({"Stage": "", "Basic Type": "", "RFP": "Total Tester Required", **tester_total})
    loading_rows.append({"Stage": "", "Basic Type": "", "RFP": "Tester Capacity", **tester_capacity_row})
    loading_rows.append({"Stage": "", "Basic Type": "", "RFP": "Remaining Tester", **tester_remaining})
    loading_rows.append({"Stage": "", "Basic Type": "", "RFP": "Tester Over Limit", **tester_over_limit})

    loading_df = pd.DataFrame(loading_rows, columns=["Stage", "Basic Type", "RFP", *ordered_week_labels])
    tester_summary_df = pd.DataFrame({
        "Week": ordered_week_labels,
        "Total Tester Required": [tester_total.get(label, 0.0) for label in ordered_week_labels],
    })
    tester_summary_df["Tester Capacity"] = float(tester_capacity)
    tester_summary_df["Remaining Tester"] = tester_summary_df["Tester Capacity"] - tester_summary_df["Total Tester Required"]
    tester_summary_df["Tester Over Limit"] = tester_summary_df["Total Tester Required"] > tester_summary_df["Tester Capacity"] + 1e-9
    return loading_df, tester_summary_df


def build_simple_weekly_outputs(wafer_by_product, product_lookup, inventory_lookup, demand_by_product, flow_df, tester_capacity, target_reach):
    product_stage_maps, all_week_labels = build_stage_maps_from_wafer_start(wafer_by_product, product_lookup, demand_by_product, flow_df)
    product_metrics = {product_values["product_key"]: product_values for product_values in product_stage_maps}

    if all_week_labels:
        output_start_week = min(all_week_labels, key=week_sort_key)
        for metrics in product_metrics.values():
            all_week_labels.update(
                label for label in metrics["Demand"].keys()
                if week_sort_key(label) >= week_sort_key(output_start_week)
            )

    ordered_weeks = expand_week_horizon(list(all_week_labels), pre_weeks=0, post_weeks=0)
    for product_key, metrics in product_metrics.items():
        stock = {}
        target_stock = {}
        stock_value = float(inventory_lookup.get(product_key, 0))
        demand = metrics["Demand"]
        product_target_reach = product_reach_setting(metrics["params"], "Target_Reach", target_reach)
        future_window = max(1, int(math.ceil(float(product_target_reach))))

        for label in ordered_weeks:
            stock_value = max(0.0, stock_value + metrics["DC Chip"].get(label, 0.0) - demand.get(label, 0.0))
            stock[label] = stock_value
            future_labels = [shift_week_label(label, offset) for offset in range(1, future_window + 1)]
            future_avg = sum(float(demand.get(future_label, 0.0)) for future_label in future_labels) / len(future_labels)
            target_stock[label] = float(product_target_reach) * future_avg

        metrics["Stock Chip"] = stock
        metrics["Reach Level"] = reach_level_by_week(stock, demand, ordered_weeks)
        metrics["Target Stock"] = target_stock

    def product_week_table(metric_name, id_columns=("Basic Type", "Row Labels")):
        rows = []
        totals = {label: 0.0 for label in ordered_weeks}
        for metrics in product_metrics.values():
            params = metrics["params"]
            row = {id_columns[0]: params["Basic_Type"], id_columns[1]: params["Product"]}
            for label in ordered_weeks:
                value = float(metrics[metric_name].get(label, 0.0))
                row[label] = value
                totals[label] += value
            rows.append(row)
        rows.append({id_columns[0]: "Grand Total", id_columns[1]: "Grand Total", **totals})
        return display_week_table(pd.DataFrame(rows, columns=[*id_columns, *ordered_weeks]))

    wafer_start_table = product_week_table("Wafer Start")

    def metric_output_table(metric_names):
        rows = []
        for metric_name in metric_names:
            totals = {label: 0.0 for label in ordered_weeks}
            for metrics in product_metrics.values():
                params = metrics["params"]
                row = {"Metric": metric_name, "Basic Type": params["Basic_Type"], "Row Labels": params["Product"]}
                for label in ordered_weeks:
                    value = float(metrics[metric_name].get(label, 0.0))
                    row[label] = value
                    if metric_name != "Reach Level":
                        totals[label] += value
                rows.append(row)
            if metric_name == "Reach Level":
                totals = {label: "" for label in ordered_weeks}
            rows.append({"Metric": metric_name, "Basic Type": "Grand Total", "Row Labels": "Grand Total", **totals})
        return display_week_table(pd.DataFrame(rows, columns=["Metric", "Basic Type", "Row Labels", *ordered_weeks]))

    prebuild_output_table = metric_output_table(["Bump Wafer", "Sort Wafer"])
    build_output_table = metric_output_table(["DPS Chip", "DC Chip", "Demand", "Stock Chip", "Reach Level"])

    stage_rows = []
    for metric_name in ["Bump Wafer", "Sort Wafer", "DPS Chip", "DC Chip", "Demand", "Stock Chip", "Reach Level"]:
        totals = {label: 0.0 for label in ordered_weeks}
        for metrics in product_metrics.values():
            params = metrics["params"]
            row = {"Metric": metric_name, "Basic Type": params["Basic_Type"], "Row Labels": params["Product"]}
            for label in ordered_weeks:
                value = float(metrics[metric_name].get(label, 0.0))
                row[label] = value
                if metric_name != "Reach Level":
                    totals[label] += value
            stage_rows.append(row)
        if metric_name == "Reach Level":
            totals = {label: "" for label in ordered_weeks}
        stage_rows.append({"Metric": metric_name, "Basic Type": "Grand Total", "Row Labels": "Grand Total", **totals})
    stage_output_table = display_week_table(pd.DataFrame(stage_rows, columns=["Metric", "Basic Type", "Row Labels", *ordered_weeks]))

    def graph_table(metric_name, total_column="Grand Total", extra_total_metric=None, include_total=True):
        rows = []
        for label in ordered_weeks:
            row = {"Row Labels": week_display_label(label)}
            total = 0.0
            for metrics in product_metrics.values():
                product = metrics["params"]["Product"]
                value = float(metrics[metric_name].get(label, 0.0))
                row[product] = value
                total += value
            if include_total:
                row[total_column] = total
            if extra_total_metric is not None:
                row[extra_total_metric] = sum(float(metrics[extra_total_metric].get(label, 0.0)) for metrics in product_metrics.values())
            rows.append(row)
        return pd.DataFrame(rows)

    tester_graph_table = graph_table("Tester Used")
    tester_graph_table["Tester Capacity"] = float(tester_capacity)
    tester_usage_table = tester_graph_table.rename(columns={"Grand Total": "Total Tester Used"}).copy()
    tester_usage_table["Remaining Tester"] = tester_usage_table["Tester Capacity"] - tester_usage_table["Total Tester Used"]
    tester_usage_table["Tester OK"] = tester_usage_table["Total Tester Used"] <= tester_usage_table["Tester Capacity"]

    demand_graph_table = graph_table("Demand")
    stock_graph_table = graph_table("Stock Chip", extra_total_metric="Target Stock")
    reach_graph_table = graph_table("Reach Level", include_total=False)

    graph_outputs = [
        ("Tester Used", tester_graph_table.rename(columns={"Total Tester Used": "Grand Total"}), "tester"),
        ("VRFN Demand", demand_graph_table, "bar"),
        ("Reach Level", reach_graph_table, "line"),
        ("Stock Development", stock_graph_table, "stock"),
    ]

    return {
        "graph_outputs": graph_outputs,
        "wafer_start_table": wafer_start_table,
        "prebuild_output_table": prebuild_output_table,
        "build_output_table": build_output_table,
        "stage_output_table": stage_output_table,
        "tester_usage_table": tester_usage_table,
    }


def first_positive_demand_week(demand_by_product):
    demand_weeks = [
        label
        for values in demand_by_product.values()
        for label, value in values.items()
        if float(value or 0) > 1e-9
    ]
    return min(demand_weeks, key=week_sort_key) if demand_weeks else None


def filter_phase_week_columns(output_df, allowed_display_labels):
    week_cols = week_columns(output_df)
    keep_cols = []
    for col in output_df.columns:
        if col not in week_cols:
            keep_cols.append(col)
            continue
        display_label = week_display_label(col) if is_internal_week_label(col) else str(col)
        if display_label in allowed_display_labels:
            keep_cols.append(col)
    return output_df.loc[:, keep_cols].copy()


def filter_phase_row_labels(output_df, allowed_display_labels):
    if output_df.empty or "Row Labels" not in output_df.columns:
        return output_df.copy()
    return output_df[output_df["Row Labels"].astype(str).isin(allowed_display_labels)].reset_index(drop=True)


def filter_phase_week_rows(output_df, allowed_display_labels):
    if output_df.empty or "Week" not in output_df.columns:
        return output_df.copy()
    return output_df[output_df["Week"].astype(str).isin(allowed_display_labels)].reset_index(drop=True)


def split_simple_outputs_by_phase(simple_outputs, loading_table, tester_summary_table, demand_by_product):
    first_demand_week = first_positive_demand_week(demand_by_product)
    all_weeks = table_week_labels(loading_table)
    if not first_demand_week:
        prebuild_weeks = all_weeks
        demand_weeks = []
    else:
        prebuild_weeks = [week for week in all_weeks if week_sort_key(week) < week_sort_key(first_demand_week)]
        demand_weeks = [week for week in all_weeks if week_sort_key(week) >= week_sort_key(first_demand_week)]

    phase_week_sets = {
        "Prebuild Phase": {week_display_label(week) for week in prebuild_weeks},
        "Demand Phase": {week_display_label(week) for week in demand_weeks},
    }

    def graph_rows(graph_df, allowed_labels):
        filtered_df = filter_phase_row_labels(graph_df, allowed_labels)
        value_cols = [col for col in filtered_df.columns if col not in ["Row Labels", "Tester Capacity", "Target Stock"]]
        if value_cols and not filtered_df.empty:
            numeric_values = filtered_df[value_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
            filtered_df = filtered_df[numeric_values.sum(axis=1) > 1e-9].reset_index(drop=True)
        return filtered_df

    reach_tester_graphs = [item for item in simple_outputs["graph_outputs"] if item[2] == "reach_tester"]
    normal_graphs = [
        item for item in simple_outputs["graph_outputs"]
        if item[2] != "reach_tester" and not (item[0].startswith("Reach Level ") and item[2] == "tester")
    ]

    phase_outputs = {}
    for phase_name, allowed_labels in phase_week_sets.items():
        phase_graphs = []
        if phase_name == "Prebuild Phase":
            phase_graphs.extend(reach_tester_graphs)
        for title, graph_df, chart_kind in normal_graphs:
            if phase_name == "Prebuild Phase" and title not in ["Tester Used", "Stock Development"]:
                continue
            if phase_name == "Demand Phase" and title not in ["Tester Used", "VRFN Demand", "Reach Level", "Stock Development"]:
                continue
            filtered_graph = graph_rows(graph_df, allowed_labels)
            if not filtered_graph.empty:
                phase_graphs.append((title, filtered_graph, chart_kind))

        phase_tables = []
        if phase_name == "Prebuild Phase":
            tester_table = filter_phase_row_labels(simple_outputs["tester_usage_table"], allowed_labels)
            if not tester_table.empty:
                phase_tables.append(("Tester Allocation", tester_table))
            phase_tables.append(("Wafer Start", filter_phase_week_columns(simple_outputs["wafer_start_table"], allowed_labels)))
            phase_tables.append(("Bump / Sort", filter_phase_week_columns(simple_outputs["prebuild_output_table"], allowed_labels)))
            phase_tables.append(("Prebuild Stock", filter_phase_week_columns(simple_outputs["build_output_table"], allowed_labels)))
        phase_tables.append(("Balanced Loading View", filter_phase_week_columns(display_week_table(loading_table), allowed_labels)))
        if phase_name == "Prebuild Phase":
            phase_tables.append(("Balanced Tester Summary", filter_phase_week_rows(display_week_table(tester_summary_table), allowed_labels)))

        phase_outputs[phase_name] = {
            "graphs": phase_graphs,
            "tables": phase_tables,
            "week_labels": sorted(allowed_labels),
        }

    return phase_outputs


def build_initial_loading_view(wspw_wafer_start_df, flow_df, product_df, demand_df, inventory_df, target_reach, tester_capacity):
    wafer_by_product, product_lookup, inventory_lookup, demand_by_product, source_type = build_initial_wafer_start_map(
        wspw_wafer_start_df,
        flow_df,
        product_df,
        demand_df,
        inventory_df,
        target_reach,
    )
    loading_df, tester_summary_df = build_loading_view_from_wafer_start(
        wafer_by_product,
        product_lookup,
        inventory_lookup,
        demand_by_product,
        flow_df,
        tester_capacity,
    )
    return loading_df, tester_summary_df, source_type


def build_reach_tester_sensitivity_graph(product_lookup, inventory_lookup, demand_by_product, flow_df, tester_capacity, reach_levels=range(1, 7)):
    demand_week_labels = sorted({week for values in demand_by_product.values() for week in values}, key=week_sort_key)
    if not demand_week_labels:
        return pd.DataFrame(columns=["Row Labels", "Tester Capacity"])

    week_values = {f"Target Reach = {reach_level}": {week: 0.0 for week in demand_week_labels} for reach_level in reach_levels}
    for reach_level in reach_levels:
        reach_col = f"Target Reach = {reach_level}"
        for product_key, params in product_lookup.items():
            demand_map = demand_by_product.get(product_key, {})
            cpw = float(params.get("CPW", 0.0))
            weekly_output = float(params.get("weekly_output", 0.0))
            if not demand_map or cpw <= 0 or weekly_output <= 0:
                continue

            previous_stock = float(inventory_lookup.get(product_key, 0.0))
            for week in demand_week_labels:
                use = float(demand_map.get(week, 0.0))
                future_weeks = [shift_week_label(week, offset) for offset in range(1, 5)]
                target_stock = sum(float(demand_map.get(future_week, 0.0)) for future_week in future_weeks) / 4 * float(reach_level)
                required_chip = max(0.0, use + target_stock - previous_stock)
                week_values[reach_col][week] += required_chip / cpw / weekly_output
                previous_stock = target_stock

    rows = []
    for week in demand_week_labels:
        row = {"Row Labels": week_display_label(week)}
        for reach_col, values in week_values.items():
            row[reach_col] = values.get(week, 0.0)
        row["Tester Capacity"] = float(tester_capacity)
        rows.append(row)
    output_df = pd.DataFrame(rows)
    reach_cols = [col for col in output_df.columns if str(col).startswith("Target Reach = ")]
    if reach_cols:
        output_df = output_df[output_df[reach_cols].sum(axis=1) > 1e-9].reset_index(drop=True)
    return output_df


def build_reach_tester_product_breakdown_graphs(product_lookup, inventory_lookup, demand_by_product, tester_capacity, reach_levels=(1, 4)):
    demand_week_labels = sorted({week for values in demand_by_product.values() for week in values}, key=week_sort_key)
    if not demand_week_labels:
        return []

    graph_outputs = []
    active_products = [product_key for product_key in sorted(demand_by_product) if product_key in product_lookup]
    for reach_level in reach_levels:
        rows = [{"Row Labels": week_display_label(week)} for week in demand_week_labels]
        for product_key in active_products:
            params = product_lookup[product_key]
            product_name = str(params.get("Product") or product_key).strip() or product_key
            demand_map = demand_by_product.get(product_key, {})
            cpw = float(params.get("CPW", 0.0))
            weekly_output = float(params.get("weekly_output", 0.0))
            previous_stock = float(inventory_lookup.get(product_key, 0.0))
            values = []

            for week in demand_week_labels:
                use = float(demand_map.get(week, 0.0))
                future_weeks = [shift_week_label(week, offset) for offset in range(1, 5)]
                target_stock = sum(float(demand_map.get(future_week, 0.0)) for future_week in future_weeks) / 4 * float(reach_level)
                required_chip = max(0.0, use + target_stock - previous_stock)
                tester_required = required_chip / cpw / weekly_output if cpw > 0 and weekly_output > 0 else 0.0
                values.append(tester_required)
                previous_stock = target_stock

            for row, value in zip(rows, values):
                row[product_name] = value

        for row in rows:
            row["Tester Capacity"] = float(tester_capacity)

        output_df = pd.DataFrame(rows)
        product_cols = [col for col in output_df.columns if col not in ["Row Labels", "Tester Capacity"]]
        if product_cols:
            output_df = output_df[output_df[product_cols].sum(axis=1) > 1e-9].reset_index(drop=True)
        graph_outputs.append((f"Reach Level {reach_level} Tester Need by Product", output_df, "tester"))
    return graph_outputs


def tester_violations_from_summary(tester_summary_df):
    if tester_summary_df.empty:
        return []
    over_rows = tester_summary_df[tester_summary_df["Tester Over Limit"].astype(bool)]
    return [
        (row["Week"], float(row["Total Tester Required"]))
        for _, row in over_rows.iterrows()
    ]


def loading_metric_rows_by_product(loading_df, metric):
    if loading_df.empty or "Stage" not in loading_df.columns:
        return {}
    matches = loading_df[loading_df["Stage"].astype(str).eq(metric)]
    if matches.empty:
        return {}

    rows = {}
    index = int(matches.index[0]) + 1
    while index < len(loading_df) and not str(loading_df.iloc[index].get("Stage", "")).strip():
        row = loading_df.iloc[index]
        product_key = planner.normalize_key(row.get("RFP"))
        if product_key and product_key.lower() != "grandtotal":
            rows[product_key] = row
        if str(row.get("RFP", "")).strip() == "Grand Total":
            break
        index += 1
    return rows


def product_reach_setting(params, key, default_value):
    value = params.get(key)
    return float(default_value) if value is None or pd.isna(value) else float(value)


def validate_product_reach_settings(product_lookup, min_reach_default, target_reach_default):
    invalid_rows = []
    for params in product_lookup.values():
        min_reach = product_reach_setting(params, "Min_Reach", min_reach_default)
        target_reach = product_reach_setting(params, "Target_Reach", target_reach_default)
        if min_reach > target_reach + 1e-9:
            invalid_rows.append(f"{params.get('Product', '')}: Min {min_reach:g} > Target {target_reach:g}")
    if invalid_rows:
        preview = "; ".join(invalid_rows[:8])
        more = " ..." if len(invalid_rows) > 8 else ""
        raise ValueError(f"Product reach settings must satisfy Min Reach <= Target Reach: {preview}{more}")


def min_reach_check_horizon(demand_by_product, pre_weeks=4):
    demand_weeks = []
    for values in demand_by_product.values():
        demand_weeks.extend(label for label, value in values.items() if float(value) > 0)
    return expand_week_horizon(demand_weeks, pre_weeks=pre_weeks, post_weeks=0)


def min_reach_violations_from_loading(loading_df, product_lookup, min_reach_default):
    week_cols = [col for col in loading_df.columns if is_internal_week_label(col)]
    reach_rows = loading_metric_rows_by_product(loading_df, "Reach Level")
    demand_rows = loading_metric_rows_by_product(loading_df, "Demand")
    violations = []

    for product_key, params in product_lookup.items():
        min_reach = product_reach_setting(params, "Min_Reach", min_reach_default)
        if min_reach <= 0 or product_key not in reach_rows or product_key not in demand_rows:
            continue
        row = reach_rows[product_key]
        demand_row = demand_rows[product_key]
        for week in week_cols:
            future_demand = []
            for offset in range(1, 5):
                demand_value = pd.to_numeric(demand_row.get(shift_week_label(week, offset)), errors="coerce")
                future_demand.append(0.0 if pd.isna(demand_value) else float(demand_value))
            if sum(future_demand) <= 0:
                continue
            reach_value = pd.to_numeric(row.get(week), errors="coerce")
            reach_value = 0.0 if pd.isna(reach_value) else float(reach_value)
            if reach_value + 1e-9 < min_reach:
                violations.append({
                    "Product": params["Product"],
                    "Week": week,
                    "Reach Level": reach_value,
                    "Min Reach Level": min_reach,
                    "Shortage Reach": min_reach - reach_value,
                })
    return pd.DataFrame(violations)


def balance_wafer_start_with_min_floor_for_tester(wafer_by_product, min_wafer_by_product, product_lookup, flow_df, tester_capacity, advance_weeks):
    offsets = build_flow_offsets(flow_df)
    sort_offset = int(offsets["sort"])
    capacity = float(tester_capacity)
    advance_weeks = max(0, int(advance_weeks))
    balanced = {product_key: {} for product_key in set(wafer_by_product) | set(min_wafer_by_product)}
    tester_usage = {}
    target_rows = []
    allocation_rows = []

    def week_number(label):
        year, week, _ = week_sort_key(label)
        return year * 52 + week

    def add_wafer(product_key, source_week, wafer):
        if wafer <= 1e-9:
            return
        product_values = balanced.setdefault(product_key, {})
        product_values[source_week] = product_values.get(source_week, 0.0) + wafer

    def schedule_task(product_key, original_source_week, wafer, reason, hard_required):
        params = product_lookup.get(product_key)
        if params is None or params["weekly_output"] <= 0 or wafer <= 1e-9:
            return 0.0, wafer / params["weekly_output"] if params is not None and params["weekly_output"] > 0 else 0.0

        weekly_output = float(params["weekly_output"])
        remaining_tester = float(wafer) / weekly_output
        required_tester = remaining_tester
        allocated_tester = 0.0
        candidate_source_weeks = [shift_week_label(original_source_week, -offset) for offset in range(0, advance_weeks + 1)]
        earliest_week = params.get("Earliest_Wafer_Start")

        for allocated_source_week in candidate_source_weeks:
            if remaining_tester <= 1e-9:
                break
            if earliest_week and week_sort_key(allocated_source_week) < week_sort_key(earliest_week):
                continue
            allocated_tester_week = shift_week_label(allocated_source_week, sort_offset)
            spare_tester = capacity - tester_usage.get(allocated_tester_week, 0.0)
            if spare_tester <= 1e-9:
                continue

            tester_to_place = min(remaining_tester, spare_tester)
            wafer_to_place = tester_to_place * weekly_output
            add_wafer(product_key, allocated_source_week, wafer_to_place)
            tester_usage[allocated_tester_week] = tester_usage.get(allocated_tester_week, 0.0) + tester_to_place
            remaining_tester -= tester_to_place
            allocated_tester += tester_to_place
            allocation_rows.append({
                "Product": params["Product"],
                "Reason": reason,
                "Required Wafer Start Week": original_source_week,
                "Allocated Wafer Start Week": allocated_source_week,
                "Required Testing Week": shift_week_label(original_source_week, sort_offset),
                "Allocated Testing Week": allocated_tester_week,
                "Wafer Allocated": wafer_to_place,
                "Tester Allocated": tester_to_place,
                "Required Tester": required_tester,
                "Advance Weeks": week_number(original_source_week) - week_number(allocated_source_week),
                "Tester Capacity": capacity,
                "Earliest Wafer Start Time": earliest_week or "",
                "Hard Required": bool(hard_required),
                "Status": "Allocated",
            })

        if remaining_tester > 1e-9:
            allocation_rows.append({
                "Product": params["Product"],
                "Reason": reason,
                "Required Wafer Start Week": original_source_week,
                "Allocated Wafer Start Week": "",
                "Required Testing Week": shift_week_label(original_source_week, sort_offset),
                "Allocated Testing Week": "",
                "Wafer Allocated": 0.0,
                "Tester Allocated": 0.0,
                "Required Tester": required_tester,
                "Advance Weeks": "",
                "Tester Capacity": capacity,
                "Earliest Wafer Start Time": earliest_week or "",
                "Hard Required": bool(hard_required),
                "Status": "Demand Coverage Shortage" if hard_required else "Reach Buffer Skipped",
                "Shortage Tester": remaining_tester,
                "Shortage Wafer": remaining_tester * weekly_output,
            })
        return allocated_tester, remaining_tester

    min_tasks = []
    target_tasks = []
    def task_sort_key(item):
        source_week = item[0]
        product_key = item[1]
        priority = float(product_lookup.get(product_key, {}).get("Priority", 0.0))
        return (week_sort_key(source_week), -priority, product_key)

    for product_key in sorted(set(wafer_by_product) | set(min_wafer_by_product)):
        source_weeks = sorted(set(wafer_by_product.get(product_key, {})) | set(min_wafer_by_product.get(product_key, {})), key=week_sort_key)
        for source_week in source_weeks:
            min_wafer = max(0.0, float(min_wafer_by_product.get(product_key, {}).get(source_week, 0.0)))
            target_wafer = max(0.0, float(wafer_by_product.get(product_key, {}).get(source_week, 0.0)))
            target_extra = max(0.0, target_wafer - min_wafer)
            if min_wafer > 1e-9:
                min_tasks.append((source_week, product_key, min_wafer))
            if target_extra > 1e-9:
                target_tasks.append((source_week, product_key, target_extra, target_wafer, min_wafer))

    for source_week, product_key, wafer in sorted(min_tasks, key=task_sort_key):
        schedule_task(product_key, source_week, wafer, "Demand Coverage", hard_required=True)

    hard_shortage_exists = any(str(row.get("Status", "")).endswith("Shortage") and bool(row.get("Hard Required")) for row in allocation_rows)
    if hard_shortage_exists:
        for source_week, product_key, wafer, target_wafer, min_wafer in sorted(target_tasks, key=task_sort_key):
            params = product_lookup.get(product_key)
            if params is None or params["weekly_output"] <= 0:
                continue
            shortage_tester = wafer / params["weekly_output"]
            allocation_rows.append({
                "Product": params["Product"],
                "Reason": "Reach Buffer",
                "Required Wafer Start Week": source_week,
                "Allocated Wafer Start Week": "",
                "Required Testing Week": shift_week_label(source_week, sort_offset),
                "Allocated Testing Week": "",
                "Wafer Allocated": 0.0,
                "Tester Allocated": 0.0,
                "Required Tester": shortage_tester,
                "Advance Weeks": "",
                "Tester Capacity": capacity,
                "Hard Required": False,
                "Status": "Reach Buffer Skipped",
                "Shortage Tester": shortage_tester,
                "Shortage Wafer": wafer,
            })
            target_rows.append({
                "Product": params["Product"],
                "Wafer Start Week": source_week,
                "Testing Week": shift_week_label(source_week, sort_offset),
                "Target Wafer Requested": target_wafer,
                "Min Required Wafer": min_wafer,
                "Target Extra Wafer Requested": wafer,
                "Target Extra Wafer Allocated": 0.0,
                "Target Extra Wafer Skipped": wafer,
                "Target Extra Tester Allocated": 0.0,
                "Target Extra Tester Skipped": shortage_tester,
                "Reason": "Reach buffer skipped because hard demand coverage is not feasible inside the max prebuild window",
            })
        for product_key in list(balanced.keys()):
            balanced[product_key] = {week: value for week, value in balanced[product_key].items() if abs(float(value)) > 1e-9}
        return balanced, pd.DataFrame(target_rows), pd.DataFrame(allocation_rows)

    for source_week, product_key, wafer, target_wafer, min_wafer in sorted(target_tasks, key=task_sort_key):
        allocated_tester, shortage_tester = schedule_task(product_key, source_week, wafer, "Reach Buffer", hard_required=False)
        params = product_lookup.get(product_key)
        if params is None or params["weekly_output"] <= 0:
            continue
        target_rows.append({
            "Product": params["Product"],
            "Wafer Start Week": source_week,
            "Testing Week": shift_week_label(source_week, sort_offset),
            "Target Wafer Requested": target_wafer,
            "Min Required Wafer": min_wafer,
            "Target Extra Wafer Requested": wafer,
            "Target Extra Wafer Allocated": allocated_tester * params["weekly_output"],
            "Target Extra Wafer Skipped": shortage_tester * params["weekly_output"],
            "Target Extra Tester Allocated": allocated_tester,
            "Target Extra Tester Skipped": shortage_tester,
            "Reason": "Reach buffer uses only tester capacity left after demand coverage latest scheduling",
        })

    for product_key in list(balanced.keys()):
        balanced[product_key] = {week: value for week, value in balanced[product_key].items() if abs(float(value)) > 1e-9}

    return balanced, pd.DataFrame(target_rows), pd.DataFrame(allocation_rows)


def style_display_table(output_df):
    numeric_cols = output_df.select_dtypes(include="number").columns.tolist()
    metric_colors = {
        "Bump": "background-color: #eef7f2",
        "Testing": "background-color: #eef3fb",
        "DPS": "background-color: #fff6e8",
    }

    def highlight_total(row):
        is_total = any(str(value) == "Grand Total" for value in row.values)
        if is_total:
            return ["font-weight: 700; background-color: #f2f4f7" for _ in row]

        metric = str(row.get("Metric", ""))
        row_color = metric_colors.get(metric, "")
        return [row_color for _ in row]

    def format_value(row, col):
        value = row[col]
        if value == "" or pd.isna(value):
            return ""
        if isinstance(value, bool):
            return "Yes" if value else "No"
        numeric_value = pd.to_numeric(value, errors="coerce")
        if pd.isna(numeric_value):
            return value

        context = " ".join(str(row.get(key, "")) for key in ["Metric", "Stage", "RFP", "Row Labels"]) + f" {col}"
        context = context.lower()
        if "reach" in context or "yield" in context or "tester" in context:
            return f"{float(numeric_value):,.2f}"
        return f"{float(numeric_value):,.0f}"

    display_df = output_df.copy()
    for col in numeric_cols:
        display_df[col] = display_df.apply(lambda row: format_value(row, col), axis=1)

    return (
        display_df.style
        .set_table_styles([
            {"selector": "th", "props": [("background-color", "#eef3f8"), ("color", "#17202a"), ("font-weight", "700"), ("border", "1px solid #d0d7de")]},
            {"selector": "td", "props": [("background-color", "#ffffff"), ("border", "1px solid #d0d7de")]},
        ])
        .apply(highlight_total, axis=1)
    )


def style_initial_loading_table(output_df):
    week_cols = week_columns(output_df)

    def format_value(value):
        if value == "" or pd.isna(value):
            return ""
        if isinstance(value, bool):
            return "Over" if value else "OK"
        numeric_value = pd.to_numeric(value, errors="coerce")
        if pd.isna(numeric_value):
            return value
        if abs(float(numeric_value) - round(float(numeric_value))) < 0.005:
            return f"{float(numeric_value):,.0f}"
        return f"{float(numeric_value):,.2f}"

    display_df = output_df.copy()
    for col in week_cols:
        display_df[col] = display_df[col].map(format_value)

    def highlight_row(row):
        if str(row.get("Stage", "")).strip():
            return ["font-weight: 700; background-color: #1f4e78; color: #ffffff" for _ in row]
        if str(row.get("RFP", "")).strip() == "Grand Total":
            return ["font-weight: 700; background-color: #d9eaf7" for _ in row]
        return ["background-color: #ffffff" for _ in row]

    def highlight_tester_cells(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        tester_over_rows = data["RFP"].astype(str).str.strip().eq("Tester Over Limit") if "RFP" in data.columns else pd.Series(False, index=data.index)
        remaining_rows = data["RFP"].astype(str).str.strip().eq("Remaining Tester") if "RFP" in data.columns else pd.Series(False, index=data.index)
        for col in week_cols:
            if col not in data.columns:
                continue
            over_values = data.loc[tester_over_rows, col].astype(str).eq("Over")
            styles.loc[tester_over_rows & over_values.reindex(data.index, fill_value=False), col] = "background-color: #f8d7da; color: #842029; font-weight: 700"
            remaining_values = pd.to_numeric(data.loc[remaining_rows, col], errors="coerce")
            styles.loc[remaining_rows & remaining_values.lt(0).reindex(data.index, fill_value=False), col] = "background-color: #f8d7da; color: #842029; font-weight: 700"
        return styles

    return (
        display_df.style
        .set_table_styles([
            {"selector": "th", "props": [("background-color", "#eef3f8"), ("color", "#17202a"), ("font-weight", "700"), ("border", "1px solid #d0d7de")]},
            {"selector": "td", "props": [("border", "1px solid #d0d7de")]},
        ])
        .apply(highlight_row, axis=1)
        .apply(highlight_tester_cells, axis=None)
    )


def make_initial_loading_bytes(initial_loading_table, initial_tester_summary_table):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        display_week_table(initial_loading_table).to_excel(writer, sheet_name="Initial_Loading_View", index=False)
        display_week_table(initial_tester_summary_table).to_excel(writer, sheet_name="Initial_Tester_Summary", index=False)
        planner.style_excel_sheet(writer.book["Initial_Loading_View"])
        planner.style_excel_sheet(writer.book["Initial_Tester_Summary"])
    output.seek(0)
    return output.getvalue()


def make_simple_plan_output_bytes(simple_outputs, loading_table, tester_summary_table, adjustment_df, min_reach_adjustment_df=None, min_reach_violation_df=None, target_buffer_reduction_df=None, phase_outputs=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        simple_outputs["wafer_start_table"].to_excel(writer, sheet_name="Wafer_Start", index=False)
        simple_outputs["prebuild_output_table"].to_excel(writer, sheet_name="Prebuild_Output", index=False)
        simple_outputs["build_output_table"].to_excel(writer, sheet_name="Build_Output", index=False)
        simple_outputs["tester_usage_table"].to_excel(writer, sheet_name="Tester_Allocation", index=False)
        display_week_table(loading_table).to_excel(writer, sheet_name="Balanced_Loading_View", index=False)
        display_week_table(tester_summary_table).to_excel(writer, sheet_name="Balanced_Tester_Summary", index=False)
        if min_reach_adjustment_df is not None and not min_reach_adjustment_df.empty:
            display_week_table(min_reach_adjustment_df).to_excel(writer, sheet_name="Min_Reach_Adjustments", index=False)
        if min_reach_violation_df is not None and not min_reach_violation_df.empty:
            display_week_table(min_reach_violation_df).to_excel(writer, sheet_name="Min_Reach_Violations", index=False)
        if adjustment_df is not None and not adjustment_df.empty:
            display_week_table(adjustment_df).to_excel(writer, sheet_name="Tester_Adjustments", index=False)

        for sheet_name in ["Wafer_Start", "Prebuild_Output", "Build_Output", "Tester_Allocation", "Balanced_Loading_View", "Balanced_Tester_Summary"]:
            planner.style_excel_sheet(writer.book[sheet_name])
        if min_reach_adjustment_df is not None and not min_reach_adjustment_df.empty:
            planner.style_excel_sheet(writer.book["Min_Reach_Adjustments"])
        if min_reach_violation_df is not None and not min_reach_violation_df.empty:
            planner.style_excel_sheet(writer.book["Min_Reach_Violations"])
        if adjustment_df is not None and not adjustment_df.empty:
            planner.style_excel_sheet(writer.book["Tester_Adjustments"])

        if phase_outputs:
            for sheet_index, phase_name in enumerate(["Prebuild Phase", "Demand Phase"]):
                phase_data = phase_outputs.get(phase_name, {})
                worksheet = writer.book.create_sheet(phase_name.replace(" ", "_"), sheet_index)
                start_row = 1
                for title, graph_df, chart_kind in phase_data.get("graphs", []):
                    if graph_df.empty:
                        continue
                    header_row, last_row, last_col = planner.add_table_block(worksheet, title, graph_df, start_row, freeze=False)
                    planner.add_excel_chart(worksheet, title, chart_kind, header_row, last_row, last_col, f"L{start_row}")
                    start_row += max(29, len(graph_df) + 5)
                for title, table_df in phase_data.get("tables", []):
                    if table_df.empty:
                        continue
                    planner.add_table_block(worksheet, title, table_df, start_row, freeze=False)
                    start_row += len(table_df) + 5
                worksheet.freeze_panes = None

        tester_graph_outputs = [item for item in simple_outputs["graph_outputs"] if item[2] in ["tester", "reach_tester"]]
        other_graph_outputs = [item for item in simple_outputs["graph_outputs"] if item[2] not in ["tester", "reach_tester"]]

        tester_graph_sheet = writer.book.create_sheet("Tester_Graph")
        for index, (title, graph_df, chart_kind) in enumerate(tester_graph_outputs):
            start_row = 1 + (index * 29)
            header_row, last_row, last_col = planner.add_table_block(tester_graph_sheet, title, graph_df, start_row, freeze=False)
            planner.add_excel_chart(tester_graph_sheet, title, chart_kind, header_row, last_row, last_col, f"L{start_row}")
        tester_graph_sheet.freeze_panes = None

        other_graph_sheet = writer.book.create_sheet("Other_Graphs")
        for index, (title, graph_df, chart_kind) in enumerate(other_graph_outputs):
            start_row = 1 + (index * 29)
            header_row, last_row, last_col = planner.add_table_block(other_graph_sheet, title, graph_df, start_row, freeze=False)
            planner.add_excel_chart(other_graph_sheet, title, chart_kind, header_row, last_row, last_col, f"L{start_row}")
        other_graph_sheet.freeze_panes = None
    output.seek(0)
    return output.getvalue()


def render_initial_loading_section(initial_loading_table, initial_tester_summary_table, initial_loading_source, tester_number):
    over_limit_weeks = int(initial_tester_summary_table["Tester Over Limit"].sum()) if not initial_tester_summary_table.empty else 0
    st.subheader("Initial Loading View")
    st.caption(
        f"Source: {initial_loading_source}. Initial loading calculates ideal production from demand plus Target Reach stock, then shifts it backward through flow. It does not optimize tester/prebuild. "
        "Before DPS, quantities are wafer counts. DPS, DC and stock are chip counts. Tester required can be fractional."
    )
    display_start_week, display_end_week = week_range_selector("Initial loading", initial_loading_table, "initial_loading_display")
    visible_initial_loading_table = filter_week_range(initial_loading_table, display_start_week, display_end_week)
    visible_initial_tester_summary_table = filter_week_range(initial_tester_summary_table, display_start_week, display_end_week)
    st.dataframe(style_initial_loading_table(display_week_table(visible_initial_loading_table)), width="stretch")

    if over_limit_weeks:
        worst_over = (initial_tester_summary_table["Total Tester Required"] - float(tester_number)).max()
        st.warning(f"Initial loading exceeds tester capacity in {over_limit_weeks} weeks. Worst weekly over limit: {worst_over:,.2f} testers.")
    else:
        st.success("Initial loading satisfies current tester capacity.")

    with st.expander("Initial Tester Summary"):
        st.dataframe(style_display_table(display_week_table(visible_initial_tester_summary_table)), width="stretch")


def draw_graphs(graph_outputs):
    st.subheader("Graph")
    for title, graph_df, chart_kind in graph_outputs:
        chart_df = graph_df[graph_df["Row Labels"].astype(str) != "Grand Total"].copy()
        x_order = chart_df["Row Labels"].astype(str).tolist()

        if chart_kind == "tester":
            product_cols = [col for col in chart_df.columns if col not in ["Row Labels", "Grand Total", "Tester Capacity"]]
            tester_long_df = chart_df.melt(id_vars="Row Labels", value_vars=product_cols, var_name="Product", value_name="Tester Used")
            tester_long_df = tester_long_df[tester_long_df["Tester Used"] > 0]
            chart = px.bar(tester_long_df, x="Row Labels", y="Tester Used", color="Product", title=title)
            chart.update_layout(barmode="stack")
            chart.add_scatter(
                x=chart_df["Row Labels"],
                y=chart_df["Tester Capacity"],
                mode="lines+markers",
                name="Tester Capacity",
                line={"color": "#d62728", "width": 3},
                marker={"size": 8},
            )
            chart.update_layout(xaxis_title="Week", yaxis_title="Tester")
            chart.update_xaxes(type="category", categoryorder="array", categoryarray=x_order)
            st.plotly_chart(chart, width="stretch")
            continue

        if chart_kind == "stock":
            target_line = chart_df[["Row Labels", "Target Stock"]].copy()
            product_cols = [col for col in chart_df.columns if col not in ["Row Labels", "Grand Total", "Target Stock"]]
            chart_df = chart_df.melt(id_vars="Row Labels", value_vars=product_cols, var_name="Product", value_name=title)
            chart = px.bar(chart_df, x="Row Labels", y=title, color="Product", title=title)
            chart.update_layout(barmode="stack")
            chart.add_scatter(
                x=target_line["Row Labels"],
                y=target_line["Target Stock"],
                mode="lines+markers",
                name="Target Stock",
                line={"color": "#d62728", "width": 3},
                marker={"size": 8},
            )
            chart.update_layout(xaxis_title="Week", yaxis_title="Stock / Target Stock")
            chart.update_xaxes(type="category", categoryorder="array", categoryarray=x_order)
            st.plotly_chart(chart, width="stretch")
            continue

        if chart_kind == "reach_tester":
            reach_cols = [col for col in chart_df.columns if str(col).startswith("Target Reach = ")]
            chart_long_df = chart_df.melt(id_vars="Row Labels", value_vars=reach_cols, var_name="Scenario", value_name="Tester Required")
            chart = px.line(chart_long_df, x="Row Labels", y="Tester Required", color="Scenario", markers=True, title=title)
            chart.add_scatter(
                x=chart_df["Row Labels"],
                y=chart_df["Tester Capacity"],
                mode="lines",
                name="Tester Capacity",
                line={"color": "#d62728", "width": 4, "dash": "dash"},
            )
            chart.update_layout(xaxis_title="Week", yaxis_title="Tester Required")
            chart.update_xaxes(type="category", categoryorder="array", categoryarray=x_order)
            st.plotly_chart(chart, width="stretch")
            continue

        chart_df = chart_df.melt(id_vars="Row Labels", var_name="Product", value_name=title)
        chart_df = chart_df[chart_df["Product"] != "Grand Total"]

        if chart_kind == "line":
            chart = px.line(chart_df, x="Row Labels", y=title, color="Product", markers=True, title=title)
            if title == "Reach Level":
                chart.update_yaxes(range=[0, 6], dtick=1)
        else:
            chart = px.bar(chart_df, x="Row Labels", y=title, color="Product", title=title)
            chart.update_layout(barmode="stack")

        chart.update_layout(xaxis_title="Week", yaxis_title=title)
        chart.update_xaxes(type="category", categoryorder="array", categoryarray=x_order)
        st.plotly_chart(chart, width="stretch")


uploaded = st.sidebar.file_uploader("Upload input Excel", type=["xlsx"])
file_bytes = uploaded.getvalue() if uploaded else None

try:
    tables = load_tables(file_bytes)
except Exception as exc:
    st.error(str(exc))
    st.stop()

if tables is None:
    st.info("Please upload an input Excel file from the sidebar, or add input_sample.xlsx to the repo.")
    st.stop()

tables["product"] = ensure_priority_column(tables["product"])
wspw_wafer_start_source_df = repair_week_sequence_columns(load_wspw_wafer_start(file_bytes))

demand_source_key = uploaded.name if uploaded else "default_input"
if st.session_state.get("simple_demand_source") != demand_source_key:
    st.session_state["simple_demand_source"] = demand_source_key
    st.session_state["simple_demand_df"] = tables["demand"].copy()
    st.session_state["simple_demand_version"] = 0
elif "simple_demand_df" not in st.session_state:
    st.session_state["simple_demand_df"] = tables["demand"].copy()
    st.session_state["simple_demand_version"] = 0

min_reach_default = float(target_value(tables["target"], "Min Reach Level", planner.DEFAULT_TESTER_CONFIG["min_REACH"]))
target_reach_default = float(target_value(tables["target"], "Target Reach Level", planner.DEFAULT_TESTER_CONFIG["target_REACH"]))
tester_number_default = int(target_value(tables["target"], "Tester Number", planner.DEFAULT_TESTER_CONFIG["available"]))
tester_smoothing_default = int(target_value(tables["target"], "Tester Smoothing Weeks", planner.DEFAULT_TESTER_CONFIG["tester_smoothing_weeks"]))
wafer_start_before_default = target_value(tables["target"], "Wafer Start Time", None)
wafer_start_before_default = "" if pd.isna(wafer_start_before_default) else str(int(wafer_start_before_default))

with st.sidebar:
    st.header("Planning Inputs")
    min_reach = st.number_input("Min Reach Level", min_value=0.0, value=min_reach_default, step=0.5)
    target_reach = st.number_input("Target Reach Level", min_value=0.0, value=target_reach_default, step=0.5)
    tester_number = st.number_input("Tester Number", min_value=1, value=tester_number_default, step=1)
    tester_smoothing_weeks = st.number_input(
        "Tester Smoothing Weeks",
        min_value=0,
        value=tester_smoothing_default,
        step=1,
        help="Allows future tester load to be pulled earlier into idle tester weeks. Set to 0 to disable smoothing.",
    )
    wafer_start_before_text = st.text_input(
        "Wafer Start Time",
        value=wafer_start_before_default,
        help="Optional. Enter how many weeks before demand wafer start is allowed. Leave blank for no limit.",
    )

tabs = st.tabs(["Flow", "Product", "Demand", "WSPW Wafer Start", "Inventory"])
with tabs[0]:
    st.caption("Transit time is shown as a column on each stage. When you run the plan, it will be converted back to Transist Time rows automatically.")
    flow_editor_df = flow_to_editor_table(tables["flow"])
    flow_headers = render_header_editor("simple_flow", flow_editor_df.columns, demand_source_key)
    flow_df = st.data_editor(
        flow_editor_df,
        width="stretch",
        num_rows="dynamic",
        key="simple_flow",
        column_config={
            "Stage": st.column_config.TextColumn(flow_headers.get("Stage", "Stage"), help="Main process step, excluding transit rows."),
            "Cycle Time Week": st.column_config.NumberColumn(flow_headers.get("Cycle Time Week", "Cycle Time Week"), min_value=0.0, step=1.0, format="%.0f"),
            "Transit Week": st.column_config.NumberColumn(flow_headers.get("Transit Week", "Transit Week"), min_value=0.0, step=1.0, format="%.0f"),
            "Yield": st.column_config.NumberColumn(flow_headers.get("Yield", "Yield"), min_value=0.0, max_value=1.0, step=0.001, format="%.3f"),
        },
    )
with tabs[1]:
    product_headers = render_header_editor("simple_product", tables["product"].columns, demand_source_key)
    product_column_config = generic_column_config(tables["product"].columns, product_headers)
    if "Priority" in tables["product"].columns:
        product_column_config["Priority"] = st.column_config.NumberColumn(
            product_headers.get("Priority", "Priority"),
            help="Higher number means higher priority. When weekly total tester capacity is not enough, products with higher Priority are allocated first.",
            step=1,
            format="%.0f",
        )
    for reach_column in ["Min Reach Level", "Target Reach Level"]:
        if reach_column in tables["product"].columns:
            product_column_config[reach_column] = st.column_config.NumberColumn(
                product_headers.get(reach_column, reach_column),
                min_value=0.0,
                step=0.5,
                format="%.1f",
            )
    if "Earliest Wafer Start Time" in tables["product"].columns:
        product_column_config["Earliest Wafer Start Time"] = st.column_config.TextColumn(
            product_headers.get("Earliest Wafer Start Time", "Earliest Wafer Start Time"),
            help="Hard constraint per product. Use YYYY-CWNN, YYMon CWNN, or YYYY-MM. Blank means no earliest wafer start limit.",
        )
    if "Total Yield" in tables["product"].columns:
        product_column_config["Total Yield"] = st.column_config.NumberColumn(
            product_headers.get("Total Yield", "Total Yield"),
            min_value=0.0,
            max_value=1.0,
            step=0.001,
            format="%.3f",
        )
    product_df = st.data_editor(
        tables["product"],
        width="stretch",
        num_rows="dynamic",
        key="simple_product",
        column_config=product_column_config,
    )
with tabs[2]:
    st.caption("Choose the demand date range. Existing month values are kept; new months are added with 0 demand.")
    default_start, default_end = demand_date_range_defaults(st.session_state["simple_demand_df"])
    range_col1, range_col2, range_col3 = st.columns([1, 1, 1])
    with range_col1:
        start_month = st.date_input("Start month", value=default_start, key="simple_demand_start_month")
    with range_col2:
        end_month = st.date_input("End month", value=default_end, key="simple_demand_end_month")
    with range_col3:
        st.write("")
        st.write("")
        apply_range = st.button("Apply date range", key="apply_demand_date_range")

    if apply_range:
        try:
            st.session_state["simple_demand_df"] = apply_demand_date_range(
                st.session_state["simple_demand_df"],
                start_month,
                end_month,
            )
            st.session_state["simple_demand_version"] += 1
            st.success("Demand date range updated")
        except Exception as exc:
            st.error(str(exc))

    demand_df = st.data_editor(
        st.session_state["simple_demand_df"],
        width="stretch",
        num_rows="dynamic",
        key=f"simple_demand_{st.session_state['simple_demand_version']}",
        column_config=generic_column_config(
            st.session_state["simple_demand_df"].columns,
            render_header_editor(
                "simple_demand",
                st.session_state["simple_demand_df"].columns,
                f"{demand_source_key}_{st.session_state['simple_demand_version']}",
            )
        ),
    )
    st.session_state["simple_demand_df"] = demand_df.copy()
with tabs[3]:
    st.caption("Reads the WSPW block from the uploaded Excel as wafer start by unified calendar week. January first week is labeled CW02, then weeks continue sequentially.")
    if wspw_wafer_start_source_df.empty:
        st.info("No WSPW block found in the current Excel. You can use this blank table as a placeholder.")
    wspw_start_week, wspw_end_week = week_range_selector("WSPW", wspw_wafer_start_source_df, "wspw_display")
    visible_wspw_wafer_start_source_df = repair_week_sequence_columns(filter_week_range(wspw_wafer_start_source_df, wspw_start_week, wspw_end_week))
    wspw_week_columns = [col for col in visible_wspw_wafer_start_source_df.columns if is_internal_week_label(col)]
    wspw_column_config = {
        "Basic Type": st.column_config.TextColumn("Basic Type"),
        "RFP": st.column_config.TextColumn("RFP"),
    }
    for week_column in wspw_week_columns:
        wspw_column_config[week_column] = st.column_config.NumberColumn(week_display_label(week_column), min_value=0.0, step=1.0, format="%.0f")
    edited_wspw_wafer_start_df = st.data_editor(
        visible_wspw_wafer_start_source_df,
        width="stretch",
        num_rows="dynamic",
        key=f"simple_wspw_wafer_start_{demand_source_key}_{wspw_start_week}_{wspw_end_week}",
        column_config=wspw_column_config,
    )
    if len(edited_wspw_wafer_start_df) == len(wspw_wafer_start_source_df):
        wspw_wafer_start_df = wspw_wafer_start_source_df.copy()
        for col in visible_wspw_wafer_start_source_df.columns:
            wspw_wafer_start_df[col] = edited_wspw_wafer_start_df[col].tolist()
    else:
        wspw_wafer_start_df = edited_wspw_wafer_start_df.copy()
with tabs[4]:
    inventory_headers = render_header_editor("simple_inventory", tables["inventory"].columns, demand_source_key)
    inventory_df = st.data_editor(
        tables["inventory"],
        width="stretch",
        num_rows="dynamic",
        key="simple_inventory",
        column_config=generic_column_config(tables["inventory"].columns, inventory_headers),
    )

try:
    edited_wafer_start_before = parse_wafer_start_before(wafer_start_before_text)
    edited_input_bytes = make_input_bytes(
        flow_df,
        product_df,
        demand_df,
        inventory_df,
        min_reach,
        target_reach,
        tester_number,
        tester_smoothing_weeks,
        edited_wafer_start_before,
    )
    st.download_button(
        "Download Edited Input Excel",
        data=edited_input_bytes,
        file_name="edited_input.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception as exc:
    st.warning(f"Cannot prepare edited input download: {exc}")

if st.button("Generate Initial Loading View", type="secondary"):
    try:
        initial_loading_table, initial_tester_summary_table, initial_loading_source = build_initial_loading_view(
            wspw_wafer_start_df,
            flow_df,
            product_df,
            demand_df,
            inventory_df,
            target_reach,
            tester_number,
        )
        st.session_state["initial_loading_result"] = {
            "loading": initial_loading_table,
            "tester_summary": initial_tester_summary_table,
            "source": initial_loading_source,
            "tester_number": tester_number,
        }
    except Exception as exc:
        st.error(str(exc))

initial_loading_result = st.session_state.get("initial_loading_result")
if initial_loading_result is not None:
    initial_loading_table = initial_loading_result["loading"]
    initial_tester_summary_table = initial_loading_result["tester_summary"]
    initial_loading_source = initial_loading_result["source"]
    render_initial_loading_section(
        initial_loading_table,
        initial_tester_summary_table,
        initial_loading_source,
        initial_loading_result["tester_number"],
    )
    initial_output_bytes = make_initial_loading_bytes(initial_loading_table, initial_tester_summary_table)
    st.download_button(
        "Download Initial Loading Excel",
        data=initial_output_bytes,
        file_name="initial_loading_view.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if bool(initial_tester_summary_table["Tester Over Limit"].any()) if not initial_tester_summary_table.empty else False:
        st.info("Initial loading exceeds tester capacity. You can run the simple plan to rebalance with current constraints.")
    else:
        st.info("Tester capacity is enough for the initial loading view. Simple plan is not required unless you want to compare results.")

if st.button("Run simple plan", type="primary"):
    try:
        wafer_start_before = parse_wafer_start_before(wafer_start_before_text)
        if min_reach > target_reach:
            raise ValueError("Reach levels must satisfy: Min Reach <= Target Reach.")
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    with st.spinner("Balancing tester loading by week..."):
        try:
            wafer_by_product, product_lookup, inventory_lookup, demand_by_product, simple_plan_source = build_initial_wafer_start_map(
                wspw_wafer_start_df,
                flow_df,
                product_df,
                demand_df,
                inventory_df,
                target_reach,
            )
            validate_product_reach_settings(product_lookup, min_reach, target_reach)
            demand_week_labels = sorted({week for values in demand_by_product.values() for week in values}, key=week_sort_key)
            if not wafer_by_product:
                raise ValueError("No wafer start could be generated from WSPW or Target Reach.")
            validate_earliest_wafer_start(wafer_by_product, product_lookup, "Initial target wafer start")

            wspw_cap_reduction_df = pd.DataFrame()
            if simple_plan_source == "WSPW Wafer Start":
                target_cap_wafer_by_product = build_initial_target_reach_wafer_start(
                    demand_by_product,
                    demand_week_labels,
                    product_lookup,
                    inventory_lookup,
                    flow_df,
                    target_reach,
                )
                wafer_by_product, wspw_cap_reduction_df = cap_wafer_start_to_target(
                    wafer_by_product,
                    target_cap_wafer_by_product,
                    product_lookup,
                )
                if not wafer_by_product or all(not values for values in wafer_by_product.values()):
                    raise ValueError("WSPW wafer start is entirely above the Target Reach ideal timing/quantity. No wafer start remains after applying the Target Reach cap.")

            target_buffer_reduction_df = wspw_cap_reduction_df
            min_reach_adjustment_df = pd.DataFrame()
            min_floor_wafer_by_product = build_min_reach_floor_wafer_start(
                demand_by_product,
                demand_week_labels,
                product_lookup,
                inventory_lookup,
                flow_df,
                min_reach,
            )
            validate_earliest_wafer_start(min_floor_wafer_by_product, product_lookup, "Min Reach floor wafer start")
            balanced_wafer_by_product, target_buffer_balance_df, adjustment_df = balance_wafer_start_with_min_floor_for_tester(
                wafer_by_product,
                min_floor_wafer_by_product,
                product_lookup,
                flow_df,
                tester_number,
                tester_smoothing_weeks,
            )
            if not target_buffer_balance_df.empty:
                target_buffer_reduction_df = pd.concat([target_buffer_reduction_df, target_buffer_balance_df], ignore_index=True)
            validate_earliest_wafer_start(balanced_wafer_by_product, product_lookup, "Balanced wafer start")
            if not adjustment_df.empty and {"Hard Required", "Status"}.issubset(adjustment_df.columns):
                hard_shortages = adjustment_df[
                    adjustment_df["Hard Required"].astype(bool)
                    & adjustment_df["Status"].astype(str).str.contains("Shortage", na=False)
                ]
                if not hard_shortages.empty:
                    preview = ", ".join(
                        f"{row['Product']} {week_display_label(row['Required Wafer Start Week'])}: {float(row.get('Shortage Tester', 0.0)):,.2f} tester"
                        for _, row in hard_shortages.head(8).iterrows()
                    )
                    more = " ..." if len(hard_shortages) > 8 else ""
                    raise ValueError(f"Hard Min Reach/Demand floor cannot fit within tester capacity and product Earliest Wafer Start Time: {preview}{more}")
            min_reach_horizon = min_reach_check_horizon(demand_by_product, pre_weeks=4)
            balanced_loading_table, balanced_tester_summary_table = build_loading_view_from_wafer_start(
                balanced_wafer_by_product,
                product_lookup,
                inventory_lookup,
                demand_by_product,
                flow_df,
                tester_number,
                min_reach_horizon,
            )
            tester_violations = tester_violations_from_summary(balanced_tester_summary_table)
            min_reach_violation_df = min_reach_violations_from_loading(balanced_loading_table, product_lookup, min_reach)
            simple_outputs = build_simple_weekly_outputs(
                balanced_wafer_by_product,
                product_lookup,
                inventory_lookup,
                demand_by_product,
                flow_df,
                tester_number,
                target_reach,
            )
            reach_tester_sensitivity_df = build_reach_tester_sensitivity_graph(
                product_lookup,
                inventory_lookup,
                demand_by_product,
                flow_df,
                tester_number,
            )
            simple_outputs["graph_outputs"].insert(1, ("Reach Level 1-6 Tester Need", reach_tester_sensitivity_df, "reach_tester"))
            phase_outputs = split_simple_outputs_by_phase(
                simple_outputs,
                balanced_loading_table,
                balanced_tester_summary_table,
                demand_by_product,
            )
            min_advance_weeks = 0
            if not adjustment_df.empty and "Advance Weeks" in adjustment_df.columns:
                min_advance_weeks = int(pd.to_numeric(adjustment_df["Advance Weeks"], errors="coerce").fillna(0).max())
            st.session_state["simple_plan_payload"] = {
                "source": simple_plan_source,
                "loading": balanced_loading_table,
                "tester_summary": balanced_tester_summary_table,
                "min_reach_adjustments": min_reach_adjustment_df,
                "target_buffer_reductions": target_buffer_reduction_df,
                "min_reach_violations": min_reach_violation_df,
                "adjustments": adjustment_df,
                "violations": tester_violations,
                "output_bytes": None,
                "simple_outputs": simple_outputs,
                "graph_outputs": simple_outputs["graph_outputs"],
                "phase_outputs": phase_outputs,
                "wafer_start_table": simple_outputs["wafer_start_table"],
                "prebuild_output_table": simple_outputs["prebuild_output_table"],
                "build_output_table": simple_outputs["build_output_table"],
                "stage_output_table": simple_outputs["stage_output_table"],
                "tester_usage_table": simple_outputs["tester_usage_table"],
                "min_advance_weeks": min_advance_weeks,
            }
        except Exception as exc:
            st.session_state.pop("simple_plan_payload", None)
            st.error(str(exc))
            st.stop()

simple_plan_payload = st.session_state.get("simple_plan_payload")
if simple_plan_payload is not None:
    st.subheader("Tester Balanced Loading View")
    st.caption(
        f"Source: {simple_plan_payload['source']}. Run simple plan now schedules Min Reach first at the latest feasible tester week, within Tester Smoothing Weeks as the max prebuild window. "
        "Target Reach is added only with remaining tester capacity. Before DPS, quantities are wafer counts. DPS, DC and stock are chip counts."
    )
    violations = simple_plan_payload["violations"]
    min_reach_violations = simple_plan_payload["min_reach_violations"]
    if violations:
        preview = ", ".join(f"{week_display_label(week)}: {value:.2f}" for week, value in violations[:8])
        more = " ..." if len(violations) > 8 else ""
        st.error(
            "Tester balancing is still over capacity inside the available pull-in window. "
            f"Over capacity weeks: {preview}{more}"
        )
        st.info("Increase Tester Smoothing Weeks or provide an earlier wafer start window, then run again.")
    if not min_reach_violations.empty:
        preview_rows = min_reach_violations.head(8)
        preview = ", ".join(
            f"{row['Product']} {week_display_label(row['Week'])}: {row['Reach Level']:.2f}/{row['Min Reach Level']:.2f}"
            for _, row in preview_rows.iterrows()
        )
        more = " ..." if len(min_reach_violations) > 8 else ""
        st.error(f"Min Reach is still not satisfied after balancing: {preview}{more}")
        st.info("Tester capacity is still respected. Increase Tester Smoothing Weeks, Tester Number, or lower Min Reach to make the Min Reach requirement feasible.")
    else:
        if not violations:
            st.success(
                "Plan completed. Every week is within tester capacity and product Min Reach is satisfied. "
                f"Minimum advance used: {simple_plan_payload['min_advance_weeks']} weeks."
            )

    st.subheader("Weekly Output")
    phase_outputs = simple_plan_payload.get("phase_outputs", {})
    phase_tabs = st.tabs(["Prebuild Phase", "Demand Phase"])
    for phase_tab, phase_name in zip(phase_tabs, ["Prebuild Phase", "Demand Phase"]):
        with phase_tab:
            phase_data = phase_outputs.get(phase_name, {"graphs": [], "tables": []})
            if phase_data.get("graphs"):
                draw_graphs(phase_data["graphs"])
            else:
                st.info("No graph data in this phase.")
            for table_title, table_df in phase_data.get("tables", []):
                if table_df.empty:
                    continue
                with st.expander(table_title, expanded=table_title in ["Tester Allocation", "Wafer Start", "Bump / Sort", "Balanced Loading View"]):
                    if table_title == "Balanced Loading View":
                        st.dataframe(style_initial_loading_table(table_df), width="stretch")
                    else:
                        st.dataframe(style_display_table(table_df), width="stretch")
    if not simple_plan_payload["min_reach_adjustments"].empty:
        with st.expander("Min Reach Added Wafer Start"):
            st.dataframe(style_display_table(display_week_table(simple_plan_payload["min_reach_adjustments"])), width="stretch")
    if not simple_plan_payload["min_reach_violations"].empty:
        with st.expander("Min Reach Violations"):
            st.dataframe(style_display_table(display_week_table(simple_plan_payload["min_reach_violations"])), width="stretch")
    if not simple_plan_payload["adjustments"].empty:
        status_values = simple_plan_payload["adjustments"]["Status"] if "Status" in simple_plan_payload["adjustments"].columns else pd.Series("", index=simple_plan_payload["adjustments"].index)
        shortage_rows = simple_plan_payload["adjustments"][status_values.astype(str).str.contains("Shortage", na=False)]
        if not shortage_rows.empty:
            st.warning(f"Tester hard constraint was respected, but {len(shortage_rows)} Min Reach scheduling rows could not fit inside the max prebuild window.")
        with st.expander("Tester Scheduling Check"):
            st.dataframe(style_display_table(display_week_table(simple_plan_payload["adjustments"])), width="stretch")
    if simple_plan_payload.get("output_bytes") is None:
        if st.button("Prepare Simple Plan Excel", type="secondary"):
            with st.spinner("Preparing Excel download..."):
                simple_plan_payload["output_bytes"] = make_simple_plan_output_bytes(
                    simple_plan_payload["simple_outputs"],
                    simple_plan_payload["loading"],
                    simple_plan_payload["tester_summary"],
                    simple_plan_payload["adjustments"],
                    simple_plan_payload["min_reach_adjustments"],
                    simple_plan_payload["min_reach_violations"],
                    simple_plan_payload["target_buffer_reductions"],
                    simple_plan_payload["phase_outputs"],
                )
                st.session_state["simple_plan_payload"] = simple_plan_payload
                st.rerun()
    else:
        st.download_button(
            "Download Simple Plan Excel",
            data=simple_plan_payload["output_bytes"],
            file_name="simple_plan_weekly_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
