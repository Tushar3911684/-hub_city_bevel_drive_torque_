from io import BytesIO
from pathlib import Path
import re

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).resolve().parent
DATABASE_FILE = APP_DIR / "HubCity_Bevel_Gear_Torque_Data.xlsx"
IN_LB_TO_NM = 0.112984829

st.set_page_config(
    page_title="Hub City Bevel Drive Torque Intelligence",
    page_icon="⚙️",
    layout="wide",
)
st.title("⚙️ Hub City Bevel Drive Torque Intelligence")
st.caption("Intelligent bevel drive torque lookup, comparison, and Excel report generation platform.")


def normalized_name(value):
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def find_column(dataframe, possible_names):
    available = {normalized_name(column): column for column in dataframe.columns}
    for name in possible_names:
        if normalized_name(name) in available:
            return available[normalized_name(name)]
    return None


def normalize_ratio(value):
    return str(value).strip().upper().replace(" ", "").replace("/", ":")


def display_ratio(value):
    return normalize_ratio(value).replace(":", "/")


def format_number(value):
    number = float(value)
    if number.is_integer():
        return f"{int(number):,}"
    return f"{number:,.2f}".rstrip("0").rstrip(".")


def ratio_sort_key(value):
    text = normalize_ratio(value)
    match = re.match(r"^(\d+(?:\.\d+)?):(\d+(?:\.\d+)?)", text)
    if not match:
        return 99, text
    numerator = float(match.group(1))
    denominator = float(match.group(2))
    direction = 0 if numerator >= denominator else 1
    return direction, numerator / denominator, text


@st.cache_data
def load_database():
    if not DATABASE_FILE.exists():
        raise FileNotFoundError(
            "HubCity_Bevel_Gear_Torque_Data.xlsx is missing. "
            "Keep it in the same folder as this Python file."
        )

    excel_file = pd.ExcelFile(DATABASE_FILE, engine="openpyxl")
    preferred_sheets = [
        "Torque_Range",
        "Torque_Ranges",
        "Torque Range",
        "Torque Ranges",
    ]
    sheet_name = next(
        (name for name in preferred_sheets if name in excel_file.sheet_names),
        None,
    )
    if sheet_name is None:
        raise ValueError("Torque_Range worksheet was not found in the database.")

    source = pd.read_excel(
        DATABASE_FILE,
        sheet_name=sheet_name,
        engine="openpyxl",
    )

    series_column = find_column(source, ["Series", "Model", "Gearbox Series"])
    ratio_column = find_column(source, ["Ratio", "Gear Ratio"])
    type_column = find_column(source, ["Gear_Type", "Gear Type", "Type"])
    min_column = find_column(
        source,
        ["Min_Torque_in_lb", "Minimum Torque", "Min Torque"],
    )
    max_column = find_column(
        source,
        ["Max_Torque_in_lb", "Maximum Torque", "Max Torque"],
    )
    direction_column = find_column(source, ["Direction"])

    required = {
        "Series": series_column,
        "Ratio": ratio_column,
        "Gear Type": type_column,
        "Minimum Torque": min_column,
        "Maximum Torque": max_column,
    }
    missing = [name for name, column in required.items() if column is None]
    if missing:
        raise ValueError("Missing database columns: " + ", ".join(missing))

    database = pd.DataFrame(
        {
            "Series": source[series_column].astype(str).str.strip().str.upper(),
            "Ratio": source[ratio_column].apply(normalize_ratio),
            "Gear_Type": source[type_column].astype(str).str.strip().str.upper(),
            "Min_Torque": pd.to_numeric(source[min_column], errors="coerce"),
            "Max_Torque": pd.to_numeric(source[max_column], errors="coerce"),
        }
    )

    if direction_column:
        database["Direction"] = source[direction_column].astype(str).str.strip()
    else:
        database["Direction"] = ""

    database["Gear_Type"] = database["Gear_Type"].replace(
        {
            "STRAIGHT": "ST",
            "STRAIGHT BEVEL": "ST",
            "SPIRAL": "SP",
            "SPIRAL BEVEL": "SP",
        }
    )
    database = database[database["Gear_Type"].isin(["ST", "SP"])]
    database = database.dropna(subset=["Min_Torque", "Max_Torque"])
    database = database.drop_duplicates(
        ["Series", "Ratio", "Gear_Type"],
        keep="first",
    )
    return database


def create_excel_report(database, selected_series):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Torque Matrix"
    worksheet.sheet_view.showGridLines = False

    navy = "1F4E78"
    cyan = "5BC0DE"
    light_blue = "D9EAF7"
    unavailable_red = "F4CCCC"
    white = "FFFFFF"
    thin = Side(style="thin", color="404040")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    current_row = 1

    for gear_type, title in [
        ("ST", "Straight Bevel Torque Range (in-lb)"),
        ("SP", "Spiral Bevel Torque Range (in-lb)"),
    ]:
        filtered = database[
            database["Series"].isin(selected_series)
            & (database["Gear_Type"] == gear_type)
        ]
        ratios = sorted(filtered["Ratio"].unique(), key=ratio_sort_key)
        values = {
            (row["Series"], row["Ratio"]): (
                f"{format_number(row['Min_Torque'])} - "
                f"{format_number(row['Max_Torque'])}"
            )
            for _, row in filtered.iterrows()
        }

        last_column = max(2, len(ratios) + 1)
        worksheet.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=last_column,
        )
        title_cell = worksheet.cell(current_row, 1, title)
        title_cell.fill = PatternFill("solid", fgColor=navy)
        title_cell.font = Font(color=white, bold=True, size=13)
        title_cell.alignment = center

        header_row = current_row + 1
        worksheet.cell(header_row, 1, "Series")
        for column_number, ratio in enumerate(ratios, start=2):
            worksheet.cell(
                header_row,
                column_number,
                f"{display_ratio(ratio)} {gear_type}",
            )

        for cell in worksheet[header_row][:last_column]:
            cell.fill = PatternFill("solid", fgColor=cyan)
            cell.font = Font(bold=True)
            cell.alignment = center
            cell.border = border

        first_data_row = header_row + 1
        for row_number, series in enumerate(selected_series, start=first_data_row):
            series_cell = worksheet.cell(row_number, 1, series)
            series_cell.fill = PatternFill("solid", fgColor=cyan)
            series_cell.font = Font(bold=True)
            series_cell.alignment = center
            series_cell.border = border

            for column_number, ratio in enumerate(ratios, start=2):
                value = values.get((series, ratio), "Not Available")
                cell = worksheet.cell(row_number, column_number, value)
                cell.fill = PatternFill(
                    "solid",
                    fgColor=(
                        light_blue if value != "Not Available" else unavailable_red
                    ),
                )
                cell.alignment = center
                cell.border = border

        current_row = first_data_row + len(selected_series) + 2

    worksheet.column_dimensions["A"].width = 18
    for column_number in range(2, worksheet.max_column + 1):
        worksheet.column_dimensions[
            get_column_letter(column_number)
        ].width = 18

    worksheet.freeze_panes = "B3"
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True

    details = workbook.create_sheet("Selected Data")
    details.append(
        [
            "Series",
            "Ratio",
            "Gear Type",
            "Direction",
            "Minimum Torque (in-lb)",
            "Maximum Torque (in-lb)",
        ]
    )

    selected_data = database[database["Series"].isin(selected_series)].sort_values(
        ["Gear_Type", "Series", "Ratio"]
    )
    for _, row in selected_data.iterrows():
        details.append(
            [
                row["Series"],
                display_ratio(row["Ratio"]),
                row["Gear_Type"],
                row["Direction"],
                row["Min_Torque"],
                row["Max_Torque"],
            ]
        )

    for cell in details[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(color=white, bold=True)
        cell.alignment = center

    details.freeze_panes = "A2"
    details.sheet_view.showGridLines = False
    for column_number, width in enumerate([18, 15, 15, 15, 24, 24], start=1):
        details.column_dimensions[get_column_letter(column_number)].width = width

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()


try:
    database = load_database()
except Exception as error:
    st.error(str(error))
    st.stop()

finder_tab, report_tab = st.tabs(
    ["🔎 Torque Finder", "📊 Excel Report Generator"]
)

with finder_tab:
    st.subheader("Find Torque Range")

    selected_series = st.selectbox(
        "Series",
        sorted(database["Series"].unique()),
        key="finder_series",
    )
    series_data = database[database["Series"] == selected_series]

    selected_ratio = st.selectbox(
        "Ratio",
        sorted(series_data["Ratio"].unique(), key=ratio_sort_key),
        format_func=display_ratio,
        key="finder_ratio",
    )
    ratio_data = series_data[series_data["Ratio"] == selected_ratio]

    selected_type = st.selectbox(
        "Gear Type",
        sorted(ratio_data["Gear_Type"].unique()),
        format_func=lambda value: (
            "ST - Straight Bevel" if value == "ST" else "SP - Spiral Bevel"
        ),
        key="finder_type",
    )

    result = ratio_data[ratio_data["Gear_Type"] == selected_type]

    if st.button("Find Torque", type="primary"):
        if result.empty:
            st.warning("No matching catalogue data was found.")
        else:
            record = result.iloc[0]
            minimum = float(record["Min_Torque"])
            maximum = float(record["Max_Torque"])

            st.success(
                "Available torque range: "
                f"{format_number(minimum)} to {format_number(maximum)} in-lb"
            )

            minimum_column, maximum_column = st.columns(2)
            minimum_column.metric(
                "Minimum Torque",
                f"{format_number(minimum)} in-lb",
            )
            minimum_column.caption(f"{minimum * IN_LB_TO_NM:,.2f} N·m")

            maximum_column.metric(
                "Maximum Torque",
                f"{format_number(maximum)} in-lb",
            )
            maximum_column.caption(f"{maximum * IN_LB_TO_NM:,.2f} N·m")

            if record["Direction"]:
                st.write(f"Direction: **{record['Direction']}**")

with report_tab:
    st.subheader("Generate ST and SP Excel Tables")

    report_series = st.multiselect(
        "Which series do you want in the Excel report?",
        options=sorted(database["Series"].unique()),
        default=[],
    )

    if report_series:
        preview = database[database["Series"].isin(report_series)]
        st.dataframe(
            preview[
                ["Series", "Ratio", "Gear_Type", "Min_Torque", "Max_Torque"]
            ],
            hide_index=True,
            use_container_width=True,
        )

        excel_report = create_excel_report(database, report_series)

        st.download_button(
            "Generate and Download Excel",
            data=excel_report,
            file_name="Hub_City_Bevel_Torque_ST_SP_Tables.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
    else:
        st.info("Select one or more gearbox series.")
