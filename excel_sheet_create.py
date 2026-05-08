from pathlib import Path
from openpyxl import load_workbook
import xlsxwriter

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "data" / "output"

SOURCE_FILE = OUTPUT_DIR / "Cleaned_Data.xlsx"
REPORT_FILE = OUTPUT_DIR / "Report.xlsx"
PCU_CONFIG_FILE = OUTPUT_DIR / "PCU_Config.xlsx"

SOURCE_SHEET = "Master_data"

DEFAULT_PCU_MAP = {
    "2 Wheeler": 0.2,
    "Auto Rickshaw": 0.8,
    "Car/Jeep/ Van": 1,
    "Taxi/Ola/Uber": 1,
    "Mini Bus": 2.1,
    "Bus (Pvt)": 4.5,
    "Bus (Gov)": 2.3,
    "LCV(4/6-Wheels)": 3.8,
    "Mini LCV/ Tata Ace": 2.3,
    "Truck": 3.8,
    "MAV (4-6 Axle)": 5.1,
    "Garbage Vehicles": 3.8,
    "Cycle": 0.4,
    "Others": 2,
}

VEHICLE_TYPES = [
    "2 Wheeler",
    "Auto Rickshaw",
    "Car/Jeep/ Van",
    "Taxi/Ola/Uber",
    "Mini Bus",
    "Bus (Pvt)",
    "Bus (Gov)",
    "LCV(4/6-Wheels)",
    "Mini LCV/ Tata Ace",
    "Truck",
    "MAV (4-6 Axle)",
    "Garbage Vehicles",
    "Cycle",
    "Others",
]


def normalize_text(value):
    return " ".join(str(value).strip().lower().split())


def get_default_pcu_value(vehicle_type):
    vehicle_type_normalized = normalize_text(vehicle_type)

    for key, value in DEFAULT_PCU_MAP.items():
        if normalize_text(key) == vehicle_type_normalized:
            return value

    return ""

def load_pcu_config():
    if PCU_CONFIG_FILE.exists():
        try:
            return load_workbook(PCU_CONFIG_FILE, data_only=True)
        except Exception:
            return None
    return None


def get_tvc_specific_pcu_value(tvc_name, vehicle_type):
    if not PCU_CONFIG_FILE.exists():
        return get_default_pcu_value(vehicle_type)

    try:
        import pandas as pd
        df = pd.read_excel(PCU_CONFIG_FILE)

        if "TVC" not in df.columns:
            return get_default_pcu_value(vehicle_type)

        match = df[df["TVC"].astype(str).str.strip() == str(tvc_name).strip()]
        if match.empty:
            return get_default_pcu_value(vehicle_type)

        row = match.iloc[0]

        if vehicle_type in row and row[vehicle_type] not in (None, ""):
            try:
                return float(row[vehicle_type])
            except (TypeError, ValueError):
                return get_default_pcu_value(vehicle_type)

        return get_default_pcu_value(vehicle_type)

    except Exception:
        return get_default_pcu_value(vehicle_type)


def get_column_index_by_header(ws, expected_header):
    headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    expected_normalized = str(expected_header).strip().lower().replace("_", " ").replace("\n", " ")

    for idx, header in enumerate(headers, start=1):
        if not header:
            continue

        header_normalized = str(header).strip().lower().replace("_", " ").replace("\n", " ")
        if header_normalized == expected_normalized:
            return idx

    return None


def get_unique_tvc_names(ws):
    tvc_col_idx = get_column_index_by_header(ws, "TVC")
    if not tvc_col_idx:
        raise ValueError("Column 'TVC' not found in Master_Data sheet")

    tvc_names = []
    seen = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        value = row[tvc_col_idx - 1]
        if value is None:
            continue

        tvc_name = str(value).strip()
        if tvc_name and tvc_name not in seen:
            seen.add(tvc_name)
            tvc_names.append(tvc_name)

    return tvc_names


def get_tvc_wise_vehicle_totals(ws):
    tvc_col_idx = get_column_index_by_header(ws, "TVC")
    if not tvc_col_idx:
        raise ValueError("Column 'TVC' not found in Master_Data sheet")

    vehicle_col_indexes = {}
    for vehicle in VEHICLE_TYPES:
        col_idx = get_column_index_by_header(ws, vehicle)
        if col_idx:
            vehicle_col_indexes[vehicle] = col_idx
        else:
            print(f"Warning: Column not found for vehicle '{vehicle}'")

    tvc_totals = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        tvc_value = row[tvc_col_idx - 1]
        if tvc_value is None:
            continue

        tvc_name = str(tvc_value).strip()
        if not tvc_name:
            continue

        if tvc_name not in tvc_totals:
            tvc_totals[tvc_name] = {vehicle: 0 for vehicle in VEHICLE_TYPES}

        for vehicle, col_idx in vehicle_col_indexes.items():
            value = row[col_idx - 1]
            if value in (None, ""):
                value = 0

            try:
                value = float(value)
            except (TypeError, ValueError):
                value = 0

            tvc_totals[tvc_name][vehicle] += value

    return tvc_totals


def copy_sheet_data_to_xlsxwriter(source_ws, target_ws):
    for r_idx, row in enumerate(source_ws.iter_rows(values_only=True)):
        for c_idx, value in enumerate(row):
            target_ws.write(r_idx, c_idx, value)


def populate_pcu_sheet_xlsxwriter(pcu_ws, master_ws):
    tvc_names = get_unique_tvc_names(master_ws)

    pcu_ws.write(0, 0, "Vehicle Type")

    for col_idx, tvc_name in enumerate(tvc_names, start=1):
        pcu_ws.write(0, col_idx, tvc_name)

    for row_idx, vehicle_type in enumerate(VEHICLE_TYPES, start=1):
        pcu_ws.write(row_idx, 0, vehicle_type)

        for col_idx, tvc_name in enumerate(tvc_names, start=1):
            pcu_value = get_tvc_specific_pcu_value(tvc_name, vehicle_type)
            pcu_ws.write(row_idx, col_idx, pcu_value)


def get_grouped_vehicle_data(vehicle_data):
    grouped = {
        "2 Wheeler": vehicle_data.get("2 Wheeler", 0),
        "Auto Rickshaw": vehicle_data.get("Auto Rickshaw", 0),
        "Car/Jeep/ Van": vehicle_data.get("Car/Jeep/ Van", 0),
        "Taxi/Ola/Uber": vehicle_data.get("Taxi/Ola/Uber", 0),
        "Bus": (
            vehicle_data.get("Mini Bus", 0)
            + vehicle_data.get("Bus (Pvt)", 0)
            + vehicle_data.get("Bus (Gov)", 0)
        ),
        "LCV(4/6-Wheels)": (
            vehicle_data.get("LCV(4/6-Wheels)", 0)
            + vehicle_data.get("Mini LCV/ Tata Ace", 0)
        ),
        "Truck": (
            vehicle_data.get("Truck", 0)
            + vehicle_data.get("MAV (4-6 Axle)", 0)
            + vehicle_data.get("Garbage Vehicles", 0)
        ),
        "Cycle": vehicle_data.get("Cycle", 0),
        "Others": vehicle_data.get("Others", 0),
    }

    # remove zero-value categories from chart
    return {k: v for k, v in grouped.items() if v > 0}


def write_grouped_chart_data(vc_ws, start_row, grouped_vehicle_data):
    """
    Writes grouped chart data in columns K:N
    K = grouped vehicle class
    L = grouped vehicle count
    M = grouped PCU total
    N = vehicle percentage
    O = PCU percentage
    """
    vc_ws.write(start_row, 10, "Vehicles")
    vc_ws.write(start_row, 11, "Vehicle Count")
    vc_ws.write(start_row, 12, "PCU")
    vc_ws.write(start_row, 13, "Vehicle %")
    vc_ws.write(start_row, 14, "PCU %")

    total_vehicle_count = sum(grouped_vehicle_data.values())

    grouped_pcu_data = {}
    for vehicle_class, count in grouped_vehicle_data.items():
        if vehicle_class == "Bus":
            pcu_total = (
                grouped_vehicle_data.get("Bus", 0)  # temp placeholder, overwritten below
            )
        elif vehicle_class == "LCV(4/6-Wheels)":
            pcu_total = (
                0  # temp placeholder, overwritten below
            )
        elif vehicle_class == "Truck":
            pcu_total = (
                0  # temp placeholder, overwritten below
            )
        else:
            pcu_total = count * get_default_pcu_value(vehicle_class)

        grouped_pcu_data[vehicle_class] = pcu_total

    # correct grouped PCUs manually
    grouped_pcu_data["Bus"] = 0
    grouped_pcu_data["LCV(4/6-Wheels)"] = 0
    grouped_pcu_data["Truck"] = 0

    total_pcu = sum(grouped_pcu_data.values())

    row = start_row + 1
    for vehicle_class, count in grouped_vehicle_data.items():
        vc_ws.write(row, 10, vehicle_class)
        vc_ws.write(row, 11, count)
        vc_ws.write(row, 12, grouped_pcu_data.get(vehicle_class, 0))
        vc_ws.write(row, 13, (count / total_vehicle_count) if total_vehicle_count else 0)
        vc_ws.write(row, 14, (grouped_pcu_data.get(vehicle_class, 0) / total_pcu) if total_pcu else 0)
        row += 1

    return start_row + 1, row - 1


def get_grouped_pcu_data_from_original(tvc_name, vehicle_data):
    return {
        "2 Wheeler": vehicle_data.get("2 Wheeler", 0) * get_tvc_specific_pcu_value(tvc_name, "2 Wheeler"),
        "Auto Rickshaw": vehicle_data.get("Auto Rickshaw", 0) * get_tvc_specific_pcu_value(tvc_name, "Auto Rickshaw"),
        "Car/Jeep/ Van": vehicle_data.get("Car/Jeep/ Van", 0) * get_tvc_specific_pcu_value(tvc_name, "Car/Jeep/ Van"),
        "Taxi/Ola/Uber": vehicle_data.get("Taxi/Ola/Uber", 0) * get_tvc_specific_pcu_value(tvc_name, "Taxi/Ola/Uber"),
        "Bus": (
            vehicle_data.get("Mini Bus", 0) * get_tvc_specific_pcu_value(tvc_name, "Mini Bus")
            + vehicle_data.get("Bus (Pvt)", 0) * get_tvc_specific_pcu_value(tvc_name, "Bus (Pvt)")
            + vehicle_data.get("Bus (Gov)", 0) * get_tvc_specific_pcu_value(tvc_name, "Bus (Gov)")
        ),
        "LCV(4/6-Wheels)": (
            vehicle_data.get("LCV(4/6-Wheels)", 0) * get_tvc_specific_pcu_value(tvc_name, "LCV(4/6-Wheels)")
            + vehicle_data.get("Mini LCV/ Tata Ace", 0) * get_tvc_specific_pcu_value(tvc_name, "Mini LCV/ Tata Ace")
        ),
        "Truck": (
            vehicle_data.get("Truck", 0) * get_tvc_specific_pcu_value(tvc_name, "Truck")
            + vehicle_data.get("MAV (4-6 Axle)", 0) * get_tvc_specific_pcu_value(tvc_name, "MAV (4-6 Axle)")
            + vehicle_data.get("Garbage Vehicles", 0) * get_tvc_specific_pcu_value(tvc_name, "Garbage Vehicles")
        ),
        "Cycle": vehicle_data.get("Cycle", 0) * get_tvc_specific_pcu_value(tvc_name, "Cycle"),
        "Others": vehicle_data.get("Others", 0) * get_tvc_specific_pcu_value(tvc_name, "Others"),
    }


def add_pie_charts(workbook, worksheet, sheet_name, chart_data_start_row, chart_data_end_row, tvc_name):
    vehicle_chart = workbook.add_chart({'type': 'pie'})

    vehicle_chart.add_series({
        'name': 'Vehicle Composition',
        'categories': [sheet_name, chart_data_start_row, 10, chart_data_end_row, 10],  # K
        'values':     [sheet_name, chart_data_start_row, 11, chart_data_end_row, 11],  # L
        'data_labels': {
            'percentage': True,
            'font': {'size': 10},
            'position': 'outside_end',
        },
    })

    vehicle_chart.set_title({
        'name': f'{tvc_name} Vehicle Composition',
        'name_font': {'size': 14, 'bold': True},
    })

    vehicle_chart.set_legend({
        'position': 'right',
        'font': {'size': 8},
    })

    vehicle_chart.set_chartarea({'border': {'none': True}})
    vehicle_chart.set_plotarea({'border': {'none': True}})

    worksheet.insert_chart(chart_data_start_row - 1, 15, vehicle_chart, {
        'x_scale': 1.00,
        'y_scale': 1.00,
    })

    pcu_chart = workbook.add_chart({'type': 'pie'})

    pcu_chart.add_series({
        'name': 'PCU Composition',
        'categories': [sheet_name, chart_data_start_row, 10, chart_data_end_row, 10],  # K
        'values':     [sheet_name, chart_data_start_row, 12, chart_data_end_row, 12],  # M
        'data_labels': {
            'percentage': True,
            'font': {'size': 10},
            'position': 'outside_end',
        },
    })

    pcu_chart.set_title({
        'name': f'{tvc_name} PCU Composition',
        'name_font': {'size': 14, 'bold': True},
    })

    pcu_chart.set_legend({
        'position': 'right',
        'font': {'size': 8},
    })

    pcu_chart.set_chartarea({'border': {'none': True}})
    pcu_chart.set_plotarea({'border': {'none': True}})

    worksheet.insert_chart(chart_data_start_row - 1, 24, pcu_chart, {
        'x_scale': 1.00,
        'y_scale': 1.00,
    })


def populate_vehical_composition_sheet_xlsxwriter(workbook, vc_ws, master_ws):
    percent_format = workbook.add_format({'num_format': '0%'})

    tvc_totals = get_tvc_wise_vehicle_totals(master_ws)
    current_row = 0

    for tvc_name, vehicle_data in tvc_totals.items():
        vc_ws.write(current_row, 0, "Location")
        vc_ws.write(current_row, 1, tvc_name)
        current_row += 2

        vc_ws.write(current_row, 0, "Values")
        vc_ws.write(current_row, 3, "Vehicles")
        vc_ws.write(current_row, 4, "Vehicle count")
        vc_ws.write(current_row, 5, "Vehicle Composition")
        vc_ws.write(current_row, 6, "PCU")
        vc_ws.write(current_row, 7, "PCUs")
        current_row += 1

        total_vehicles = sum(vehicle_data.values())

        for vehicle in VEHICLE_TYPES:
            vehicles_count = vehicle_data.get(vehicle, 0)
            pcu_value = get_tvc_specific_pcu_value(tvc_name, vehicle)

            composition = (vehicles_count / total_vehicles) if total_vehicles > 0 else 0
            pcu_total = vehicles_count * pcu_value if pcu_value != "" else 0

            vc_ws.write(current_row, 0, f"Sum of {vehicle}")
            vc_ws.write(current_row, 1, vehicles_count)
            vc_ws.write(current_row, 3, vehicle)
            vc_ws.write(current_row, 4, vehicles_count)
            vc_ws.write(current_row, 5, composition, percent_format)
            vc_ws.write(current_row, 6, pcu_value)
            vc_ws.write(current_row, 7, pcu_total)
            current_row += 1

        vc_ws.write(current_row, 1, total_vehicles)

        grouped_vehicle_data = get_grouped_vehicle_data(vehicle_data)
        grouped_pcu_data = get_grouped_pcu_data_from_original(tvc_name, vehicle_data)

        # write grouped chart data
        chart_header_row = current_row - len(VEHICLE_TYPES) - 1
        vc_ws.write(chart_header_row, 10, "Vehicles")
        vc_ws.write(chart_header_row, 11, "Vehicle Count")
        vc_ws.write(chart_header_row, 12, "PCU")

        chart_data_start_row = chart_header_row + 1
        temp_row = chart_data_start_row

        for category, vehicle_count in grouped_vehicle_data.items():
            vc_ws.write(temp_row, 10, category)
            vc_ws.write(temp_row, 11, vehicle_count)
            vc_ws.write(temp_row, 12, grouped_pcu_data.get(category, 0))
            temp_row += 1

        chart_data_end_row = temp_row - 1

        add_pie_charts(
            workbook=workbook,
            worksheet=vc_ws,
            sheet_name="Vehical_Composition",
            chart_data_start_row=chart_data_start_row,
            chart_data_end_row=chart_data_end_row,
            tvc_name=tvc_name
        )

        current_row += 18


def create_report():
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"Source file not found: {SOURCE_FILE}")

    source_wb = load_workbook(SOURCE_FILE, data_only=True)
    print("Available sheets:", source_wb.sheetnames)

    if SOURCE_SHEET not in source_wb.sheetnames:
        raise ValueError(
            f"{SOURCE_SHEET} sheet not found in {SOURCE_FILE.name}. "
            f"Available sheets: {source_wb.sheetnames}"
        )

    source_ws = source_wb[SOURCE_SHEET]

    report_wb = xlsxwriter.Workbook(str(REPORT_FILE))

    master_ws = report_wb.add_worksheet("Master_Data")
    pcu_ws = report_wb.add_worksheet("PCU")
    vc_ws = report_wb.add_worksheet("Vehical_Composition")

    copy_sheet_data_to_xlsxwriter(source_ws, master_ws)
    populate_pcu_sheet_xlsxwriter(pcu_ws, source_ws)
    populate_vehical_composition_sheet_xlsxwriter(report_wb, vc_ws, source_ws)

    report_wb.close()
    print(f"Report created successfully: {REPORT_FILE}")


if __name__ == "__main__":
    create_report()