from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "data" / "output"

SOURCE_FILE = OUTPUT_DIR / "Cleaned_Data.xlsx"
REPORT_FILE = OUTPUT_DIR / "Report.xlsx"

SOURCE_SHEET = "Master_data"
TARGET_SHEETS = ["Master_Data", "PCU", "Vehical_Composition"]

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


def copy_sheet_data(source_ws, target_ws):
    for row in source_ws.iter_rows(values_only=True):
        target_ws.append(list(row))


def debug_headers(ws):
    headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    print("Master_Data headers:", list(headers))


def debug_sample_data(ws):
    for row in ws.iter_rows(min_row=2, max_row=6, values_only=True):
        print("Sample row:", row)


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

def populate_vehical_composition_sheet(master_ws, vc_ws):
    tvc_totals = get_tvc_wise_vehicle_totals(master_ws)
    current_row = 1

    for tvc_name, vehicle_data in tvc_totals.items():
        vc_ws.cell(row=current_row, column=1, value="Location")
        vc_ws.cell(row=current_row, column=2, value=tvc_name)
        current_row += 2

        vc_ws.cell(row=current_row, column=1, value="Values")
        vc_ws.cell(row=current_row, column=4, value="Vehicles")
        vc_ws.cell(row=current_row, column=5, value="Vehicle count")
        vc_ws.cell(row=current_row, column=6, value="Vehicle Composition")
        vc_ws.cell(row=current_row, column=7, value="PCU")
        vc_ws.cell(row=current_row, column=8, value="PCUs")
        current_row += 1

        data_start_row = current_row
        total_vehicles = sum(vehicle_data.values())

        for vehicle in VEHICLE_TYPES:
            vehicles_count = vehicle_data.get(vehicle, 0)
            pcu_value = get_default_pcu_value(vehicle)

            if total_vehicles > 0:
                composition = vehicles_count / total_vehicles
            else:
                composition = 0

            pcu_total = vehicles_count * pcu_value if pcu_value != "" else 0

            vc_ws.cell(row=current_row, column=1, value=f"Sum of {vehicle}")
            vc_ws.cell(row=current_row, column=2, value=vehicles_count)
            vc_ws.cell(row=current_row, column=4, value=vehicle)
            vc_ws.cell(row=current_row, column=5, value=vehicles_count)

            composition_cell = vc_ws.cell(row=current_row, column=6, value=composition)
            composition_cell.number_format = '0%'

            vc_ws.cell(row=current_row, column=7, value=pcu_value)
            vc_ws.cell(row=current_row, column=8, value=pcu_total)

            current_row += 1

        data_end_row = current_row - 1

        vc_ws.cell(row=current_row, column=2, value=total_vehicles)

        add_pie_charts(vc_ws, data_start_row, data_end_row, tvc_name)

        current_row += 18

def add_pie_charts(ws, start_row, end_row, tvc_name):
    # Vehicle Composition chart
    vehicle_chart = PieChart()
    labels = Reference(ws, min_col=4, min_row=start_row, max_row=end_row)   # Vehicles names in col D
    data = Reference(ws, min_col=5, min_row=start_row - 1, max_row=end_row) # Vehicle counts in col E, with header
    vehicle_chart.add_data(data, titles_from_data=True)
    vehicle_chart.set_categories(labels)
    vehicle_chart.title = f"{tvc_name} Vehicle Composition"
    vehicle_chart.height = 12
    vehicle_chart.width = 16
    vehicle_chart.legend.position = "r"
    vehicle_chart.dataLabels = DataLabelList()
    vehicle_chart.dataLabels.showPercent = True

    # PCU Composition chart
    pcu_chart = PieChart()
    pcu_labels = Reference(ws, min_col=4, min_row=start_row, max_row=end_row)   # Vehicles names in col D
    pcu_data = Reference(ws, min_col=8, min_row=start_row - 1, max_row=end_row) # PCUs in col H, with header
    pcu_chart.add_data(pcu_data, titles_from_data=True)
    pcu_chart.set_categories(pcu_labels)
    pcu_chart.title = f"{tvc_name} PCU Composition"
    pcu_chart.height = 12
    pcu_chart.width = 16
    pcu_chart.legend.position = "r"
    pcu_chart.dataLabels = DataLabelList()
    pcu_chart.dataLabels.showPercent = True

    # Place charts after column J
    ws.add_chart(vehicle_chart, f"K{start_row - 2}")
    ws.add_chart(pcu_chart, f"U{start_row - 2}")

def normalize_text(value):
    return " ".join(str(value).strip().lower().split())


def get_default_pcu_value(vehicle_type):
    vehicle_type_normalized = normalize_text(vehicle_type)

    for key, value in DEFAULT_PCU_MAP.items():
        if normalize_text(key) == vehicle_type_normalized:
            return value

    return ""


def populate_pcu_sheet(master_ws, pcu_ws):
    tvc_names = get_unique_tvc_names(master_ws)

    pcu_ws.cell(row=1, column=1, value="Vehicle Type")

    for col_idx, tvc_name in enumerate(tvc_names, start=2):
        pcu_ws.cell(row=1, column=col_idx, value=tvc_name)

    for row_idx, vehicle_type in enumerate(VEHICLE_TYPES, start=2):
        pcu_ws.cell(row=row_idx, column=1, value=vehicle_type)

        default_pcu = get_default_pcu_value(vehicle_type)
        for col_idx in range(2, len(tvc_names) + 2):
            pcu_ws.cell(row=row_idx, column=col_idx, value=default_pcu)


def create_report():
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"Source file not found: {SOURCE_FILE}")

    source_wb = load_workbook(SOURCE_FILE)
    print("Available sheets:", source_wb.sheetnames)

    if SOURCE_SHEET not in source_wb.sheetnames:
        raise ValueError(
            f"{SOURCE_SHEET} sheet not found in {SOURCE_FILE.name}. "
            f"Available sheets: {source_wb.sheetnames}"
        )

    source_ws = source_wb[SOURCE_SHEET]

    report_wb = Workbook()
    report_wb.active.title = TARGET_SHEETS[0]

    for sheet_name in TARGET_SHEETS[1:]:
        report_wb.create_sheet(sheet_name)

    target_ws = report_wb[TARGET_SHEETS[0]]
    copy_sheet_data(source_ws, target_ws)

    debug_headers(target_ws)
    debug_sample_data(target_ws)

    pcu_ws = report_wb["PCU"]
    populate_pcu_sheet(target_ws, pcu_ws)

    vc_ws = report_wb["Vehical_Composition"]
    populate_vehical_composition_sheet(target_ws, vc_ws)

    report_wb.save(REPORT_FILE)
    print(f"Report created successfully: {REPORT_FILE}")


if __name__ == "__main__":
    create_report()