import re
from pathlib import Path
import pandas as pd

INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")
OUTPUT_FILE = OUTPUT_DIR / "Output_report.xlsx"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


FINAL_COLUMNS = [
    "TVC",
    "Direction",
    "Time period-from",
    "Time period-to",
    "Time",
    "Time period-Hour",
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


def normalize_tvc(text: str, fallback: str = "") -> str:
    if text is None:
        text = ""

    text = str(text).strip()

    match = re.search(r"TVC[\s\-]*(\d+)", text, flags=re.IGNORECASE)
    if match:
        num = int(match.group(1))
        return f"TVC-{num:02d}"

    match = re.search(r"(TVC[\s\-]?\d+)", fallback, flags=re.IGNORECASE)
    if match:
        num_match = re.search(r"(\d+)", match.group(1))
        if num_match:
            num = int(num_match.group(1))
            return f"TVC-{num:02d}"

    return fallback.replace(".xlsx", "").replace(".xls", "").strip()


def is_15_min_time(value: str) -> bool:
    if value is None:
        return False

    value = str(value).strip()

    pattern_1 = r"^\d{2}:\d{2}\s*-\s*\d{2}:\d{2}$"
    pattern_2 = r"^\d{4}\s*-\s*\d{4}$"

    return bool(re.match(pattern_1, value) or re.match(pattern_2, value))


def standardize_time(value: str):
    value = str(value).strip()

    if re.match(r"^\d{2}:\d{2}\s*-\s*\d{2}:\d{2}$", value):
        parts = [p.strip() for p in value.split("-")]
        start = parts[0]
        end = parts[1]

    elif re.match(r"^\d{4}\s*-\s*\d{4}$", value):
        parts = [p.strip() for p in value.split("-")]
        start = f"{parts[0][:2]}:{parts[0][2:]}"
        end = f"{parts[1][:2]}:{parts[1][2:]}"
    else:
        return None, None, None, None

    start_hour = start[:2]
    time_hour = f"{start_hour}:00 - {int(start_hour)+1:02d}:00"
    full_time = f"{start} - {end}"

    return start, end, full_time, time_hour


def clean_numeric(val):
    if pd.isna(val):
        return 0
    try:
        return int(float(val))
    except Exception:
        return 0


def extract_tvc_from_tvc01(df_raw, file_name):
    try:
        value = df_raw.iloc[0, 1]
        return normalize_tvc(value, file_name)
    except Exception:
        return normalize_tvc("", file_name)


def extract_tvc_from_tvc07(df_raw, file_name):
    try:
        for i in range(min(10, len(df_raw))):
            row_vals = df_raw.iloc[i].astype(str).tolist()
            row_text = " ".join(row_vals)
            tvc = normalize_tvc(row_text, file_name)
            if "TVC-" in tvc:
                return tvc
    except Exception:
        pass
    return normalize_tvc("", file_name)


def parse_tvc01(file_path: Path) -> pd.DataFrame:
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
    tvc = extract_tvc_from_tvc01(df_raw, file_path.name)

    all_rows = []

    # First direction block
    dir_1 = str(df_raw.iloc[1, 1]).strip()
    header_row_1 = 4
    data_start_1 = 5
    daily_total_row_1 = df_raw[df_raw.iloc[:, 0].astype(str).str.strip().eq("Daily Total")].index[0]

    # Second direction block
    dir_2 = str(df_raw.iloc[105, 1]).strip()
    header_row_2 = 108
    data_start_2 = 109
    daily_total_candidates = df_raw[df_raw.iloc[:, 0].astype(str).str.strip().eq("Daily Total")].index.tolist()
    daily_total_row_2 = daily_total_candidates[-1]

    def process_block(direction, start_idx, end_idx):
        for i in range(start_idx, end_idx):
            time_val = str(df_raw.iloc[i, 0]).strip()

            if not is_15_min_time(time_val):
                continue

            start_t, end_t, full_time, hour_time = standardize_time(time_val)

            row = {
                "TVC": tvc,
                "Direction": direction,
                "Time period-from": start_t,
                "Time period-to": end_t,
                "Time": full_time,
                "Time period-Hour": hour_time,
                "2 Wheeler": clean_numeric(df_raw.iloc[i, 1]),
                "Auto Rickshaw": clean_numeric(df_raw.iloc[i, 2]),
                "Car/Jeep/ Van": clean_numeric(df_raw.iloc[i, 3]),
                "Taxi/Ola/Uber": clean_numeric(df_raw.iloc[i, 4]),
                "Mini Bus": clean_numeric(df_raw.iloc[i, 5]),
                "Bus (Pvt)": clean_numeric(df_raw.iloc[i, 6]),
                "Bus (Gov)": clean_numeric(df_raw.iloc[i, 7]),
                "LCV(4/6-Wheels)": clean_numeric(df_raw.iloc[i, 9]),
                "Mini LCV/ Tata Ace": clean_numeric(df_raw.iloc[i, 8]),
                "Truck": clean_numeric(df_raw.iloc[i, 10]) + clean_numeric(df_raw.iloc[i, 11]) + clean_numeric(df_raw.iloc[i, 13]),
                "MAV (4-6 Axle)": clean_numeric(df_raw.iloc[i, 12]),
                "Garbage Vehicles": clean_numeric(df_raw.iloc[i, 14]),
                "Cycle": clean_numeric(df_raw.iloc[i, 18]),
                "Others": clean_numeric(df_raw.iloc[i, 19]),
            }
            all_rows.append(row)

    process_block(dir_1, data_start_1, daily_total_row_1)
    process_block(dir_2, data_start_2, daily_total_row_2)

    return pd.DataFrame(all_rows, columns=FINAL_COLUMNS)


def parse_tvc07(file_path: Path) -> pd.DataFrame:
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
    tvc = extract_tvc_from_tvc07(df_raw, file_path.name)

    direction_1 = str(df_raw.iloc[10, 1]).strip()
    direction_2 = str(df_raw.iloc[10, 22]).strip()

    dir1_rows = []
    dir2_rows = []

    for i in range(14, len(df_raw)):
        time_val = str(df_raw.iloc[i, 0]).strip()

        if not is_15_min_time(time_val):
            continue

        start_t, end_t, full_time, hour_time = standardize_time(time_val)

        row_1 = {
            "TVC": tvc,
            "Direction": direction_1,
            "Time period-from": start_t,
            "Time period-to": end_t,
            "Time": full_time,
            "Time period-Hour": hour_time,
            "2 Wheeler": clean_numeric(df_raw.iloc[i, 1]),
            "Auto Rickshaw": clean_numeric(df_raw.iloc[i, 4]),
            "Car/Jeep/ Van": clean_numeric(df_raw.iloc[i, 2]),
            "Taxi/Ola/Uber": clean_numeric(df_raw.iloc[i, 3]),
            "Mini Bus": clean_numeric(df_raw.iloc[i, 7]),
            "Bus (Pvt)": clean_numeric(df_raw.iloc[i, 9]) + clean_numeric(df_raw.iloc[i, 11]),
            "Bus (Gov)": clean_numeric(df_raw.iloc[i, 10]) + clean_numeric(df_raw.iloc[i, 12]),
            "LCV(4/6-Wheels)": clean_numeric(df_raw.iloc[i, 15]),
            "Mini LCV/ Tata Ace": clean_numeric(df_raw.iloc[i, 14]),
            "Truck": clean_numeric(df_raw.iloc[i, 16]),
            "MAV (4-6 Axle)": clean_numeric(df_raw.iloc[i, 17]),
            "Garbage Vehicles": 0,
            "Cycle": clean_numeric(df_raw.iloc[i, 18]),
            "Others": clean_numeric(df_raw.iloc[i, 20]),
        }

        row_2 = {
            "TVC": tvc,
            "Direction": direction_2,
            "Time period-from": start_t,
            "Time period-to": end_t,
            "Time": full_time,
            "Time period-Hour": hour_time,
            "2 Wheeler": clean_numeric(df_raw.iloc[i, 22]),
            "Auto Rickshaw": clean_numeric(df_raw.iloc[i, 25]),
            "Car/Jeep/ Van": clean_numeric(df_raw.iloc[i, 23]),
            "Taxi/Ola/Uber": clean_numeric(df_raw.iloc[i, 24]),
            "Mini Bus": clean_numeric(df_raw.iloc[i, 28]),
            "Bus (Pvt)": clean_numeric(df_raw.iloc[i, 30]) + clean_numeric(df_raw.iloc[i, 32]),
            "Bus (Gov)": clean_numeric(df_raw.iloc[i, 31]) + clean_numeric(df_raw.iloc[i, 33]),
            "LCV(4/6-Wheels)": clean_numeric(df_raw.iloc[i, 36]),
            "Mini LCV/ Tata Ace": clean_numeric(df_raw.iloc[i, 35]),
            "Truck": clean_numeric(df_raw.iloc[i, 37]),
            "MAV (4-6 Axle)": clean_numeric(df_raw.iloc[i, 38]),
            "Garbage Vehicles": 0,
            "Cycle": clean_numeric(df_raw.iloc[i, 39]),
            "Others": clean_numeric(df_raw.iloc[i, 41]),
        }

        dir1_rows.append(row_1)
        dir2_rows.append(row_2)

    all_rows = dir1_rows + dir2_rows
    return pd.DataFrame(all_rows, columns=FINAL_COLUMNS)


def detect_and_parse(file_path: Path) -> pd.DataFrame:
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None, nrows=15)

    preview_text = " ".join(df_raw.astype(str).fillna("").values.flatten().tolist()).lower()

    if "agar naka" in preview_text or "dewas gate" in preview_text:
        return parse_tvc07(file_path)

    if "daravali village" in preview_text or "borivali" in preview_text:
        return parse_tvc01(file_path)

    file_name_lower = file_path.name.lower()
    if "tvc-07" in file_name_lower or "tvc_07" in file_name_lower:
        return parse_tvc07(file_path)
    if "tvc-01" in file_name_lower or "tvc_01" in file_name_lower or "tvc1" in file_name_lower:
        return parse_tvc01(file_path)

    raise ValueError(f"Unsupported file format: {file_path.name}")


def main():
    excel_files = list(INPUT_DIR.glob("*.xlsx")) + list(INPUT_DIR.glob("*.xls"))

    if not excel_files:
        print("No Excel files found in data/input/")
        return

    master_df_list = []

    for file_path in excel_files:
        try:
            cleaned_df = detect_and_parse(file_path)
            master_df_list.append(cleaned_df)
            print(f"Processed: {file_path.name}")
        except Exception as e:
            print(f"Error processing {file_path.name}: {e}")

    if not master_df_list:
        print("No valid data extracted.")
        return

    master_df = pd.concat(master_df_list, ignore_index=True)
    master_df = master_df[FINAL_COLUMNS]

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        master_df.to_excel(writer, sheet_name="Master_data", index=False)

    print(f"Output created: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()