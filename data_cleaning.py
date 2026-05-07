import re
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd

INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")
OUTPUT_FILE = OUTPUT_DIR / "Cleaned_Data.xlsx"

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

TOTAL_PATTERNS = [
    r"hour\s*total",
    r"daily\s*total",
    r"grand\s*total",
    r"pcu\s*total",
    r"^\s*total\s*$",
]

GENERIC_PATTERNS = {
    "2 Wheeler": [r"2\s*wheeler", r"two\s*wheeler"],
    "Auto Rickshaw": [r"auto", r"rickshaw", r"3\s*wheeler"],
    "Car/Jeep/ Van": [r"car", r"jeep", r"van"],
    "Taxi/Ola/Uber": [r"taxi", r"ola", r"uber", r"cab"],
    "Mini Bus": [r"mini\s*bus"],
    "Bus (Pvt)": [r"bus\s*\(pvt\)", r"standard\s*bus\s*\(pvt\)", r"private", r"ac\s*private", r"non\s*ac\s*private"],
    "Bus (Gov)": [r"bus\s*\(gov\)", r"standard\s*bus\s*\(gov\)", r"government", r"\bgov\b", r"ac\s*government", r"non\s*ac\s*government"],
    "LCV(4/6-Wheels)": [r"\blcv\b", r"lcv\s*\(4/6[-\s]*wheels?\)"],
    "Mini LCV/ Tata Ace": [r"mini\s*lcv", r"tata\s*ace"],
    "Truck": [r"truck", r"2-axle", r"3-axle", r"2-3 axle", r"oversized"],
    "MAV (4-6 Axle)": [r"\bmav\b", r"4-6 axle"],
    "Garbage Vehicles": [r"garbage"],
    "Cycle": [r"\bcycle\b"],
    "Others": [r"\bothers\b"],
}


def norm_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def clean_numeric(val):
    if pd.isna(val):
        return 0
    s = str(val).strip().replace(",", "")
    if s == "":
        return 0
    try:
        return int(float(s))
    except:
        return 0


def normalize_tvc(text, fallback=""):
    text = norm_text(text)
    m = re.search(r"TVC[\s\-]*(\d+)", text, flags=re.I)
    if m:
        return f"TVC-{int(m.group(1)):02d}"

    fallback = norm_text(fallback)
    m2 = re.search(r"TVC[\s\-]*(\d+)", fallback, flags=re.I)
    if m2:
        return f"TVC-{int(m2.group(1)):02d}"

    return fallback.replace(".xlsx", "").replace(".xls", "")


def looks_like_total(text):
    t = norm_text(text).lower()
    return any(re.search(p, t) for p in TOTAL_PATTERNS)

def clean_direction_text(text):
    text = norm_text(text)
    text = re.sub(r"^\s*Direction\s*", "", text, flags=re.I)
    text = re.sub(r"\s*Day\s*\d{4}-\d{2}-\d{2}.*$", "", text, flags=re.I)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def is_time_interval(value):
    s = norm_text(value)
    if re.match(r"^\d{2}:\d{2}\s*-\s*\d{2}:\d{2}$", s):
        return True
    if re.match(r"^\d{4}\s*-\s*\d{4}$", s):
        return True
    return False


def standardize_time(value):
    s = norm_text(value)

    if re.match(r"^\d{2}:\d{2}\s*-\s*\d{2}:\d{2}$", s):
        p1, p2 = [x.strip() for x in s.split("-")]
        start = p1
        end = p2
    elif re.match(r"^\d{4}\s*-\s*\d{4}$", s):
        p1, p2 = [x.strip() for x in s.split("-")]
        start = f"{p1[:2]}:{p1[2:]}"
        end = f"{p2[:2]}:{p2[2:]}"
    else:
        return None

    dt = datetime.strptime(start, "%H:%M")
    hour_start = dt.replace(minute=0)
    hour_end = hour_start + timedelta(hours=1)

    return {
        "Time period-from": start,
        "Time period-to": end,
        "Time": f"{start} - {end}",
        "Time period-Hour": f"{hour_start.strftime('%H:%M')} - {hour_end.strftime('%H:%M')}",
    }


def make_record(tvc, direction, time_parts, values):
    rec = {
        "TVC": tvc,
        "Direction": direction,
        "Time period-from": time_parts["Time period-from"],
        "Time period-to": time_parts["Time period-to"],
        "Time": time_parts["Time"],
        "Time period-Hour": time_parts["Time period-Hour"],
        "2 Wheeler": 0,
        "Auto Rickshaw": 0,
        "Car/Jeep/ Van": 0,
        "Taxi/Ola/Uber": 0,
        "Mini Bus": 0,
        "Bus (Pvt)": 0,
        "Bus (Gov)": 0,
        "LCV(4/6-Wheels)": 0,
        "Mini LCV/ Tata Ace": 0,
        "Truck": 0,
        "MAV (4-6 Axle)": 0,
        "Garbage Vehicles": 0,
        "Cycle": 0,
        "Others": 0,
    }
    rec.update(values)
    return rec


def extract_tvc_from_sheet(df, file_name):
    text = " ".join(norm_text(v) for v in df.head(25).fillna("").values.flatten().tolist())
    return normalize_tvc(text, file_name)


def score_header_row(row_values):
    score = 0
    texts = [norm_text(v).lower() for v in row_values]
    for cell in texts:
        if "time" in cell:
            score += 2
        for pats in GENERIC_PATTERNS.values():
            if any(re.search(p, cell) for p in pats):
                score += 1
                break
    return score


def find_header_candidates(df):
    candidates = []
    for i in range(min(80, len(df))):
        score = score_header_row(df.iloc[i].tolist())
        if score >= 5:
            candidates.append(i)
    return candidates


def split_blocks(row):
    non_empty = [i for i, v in enumerate(row) if norm_text(v) != ""]
    if not non_empty:
        return []

    blocks = []
    start = non_empty[0]
    prev = non_empty[0]

    for idx in non_empty[1:]:
        if idx - prev > 3:
            blocks.append((start, prev + 1))
            start = idx
        prev = idx
    blocks.append((start, prev + 1))
    return blocks


def find_direction(df, header_row_idx, c1, c2):
    for r in range(max(0, header_row_idx - 6), min(len(df), header_row_idx + 1)):
        vals = [norm_text(df.iloc[r, c]) for c in range(c1, min(c2, df.shape[1]))]
        joined = " ".join([v for v in vals if v]).strip()
        low = joined.lower()
        if " to " in low or " - " in low or "gate" in low or "naka" in low or "village" in low:
            return joined
    return "Unknown Direction"


def build_colmap(header_cells):
    colmap = {}
    for idx, cell in enumerate(header_cells):
        txt = norm_text(cell).lower()

        if txt == "time" or "time" in txt:
            colmap["time"] = idx

        for key, pats in GENERIC_PATTERNS.items():
            if any(re.search(p, txt) for p in pats):
                colmap.setdefault(key, []).append(idx)

    return colmap


def extract_dynamic(df, file_name):
    tvc = extract_tvc_from_sheet(df, file_name)
    out = []

    for hr in find_header_candidates(df):
        row = df.iloc[hr].tolist()
        blocks = split_blocks(row)

        for c1, c2 in blocks:
            header_cells = row[c1:c2]
            colmap = build_colmap(header_cells)

            if "time" not in colmap:
                continue

            direction = find_direction(df, hr, c1, c2)

            r = hr + 1
            while r < len(df):
                row_slice = df.iloc[r, c1:c2].tolist()
                joined = " ".join(norm_text(x) for x in row_slice).strip()

                if not joined:
                    if r > hr + 5:
                        break
                    r += 1
                    continue

                if looks_like_total(joined):
                    r += 1
                    continue

                time_idx = colmap["time"]
                time_val = norm_text(row_slice[time_idx]) if time_idx < len(row_slice) else ""

                if not is_time_interval(time_val):
                    r += 1
                    continue

                time_parts = standardize_time(time_val)
                if not time_parts:
                    r += 1
                    continue

                vals = {k: 0 for k in GENERIC_PATTERNS.keys()}
                for key in GENERIC_PATTERNS.keys():
                    if key in colmap:
                        for idx in colmap[key]:
                            if idx < len(row_slice):
                                vals[key] += clean_numeric(row_slice[idx])

                if sum(vals.values()) > 0:
                    out.append(make_record(tvc, direction, time_parts, vals))

                r += 1

    return out


def parse_tvc07_style(df, file_name):
    tvc = extract_tvc_from_sheet(df, file_name)
    results = []

    direction_1 = ""
    direction_2 = ""

    for r in range(min(20, len(df))):
        for c in range(min(df.shape[1], 25)):
            txt = norm_text(df.iloc[r, c])
            low = txt.lower()
            if not direction_1 and ("agar naka" in low or "dewas gate" in low):
                direction_1 = txt

            if c + 21 < df.shape[1]:
                txt2 = norm_text(df.iloc[r, c + 21])
                low2 = txt2.lower()
                if not direction_2 and ("agar naka" in low2 or "dewas gate" in low2):
                    direction_2 = txt2

    if not direction_1:
        direction_1 = "Agar Naka - Dewas Gate"
    if not direction_2:
        direction_2 = "Dewas Gate - Agar Naka"

    dir1_rows = []
    dir2_rows = []

    for i in range(len(df)):
        time_val = norm_text(df.iloc[i, 0]) if df.shape[1] > 0 else ""
        if not is_time_interval(time_val):
            continue

        time_parts = standardize_time(time_val)
        if not time_parts:
            continue

        if df.shape[1] >= 21:
            row1 = make_record(
                tvc,
                direction_1,
                time_parts,
                {
                    "2 Wheeler": clean_numeric(df.iloc[i, 1]),
                    "Auto Rickshaw": clean_numeric(df.iloc[i, 4]),
                    "Car/Jeep/ Van": clean_numeric(df.iloc[i, 2]),
                    "Taxi/Ola/Uber": clean_numeric(df.iloc[i, 3]),
                    "Mini Bus": clean_numeric(df.iloc[i, 7]),
                    "Bus (Pvt)": clean_numeric(df.iloc[i, 9]) + clean_numeric(df.iloc[i, 11]),
                    "Bus (Gov)": clean_numeric(df.iloc[i, 10]) + clean_numeric(df.iloc[i, 12]),
                    "Mini LCV/ Tata Ace": clean_numeric(df.iloc[i, 14]),
                    "LCV(4/6-Wheels)": clean_numeric(df.iloc[i, 15]),
                    "Truck": clean_numeric(df.iloc[i, 16]),
                    "MAV (4-6 Axle)": clean_numeric(df.iloc[i, 17]),
                    "Garbage Vehicles": 0,
                    "Cycle": clean_numeric(df.iloc[i, 18]),
                    "Others": clean_numeric(df.iloc[i, 20]),
                },
            )
            if sum(row1[c] for c in FINAL_COLUMNS[6:]) > 0:
                dir1_rows.append(row1)

        if df.shape[1] >= 42:
            row2 = make_record(
                tvc,
                direction_2,
                time_parts,
                {
                    "2 Wheeler": clean_numeric(df.iloc[i, 22]),
                    "Auto Rickshaw": clean_numeric(df.iloc[i, 25]),
                    "Car/Jeep/ Van": clean_numeric(df.iloc[i, 23]),
                    "Taxi/Ola/Uber": clean_numeric(df.iloc[i, 24]),
                    "Mini Bus": clean_numeric(df.iloc[i, 28]),
                    "Bus (Pvt)": clean_numeric(df.iloc[i, 30]) + clean_numeric(df.iloc[i, 32]),
                    "Bus (Gov)": clean_numeric(df.iloc[i, 31]) + clean_numeric(df.iloc[i, 33]),
                    "Mini LCV/ Tata Ace": clean_numeric(df.iloc[i, 35]),
                    "LCV(4/6-Wheels)": clean_numeric(df.iloc[i, 36]),
                    "Truck": clean_numeric(df.iloc[i, 37]),
                    "MAV (4-6 Axle)": clean_numeric(df.iloc[i, 38]),
                    "Garbage Vehicles": 0,
                    "Cycle": clean_numeric(df.iloc[i, 39]),
                    "Others": clean_numeric(df.iloc[i, 41]),
                },
            )
            if sum(row2[c] for c in FINAL_COLUMNS[6:]) > 0:
                dir2_rows.append(row2)

    results.extend(dir1_rows)
    results.extend(dir2_rows)
    return results


def parse_tvc01_style(df, file_name):
    tvc = extract_tvc_from_sheet(df, file_name)
    results = []

    possible_directions = []
    for r in range(min(len(df), 250)):
        row_joined = " | ".join(norm_text(df.iloc[r, c]) for c in range(min(df.shape[1], 6)))
        low = row_joined.lower()

        if "direction" in low and " to " in low:
            direction_text = ""

            for c in range(min(df.shape[1], 6)):
                txt = norm_text(df.iloc[r, c])
                if " to " in txt.lower():
                    direction_text = txt
                    break

            if direction_text:
                direction_text = clean_direction_text(direction_text)
                if direction_text not in [d[1] for d in possible_directions]:
                    possible_directions.append((r, direction_text))

    if len(possible_directions) < 2:
        return results

    blocks = [
        (possible_directions[0][0], possible_directions[0][1]),
        (possible_directions[1][0], possible_directions[1][1]),
    ]

    for block_idx, (start_row, direction) in enumerate(blocks):
        end_row = blocks[block_idx + 1][0] if block_idx + 1 < len(blocks) else len(df)

        header_row = None
        for r in range(start_row, min(start_row + 10, len(df))):
            first_cell = norm_text(df.iloc[r, 0]).lower()
            if first_cell == "time":
                header_row = r
                break

        if header_row is None:
            continue

        for r in range(header_row + 1, end_row):
            time_val = norm_text(df.iloc[r, 0]) if df.shape[1] > 0 else ""

            if looks_like_total(time_val):
                continue

            if not is_time_interval(time_val):
                continue

            time_parts = standardize_time(time_val)
            if not time_parts:
                continue

            row_record = make_record(
                tvc,
                direction,
                time_parts,
                {
                    "2 Wheeler": clean_numeric(df.iloc[r, 1]) if df.shape[1] > 1 else 0,
                    "Auto Rickshaw": clean_numeric(df.iloc[r, 2]) if df.shape[1] > 2 else 0,
                    "Car/Jeep/ Van": clean_numeric(df.iloc[r, 3]) if df.shape[1] > 3 else 0,
                    "Taxi/Ola/Uber": clean_numeric(df.iloc[r, 4]) if df.shape[1] > 4 else 0,
                    "Mini Bus": clean_numeric(df.iloc[r, 5]) if df.shape[1] > 5 else 0,
                    "Bus (Pvt)": clean_numeric(df.iloc[r, 6]) if df.shape[1] > 6 else 0,
                    "Bus (Gov)": clean_numeric(df.iloc[r, 7]) if df.shape[1] > 7 else 0,
                    "Mini LCV/ Tata Ace": clean_numeric(df.iloc[r, 8]) if df.shape[1] > 8 else 0,
                    "LCV(4/6-Wheels)": clean_numeric(df.iloc[r, 9]) if df.shape[1] > 9 else 0,
                    "Truck": (
                        (clean_numeric(df.iloc[r, 10]) if df.shape[1] > 10 else 0)
                        + (clean_numeric(df.iloc[r, 11]) if df.shape[1] > 11 else 0)
                        + (clean_numeric(df.iloc[r, 13]) if df.shape[1] > 13 else 0)
                    ),
                    "MAV (4-6 Axle)": clean_numeric(df.iloc[r, 12]) if df.shape[1] > 12 else 0,
                    "Garbage Vehicles": clean_numeric(df.iloc[r, 14]) if df.shape[1] > 14 else 0,
                    "Cycle": clean_numeric(df.iloc[r, 18]) if df.shape[1] > 18 else 0,
                    "Others": clean_numeric(df.iloc[r, 19]) if df.shape[1] > 19 else 0,
                }
            )

            if sum(row_record[c] for c in FINAL_COLUMNS[6:]) > 0:
                results.append(row_record)

    return results


def deduplicate_records(records):
    if not records:
        return pd.DataFrame(columns=FINAL_COLUMNS)

    df = pd.DataFrame(records)

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = 0 if col in FINAL_COLUMNS[6:] else ""

    df = df[FINAL_COLUMNS].drop_duplicates().reset_index(drop=True)
    return df


def process_file(file_path):
    all_rows = []

    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Failed to open {file_path.name}: {e}")
        return all_rows

    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            fname_lower = file_path.name.lower()

            if "01" in fname_lower:
                rows = parse_tvc01_style(df, file_path.name)
                if not rows:
                    rows = extract_dynamic(df, file_path.name)

            elif "07" in fname_lower:
                rows = parse_tvc07_style(df, file_path.name)
                if not rows:
                    rows = extract_dynamic(df, file_path.name)

            else:
                rows = extract_dynamic(df, file_path.name)
                if not rows:
                    rows = parse_tvc07_style(df, file_path.name)
                if not rows:
                    rows = parse_tvc01_style(df, file_path.name)

            all_rows.extend(rows)

        except Exception as e:
            print(f"  Sheet skipped: {sheet_name} | {e}")

    return all_rows


def main():
    excel_files = list(INPUT_DIR.glob("*.xlsx")) + list(INPUT_DIR.glob("*.xls"))

    if not excel_files:
        print("No Excel files found in data/input/")
        return

    all_records = []

    for file_path in excel_files:
        print(f"Processing: {file_path.name}")
        rows = process_file(file_path)
        print(f"  Extracted rows: {len(rows)}")
        all_records.extend(rows)

    master_df = deduplicate_records(all_records)

    if master_df.empty:
        print("No valid traffic data extracted.")
        return

    try:
        master_df["_sort_time"] = pd.to_datetime(
            master_df["Time period-from"],
            format="%H:%M",
            errors="coerce"
        )
        master_df = master_df.sort_values(
            by=["TVC", "Direction", "_sort_time"],
            ascending=[True, True, True]
        ).drop(columns=["_sort_time"]).reset_index(drop=True)
    except Exception:
        master_df = master_df.reset_index(drop=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        master_df.to_excel(writer, sheet_name="Master_data", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Master_data"]

        header_fmt = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "border": 1
        })

        cell_fmt = workbook.add_format({"border": 1})

        for col_num, value in enumerate(master_df.columns):
            worksheet.write(0, col_num, value, header_fmt)

        for idx, col in enumerate(master_df.columns):
            max_len = max(
                len(str(col)),
                master_df[col].astype(str).map(len).max() if not master_df.empty else 10
            )
            worksheet.set_column(idx, idx, max(14, min(28, max_len + 2)), cell_fmt)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(master_df), len(master_df.columns) - 1)

    print(f"\nOutput created successfully: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()