from pathlib import Path
import subprocess
import sys
import pandas as pd
import streamlit as st

# -------------------------------------------------
# Paths
# -------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"
CLEANED_FILE = OUTPUT_DIR / "Cleaned_Data.xlsx"
REPORT_FILE = OUTPUT_DIR / "Report.xlsx"
REPORT_SCRIPT = BASE_DIR / "excel_sheet_create.py"

INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# -------------------------------------------------
# Page config
# -------------------------------------------------
st.set_page_config(
    page_title="Excel Automation Dashboard",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------
# Helper functions
# -------------------------------------------------
def get_existing_input_files():
    return sorted([f for f in INPUT_DIR.iterdir() if f.is_file()])


def get_unique_file_path(file_name: str) -> Path:
    target = INPUT_DIR / file_name
    if not target.exists():
        return target

    stem = target.stem
    suffix = target.suffix
    counter = 1

    while True:
        new_target = INPUT_DIR / f"{stem}_{counter}{suffix}"
        if not new_target.exists():
            return new_target
        counter += 1


def save_uploaded_files(uploaded_files):
    saved_files = []
    for uploaded_file in uploaded_files:
        file_path = get_unique_file_path(uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        saved_files.append(file_path.name)
    return saved_files


def clear_input_folder():
    for file in INPUT_DIR.iterdir():
        if file.is_file():
            file.unlink()


def run_data_cleaning():
    return subprocess.run(
        [sys.executable, str(BASE_DIR / "data_cleaning.py")],
        capture_output=True,
        text=True,
        cwd=str(BASE_DIR)
    )


def run_report_generation():
    return subprocess.run(
        [sys.executable, str(REPORT_SCRIPT)],
        capture_output=True,
        text=True,
        cwd=str(BASE_DIR)
    )


def read_tvc_list_from_cleaned_file():
    if not CLEANED_FILE.exists():
        return None, f"{CLEANED_FILE.name} not found in data/output/."

    try:
        df = pd.read_excel(CLEANED_FILE)

        normalized_cols = {str(col).strip().lower(): col for col in df.columns}

        if "tvc" not in normalized_cols:
            return None, f"TVC column not found in {CLEANED_FILE.name}. Available columns: {list(df.columns)}"

        tvc_col = normalized_cols["tvc"]

        tvc_values = (
            df[tvc_col]
            .dropna()
            .astype(str)
            .str.strip()
        )

        tvc_values = tvc_values[tvc_values != ""]
        unique_tvc = sorted(tvc_values.unique().tolist())

        return unique_tvc, None

    except Exception as e:
        return None, f"Error reading {CLEANED_FILE.name}: {e}"


# -------------------------------------------------
# Minimal CSS
# -------------------------------------------------
st.markdown("""
<style>
    .block-container {
        max-width: 1000px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.25rem;
    }
    .sub-title {
        color: #6b7280;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# Header
# -------------------------------------------------
st.markdown('<div class="main-title">Excel Automation Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Upload files, manage input files, run data cleaning, generate report, and review TVCs from Cleaned_Data.xlsx.</div>',
    unsafe_allow_html=True
)

# -------------------------------------------------
# Metrics
# -------------------------------------------------
input_files = get_existing_input_files()

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Input files available", len(input_files))
with col2:
    st.metric("Cleaned file exists", "Yes" if CLEANED_FILE.exists() else "No")
with col3:
    st.metric("Report file exists", "Yes" if REPORT_FILE.exists() else "No")

st.divider()

# -------------------------------------------------
# Upload files
# -------------------------------------------------
st.subheader("Upload Input Excel Files")
st.write("Select one or more Excel files to store in `data/input/`.")

uploaded_files = st.file_uploader(
    "Choose Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files and st.button("Upload files", use_container_width=True):
    saved_files = save_uploaded_files(uploaded_files)
    st.success("Files uploaded successfully.")
    st.write("Saved files:")
    for name in saved_files:
        st.write(f"- {name}")

st.divider()

# -------------------------------------------------
# Current input files
# -------------------------------------------------
st.subheader("Current Input Files")
input_files = get_existing_input_files()

if input_files:
    for file in input_files:
        st.write(f"- {file.name}")
else:
    st.info("No files currently available in data/input/.")

if st.button("Clear all input files", use_container_width=True):
    clear_input_folder()
    st.warning("All files removed from data/input/.")
    st.rerun()

st.divider()

# -------------------------------------------------
# Run cleaning
# -------------------------------------------------
st.subheader("Run Data Cleaning")
st.write("Run `data_cleaning.py` after uploading files.")

if st.button("Run data cleaning", use_container_width=True):
    with st.spinner("Running data cleaning..."):
        try:
            result = run_data_cleaning()
            if result.returncode == 0:
                st.success("Data cleaning completed successfully.")
                if result.stdout.strip():
                    st.code(result.stdout, language="bash")
            else:
                st.error("Data cleaning failed.")
                if result.stderr.strip():
                    st.code(result.stderr, language="bash")
        except Exception as e:
            st.error(f"Failed to run data_cleaning.py: {e}")

st.divider()

# -------------------------------------------------
# TVC analysis
# -------------------------------------------------
st.subheader("TVC Analysis")
st.write("Read `data/output/Cleaned_Data.xlsx`, count unique TVCs, and list them.")

if st.button("Load TVCs from Cleaned_Data.xlsx", use_container_width=True):
    tvc_list, error = read_tvc_list_from_cleaned_file()

    if error:
        st.error(error)
    else:
        st.success(f"Found {len(tvc_list)} unique TVC value(s).")

        if tvc_list:
            st.write("TVC list:")
            for i, tvc in enumerate(tvc_list, start=1):
                st.write(f"{i}. {tvc}")
        else:
            st.info("No TVC values found in Cleaned_Data.xlsx.")

st.divider()

# -------------------------------------------------
# Generate report
# -------------------------------------------------
st.subheader("Generate Report")
st.write("Generate `Report.xlsx` from `data/output/Cleaned_Data.xlsx` and download it.")

col1, col2 = st.columns(2)

with col1:
    if st.button("Generate report", use_container_width=True):
        with st.spinner("Generating report..."):
            try:
                if not CLEANED_FILE.exists():
                    st.error(f"{CLEANED_FILE.name} not found in data/output/. Please run data cleaning first.")
                elif not REPORT_SCRIPT.exists():
                    st.error(f"{REPORT_SCRIPT.name} not found in project folder.")
                else:
                    result = run_report_generation()
                    if result.returncode == 0:
                        st.success("Report generated successfully.")
                        if result.stdout.strip():
                            st.code(result.stdout, language="bash")
                    else:
                        st.error("Report generation failed.")
                        if result.stderr.strip():
                            st.code(result.stderr, language="bash")
            except Exception as e:
                st.error(f"Failed to generate report: {e}")

with col2:
    if REPORT_FILE.exists():
        with open(REPORT_FILE, "rb") as f:
            st.download_button(
                label="Download Report.xlsx",
                data=f,
                file_name="Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("Report.xlsx is not available yet.")