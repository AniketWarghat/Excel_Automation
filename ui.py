from pathlib import Path
import subprocess
import sys
import streamlit as st

# -------------------------------------------------
# Paths
# -------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

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


def get_existing_output_files():
    return sorted([f for f in OUTPUT_DIR.iterdir() if f.is_file()])


def get_unique_file_path(file_name: str) -> Path:
    target = INPUT_DIR / file_name

    if not target.exists():
        return target

    stem = target.stem
    suffix = target.suffix
    counter = 1

    while True:
        new_name = f"{stem}_{counter}{suffix}"
        new_target = INPUT_DIR / new_name
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
    try:
        result = subprocess.run(
            [sys.executable, str(BASE_DIR / "data_cleaning.py")],
            capture_output=True,
            text=True,
            cwd=str(BASE_DIR)
        )
        return result
    except Exception as e:
        return e


# -------------------------------------------------
# Custom CSS
# -------------------------------------------------
st.markdown("""
<style>
    .main-title {
        font-size: 34px;
        font-weight: 700;
        color: #16324f;
        margin-bottom: 8px;
    }
    .sub-title {
        font-size: 16px;
        color: #5c6b7a;
        margin-bottom: 24px;
    }
    .card {
        background-color: #f8fbff;
        border: 1px solid #d7e3f1;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        margin-bottom: 16px;
    }
    .card-title {
        font-size: 22px;
        font-weight: 600;
        color: #12344d;
        margin-bottom: 10px;
    }
    .card-text {
        font-size: 15px;
        color: #536271;
        margin-bottom: 18px;
        line-height: 1.6;
    }
    .metric-box {
        background-color: #ffffff;
        border: 1px solid #e1e8f0;
        border-radius: 14px;
        padding: 10px 16px;
        text-align: center;
    }
    .small-label {
        color: #5c6b7a;
        font-size: 14px;
    }
    .big-number {
        color: #16324f;
        font-size: 28px;
        font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# Header
# -------------------------------------------------
st.markdown('<div class="main-title">Excel Automation Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Upload Excel files, manage input files, run data cleaning, and download output files — all from one page.</div>',
    unsafe_allow_html=True
)

# -------------------------------------------------
# Metrics
# -------------------------------------------------
input_files = get_existing_input_files()
output_files = get_existing_output_files()

m1, m2 = st.columns(2)

with m1:
    st.markdown(
        f"""
        <div class="metric-box">
            <div class="small-label">Input files available</div>
            <div class="big-number">{len(input_files)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with m2:
    st.markdown(
        f"""
        <div class="metric-box">
            <div class="small-label">Output files available</div>
            <div class="big-number">{len(output_files)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# -------------------------------------------------
# Upload + Cleaning sections
# -------------------------------------------------
col1, col2 = st.columns(2)

# ---------------- Upload Section ----------------
with col1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Step 1: Upload Input Excel Files</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="card-text">Select one or more Excel files and store them in <b>data/input/</b>.</div>',
        unsafe_allow_html=True
    )

    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("Upload Files", use_container_width=True):
            saved_files = save_uploaded_files(uploaded_files)
            st.success("Files uploaded successfully.")

            st.write("Saved files:")
            for name in saved_files:
                st.write(f"- {name}")

    st.markdown("### Current files in data/input/")
    input_files = get_existing_input_files()

    if input_files:
        for file in input_files:
            st.write(f"- {file.name}")
    else:
        st.info("No files currently available in data/input/.")

    if st.button("Clear all files from input folder", use_container_width=True):
        clear_input_folder()
        st.warning("All files removed from data/input/.")
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- Cleaning Section ----------------
with col2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Step 2: Data Cleaning & Output</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="card-text">Run <b>data_cleaning.py</b> after uploading files. Cleaned files will be generated in <b>data/output/</b>.</div>',
        unsafe_allow_html=True
    )

    if st.button("Run Data Cleaning", use_container_width=True):
        with st.spinner("Running data cleaning..."):
            result = run_data_cleaning()

        if isinstance(result, Exception):
            st.error(f"Failed to run data_cleaning.py: {result}")
        else:
            if result.returncode == 0:
                st.success("Data cleaning completed successfully.")

                if result.stdout.strip():
                    st.markdown("### Process log")
                    st.code(result.stdout, language="bash")
            else:
                st.error("Data cleaning failed.")
                if result.stderr.strip():
                    st.markdown("### Error log")
                    st.code(result.stderr, language="bash")

    st.markdown("### Current files in data/output/")
    output_files = get_existing_output_files()

    if output_files:
        for file in output_files:
            st.write(f"- {file.name}")

        st.markdown("### Download output files")
        for file in output_files:
            try:
                with open(file, "rb") as f:
                    st.download_button(
                        label=f"Download {file.name}",
                        data=f,
                        file_name=file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"download_{file.name}"
                    )
            except Exception as e:
                st.warning(f"Could not prepare download for {file.name}: {e}")
    else:
        st.info("No files currently available in data/output/.")

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------------------------------
# Footer
# -------------------------------------------------
st.markdown("---")
st.subheader("Workflow")
st.markdown("""
- Upload one or more Excel files in **Step 1**
- Confirm they appear in **data/input/**
- Click **Run Data Cleaning** in **Step 2**
- Download the generated files from **data/output/**
""")