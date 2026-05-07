from pathlib import Path
import streamlit as st

INPUT_DIR = Path("data/input")
INPUT_DIR.mkdir(parents=True, exist_ok=True)


def get_existing_files():
    return sorted([f for f in INPUT_DIR.iterdir() if f.is_file()])


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


st.set_page_config(page_title="Input File Upload", layout="centered")
st.title("Upload Input Excel Files")
st.write("Select one or more Excel files to store in `data/input/`.")

uploaded_files = st.file_uploader(
    "Choose Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    saved_files = save_uploaded_files(uploaded_files)
    st.success("Files uploaded successfully.")

    st.write("Saved files:")
    for name in saved_files:
        st.write(f"- {name}")

st.subheader("Current files in data/input/")
existing_files = get_existing_files()

if existing_files:
    for file in existing_files:
        st.write(f"- {file.name}")
else:
    st.info("No files currently available in data/input/.")

if st.button("Clear all files from input folder"):
    clear_input_folder()
    st.warning("All files removed from data/input/.")
    st.rerun()

#use the following command to run the Streamlit app:
#python -m streamlit run input.py --server.port 8501 --server.address 0.0.0.0