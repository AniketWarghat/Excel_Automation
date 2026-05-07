import os
from pathlib import Path

INPUT_DIR = Path("data/input")


def ensure_input_dir():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)


def get_input_files():
    ensure_input_dir()
    files = list(INPUT_DIR.glob("*.xlsx")) + list(INPUT_DIR.glob("*.xls"))
    return files


def main():
    ensure_input_dir()
    files = get_input_files()

    if not files:
        print(f"No Excel files found in: {INPUT_DIR}")
        print("Please upload or copy input files into the data/input folder.")
        return

    print("Input files found:")
    for file in files:
        print(file)


if __name__ == "__main__":
    main()