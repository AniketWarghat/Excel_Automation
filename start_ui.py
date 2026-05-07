import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

subprocess.run([
    sys.executable,
    "-m",
    "streamlit",
    "run",
    str(BASE_DIR / "ui.py"),
    "--server.port", "5500",
    "--server.address", "0.0.0.0"
], cwd=BASE_DIR)