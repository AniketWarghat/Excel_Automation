import subprocess
import sys
from pathlib import Path

app_path = Path("KMZ img embed") / "KMZ_img_embed.py"

subprocess.run([
    sys.executable,
    "-m",
    "streamlit",
    "run",
    str(app_path)
])