"""Run the BI automation workflow from the project root."""

from pathlib import Path
import runpy
import sys

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"

# The existing workflow imports database.py from src and settings.py from config.
for path in (str(PROJECT_ROOT), str(SRC_DIR)):
    if path not in sys.path:
        sys.path.insert(0, path)

runpy.run_path(str(SRC_DIR / "main.py"), run_name="__main__")
