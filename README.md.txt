@'
# RGGP7 – R&D Final Report Builder (Tkinter)

**What it does:** Windows desktop app to generate R&D Final Reports with our standard sections and formatting.

## How to run (end users without Python)
1. Download the latest **RGGP7.exe** from the **Releases** page.
2. Double-click `RGGP7.exe` to launch the app (no Python needed).

## How to build the EXE (developers)
- Install Python 3.11+
- `pip install -r requirements.txt`
- `pyinstaller --onefile --windowed --add-data "assets;assets" RGGP7.py`
- Output EXE appears in `dist/`.

## Folder structure
