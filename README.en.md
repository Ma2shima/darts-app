# DartsApp — Darts Improvement Tracker

An app that lets a participant throw two rounds of darts (3 throws each) and quantifies their improvement by comparing the grouping radius between rounds.

---

## Screenshots

| Round 1 Input | Results Display |
|---|---|
| Click the dartboard to record each landing point | Shows grouping spread and accuracy improvement rate |

---

## Features

- **Click-to-record input** — Click the on-screen dartboard to log each dart's landing point
- **Miss support** — Clicking outside the board records a miss (fixed distance of 300 mm)
- **Auto round transition** — Automatically advances to Round 2 after 3 throws in Round 1
- **Edit recorded throws** — Click any recorded throw label to re-enter it
- **Grouping calculation** — Quantifies spread as the maximum distance from the centroid (mm)
- **Improvement rate** — Calculates `(Round 1 − Round 2) / Round 1 × 100%`
- **Excel export** — Saves coordinates, distances, and comparison results as a styled `.xlsx` file
- **Auto-reset after export** — Data is automatically cleared after saving to Excel
- **Language toggle** — Switch between Japanese and English with the header button

---

## Requirements

| Platform | Distribution file | Requirements |
|---|---|---|
| macOS 11 or later | `DartsApp_Mac.zip` | None (Python not required) |
| Windows 10 / 11 | `DartsApp_Windows.zip` | None (Python not required) |
| Python environment | `darts_app.py` | Python 3.10+, openpyxl |

---

## Installation & Launch

### Mac

1. Unzip `DartsApp_Mac.zip`
2. Double-click `DartsApp.app`

> **If you see a security warning on first launch:**
> "Apple cannot verify the developer" — Right-click (or Control+click) the app → **Open** → **Open**
> Or go to: System Settings → Privacy & Security → **Open Anyway**

### Windows

1. Unzip `DartsApp_Windows.zip`
2. Double-click `DartsApp.exe`

> **If Windows SmartScreen appears:**
> Click **More info** → **Run anyway**

### Run from Python (for developers)

```bash
pip install openpyxl
python darts_app.py
```

---

## How to Use

```
1. Enter the participant's name
        ↓
2. Round 1: Click the dartboard 3 times to record landing points
   (Click outside the board = Miss)
        ↓
3. Automatically advances to Round 2 (no button press needed)
        ↓
4. Round 2: Click the dartboard 3 times in the same way
        ↓
5. Results appear automatically in the right-hand panel
        ↓
6. Click "Export to Excel" to save the data
```

### Correcting a recorded throw

Click any recorded throw label (marked with `✎`) to enter edit mode.  
Then click the dartboard to overwrite that throw's data.  
Click the same label again to cancel without changes.

---

## Calculation Details

### Grouping Radius (Spread)

```
Centroid  = average position of all 3 throws
Grouping radius = maximum distance from centroid to any throw (mm)
```

Miss throws are included in the calculation as a point 300 mm from the bull's-eye.

### Improvement Rate

```
Improvement (%) = (Round 1 radius − Round 2 radius) / Round 1 radius × 100
```

Positive = improved; Negative = got worse.

### Scale

The outer edge of the scoring area (double ring) on screen equals 170 mm in real life.

---

## Build Instructions (for developers)

### Mac

```bash
pip install pyinstaller openpyxl
pyinstaller --windowed --name "DartsApp" --icon darts.icns -y darts_app.py
# Output: dist/DartsApp.app
```

### Windows

```powershell
pip install pyinstaller openpyxl
pyinstaller --windowed --name "DartsApp" --icon darts.ico --onefile -y darts_app.py
# Output: dist/DartsApp.exe
```

### GitHub Actions (build Mac & Windows simultaneously)

A push to the `main` branch triggers an automatic build for both platforms.  
See `.github/workflows/build.yml` for details.

---

## File Structure

```
darts_app/
├── darts_app.py          # Main application (Python)
├── darts.icns            # App icon (Mac)
├── darts.ico             # App icon (Windows)
├── .github/
│   └── workflows/
│       └── build.yml     # GitHub Actions build workflow
├── README.md             # Japanese README
└── README.en.md          # This file
```

---

## Dependencies

| Library | Purpose | Notes |
|---|---|---|
| tkinter | GUI framework | Python standard library |
| openpyxl | Excel export | `pip install openpyxl` |
| math / datetime | Calculations & timestamps | Python standard library |

---

## License

MIT License
