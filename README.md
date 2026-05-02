# Water Supply Pipe Sizing Tool

An automated Excel workbook for hydraulic pipe sizing of rural water supply systems, built with Python and openpyxl. Implements the **Hazen-Williams** equation per the **Uganda Water Supply Design Manual (2nd Edition)** using **HDPE PE100 pipes (DIN 8074, Gentex catalog)**.

---

## Features

- **Auto pipe selection** — smallest OD from catalog where internal diameter ≥ D_req, using Excel `AGGREGATE` formula
- **Auto pressure rating (PN)** — PN-6 / PN-10 / PN-16 / PN-20 selected from max working pressure with safety factor
- **TDH calculation** — production well → storage reservoir, including static head, friction, minor losses, and velocity head
- **Pump sizing** — hydraulic power, motor power (input efficiencies), daily volume supplied
- **HGL chaining** — hydraulic grade line propagates automatically through transmission segments
- **Live status flags** — ✓ OK / ⚠ LOW VELOCITY / ⚠ HIGH VELOCITY / ⚠ LOW PRESSURE / ⚠ SIZE N/A per segment
- **Node pressure summary** — 10-node check table in distribution sheet
- **Color-coded cells** — yellow = user input, black = formula, green = cross-sheet link

---

## Workbook Sheets

| Sheet | Purpose |
|---|---|
| `COVER` | Project info, auto-populated design summary, navigation guide |
| `PARAMETERS` | All design constants — H-W C, velocities, PN ratings, demand factors |
| `PIPE_DB` | Full Gentex HDPE PE100 catalog — 74 entries, OD 20–450 mm, PN-6 to PN-20 |
| `TDH_CALC` | Total Dynamic Head: 12 rising main segments + pump requirements |
| `TRANSMISSION` | Transmission main: 15 segments, auto OD/PN, HGL profile |
| `DISTRIBUTION` | Distribution network: demand calc, 20 branches, node pressure checks |

---

## Design Standards

- Uganda Water Supply Design Manual, 2nd Edition
- Hazen-Williams equation: `hf = 10.67 · L · Q^1.852 / (C^1.852 · D^4.87)`
- HDPE PE100 pipe catalog: Gentex DIN 8074 (price list effective July 2018)
- Minimum residual pressure: 7 m (distribution consumer), 2 m (reservoir/transmission end)
- Allowable head loss gradient: 5 m/km (transmission), 8 m/km (distribution)
- Flow velocities: 0.5–2.5 m/s (transmission), 0.3–1.5 m/s (distribution)

---

## How to Run (Google Colab)

### 1. Install dependencies
```python
!pip install openpyxl xlrd
```

### 2. Upload the pipe catalog
Upload `HDPE_Pipes_-Gentex_1_.xls` using the Colab Files panel (left sidebar).

### 3. Update file paths at the top of Stage 1
```python
HDPE_FILE   = '/content/HDPE_Pipes_-Gentex_1_.xls'
OUTPUT_FILE = '/content/Uganda_Pipe_Sizing_Tool.xlsx'
```

### 4. Run stages in order
Each stage is one Colab cell. All stages share the same Python kernel session.

| Stage | File | Builds |
|---|---|---|
| 1 | `stage1.py` | Imports, colour constants, style helpers, pipe data, formula builders |
| 2 | `stage2.py` | COVER, PARAMETERS, PIPE_DB sheets |
| 3 | `stage3.py` | TDH_CALC sheet |
| 4 | `stage4.py` | TRANSMISSION sheet |
| 5 | `stage5.py` | DISTRIBUTION sheet + saves workbook |

### 5. Download output
Files panel → right-click `Uganda_Pipe_Sizing_Tool.xlsx` → **Download**

---

## How to Use the Workbook

1. **PARAMETERS** — review and adjust defaults (velocities, pressure ratings, demand factors)
2. **TDH_CALC** — enter well survey data and reservoir elevations → TDH and pump power auto-calculated
3. **TRANSMISSION** — enter segment lengths, elevations, and flow → OD and PN auto-selected
4. **DISTRIBUTION** — enter population, reservoir HWL, and branch data → node pressures checked

### Branching in TRANSMISSION / DISTRIBUTION
For branches, manually set the **HGL Start** cell of the branch segment to reference the junction node's HGL End column:
```
=T16   ← links to HGL End of row 16 (the junction)
```

---

## Requirements

- Python 3.8+
- `openpyxl >= 3.0`
- `xlrd >= 2.0`
- `pandas`
- Microsoft Excel 2010+ (uses `AGGREGATE` and `SUMPRODUCT` functions)

---

## Project Structure

```
water-supply-pipe-sizing-tool/
├── stage1.py                        # Setup, helpers, pipe data, formula builders
├── stage2.py                        # COVER, PARAMETERS, PIPE_DB sheets
├── stage3.py                        # TDH_CALC sheet
├── stage4.py                        # TRANSMISSION sheet
├── stage5.py                        # DISTRIBUTION sheet + save
├── Uganda_Pipe_Sizing_Tool.xlsx     # Pre-built Excel workbook (ready to use)
└── README.md
```

> **Note:** The Gentex pipe catalog file (`HDPE_Pipes_-Gentex_1_.xls`) is not included in this repository. You will need to supply your own copy and upload it to Colab before running the stages.

---

## License

This project is released for use in water supply engineering projects. Attribution appreciated.
