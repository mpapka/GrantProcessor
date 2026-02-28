# Grant Expenditure Processor

Tools for parsing university grant expenditure data and generating interactive
visualizations. Built for quarterly COE (College of Engineering) financial
reports, but adaptable to any institution that follows the same spreadsheet
layout.

**Author:** Michael Papka

## Overview

| Script | Purpose |
|--------|---------|
| `grantProcessor.py` | Parse raw quarterly Excel files into clean per-department or cross-department summaries with indirect cost distributions |
| `plotGrantSummary.py` | Generate static plots (PNG/PDF) and an interactive D3.js dashboard from the combined summary |

## Requirements

- Python 3.8+
- `openpyxl` — Excel reading/writing
- `matplotlib` — static chart generation (PNG/PDF)
- `plotly` — standalone interactive HTML charts

```bash
pip install openpyxl matplotlib plotly
```

The interactive dashboard (`dashboard.html`) uses **D3.js v7** loaded from CDN
and has no additional Python dependencies.

## Quick Start

An example input file (`Example-Q4AY25.xlsx`) with fictitious data is included
so you can try the full pipeline immediately:

```bash
# Step 1: Process the raw input into a combined summary
python grantProcessor.py Example-Q4AY25.xlsx -d all -q full-year --combined \
    -o Example-Combined.xlsx

# Step 2: Generate the interactive dashboard
python plotGrantSummary.py Example-Combined.xlsx --format dashboard

# Step 3: Open in your browser (no web server needed)
open dashboard.html        # macOS
xdg-open dashboard.html    # Linux
start dashboard.html       # Windows
```

## Source Data Format

The input file contains a single sheet organized into department sections:

- **Department headers** — ALL-CAPS text in column B (e.g., `COMPUTER SCIENCE`)
- **Investigator rows** — name in column B, grant expenditures in columns G-J:
  - Column G: Q1 (Jul-Sep)
  - Column H: Q2 (Oct-Dec)
  - Column I: Q3 (Jan-Mar)
  - Column J: Q4 (Apr-Jun)
- **Subtotal rows** — column B empty, appear after multi-grant investigators
- **FY summary rows** — `FY25`, `FY26`, etc. mark the end of each section

---

## grantProcessor.py

### Usage

```
python grantProcessor.py <input_file> [options]
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `-d`, `--department` | `COMPUTER SCIENCE` | Department name (case-insensitive) or `all` |
| `-q`, `--quarter` | `Q1` | `Q1`, `Q2`, `Q3`, `Q4`, or `full-year` |
| `-o`, `--output` | auto-generated | Output file path |
| `-c`, `--combined` | off | Single-sheet mode with Department column (implies `all`) |
| `--indirect-rate` | `0.56` | F&A indirect cost rate |
| `--campus-rate` | `0.448` | Campus share of indirect |
| `--ovcr-rate` | `0.07` | OVCR share of indirect |
| `--coe-rate` | `0.275` | COE share of indirect |
| `--dept-rate` | `0.20` | Department share of indirect |

### Output Columns

| Column | Description |
|--------|-------------|
| Department | Abbreviation (combined mode only) |
| Name | Investigator name |
| Funds | Total expenditures for the selected quarter(s) |
| Percentage | Share of total (department or college-wide) |
| Indirect | `Funds * indirect_rate` |
| Campus / OVCR / COE / Dept | Splits of the indirect amount |
| Q1-Q4 Funds | Per-quarter breakdown (full-year combined mode only) |

### Examples

```bash
# Single department, single quarter (defaults)
python grantProcessor.py COE-Q1AY26.xlsx
# -> CS-Expenditures-Q1.xlsx

# Specific department and quarter
python grantProcessor.py COE-Q1AY26.xlsx -d "CHEMICAL ENGINEERING" -q Q2
# -> ChE-Expenditures-Q2.xlsx

# All departments, one sheet each
python grantProcessor.py COE-Q1AY26.xlsx -d all -q Q1
# -> COE-Expenditures-Q1.xlsx (sheets: ERC, MIE, BME, ...)

# Full-year combined (all departments, single sheet, includes Q1-Q4 columns)
python grantProcessor.py COE-Q4AY25.xlsx -d all -q full-year --combined
# -> COE-Combined-full-year.xlsx

# Custom indirect rates
python grantProcessor.py COE-Q1AY26.xlsx --indirect-rate 0.60 --campus-rate 0.45
```

### Department Abbreviations

New departments are auto-detected from the input file. Built-in abbreviations:

| Full Name | Abbreviation |
|-----------|-------------|
| COMPUTER SCIENCE | CS |
| MECHANICAL AND INDUSTRIAL ENGINEERING | MIE |
| BIOMEDICAL ENGINEERING | BME |
| CHEMICAL ENGINEERING | ChE |
| CIVIL AND MATERIAL ENGINEERING | CME |
| ELECTRICAL AND COMPUTER ENGINEERING | ECE |
| ENERGY RESOURCE CENTER | ERC |

ADMINISTRATION and COLLEGE TOTAL sections are automatically skipped.

---

## plotGrantSummary.py

### Usage

```
python plotGrantSummary.py <combined_file> [options]
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--top` | `20` | Number of investigators in the top-N chart |
| `--format` | `all` | `png`, `pdf`, `html`, `dashboard`, or `all` |
| `-o`, `--output-dir` | `.` | Directory for output files |

### Generated Outputs

| Output | Description |
|--------|-------------|
| `topInvestigators.png` | Horizontal bar chart of top-N investigators |
| `deptTotals.png` | Vertical bar chart of department totals |
| `indirectBreakdown.png` | Stacked bar chart of indirect cost splits |
| `summary.pdf` | All three charts in one multi-page PDF |
| `*.html` (individual) | Interactive Plotly versions of each chart |
| `dashboard.html` | Self-contained D3.js interactive dashboard |

### Interactive Dashboard

The dashboard is a single HTML file with no server requirements. It loads
D3.js v7 from CDN and embeds all data inline as JSON. Features:

- **Three tabbed views**: Top Investigators, Department Totals, Indirect Breakdown
- **Top-N selector**: 5 / 10 / 15 / 20 / 30 / All
- **Sort toggle**: by Funds or by Department
- **Department filter**: checkboxes to show/hide departments
- **Quarter selector**: when the combined file includes per-quarter data
  (generated with `-q full-year`), a dropdown lets you view Full Year,
  Q1, Q2, Q3, or Q4 individually
- **Animated transitions**: bars animate smoothly when toggling departments,
  changing top-N, or switching sort order
- **Hover tooltips**: exact dollar amounts and investigator details
- **Department Totals**: each department bar is stacked by individual
  investigator with shaded gradients; negative corrections (if any) appear
  as red segments
- **Statistics row**: PI count, average, and median displayed below each
  department bar
- **Responsive**: charts redraw on browser resize

### Examples

```bash
# Full pipeline: process + dashboard
python grantProcessor.py COE-Q4AY25.xlsx -d all -q full-year --combined
python plotGrantSummary.py COE-Combined-full-year.xlsx --format dashboard
open dashboard.html

# All output formats
python plotGrantSummary.py COE-Combined-Q1.xlsx

# Top 10 investigators, PNG only
python plotGrantSummary.py COE-Combined-Q1.xlsx --top 10 --format png

# Dashboard to a specific directory
python plotGrantSummary.py COE-Combined-Q1.xlsx --format dashboard -o reports/
```

---

## License

MIT
