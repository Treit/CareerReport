# CareerReport

Generate an HTML career-timeline report from an HR-history spreadsheet export. The report shows every manager you've reported to and every building you've worked in as Gantt-style bars, plus summary tables ranked by total time.

## Install

Requires Python 3.10+.

```
pip install pandas openpyxl plotly
```

For legacy `.xls` (OLE2) inputs, also install `xlrd`:

```
pip install xlrd
```

On Windows, for HR exports that use an unusual OLE2 encapsulation (some corporate HR systems produce these), install `pywin32` to enable an Excel COM fallback:

```
pip install pywin32
```

## Usage

```
python build_report.py -i my_history.xlsx
python build_report.py -i my_history.xlsx -n "Jane Smith" -o jane.html
python build_report.py --help
```

Options:

| Flag | Description |
|------|-------------|
| `-i`, `--input` | Path to the history export (`.xlsx` or `.xls`). Default: `sample_career.xlsx`. |
| `-o`, `--output` | Path for the HTML report. Default: `<input-stem>_report.html` next to the input. |
| `-n`, `--name` | Employee name for the report header. If omitted, auto-detected from a `Full Name`, `Employee Name`, or `Name` column. |
| `-h`, `--help` | Show the standard help. |

The report is a single self-contained HTML file. It loads Plotly from a CDN so opening it requires internet access.

## Input schema

The script expects an Excel workbook whose first sheet contains one row per personnel activity. Required columns:

| Column | Type | Purpose |
|--------|------|---------|
| `Activity Date` | date | When the change happened. Sorted ascending to reconstruct history. |
| `Reports To Full Name` | text | Manager on the given date. Used to build the manager timeline. |
| `Building Name` | text | Building on the given date. Used to build the building timeline. |

Recommended (used when present):

| Column | Purpose |
|--------|---------|
| `Full Name` / `Employee Name` / `Name` | Auto-fills the report header. |

Any additional columns (e.g. `Personnel Number`, `Alias`, `Activity Detail`, `Position Number`, `Area`, `City`, `Work Country`) are ignored. This matches the standard SAP-HR-style export format used by many large employers.

## Try the sample

A fictional John Doe career is bundled for testing. Regenerate it any time:

```
python generate_sample.py
python build_report.py -i sample_career.xlsx
```

Then open `sample_career_report.html` in a browser.

## Privacy note

The script itself and the schema above are generic and safe to share. **Real HR exports may contain data your employer considers internal** (building codes, cost centers, org structure, colleague names). Review any real report before publishing screenshots or hosting it publicly.

## License

MIT
