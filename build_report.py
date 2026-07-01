"""Generate an interactive HTML career-timeline report from an HR history export.

The input is an Excel workbook (.xlsx or legacy .xls) containing one row per
personnel activity (change of manager, building, level, etc.). See README.md
for the expected column schema.
"""

import argparse
import json
import os
import sys
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

import pandas as pd


REQUIRED_COLUMNS = ["Activity Date", "Reports To Full Name", "Building Name"]


def read_history(path: Path) -> pd.DataFrame:
    """Read an xlsx or legacy xls (including OLE2 files misnamed .xlsx)."""
    errors = []
    for engine in ("openpyxl", "xlrd"):
        try:
            df = pd.read_excel(path, engine=engine)
            break
        except Exception as e:
            errors.append(f"{engine}: {e}")
    else:
        df = _read_via_excel_com(path, errors)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"Input is missing required columns: {missing}\n"
            f"Expected at least: {REQUIRED_COLUMNS}\n"
            f"Found: {list(df.columns)}"
        )

    df = df.dropna(subset=["Activity Date"]).copy()
    df["Date"] = pd.to_datetime(df["Activity Date"])
    df = df.sort_values("Date").reset_index(drop=True)
    return df


def _read_via_excel_com(path: Path, prior_errors: list) -> pd.DataFrame:
    """Last-resort Windows fallback: use Excel COM to export to CSV, then read it.

    Handles OLE2 workbooks that openpyxl and xlrd cannot parse.
    """
    if sys.platform != "win32":
        raise SystemExit(
            f"Could not read '{path}':\n  " + "\n  ".join(prior_errors)
        )
    try:
        import win32com.client  # type: ignore
    except ImportError:
        raise SystemExit(
            f"Could not read '{path}' with pandas engines:\n  "
            + "\n  ".join(prior_errors)
            + "\n\nOn Windows, install pywin32 to enable Excel COM fallback:\n"
            + "  pip install pywin32"
        )

    print(
        f"Note: '{path.name}' is an OLE2-encapsulated workbook that pandas cannot parse directly.\n"
        f"      Falling back to Excel COM to convert it. This can take 20-40 seconds...",
        file=sys.stderr,
        flush=True,
    )

    tmp_csv = path.with_suffix(".converted.csv")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow: suppress MOTW/Protected View prompts
    except Exception:
        pass
    try:
        wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        wb.Sheets(1).SaveAs(str(tmp_csv), 6)  # 6 = xlCSV
        wb.Close(False)
    finally:
        excel.Quit()
    df = pd.read_csv(tmp_csv)
    try:
        tmp_csv.unlink()
    except OSError:
        pass
    print("      Done reading via Excel COM.", file=sys.stderr, flush=True)
    return df


def collapse_runs(df: pd.DataFrame, column: str, today: pd.Timestamp):
    """Return contiguous (value, start, end) runs for a column, sorted by date."""
    runs = []
    prev_val = None
    run_start = None
    for _, row in df.iterrows():
        val = row[column]
        if pd.isna(val):
            continue
        if val != prev_val:
            if prev_val is not None:
                runs.append((prev_val, run_start, row["Date"]))
            prev_val = val
            run_start = row["Date"]
    if prev_val is not None:
        runs.append((prev_val, run_start, today))
    return runs


def totals(runs):
    out = OrderedDict()
    for name, start, end in runs:
        days = (end - start).days
        if name not in out:
            out[name] = {"days": 0, "first": start, "last": end}
        out[name]["days"] += days
        out[name]["first"] = min(out[name]["first"], start)
        out[name]["last"] = max(out[name]["last"], end)
    return out


def fmt_duration(days: int) -> str:
    years = days // 365
    months = (days % 365) // 30
    parts = []
    if years:
        parts.append(f"{years}y")
    if months:
        parts.append(f"{months}m")
    if not parts:
        parts.append(f"{days}d")
    return " ".join(parts)


def detect_employee_name(df: pd.DataFrame) -> str | None:
    for col in ("Full Name", "Employee Name", "Name"):
        if col in df.columns:
            values = df[col].dropna().unique()
            if len(values) == 1:
                return str(values[0])
            if len(values) > 1:
                return str(values[0])
    return None


def render_html(employee_name: str, df: pd.DataFrame, today: pd.Timestamp) -> str:
    mgr_runs = collapse_runs(df, "Reports To Full Name", today)
    bldg_runs = collapse_runs(df, "Building Name", today)

    mgr_totals = totals(mgr_runs)
    bldg_totals = totals(bldg_runs)

    def gantt(runs):
        return [
            {
                "name": str(name),
                "start": start.strftime("%Y-%m-%d"),
                "end": end.strftime("%Y-%m-%d"),
            }
            for name, start, end in runs
        ]

    def table_rows(t):
        return sorted(
            (
                {
                    "name": str(n),
                    "duration": fmt_duration(info["days"]),
                    "days": info["days"],
                    "first": info["first"].strftime("%Y-%m-%d"),
                    "last": info["last"].strftime("%Y-%m-%d"),
                }
                for n, info in t.items()
            ),
            key=lambda r: -r["days"],
        )

    start_date = df["Date"].min()
    total_years = (today - start_date).days / 365.25

    mgr_gantt = gantt(mgr_runs)
    bldg_gantt = gantt(bldg_runs)
    mgr_table = table_rows(mgr_totals)
    bldg_table = table_rows(bldg_totals)

    mgr_tr = "".join(
        f'<tr><td>{r["name"]}</td><td class="num">{r["duration"]}</td>'
        f'<td>{r["first"]}</td><td>{r["last"]}</td></tr>'
        for r in mgr_table
    )
    bldg_tr = "".join(
        f'<tr><td>{r["name"]}</td><td class="num">{r["duration"]}</td>'
        f'<td>{r["first"]}</td><td>{r["last"]}</td></tr>'
        for r in bldg_table
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Career Timeline &mdash; {employee_name}</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  * {{ box-sizing: border-box; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: #f6f8fa; color: #1f2328; margin: 0; padding: 32px;
    max-width: 1200px; margin-left: auto; margin-right: auto;
  }}
  h1 {{ font-size: 28px; margin: 0 0 4px 0; border-bottom: 2px solid #0969da; padding-bottom: 8px; }}
  .subtitle {{ color: #57606a; margin-bottom: 32px; }}
  .stats {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 32px; }}
  .stat {{ background: white; border: 1px solid #d0d7de; border-radius: 8px; padding: 16px; text-align: center; }}
  .stat .value {{ font-size: 28px; font-weight: 600; color: #0969da; }}
  .stat .label {{ font-size: 13px; color: #57606a; margin-top: 4px; }}
  h2 {{ font-size: 20px; margin: 32px 0 12px 0; }}
  .card {{ background: white; border: 1px solid #d0d7de; border-radius: 8px; padding: 20px; margin-bottom: 24px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
  th {{ text-align: left; background: #f6f8fa; border-bottom: 2px solid #d0d7de; padding: 8px 12px; font-weight: 600; }}
  td {{ padding: 8px 12px; border-bottom: 1px solid #eaeef2; }}
  tr:last-child td {{ border-bottom: none; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  .chart {{ width: 100%; height: 480px; }}
  .footer {{ color: #57606a; font-size: 12px; margin-top: 32px; text-align: center; }}
</style>
</head>
<body>

<h1>Career Timeline &mdash; {employee_name}</h1>
<div class="subtitle">Managers and buildings since {start_date.strftime('%B %Y')}</div>

<div class="stats">
  <div class="stat"><div class="value">{total_years:.1f}</div><div class="label">Years of history</div></div>
  <div class="stat"><div class="value">{len(mgr_totals)}</div><div class="label">Unique managers</div></div>
  <div class="stat"><div class="value">{len(bldg_totals)}</div><div class="label">Unique buildings</div></div>
  <div class="stat"><div class="value">{len(mgr_runs)}</div><div class="label">Manager changes</div></div>
</div>

<h2>Managers over time</h2>
<div class="card"><div id="mgr-chart" class="chart"></div></div>
<div class="card">
<table>
<thead><tr><th>Manager</th><th class="num">Time</th><th>First</th><th>Last</th></tr></thead>
<tbody>{mgr_tr}</tbody>
</table>
</div>

<h2>Buildings over time</h2>
<div class="card"><div id="bldg-chart" class="chart"></div></div>
<div class="card">
<table>
<thead><tr><th>Building</th><th class="num">Time</th><th>First</th><th>Last</th></tr></thead>
<tbody>{bldg_tr}</tbody>
</table>
</div>

<div class="footer">Generated {today.strftime('%Y-%m-%d')}</div>

<script>
const mgrData = {json.dumps(mgr_gantt)};
const bldgData = {json.dumps(bldg_gantt)};

function makeGantt(divId, data, colorBase) {{
  const ordered = data.slice().reverse();
  const traces = ordered.map((d) => ({{
    x: [d.start, d.end],
    y: [d.name, d.name],
    mode: 'lines',
    line: {{ width: 18, color: colorBase }},
    hovertemplate: `<b>${{d.name}}</b><br>${{d.start}} \u2192 ${{d.end}}<extra></extra>`,
    showlegend: false
  }}));
  const layout = {{
    margin: {{ l: 160, r: 30, t: 20, b: 40 }},
    xaxis: {{ type: 'date', gridcolor: '#eaeef2', fixedrange: true }},
    yaxis: {{ automargin: true, tickfont: {{ size: 12 }}, fixedrange: true }},
    plot_bgcolor: 'white', paper_bgcolor: 'white', hovermode: false
  }};
  Plotly.newPlot(divId, traces, layout, {{staticPlot: true, responsive: true}});
}}

makeGantt('mgr-chart', mgrData, '#0969da');
makeGantt('bldg-chart', bldgData, '#1f883d');
</script>
</body>
</html>
"""


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Generate an HTML career-timeline report from an HR history export (.xlsx / .xls).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "-i", "--input",
        default="sample_career.xlsx",
        help="Path to the HR history export workbook.",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Path to write the HTML report. Defaults to <input-stem>_report.html next to the input.",
    )
    parser.add_argument(
        "-n", "--name",
        default=None,
        help="Employee name to show in the report header. If omitted, the script tries to detect it from a 'Full Name', 'Employee Name', or 'Name' column.",
    )
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2

    df = read_history(input_path)

    name = args.name or detect_employee_name(df) or "Employee"

    output_path = (
        Path(args.output).resolve()
        if args.output
        else input_path.with_name(f"{input_path.stem}_report.html")
    )

    today = pd.Timestamp(datetime.now().date())
    html = render_html(name, df, today)
    output_path.write_text(html, encoding="utf-8")

    print(f"Employee: {name}")
    print(f"History rows: {len(df)}")
    print(f"Report: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
