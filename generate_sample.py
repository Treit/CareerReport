"""Generate a fictional sample_career.xlsx for testing / demoing build_report.py.

The data is entirely made up. Names, buildings, and cost centers are fictional.
"""

from pathlib import Path

import pandas as pd


ROWS = [
    ("6/15/2025", "Current Row",           "Priya Ramanathan", "TOWER-1",   "Portland"),
    ("3/1/2025",  "Change of Building",    "Priya Ramanathan", "TOWER-1",   "Portland"),
    ("9/1/2024",  "Review",                "Priya Ramanathan", "RIVERSIDE", "Portland"),
    ("8/30/2024", "Stock Grant",           "Priya Ramanathan", "RIVERSIDE", "Portland"),
    ("2/1/2024",  "Change in Reports To",  "Priya Ramanathan", "RIVERSIDE", "Portland"),
    ("9/1/2023",  "Change of Pay",         "Marcus Alvarado",  "RIVERSIDE", "Portland"),
    ("9/1/2023",  "Review",                "Marcus Alvarado",  "RIVERSIDE", "Portland"),
    ("5/12/2023", "Change of Building",    "Marcus Alvarado",  "RIVERSIDE", "Portland"),
    ("1/10/2023", "Change in Reports To",  "Marcus Alvarado",  "HARBOR-3",  "Seattle"),
    ("9/1/2022",  "Review",                "Elena Kowalski",   "HARBOR-3",  "Seattle"),
    ("8/31/2022", "Stock Grant",           "Elena Kowalski",   "HARBOR-3",  "Seattle"),
    ("6/1/2021",  "Change of Building",    "Elena Kowalski",   "HARBOR-3",  "Seattle"),
    ("9/1/2020",  "Change of Pay",         "Elena Kowalski",   "HARBOR-1",  "Seattle"),
    ("9/1/2020",  "Review",                "Elena Kowalski",   "HARBOR-1",  "Seattle"),
    ("4/15/2020", "Change in Reports To",  "Elena Kowalski",   "HARBOR-1",  "Seattle"),
    ("9/1/2019",  "Review",                "Devon Fitzgerald", "HARBOR-1",  "Seattle"),
    ("9/1/2018",  "Change of Level",       "Devon Fitzgerald", "HARBOR-1",  "Seattle"),
    ("9/1/2018",  "Change of Pay",         "Devon Fitzgerald", "HARBOR-1",  "Seattle"),
    ("2/1/2018",  "Change of Building",    "Devon Fitzgerald", "HARBOR-1",  "Seattle"),
    ("9/1/2017",  "Review",                "Devon Fitzgerald", "PIONEER-2", "Seattle"),
    ("11/3/2016", "Change in Reports To",  "Devon Fitzgerald", "PIONEER-2", "Seattle"),
    ("9/1/2016",  "Change of Pay",         "Rina Bergstrom",   "PIONEER-2", "Seattle"),
    ("9/1/2015",  "Review",                "Rina Bergstrom",   "PIONEER-2", "Seattle"),
    ("6/1/2015",  "Change of Building",    "Rina Bergstrom",   "PIONEER-2", "Seattle"),
    ("9/1/2014",  "Review",                "Rina Bergstrom",   "PIONEER-A", "Seattle"),
    ("3/10/2013", "Change in Reports To",  "Rina Bergstrom",   "PIONEER-A", "Seattle"),
    ("9/1/2012",  "Change of Pay",         "Hiroshi Tanaka",   "PIONEER-A", "Seattle"),
    ("9/1/2011",  "Review",                "Hiroshi Tanaka",   "PIONEER-A", "Seattle"),
    ("8/15/2010", "Change of Building",    "Hiroshi Tanaka",   "PIONEER-A", "Seattle"),
    ("9/1/2009",  "Change of Pay",         "Hiroshi Tanaka",   "CENTRAL-4", "Seattle"),
    ("6/1/2008",  "Change in Reports To",  "Hiroshi Tanaka",   "CENTRAL-4", "Seattle"),
    ("9/1/2007",  "Review",                "Anna Petrova",     "CENTRAL-4", "Seattle"),
    ("4/2/2006",  "Change of Building",    "Anna Petrova",     "CENTRAL-4", "Seattle"),
    ("3/15/2005", "Hire",                  "Anna Petrova",     "CENTRAL-2", "Seattle"),
]


def main() -> None:
    df = pd.DataFrame(
        [
            {
                "Personnel Number": 12345,
                "Alias": "JDOE",
                "Full Name": "John Doe",
                "Activity": "Miscellaneous" if activity != "Current Row" else "Current",
                "Person Status Group": "Active",
                "Activity Date": date,
                "Person Status": "Active",
                "Activity Detail": activity,
                "Reports To Full Name": manager,
                "Position Number": 90000000,
                "Manager Indicator": "N",
                "Area": "United States",
                "Area Detail": "Washington" if city == "Seattle" else "Oregon",
                "City Summary Group": "Pacific Northwest",
                "City": city.upper(),
                "Building Name": building,
                "Work Country": "USA",
            }
            for date, activity, manager, building, city in ROWS
        ]
    )

    out = Path(__file__).parent / "sample_career.xlsx"
    df.to_excel(out, index=False, sheet_name="Export", engine="openpyxl")
    print(f"Wrote {out} ({len(df)} rows)")


if __name__ == "__main__":
    main()
