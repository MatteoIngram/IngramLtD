"""
Export BDC² index data from the Excel model to data/bdc2.js for the website.

Usage (run after updating the workbook each month and saving it in Excel):
    pip install openpyxl
    python export_data.py BDC2_Index_Fund_Model.xlsx

Reads the cached values Excel saves, so open and save the workbook in Excel
(or LibreOffice) before running, otherwise formula cells come back empty.
"""
import json, sys, datetime as dt
import openpyxl

FUNDS = [  # code, fund, sponsor, row in 'Fund Model' with the annualised yield
    ("ASIF", "Ares Strategic Income Fund", "Ares", 3),
    ("BCRED", "Blackstone Private Credit Fund", "Blackstone", 4),
    ("OCIC", "Blue Owl Credit Income Corp", "Blue Owl", 5),
    ("HLEND", "HPS Corporate Lending Fund", "HPS / BlackRock", 6),
    ("ADS", "Apollo Debt Solutions BDC", "Apollo", 7),
]

def month_end(d: dt.datetime) -> str:
    nxt = (d.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
    return (nxt - dt.timedelta(days=1)).date().isoformat()

def main(path: str) -> None:
    wb = openpyxl.load_workbook(path, data_only=True)
    ix = wb["Index"]
    head = {ix.cell(row=2, column=c).value: c for c in range(1, 10)}
    c_date, c_ret = head["Date"], head["Total return (EW)"]
    c_tr, c_ds = head["TR Index Level"], head["De-smoothed Index Level"]
    series = []
    for r in range(3, ix.max_row + 1):
        d, lvl = ix.cell(row=r, column=c_date).value, ix.cell(row=r, column=c_tr).value
        if not isinstance(d, dt.datetime) or not isinstance(lvl, (int, float)):
            continue
        ret = ix.cell(row=r, column=c_ret).value
        ds = ix.cell(row=r, column=c_ds).value
        series.append({"date": month_end(d), "level": round(lvl, 2),
                       "desmoothed": round(ds, 2) if isinstance(ds, (int, float)) else None,
                       "ret": ret if isinstance(ret, (int, float)) else None})
    if not series:
        sys.exit("No index values found: open and save the workbook in Excel first.")
    stats = {}
    for r in range(1, 20):
        k, v = ix[f"H{r}"].value, ix[f"I{r}"].value
        if k and isinstance(v, (int, float)):
            stats[k] = v
    fm = wb["Fund Model"]
    constituents = [{"code": c, "fund": n, "sponsor": s, "weight": 0.2,
                     "yield": fm[f"B{row}"].value} for c, n, s, row in FUNDS]
    out = {"asOf": series[-1]["date"], "base": {"date": series[0]["date"], "level": 1000},
           "series": series, "stats": stats, "constituents": constituents,
           "generated": dt.datetime.now().isoformat(timespec="minutes")}
    with open("data/bdc2.js", "w", encoding="utf-8") as f:
        f.write("window.BDC2 = " + json.dumps(out, indent=1) + ";\n")
    print(f"Wrote data/bdc2.js: {len(series)} months, as of {out['asOf']}, level {series[-1]['level']}")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "BDC2_Index_Fund_Model.xlsx")
