"""Combine Harbour Street Cafe's Jan-Jun 2026 sales exports into one sheet
and precompute the figures the dashboard embeds.

Usage: python3 build_data.py <folder with the six source files>
Source files are only read, never modified.
"""
import csv, json, sys, datetime as dt
from collections import OrderedDict, Counter
from pathlib import Path
import openpyxl

SRC = Path(sys.argv[1])
OUT = Path(__file__).parent
MONTHS = [
    ("January", "2026-01-sales.csv"), ("February", "2026-02-sales.csv"),
    ("March", "2026-03-sales.xlsx"), ("April", "2026-04-sales.csv"),
    ("May", "2026-05-sales.csv"), ("June", "2026-06-sales.csv"),
]
ITEMS = ["Flat White", "Latte", "Long Black", "Iced Kopi",
         "Croissant", "Banana Bread", "Kaya Toast", "Chicken Sandwich"]

def find(name):
    hits = [p for p in SRC.iterdir() if p.name.endswith(name)]
    assert len(hits) == 1, (name, hits)
    return hits[0]

def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    for fmt in ("%Y-%m-%d", "%d %b %Y"):
        try:
            return dt.datetime.strptime(v.strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(v)

def read(path):
    if path.suffix == ".xlsx":
        raw = list(openpyxl.load_workbook(path).active.iter_rows(values_only=True))
    else:
        raw = list(csv.reader(path.open(newline="")))
    header, body = raw[0], raw[1:]
    assert [h for h in header][:4] == ["Date", "Item", "Qty", "Unit Price"], header
    # lines: every line in file order; data lines as tuples, others kept raw.
    rows, lines, extra = [], [], Counter()
    for r in body:
        if not r or not any(c not in (None, "") for c in r):
            extra["blank lines"] += 1
            lines.append(("raw", ["", "", "", "", ""]))
        elif r[0] in (None, "") and str(r[1]).strip().upper() == "TOTAL":
            extra["TOTAL line"] += 1
            extra_total = float(r[4])
            lines.append(("raw", ["", r[1], "", "", r[4]]))
        else:
            t = (parse_date(r[0]), r[1], int(r[2]), float(r[3]), float(r[4]))
            rows.append(t)
            lines.append(("row", t))
    if extra["TOTAL line"]:
        assert abs(extra_total - sum(r[4] for r in rows)) < 0.005, "TOTAL line disagrees with rows"
    return header, rows, lines, dict(extra)

combined, months, notes = [], OrderedDict(), []
for month, name in MONTHS:
    path = find(name)
    header, rows, lines, extra = read(path)
    file_total = round(sum(r[4] for r in rows), 2)
    # Repeated (date, item) rows are kept as the owner asked; only reported.
    seen, dupes = set(), []
    for r in rows:
        key = (r[0], r[1])
        if key in seen:
            dupes.append(r)
        seen.add(key)
    kept = rows
    for r in kept:
        assert r[1] in ITEMS and r[0].strftime("%B") == month, r
        assert abs(r[2] * r[3] - r[4]) < 0.005, r
    days = sorted({r[0] for r in kept})
    by_item = {i: {"revenue": 0.0, "units": 0} for i in ITEMS}
    for r in kept:
        by_item[r[1]]["revenue"] += r[4]
        by_item[r[1]]["units"] += r[2]
    for v in by_item.values():
        v["revenue"] = round(v["revenue"], 2)
    months[month] = {
        "file": path.name.split("-", 1)[1],
        "fileRows": len(rows), "fileTotal": file_total,
        "rows": len(kept), "revenue": round(sum(r[4] for r in kept), 2),
        "units": sum(r[2] for r in kept), "tradingDays": len(days),
        "firstDay": days[0].isoformat(), "lastDay": days[-1].isoformat(),
        "repeatedRows": len(dupes),
        "repeatedRevenue": round(sum(r[4] for r in dupes), 2),
        "repeatedDays": sorted({r[0].isoformat() for r in dupes}),
        "extraLines": extra,
        "revenueCol": header[4],
        "items": by_item,
        "days": {d.isoformat(): round(sum(r[4] for r in kept if r[0] == d), 2) for d in days},
    }
    combined += [(l, month) for l in lines]

# Files are concatenated in month order, each in its own line order, so the
# kept blank and TOTAL lines stay where they were.
with (OUT / "harbour-street-cafe-sales-2026-H1.csv").open("w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["Date", "Item", "Qty", "Unit Price", "Revenue", "Month"])
    for (kind, r), m in combined:
        if kind == "raw":
            w.writerow(r + [m])
        else:
            w.writerow([r[0].isoformat(), r[1], r[2], f"{r[3]:g}", f"{r[4]:.2f}", m])

data = {"generated": dt.date.today().isoformat(), "items": ITEMS, "months": months}
(OUT / "dashboard-data.json").write_text(json.dumps(data, indent=1))

print(f"{'Month':10}{'file rows':>10}{'kept':>6}{'file total':>12}{'dashboard':>12}  notes")
for m, v in months.items():
    note = []
    if v["repeatedRows"]: note.append(f"kept {v['repeatedRows']} repeated rows ({', '.join(v['repeatedDays'])}) = {v['repeatedRevenue']:.2f}")
    if v["extraLines"]: note.append("kept, not counted: " + ", ".join(f"{n} {k}" for k, n in v["extraLines"].items()))
    print(f"{m:10}{v['fileRows']:>10}{v['rows']:>6}{v['fileTotal']:>12,.2f}{v['revenue']:>12,.2f}  {'; '.join(note)}")
print(f"{'TOTAL':10}{sum(v['fileRows'] for v in months.values()):>10}{sum(v['rows'] for v in months.values()):>6}"
      f"{sum(v['fileTotal'] for v in months.values()):>12,.2f}{sum(v['revenue'] for v in months.values()):>12,.2f}")
tot = {i: sum(v["items"][i]["revenue"] for v in months.values()) for i in ITEMS}
print("H1 item ranking:", sorted(((round(t,2), i) for i, t in tot.items()), reverse=True))

# Embed the figures into the dashboard page.
page = (OUT / "template.html").read_text().replace("/*__DATA__*/", json.dumps(data, separators=(",", ":")))
(OUT / "harbour-street-cafe-dashboard.html").write_text(page)
print("Wrote", OUT / "harbour-street-cafe-dashboard.html")
