# parse_calendar.py
import re, argparse
from ics import Calendar
import pandas as pd
from zoneinfo import ZoneInfo
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--ics", required=True, help="YYYYMMDD-YYYYMMDD_Name.ics")
ap.add_argument("--tz", default="Europe/Zurich")
args = ap.parse_args()

ics_path = Path(args.ics)
m = re.match(r"(?i)(\d{8}-\d{8}_.+)\.ics$", ics_path.name)
if not m:
    raise SystemExit("❌ Dateiname muss dem Muster folgen: YYYYMMDD-YYYYMMDD_Name.ics")
base = m.group(1)                                    # <-- sadece dosya adından
outfile = f"{base}_export.xlsx"

tz = ZoneInfo(args.tz)
with open(ics_path, "r", encoding="utf-8") as f:
    cal = Calendar(f.read())

rows = []
for e in cal.events:
    s = e.begin.datetime
    e_ = e.end.datetime
    if s.tzinfo:  s  = s.astimezone(tz)
    if e_.tzinfo: e_ = e_.astimezone(tz)
    rows.append({
        "Betreff": e.name or "",
        "Startdatum": s.date(), "Startzeit": s.strftime("%H:%M:%S"),
        "Enddatum": e_.date(),  "Endzeit":  e_.strftime("%H:%M:%S"),
        "Ort": e.location or "", "Beschreibung": e.description or ""
    })

df = pd.DataFrame(rows).sort_values(["Startdatum","Startzeit"])
df.to_excel(outfile, index=False)
print(f"✅ Excel-Datei erstellt: {outfile} (Basis: {base})")
