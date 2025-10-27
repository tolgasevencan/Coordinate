# parse_calendar.py
import re, argparse
from ics import Calendar
import pandas as pd
from zoneinfo import ZoneInfo
from pathlib import Path

def slug(s): return re.sub(r"[^A-Za-z0-9._-]+", "_", s).strip("_")

ap = argparse.ArgumentParser()
ap.add_argument("--ics", required=True, help="ICS-Datei")
ap.add_argument("--name", help="Name der Person (optional; wird aus dem Dateinamen gelesen)")
ap.add_argument("--tz", default="Europe/Zurich")
args = ap.parse_args()

ics_path = Path(args.ics)
m = re.match(r"(?i)(\d{8})-(\d{8})_(.+)\.ics$", ics_path.name)
name_from_file = slug(m.group(3)) if m else None
name = slug(args.name) if args.name else (name_from_file or "Kalender")

tz = ZoneInfo(args.tz)
with open(ics_path, "r", encoding="utf-8") as f:
    cal = Calendar(f.read())

rows=[]
for e in cal.events:
    s = e.begin.datetime
    e_ = e.end.datetime
    if s.tzinfo: s = s.astimezone(tz)
    if e_.tzinfo: e_ = e_.astimezone(tz)
    rows.append({
        "Betreff": e.name or "",
        "Startdatum": s.date(), "Startzeit": s.strftime("%H:%M:%S"),
        "Enddatum": e_.date(), "Endzeit": e_.strftime("%H:%M:%S"),
        "Ort": e.location or "", "Beschreibung": e.description or ""
    })

df = pd.DataFrame(rows).sort_values(["Startdatum","Startzeit"])
if df.empty: raise SystemExit("❌ Keine Termine in der ICS-Datei gefunden.")

dmin = df["Startdatum"].min().strftime("%Y%m%d")
dmax = df["Enddatum"].max().strftime("%Y%m%d")
base = f"{dmin}-{dmax}_{name}"
outfile = f"{base}_export.xlsx"

df.to_excel(outfile, index=False)
print(f"✅ Excel-Datei erstellt: {outfile}")
print(f"📅 Zeitraum: {dmin} - {dmax}")
print(f"👤 Kalender: {name}")
