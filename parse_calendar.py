# parse_calendar.py
import re, argparse
from ics import Calendar
import pandas as pd
from zoneinfo import ZoneInfo
from pathlib import Path

def slug(s): return re.sub(r"[^A-Za-z0-9._-]+", "_", s).strip("_")

ap = argparse.ArgumentParser()
ap.add_argument("--ics", required=True, help="ICS dosyası")
ap.add_argument("--name", help="Kişi adı (opsiyonel; dosya adından alınır)")
ap.add_argument("--tz", default="Europe/Zurich")
args = ap.parse_args()

ics_path = Path(args.ics)
m = re.match(r"(?i)(\d{8})-(\d{8})_(.+)\.ics$", ics_path.name)
name_from_file = slug(m.group(3)) if m else None
name = slug(args.name) if args.name else (name_from_file or "Calendar")

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
        "Subject": e.name or "",
        "Start Date": s.date(), "Start Time": s.strftime("%H:%M:%S"),
        "End Date": e_.date(),   "End Time": e_.strftime("%H:%M:%S"),
        "Location": e.location or "", "Description": e.description or ""
    })

df = pd.DataFrame(rows).sort_values(["Start Date","Start Time"])
if df.empty: raise SystemExit("ICS içinde etkinlik yok.")
dmin = df["Start Date"].min().strftime("%Y%m%d")
dmax = df["End Date"].max().strftime("%Y%m%d")
base = f"{dmin}-{dmax}_{name}"
outfile = f"{base}_export.xlsx"

df.to_excel(outfile, index=False)
print(f"✅ Excel hazır: {outfile}")
