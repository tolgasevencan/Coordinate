# parse_calendar.py
import sys
from ics import Calendar
import pandas as pd
from datetime import timezone
from zoneinfo import ZoneInfo  # Python 3.9+

tz = ZoneInfo("Europe/Zurich")

infile = sys.argv[1] if len(sys.argv) > 1 else "Beat Majoleth Calendar.ics"
outfile = "outlook_export.xlsx"

with open(infile, "r", encoding="utf-8") as f:
    cal = Calendar(f.read())

rows = []
for e in cal.events:
    start = e.begin.datetime
    end = e.end.datetime
    # TZ'yi İsviçre’ye çevir (tz-aware ise)
    if start.tzinfo:
        start = start.astimezone(tz)
    if end.tzinfo:
        end = end.astimezone(tz)
    rows.append({
        "Subject": e.name or "",
        "Start Date": start.date(),
        "Start Time": start.strftime("%H:%M:%S"),
        "End Date": end.date(),
        "End Time": end.strftime("%H:%M:%S"),
        "Location": e.location or "",
        "Description": e.description or ""
    })

df = pd.DataFrame(rows).sort_values(["Start Date","Start Time"])
df.to_excel(outfile, index=False)
print(f"✅ Excel hazır: {outfile}")
