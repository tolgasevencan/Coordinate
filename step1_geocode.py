# step1_geocode.py  (DE)
import argparse, time, pandas as pd
from geopy.geocoders import Nominatim
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--infile", required=True, help="Eingabedatei: <BASE>_export.xlsx")
ap.add_argument("--outfile", help="Optionaler Ausgabepfad")
args = ap.parse_args()

df = pd.read_excel(args.infile)

geolocator = Nominatim(user_agent="route_optimizer")
if "Breitengrad" not in df.columns: df["Breitengrad"] = None
if "Längengrad" not in df.columns: df["Längengrad"] = None

print(f"📍 Starte Geokodierung ({len(df)} Zeilen)...")
for i, addr in df["Ort"].fillna("").items():
    if not addr: continue
    try:
        loc = geolocator.geocode(addr)
        if loc:
            df.at[i, "Breitengrad"] = loc.latitude
            df.at[i, "Längengrad"]  = loc.longitude
            print(f"  ✅ {addr} → ({loc.latitude:.5f}, {loc.longitude:.5f})")
        else:
            print(f"  ⚠️ Keine Koordinaten: {addr}")
    except Exception as e:
        print(f"  ❌ Fehler {addr}: {e}")
    time.sleep(1)

base_stem = Path(args.infile).stem              # z.B. <BASE>_export
# deterministik: her zaman _export_geocoded
outfile = args.outfile or f"{base_stem}_geocoded.xlsx"

df.to_excel(outfile, index=False)
print(f"✅ Geokodierung abgeschlossen: {outfile}")
