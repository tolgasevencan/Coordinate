# step1_geocode.py
import argparse, time, pandas as pd
from geopy.geocoders import Nominatim
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--infile", required=True, help="Eingabedatei (Excel)")
ap.add_argument("--outfile", help="Optionaler Ausgabepfad")
args = ap.parse_args()

df = pd.read_excel(args.infile)
geolocator = Nominatim(user_agent="route_optimizer")
df["Breitengrad"] = None
df["Längengrad"] = None

print(f"📍 Starte Geokodierung ({len(df)} Adressen)...")

for i, addr in df["Ort"].fillna("").items():
    if addr:
        try:
            loc = geolocator.geocode(addr)
            if loc:
                df.at[i,"Breitengrad"]=loc.latitude
                df.at[i,"Längengrad"]=loc.longitude
                print(f"  ✅ {addr} → ({loc.latitude:.5f}, {loc.longitude:.5f})")
            else:
                print(f"  ⚠️ Keine Koordinaten gefunden: {addr}")
        except Exception as e:
            print(f"  ❌ Fehler bei {addr}: {e}")
        time.sleep(1)

base = Path(args.infile).name.replace("_export.xlsx","")
outfile = args.outfile or f"{base}_geocoded.xlsx"
df.to_excel(outfile, index=False)
print(f"✅ Geokodierung abgeschlossen: {outfile}")
