# step1_geocode.py
import argparse, time, pandas as pd
from geopy.geocoders import Nominatim
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--infile", required=True)
ap.add_argument("--outfile")
args = ap.parse_args()

df = pd.read_excel(args.infile)
geolocator = Nominatim(user_agent="route_optimizer")
df["Latitude"] = None; df["Longitude"] = None

for i, addr in df["Location"].fillna("").items():
    if addr:
        try:
            loc = geolocator.geocode(addr)
            if loc:
                df.at[i,"Latitude"]=loc.latitude
                df.at[i,"Longitude"]=loc.longitude
        except Exception as e:
            print("Geocode hata:", addr, e)
        time.sleep(1)

base = Path(args.infile).name.replace("_export.xlsx","")
outfile = args.outfile or f"{base}_geocoded.xlsx"
df.to_excel(outfile, index=False)
print(f"✅ Geocoded: {outfile}")
