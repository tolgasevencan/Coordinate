import pandas as pd
from geopy.geocoders import Nominatim
import time

# 1. Excel dosyasını yükle
df = pd.read_excel('outlook_export.xlsx')

# 2. Geocoder başlat
geolocator = Nominatim(user_agent="route_optimizer")

# 3. Koordinatları saklamak için kolonlar
df["Latitude"] = None
df["Longitude"] = None

# 4. Adresleri geocode et
for i, row in df.iterrows():
    address = row["Location"]
    if pd.notna(address):
        try:
            location = geolocator.geocode(address)
            if location:
                df.at[i, "Latitude"] = location.latitude
                df.at[i, "Longitude"] = location.longitude
        except Exception as e:
            print(f"Hata: {address} -> {e}")
        time.sleep(1)  # API kısıtlamalarına takılmamak için

# 5. Yeni dosyaya kaydet
df.to_excel("outlook_export_geocoded.xlsx", index=False)

print("✅ Geocoding işlemi tamamlandı! 'outlook_export_geocoded.xlsx' oluşturuldu.")
