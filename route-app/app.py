# app.py — End-to-End: ICS -> Excel -> Geokodierung -> Routenanalyse (Deutsch)
import io, time, math, pandas as pd, streamlit as st
from ics import Calendar
from geopy.geocoders import Nominatim
from io import BytesIO

st.set_page_config(page_title="Aussendienst Analyse", page_icon="🚚", layout="wide")
geocoder = Nominatim(user_agent="aussendienst-route-optimizer (contact: tolgasevencan@icloud.com)")

# ---------- Helfer ----------
def parse_ics(ics_bytes: bytes) -> pd.DataFrame:
    cal = Calendar(ics_bytes.decode("utf-8", errors="ignore"))
    rows = []
    for e in cal.events:
        s = e.begin.datetime
        eend = e.end.datetime
        rows.append({
            "Betreff": e.name or "",
            "Startdatum": s.date(),
            "Startzeit": s.strftime("%H:%M:%S"),
            "Enddatum": eend.date(),
            "Endzeit": eend.strftime("%H:%M:%S"),
            "Ort": (e.location or "").strip(),
            "Beschreibung": e.description or ""
        })
    df = pd.DataFrame(rows).sort_values(["Startdatum", "Startzeit"]).reset_index(drop=True)
    return df

def geocode_df(df: pd.DataFrame, addr_col="Ort", sleep_sec=1.0) -> pd.DataFrame:
    out = df.copy()
    out["Breitengrad"] = None
    out["Längengrad"] = None
    for i, row in out.iterrows():
        adr = (row.get(addr_col) or "").strip()
        if not adr:
            continue
        try:
            loc = geocoder.geocode(adr)
            if loc:
                out.at[i, "Breitengrad"] = loc.latitude
                out.at[i, "Längengrad"] = loc.longitude
        except Exception:
            pass
        time.sleep(sleep_sec)   # Respektiere Nominatim-Rate-Limits
    return out

def haversine_km(lat1, lon1, lat2, lon2):
    import math
    R = 6371.0
    p1 = math.radians(lat1); p2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dlambda/2)**2
    return 2*R*math.asin(math.sqrt(a))

def distance_duration_mats(coords, kmh=35.0):
    n = len(coords)
    dist = [[0.0]*n for _ in range(n)]
    dur  = [[0.0]*n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            d = haversine_km(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
            dist[i][j] = d
            dur[i][j]  = (d / max(kmh, 1e-6)) * 60.0  # Minuten
    return dur, dist

def nearest_neighbor_order(dur):
    n = len(dur); visited = [False]*n; order = [0]; visited[0] = True
    for _ in range(n-1):
        last = order[-1]
        nxt = min((j for j in range(n) if not visited[j]), key=lambda j: dur[last][j])
        visited[nxt] = True
        order.append(nxt)
    return order

def excel_download(df: pd.DataFrame, filename: str, sheet_name="Export"):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet_name)
    st.download_button(
        "📥 Excel herunterladen",
        data=buf.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- UI ----------
st.title("🚚 Aussendienst – ICS ➜ Excel ➜ Geokodierung ➜ Routenanalyse")
st.caption("Ein einziger Ablauf: ICS hochladen, automatisch geokodieren und Route berechnen.")

uploaded = st.file_uploader("ICS-Datei hochladen (z.B. 20250901-20250930_Name.ics)", type=["ics"])
kmh = st.slider("Fahrgeschwindigkeit (km/h) – grobe Annahme", 20, 80, 35, 1)
start_as_home = st.text_input("Startadresse (optional; leer = erster Termin)", "")

if not uploaded:
    st.info("Bitte eine ICS-Datei hochladen.")
    st.stop()

base_name = uploaded.name.replace(".ics", "")
st.write(f"**Datei:** {uploaded.name}")

# 1) ICS -> Excel
with st.status("ICS wird gelesen …", expanded=False):
    df = parse_ics(uploaded.getvalue())
st.success(f"Termine insgesamt: {len(df)}")

c1, c2 = st.columns(2)
with c1:
    st.subheader("Rohdaten (ICS → Tabelle)")
    st.dataframe(df, use_container_width=True, height=320)
with c2:
    excel_download(df, f"{base_name}_export.xlsx")

# 2) Geokodierung
with st.status("Adressen werden geokodiert (Nominatim) …", expanded=False):
    df_geo = geocode_df(df, "Ort", sleep_sec=1.0)
ok_rows = df_geo.dropna(subset=["Breitengrad", "Längengrad"])
miss = len(df_geo) - len(ok_rows)
st.success(f"Geokodierung abgeschlossen: {len(ok_rows)} Einträge, übersprungen: {miss}")

c3, c4 = st.columns(2)
with c3:
    st.subheader("Geokodierte Tabelle")
    st.dataframe(df_geo, use_container_width=True, height=320)
with c4:
    excel_download(df_geo, f"{base_name}_export_geocoded.xlsx")

# 3) Routenanalyse (ein Tag, einfache Heuristik)
coords = ok_rows[["Breitengrad", "Längengrad"]].to_numpy().tolist()
labels = ok_rows["Betreff"].fillna("").astype(str).tolist()

# Optionaler Startpunkt
if start_as_home.strip():
    try:
        loc = geocoder.geocode(start_as_home.strip())
        if loc:
            coords = [[loc.latitude, loc.longitude]] + coords
            labels = [f"Start/Zuhause: {start_as_home.strip()}"] + labels
            st.info("Startpunkt hinzugefügt.")
            time.sleep(1.0)
    except Exception:
        pass

if len(coords) < 2:
    st.warning("Für eine Route werden mindestens 2 gültige Koordinaten benötigt.")
    st.stop()

dur, dist = distance_duration_mats(coords, kmh=kmh)
opt = nearest_neighbor_order(dur)

# KPIs
plan_min = sum(dur[i][i+1] for i in range(len(opt)-1))
plan_km  = sum(dist[i][i+1] for i in range(len(opt)-1))
kpi_df = pd.DataFrame([{
    "Gesamtdauer (Minuten)": round(plan_min, 1),
    "Gesamtdistanz (km)": round(plan_km, 1),
    "Anzahl Stopps": len(coords)
}])

# Report-Tabellen
idx_labels = [f"{i}" for i in range(len(labels))]
df_dur  = pd.DataFrame(dur, index=idx_labels, columns=idx_labels)
df_dist = pd.DataFrame(dist, index=idx_labels, columns=idx_labels)

route_labels = [labels[i] for i in opt]
route_df = pd.DataFrame({
    "Reihenfolge": list(range(len(route_labels))),
    "Bezeichnung": route_labels
})

st.subheader("🧭 Routen-Übersicht")
c5, c6 = st.columns([1,1])
with c5:
    st.metric("Gesamtdauer (Min)", f"{round(plan_min,1)}")
    st.metric("Gesamtdistanz (km)", f"{round(plan_km,1)}")
with c6:
    st.dataframe(kpi_df, use_container_width=True, height=120)

st.subheader("Besuchsreihenfolge")
st.dataframe(route_df, use_container_width=True, height=260)

# Sammel-Excel (Report)
out_buf = BytesIO()
with pd.ExcelWriter(out_buf, engine="xlsxwriter") as xw:
    df_dur.to_excel(xw, sheet_name="Dauer (Minuten)")
    df_dist.to_excel(xw, sheet_name="Distanz (km)")
    ok_rows.to_excel(xw, index=False, sheet_name="Besuche")
    route_df.to_excel(xw, index=False, sheet_name="Route")
    kpi_df.to_excel(xw, index=False, sheet_name="KPIs")

st.download_button(
    "📥 Routen-Report herunterladen",
    data=out_buf.getvalue(),
    file_name=f"{base_name}_export_route_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Hinweis: Geokodierung via Nominatim (OpenStreetMap). Zwischen Anfragen wird ~1 s pausiert. "
           "Reisezeiten basieren auf Haversine-Distanz + geschätzter Geschwindigkeit – nur Näherung.")
