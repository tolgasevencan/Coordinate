# app.py — Aussendienst • ICS ➜ Excel ➜ Geokodierung ➜ Tagesbasierter Routenreport
import io, time, json, requests
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
from ics import Calendar
from geopy.geocoders import Nominatim

# ───────────────────────────────────────────────────────────────────────────────
# BASICS
TZ = "Europe/Zurich"
st.set_page_config(page_title="Aussendienst Analyse", page_icon="🚗")
st.title("🚗 Aussendienst – ICS ➜ Excel ➜ Geokodierung ➜ Routenanalyse")

# Nominatim kullanım şartı gereği: user_agent içinde iletişim bilgini yaz
USER_AGENT = "aussendienst-route-optimizer (contact: you@example.com)"
geocoder = Nominatim(user_agent=USER_AGENT, timeout=15)

# ───────────────────────────────────────────────────────────────────────────────
# HELPERS

def parse_ics_to_df(ics_bytes: bytes) -> pd.DataFrame:
    cal = Calendar(ics_bytes.decode("utf-8", errors="ignore"))
    rows = []
    for e in cal.events:
        s = e.begin
        t_start = s.to("local").datetime if s else None
        eend = e.end
        t_end = eend.to("local").datetime if eend else None
        rows.append({
            "Betreff": e.name or "",
            "Startdatum": t_start.date().isoformat() if t_start else "",
            "Startzeit": t_start.strftime("%H:%M:%S") if t_start else "",
            "Enddatum": t_end.date().isoformat() if t_end else "",
            "Endzeit": t_end.strftime("%H:%M:%S") if t_end else "",
            "Ort": (e.location or "").strip(),
            "Beschreibung": (e.description or "").strip()
        })
    df = pd.DataFrame(rows).sort_values(["Startdatum","Startzeit"]).reset_index(drop=True)
    return df


def geocode_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Breitengrad"] = None
    out["Längengrad"] = None
    for i, row in out.iterrows():
        addr = str(row["Ort"]).strip()
        if not addr:
            continue
        try:
            loc = geocoder.geocode(addr)
            if loc:
                out.at[i, "Breitengrad"] = round(float(loc.latitude), 5)
                out.at[i, "Längengrad"] = round(float(loc.longitude), 5)
        except Exception:
            pass
        time.sleep(1)  # Nominatim nezaketi (rate limit)
    return out


def osrm_table(coords):
    """coords = [(lat, lon), ...]  →  minutes matrix, km matrix"""
    if len(coords) < 2:
        return None, None, "Zu wenige Koordinaten für Matrix."
    coord_str = ";".join([f"{lon},{lat}" for lat, lon in coords])
    url = f"https://router.project-osrm.org/table/v1/driving/{coord_str}?annotations=duration,distance"
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        js = r.json()
        if "durations" not in js or "distances" not in js:
            return None, None, "Unerwartete OSRM-Antwort."
        dur = [[round((d or 0)/60, 1) for d in row] for row in js["durations"]]   # Minuten
        dist = [[round((d or 0)/1000, 2) for d in row] for row in js["distances"]]  # km
        return dur, dist, None
    except Exception as e:
        return None, None, f"OSRM-Anfrage fehlgeschlagen: {e}"


def nn(D, start=0):
    n = len(D); left = set(range(n)); left.remove(start)
    r = [start]
    while left:
        j = min(left, key=lambda x: D[r[-1]][x])
        r.append(j); left.remove(j)
    return r


def two_opt(route, D):
    def L(rt): return sum(D[rt[i]][rt[i+1]] for i in range(len(rt)-1))
    best = route[:]; improved = True
    while improved:
        improved = False
        for i in range(1, len(best)-2):
            for k in range(i+1, len(best)-1):
                new = best[:i] + best[i:k+1][::-1] + best[k+1:]
                if L(new) < L(best):
                    best = new; improved = True
    return best


def label_for_row(row):
    t = str(row["Startzeit"])[:5] if isinstance(row["Startzeit"], str) else ""
    o = str(row["Ort"]).split(",")[0][:40]
    return f"{t} {o}".strip()


def build_daywise_report(df_all: pd.DataFrame, base_name: str) -> bytes:
    """Gün-gün Excel raporu (çoklu sheet) üretir, bytes döner."""
    # tip: tarih alanı
    df_all = df_all.copy()
    df_all["Startdatum"] = pd.to_datetime(df_all["Startdatum"]).dt.date
    days = sorted(df_all["Startdatum"].dropna().unique().tolist())

    overview_rows = []
    hints = []

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        for day in days:
            df = df_all[df_all["Startdatum"] == day].copy()
            # planlanan sıraya göre saat
            df["__t"] = pd.to_datetime(df["Startzeit"], format="%H:%M:%S", errors="coerce")
            df = df.sort_values("__t").drop(columns="__t").reset_index(drop=True)
            df = df.dropna(subset=["Breitengrad","Längengrad"])

            if len(df) < 2:
                overview_rows.append({
                    "Datum": str(day), "Besuche": len(df),
                    "Geplante Dauer (Min)": 0.0, "Geplante Distanz (km)": 0.0,
                    "Optimierte Dauer (Min)": 0.0, "Optimierte Distanz (km)": 0.0,
                    "Ersparnis (Min)": 0.0, "Ersparnis (%)": 0.0, "Hinweis": "Zu wenige Adressen"
                })
                df.to_excel(w, index=False, sheet_name=f"{day} Besuche")
                continue

            labels = [f"{i+1}. {label_for_row(r)}" for i, r in df.iterrows()]
            coords = list(zip(df["Breitengrad"], df["Längengrad"]))

            dur, dist, err = osrm_table(coords)
            if err:
                hints.append(f"{day}: {err}")
                overview_rows.append({
                    "Datum": str(day), "Besuche": len(df),
                    "Geplante Dauer (Min)": 0.0, "Geplante Distanz (km)": 0.0,
                    "Optimierte Dauer (Min)": 0.0, "Optimierte Distanz (km)": 0.0,
                    "Ersparnis (Min)": 0.0, "Ersparnis (%)": 0.0, "Hinweis": err
                })
                df.to_excel(w, index=False, sheet_name=f"{day} Besuche")
                continue

            # TSP kaba+2-opt
            init = nn(dist, start=0)
            opt = two_opt(init, dist)

            plan_min = sum(dur[i][i+1] for i in range(len(dur)-1))
            opt_min  = sum(dur[opt[i]][opt[i+1]] for i in range(len(opt)-1))
            plan_km  = sum(dist[i][i+1] for i in range(len(dist)-1))
            opt_km   = sum(dist[opt[i]][opt[i+1]] for i in range(len(opt)-1))
            save_min = round(plan_min - opt_min, 1)
            save_pct = round(100*save_min/plan_min, 1) if plan_min > 0 else 0.0

            # görsel tablolar
            df_dur  = pd.DataFrame(dur,  index=labels, columns=labels)
            df_dist = pd.DataFrame(dist, index=labels, columns=labels)

            vis = df.copy()
            vis.insert(0, "Geplante Reihenfolge", range(1, len(df)+1))
            vis.insert(1, "Optimale Reihenfolge", [opt.index(i)+1 for i in range(len(df))])

            route_df = pd.DataFrame({
                "Geplante Reihenfolge": range(1, len(labels)+1),
                "Geplanter Stopp":      labels,
                "Optimale Reihenfolge": range(1, len(labels)+1),
                "Optimaler Stopp":      [labels[i] for i in opt]
            })

            kpis = pd.DataFrame([{
                "Datum": str(day),
                "Besuche": len(df),
                "Geplante Dauer (Min)": round(plan_min,1),
                "Geplante Distanz (km)": round(plan_km,2),
                "Optimierte Dauer (Min)": round(opt_min,1),
                "Optimierte Distanz (km)": round(opt_km,2),
                "Ersparnis (Min)": save_min,
                "Ersparnis (%)": save_pct
            }])

            df_dur.to_excel(w,  sheet_name=f"{day} Dauer")
            df_dist.to_excel(w, sheet_name=f"{day} Distanz")
            vis.to_excel(w,     index=False, sheet_name=f"{day} Besuche")
            route_df.to_excel(w,index=False, sheet_name=f"{day} Route")
            kpis.to_excel(w,    index=False, sheet_name=f"{day} KPIs")

            overview_rows.append({
                "Datum": str(day), "Besuche": len(df),
                "Geplante Dauer (Min)": round(plan_min,1),
                "Optimierte Dauer (Min)": round(opt_min,1),
                "Ersparnis (Min)": save_min,
                "Ersparnis (%)": save_pct,
                "Hinweis": ""
            })

        # Gesamt-Übersicht + Hinweise
        ueb = pd.DataFrame(overview_rows)
        if not ueb.empty:
            totals = pd.DataFrame([{
                "Datum": "SUMME",
                "Besuche": ueb["Besuche"].sum(),
                "Geplante Dauer (Min)": round(ueb["Geplante Dauer (Min)"].sum(),1),
                "Optimierte Dauer (Min)": round(ueb["Optimierte Dauer (Min)"].sum(),1),
                "Ersparnis (Min)": round(ueb["Ersparnis (Min)"].sum(),1),
                "Ersparnis (%)": ""
            }])
            ueb = pd.concat([ueb, totals], ignore_index=True)
        ueb.to_excel(w, index=False, sheet_name="Übersicht")

        if hints:
            pd.DataFrame({"Hinweise": hints}).to_excel(w, index=False, sheet_name="Hinweise")

    buf.seek(0)
    filename = f"{base_name}_export_route_report.xlsx"
    st.success("Tagesbasierter Routenreport erstellt.")
    st.download_button("📥 Excel exportieren", buf, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return buf.getvalue()

# ───────────────────────────────────────────────────────────────────────────────
# UI (3 Adım)

tab1, tab2, tab3 = st.tabs(["ICS ➜ Excel", "Geokodierung (Excel)", "Routenanalyse & Report"])

with tab1:
    st.subheader("Schritt 1: ICS-Datei in Excel konvertieren")
    up = st.file_uploader("ICS-Datei hochladen (z.B. 20250901-20250930_Name.ics)", type="ics")
    if up:
        df_ics = parse_ics_to_df(up.getvalue())
        st.info(f"**Gesamt Ereigniszahl:** {len(df_ics)}")
        st.dataframe(df_ics, use_container_width=True)
        base_name = Path(up.name).stem
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
            df_ics.to_excel(w, index=False, sheet_name="Termine")
        bio.seek(0)
        st.download_button("📥 Excel exportieren", bio, file_name=base_name+"_export.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab2:
    st.subheader("Schritt 2: Geokodierung (Excel ➜ Koordinaten)")
    up2 = st.file_uploader("Excel mit Terminen (aus Schritt 1)", type=["xlsx"], key="geocode")
    colA, colB = st.columns(2)
    with colA:
        st.caption("Nominatim-Hinweis: Bitte ein paar Sekunden Geduld (Rate-Limit)")
    if up2:
        df_in = pd.read_excel(up2)
        need = {"Startdatum","Startzeit","Ort"}
        miss = need - set(df_in.columns)
        if miss:
            st.error(f"Fehlende Spalten: {miss}")
        else:
            df_out = geocode_df(df_in)
            st.dataframe(df_out.head(50), use_container_width=True)
            base_name = Path(up2.name).stem.replace("_export", "")
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
                df_out.to_excel(w, index=False, sheet_name="Termine_geocoded")
            bio.seek(0)
            st.download_button("📥 Geokodiertes Excel exportieren", bio,
                               file_name=base_name+"_export_geocoded.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab3:
    st.subheader("Schritt 3: Tagesbasierte Routenanalyse & Report")
    up3 = st.file_uploader("Geokodiertes Excel (aus Schritt 2)", type=["xlsx"], key="route")
    if up3:
        df_geo = pd.read_excel(up3)
        need = {"Startdatum","Startzeit","Ort","Breitengrad","Längengrad"}
        miss = need - set(df_geo.columns)
        if miss:
            st.error(f"Fehlende Spalten im geokodierten Excel: {miss}")
        else:
            st.write("**Hinweis:** Für jeden Tag werden fünf Sheets erzeugt: *Dauer*, *Distanz*, *Besuche*, *Route*, *KPIs*. Zusätzlich *Übersicht* und ggf. *Hinweise*.")
            base_name = Path(up3.name).stem.replace("_export_geocoded", "")
            build_daywise_report(df_geo, base_name)

# Test
