# app.py — Aussendienst • ICS ➜ Geokodierung ➜ Tagesbasierter Routenreport

import io
import time
import json
import requests
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from ics import Calendar
from geopy.geocoders import Nominatim

# ───────────────────────────────────────────────────────────────────────────────
# GRUNDEINSTELLUNGEN

TZ = "Europe/Zurich"
st.set_page_config(page_title="Aussendienst Analyse", page_icon="🚗")
st.title("🚗 Aussendienst – ICS ➜ Excel ➜ Geokodierung ➜ Routenanalyse")

# Hinweis laut Nominatim-Nutzungsbedingungen: user_agent muss Kontaktinformation enthalten
USER_AGENT = "aussendienst-route-optimizer (contact: you@example.com)"
geocoder = Nominatim(user_agent=USER_AGENT, timeout=15)

# ───────────────────────────────────────────────────────────────────────────────
# HILFSFUNKTIONEN


def parse_ics_to_df(ics_bytes: bytes) -> pd.DataFrame:
    """ICS-Datei einlesen und als DataFrame mit Terminen zurückgeben."""
    cal = Calendar(ics_bytes.decode("utf-8", errors="ignore"))
    rows = []
    for e in cal.events:
        s = e.begin
        t_start = s.to("local").datetime if s else None
        eend = e.end
        t_end = eend.to("local").datetime if eend else None
        rows.append(
            {
                "Betreff": e.name or "",
                "Startdatum": t_start.date().isoformat() if t_start else "",
                "Startzeit": t_start.strftime("%H:%M:%S") if t_start else "",
                "Enddatum": t_end.date().isoformat() if t_end else "",
                "Endzeit": t_end.strftime("%H:%M:%S") if t_end else "",
                "Ort": (e.location or "").strip(),
                "Beschreibung": (e.description or "").strip(),
            }
        )

    df = (
        pd.DataFrame(rows)
        .sort_values(["Startdatum", "Startzeit"])
        .reset_index(drop=True)
    )
    return df


def geocode_df(df: pd.DataFrame) -> pd.DataFrame:
    """Ort-Spalte geokodieren und Breiten-/Längengrad ergänzen."""
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
            # Fehler bei einem einzelnen Geocode-Request ignorieren
            pass

        # Höflichkeit gegenüber Nominatim (Rate Limit)
        time.sleep(1)

    return out


def osrm_table(coords):
    """coords = [(lat, lon), ...]  →  Minutenmatrix, km-Matrix"""
    if len(coords) < 2:
        return None, None, "Zu wenige Koordinaten für die Matrix."

    coord_str = ";".join([f"{lon},{lat}" for lat, lon in coords])
    url = (
        f"https://router.project-osrm.org/table/v1/driving/"
        f"{coord_str}?annotations=duration,distance"
    )

    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        js = r.json()
        if "durations" not in js or "distances" not in js:
            return None, None, "Unerwartete OSRM-Antwort."

        dur = [
            [round((d or 0) / 60, 1) for d in row] for row in js["durations"]
        ]  # Minuten
        dist = [
            [round((d or 0) / 1000, 2) for d in row] for row in js["distances"]
        ]  # km
        return dur, dist, None
    except Exception as e:
        return None, None, f"OSRM-Anfrage fehlgeschlagen: {e}"


def nn(D, start=0):
    """Nearest-Neighbor-Heuristik."""
    n = len(D)
    left = set(range(n))
    left.remove(start)
    r = [start]
    while left:
        j = min(left, key=lambda x: D[r[-1]][x])
        r.append(j)
        left.remove(j)
    return r


def two_opt(route, D):
    """2-Opt-Verbesserung der Route."""

    def L(rt):
        return sum(D[rt[i]][rt[i + 1]] for i in range(len(rt) - 1))

    best = route[:]
    improved = True
    while improved:
        improved = False
        for i in range(1, len(best) - 2):
            for k in range(i + 1, len(best) - 1):
                new = best[:i] + best[i : k + 1][::-1] + best[k + 1 :]
                if L(new) < L(best):
                    best = new
                    improved = True
    return best


def label_for_row(row):
    t = str(row["Startzeit"])[:5] if isinstance(row["Startzeit"], str) else ""
    o = str(row["Ort"]).split(",")[0][:40]
    return f"{t} {o}".strip()


def build_daywise_report(df_all: pd.DataFrame, base_name: str) -> bytes:
    """
    Erstellt einen tagesbasierten Excel-Report mit mehreren Sheets
    und gibt die Bytes des erzeugten Files zurück.
    """

    df_all = df_all.copy()
    df_all["Startdatum"] = pd.to_datetime(df_all["Startdatum"]).dt.date
    days = sorted(df_all["Startdatum"].dropna().unique().tolist())

    overview_rows = []
    hints = []

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        for day in days:
            df = df_all[df_all["Startdatum"] == day].copy()

            # Nach geplanter Startzeit sortieren
            df["__t"] = pd.to_datetime(
                df["Startzeit"], format="%H:%M:%S", errors="coerce"
            )
            df = df.sort_values("__t").drop(columns="__t").reset_index(drop=True)

            # Nur Zeilen mit Geokoordinaten verwenden
            df = df.dropna(subset=["Breitengrad", "Längengrad"])

            if len(df) < 2:
                overview_rows.append(
                    {
                        "Datum": str(day),
                        "Besuche": len(df),
                        "Geplante Dauer (Min)": 0.0,
                        "Geplante Distanz (km)": 0.0,
                        "Optimierte Dauer (Min)": 0.0,
                        "Optimierte Distanz (km)": 0.0,
                        "Ersparnis (Min)": 0.0,
                        "Ersparnis (%)": 0.0,
                        "Hinweis": "Zu wenige Adressen",
                    }
                )
                df.to_excel(w, index=False, sheet_name=f"{day} Besuche")
                continue

            labels = [f"{i+1}. {label_for_row(r)}" for i, r in df.iterrows()]
            coords = list(zip(df["Breitengrad"], df["Längengrad"]))

            dur, dist, err = osrm_table(coords)
            if err:
                hints.append(f"{day}: {err}")
                overview_rows.append(
                    {
                        "Datum": str(day),
                        "Besuche": len(df),
                        "Geplante Dauer (Min)": 0.0,
                        "Geplante Distanz (km)": 0.0,
                        "Optimierte Dauer (Min)": 0.0,
                        "Optimierte Distanz (km)": 0.0,
                        "Ersparnis (Min)": 0.0,
                        "Ersparnis (%)": 0.0,
                        "Hinweis": err,
                    }
                )
                df.to_excel(w, index=False, sheet_name=f"{day} Besuche")
                continue

            # TSP-Heuristik (Nearest Neighbor + 2-Opt)
            init = nn(dist, start=0)
            opt = two_opt(init, dist)

            plan_min = sum(dur[i][i + 1] for i in range(len(dur) - 1))
            opt_min = sum(dur[opt[i]][opt[i + 1]] for i in range(len(opt) - 1))
            plan_km = sum(dist[i][i + 1] for i in range(len(dist) - 1))
            opt_km = sum(dist[opt[i]][opt[i + 1]] for i in range(len(opt) - 1))
            save_min = round(plan_min - opt_min, 1)
            save_pct = round(100 * save_min / plan_min, 1) if plan_min > 0 else 0.0

            # Visuelle Tabellen
            df_dur = pd.DataFrame(dur, index=labels, columns=labels)
            df_dist = pd.DataFrame(dist, index=labels, columns=labels)

            vis = df.copy()
            vis.insert(0, "Geplante Reihenfolge", range(1, len(df) + 1))
            vis.insert(
                1, "Optimale Reihenfolge", [opt.index(i) + 1 for i in range(len(df))]
            )

            route_df = pd.DataFrame(
                {
                    "Geplante Reihenfolge": range(1, len(labels) + 1),
                    "Geplanter Stopp": labels,
                    "Optimale Reihenfolge": range(1, len(labels) + 1),
                    "Optimaler Stopp": [labels[i] for i in opt],
                }
            )

            kpis = pd.DataFrame(
                [
                    {
                        "Datum": str(day),
                        "Besuche": len(df),
                        "Geplante Dauer (Min)": round(plan_min, 1),
                        "Geplante Distanz (km)": round(plan_km, 2),
                        "Optimierte Dauer (Min)": round(opt_min, 1),
                        "Optimierte Distanz (km)": round(opt_km, 2),
                        "Ersparnis (Min)": save_min,
                        "Ersparnis (%)": save_pct,
                    }
                ]
            )

            df_dur.to_excel(w, sheet_name=f"{day} Dauer")
            df_dist.to_excel(w, sheet_name=f"{day} Distanz")
            vis.to_excel(w, index=False, sheet_name=f"{day} Besuche")
            route_df.to_excel(w, index=False, sheet_name=f"{day} Route")
            kpis.to_excel(w, index=False, sheet_name=f"{day} KPIs")

            overview_rows.append(
                {
                    "Datum": str(day),
                    "Besuche": len(df),
                    "Geplante Dauer (Min)": round(plan_min, 1),
                    "Geplante Distanz (km)": round(plan_km, 2),
                    "Optimierte Dauer (Min)": round(opt_min, 1),
                    "Optimierte Distanz (km)": round(opt_km, 2),
                    "Ersparnis (Min)": save_min,
                    "Ersparnis (%)": save_pct,
                    "Hinweis": "",
                }
            )

        # Gesamtübersicht + Hinweise
        ueb = pd.DataFrame(overview_rows)
        if not ueb.empty:
            totals = pd.DataFrame(
                [
                    {
                        "Datum": "SUMME",
                        "Besuche": ueb["Besuche"].sum(),
                        "Geplante Dauer (Min)": round(
                            ueb["Geplante Dauer (Min)"].sum(), 1
                        ),
                        "Geplante Distanz (km)": round(
                            ueb["Geplante Distanz (km)"].sum(), 2
                        ),
                        "Optimierte Dauer (Min)": round(
                            ueb["Optimierte Dauer (Min)"].sum(), 1
                        ),
                        "Optimierte Distanz (km)": round(
                            ueb["Optimierte Distanz (km)"].sum(), 2
                        ),
                        "Ersparnis (Min)": round(ueb["Ersparnis (Min)"].sum(), 1),
                        "Ersparnis (%)": "",
                        "Hinweis": "",
                    }
                ]
            )
            ueb = pd.concat([ueb, totals], ignore_index=True)

        ueb.to_excel(w, index=False, sheet_name="Übersicht")

        if hints:
            pd.DataFrame({"Hinweise": hints}).to_excel(
                w, index=False, sheet_name="Hinweise"
            )

    buf.seek(0)
    filename = f"{base_name}_export_route_report.xlsx"
    st.success("Tagesbasierter Routenreport erstellt.")
    st.download_button(
        "📥 Excel exportieren",
        buf,
        file_name=filename,
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
    return buf.getvalue()


# ───────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI – ALLES IN EINEM SCHRITT

st.header("1️⃣ ICS hochladen und vollständigen Routenreport erstellen")

uploaded_ics = st.file_uploader("Outlook-ICS-Datei hochladen", type=["ics"])

if uploaded_ics is not None:
    base_name = Path(uploaded_ics.name).stem or "kalender"

    try:
        ics_bytes = uploaded_ics.read()

        with st.spinner("📆 ICS wird analysiert..."):
            df_raw = parse_ics_to_df(ics_bytes)
            if df_raw.empty:
                st.warning("In der ICS-Datei wurden keine Termine gefunden.")
            else:
                st.subheader("Aus dem Kalender gelesene Termine")
                st.dataframe(df_raw)

        if not df_raw.empty:
            with st.spinner("📍 Adressen werden geokodiert..."):
                df_geo = geocode_df(df_raw)
                st.subheader("Termine mit Geokoordinaten")
                st.dataframe(df_geo)

            with st.spinner("🧮 Routenoptimierung & Excel-Report wird erstellt..."):
                _ = build_daywise_report(df_geo, base_name)

    except Exception as e:
        st.error(f"Fehler bei der Verarbeitung der ICS-Datei: {e}")
else:
    st.info("Bitte oben eine ICS-Datei hochladen.")
