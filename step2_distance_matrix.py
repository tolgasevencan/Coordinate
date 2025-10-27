# step2_distance_matrix.py  (DE) — Tagesbasierte Analyse + verständliche Auswertung
import argparse, pandas as pd, requests
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--infile", required=True, help="Eingabe: <BASE>_export_geocoded.xlsx")
ap.add_argument("--outfile", help="Optionaler Ausgabepfad")
args = ap.parse_args()

df_all = pd.read_excel(args.infile)

# Pflichtspalten prüfen
need = {"Startdatum","Startzeit","Ort","Breitengrad","Längengrad"}
missing = need - set(df_all.columns)
if missing:
    raise SystemExit(f"❌ Fehlende Spalten in {args.infile}: {missing}")

# Datum ins richtige Format
df_all["Startdatum"] = pd.to_datetime(df_all["Startdatum"]).dt.date
# Tagesliste
tage = sorted(df_all["Startdatum"].unique())

def osrm_table(coords):
    coord_str = ";".join([f"{lon},{lat}" for lat,lon in coords])
    try:
        r = requests.get(
            f"https://router.project-osrm.org/table/v1/driving/{coord_str}?annotations=duration,distance",
            timeout=30
        )
        r.raise_for_status()
        js = r.json()
    except Exception as e:
        return None, None, f"OSRM-Anfrage fehlgeschlagen: {e}"
    if "durations" not in js or "distances" not in js:
        return None, None, f"Unerwartete OSRM-Antwort: {js}"
    dur  = [[round((d or 0)/60, 1) for d in row] for row in js["durations"]]   # Min
    dist = [[round((d or 0)/1000, 2) for d in row] for row in js["distances"]] # km
    return dur, dist, None

def nn(D, start=0):
    n=len(D); left=set(range(n)); left.remove(start); r=[start]
    while left:
        j=min(left, key=lambda x:D[r[-1]][x]); r.append(j); left.remove(j)
    return r

def two_opt(route, D):
    def L(rt): return sum(D[rt[i]][rt[i+1]] for i in range(len(rt)-1))
    best=route[:]; improved=True
    while improved:
        improved=False
        for i in range(1, len(best)-2):
            for k in range(i+1, len(best)-1):
                new = best[:i] + best[i:k+1][::-1] + best[k+1:]
                if L(new) < L(best): best=new; improved=True
    return best

# Ausgabename deterministisch
in_stem = Path(args.infile).stem                     # <BASE>_export_geocoded
out_stem = in_stem.replace("_geocoded", "_route_report")  # Tagesbasiert, aber gleicher Name
outfile = args.outfile or f"{out_stem}.xlsx"

uebersicht_rows = []   # pro Tag KPI-Zeile
hinweise = []          # Fehlermeldungen/Notizen pro Tag

with pd.ExcelWriter(outfile, engine="xlsxwriter") as w:
    # jeden Tag separat berechnen
    for tag in tage:
        df = df_all[df_all["Startdatum"] == tag].copy()
        # sortiert nach Startzeit (geplante Reihenfolge)
        df["Startzeit_ts"] = pd.to_datetime(df["Startzeit"], format="%H:%M:%S", errors="coerce")
        df = df.sort_values("Startzeit_ts", na_position="last").drop(columns=["Startzeit_ts"])
        # nur gültige Koordinaten
        df = df.dropna(subset=["Breitengrad","Längengrad"]).reset_index(drop=True)

        if len(df) < 2:
            # Mindestens 2 Adressen nötig
            uebersicht_rows.append({
                "Datum": str(tag),
                "Besuche": len(df),
                "Geplante Dauer (Min)": 0.0,
                "Optimierte Dauer (Min)": 0.0,
                "Ersparnis (Min)": 0.0,
                "Ersparnis (%)": 0.0,
                "Hinweis": "Zu wenige Adressen (mindestens 2 erforderlich)"
            })
            # Trotzdem eine kurze Tages-Seite mit Hinweisen
            pd.DataFrame(df).to_excel(w, index=False, sheet_name=f"{tag} Besuche")
            continue

        # Labels lesbarer machen: "HH:MM Ort"
        def short_label(row):
            t = str(row["Startzeit"])[:5] if isinstance(row["Startzeit"], str) else ""
            o = str(row["Ort"]).split(",")[0][:40]
            return f"{t} {o}".strip()
        labels = [f"{i+1}. {short_label(r)}" for i, r in df.iterrows()]
        coords = list(zip(df["Breitengrad"], df["Längengrad"]))

        # OSRM Matrix
        dur, dist, err = osrm_table(coords)
        if err:
            uebersicht_rows.append({
                "Datum": str(tag),
                "Besuche": len(df),
                "Geplante Dauer (Min)": 0.0,
                "Optimierte Dauer (Min)": 0.0,
                "Ersparnis (Min)": 0.0,
                "Ersparnis (%)": 0.0,
                "Hinweis": f"OSRM-Fehler: {err}"
            })
            hinweise.append(f"{tag}: {err}")
            # Tages-Besuche trotzdem speichern
            df.to_excel(w, index=False, sheet_name=f"{tag} Besuche")
            continue

        # Geplante & Optimale Reihenfolge
        init = nn(dist, start=0)
        opt  = two_opt(init, dist)

        plan_min = sum(dur[i][i+1] for i in range(len(dur)-1))
        opt_min  = sum(dur[opt[i]][opt[i+1]] for i in range(len(opt)-1))
        plan_km  = sum(dist[i][i+1] for i in range(len(dist)-1))
        opt_km   = sum(dist[opt[i]][opt[i+1]] for i in range(len(opt)-1))
        save_min = round(plan_min - opt_min, 1)
        save_pct = round(100*save_min/plan_min, 1) if plan_min>0 else 0.0

        # Tabellen pro Tag
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
            "Datum": str(tag),
            "Besuche": len(df),
            "Geplante Dauer (Min)": round(plan_min,1),
            "Geplante Distanz (km)": round(plan_km,2),
            "Optimierte Dauer (Min)": round(opt_min,1),
            "Optimierte Distanz (km)": round(opt_km,2),
            "Ersparnis (Min)": save_min,
            "Ersparnis (%)": save_pct
        }])

        # Sheets schreiben (pro Tag)
        df_dur.to_excel(w,  sheet_name=f"{tag} Dauer")
        df_dist.to_excel(w, sheet_name=f"{tag} Distanz")
        vis.to_excel(w,     index=False, sheet_name=f"{tag} Besuche")
        route_df.to_excel(w, index=False, sheet_name=f"{tag} Route")
        kpis.to_excel(w,    index=False, sheet_name=f"{tag} KPIs")

        # Übersicht-Zeile sammeln
        uebersicht_rows.append({
            "Datum": str(tag),
            "Besuche": len(df),
            "Geplante Dauer (Min)": round(plan_min,1),
            "Optimierte Dauer (Min)": round(opt_min,1),
            "Ersparnis (Min)": save_min,
            "Ersparnis (%)": save_pct,
            "Hinweis": ""
        })

    # Gesamt-Übersicht
    ueb = pd.DataFrame(uebersicht_rows)
    if not ueb.empty:
        totals = pd.DataFrame([{
            "Datum": "SUMME",
            "Besuche": ueb["Besuche"].sum(),
            "Geplante Dauer (Min)": round(ueb["Geplante Dauer (Min)"].sum(),1),
            "Optimierte Dauer (Min)": round(ueb["Optimierte Dauer (Min)"].sum(),1),
            "Ersparnis (Min)": round(ueb["Ersparnis (Min)"].sum(),1),
            "Ersparnis (%)": ""  # Prozent-Summe anlamsız, boş bırakıyoruz
        }])
        ueb = pd.concat([ueb, totals], ignore_index=True)
    ueb.to_excel(w, index=False, sheet_name="Übersicht")

    # Kurze Legende/Hinweise
    if hinweise:
        pd.DataFrame({"Hinweise": hinweise}).to_excel(w, index=False, sheet_name="Hinweise")

print(f"✅ Tagesbasierter Routenreport erstellt: {outfile}")