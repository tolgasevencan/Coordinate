# step2_distance_matrix.py
import pandas as pd
import requests

INPUT = "outlook_export_geocoded.xlsx"   # Adım 1’in çıktısı
OUT_XLSX = "step2_route_report.xlsx"
OSRM = "https://router.project-osrm.org"  # public OSRM

def osrm_table(coords):
    # coords: [(lat,lon), ...]
    if len(coords) < 2:
        return [[0]], [[0]]
    coord_str = ";".join([f"{lon},{lat}" for lat,lon in coords])
    url = f"{OSRM}/table/v1/driving/{coord_str}?annotations=duration,distance"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    js = r.json()
    dur_min = [[round((d or 0)/60, 1) for d in row] for row in js["durations"]]   # sec -> min
    dist_km = [[round((d or 0)/1000, 2) for d in row] for row in js["distances"]] # m -> km
    return dur_min, dist_km

def nearest_neighbor(dist, start=0):
    n = len(dist)
    un = set(range(n)); un.remove(start)
    route = [start]
    while un:
        last = route[-1]
        nxt = min(un, key=lambda j: dist[last][j])
        route.append(nxt); un.remove(nxt)
    return route

def two_opt(route, dist):
    def length(rt): return sum(dist[rt[i]][rt[i+1]] for i in range(len(rt)-1))
    best = route[:]; best_len = length(best); improved = True
    while improved:
        improved = False
        for i in range(1, len(best)-2):
            for k in range(i+1, len(best)-1):
                new = best[:i] + best[i:k+1][::-1] + best[k+1:]
                nl = length(new)
                if nl < best_len - 1e-6:
                    best, best_len = new, nl; improved = True
    return best

def main():
    df = pd.read_excel(INPUT)
    if not {"Latitude","Longitude"}.issubset(df.columns):
        raise SystemExit("Latitude/Longitude yok. Önce step1_geocode.py çalıştırın.")

    # Sırayı takvimdeki gibi koruyoruz
    visits = df[df["Latitude"].notna() & df["Longitude"].notna()].reset_index(drop=True)
    coords = list(zip(visits["Latitude"], visits["Longitude"]))
    if len(coords) < 2:
        raise SystemExit("En az 2 koordinat gerekli.")

    # OSRM matrisleri
    dur, dist = osrm_table(coords)

    # Planlanan sıranın toplamları
    plan_min = sum(dur[i][i+1] for i in range(len(dur)-1))
    plan_km  = sum(dist[i][i+1] for i in range(len(dist)-1))

    # Optimal rota (NN + 2-opt)
    init = nearest_neighbor(dist, start=0)
    opt = two_opt(init, dist)
    opt_min = sum(dur[opt[i]][opt[i+1]] for i in range(len(opt)-1))
    opt_km  = sum(dist[opt[i]][opt[i+1]] for i in range(len(opt)-1))

    # Rapor tabloları
    visits_out = visits.copy()
    visits_out.insert(0, "planned_order", range(1, len(visits)+1))
    visits_out.insert(1, "optimal_order", [opt.index(i)+1 for i in range(len(visits))])

    kpis = pd.DataFrame([{
        "planned_total_duration_min": round(plan_min,1),
        "planned_total_distance_km": round(plan_km,2),
        "optimal_total_duration_min": round(opt_min,1),
        "optimal_total_distance_km": round(opt_km,2),
        "time_saving_min": round(plan_min - opt_min,1),
        "time_saving_pct": round(100*(plan_min - opt_min)/plan_min,1) if plan_min>0 else 0.0
    }])

    # Excel’e yaz
    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as w:
        pd.DataFrame(dur).to_excel(w, index=False, header=False, sheet_name="durations_min")
        pd.DataFrame(dist).to_excel(w, index=False, header=False, sheet_name="distances_km")
        visits_out.to_excel(w, index=False, sheet_name="visits")
        kpis.to_excel(w, index=False, sheet_name="kpis")

    print(f"✅ Rapor hazır: {OUT_XLSX}")

if __name__ == "__main__":
    main()
