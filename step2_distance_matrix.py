# step2_distance_matrix.py (labeled)
import pandas as pd
import requests

INPUT = "outlook_export_geocoded.xlsx"
OUT_XLSX = "step2_route_report.xlsx"
OSRM = "https://router.project-osrm.org"

def osrm_table(coords):
    coord_str = ";".join([f"{lon},{lat}" for lat,lon in coords])
    url = f"{OSRM}/table/v1/driving/{coord_str}?annotations=duration,distance"
    r = requests.get(url, timeout=30); r.raise_for_status()
    js = r.json()
    dur_min = [[round((d or 0)/60, 1) for d in row] for row in js["durations"]]
    dist_km = [[round((d or 0)/1000, 2) for d in row] for row in js["distances"]]
    return dur_min, dist_km

def nearest_neighbor(dist, start=0):
    n=len(dist); un=set(range(n)); un.remove(start); route=[start]
    while un:
        last=route[-1]
        nxt=min(un, key=lambda j: dist[last][j])
        route.append(nxt); un.remove(nxt)
    return route

def two_opt(route, dist):
    def length(rt): return sum(dist[rt[i]][rt[i+1]] for i in range(len(rt)-1))
    best=route[:]; best_len=length(best); improved=True
    while improved:
        improved=False
        for i in range(1, len(best)-2):
            for k in range(i+1, len(best)-1):
                new=best[:i]+best[i:k+1][::-1]+best[k+1:]
                nl=length(new)
                if nl < best_len - 1e-6:
                    best, best_len=new, nl; improved=True
    return best

def main():
    df = pd.read_excel(INPUT)
    need = {"Location","Latitude","Longitude"}
    if not need.issubset(df.columns):
        raise SystemExit(f"Eksik kolonlar: {need - set(df.columns)}")

    # Takvim sırası korunur
    visits = df[df["Latitude"].notna() & df["Longitude"].notna()].reset_index(drop=True)

    # Etiketler (kısa ve anlaşılır)
    def short_label(loc):
        s = str(loc).split(",")[0].strip()
        return s[:35]  # çok uzun olmasın
    labels = [f"{i+1}. {short_label(v)}" for i, v in enumerate(visits["Location"])]

    coords = list(zip(visits["Latitude"], visits["Longitude"]))
    dur, dist = osrm_table(coords)

    # Planlanan ve optimal
    plan_min = sum(dur[i][i+1] for i in range(len(dur)-1))
    plan_km  = sum(dist[i][i+1] for i in range(len(dist)-1))

    init = nearest_neighbor(dist, start=0)
    opt  = two_opt(init, dist)
    opt_min = sum(dur[opt[i]][opt[i+1]] for i in range(len(opt)-1))
    opt_km  = sum(dist[opt[i]][opt[i+1]] for i in range(len(opt)-1))

    # Etiketli DataFrame'ler
    df_dur  = pd.DataFrame(dur,  index=labels, columns=labels)
    df_dist = pd.DataFrame(dist, index=labels, columns=labels)

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

    # İnsan-okur “route” sayfası
    planned_route = [labels[i] for i in range(len(labels))]
    optimal_route = [labels[i] for i in opt]
    route_df = pd.DataFrame({
        "planned_order": list(range(1, len(planned_route)+1)),
        "planned_label": planned_route,
        "optimal_order": list(range(1, len(optimal_route)+1)),
        "optimal_label": optimal_route,
    })

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as w:
        df_dur.to_excel(w, sheet_name="durations_min")
        df_dist.to_excel(w, sheet_name="distances_km")
        visits_out.to_excel(w, index=False, sheet_name="visits")
        route_df.to_excel(w, index=False, sheet_name="route")
        kpis.to_excel(w, index=False, sheet_name="kpis")

    print(f"✅ Etiketli rapor hazır: {OUT_XLSX}")

if __name__ == "__main__":
    main()
