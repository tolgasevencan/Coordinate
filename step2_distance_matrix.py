# step2_distance_matrix.py
import argparse, pandas as pd, requests
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--infile", required=True, help="Eingabedatei (Excel mit Koordinaten)")
ap.add_argument("--outfile", help="Optionaler Ausgabepfad")
args = ap.parse_args()

df = pd.read_excel(args.infile).dropna(subset=["Breitengrad","Längengrad"])
labels=[f"{i+1}. {str(x).split(',')[0][:35]}" for i,x in enumerate(df["Ort"])]
coords=list(zip(df["Breitengrad"], df["Längengrad"]))

print(f"🧭 Berechne Distanzmatrix für {len(coords)} Standorte...")

coord_str=";".join([f"{lon},{lat}" for lat,lon in coords])
js=requests.get(f"https://router.project-osrm.org/table/v1/driving/{coord_str}?annotations=duration,distance",timeout=30).json()
dur=[[round((d or 0)/60,1) for d in row] for row in js["durations"]]
dist=[[round((d or 0)/1000,2) for d in row] for row in js["distances"]]

def nn(D, start=0):
    n=len(D); left=set(range(n)); left.remove(start); r=[start]
    while left: j=min(left, key=lambda x:D[r[-1]][x]); r.append(j); left.remove(j)
    return r
def two_opt(route, D):
    def L(rt): return sum(D[rt[i]][rt[i+1]] for i in range(len(rt)-1))
    best=route[:]; improved=True
    while improved:
        improved=False
        for i in range(1,len(best)-2):
            for k in range(i+1,len(best)-1):
                new=best[:i]+best[i:k+1][::-1]+best[k+1:]
                if L(new) < L(best): best=new; improved=True
    return best

opt=two_opt(nn(dist), dist)
plan_min=sum(dur[i][i+1] for i in range(len(dur)-1))
opt_min=sum(dur[opt[i]][opt[i+1]] for i in range(len(opt)-1))
plan_km=sum(dist[i][i+1] for i in range(len(dist)-1))
opt_km=sum(dist[opt[i]][opt[i+1]] for i in range(len(opt)-1))

df_dur=pd.DataFrame(dur, index=labels, columns=labels)
df_dist=pd.DataFrame(dist, index=labels, columns=labels)
vis=df.copy()
vis.insert(0,"Geplante Reihenfolge", range(1,len(df)+1))
vis.insert(1,"Optimale Reihenfolge", [opt.index(i)+1 for i in range(len(df))])
route_df=pd.DataFrame({"Geplante Reihenfolge":range(1,len(labels)+1),
                       "Geplanter Standort":labels,
                       "Optimale Reihenfolge":range(1,len(labels)+1),
                       "Optimaler Standort":[labels[i] for i in opt]})
kpis=pd.DataFrame([{
    "Geplante Gesamtdauer (Minuten)": round(plan_min,1),
    "Geplante Distanz (km)": round(plan_km,2),
    "Optimale Gesamtdauer (Minuten)": round(opt_min,1),
    "Optimale Distanz (km)": round(opt_km,2),
    "Ersparnis (Minuten)": round(plan_min-opt_min,1),
    "Ersparnis (%)": round(100*(plan_min-opt_min)/plan_min,1) if plan_min>0 else 0.0
}])

base = Path(args.infile).name.replace("_geocoded.xlsx","")
outfile = args.outfile or f"{base}_route_report.xlsx"
with pd.ExcelWriter(outfile, engine="xlsxwriter") as w:
    df_dur.to_excel(w, sheet_name="Dauer (Minuten)")
    df_dist.to_excel(w, sheet_name="Distanz (km)")
    vis.to_excel(w, index=False, sheet_name="Besuche")
    route_df.to_excel(w, index=False, sheet_name="Route")
    kpis.to_excel(w, index=False, sheet_name="KPIs")

print(f"✅ Routenanalyse abgeschlossen: {outfile}")
print(f"⏱️ Geplante Dauer: {plan_min:.1f} Min → Optimiert: {opt_min:.1f} Min")
print(f"📉 Ersparnis: {round(plan_min-opt_min,1)} Min ({round(100*(plan_min-opt_min)/plan_min,1) if plan_min>0 else 0} %)")
