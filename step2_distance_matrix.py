# step2_distance_matrix.py
import argparse, pandas as pd, requests
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--infile", required=True)
ap.add_argument("--outfile")
args = ap.parse_args()

df = pd.read_excel(args.infile).dropna(subset=["Latitude","Longitude"])
labels=[f"{i+1}. {str(x).split(',')[0][:35]}" for i,x in enumerate(df["Location"])]
coords=list(zip(df["Latitude"], df["Longitude"]))

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
vis.insert(0,"planned_order", range(1,len(df)+1))
vis.insert(1,"optimal_order", [opt.index(i)+1 for i in range(len(df))])
route_df=pd.DataFrame({"planned_order":range(1,len(labels)+1),
                       "planned_label":labels,
                       "optimal_order":range(1,len(labels)+1),
                       "optimal_label":[labels[i] for i in opt]})
kpis=pd.DataFrame([{
    "planned_total_duration_min": round(plan_min,1),
    "planned_total_distance_km": round(plan_km,2),
    "optimal_total_duration_min": round(opt_min,1),
    "optimal_total_distance_km": round(opt_km,2),
    "time_saving_min": round(plan_min-opt_min,1),
    "time_saving_pct": round(100*(plan_min-opt_min)/plan_min,1) if plan_min>0 else 0.0
}])

base = Path(args.infile).name.replace("_geocoded.xlsx","")
outfile = args.outfile or f"{base}_route_report.xlsx"
with pd.ExcelWriter(outfile, engine="xlsxwriter") as w:
    df_dur.to_excel(w, sheet_name="durations")
    df_dist.to_excel(w, sheet_name="distances")
    vis.to_excel(w, index=False, sheet_name="visits")
    route_df.to_excel(w, index=False, sheet_name="route")
    kpis.to_excel(w, index=False, sheet_name="kpis")
print(f"✅ Rapor: {outfile}")
