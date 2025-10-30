import io, pandas as pd, streamlit as st
from ics import Calendar
from zoneinfo import ZoneInfo

TZ = ZoneInfo("Europe/Zurich")

st.set_page_config(page_title="Aussendienst Analyse", page_icon="🚗")
st.title("🚗 Aussendienst – ICS → Excel (v1)")
st.caption("ICS yükle • Takvim etkinliklerini tabloya çevir • Excel indir")

def parse_ics(ics_bytes: bytes) -> pd.DataFrame:
    cal = Calendar(ics_bytes.decode("utf-8", errors="ignore"))
    rows = []
    for e in cal.events:
        s, t = e.begin.datetime, e.end.datetime
        if s.tzinfo: s = s.astimezone(TZ)
        if t.tzinfo: t = t.astimezone(TZ)
        rows.append({
            "Betreff": e.name or "",
            "Startdatum": s.date(),
            "Startzeit": s.strftime("%H:%M:%S"),
            "Enddatum": t.date(),
            "Endzeit": t.strftime("%H:%M:%S"),
            "Ort": e.location or "",
            "Beschreibung": e.description or ""
        })
    return pd.DataFrame(rows).sort_values(["Startdatum","Startzeit"])

uploaded = st.file_uploader("ICS-Datei hochladen (z.B. 20250901-20250930_Name.ics)", type=["ics"])
if uploaded:
    df = parse_ics(uploaded.getvalue())
    st.success(f"Toplam etkinlik: {len(df)}")
    st.dataframe(df, use_container_width=True)

    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    st.download_button(
        "📥 Excel indir",
        data=buf.getvalue(),
        file_name=uploaded.name.replace(".ics", "_export.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Lütfen bir ICS dosyası yükleyin.")
