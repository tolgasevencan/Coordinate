# ───────────────────────────────────────────────────────────────────────────────
# BENUTZEROBERFLÄCHE – ICS → (parse) → (geocode) → Excel-Route-Report in 1 Schritt

st.header("1️⃣ ICS hochladen und vollständigen Routenreport erstellen")

uploaded_ics = st.file_uploader("Outlook-ICS-Datei hochladen", type=["ics"])

if uploaded_ics is not None:
    # Basisname für das Export-File
    base_name = Path(uploaded_ics.name).stem or "kalender"

    try:
        ics_bytes = uploaded_ics.read()

        with st.spinner("📆 ICS wird analysiert..."):
            df_raw = parse_ics_to_df(ics_bytes)
            st.subheader("Aus dem Kalender gelesene Termine")
            st.dataframe(df_raw)

        with st.spinner("📍 Adressen werden geokodiert..."):
            df_geo = geocode_df(df_raw)
            st.subheader("Termine mit Geokoordinaten")
            st.dataframe(df_geo)

        with st.spinner("🧮 Routenoptimierung & Excel-Report wird erstellt..."):
            _ = build_daywise_report(df_geo, base_name)

    except Exception as e:
        st.error(f"Fehler bei der Verarbeitung der ICS-Datei: {e}")