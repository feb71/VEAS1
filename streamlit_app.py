import streamlit as st
import pandas as pd
from io import BytesIO

def process_data(df):
    """
    For hver unik S_OBJID:
      - Finn raden med lavest Høyde og oppdater:
          S_FCODE = "KUM"
          Kumform = "R"
          Kjegle = "S"
          Høydereferanse = "BUNN_INNVENDIG"
          Bredde = verdi fra VEAS_VA.Dimensjon (mm)
      - Hvis gruppen har mer enn én rad, finn raden med høyest Høyde og oppdater:
          S_FCODE = "LOK"
          Høydereferanse = "TOPP_UTVENDIG"
          Bredde = verdi fra VEAS_VA.Diameter kumlokk (mm)
    """
    # Sjekk at nødvendige kolonner finnes
    required_cols = [
        "S_OBJID",
        "Høyde",
        "S_FCODE",
        "Kumform",
        "Kjegle",
        "Høydereferanse",
        "Bredde",
        "VEAS_VA.Dimensjon (mm)",
        "VEAS_VA.Diameter kumlokk (mm)"
    ]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Mangler kolonne(r) i Excel: {missing}")
    
    # For hver S_OBJID
    for s_objid, group in df.groupby("S_OBJID"):
        if group.empty:
            continue

        # Finn raden med lavest Høyde
        min_idx = group["Høyde"].idxmin()
        df.at[min_idx, "S_FCODE"] = "KUM"
        df.at[min_idx, "Kumform"] = "R"
        df.at[min_idx, "Kjegle"] = "S"
        df.at[min_idx, "Høydereferanse"] = "BUNN_INNVENDIG"
        df.at[min_idx, "Bredde"] = df.at[min_idx, "VEAS_VA.Dimensjon (mm)"]
        
        # Hvis det finnes mer enn én rad, finn raden med høyest Høyde
        if len(group) > 1:
            max_idx = group["Høyde"].idxmax()
            # Sørg for at den med høyest ikke er den samme som med lavest
            if max_idx != min_idx:
                df.at[max_idx, "S_FCODE"] = "LOK"
                df.at[max_idx, "Høydereferanse"] = "TOPP_UTVENDIG"
                df.at[max_idx, "Bredde"] = df.at[max_idx, "VEAS_VA.Diameter kumlokk (mm)"]
    
    return df

def to_excel(df):
    """
    Konverterer DataFrame til en Excel-fil lagret i minnet (BytesIO)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- Streamlit-app ---
st.title("VEAS Koordinatprosesserer - KUM & LOK")

st.write("""
Last opp en Excel-fil som inneholder følgende kolonner:

- **S_OBJID**
- **Høyde**
- **S_FCODE**
- **Kumform**
- **Kjegle**
- **Høydereferanse**
- **Bredde**
- **VEAS_VA.Dimensjon (mm)**
- **VEAS_VA.Diameter kumlokk (mm)**

Programmet gjør følgende:
- For hver unik S_OBJID finner vi raden med **lavest Høyde** og setter:
  - `S_FCODE = "KUM"`
  - `Kumform = "R"`
  - `Kjegle = "S"`
  - `Høydereferanse = "BUNN_INNVENDIG"`
  - `Bredde` = verdi fra `VEAS_VA.Dimensjon (mm)`

- Dersom gruppen har mer enn én rad, finner vi også raden med **høyest Høyde** og setter:
  - `S_FCODE = "LOK"`
  - `Høydereferanse = "TOPP_UTVENDIG"`
  - `Bredde` = verdi fra `VEAS_VA.Diameter kumlokk (mm)`
""")

uploaded_file = st.file_uploader("Last opp Excel-fil", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.subheader("Innlest data (første 5 rader)")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Feil ved lesing av filen: {e}")
        st.stop()
    
    if st.button("Prosesser data"):
        try:
            result_df = process_data(df)
            st.subheader("Oppdatert data (første 10 rader)")
            st.dataframe(result_df.head(10))
            
            excel_data = to_excel(result_df)
            st.download_button(
                label="Last ned oppdatert Excel-fil",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Feil under behandling av data: {e}")
