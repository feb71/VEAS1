import streamlit as st
import pandas as pd
from io import BytesIO

def process_data(df):
    """
    For hver unik S_OBJID:
    - Finn raden med lavest Høyde og sett S_FCODE = "KUM"
    - Finn raden med høyest Høyde og sett S_FCODE = "LOK"
    Dersom begge er samme rad (kun én rad i gruppa), gis den S_FCODE = "KUM"
    """
    group_col = "S_OBJID"
    selected_rows = []
    
    # Gruppér på S_OBJID
    for s_objid, group in df.groupby(group_col):
        # Finn raden med lavest Høyde
        min_idx = group["Høyde"].idxmin()
        min_row = group.loc[min_idx].copy()
        min_row["S_FCODE"] = "KUM"
        selected_rows.append(min_row)
        
        # Finn raden med høyest Høyde, hvis den er en annen rad
        max_idx = group["Høyde"].idxmax()
        if max_idx != min_idx:
            max_row = group.loc[max_idx].copy()
            max_row["S_FCODE"] = "LOK"
            selected_rows.append(max_row)
    
    result_df = pd.DataFrame(selected_rows)
    return result_df

def to_excel(df):
    """
    Konverter DataFrame til en Excel-fil i minnet (BytesIO)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

st.title("Excel Koordinatprosesserer")

st.write("""
Last opp en Excel-fil med kolonnene:
- **S_OBJID**
- **Øst**
- **Nord**
- **Høyde**

Programmet grupperer dataene basert på S_OBJID og setter:
- Den med lavest Høyde får **S_FCODE = "KUM"**
- Den med høyest Høyde får **S_FCODE = "LOK"**
""")

# Filopplasting
uploaded_file = st.file_uploader("Last opp Excel-fil", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.subheader("Input-data")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Feil ved lesing av filen: {e}")
    
    if st.button("Prosesser data"):
        try:
            result_df = process_data(df)
            st.subheader("Prosessert data")
            st.dataframe(result_df)
            
            # Konverter til Excel og legg til en nedlastingsknapp
            excel_data = to_excel(result_df)
            st.download_button(
                label="Last ned prosessert Excel-fil",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Feil under behandling av data: {e}")
