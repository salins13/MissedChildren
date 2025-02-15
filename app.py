import pandas as pd
import streamlit as st

def process_excel(input_file):
    # Read Excel file
    df = pd.read_excel(input_file, dtype=str)
    
    # Convert date columns explicitly
    df["mon_date"] = pd.to_datetime(df["mon_date"], errors="coerce").dt.date
    df["details_dob"] = pd.to_datetime(df["details_dob"], errors="coerce").dt.date
    
    # Convert details_age_in_month to integer
    df["details_age_in_month"] = pd.to_numeric(df["details_age_in_month"], errors="coerce").fillna(0).astype(int)
    
    # Drop rows where 'at_16m_reason_missed_vaccination' is NaN
    df = df.dropna(subset=['at_16m_reason_missed_vaccination'])
    
    # Select required columns
    selected_columns = [
        "mon_date", "block", "phc_uphc", "subcentre_uhp", "session_site", "vill_moh_ward", 
        "details_child_name", "details_mother_father_name", "details_child_sex", "details_dob", "details_age_in_month",
        "birth_bcg", "at_6w_opv1", "at_6w_rota1", "at_6w_ipv_f_IPV1", "at_6w_pcv1", "at_6w_penta1",
        "at_10w_opv2", "at_10w_rota2", "at_10w_penta2",
        "at_14w_opv3", "at_14w_rota3", "at_14w_ipv_f_IPV2", "at_14w_pcv2", "at_14w_penta3",
        "at_9m_f_IPV3", "at_9m_mcv1_mr1", "at_9m_pcv3", "at_9m_je1",
        "at_16m_opvb", "at_16m_mcv2_mr2", "at_16m_je2", "at_16m_dptb",
        "at_16m_child_due_any_dose"
    ]
    df_selected = df[selected_columns]
    
    # Save processed data into an Excel file with separate sheets for each block
    output_file = "processed_data.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for block in df_selected["block"].dropna().unique():
            df_selected[df_selected["block"] == block].to_excel(writer, sheet_name=str(block), index=False)
    
    return output_file

# Streamlit UI
st.title("Excel Data Processor")
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file:
    st.write("Processing...")
    output_path = process_excel(uploaded_file)
    st.success("Processing complete!")
    with open(output_path, "rb") as f:
        st.download_button("Download Processed Excel", f, file_name="processed_data.xlsx")
