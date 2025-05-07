import os
import pandas as pd
import streamlit as st
from io import BytesIO
import tempfile

def combine_excel_sheets(file_list):
    """
    Menggabungkan semua sheet dari berbagai file Excel dalam satu file Excel baru.
    Kolom "NOMOR AJU" dan "NOMOR IDENTITAS" diformat sebagai string, 
    tetapi tetap membiarkan nilai kosong sebagai NaN.
    """
    data_dict = {}

    if not file_list:
        return None, "Tidak ada file Excel diunggah."

    for uploaded_file in file_list:
        try:
            excel_data = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
            for sheet_name, df in excel_data.items():
                if sheet_name not in data_dict:
                    data_dict[sheet_name] = []

                # Format kolom penting sebagai string tanpa mengubah NaN
                if 'NOMOR AJU' in df.columns:
                    df['NOMOR AJU'] = df['NOMOR AJU'].apply(lambda x: str(x) if pd.notna(x) else x)

                if sheet_name.strip().lower() == 'entitas' and 'NOMOR IDENTITAS' in df.columns:
                    df['NOMOR IDENTITAS'] = df['NOMOR IDENTITAS'].apply(lambda x: str(x) if pd.notna(x) else x)

                if 'Source.Name' in df.columns:
                    df.drop(columns=['Source.Name'], inplace=True)

                data_dict[sheet_name].append(df)

        except Exception as e:
            return None, f"Terjadi error saat membaca file {uploaded_file.name}: {str(e)}"

    if not data_dict:
        return None, "Tidak ada data yang berhasil digabungkan."

    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df_list in data_dict.items():
                combined_df = pd.concat(df_list, ignore_index=True)
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        return output, None
    except Exception as e:
        return None, f"Terjadi error saat menulis file hasil: {str(e)}"

# Streamlit UI
st.title("Combiner Excel - EDII")

uploaded_files = st.file_uploader(
    "Unggah beberapa file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

output_filename = st.text_input("Nama file:", "Combined_Data.xlsx")

if st.button("Gabungkan"):
    if not uploaded_files:
        st.error("Apa yang mau digabung orang kosong.")
    else:
        with st.spinner("Sabang Dor..."):
            result, error = combine_excel_sheets(uploaded_files)
            if error:
                st.error(error)
            else:
                st.success("Udeh kelar!")
                st.download_button(
                    label="📥 Download Dimari",
                    data=result,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.divider()
st.markdown("<p style='text-align: center;'>Made with ♥ By Ibnu P</p>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>May 2025</p>", unsafe_allow_html=True)
    
