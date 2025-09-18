# -*- coding: utf-8 -*-
import streamlit as st, pandas as pd
from io import BytesIO
from transformer import transform_yanitlar_to_table

st.set_page_config(page_title="Rapor Tablosu Oluşturucu", layout="wide")
st.title("📊 Rapor Tablosu Oluşturucu")
st.caption("yanitlar.xlsx → rapor.xlsx içindeki 'Fonksiyonlar Data' tablosunu **%100 aynı kolonlarla** ve **Excel Table (filtreli)** oluşturur.")

with st.sidebar:
    st.header("Ayarlar")
    faz_value = st.text_input("Faz", value="Faz 6")
    devam_esik = st.number_input("Devamlılık eşiği (OK için minimum puan)", min_value=1, max_value=5, value=4, step=1)

file_up = st.file_uploader("yanitlar.xlsx dosyasını yükleyin", type=["xlsx"])

if not file_up:
    st.info("Başlamak için yanitlar.xlsx yükleyin.")
    st.stop()

# Kaynak oku
xls = pd.ExcelFile(file_up)
df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

st.subheader("Kaynak önizleme (ilk 20 satır)")
st.dataframe(df_raw.head(20), use_container_width=True)

# Dönüştür
out_df = transform_yanitlar_to_table(df_raw, faz_value=faz_value, devamlilik_threshold=devam_esik)

st.subheader("Çıktı önizleme (ilk 50 satır)")
st.dataframe(out_df.head(50), use_container_width=True)

# Excel Table (filtreli) olarak indir
from xlsxwriter.utility import xl_range
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    sheet_name = "Fonksiyonlar Data"
    out_df.to_excel(writer, sheet_name=sheet_name, index=False)
    wb = writer.book
    ws = writer.sheets[sheet_name]

    nrows, ncols = out_df.shape
    # add_table uses inclusive coordinates (0-based)
    ws.add_table(0, 0, nrows, ncols-1, {
        "name": "FonksiyonlarData",
        "columns": [{"header": c} for c in out_df.columns],
        "autofilter": True
    })

buffer.seek(0)
st.download_button("⬇️ Excel'i indir (filtreli tablo)", data=buffer,
                   file_name="rapor_tablosu.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
