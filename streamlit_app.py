# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from transformer import transform  # <— ImportError buradan çözülüyor

st.set_page_config(page_title="ReportCaster — Rapor Tablosu", layout="wide")
st.title("📊 ReportCaster — Fonksiyonlar Data Tablosu")
st.caption("yanitlar.xlsx → rapor.xlsx içindeki 'Fonksiyonlar Data' kolonlarıyla, filtreli Excel Table çıktısı üretir.")

with st.sidebar:
    st.header("Ayarlar")
    faz_value = st.text_input("Faz", value="Faz 6")
    devam_esik = st.number_input("Devamlılık eşiği (OK için min puan)", min_value=1, max_value=5, value=4, step=1)

src_file = st.file_uploader("yanitlar.xlsx dosyasını yükleyin", type=["xlsx"])

if not src_file:
    st.info("Başlamak için yanitlar.xlsx yükleyin.")
    st.stop()

# Kaynağı oku
xls = pd.ExcelFile(src_file)
df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

st.subheader("Kaynak önizleme (ilk 20)")
st.dataframe(df_raw.head(20), use_container_width=True)

# Dönüştür
out_df = transform(df_raw, faz_value=faz_value, devamlilik_threshold=int(devam_esik))

st.subheader("Çıktı önizleme (ilk 50)")
st.dataframe(out_df.head(50), use_container_width=True)

# Excel Table (filtreli) olarak indir
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    sheet_name = "Fonksiyonlar Data"
    out_df.to_excel(writer, sheet_name=sheet_name, index=False)
    wb = writer.book
    ws = writer.sheets[sheet_name]

    nrows, ncols = out_df.shape
    # add_table: inclusive 0-based koordinatlar
    ws.add_table(0, 0, nrows, ncols-1, {
        "name": "FonksiyonlarData",
        "columns": [{"header": c} for c in out_df.columns],
        "autofilter": True
    })

buffer.seek(0)
st.download_button("⬇️ Excel'i indir (filtreli tablo)", data=buffer,
                   file_name="rapor_tablosu.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
