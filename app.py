# -*- coding: utf-8 -*-
import streamlit as st, pandas as pd
from io import BytesIO
from transformer import build_from_yanitlar
from writer import write_into_template, write_portable_with_tables

st.set_page_config(page_title="Rapor Otomasyonu", layout="wide")
st.title("📊 Rapor Otomasyonu — yanitlar.xlsx → rapor.xlsx benzeri çıktı")

with st.sidebar:
    st.header("Ayarlar")
    faz_value = st.text_input("Faz", value="Faz 6")
    esik = st.number_input("Devamlılık eşiği (OK için min puan)", 1, 5, 4, 1)

st.markdown("**1) yanitlar.xlsx** ve **2) rapor.xlsx (şablon)** dosyalarını yükleyin.")
col1, col2 = st.columns(2)
with col1:
    yan_file = st.file_uploader("yanitlar.xlsx", type=["xlsx"], key="yan")
with col2:
    tpl_file = st.file_uploader("rapor.xlsx (şablon)", type=["xlsx"], key="tpl")

if not yan_file:
    st.info("Önce yanitlar.xlsx dosyasını yükleyin.")
    st.stop()

xls = pd.ExcelFile(yan_file)

# Şablondan kolon başlıklarını almak → birebir sıra/isim garantisi
fonk_cols = up_cols = None
if tpl_file:
    try:
        fonk_cols = list(pd.read_excel(tpl_file, sheet_name="Fonksiyonlar Data").columns)
        up_cols   = list(pd.read_excel(tpl_file, sheet_name="UploadDownload Data").columns)
    except Exception as e:
        st.warning(f"Şablondan kolon başlıkları alınamadı: {e}")

fonk_df, up_df = build_from_yanitlar(
    xls, faz_value=faz_value, devamlilik_threshold=int(esik),
    fonk_cols=fonk_cols, updown_cols=up_cols
)

st.subheader("Fonksiyonlar Data (örnek 50)")
st.dataframe(fonk_df.head(50), use_container_width=True)
st.subheader("UploadDownload Data (örnek 50)")
st.dataframe(up_df.head(50), use_container_width=True)

st.markdown("---")
st.markdown("### 📥 Çıktı")

# 1) Taşınabilir Excel (tablolar & filtreler)
port_bytes = write_portable_with_tables(fonk_df, up_df)
st.download_button("⬇️ Taşınabilir Excel (tablolar & filtreler)", data=port_bytes,
                   file_name="rapor_portable.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 2) Şablona yaz (rapor.xlsx yapısı korunur — tablo aralıkları güncellenir)
if tpl_file:
    try:
        tpl_bytes = tpl_file.read()
        out_bytes = write_into_template(tpl_bytes, fonk_df, up_df)
        st.download_button("⬇️ Şablona Yazılmış Excel (rapor yapısı korunur)", data=out_bytes,
                           file_name="rapor_sablonlu.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Şablona yazılamadı: {e}")
        st.info("Yedek olarak 'Taşınabilir Excel' dosyasını kullanabilirsiniz.")
else:
    st.info("Şablona birebir yazım için rapor.xlsx dosyasını da yükleyin.")
