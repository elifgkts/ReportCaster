
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from transformer import transform_exact, TEMPLATE_COLS

st.set_page_config(page_title="Form → Fonksiyonlar (Birebir Şablon)", layout="wide")
st.title("📑 Form → 📊 Fonksiyonlar Data (Birebir Şablon)")

st.caption("Çıktı, **template.xlsx** içindeki 'Fonksiyonlar Data' sayfasının sütunları ve çalışma kitabı yapısı ile birebir aynı şekilde oluşturulur.")

with st.sidebar:
    st.header("Ayarlar")
    faz_value = st.text_input("Faz", value="Faz 6")
    st.markdown("İsterseniz kendi şablonunuzu yükleyin (boş bırakılırsa repo'daki `template.xlsx` kullanılır):")
    tpl_file = st.file_uploader("Şablon (2 numaralı format)", type=["xlsx"], key="tpl")

src_file = st.file_uploader("1 numaralı formatta Excel yükleyin (.xlsx)", type=["xlsx"], key="src")

if src_file:
    try:
        xls = pd.ExcelFile(src_file)
        df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    except Exception as e:
        st.error(f"Kaynak Excel okunamadı: {e}")
        st.stop()

    st.subheader("Kaynak örnek")
    st.dataframe(df_raw.head(20), use_container_width=True)

    out_df = transform_exact(df_raw, faz_value=faz_value)

    st.subheader("Dönüştürülmüş veri (önizleme)")
    st.dataframe(out_df.head(50), use_container_width=True)

    # --- Write into template workbook, preserving other sheets ---
    if tpl_file:
        tpl_bytes = tpl_file.read()
    else:
        with open("template.xlsx", "rb") as f:
            tpl_bytes = f.read()

    wb = load_workbook(BytesIO(tpl_bytes))
    if "Fonksiyonlar Data" not in wb.sheetnames:
        st.error("Şablonda 'Fonksiyonlar Data' sayfası bulunamadı.")
        st.stop()
    ws = wb["Fonksiyonlar Data"]

    # Clear existing data below header (assume headers in row 1)
    ws.delete_rows(2, ws.max_row)

    # Ensure headers exactly as template's existing row 1
    template_headers = [cell.value for cell in ws[1]]
    if template_headers != TEMPLATE_COLS:
        # If they differ, rewrite header to match what our transform produced
        for col_idx, name in enumerate(TEMPLATE_COLS, start=1):
            ws.cell(row=1, column=col_idx, value=name)
        # Remove any extra trailing columns from template header
        if ws.max_column > len(TEMPLATE_COLS):
            ws.delete_cols(len(TEMPLATE_COLS)+1, ws.max_column - len(TEMPLATE_COLS))

    # Write data rows
    for r_idx, (_, r) in enumerate(out_df.iterrows(), start=2):
        for c_idx, col in enumerate(TEMPLATE_COLS, start=1):
            ws.cell(row=r_idx, column=c_idx, value=r.get(col, None))

    # Produce a download file
    out_bytes = BytesIO()
    wb.save(out_bytes)
    out_bytes.seek(0)

    st.download_button("⬇️ Şablona yazılmış Excel'i indir", data=out_bytes, file_name="rapor_birebir.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Başlamak için 1 numaralı formatta bir Excel yükleyin.")

st.markdown("---")
st.caption("Notlar: Şablon yüklemezseniz repo'daki `template.xlsx` kullanılır. Çıktı, şablonun çalışma kitabı yapısını korur; 'Fonksiyonlar Data' sayfasının verileri başlıktan itibaren yenilenir.")
