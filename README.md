# Rapor Tablosu Oluşturucu

**Amaç:** `yanitlar.xlsx` dosyasını alır, `rapor.xlsx` içindeki **Fonksiyonlar Data** tablosunu **birebir kolonlarla** ve **Excel Table (filtreli)** oluşturur.

## Çalıştırma
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notlar
- Çıktı sayfası adı: **Fonksiyonlar Data**
- Oluşturulan Excel Table adı: **FonksiyonlarData** (filtreler aktif)
- Devamlılık = puan ≥ eşiğe **OK**, aksi **NOK** (varsayılan eşik: 4)
- Test Alanı haritası ve fuzzy sütun eşleşmeleri `transformer.py` içindedir.
