# Rapor Otomasyonu — yanitlar.xlsx → rapor.xlsx

Uygulama, yanitlar.xlsx içindeki tüm sekmeleri analiz eder; Fonksiyonlar Data ve UploadDownload Data tablolarını üretir. 
İki çıktı sunar:
1) Taşınabilir Excel: xlsxwriter ile, her sayfada Excel Table (filtreli)
2) Şablona Yazılmış Excel: rapor.xlsx şablonuna yazıp tablo aralıklarını büyütür; filtre/biçimler korunur.

Kurulum:
pip install -r requirements.txt
streamlit run app.py

Not: Şablonda pivot/slicer varsa, veri güncelledikten sonra Excel içinde “Yenile” yapmanız gerekebilir.
