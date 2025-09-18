# -*- coding: utf-8 -*-
from __future__ import annotations
import pandas as pd, numpy as np, re, unicodedata, difflib
from datetime import datetime

FONK_COLS_DEFAULT = [
    'Faz','Column1','Katılımcı','Devamlılık','Tarih','Test Alanı',
    'Cihaz OS','Uygulama','wifi/lte','Versiyon','Puan',
    'Bip Yorum','Whatsapp yorum','Telegram Yorum','cihaz'
]
UPDOWN_COLS_DEFAULT = [
    'Faz','Katılımcı','Devamlılık','Cihaz OS','Tarih','Uygulama',
    'Upload/Download','Gönderim tipi','Dosya Boyutu','Karşıdaki boyut',
    'Süre (sn)','Hız (mb/sn)','versiyon','wifi/lte'
]

TEST_ALANI_MAP = {
    "txt": "IM - 1-1 txt mesaj",
    "gm": "IM - Grup mesajlaşması",
    "im": "IM - Genel",
    "media": "IM - Medya paylaşımı",
    "voip": "Voip - 1-1 Görüntülü görüşme",
    "gsg": "Voip - Grup sesli görüşme",
    "ggg": "Voip - Grup görüntülü görüşme",
    "call": "Voip - 1-1 Sesli görüşme",
}

APPS = [("Bip","bip"), ("Whatsapp","whatsapp"), ("Telegram","telegram")]

def _norm(s: str) -> str:
    if s is None: return ""
    s = str(s)
    tr = str.maketrans({"ı":"i","İ":"i","ş":"s","Ş":"s","ğ":"g","Ğ":"g","ü":"u","Ü":"u","ö":"o","Ö":"o","ç":"c","Ç":"c"})
    s = s.translate(tr).lower().strip()
    s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)

def _best(needle: str, haystack: list[str]) -> str | None:
    norm = _norm(needle).replace(" ", "")
    cand = [_norm(c).replace(" ", "") for c in haystack]
    if not cand: return None
    import difflib
    m = difflib.get_close_matches(norm, cand, n=1, cutoff=0.7)
    return haystack[cand.index(m[0])] if m else None

def _excel_serial(dt) -> int | None:
    if pd.isna(dt): return None
    try:
        ts = pd.to_datetime(dt)
        return int((ts - pd.Timestamp("1899-12-30")).days)
    except Exception:
        return None

def build_from_yanitlar(xls: pd.ExcelFile, faz_value: str = "Faz 6",
                        devamlilik_threshold: int = 4,
                        fonk_cols: list[str] | None = None,
                        updown_cols: list[str] | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    sheets = xls.sheet_names
    dfs = {name: xls.parse(name) for name in sheets}

    # Heuristik: Puan içeren formlar
    form_candidates = [df for df in dfs.values() if any(" Puan" in c for c in df.columns)]
    # Heuristik: Upload/Download formu
    form2_candidates = [df for df in dfs.values() if any("Gönderilen fotoğraf" in c or "Video gönderme" in c for c in df.columns)]
    if not form_candidates and dfs:
        form_candidates = [list(dfs.values())[0]]
    if not form2_candidates and len(dfs) > 1:
        form2_candidates = [list(dfs.values())[1]]

    # ---------- Fonksiyonlar Data ----------
    fonk_rows = []
    id_map, next_id = {}, 1

    def proc_form(df: pd.DataFrame):
        nonlocal next_id
        cols = list(df.columns)
        col_name = _best("Adınız, soyadınız", cols) or _best("Ad Soyad", cols) or _best("Katılımcı", cols)
        col_date = _best("Tarih", cols) or _best("Zaman damgası", cols)
        col_conn = _best("Bağlantı türü", cols) or _best("Bağlantı", cols)
        ver_cols = {
            "Bip": _best("Bip Uygulama Versiyon", cols) or _best("Bip Versiyon", cols),
            "Whatsapp": _best("Whatsapp Uygulama Versiyon", cols) or _best("Whatsapp Versiyon", cols),
            "Telegram": _best("Telegram Uygulama Versiyon", cols) or _best("Telegram Versiyon", cols),
        }
        for col in cols:
            if not col.endswith(" Puan"):
                continue
            m = re.match(r'^(Bip|Whatsapp|Telegram)\s+(.+)\s+Puan$', col)
            if not m:
                continue
            app, scen = m.groups()
            scen_code = _norm(scen)
            scen_code = re.sub(r'[^a-z0-9]+','',scen_code)
            test_alani = TEST_ALANI_MAP.get(scen_code, scen)
            comment_col = None
            idx = cols.index(col)
            if idx+1 < len(cols) and 'yorum' in _norm(cols[idx+1]):
                comment_col = cols[idx+1]
            for i, val in df[col].items():
                if pd.isna(val) or str(val) == "":
                    continue
                try:
                    puan = int(val)
                except Exception:
                    try:
                        puan = int(float(str(val).replace(",", ".")))
                    except Exception:
                        continue
                name = str(df.at[i, col_name]) if col_name in df.columns else ""
                if not name:
                    continue
                if name not in id_map:
                    id_map[name] = str(next_id); next_id += 1
                tarih = _excel_serial(df.at[i, col_date]) if col_date in df.columns else None
                conn  = str(df.at[i, col_conn]) if col_conn in df.columns else None
                vers  = str(df.at[i, ver_cols[app]]) if ver_cols[app] in df.columns and pd.notna(df.at[i, ver_cols[app]]) else ""
                devam = "OK" if puan >= devamlilik_threshold else "NOK"
                yorum = ""
                if comment_col and comment_col in df.columns:
                    cv = df.at[i, comment_col]
                    if pd.notna(cv): yorum = str(cv)
                fonk_rows.append({
                    'Faz': faz_value, 'Column1': id_map[name], 'Katılımcı': name,
                    'Devamlılık': devam, 'Tarih': tarih, 'Test Alanı': test_alani,
                    'Cihaz OS': None, 'Uygulama': app.lower(), 'wifi/lte': conn,
                    'Versiyon': vers, 'Puan': puan,
                    'Bip Yorum': yorum if app=="Bip" else None,
                    'Whatsapp yorum': yorum if app=="Whatsapp" else None,
                    'Telegram Yorum': yorum if app=="Telegram" else None,
                    'cihaz': None
                })

    for df in form_candidates:
        proc_form(df)

    fonk_df = pd.DataFrame(fonk_rows)
    if fonk_df.empty:
        fonk_df = pd.DataFrame(columns=FONK_COLS_DEFAULT)

    if fonk_cols:
        for c in fonk_cols:
            if c not in fonk_df.columns: fonk_df[c] = None
        fonk_df = fonk_df[fonk_cols]
    else:
        for c in FONK_COLS_DEFAULT:
            if c not in fonk_df.columns: fonk_df[c] = None
        fonk_df = fonk_df[FONK_COLS_DEFAULT]
    fonk_df = fonk_df.sort_values(["Katılımcı","Tarih","Uygulama","Test Alanı"], kind="stable").reset_index(drop=True)

    # ---------- UploadDownload Data ----------
    up_rows = []
    def tofloat(x):
        if pd.isna(x) or x is None: return None
        try: return float(x)
        except:
            try: return float(str(x).replace(",", "."))
            except: return None

    form2_candidates = form2_candidates or []
    for df in form2_candidates:
        cols = list(df.columns)
        col_name = _best("Adınız, soyadınız", cols) or _best("Ad Soyad", cols) or _best("Katılımcı", cols)
        col_date = _best("Tarih", cols) or _best("Zaman damgası", cols)
        col_conn = _best("Bağlantı türü", cols) or _best("Bağlantı", cols)
        ver_cols = {
            "Bip": _best("Bip Uygulama Versiyon", cols) or _best("Bip Versiyon", cols),
            "Whatsapp": _best("Whatsapp Uygulama Versiyon", cols) or _best("Whatsapp Versiyon", cols),
            "Telegram": _best("Telegram Uygulama Versiyon", cols) or _best("Telegram Versiyon", cols),
        }
        g_photo = _best("Gönderilen fotoğraf boyutu", cols) or _best("Gönderilen fotoğraf boyutu (mb cinsinden)", cols)
        g_video = _best("Gönderilen video boyutu", cols) or _best("Gönderilen video boyutu (mb cinsinden)", cols)

        for i in range(len(df)):
            name = str(df.at[i, col_name]) if col_name in df.columns else ""
            if not name: continue
            tarih = _excel_serial(df.at[i, col_date]) if col_date in df.columns else None
            conn  = str(df.at[i, col_conn]) if col_conn in df.columns else None
            gphoto = tofloat(df.at[i, g_photo]) if g_photo in df.columns else None
            gvideo = tofloat(df.at[i, g_video]) if g_video in df.columns else None

            for app, app_l in APPS:
                vers = str(df.at[i, ver_cols[app]]) if ver_cols[app] in df.columns and pd.notna(df.at[i, ver_cols[app]]) else ""
                # Foto upload
                if gphoto is not None:
                    remote = tofloat(df.at[i, _best(f"{app} Alıcıya ulaşan fotoğraf boyutu", cols) or _best(f"{app} Alıcıya ulaşan fotoğraf boyutu (mb)", cols)])
                    tsend  = tofloat(df.at[i, _best(f"{app} Fotoğraf gönderme süresi", cols) or _best(f"{app} Fotoğraf gönderme süresi (sn)", cols)])
                    speed  = (gphoto/tsend) if (tsend and tsend>0) else None
                    up_rows.append({
                        'Faz': faz_value,'Katılımcı': name,'Devamlılık': "OK",'Cihaz OS': None,'Tarih': tarih,
                        'Uygulama': app_l,'Upload/Download': 'upload','Gönderim tipi': 'fotoğraf',
                        'Dosya Boyutu': gphoto,'Karşıdaki boyut': remote,'Süre (sn)': tsend,'Hız (mb/sn)': speed,
                        'versiyon': vers,'wifi/lte': conn
                    })
                # Foto download
                rtime   = tofloat(df.at[i, _best(f"{app} Fotoğraf alma süresi", cols)])
                rlocal  = tofloat(df.at[i, _best(f"{app} Alınan fotoğraf boyutu", cols) or _best(f"{app} Alınan fotoğraf boyutu (mb)", cols)])
                rremote = tofloat(df.at[i, _best(f"{app} Kaynak fotoğraf boyutu", cols) or _best(f"{app} Kaynak fotoğraf boyutu (mb)", cols)])
                speed2  = (rlocal/rtime) if (rlocal and rtime and rtime>0) else None
                if (rtime is not None) or (rlocal is not None) or (rremote is not None):
                    up_rows.append({
                        'Faz': faz_value,'Katılımcı': name,'Devamlılık': "OK",'Cihaz OS': None,'Tarih': tarih,
                        'Uygulama': app_l,'Upload/Download': 'download','Gönderim tipi': 'fotoğraf',
                        'Dosya Boyutu': rlocal,'Karşıdaki boyut': rremote,'Süre (sn)': rtime,'Hız (mb/sn)': speed2,
                        'versiyon': vers,'wifi/lte': conn
                    })
                # Video upload
                if gvideo is not None:
                    remote_v = tofloat(df.at[i, _best(f"{app} Alıcıya ulaşan video boyutu", cols) or _best(f"{app} Alıcıya ulaşan video boyutu (mb)", cols)])
                    tsend_v  = tofloat(df.at[i, _best(f"{app} Video gönderme süresi", cols) or _best(f"{app} Video gönderme süresi (sn)", cols)])
                    speed_v  = (gvideo/tsend_v) if (tsend_v and tsend_v>0) else None
                    up_rows.append({
                        'Faz': faz_value,'Katılımcı': name,'Devamlılık': "OK",'Cihaz OS': None,'Tarih': tarih,
                        'Uygulama': app_l,'Upload/Download': 'upload','Gönderim tipi': 'video',
                        'Dosya Boyutu': gvideo,'Karşıdaki boyut': remote_v,'Süre (sn)': tsend_v,'Hız (mb/sn)': speed_v,
                        'versiyon': vers,'wifi/lte': conn
                    })
                # Video download
                rtime_v   = tofloat(df.at[i, _best(f"{app} Video alma süresi", cols)])
                rlocal_v  = tofloat(df.at[i, _best(f"{app} Alınan video boyutu", cols) or _best(f"{app} Alınan video boyutu (mb)", cols)])
                rremote_v = tofloat(df.at[i, _best(f"{app} Kaynak video boyutu", cols) or _best(f"{app} Kaynak video boyutu (mb)", cols)])
                speed_v2  = (rlocal_v/rtime_v) if (rlocal_v and rtime_v and rtime_v>0) else None
                if (rtime_v is not None) or (rlocal_v is not None) or (rremote_v is not None):
                    up_rows.append({
                        'Faz': faz_value,'Katılımcı': name,'Devamlılık': "OK",'Cihaz OS': None,'Tarih': tarih,
                        'Uygulama': app_l,'Upload/Download': 'download','Gönderim tipi': 'video',
                        'Dosya Boyutu': rlocal_v,'Karşıdaki boyut': rremote_v,'Süre (sn)': rtime_v,'Hız (mb/sn)': speed_v2,
                        'versiyon': vers,'wifi/lte': conn
                    })

    up_df = pd.DataFrame(up_rows)
    if up_df.empty:
        up_df = pd.DataFrame(columns=UPDOWN_COLS_DEFAULT)

    if updown_cols:
        for c in updown_cols:
            if c not in up_df.columns: up_df[c] = None
        up_df = up_df[updown_cols]
    else:
        for c in UPDOWN_COLS_DEFAULT:
            if c not in up_df.columns: up_df[c] = None
        up_df = up_df[UPDOWN_COLS_DEFAULT]

    return fonk_df, up_df
# --- geriye dönük uyumluluk için alias ---
def transform(df_raw, faz_value=None, devamlilik_threshold=4):
    """
    reportcaster/streamlit_app.py bu imzayı bekliyor.
    Girdi: df_raw (DataFrame)
    Çıktı: Fonksiyonlar Data (DataFrame)
    """
    try:
        # Eğer dosyanda bu fonksiyon varsa onu kullan:
        return transform_yanitlar_to_table(df_raw, faz_value=faz_value, devamlilik_threshold=devamlilik_threshold)
    except NameError:
        # Alternatif isimli fonksiyonun varsa onu çağır (ör: _core_transform, build_* vs.)
        return _core_transform(df_raw, faz_value=faz_value, devamlilik_threshold=devamlilik_threshold)
