# -*- coding: utf-8 -*-
from __future__ import annotations
import pandas as pd, numpy as np, re, unicodedata, difflib
from datetime import datetime

# Target schema (rapor.xlsx → "Fonksiyonlar Data")
RAPOR_COLUMNS = [
    'Faz','Column1','Katılımcı','Devamlılık','Tarih','Test Alanı',
    'Cihaz OS','Uygulama','wifi/lte','Versiyon','Puan',
    'Bip Yorum','Whatsapp yorum','Telegram Yorum','cihaz'
]

# Mapping for feature names in yanitlar.xlsx to "Test Alanı" display values
TEST_ALANI_MAP = {
    "txt": "IM - 1-1 txt mesaj",
    "gm": "IM - Grup mesajlaşması",
    "call": "Voip - 1-1 Sesli görüşme",
    "media": "IM - Medya paylaşımı",
    "im": "IM - Genel",
    "voip": "Voip - 1-1 Görüntülü görüşme",
}

# App labels
APPS = [("Bip","bip"), ("Whatsapp","whatsapp"), ("Telegram","telegram")]

def _normalize_tr(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    tr_map = str.maketrans({"ı":"i","İ":"i","ş":"s","Ş":"s","ğ":"g","Ğ":"g","ü":"u","Ü":"u","ö":"o","Ö":"o","ç":"c","Ç":"c"})
    s = s.translate(tr_map).lower().strip()
    s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s)
    return s

def _best_match(needle: str, haystack: list[str]) -> str | None:
    norm = _normalize_tr(needle).replace(" ", "")
    cand_norm = [_normalize_tr(c).replace(" ", "") for c in haystack]
    if not cand_norm:
        return None
    matches = difflib.get_close_matches(norm, cand_norm, n=1, cutoff=0.7)
    if not matches:
        return None
    hit = matches[0]
    return haystack[cand_norm.index(hit)]

def transform_yanitlar_to_table(df_raw: pd.DataFrame, faz_value: str | None = None, devamlilik_threshold: int = 4) -> pd.DataFrame:
    """Convert yanitlar.xlsx wide-form to the long table with exact RAPOR_COLUMNS order."""
    src_cols = list(df_raw.columns)

    # Base columns
    col_adsoyad = _best_match("Adınız, soyadınız", src_cols) or _best_match("Ad Soyad", src_cols)
    col_tarih    = _best_match("tarih", src_cols) or _best_match("Zaman damgası", src_cols)
    col_baglanti = _best_match("Bağlantı türü", src_cols)
    col_ver_bip  = _best_match("Bip Uygulama Versiyon", src_cols)
    col_ver_wha  = _best_match("Whatsapp Uygulama Versiyon", src_cols)
    col_ver_tel  = _best_match("Telegram Uygulama Versiyon", src_cols)

    # Feature columns discovery
    features = []
    for base in ["Txt","GM","Call","Media","IM","Voip"]:
        pretty = TEST_ALANI_MAP.get(base.lower(), base)
        for app_prefix, app_value in APPS:
            puan_col  = _best_match(f"{app_prefix} {base} Puan", src_cols)
            yorum_col = _best_match(f"{app_prefix} {base} 3 ve 3'ün altında verilen puan yorumu", src_cols)
            if puan_col:  # yorum_col olmayabilir
                features.append({
                    "test_alani": pretty,
                    "app_value" : app_value,
                    "puan_col"  : puan_col,
                    "yorum_col" : yorum_col
                })

    rows = []
    for _, row in df_raw.iterrows():
        katilimci = row.get(col_adsoyad, None)
        tarih = row.get(col_tarih, None)

        # robust tarih parse
        try:
            if isinstance(tarih, str):
                for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
                    try:
                        tarih = datetime.strptime(tarih.split()[0], fmt).date().isoformat()
                        break
                    except: 
                        pass
            else:
                tarih = pd.to_datetime(tarih).date().isoformat()
        except:
            # leave as-is if parsing fails
            pass

        baglanti = row.get(col_baglanti, None)
        ver_bip  = row.get(col_ver_bip , None)
        ver_wha  = row.get(col_ver_wha , None)
        ver_tel  = row.get(col_ver_tel , None)

        for f in features:
            puan = row.get(f["puan_col"], np.nan)
            if pd.isna(puan):
                continue
            # normalize puan to int
            try:
                puan_int = int(puan)
            except:
                try:
                    puan_int = int(float(str(puan).replace(",", ".")))
                except:
                    continue  # skip if can't coerce

            yorum_val = row.get(f["yorum_col"], None) if f["yorum_col"] in df_raw.columns else None
            vers = ver_bip if f["app_value"]=="bip" else (ver_wha if f["app_value"]=="whatsapp" else ver_tel)
            devam = "OK" if (pd.notna(puan_int) and puan_int >= devamlilik_threshold) else "NOK"

            rows.append({
                'Faz': faz_value,
                'Column1': None,
                'Katılımcı': katilimci,
                'Devamlılık': devam,
                'Tarih': tarih,
                'Test Alanı': f["test_alani"],
                'Cihaz OS': None,
                'Uygulama': f["app_value"],
                'wifi/lte': baglanti,
                'Versiyon': vers,
                'Puan': puan_int,
                'Bip Yorum': yorum_val if f["app_value"]=="bip" else None,
                'Whatsapp yorum': yorum_val if f["app_value"]=="whatsapp" else None,
                'Telegram Yorum': yorum_val if f["app_value"]=="telegram" else None,
                'cihaz': None
            })

    df_out = pd.DataFrame(rows)
    # Ensure exact columns & order
    for c in RAPOR_COLUMNS:
        if c not in df_out.columns:
            df_out[c] = None
    df_out = df_out[RAPOR_COLUMNS].sort_values(["Katılımcı", "Tarih", "Uygulama", "Test Alanı"], kind="stable").reset_index(drop=True)
    return df_out
