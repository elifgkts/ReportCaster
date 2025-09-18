# -*- coding: utf-8 -*-
from __future__ import annotations
import pandas as pd, numpy as np, re, unicodedata, difflib
from datetime import datetime

# ÇIKTI ŞEMASI (rapor.xlsx → "Fonksiyonlar Data")
RAPOR_COLUMNS = [
    'Faz','Column1','Katılımcı','Devamlılık','Tarih','Test Alanı',
    'Cihaz OS','Uygulama','wifi/lte','Versiyon','Puan',
    'Bip Yorum','Whatsapp yorum','Telegram Yorum','cihaz'
]

# Test alanı adlandırmaları (ihtiyaca göre genişletebilirsin)
TEST_ALANI_MAP = {
    "txt":   "IM - 1-1 txt mesaj",
    "gm":    "IM - Grup mesajlaşması",
    "im":    "IM - Genel",
    "media": "IM - Medya paylaşımı",
    "voip":  "Voip - 1-1 Görüntülü görüşme",
    "gsg":   "Voip - Grup sesli görüşme",
    "ggg":   "Voip - Grup görüntülü görüşme",
    "call":  "Voip - 1-1 Sesli görüşme",
}

# Uygulama etiketleri
APPS = [("Bip","bip"), ("Whatsapp","whatsapp"), ("Telegram","telegram")]

# ----------------- yardımcılar -----------------
def _normalize_tr(s: str) -> str:
    if s is None: return ""
    s = str(s)
    tr = str.maketrans({"ı":"i","İ":"i","ş":"s","Ş":"s","ğ":"g","Ğ":"g","ü":"u","Ü":"u","ö":"o","Ö":"o","ç":"c","Ç":"c"})
    s = s.translate(tr).lower().strip()
    s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)

def _best_match(needle: str, haystack: list[str]) -> str | None:
    norm = _normalize_tr(needle).replace(" ", "")
    cand = [_normalize_tr(c).replace(" ", "") for c in haystack]
    if not cand: return None
    m = difflib.get_close_matches(norm, cand, n=1, cutoff=0.7)
    return haystack[cand.index(m[0])] if m else None

def _excel_serial(dt) -> int | None:
    """Excel seri tarih (A1 tarzı) -> gün sayısı. Hata olursa None."""
    if pd.isna(dt): return None
    try:
        ts = pd.to_datetime(dt)
        return int((ts - pd.Timestamp("1899-12-30")).days)
    except Exception:
        return None

# ----------------- asıl dönüşüm -----------------
def transform_yanitlar_to_table(df_raw: pd.DataFrame,
                                faz_value: str | None = "Faz 6",
                                devamlilik_threshold: int = 4) -> pd.DataFrame:
    """
    yanitlar.xlsx’teki **tek bir form sayfasını** (Puan kolonları olan) alır ve
    rapor.xlsx’teki 'Fonksiyonlar Data' şemasına çevirir.
    """
    cols = list(df_raw.columns)

    # kolon keşfi (fuzzy)
    col_name = _best_match("Adınız, soyadınız", cols) or _best_match("Ad Soyad", cols) or _best_match("Katılımcı", cols)
    col_date = _best_match("Tarih", cols) or _best_match("Zaman damgası", cols)
    col_conn = _best_match("Bağlantı türü", cols) or _best_match("Bağlantı", cols)
    ver_cols = {
        "Bip": _best_match("Bip Uygulama Versiyon", cols) or _best_match("Bip Versiyon", cols),
        "Whatsapp": _best_match("Whatsapp Uygulama Versiyon", cols) or _best_match("Whatsapp Versiyon", cols),
        "Telegram": _best_match("Telegram Uygulama Versiyon", cols) or _best_match("Telegram Versiyon", cols),
    }

    rows = []
    id_map, next_id = {}, 1

    # Tüm "… Puan" kolonlarını dolaş
    for col in cols:
        if not col.endswith(" Puan"):
            continue
        m = re.match(r'^(Bip|Whatsapp|Telegram)\s+(.+)\s+Puan$', col)
        if not m:
            continue
        app, scen = m.groups()

        # senaryo → Test Alanı
        scen_code = _normalize_tr(scen)
        scen_code = re.sub(r'[^a-z0-9]+','',scen_code)
        test_alani = TEST_ALANI_MAP.get(scen_code, scen)

        # yanındaki yorum kolonu
        comment_col = None
        idx = cols.index(col)
        if idx+1 < len(cols) and 'yorum' in _normalize_tr(cols[idx+1]):
            comment_col = cols[idx+1]

        # satır satır işle
        for i, val in df_raw[col].items():
            if pd.isna(val) or str(val) == "":
                continue
            # puanı int'e zorla
            try:
                puan = int(val)
            except Exception:
                try:
                    puan = int(float(str(val).replace(",", ".")))
                except Exception:
                    continue

            # zorunlu alanlar
            name = str(df_raw.at[i, col_name]) if col_name in df_raw.columns else ""
            if not name:
                continue
            if name not in id_map:
                id_map[name] = str(next_id); next_id += 1

            tarih = _excel_serial(df_raw.at[i, col_date]) if col_date in df_raw.columns else None
            conn  = str(df_raw.at[i, col_conn]) if col_conn in df_raw.columns else None
            vers  = ""
            if ver_cols[app] in df_raw.columns and pd.notna(df_raw.at[i, ver_cols[app]]):
                vers = str(df_raw.at[i, ver_cols[app]])

            devam = "OK" if puan >= devamlilik_threshold else "NOK"

            yorum = None
            if comment_col and comment_col in df_raw.columns:
                cv = df_raw.at[i, comment_col]
                if pd.notna(cv) and str(cv) != "":
                    yorum = str(cv)

            # uygulama özel yorum sütunları
            bip_y, wa_y, tg_y = None, None, None
            if app == "Bip": bip_y = yorum
            elif app == "Whatsapp": wa_y = yorum
            elif app == "Telegram": tg_y = yorum

            rows.append({
                'Faz': faz_value,
                'Column1': id_map[name],
                'Katılımcı': name,
                'Devamlılık': devam,
                'Tarih': tarih,
                'Test Alanı': test_alani,
                'Cihaz OS': None,
                'Uygulama': app.lower(),
                'wifi/lte': conn,
                'Versiyon': vers,
                'Puan': puan,
                'Bip Yorum': bip_y,
                'Whatsapp yorum': wa_y,
                'Telegram Yorum': tg_y,
                'cihaz': None
            })

    df = pd.DataFrame(rows)
    # tam kolon sırası garantisi
    for c in RAPOR_COLUMNS:
        if c not in df.columns:
            df[c] = None
    df = df[RAPOR_COLUMNS].sort_values(
        ["Katılımcı","Tarih","Uygulama","Test Alanı"], kind="stable"
    ).reset_index(drop=True)
    return df

# ---- reportcaster/streamlit_app.py uyumluluğu için ----
def transform(df_raw: pd.DataFrame, faz_value: str | None = None, devamlilik_threshold: int = 4) -> pd.DataFrame:
    """
    reportcaster/streamlit_app.py şu imzayı çağırıyor:
        transform(df_raw, faz_value=..., devamlilik_threshold=...)
    """
    return transform_yanitlar_to_table(
        df_raw, faz_value=faz_value, devamlilik_threshold=devamlilik_threshold
    )
