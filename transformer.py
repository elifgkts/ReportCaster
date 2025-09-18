# -*- coding: utf-8 -*-
from __future__ import annotations
import pandas as pd, numpy as np, re, unicodedata, difflib

# Nihai tablo şeması (BİREBİR)
TARGET_COLS = [
    'Faz', 'Test Alanı', 'OS', 'Uygulama', 'Network', 'Versiyon', 'Puan', 'Yorum', 'Cihaz'
]

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
    m = difflib.get_close_matches(norm, cand, n=1, cutoff=0.7)
    return haystack[cand.index(m[0])] if m else None

def transform(df_raw: pd.DataFrame, faz_value: str | None = "Faz 6", devamlilik_threshold: int = 4) -> pd.DataFrame:
    """
    Eski imza ile çağrılıyor. df_raw (yanitlar sekmesi) → tek tablolu çıktı:
    ['Faz','Test Alanı','OS','Uygulama','Network','Versiyon','Puan','Yorum','Cihaz']
    """
    cols = list(df_raw.columns)

    # Lazım olabilecek kolonlar (fuzzy)
    col_os   = _best("Cihaz OS", cols) or _best("OS", cols)
    col_net  = _best("Bağlantı türü", cols) or _best("Network", cols) or _best("wifi/lte", cols)
    col_dev  = _best("cihaz", cols) or _best("Cihaz", cols)
    ver_cols = {
        "Bip": _best("Bip Uygulama Versiyon", cols) or _best("Bip Versiyon", cols),
        "Whatsapp": _best("Whatsapp Uygulama Versiyon", cols) or _best("Whatsapp Versiyon", cols),
        "Telegram": _best("Telegram Uygulama Versiyon", cols) or _best("Telegram Versiyon", cols),
    }

    out_rows = []

    for col in cols:
        # Yalnızca "... Puan" biten sütunları dolaş
        if not col.endswith(" Puan"):
            continue
        m = re.match(r'^(Bip|Whatsapp|Telegram)\s+(.+)\s+Puan$', col)
        if not m:
            continue
        app_pretty, scen = m.groups()
        app_val = app_pretty.lower()

        # Yorum sütunu: genellikle puan kolonunun hemen sağında "... yorum"
        comment_col = None
        i = cols.index(col)
        if i + 1 < len(cols) and "yorum" in _norm(cols[i+1]):
            comment_col = cols[i+1]

        # Test Alanı ismini map’ten bul
        scen_code = _norm(scen)
        scen_code = re.sub(r'[^a-z0-9]+','',scen_code)
        test_alani = TEST_ALANI_MAP.get(scen_code, scen)

        # Satır satır değerleri al
        for ridx, val in df_raw[col].items():
            if pd.isna(val) or str(val) == "":
                continue
            # Puan int’e
            try:
                puan = int(val)
            except Exception:
                try:
                    puan = int(float(str(val).replace(",", ".")))
                except Exception:
                    continue

            # OS / Network / Versiyon / Cihaz / Yorum
            os_val  = str(df_raw.at[ridx, col_os]).strip()  if col_os  in df_raw.columns and pd.notna(df_raw.at[ridx, col_os])  else None
            net_val = str(df_raw.at[ridx, col_net]).strip() if col_net in df_raw.columns and pd.notna(df_raw.at[ridx, col_net]) else None
            dev_val = str(df_raw.at[ridx, col_dev]).strip() if col_dev in df_raw.columns and pd.notna(df_raw.at[ridx, col_dev]) else None
            ver_col = ver_cols.get(app_pretty)
            ver_val = str(df_raw.at[ridx, ver_col]).strip() if ver_col in df_raw.columns and pd.notna(df_raw.at[ridx, ver_col]) else None
            yorum   = None
            if comment_col and comment_col in df_raw.columns:
                cv = df_raw.at[ridx, comment_col]
                if pd.notna(cv) and str(cv) != "":
                    yorum = str(cv)

            out_rows.append({
                'Faz': faz_value,
                'Test Alanı': test_alani,
                'OS': os_val,
                'Uygulama': app_val,
                'Network': net_val,
                'Versiyon': ver_val,
                'Puan': puan,
                'Yorum': yorum,
                'Cihaz': dev_val,
            })

    out = pd.DataFrame(out_rows)
    # Kolonları garanti altına al ve sıraya koy
    for c in TARGET_COLS:
        if c not in out.columns:
            out[c] = None
    out = out[TARGET_COLS].reset_index(drop=True)
    return out
