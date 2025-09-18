# -*- coding: utf-8 -*-
from __future__ import annotations
import pandas as pd, numpy as np, re, unicodedata, difflib

TARGET_COLS = ['Faz','Test Alanı','OS','Uygulama','Network','Versiyon','Puan','Yorum','Cihaz']

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

# ---- helpers -------------------------------------------------------------
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

def _infer_os_from_device(dev: str | None) -> str | None:
    if not dev: return None
    s = _norm(dev)
    # iOS ipuçları
    if any(k in s for k in ["iphone", "ipad", "ios", "apple", "iphone14", "iphone13", "iphone12", "se (", "pro max"]):
        return "ios"
    # Android marka ipuçları
    if any(k in s for k in [
        "samsung","xiaomi","redmi","mi ","poco","huawei","honor","oppo","realme","oneplus",
        "vivo","tecno","infinix","lenovo","nokia","gm ","general mobile","casper","reeder","tcl","asus","pixel"
    ]):
        return "android"
    # doğrudan "android" geçenler
    if "android" in s: return "android"
    return None

def _expand_candidate(cols: list[str], *cands: str) -> str | None:
    # çoklu adaylardan ilk uyuşanı döndür (fuzzy)
    for c in cands:
        hit = _best(c, cols)
        if hit: return hit
    return None

# ---- main ----------------------------------------------------------------
def transform(df_raw: pd.DataFrame, faz_value: str | None = "Faz 6", devamlilik_threshold: int = 4) -> pd.DataFrame:
    """
    Girdi: yanitlar.xlsx içindeki bir sayfa (puan/yorum sütunları olan)
    Çıktı: ['Faz','Test Alanı','OS','Uygulama','Network','Versiyon','Puan','Yorum','Cihaz']
    """
    cols = list(df_raw.columns)

    # OS / Network / Cihaz için geniş aday listeleri
    col_os = _expand_candidate(cols,
        "Cihaz OS","OS","Telefon OS","İşletim sistemi","Operating System","Platform","Mobil OS"
    )
    col_net = _expand_candidate(cols,
        "Bağlantı türü","Bağlantı","Network","wifi/lte","Ağ türü","Şebeke","Bağlantı Tipi"
    )
    col_dev = _expand_candidate(cols,
        "Cihaz","cihaz","Cihaz modeli","Telefon modeli","Model","Device","Telefon","Model Adı"
    )

    ver_cols = {
        "Bip": _expand_candidate(cols, "Bip Uygulama Versiyon","Bip Versiyon","Bip version","BiP Version"),
        "Whatsapp": _expand_candidate(cols, "Whatsapp Uygulama Versiyon","Whatsapp Versiyon","WhatsApp version"),
        "Telegram": _expand_candidate(cols, "Telegram Uygulama Versiyon","Telegram Versiyon","Telegram version"),
    }

    out_rows = []

    for col in cols:
        if not col.endswith(" Puan"):
            continue
        m = re.match(r'^(Bip|Whatsapp|Telegram)\s+(.+)\s+Puan$', col)
        if not m:
            continue
        app_pretty, scen = m.groups()
        app_val = app_pretty.lower()

        # yanındaki "yorum" kolonu (çoğunlukla)
        comment_col = None
        i = cols.index(col)
        if i + 1 < len(cols) and "yorum" in _norm(cols[i+1]):
            comment_col = cols[i+1]

        # Test Alanı metni
        scen_code = re.sub(r'[^a-z0-9]+','', _norm(scen))
        test_alani = TEST_ALANI_MAP.get(scen_code, scen)

        for ridx, val in df_raw[col].items():
            if pd.isna(val) or str(val) == "":
                continue
            # Puan'ı int'e zorla
            try:
                puan = int(val)
            except Exception:
                try:
                    puan = int(float(str(val).replace(",", ".")))
                except Exception:
                    continue

            # OS
            os_val = None
            if col_os and col_os in df_raw.columns and pd.notna(df_raw.at[ridx, col_os]):
                os_val = str(df_raw.at[ridx, col_os]).strip().lower()
            # Cihaz
            dev_val = None
            if col_dev and col_dev in df_raw.columns and pd.notna(df_raw.at[ridx, col_dev]):
                dev_val = str(df_raw.at[ridx, col_dev]).strip()
            # OS boşsa cihazdan tahmin et
            if not os_val:
                os_guess = _infer_os_from_device(dev_val)
                if os_guess: os_val = os_guess

            # Network
            net_val = None
            if col_net and col_net in df_raw.columns and pd.notna(df_raw.at[ridx, col_net]):
                net_val = str(df_raw.at[ridx, col_net]).strip()

            # Versiyon
            ver_col = ver_cols.get(app_pretty)
            ver_val = None
            if ver_col and ver_col in df_raw.columns and pd.notna(df_raw.at[ridx, ver_col]):
                ver_val = str(df_raw.at[ridx, ver_col]).strip()

            # Yorum
            yorum = None
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

    # Son güvenlik: OS boş ama cihaz iPhone/iPad ise iOS; belirgin Android markası varsa Android
    def _finalize_os(row):
        if row.get("OS"): return row["OS"]
        return _infer_os_from_device(row.get("Cihaz"))
    if not out.empty:
        out["OS"] = out.apply(_finalize_os, axis=1)

    # Kolon sırası/varlığı garantisi
    for c in TARGET_COLS:
        if c not in out.columns:
            out[c] = None
    return out[TARGET_COLS].reset_index(drop=True)
