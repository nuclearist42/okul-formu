# ==========================================================================
# KONYA LİSESİ ÖĞRENCİ BİLGİ TOPLAMA SİSTEMİ
# ==========================================================================
# BU SÜRÜMDE YAPILAN KRİTİK DÜZELTMELER (veri kaybı / "sheet siliniyor" sorunu için):
#
# 1) ESKİ KOD: Tek bir öğrenci cevabı eklemek için TÜM "yanitlar" sayfası
#    okunuyor, pandas'ta yeni satır ekleniyor, sonra sayfanın TAMAMI
#    conn.update() ile yeniden yazılıyordu. İki öğrenci neredeyse aynı anda
#    form gönderdiğinde, ikincisi birincinin eklediği satırı görmeden eski
#    (bayat) tabloyu yeniden yazıyor ve birincinin cevabını sessizce siliyordu.
#    YENİ KOD: gspread ile SADECE ilgili tek satır ekleniyor (append_row) ya da
#    güncelleniyor (targeted range update). Başka hiçbir satıra dokunulmuyor.
#
# 2) ESKİ KOD: threading.Lock() sadece yazma anını kilitliyordu; okuma
#    (mevcut veriyi çekme) kilidin DIŞINDA yapılıyordu. Bu da "kontrol et,
#    sonra yaz" (TOCTOU) yarış durumuna açık kapı bırakıyordu.
#    YENİ KOD: okuma + satır bulma + yazma işleminin TAMAMI tek bir kilit
#    bloğu içinde, tek bir atomik işlem olarak yapılıyor.
#
# 3) ESKİ KOD: Google Sheets'e ağ/oturum hatası olduğunda fonksiyon None
#    dönüyor ve bu None, st.cache_data tarafından 5 dakika boyunca "geçerli
#    sonuç" gibi ÖNBELLEĞE ALINIYORDU.
#    YENİ KOD: Hatalar önbelleğe hiç yazılmıyor; sadece başarılı okumalar
#    önbelleklenir, hata durumunda bir sonraki çağrıda tekrar denenir.
#
# 4) ESKİ KOD: Her tek öğrenci gönderiminde "istatistik" ve "doldurmayanlar"
#    sayfaları da tam olarak yeniden hesaplanıp yazılıyordu → tek gönderim
#    başına 6-7 Google Sheets API çağrısı. Sınıfın tamamı aynı anda form
#    doldurunca Google'ın dakikalık kotasına takılma ve yarım kalan
#    yazma riski artıyordu.
#    YENİ KOD: İstatistik senkronizasyonu en fazla dakikada bir otomatik
#    çalışır (throttle); admin panelindeki "Yenile" butonu istenildiğinde
#    anında (force=True) çalıştırabilir.
#
# 5) ESKİ KOD: e-Okul Excel yükleme, öğrenci listesinin TAMAMINI hiçbir
#    onay istemeden (allow_delete=True) üzerine yazıyordu. Yanlış/eksik
#    dosya yüklenirse tüm öğrenci listesi tek seferde küçülebiliyordu.
#    YENİ KOD: Yeni liste, mevcut listeden belirgin şekilde küçükse admin
#    açıkça onaylamadan işlem durur.
#
# NOT: secrets.toml dosyanızı DEĞİŞTİRMENİZE gerek yok. Aynı
# [connections.gsheets] bloğunu (service account bilgileri + "spreadsheet"
# alanı) kullanmaya devam ediyoruz; sadece streamlit-gsheets sarmalayıcısı
# yerine gspread'i doğrudan, hedefli (tek satır) işlemler için kullanıyoruz.
#
# requirements.txt için: gspread ve google-auth paketlerinin kurulu olması
# yeterli (genelde streamlit-gsheets'in bağımlılığı olarak zaten kuruludur).
# ==========================================================================

import streamlit as st
import pandas as pd
import datetime
import re
import io
import time
import threading

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

# TÜM ÖĞRENCİLERİ SIRAYA SOKACAK KÜRESEL TURNİKE (MUTEX LOCK)
sheet_lock = threading.Lock()

st.set_page_config(page_title="Konya Lisesi Bilgi Toplama Sistemi", layout="wide")

# ==========================================
# 0. GOOGLE SHEETS BAĞLANTISI (gspread doğrudan)
# ==========================================
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


@st.cache_resource(show_spinner=False)
def get_gspread_client():
    creds_info = dict(st.secrets["connections"]["gsheets"])
    creds_info.pop("spreadsheet", None)
    creds_info.pop("spreadsheet_id", None)
    creds = Credentials.from_service_account_info(creds_info, scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_gspread_client()
    conf = st.secrets["connections"]["gsheets"]
    sheet_ref = conf.get("spreadsheet") or conf.get("spreadsheet_id")
    if not sheet_ref:
        raise RuntimeError(
            "secrets.toml içindeki [connections.gsheets] bloğunda 'spreadsheet' "
            "alanı bulunamadı."
        )
    if str(sheet_ref).startswith("http"):
        return client.open_by_url(sheet_ref)
    return client.open_by_key(sheet_ref)


def get_ws(worksheet_name):
    """Güncel worksheet handle'ı döner (bu sadece metadata çağrısıdır, ucuzdur)."""
    return get_spreadsheet().worksheet(worksheet_name)


def api_call_with_retry(func, *args, retries=3, base_delay=1.2, **kwargs):
    """Geçici Google API hatalarında (429/500/503 vb.) küçük bekleme ile tekrar dener."""
    last_err = None
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            last_err = e
            time.sleep(base_delay * (attempt + 1))
        except Exception as e:
            last_err = e
            break
    raise last_err


def col_letter(n):
    """n. sütunun harfini döner (1 -> A, 27 -> AA ...)."""
    a1 = rowcol_to_a1(1, n)
    return re.sub(r"\d+$", "", a1)


# ==========================================
# 1. TEMİZLEME / NORMALİZASYON YARDIMCILARI (değişmedi)
# ==========================================
def clean_val(val, default=""):
    if pd.isna(val) or val is None:
        return default
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", "<na>", ""]:
        return default
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str


def clean_id(val):
    if pd.isna(val) or val is None:
        return "0"
    s = str(val).strip()
    if s.lower() in ["nan", "none", "<na>", "", "0"]:
        return "0"
    s = s.replace(".", ",")
    parts = []
    for p in s.split(","):
        p_str = clean_val(p)
        if p_str and p_str != "0":
            parts.append(p_str)
    return ",".join(parts) if parts else "0"


def tr_norm(text):
    if pd.isna(text) or text is None:
        return ""
    t = str(text).strip()
    mapping = {
        "i": "I", "İ": "I", "ı": "I", "I": "I",
        "ş": "S", "Ş": "S",
        "ğ": "G", "Ğ": "G",
        "ü": "U", "Ü": "U",
        "ö": "O", "Ö": "O",
        "ç": "C", "Ç": "C",
    }
    res = [mapping.get(ch, ch.upper()) for ch in t]
    norm_str = "".join(res)
    return re.sub(r"[^A-Z0-9]", "", norm_str)


def get_ans_for_id(pid, answers_map):
    pid_str = clean_val(pid)
    widget_key = f"widget_{pid_str}"
    if widget_key in st.session_state:
        val = str(st.session_state[widget_key]).strip()
        if val and val != "SEÇİNİZ":
            return val
    if answers_map and pid_str in answers_map:
        val = str(answers_map.get(pid_str, "")).strip()
        if val and val != "SEÇİNİZ":
            return val
    return ""


def is_question_visible(q_row, answers_map):
    parent_id_str = clean_id(q_row.get("bagli_parent_id"))
    parent_target_str = clean_val(q_row.get("bagli_parent_deger"), default="")

    if parent_id_str in ["0", "", "nan", "none"]:
        return True

    p_ids = [clean_val(x) for x in str(parent_id_str).split(",") if clean_val(x) and clean_val(x) != "0"]
    p_targets = [clean_val(x) for x in str(parent_target_str).split(",") if clean_val(x)]

    if not p_ids:
        return True

    for idx, pid in enumerate(p_ids):
        target = p_targets[idx] if idx < len(p_targets) else (p_targets[0] if p_targets else "")
        target_norm = tr_norm(target)

        user_ans = get_ans_for_id(pid, answers_map)
        user_ans_norm = tr_norm(user_ans)

        if not user_ans_norm:
            return False

        if target_norm:
            match = (
                target_norm == user_ans_norm
                or target_norm in user_ans_norm
                or user_ans_norm in target_norm
                or (target_norm in ["SAG", "HAYATTA"] and user_ans_norm in ["SAG", "HAYATTA"])
            )
            if not match:
                return False

    return True


def sort_sinif_sube_key(item):
    m = re.search(r"(\d+)", str(item))
    num = int(m.group(1)) if m else 999
    text = re.sub(r"\d+", "", str(item)).strip()
    return (num, text)


def format_phone(phone_str):
    digits = re.sub(r"\D", "", str(phone_str))
    if digits.startswith("0"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"{digits[:3]} {digits[3:6]} {digits[6:8]} {digits[8:10]}"
    return phone_str


def mask_name(name):
    words = str(name).strip().split()
    return " ".join([w[:2] + "*" * (len(w) - 2) if len(w) > 2 else w[0] + "*" for w in words])


# ==========================================
# 2. DÜŞÜK SEVİYE OKUMA/YAZMA (gspread tabanlı, atomik)
# ==========================================
def ws_values_to_df(values):
    if not values:
        return pd.DataFrame()
    header = values[0]
    seen = {}
    clean_header = []
    for h in header:
        h2 = (h or "").strip()
        if not h2:
            h2 = f"col_{len(clean_header)}"
        if h2 in seen:
            seen[h2] += 1
            h2 = f"{h2}_{seen[h2]}"
        else:
            seen[h2] = 0
        clean_header.append(h2)

    rows = values[1:]
    max_len = len(clean_header)
    fixed_rows = [r + [""] * (max_len - len(r)) if len(r) < max_len else r[:max_len] for r in rows]
    df = pd.DataFrame(fixed_rows, columns=clean_header)
    for col in df.columns:
        df[col] = df[col].apply(clean_val)
    return df


def read_sheet_live(worksheet_name):
    """Her zaman Google Sheets'ten anlık veri okur. Hata olursa None döner
    ve bu None ASLA önbelleğe yazılmaz (önceki sürümdeki hata buydu)."""
    try:
        ws = get_ws(worksheet_name)
        values = api_call_with_retry(ws.get_all_values)
        return ws_values_to_df(values)
    except Exception as e:
        st.session_state["_last_gsheet_error"] = str(e)
        return None


def get_data_live(worksheet_name):
    """Bağlantı koparsa None döner (güvenliği tetiklemek için)."""
    df = read_sheet_live(worksheet_name)
    if df is None:
        st.error(f"⚠️ Google Sheets bağlantı hatası: {st.session_state.get('_last_gsheet_error', 'bilinmeyen hata')}")
    return df


# Basit, elle yönetilen TTL önbelleği (sadece BAŞARILI okumalar saklanır)
_CACHE_STORE = {}
_CACHE_TTL_SECONDS = 20


def get_data(worksheet_name):
    now = time.time()
    cached = _CACHE_STORE.get(worksheet_name)
    if cached and (now - cached["ts"] < _CACHE_TTL_SECONDS):
        return cached["df"].copy()

    df = read_sheet_live(worksheet_name)
    ss_key = f"backup_df_{worksheet_name}"

    if df is not None:
        _CACHE_STORE[worksheet_name] = {"df": df, "ts": now}
        if not df.empty:
            st.session_state[ss_key] = df
        return df.copy()

    if ss_key in st.session_state:
        return st.session_state[ss_key].copy()
    return pd.DataFrame()


def clear_all_caches():
    _CACHE_STORE.clear()
    for key in list(st.session_state.keys()):
        if key.startswith("backup_df_"):
            del st.session_state[key]


def ensure_headers(ws, needed_columns):
    """Başlık satırına eksik sütunları ekler; MEVCUT VERİ SATIRLARINA DOKUNMAZ."""
    header = api_call_with_retry(ws.row_values, 1)
    if not header:
        header = list(needed_columns)
        api_call_with_retry(ws.update, "1:1", [header], value_input_option="RAW")
        return header
    missing = [c for c in needed_columns if c not in header]
    if missing:
        new_header = header + missing
        api_call_with_retry(ws.update, "1:1", [new_header], value_input_option="RAW")
        return new_header
    return header


def find_row_by_key(ws, header, key_col, key_val):
    if key_col not in header:
        return None
    key_idx = header.index(key_col)
    values = api_call_with_retry(ws.get_all_values)
    key_val_str = str(key_val)
    for i, r in enumerate(values[1:], start=2):
        cell = r[key_idx] if len(r) > key_idx else ""
        if clean_val(cell) == key_val_str:
            return i
    return None


def upsert_row(worksheet_name, key_col, key_val, row_dict):
    """`key_col`==`key_val` olan satırı günceller; yoksa yeni satır ekler.
    Okuma + arama + yazma TEK bir kilit bloğunda, ATOMİK olarak yapılır.
    Sadece TEK satır etkilenir; tablonun geri kalanına ASLA dokunulmaz."""
    with sheet_lock:
        ws = get_ws(worksheet_name)
        header = ensure_headers(ws, list(row_dict.keys()) + [key_col])
        row_idx = find_row_by_key(ws, header, key_col, key_val)
        row_values = [str(row_dict.get(h, "")) for h in header]

        if row_idx:
            rng = f"A{row_idx}:{col_letter(len(header))}{row_idx}"
            api_call_with_retry(ws.update, rng, [row_values], value_input_option="RAW")
            return True, "updated"
        else:
            api_call_with_retry(ws.append_row, row_values, value_input_option="RAW")
            return True, "inserted"


def delete_row_by_key(worksheet_name, key_col, key_val):
    with sheet_lock:
        ws = get_ws(worksheet_name)
        header = api_call_with_retry(ws.row_values, 1)
        row_idx = find_row_by_key(ws, header, key_col, key_val)
        if row_idx:
            api_call_with_retry(ws.delete_rows, row_idx)
            return True
        return False


def replace_full_sheet(worksheet_name, df, allow_shrink=False):
    """Bilinçli TOPLU değişiklikler için (örn. e-Okul listesi yükleme).
    Öğrenci/soru/yanıt kaybını önlemek için: yeni veri, mevcut veriden
    belirgin şekilde küçükse `allow_shrink=True` verilmediği sürece durur."""
    with sheet_lock:
        ws = get_ws(worksheet_name)
        current_values = api_call_with_retry(ws.get_all_values)
        current_rows = max(len(current_values) - 1, 0)
        new_rows = len(df)

        if current_rows > 0 and new_rows < current_rows * 0.9 and not allow_shrink:
            return False, current_rows, new_rows

        values = [list(df.columns)] + df.astype(str).values.tolist() if not df.empty else [list(df.columns)]
        api_call_with_retry(ws.clear)
        api_call_with_retry(ws.update, "A1", values, value_input_option="RAW")
        return True, current_rows, new_rows


def write_report_sheet(worksheet_name, df):
    """istatistik / doldurmayanlar gibi TÜRETİLMİŞ (kaynak olmayan) rapor
    sayfaları için: her seferinde sıfırdan yeniden hesaplandığından tam
    üzerine yazma güvenlidir (kalıcı veri kaybı riski taşımaz)."""
    try:
        with sheet_lock:
            ws = get_ws(worksheet_name)
            values = [list(df.columns)] + df.astype(str).values.tolist() if not df.empty else [list(df.columns)]
            api_call_with_retry(ws.clear)
            api_call_with_retry(ws.update, "A1", values, value_input_option="RAW")
    except Exception:
        pass


# ==========================================
# 3. İSTATİSTİK SENKRONİZASYONU (throttled — dakikada en fazla 1 kez otomatik)
# ==========================================
_LAST_STATS_SYNC = {"ts": 0.0}
STATS_SYNC_MIN_INTERVAL = 60.0


def sync_stats_to_gsheet(force=False):
    now = time.time()
    if not force and (now - _LAST_STATS_SYNC["ts"] < STATS_SYNC_MIN_INTERVAL):
        return
    _LAST_STATS_SYNC["ts"] = now
    try:
        df_students = get_data_live("ogrenciler")
        df_yanitlar = get_data_live("yanitlar")

        if df_students is None or df_yanitlar is None:
            return
        if df_students.empty:
            return

        if not df_yanitlar.empty:
            cols_to_use = ["numara"] + [c for c in df_yanitlar.columns if c not in df_students.columns]
            merged_auto = pd.merge(df_students, df_yanitlar[cols_to_use], on="numara", how="left")
        else:
            merged_auto = df_students.copy()

        merged_auto["FORM DURUMU"] = merged_auto.get("tarih", pd.Series([None] * len(merged_auto))).apply(
            lambda x: "DOLDURDU" if pd.notna(x) and str(x).strip() not in ["", "nan"] else "DOLDURMADI"
        )
        merged_auto["sinif_sube"] = merged_auto["sinif"].astype(str) + "/" + merged_auto["sube"].astype(str)

        ogretmen_map = merged_auto.groupby("sinif_sube")["ogretmen"].first()
        stats_df = merged_auto.groupby("sinif_sube")["FORM DURUMU"].value_counts().unstack(fill_value=0)
        if "DOLDURDU" not in stats_df.columns:
            stats_df["DOLDURDU"] = 0
        if "DOLDURMADI" not in stats_df.columns:
            stats_df["DOLDURMADI"] = 0
        stats_df["TOPLAM"] = stats_df["DOLDURDU"] + stats_df["DOLDURMADI"]
        stats_df["SINIF ÖĞRETMENİ"] = stats_df.index.map(ogretmen_map)

        stats_to_export = stats_df[["TOPLAM", "DOLDURDU", "DOLDURMADI", "SINIF ÖĞRETMENİ"]].reset_index()
        write_report_sheet("istatistik", stats_to_export)

        doldurmayanlar_df = merged_auto[merged_auto["FORM DURUMU"] == "DOLDURMADI"][
            ["sinif_sube", "numara", "ad_soyad", "ogretmen"]
        ]
        write_report_sheet("doldurmayanlar", doldurmayanlar_df)
    except Exception:
        pass


# ==========================================
# 4. e-OKUL EXCEL PARSER (sadece PARSE eder, kaydetmez)
# ==========================================
def parse_eokul_file(file_buffer):
    def clean(val):
        return str(val).replace("\xa0", " ").strip() if pd.notna(val) else ""

    xl = pd.ExcelFile(file_buffer)
    all_students = []

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)
        c_sinif, c_sube, c_ogretmen = None, None, None
        num_col, ad_col, soyad_col = None, None, None

        for idx, row in df.iterrows():
            row_cells = [clean(val) for val in row.values]
            row_text = " ".join([c for c in row_cells if c])

            match_sinif = re.search(r"(\d+)\s*\.\s*Sınıf", row_text, re.IGNORECASE)
            match_sube = re.search(r"/\s*([A-ZÇĞİÖŞÜa-zçğıöşü0-9]+)\s*Şubesi", row_text, re.IGNORECASE)
            if match_sinif:
                c_sinif = match_sinif.group(1).strip()
            if match_sube:
                c_sube = match_sube.group(1).strip().upper()

            if "Sınıf Öğretmeni:" in row_text:
                parts = row_text.split("Sınıf Öğretmeni:")
                if len(parts) > 1:
                    raw_t = parts[1].split("Sınıf")[0].split("Müdür")[0].strip()
                    if raw_t:
                        c_ogretmen = raw_t

            for c_idx, cell in enumerate(row_cells):
                cell_lower = cell.lower()
                if "öğrenci no" in cell_lower:
                    num_col = c_idx
                elif cell_lower in ["adı", "ad"]:
                    ad_col = c_idx
                elif cell_lower in ["soyadı", "soyad"]:
                    soyad_col = c_idx

            use_num = num_col if num_col is not None else 1
            use_ad = ad_col if ad_col is not None else 2
            use_soyad = soyad_col if soyad_col is not None else (7 if len(row_cells) > 7 else 3)

            if len(row_cells) > use_num and len(row_cells) > use_ad:
                val_num, val_ad = row_cells[use_num], row_cells[use_ad]
                val_soyad = row_cells[use_soyad] if len(row_cells) > use_soyad else ""

                numara_str = clean_val(val_num)
                if numara_str.isdigit() and val_ad and val_ad.lower() not in ["adı", "ad", "öğrenci no"]:
                    all_students.append(
                        {
                            "numara": numara_str,
                            "sinif": clean_val(c_sinif) if c_sinif else "Tanımsız",
                            "sube": clean_val(c_sube) if c_sube else "Tanımlanmadı",
                            "ogretmen": c_ogretmen if c_ogretmen else "Tanımlanmadı",
                            "ad_soyad": f"{val_ad} {val_soyad}".strip(),
                        }
                    )

    df_new = pd.DataFrame(all_students) if all_students else pd.DataFrame(
        columns=["numara", "sinif", "sube", "ogretmen", "ad_soyad"]
    )
    return df_new, len(all_students)


# ==========================================
# 5. STREAMLIT ARAYÜZÜ
# ==========================================
try:
    get_spreadsheet()
except Exception as e:
    st.error(f"❌ Google Sheets bağlantısı kurulamadı: {e}")
    st.stop()

st.title("Konya Lisesi Öğrenci Bilgi Formu")

st.markdown(
    """
ORTAÖĞRETİM KURUMLARINA YERLEŞTİRME İŞLEMİNDE ÖĞRENCİLERİN E-OKUL BİLGİLERİ KULLANILMAKTADIR. BU NEDENLE ÖĞRENCİLERE AİT BİLGİLERİN TAM VE GÜNCEL OLMASI GEREKMEKTEDİR. AİLE VE ÖĞRENCİ İLE İLGİLİ TÜM GELİŞME VE DEĞİŞİKLİKLERİN ZAMANINDA OKUL İDARESİNE BİLDİRİLMESİ GEREKMEKTEDİR. ELEKTRONİK FORMUN İLERLEYİŞİ VERİLEN CEVAPLARA GÖRE OLACAĞI İÇİN BİLGİLERİN DOĞRU VE GÜNCEL OLMASI ÖNEM ARZ ETMEKTEDİR.

BİLGİ GİRİŞLERİ TAMAMLANINCA OKUL REHBERLİK SERVİSİ TARAFINDAN ÇIKTI ALINARAK İMZALAMASI İÇİN VELİLERE ULAŞTIRILACAKTIR.

FORM BİLGİLERİ YALNIZCA OKUL İDARESİ VE REHBERLİK SERVİSİ TARAFINDAN GÖRÜNTÜLENEBİLECEKTİR. ÜÇÜNCÜ KİŞİLERLE PAYLAŞILMAMAKTADIR.
"""
)

tab1, tab2 = st.tabs(["📝 Öğrenci Formu", "⚙️ Panel"])

# --- TAB 1: ÖĞRENCİ FORMU ---
with tab1:
    df_students = get_data("ogrenciler")
    df_questions = get_data("sorular")

    if not df_questions.empty and "sira" in df_questions.columns:
        df_questions["sira"] = pd.to_numeric(df_questions["sira"], errors="coerce").fillna(999)
        df_questions = df_questions.sort_values(by=["sira", "id"])

    if df_students.empty or "numara" not in df_students.columns:
        st.info("Sistemde henüz öğrenci listesi tanımlı değil. Lütfen Yönetim Paneli'nden e-Okul listesi yükleyin.")
    elif df_questions.empty or "id" not in df_questions.columns:
        st.warning("Yönetim panelinden henüz soru eklenmemiş. Lütfen soruları oluşturun.")
    else:
        siniflar = sorted(df_students["sinif"].astype(str).unique(), key=sort_sinif_sube_key)
        secilen_sinif = st.selectbox("Sınıfınızı Seçin:", ["SEÇİNİZ"] + siniflar)

        if secilen_sinif != "SEÇİNİZ":
            subeler = sorted(df_students[df_students["sinif"].astype(str) == secilen_sinif]["sube"].astype(str).unique())
            secilen_sube = st.selectbox("Şubenizi Seçin:", ["SEÇİNİZ"] + subeler)

            if secilen_sube != "SEÇİNİZ":
                filtered = df_students[
                    (df_students["sinif"].astype(str) == secilen_sinif) & (df_students["sube"].astype(str) == secilen_sube)
                ].copy()
                filtered["display"] = filtered.apply(lambda r: f"{r['numara']} - {mask_name(r['ad_soyad'])}", axis=1)

                secilen_ogrenci = st.selectbox("Numaranızı ve Adınızı Seçin:", ["SEÇİNİZ"] + list(filtered["display"]))

                if secilen_ogrenci != "SEÇİNİZ":
                    secilen_no = clean_val(secilen_ogrenci.split(" - ")[0])
                    student_row = filtered[filtered["numara"].astype(str) == secilen_no].iloc[0]

                    df_yanitlar_check = get_data_live("yanitlar")

                    mevcut_yanit = pd.DataFrame()
                    if df_yanitlar_check is not None and not df_yanitlar_check.empty and "numara" in df_yanitlar_check.columns:
                        mevcut_yanit = df_yanitlar_check[df_yanitlar_check["numara"].astype(str) == secilen_no]

                    can_submit, is_update = True, False

                    if not mevcut_yanit.empty:
                        st.warning(f"⚠️ **{secilen_no}** numaralı öğrenci olarak daha önce form doldurulmuştur.")
                        if st.checkbox("Yanıtlarımı güncellemek istiyorum."):
                            is_update = True
                            st.info(
                                "🔒 **Gizlilik ve Güvenlik Bildirimi:** Kişisel verilerinizin gizliliği gereği daha önce "
                                "girmiş olduğunuz bilgiler ekranda gösterilmemektedir. Bilgilerinizi güncellemek için "
                                "lütfen aşağıdaki alanları sıfırdan doldurunuz. Göndereceğiniz yeni veriler eski "
                                "kaydınızın üzerine yazılacaktır."
                            )
                        else:
                            can_submit = False

                    if can_submit:
                        st.divider()
                        st.subheader("Form Soruları")

                        if "answers" not in st.session_state or st.session_state.get("current_no") != secilen_no:
                            st.session_state["answers"] = {}
                            st.session_state["current_no"] = secilen_no

                        validation_errors = []

                        for _, q in df_questions.iterrows():
                            q_id = clean_val(q["id"])
                            q_metni = q["soru_metni"]
                            q_type = q["soru_tipi"]

                            if not is_question_visible(q, st.session_state["answers"]):
                                st.session_state["answers"][q_id] = ""
                                continue

                            default_val = get_ans_for_id(q_id, st.session_state["answers"])
                            raw_sec = clean_val(q.get("secenekler", ""), default="")

                            if q_type == "coktan_secmeli":
                                opts = ["SEÇİNİZ"] + [opt.strip() for opt in raw_sec.split(",") if opt.strip()]
                                idx = opts.index(default_val) if default_val in opts else 0
                                selected = st.selectbox(f"📌 {q_metni}", opts, index=idx, key=f"widget_{q_id}")
                                st.session_state["answers"][q_id] = selected

                            elif q_type == "coklu_secim":
                                opts = [opt.strip() for opt in raw_sec.split(",") if opt.strip()]
                                def_list = [x.strip() for x in default_val.split(",") if x.strip()] if default_val else []
                                valid_def_list = [x for x in def_list if x in opts]
                                sel_list = st.multiselect(f"📌 {q_metni}", opts, default=valid_def_list, key=f"widget_{q_id}")
                                st.session_state["answers"][q_id] = ", ".join(sel_list)

                            elif q_type == "tc_no":
                                val = st.text_input(f"📌 {q_metni}", value=default_val, max_chars=11, key=f"widget_{q_id}")
                                st.session_state["answers"][q_id] = val

                            elif q_type == "telefon":
                                val = st.text_input(f"📌 {q_metni}", value=default_val, max_chars=14, key=f"widget_{q_id}")
                                st.session_state["answers"][q_id] = val

                            else:
                                val = st.text_input(f"📌 {q_metni}", value=default_val, key=f"widget_{q_id}")
                                st.session_state["answers"][q_id] = val

                        st.write("")
                        if st.button("💾 Formu Gönder / Kaydet", type="primary"):
                            tarih = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                            new_row_data = {
                                "numara": secilen_no,
                                "sinif": clean_val(student_row["sinif"]),
                                "sube": clean_val(student_row["sube"]),
                                "ogretmen": clean_val(student_row["ogretmen"]),
                                "ad_soyad": clean_val(student_row["ad_soyad"]),
                            }

                            for _, q in df_questions.iterrows():
                                q_id = clean_val(q["id"])
                                q_metni = q["soru_metni"]
                                q_val = get_ans_for_id(q_id, st.session_state["answers"])
                                q_type = q["soru_tipi"]

                                if not is_question_visible(q, st.session_state["answers"]):
                                    new_row_data[q_metni] = ""
                                    continue

                                if q_type == "coktan_secmeli" and (q_val == "SEÇİNİZ" or not q_val):
                                    validation_errors.append(f"❌ **{q_metni}** sorusu için seçim yapınız.")
                                elif q_type == "tc_no" and q_val and not (q_val.isdigit() and len(q_val) == 11):
                                    validation_errors.append(f"❌ **{q_metni}** 11 haneli rakam olmalıdır.")
                                elif q_type == "telefon" and q_val:
                                    clean_p = re.sub(r"\D", "", q_val)
                                    if clean_p.startswith("0"):
                                        clean_p = clean_p[1:]
                                    if len(clean_p) != 10:
                                        validation_errors.append(f"❌ **{q_metni}** 10 haneli olmalıdır.")
                                    else:
                                        q_val = format_phone(clean_p)

                                new_row_data[q_metni] = q_val

                            new_row_data["tarih"] = tarih

                            if validation_errors:
                                for err in validation_errors:
                                    st.error(err)
                            else:
                                # --- KRİTİK DÜZELTME ---
                                # Tüm tabloyu okuyup yeniden yazmak yerine, SADECE bu
                                # öğrencinin satırını atomik biçimde ekliyor/güncelliyoruz.
                                # Başka bir öğrencinin cevabına dokunulmaz.
                                try:
                                    ok, action = upsert_row("yanitlar", "numara", secilen_no, new_row_data)
                                    if ok:
                                        clear_all_caches()
                                        sync_stats_to_gsheet(force=False)
                                        st.success("✅ Form yanıtlarınız başarıyla kaydedildi!")
                                except Exception as e:
                                    st.error(
                                        "❌ Google Sheets'e yazarken bir hata oluştu. Verileriniz kaybolmadı, "
                                        f"lütfen 'Formu Gönder' butonuna tekrar basın. (Hata: {e})"
                                    )

# --- TAB 2: YÖNETİCİ & ÖĞRETMEN PANELİ ---
with tab1 if False else tab2:
    st.subheader("Panel")
    sifre = st.text_input("Yönetici Şifresi:", type="password")

    admin_pass = st.secrets.get("ADMIN_PASSWORD", "bettiyin")

    if sifre == admin_pass:
        if st.button("🔄 Google Sheets Verilerini Yenile / Önbelleği Temizle"):
            clear_all_caches()
            sync_stats_to_gsheet(force=True)
            clear_all_caches()
            st.success("✅ Önbellek temizlendi ve Google Sheets sayfaları (istatistik/doldurmayanlar) güncellendi!")
            st.rerun()

        df_o = get_data("ogrenciler")
        df_y = get_data("yanitlar")
        df_q = get_data("sorular")

        if df_o.empty or "numara" not in df_o.columns:
            df_o = pd.DataFrame(columns=["numara", "sinif", "sube", "ogretmen", "ad_soyad"])

        if df_y.empty or "numara" not in df_y.columns:
            df_y = pd.DataFrame(columns=["numara", "sinif", "sube", "ogretmen", "ad_soyad"])

        if df_q.empty or "id" not in df_q.columns:
            df_q = pd.DataFrame(columns=["id", "soru_metni", "soru_tipi", "secenekler", "sira", "bagli_parent_id", "bagli_parent_deger"])
        else:
            if "sira" in df_q.columns:
                df_q["sira"] = pd.to_numeric(df_q["sira"], errors="coerce").fillna(999)
                df_q = df_q.sort_values(by=["sira", "id"])

        if not df_o.empty:
            if not df_y.empty:
                cols_to_use = ["numara"] + [c for c in df_y.columns if c not in df_o.columns]
                merged_all = pd.merge(df_o, df_y[cols_to_use], on="numara", how="left")
            else:
                merged_all = df_o.copy()

            merged_all["FORM DURUMU"] = merged_all.get("tarih", pd.Series([None] * len(merged_all))).apply(
                lambda x: "DOLDURDU" if pd.notna(x) and str(x).strip() not in ["", "nan"] else "DOLDURMADI"
            )
            merged_all["sinif_sube"] = merged_all["sinif"].astype(str) + "/" + merged_all["sube"].astype(str)
        else:
            merged_all = pd.DataFrame()

        sub_tab1, sub_tab2, sub_tab3 = st.tabs(["📊 İstatistikler", "📗 Excel Raporu", "🛠️ Soru & e-Okul Yönetimi"])

        with sub_tab1:
            st.markdown("### 📈 Genel ve Sınıf Bazlı Durum Takibi")
            if not merged_all.empty:
                toplam_ogr = len(df_o)
                dolduran_ogr = len(merged_all[merged_all["FORM DURUMU"] == "DOLDURDU"])
                doldurmayan_ogr = toplam_ogr - dolduran_ogr
                orani = int((dolduran_ogr / toplam_ogr) * 100) if toplam_ogr > 0 else 0

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Toplam Öğrenci", toplam_ogr)
                col2.metric("Form Dolduran", dolduran_ogr)
                col3.metric("Doldurmayan", doldurmayan_ogr)
                col4.metric("Tamamlanma Oranı", f"%{orani}")

                st.divider()
                st.markdown("#### Sınıf/Şube Bazında Doldurma Durumları")
                ogretmen_map = merged_all.groupby("sinif_sube")["ogretmen"].first()
                stats_df = merged_all.groupby("sinif_sube")["FORM DURUMU"].value_counts().unstack(fill_value=0)
                if "DOLDURDU" not in stats_df.columns:
                    stats_df["DOLDURDU"] = 0
                if "DOLDURMADI" not in stats_df.columns:
                    stats_df["DOLDURMADI"] = 0
                stats_df["TOPLAM"] = stats_df["DOLDURDU"] + stats_df["DOLDURMADI"]
                stats_df["SINIF ÖĞRETMENİ"] = stats_df.index.map(ogretmen_map)

                st.dataframe(stats_df[["TOPLAM", "DOLDURDU", "DOLDURMADI", "SINIF ÖĞRETMENİ"]], use_container_width=True)

                with st.expander("🚨 Formu Henüz Doldurmayan Öğrenciler Listesi"):
                    doldurmayanlar = merged_all[merged_all["FORM DURUMU"] == "DOLDURMADI"][
                        ["sinif_sube", "numara", "ad_soyad", "ogretmen"]
                    ]
                    st.dataframe(doldurmayanlar, use_container_width=True)
            else:
                st.info("Sistemde henüz öğrenci verisi yok. Lütfen 'Soru & e-Okul Yönetimi' sekmesinden e-Okul Excel dosyasını yükleyin.")

        with sub_tab2:
            st.markdown("### 📗 Toplu Excel İndirme")
            if not merged_all.empty and not df_q.empty:
                excel_rows = []
                for idx, r in merged_all.iterrows():
                    row_data = {
                        "SINIF/ŞUBE": r.get("sinif_sube", ""),
                        "OKUL NO": r["numara"],
                        "ADI SOYADI": r["ad_soyad"],
                        "DURUM": r["FORM DURUMU"],
                    }
                    for _, q in df_q.iterrows():
                        q_metni = q["soru_metni"]
                        val = str(r.get(q_metni, "")).strip()
                        if val.lower() in ["nan", "none"]:
                            val = ""
                        row_data[q_metni] = val

                    excel_rows.append(row_data)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    pd.DataFrame(excel_rows).to_excel(writer, index=False)
                st.download_button("📊 TÜM VERİLERİ EXCEL OLARAK İNDİR", data=output.getvalue(), file_name="OKUL_BILGI_FORMU_RAPOR.xlsx")
            else:
                st.info("Rapor oluşturmak için yeterli veri bulunamadı.")

        with sub_tab3:
            st.markdown("### 🛠️ Sistem Yönetim Paneli")

            with st.expander("📥 e-Okul Excel Listesi Yükle / Güncelle", expanded=False):
                uploaded_file = st.file_uploader(
                    "e-Okul'dan aldığınız Sinif_Listesi.xls/xlsx dosyasını seçin:", type=["xls", "xlsx"]
                )

                if uploaded_file and st.button("Dosyayı Analiz Et"):
                    df_parsed, toplam_parsed = parse_eokul_file(uploaded_file)
                    st.session_state["_pending_ogrenci_df"] = df_parsed
                    st.session_state["_pending_ogrenci_toplam"] = toplam_parsed

                pending_df = st.session_state.get("_pending_ogrenci_df")
                if pending_df is not None:
                    mevcut_df_check = get_data_live("ogrenciler")
                    mevcut_sayi = len(mevcut_df_check) if mevcut_df_check is not None else 0
                    yeni_sayi = len(pending_df)

                    st.info(f"📋 Mevcut kayıtlı öğrenci sayısı: **{mevcut_sayi}**  |  Yeni dosyadaki öğrenci sayısı: **{yeni_sayi}**")

                    onay_gerekli = mevcut_sayi > 0 and yeni_sayi < mevcut_sayi * 0.9
                    force_ok = True
                    if onay_gerekli:
                        st.warning(
                            "⚠️ Yeni dosyadaki öğrenci sayısı, mevcut listeden belirgin şekilde daha az! "
                            "Yanlış veya eksik bir dosya seçmiş olabilirsiniz."
                        )
                        force_ok = st.checkbox("Evet, öğrenci sayısının azalmasını onaylıyorum, üzerine yazmak istiyorum.")

                    if st.button("✅ Veritabanına İşle / Kaydet", type="primary", disabled=(onay_gerekli and not force_ok)):
                        ok, eski, yeni = replace_full_sheet("ogrenciler", pending_df, allow_shrink=True)
                        if ok:
                            st.success(f"✅ {yeni_sayi} öğrenci başarıyla Google Sheets veritabanına aktarıldı!")
                            del st.session_state["_pending_ogrenci_df"]
                            clear_all_caches()
                            st.rerun()
                        else:
                            st.error("❌ İşlem güvenlik nedeniyle durduruldu.")

            with st.expander("🗑️ Doldurulmuş Öğrenci Form Yanıtını Sil / Sıfırla", expanded=True):
                df_y_del = get_data_live("yanitlar")
                if df_y_del is not None and not df_y_del.empty and "numara" in df_y_del.columns:
                    df_valid_y = df_y_del[df_y_del["numara"].astype(str).str.strip() != ""].copy()
                    if not df_valid_y.empty:
                        df_valid_y["disp_sil"] = df_valid_y.apply(
                            lambda r: f"{r['numara']} - {r.get('ad_soyad', '')} ({r.get('sinif', '')}/{r.get('sube', '')})", axis=1
                        )
                        silinecek_ogrenci = st.selectbox("Yanıtı Silinecek Öğrenciyi Seçin:", ["SEÇİNİZ"] + list(df_valid_y["disp_sil"]))
                        if silinecek_ogrenci != "SEÇİNİZ":
                            sil_no = clean_val(silinecek_ogrenci.split(" - ")[0])
                            if st.button(f"🗑️ {sil_no} Numaralı Öğrencinin Yanıtlarını Tamamen Sil", type="primary"):
                                # Sadece ilgili TEK satır silinir; başka hiçbir kayda dokunulmaz.
                                if delete_row_by_key("yanitlar", "numara", sil_no):
                                    sync_stats_to_gsheet(force=True)
                                    clear_all_caches()
                                    st.success(f"✅ {sil_no} numaralı öğrencinin yanıtları başarıyla silindi!")
                                    st.rerun()
                                else:
                                    st.error("❌ Kayıt bulunamadı ya da silinemedi.")
                    else:
                        st.info("Sistemde henüz doldurulmuş yanıt bulunmuyor.")
                else:
                    st.info("Sistemde henüz doldurulmuş yanıt bulunmuyor.")

            st.divider()
            st.markdown("### 📝 Form Sorularını Yönet (CRUD)")

            q_id_to_title = {}
            if not df_q.empty and "id" in df_q.columns:
                for _, q_item in df_q.iterrows():
                    c_id = clean_val(q_item["id"])
                    if c_id and c_id != "0":
                        q_id_to_title[c_id] = q_item["soru_metni"]

            if not df_q.empty:
                st.markdown("#### 📋 Mevcut Soru Listesi ve Şartlı Bağlantı Kontrolü")
                df_q_disp = df_q.copy()

                def format_parents_summary(row):
                    p_id_str = clean_id(row.get("bagli_parent_id"))
                    if p_id_str in ["0", "", "nan"]:
                        return "Ana Soru (Her Zaman Görünür)"
                    ids = p_id_str.split(",")
                    titles = [f"[ID:{i}] {q_id_to_title.get(i, 'Bilinmeyen Soru')}" for i in ids]
                    return ", ".join(titles)

                df_q_disp["Bağlı Olduğu Üst Soru(lar)"] = df_q_disp.apply(format_parents_summary, axis=1)
                disp_cols = [
                    c for c in ["id", "sira", "soru_metni", "soru_tipi", "secenekler", "Bağlı Olduğu Üst Soru(lar)", "bagli_parent_deger"]
                    if c in df_q_disp.columns
                ]
                st.dataframe(df_q_disp[disp_cols], use_container_width=True)

            q_islem = st.radio("Yapmak istediğiniz işlem:", ["Yeni Soru Ekle", "Mevcut Soruyu Düzenle / Sil"], horizontal=True)

            parent_opts = {}
            if not df_q.empty and "id" in df_q.columns:
                for _, q_item in df_q.iterrows():
                    c_id = clean_val(q_item["id"])
                    if c_id and c_id != "0":
                        parent_opts[f"ID:{c_id} - {q_item['soru_metni']}"] = c_id

            if q_islem == "Yeni Soru Ekle":
                with st.form("yeni_soru_form"):
                    y_metin = st.text_input("Soru Metni:")
                    y_tip = st.selectbox("Soru Tipi:", ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"])
                    y_secenekler = st.text_input(
                        "Seçenekler (Çoktan seçmeli veya çoklu seçim ise virgülle ayırın):", help="Örn: SAĞ,ÖLÜ veya BİRLİKTE,AYRI"
                    )
                    y_sira = st.number_input("Soru Sırası:", min_value=1, value=len(df_q) + 1 if not df_q.empty else 1)

                    selected_y_parents = st.multiselect(
                        "Bağlı Olduğu Üst Soru(lar) (Şartlı Gösterim):",
                        options=list(parent_opts.keys()),
                        help="Bu sorunun görünmesi için yanıtlanması gereken üst soruları seçin.",
                    )
                    y_parent_val = st.text_input("Şart Değer(leri):", help="Örn: SAĞ (Eğer iki soru seçtiyseniz ikisi için de geçerli olur)")

                    if st.form_submit_button("➕ Soruyu Kaydet"):
                        if y_metin:
                            try:
                                max_id = int(max([int(clean_val(x)) for x in df_q["id"] if clean_val(x).isdigit()]))
                            except Exception:
                                max_id = 0
                            new_id = str(max_id + 1)

                            y_p_ids = ",".join([parent_opts[k] for k in selected_y_parents]) if selected_y_parents else "0"

                            new_q = {
                                "id": new_id,
                                "soru_metni": y_metin,
                                "soru_tipi": y_tip,
                                "secenekler": y_secenekler,
                                "sira": str(y_sira),
                                "bagli_parent_id": clean_id(y_p_ids),
                                "bagli_parent_deger": clean_val(y_parent_val),
                            }
                            # Sadece TEK satır eklenir; diğer sorulara dokunulmaz.
                            upsert_row("sorular", "id", new_id, new_q)
                            st.success("✅ Yeni soru kaydedildi!")
                            clear_all_caches()
                            st.rerun()
                        else:
                            st.error("Lütfen soru metnini boş bırakmayın.")

            elif q_islem == "Mevcut Soruyu Düzenle / Sil":
                if not df_q.empty and "id" in df_q.columns:
                    q_dict = {f"ID:{clean_val(r['id'])} - {r['soru_metni']}": clean_val(r["id"]) for _, r in df_q.iterrows() if clean_val(r["id"])}
                    if q_dict:
                        secilen_q_label = st.selectbox("Düzenlenecek Soruyu Seçin:", list(q_dict.keys()))
                        secilen_q_id = q_dict[secilen_q_label]
                        q_row = df_q[df_q["id"].apply(lambda x: clean_val(x)) == secilen_q_id].iloc[0]

                        cur_p_ids = [clean_val(x) for x in str(clean_id(q_row.get("bagli_parent_id"))).split(",") if clean_val(x) and clean_val(x) != "0"]
                        default_selected_parents = [k for k, v in parent_opts.items() if v in cur_p_ids]

                        with st.form("duzenle_soru_form"):
                            d_metin = st.text_input("Soru Metni:", value=str(q_row["soru_metni"]))
                            d_tip_idx = (
                                ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"].index(q_row["soru_tipi"])
                                if q_row["soru_tipi"] in ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"]
                                else 0
                            )
                            d_tip = st.selectbox("Soru Tipi:", ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"], index=d_tip_idx)
                            d_secenekler = st.text_input("Seçenekler:", value=clean_val(q_row.get("secenekler")))
                            d_sira = st.number_input("Soru Sırası:", value=int(clean_val(q_row.get("sira"), "1") or 1))

                            selected_d_parents = st.multiselect(
                                "Bağlı Olduğu Üst Soru(lar) (Şartlı Gösterim):",
                                options=list(parent_opts.keys()),
                                default=default_selected_parents,
                            )
                            d_parent_val = st.text_input("Şart Değer(leri):", value=clean_val(q_row.get("bagli_parent_deger")), help="Örn: SAĞ")

                            btn_col1, btn_col2 = st.columns(2)
                            guncelle = btn_col1.form_submit_button("💾 Güncelle")
                            sil = btn_col2.form_submit_button("🗑️ Soruyu Sil", type="primary")

                            if guncelle:
                                d_p_ids = ",".join([parent_opts[k] for k in selected_d_parents]) if selected_d_parents else "0"
                                updated_q = {
                                    "id": secilen_q_id,
                                    "soru_metni": str(d_metin),
                                    "soru_tipi": str(d_tip),
                                    "secenekler": str(d_secenekler),
                                    "sira": str(d_sira),
                                    "bagli_parent_id": clean_id(d_p_ids),
                                    "bagli_parent_deger": str(d_parent_val),
                                }
                                # Sadece BU sorunun satırı güncellenir; diğer sorular etkilenmez.
                                upsert_row("sorular", "id", secilen_q_id, updated_q)
                                st.success("✅ Soru güncellendi!")
                                clear_all_caches()
                                st.rerun()

                            if sil:
                                delete_row_by_key("sorular", "id", secilen_q_id)
                                st.success("🗑️ Soru silindi!")
                                clear_all_caches()
                                st.rerun()
                    else:
                        st.info("Düzenlenecek soru bulunamadı.")
