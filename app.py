import streamlit as st
import pandas as pd
import datetime
import json
import re
import io
from fpdf import FPDF
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. GOOGLE SHEETS BAGLANTISI VE YARDIMCILAR
# ==========================================
st.set_page_config(page_title="Konya Lisesi Bilgi Toplama Sistemi", layout="wide")
conn_gs = st.connection("gsheets", type=GSheetsConnection)

def get_data(worksheet_name):
    """Google Sheet uzerindeki ilgili sekmeden veriyi canlı çeker."""
    try:
        df = conn_gs.read(worksheet=worksheet_name, ttl=0)
        return df if df is not None else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def save_data(worksheet_name, df):
    """Google Sheet üzerindeki ilgili sekmeyi günceller."""
    conn_gs.update(worksheet=worksheet_name, data=df)

def sort_sinif_sube_key(item):
    m = re.search(r'(\d+)', str(item))
    num = int(m.group(1)) if m else 999
    text = re.sub(r'\d+', '', str(item)).strip()
    return (num, text)

def format_phone(phone_str):
    digits = re.sub(r'\D', '', str(phone_str))
    if digits.startswith('0'):
        digits = digits[1:]
    if len(digits) == 10:
        return f"{digits[:3]} {digits[3:6]} {digits[6:8]} {digits[8:10]}"
    return phone_str

def mask_name(name):
    words = str(name).strip().split()
    return " ".join([w[:2] + "*" * (len(w) - 2) if len(w) > 2 else w[0] + "*" for w in words])

def tr_fix(text):
    mapping = {'İ': 'I', 'ı': 'i', 'Ş': 'S', 'ş': 's', 'Ğ': 'G', 'ğ': 'g', 'Ü': 'U', 'ü': 'u', 'Ö': 'O', 'ö': 'o', 'Ç': 'C', 'ç': 'c'}
    for tr, en in mapping.items(): 
        text = str(text).replace(tr, en)
    return text

# ==========================================
# 2. VARSAYILAN TABLO YAPISI KONTROLU
# ==========================================
def init_sheets():
    df_q = get_data("sorular")
    if df_q.empty or "soru_metni" not in df_q.columns:
        varsayilan_sorular = [
            {"id": 1, "soru_metni": "VELİSİ KİM?", "soru_tipi": "coktan_secmeli", "secenekler": "ANNE,BABA,DİĞER", "sira": 1, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 2, "soru_metni": "VELİNİN TC KİMLİK NUMARASI", "soru_tipi": "tc_no", "secenekler": "", "sira": 2, "bagli_parent_id": 1, "bagli_parent_deger": "DİĞER"},
            {"id": 3, "soru_metni": "VELİNİN ADI SOYADI", "soru_tipi": "metin", "secenekler": "", "sira": 3, "bagli_parent_id": 1, "bagli_parent_deger": "DİĞER"},
            {"id": 4, "soru_metni": "VELİNİN TELEFON NUMARASI", "soru_tipi": "telefon", "secenekler": "", "sira": 4, "bagli_parent_id": 1, "bagli_parent_deger": "DİĞER"},
            {"id": 5, "soru_metni": "YAKINLIK DERECESİ", "soru_tipi": "coktan_secmeli", "secenekler": "DAYI,AMCA,TEYZE,HALA,DEDE,NİNE,DİĞER", "sira": 5, "bagli_parent_id": 1, "bagli_parent_deger": "DİĞER"},
            {"id": 6, "soru_metni": "ÖĞRENCİ ÖZEL DURUMU", "soru_tipi": "coklu_secim", "secenekler": "AİLE BÖLÜNMÜŞ (ANNE-BABA AYRI),YETİM (BABA VEFAT),ÖKSÜZ (ANNE VEFAT),ŞEHİT/GAZİ YAKINI,ENGEL DURUMU VAR,YOK", "sira": 6, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 7, "soru_metni": "SÜREKLİ HASTALIĞI / KRONİK RAHATSIZLIK", "soru_tipi": "coklu_secim", "secenekler": "DİYABET,ASTIM,ALERJİ,KALP HASTALIĞI,TANSİYON,EPRİLEPSİ,YOK", "sira": 7, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 8, "soru_metni": "GEÇİRDİĞİ AMELİYAT VEYA KAZA", "soru_tipi": "metin", "secenekler": "", "sira": 8, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 9, "soru_metni": "KİMİNLE OTURUYOR?", "soru_tipi": "coktan_secmeli", "secenekler": "AİLE,ANNE,BABA,AKRABA,YURT,YALNIZ", "sira": 9, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 10, "soru_metni": "OTURDUĞU EV", "soru_tipi": "coktan_secmeli", "secenekler": "KENDİ EVİMİZ,KİRA,LOJMAN", "sira": 10, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 11, "soru_metni": "KENDİ ODASI VAR MI?", "soru_tipi": "coktan_secmeli", "secenekler": "VAR,YOK", "sira": 11, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 12, "soru_metni": "EV İSINMA TİPİ", "soru_tipi": "coktan_secmeli", "secenekler": "KALORİFER,DOĞALGAZ,SOBA,KLİMA", "sira": 12, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 13, "soru_metni": "EVDE İNTERNET VAR MI?", "soru_tipi": "coktan_secmeli", "secenekler": "VAR,YOK", "sira": 13, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 14, "soru_metni": "KENDİSİNE AİT BİLGİSAYAR/TABLET VAR MI?", "soru_tipi": "coktan_secmeli", "secenekler": "VAR,YOK", "sira": 14, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 15, "soru_metni": "BABA MESLEĞİ", "soru_tipi": "metin", "secenekler": "", "sira": 15, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 16, "soru_metni": "ANNE MESLEĞİ", "soru_tipi": "metin", "secenekler": "", "sira": 16, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 17, "soru_metni": "AİLE AYLIK ORTALAMA GELİRİ", "soru_tipi": "coktan_secmeli", "secenekler": "ASGARİ ÜCRET ALTI,ASGARİ ÜCRET,ASGARİ ÜCRET ÜSTÜ", "sira": 17, "bagli_parent_id": 0, "bagli_parent_deger": ""},
            {"id": 18, "soru_metni": "OKULA ULAŞIM ŞEKLİ", "soru_tipi": "coktan_secmeli", "secenekler": "YÜRÜYEREK,SERVİS,TOPLU TAŞIMA,ÖZEL ARAÇ", "sira": 18, "bagli_parent_id": 0, "bagli_parent_deger": ""}
        ]
        save_data("sorular", pd.DataFrame(varsayilan_sorular))

init_sheets()

# ==========================================
# 3. e-OKUL EXCEL PARSER & PDF
# ==========================================
def parse_and_save_eokul(file_buffer):
    def clean(val):
        return str(val).replace('\xa0', ' ').strip() if pd.notna(val) else ""

    xl = pd.ExcelFile(file_buffer)
    all_students = []
    
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)
        c_sinif, c_sube, c_ogretmen = None, None, None
        num_col, ad_col, soyad_col = None, None, None
        
        for idx, row in df.iterrows():
            row_cells = [clean(val) for val in row.values]
            row_text = " ".join([c for c in row_cells if c])
            
            match_sinif = re.search(r'(\d+)\s*\.\s*Sınıf', row_text, re.IGNORECASE)
            match_sube = re.search(r'/\s*([A-ZÇĞİÖŞÜa-zçğıöşü0-9]+)\s*Şubesi', row_text, re.IGNORECASE)
            if match_sinif: c_sinif = match_sinif.group(1).strip()
            if match_sube: c_sube = match_sube.group(1).strip().upper()
                
            if "Sınıf Öğretmeni:" in row_text:
                parts = row_text.split("Sınıf Öğretmeni:")
                if len(parts) > 1:
                    raw_t = parts[1].split("Sınıf")[0].split("Müdür")[0].strip()
                    if raw_t: c_ogretmen = raw_t
            
            for c_idx, cell in enumerate(row_cells):
                cell_lower = cell.lower()
                if "öğrenci no" in cell_lower: num_col = c_idx
                elif cell_lower in ["adı", "ad"]: ad_col = c_idx
                elif cell_lower in ["soyadı", "soyad"]: soyad_col = c_idx
            
            use_num = num_col if num_col is not None else 1
            use_ad = ad_col if ad_col is not None else 2
            use_soyad = soyad_col if soyad_col is not None else (7 if len(row_cells) > 7 else 3)
            
            if len(row_cells) > use_num and len(row_cells) > use_ad:
                val_num, val_ad = row_cells[use_num], row_cells[use_ad]
                val_soyad = row_cells[use_soyad] if len(row_cells) > use_soyad else ""
                
                numara_str = None
                try:
                    num_float = float(val_num)
                    if num_float > 0 and num_float.is_integer():
                        numara_str = str(int(num_float))
                except ValueError: pass
                
                if numara_str and val_ad and val_ad.lower() not in ["adı", "ad", "öğrenci no"]:
                    all_students.append({
                        "numara": str(numara_str),
                        "sinif": c_sinif if c_sinif else "Tanımsız",
                        "sube": c_sube if c_sube else "Tanımsız",
                        "ogretmen": c_ogretmen if c_ogretmen else "Tanımlanmadı",
                        "ad_soyad": f"{val_ad} {val_soyad}".strip()
                    })
    
    if all_students:
        df_new = pd.DataFrame(all_students)
        save_data("ogrenciler", df_new)
    return len(all_students)

def generate_class_pdf(df_sube_merged, questions_df, sinif_sube_adi):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    for idx, st_row in df_sube_merged.iterrows():
        pdf.add_page()
        pdf.set_font("Helvetica", 'B', 13)
        pdf.cell(0, 7, tr_fix("KONYA LISESI OGRENCI BILGI FORMU"), ln=True, align="C")
        pdf.set_font("Helvetica", 'I', 9)
        pdf.cell(0, 5, tr_fix(f"Sinif / Sube: {sinif_sube_adi}  |  Sinif Ogretmeni: {st_row['ogretmen']}"), ln=True, align="C")
        pdf.ln(4)
        
        pdf.set_font("Helvetica", 'B', 10)
        pdf.set_fill_color(220, 230, 242)
        pdf.cell(0, 7, tr_fix(f" OGR. NO: {st_row['numara']}  |  ADI SOYADI: {st_row['ad_soyad']}  |  DURUM: {st_row['FORM DURUMU']}"), ln=True, fill=True)
        pdf.ln(3)
        
        ans_json = json.loads(st_row['yanitlar_json']) if pd.notna(st_row['yanitlar_json']) and st_row['yanitlar_json'] else {}

        if st_row['FORM DURUMU'] == "DOLDURDU":
            for _, q in questions_df.iterrows():
                q_id = str(q['id'])
                parent_id = str(q.get('bagli_parent_id', 0))
                parent_target = str(q.get('bagli_parent_deger', '')).strip()
                
                if parent_id != "0":
                    parent_ans = str(ans_json.get(parent_id, "")).strip()
                    if parent_ans != parent_target:
                        continue
                
                q_title = tr_fix(q['soru_metni'])
                q_ans = tr_fix(ans_json.get(q_id, "-"))
                pdf.set_font("Helvetica", 'B', 8)
                pdf.cell(75, 5, f"{q_title}:", border=0)
                pdf.set_font("Helvetica", size=8)
                pdf.cell(0, 5, f" {q_ans}", border=0, ln=True)
            
            pdf.ln(10)
            pdf.set_font("Helvetica", 'B', 8)
            pdf.cell(95, 5, tr_fix("Sinif Rehber Ogretmeni Imza:"), align="L")
            pdf.cell(95, 5, tr_fix("Veli Imza:"), align="R", ln=True)
        else:
            pdf.set_font("Helvetica", 'I', 10)
            pdf.cell(0, 10, tr_fix("Bu ogrenci henüz formu doldurmamistir."), ln=True)
            
    return pdf.output()

# ==========================================
# 4. STREAMLIT ARAYÜZÜ
# ==========================================
st.title("🏫 Öğrenci Bilgi Formu & Raporlama Sistemi")
tab1, tab2 = st.tabs(["📝 Öğrenci Formu", "⚙️ Yönetici & Öğretmen Paneli"])

# --- TAB 1: ÖĞRENCİ FORMU ---
with tab1:
    df_students = get_data("ogrenciler")
    df_questions = get_data("sorular")
    
    if not df_questions.empty and "sira" in df_questions.columns:
        df_questions = df_questions.sort_values(by=["sira", "id"])

    if df_students.empty:
        st.info("Sistemde henüz öğrenci listesi tanımlı değil.")
    elif df_questions.empty:
        st.warning("Formda henüz soru tanımlanmamış.")
    else:
        df_students["numara"] = df_students["numara"].astype(str)
        siniflar = sorted(df_students["sinif"].unique(), key=sort_sinif_sube_key)
        secilen_sinif = st.selectbox("Sınıfınızı Seçin:", ["SEÇİNİZ"] + siniflar)
        
        if secilen_sinif != "SEÇİNİZ":
            subeler = sorted(df_students[df_students["sinif"] == secilen_sinif]["sube"].unique())
            secilen_sube = st.selectbox("Şubenizi Seçin:", ["SEÇİNİZ"] + subeler)
            
            if secilen_sube != "SEÇİNİZ":
                filtered = df_students[(df_students["sinif"] == secilen_sinif) & (df_students["sube"] == secilen_sube)].copy()
                filtered["display"] = filtered.apply(lambda r: f"{r['numara']} - {mask_name(r['ad_soyad'])}", axis=1)
                
                secilen_ogrenci = st.selectbox("Numaranızı ve Adınızı Seçin:", ["SEÇİNİZ"] + list(filtered["display"]))
                
                if secilen_ogrenci != "SEÇİNİZ":
                    secilen_no = str(secilen_ogrenci.split(" - ")[0])
                    
                    df_yanitlar = get_data("yanitlar")
                    mevcut_yanit = pd.DataFrame()
                    if not df_yanitlar.empty and "numara" in df_yanitlar.columns:
                        df_yanitlar["numara"] = df_yanitlar["numara"].astype(str)
                        mevcut_yanit = df_yanitlar[df_yanitlar["numara"] == secilen_no]
                    
                    can_submit, is_update = True, False
                    eski_cevaplar = {}
                    if not mevcut_yanit.empty:
                        eski_cevaplar = json.loads(mevcut_yanit.iloc[0]["yanitlar_json"])
                        st.warning(f"⚠️ **{secilen_no}** numaralı öğrenci olarak daha önce form doldurulmuş.")
                        if st.checkbox("Yanıtlarımı güncellemek istiyorum."): is_update = True
                        else: can_submit = False
                    
                    if can_submit:
                        st.divider()
                        st.subheader("Form Soruları")
                        
                        if "answers" not in st.session_state or st.session_state.get("current_no") != secilen_no:
                            st.session_state["answers"] = eski_cevaplar.copy()
                            st.session_state["current_no"] = secilen_no
                            
                        validation_errors = []
                        
                        for _, q in df_questions.iterrows():
                            q_id = str(q["id"])
                            q_metni = q['soru_metni']
                            q_type = q["soru_tipi"]
                            parent_id = str(q.get('bagli_parent_id', 0))
                            parent_target = str(q.get('bagli_parent_deger', '')).strip()
                            
                            if parent_id != "0":
                                parent_ans = str(st.session_state["answers"].get(parent_id, "")).strip()
                                if parent_ans != parent_target:
                                    st.session_state["answers"][q_id] = ""
                                    continue
                                    
                            default_val = st.session_state["answers"].get(q_id, "")
                            
                            if q_type == "coktan_secmeli":
                                opts = ["SEÇİNİZ"] + [opt.strip() for opt in str(q["secenekler"]).split(",") if opt.strip()]
                                idx = opts.index(default_val) if default_val in opts else 0
                                selected = st.selectbox(f"📌 {q_metni}", opts, index=idx, key=f"widget_{q_id}")
                                st.session_state["answers"][q_id] = selected
                                
                            elif q_type == "coklu_secim":
                                opts = [opt.strip() for opt in str(q["secenekler"]).split(",") if opt.strip()]
                                def_list = [x.strip() for x in default_val.split(",") if x.strip()] if default_val else []
                                sel_list = st.multiselect(f"📌 {q_metni}", opts, default=def_list, key=f"widget_{q_id}")
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
                            for _, q in df_questions.iterrows():
                                q_id = str(q["id"])
                                q_metni = q['soru_metni']
                                q_val = str(st.session_state["answers"].get(q_id, "")).strip()
                                q_type = q["soru_tipi"]
                                parent_id = str(q.get('bagli_parent_id', 0))
                                parent_target = str(q.get('bagli_parent_deger', '')).strip()
                                
                                if parent_id != "0":
                                    parent_ans = str(st.session_state["answers"].get(parent_id, "")).strip()
                                    if parent_ans != parent_target:
                                        continue
                                
                                if q_type == "coktan_secmeli" and (q_val == "SEÇİNİZ" or not q_val):
                                    validation_errors.append(f"❌ **{q_metni}** sorusu için seçim yapınız.")
                                elif q_type == "tc_no" and q_val and not (q_val.isdigit() and len(q_val) == 11):
                                    validation_errors.append(f"❌ **{q_metni}** 11 haneli rakam olmalıdır.")
                                elif q_type == "telefon" and q_val:
                                    clean_p = re.sub(r'\D', '', q_val)
                                    if clean_p.startswith('0'): clean_p = clean_p[1:]
                                    if len(clean_p) != 10:
                                        validation_errors.append(f"❌ **{q_metni}** 10 haneli olmalıdır.")
                                    else:
                                        st.session_state["answers"][q_id] = format_phone(clean_p)
                            
                            if validation_errors:
                                for err in validation_errors: st.error(err)
                            else:
                                tarih = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                json_data = json.dumps(st.session_state["answers"], ensure_ascii=False)
                                
                                if df_yanitlar.empty:
                                    df_yanitlar = pd.DataFrame(columns=["numara", "yanitlar_json", "tarih"])
                                
                                if is_update:
                                    df_yanitlar.loc[df_yanitlar["numara"] == secilen_no, ["yanitlar_json", "tarih"]] = [json_data, tarih]
                                else:
                                    new_row = pd.DataFrame([{"numara": secilen_no, "yanitlar_json": json_data, "tarih": tarih}])
                                    df_yanitlar = pd.concat([df_yanitlar, new_row], ignore_index=True)
                                
                                save_data("yanitlar", df_yanitlar)
                                st.success("✅ Form yanıtlarınız Google Sheets üzerine kaydedildi!")

# --- TAB 2: YÖNETİCİ & ÖĞRETMEN PANATELİ ---
with tab2:
    st.subheader("Yönetici & Öğretmen Paneli")
    sifre = st.text_input("Yönetici Şifresi:", type="password")
    
    if sifre == "admin123":
        df_o = get_data("ogrenciler")
        df_y = get_data("yanitlar")
        df_q = get_data("sorular")
        
        if not df_o.empty: df_o["numara"] = df_o["numara"].astype(str)
        if not df_y.empty: df_y["numara"] = df_y["numara"].astype(str)
        if not df_q.empty and "sira" in df_q.columns: df_q = df_q.sort_values(by=["sira", "id"])
        
        merged_all = pd.merge(df_o, df_y, on='numara', how='left') if not df_o.empty else pd.DataFrame()
        if not merged_all.empty:
            merged_all['FORM DURUMU'] = merged_all['yanitlar_json'].apply(lambda x: "DOLDURDU" if pd.notna(x) and str(x) != "" else "DOLDURMADI")
            merged_all['sinif_sube'] = merged_all['sinif'] + merged_all['sube']
        
        sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs([
            "📊 İstatistikler", "📗 Excel Raporu", "📄 PDF Dökümleri", "🛠️ Soruları Yönet"
        ])
        
        with sub_tab1:
            st.markdown("### 📈 Doldurma Durum Takibi")
            if not merged_all.empty:
                toplam_ogr = len(df_o)
                dolduran_ogr = len(merged_all[merged_all['FORM DURUMU'] == 'DOLDURDU'])
                st.metric("Toplam Öğrenci", toplam_ogr)
                st.metric("Form Dolduran", dolduran_ogr)

        with sub_tab2:
            st.markdown("### 📗 Toplu Excel İndirme")
            if not merged_all.empty:
                excel_rows = []
                for idx, r in merged_all.iterrows():
                    ans_dict = json.loads(r['yanitlar_json']) if pd.notna(r['yanitlar_json']) and r['yanitlar_json'] else {}
                    row_data = {"SINIFI": r.get('sinif_sube', ''), "OKUL NO": r['numara'], "ADI SOYADI": r['ad_soyad']}
                    for _, q in df_q.iterrows():
                        row_data[q['soru_metni']] = ans_dict.get(str(q['id']), "")
                    excel_rows.append(row_data)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pd.DataFrame(excel_rows).to_excel(writer, index=False)
                st.download_button("📊 EXCEL İNDİR", data=output.getvalue(), file_name="RAPOR.xlsx")

        with sub_tab3:
            st.markdown("### 📄 Sınıf PDF Dökümleri")
            if not merged_all.empty:
                tum_subeler = sorted(merged_all['sinif_sube'].unique().tolist(), key=sort_sinif_sube_key)
                secilen_pdf_sube = st.selectbox("Sınıf Seçin:", tum_subeler)
                if st.button("PDF Oluştur"):
                    sube_students = merged_all[merged_all['sinif_sube'] == secilen_pdf_sube]
                    pdf_bytes = generate_class_pdf(sube_students, df_q, secilen_pdf_sube)
                    st.download_button(f"📥 {secilen_pdf_sube}.pdf İndir", data=bytes(pdf_bytes), file_name=f"{secilen_pdf_sube}.pdf")

        with sub_tab4:
            st.markdown("#### e-Okul Excel Listesi Yükle")
            uploaded_file = st.file_uploader("Excel Yükle", type=["xls", "xlsx"])
            if uploaded_file and st.button("Veritabanına İşle"):
                toplam = parse_and_save_eokul(uploaded_file)
                st.success(f"✅ {toplam} öğrenci aktarıldı!")