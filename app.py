import streamlit as st
import pandas as pd
import datetime
import json
import re
import io
from fpdf import FPDF
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. GOOGLE SHEETS BAĞLANTISI VE YARDIMCILAR
# ==========================================
st.set_page_config(page_title="Konya Lisesi Bilgi Toplama Sistemi", layout="wide")
conn_gs = st.connection("gsheets", type=GSheetsConnection)

def clean_val(val, default=""):
    if pd.isna(val) or val is None:
        return default
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", "<na>", ""]:
        return default
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

def get_data(worksheet_name):
    try:
        df = conn_gs.read(worksheet=worksheet_name, ttl=600)
        if df is not None and not df.empty:
            for col in ["numara", "sinif", "sube", "id", "sira"]:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: clean_val(x))
            return df.copy()
        return pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def save_data(worksheet_name, df):
    conn_gs.update(worksheet=worksheet_name, data=df)
    st.cache_data.clear()

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
# 2. e-OKUL EXCEL PARSER & PDF
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
                
                numara_str = clean_val(val_num)
                if numara_str.isdigit() and val_ad and val_ad.lower() not in ["adı", "ad", "öğrenci no"]:
                    all_students.append({
                        "numara": numara_str,
                        "sinif": clean_val(c_sinif) if c_sinif else "Tanımsız",
                        "sube": clean_val(c_sube) if c_sube else "Tanımsız",
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
        
        if st_row['FORM DURUMU'] == "DOLDURDU":
            for _, q in questions_df.iterrows():
                q_title = q['soru_metni']
                parent_id = clean_val(q.get('bagli_parent_id'), default="0")
                parent_target = clean_val(q.get('bagli_parent_deger'), default="")
                
                if parent_id != "0" and parent_id != "":
                    parent_q_rows = questions_df[questions_df['id'].astype(str) == parent_id]
                    if not parent_q_rows.empty:
                        parent_title = parent_q_rows.iloc[0]['soru_metni']
                        parent_ans = str(st_row.get(parent_title, "")).strip()
                        if parent_ans != parent_target:
                            continue
                
                q_ans = tr_fix(str(st_row.get(q_title, "-")).strip())
                if q_ans.lower() in ["nan", "none", ""]: q_ans = "-"
                
                pdf.set_font("Helvetica", 'B', 8)
                pdf.cell(75, 5, f"{tr_fix(q_title)}:", border=0)
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
# 3. STREAMLIT ARAYÜZÜ
# ==========================================
st.title("🏫 Öğrenci Bilgi Formu & Raporlama Sistemi")
tab1, tab2 = st.tabs(["📝 Öğrenci Formu", "⚙️ Yönetici & Öğretmen Paneli"])

# --- TAB 1: ÖĞRENCİ FORMU ---
with tab1:
    df_students = get_data("ogrenciler")
    df_questions = get_data("sorular")
    
    if not df_questions.empty and "sira" in df_questions.columns:
        df_questions["sira"] = pd.to_numeric(df_questions["sira"], errors='coerce').fillna(999)
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
                filtered = df_students[(df_students["sinif"].astype(str) == secilen_sinif) & (df_students["sube"].astype(str) == secilen_sube)].copy()
                filtered["display"] = filtered.apply(lambda r: f"{r['numara']} - {mask_name(r['ad_soyad'])}", axis=1)
                
                secilen_ogrenci = st.selectbox("Numaranızı ve Adınızı Seçin:", ["SEÇİNİZ"] + list(filtered["display"]))
                
                if secilen_ogrenci != "SEÇİNİZ":
                    secilen_no = clean_val(secilen_ogrenci.split(" - ")[0])
                    
                    df_yanitlar = get_data("yanitlar")
                    if df_yanitlar.empty or "numara" not in df_yanitlar.columns:
                        df_yanitlar = pd.DataFrame(columns=["numara", "tarih"])
                    
                    mevcut_yanit = df_yanitlar[df_yanitlar["numara"] == secilen_no] if not df_yanitlar.empty else pd.DataFrame()
                    
                    can_submit, is_update = True, False
                    eski_cevaplar = {}
                    
                    if not mevcut_yanit.empty:
                        row = mevcut_yanit.iloc[0]
                        for _, q in df_questions.iterrows():
                            q_id = clean_val(q["id"])
                            q_metni = q["soru_metni"]
                            if q_metni in row:
                                eski_cevaplar[q_id] = clean_val(row[q_metni])
                        
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
                            q_id = clean_val(q["id"])
                            q_metni = q['soru_metni']
                            q_type = q["soru_tipi"]
                            
                            parent_id = clean_val(q.get('bagli_parent_id'), default="0")
                            parent_target = clean_val(q.get('bagli_parent_deger'), default="")
                            
                            if parent_id != "0" and parent_id != "":
                                parent_ans = str(st.session_state["answers"].get(parent_id, "")).strip()
                                if parent_ans != parent_target:
                                    st.session_state["answers"][q_id] = ""
                                    continue
                                    
                            default_val = st.session_state["answers"].get(q_id, "")
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
                            new_row_data = {"numara": secilen_no, "tarih": tarih}
                            
                            for _, q in df_questions.iterrows():
                                q_id = clean_val(q["id"])
                                q_metni = q['soru_metni']
                                q_val = str(st.session_state["answers"].get(q_id, "")).strip()
                                q_type = q["soru_tipi"]
                                
                                parent_id = clean_val(q.get('bagli_parent_id'), default="0")
                                parent_target = clean_val(q.get('bagli_parent_deger'), default="")
                                
                                if parent_id != "0" and parent_id != "":
                                    parent_ans = str(st.session_state["answers"].get(parent_id, "")).strip()
                                    if parent_ans != parent_target:
                                        new_row_data[q_metni] = ""
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
                                        q_val = format_phone(clean_p)
                                        st.session_state["answers"][q_id] = q_val
                                
                                new_row_data[q_metni] = q_val
                            
                            if validation_errors:
                                for err in validation_errors: st.error(err)
                            else:
                                if df_yanitlar.empty or "numara" not in df_yanitlar.columns:
                                    df_yanitlar = pd.DataFrame(columns=["numara", "tarih"])
                                
                                if is_update and not df_yanitlar.empty and secilen_no in df_yanitlar["numara"].values:
                                    idx_to_update = df_yanitlar[df_yanitlar["numara"] == secilen_no].index[0]
                                    for col, val in new_row_data.items():
                                        df_yanitlar.loc[idx_to_update, col] = val
                                else:
                                    new_row_df = pd.DataFrame([new_row_data])
                                    df_yanitlar = pd.concat([df_yanitlar, new_row_df], ignore_index=True)
                                
                                save_data("yanitlar", df_yanitlar)
                                st.success("✅ Form yanıtlarınız başarıyla kaydedildi!")

# --- TAB 2: YÖNETİCİ & ÖĞRETMEN PANATELİ ---
with tab2:
    st.subheader("Yönetici & Öğretmen Paneli")
    sifre = st.text_input("Yönetici Şifresi:", type="password")
    
    if sifre == "admin123":
        df_o = get_data("ogrenciler")
        df_y = get_data("yanitlar")
        df_q = get_data("sorular")
        
        if df_o.empty or "numara" not in df_o.columns:
            df_o = pd.DataFrame(columns=["numara", "sinif", "sube", "ogretmen", "ad_soyad"])
            
        if df_y.empty or "numara" not in df_y.columns:
            df_y = pd.DataFrame(columns=["numara", "tarih"])
            
        if df_q.empty or "id" not in df_q.columns:
            df_q = pd.DataFrame(columns=["id", "soru_metni", "soru_tipi", "secenekler", "sira", "bagli_parent_id", "bagli_parent_deger"])
        else:
            if "sira" in df_q.columns: 
                df_q["sira"] = pd.to_numeric(df_q["sira"], errors='coerce').fillna(999)
                df_q = df_q.sort_values(by=["sira", "id"])
        
        merged_all = pd.merge(df_o, df_y, on='numara', how='left') if not df_o.empty else pd.DataFrame()
        if not merged_all.empty:
            merged_all['FORM DURUMU'] = merged_all['tarih'].apply(lambda x: "DOLDURDU" if pd.notna(x) and str(x).strip() not in ["", "nan"] else "DOLDURMADI")
            merged_all['sinif_sube'] = merged_all['sinif'].astype(str) + "/" + merged_all['sube'].astype(str)
        
        sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs([
            "📊 İstatistikler", "📗 Excel Raporu", "📄 PDF Dökümleri", "🛠️ Soru & e-Okul Yönetimi"
        ])
        
        with sub_tab1:
            st.markdown("### 📈 Genel ve Sınıf Bazlı Durum Takibi")
            if not merged_all.empty:
                toplam_ogr = len(df_o)
                dolduran_ogr = len(merged_all[merged_all['FORM DURUMU'] == 'DOLDURDU'])
                doldurmayan_ogr = toplam_ogr - dolduran_ogr
                orani = int((dolduran_ogr / toplam_ogr) * 100) if toplam_ogr > 0 else 0
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Toplam Öğrenci", toplam_ogr)
                col2.metric("Form Dolduran", dolduran_ogr)
                col3.metric("Doldurmayan", doldurmayan_ogr)
                col4.metric("Tamamlanma Oranı", f"%{orani}")
                
                st.divider()
                st.markdown("#### Sınıf/Şube Bazında Doldurma Durumları")
                stats_df = merged_all.groupby('sinif_sube')['FORM DURUMU'].value_counts().unstack(fill_value=0)
                if 'DOLDURDU' not in stats_df.columns: stats_df['DOLDURDU'] = 0
                if 'DOLDURMADI' not in stats_df.columns: stats_df['DOLDURMADI'] = 0
                stats_df['TOPLAM'] = stats_df['DOLDURDU'] + stats_df['DOLDURMADI']
                stats_df['TAMAMLANMA %'] = ((stats_df['DOLDURDU'] / stats_df['TOPLAM']) * 100).round(1)
                st.dataframe(stats_df[['TOPLAM', 'DOLDURDU', 'DOLDURMADI', 'TAMAMLANMA %']], use_container_width=True)
                
                with st.expander("🚨 Formu Henüz Doldurmayan Öğrenciler Listesi"):
                    doldurmayanlar = merged_all[merged_all['FORM DURUMU'] == 'DOLDURMADI'][['sinif_sube', 'numara', 'ad_soyad', 'ogretmen']]
                    st.dataframe(doldurmayanlar, use_container_width=True)
            else:
                st.info("Sistemde henüz öğrenci verisi yok. Lütfen 'Soru & e-Okul Yönetimi' sekmesinden e-Okul Excel dosyasını yükleyin.")

        with sub_tab2:
            st.markdown("### 📗 Toplu Excel İndirme")
            if not merged_all.empty and not df_q.empty:
                excel_rows = []
                for idx, r in merged_all.iterrows():
                    row_data = {
                        "SINIF/ŞUBE": r.get('sinif_sube', ''), 
                        "OKUL NO": r['numara'], 
                        "ADI SOYADI": r['ad_soyad'],
                        "DURUM": r['FORM DURUMU']
                    }
                    for _, q in df_q.iterrows():
                        q_metni = q['soru_metni']
                        val = str(r.get(q_metni, "")).strip()
                        if val.lower() in ["nan", "none"]: val = ""
                        row_data[q_metni] = val
                        
                    excel_rows.append(row_data)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pd.DataFrame(excel_rows).to_excel(writer, index=False)
                st.download_button("📊 TÜM VERİLERİ EXCEL OLARAK İNDİR", data=output.getvalue(), file_name="OKUL_BILGI_FORMU_RAPOR.xlsx")
            else:
                st.info("Rapor oluşturmak için yeterli veri bulunamadı.")

        with sub_tab3:
            st.markdown("### 📄 Sınıf Bazlı PDF Dökümleri")
            if not merged_all.empty and not df_q.empty and 'sinif_sube' in merged_all.columns:
                tum_subeler = sorted(merged_all['sinif_sube'].unique().tolist(), key=sort_sinif_sube_key)
                secilen_pdf_sube = st.selectbox("Sınıf Seçin:", tum_subeler)
                if st.button("PDF Raporu Oluştur"):
                    sube_students = merged_all[merged_all['sinif_sube'] == secilen_pdf_sube]
                    pdf_bytes = generate_class_pdf(sube_students, df_q, secilen_pdf_sube)
                    st.download_button(f"📥 {secilen_pdf_sube.replace('/', '_')}_Formlar.pdf İndir", data=bytes(pdf_bytes), file_name=f"{secilen_pdf_sube.replace('/', '_')}_Formlar.pdf")
            else:
                st.info("PDF oluşturmak için öğrenci ve soru verisi gereklidir.")

        with sub_tab4:
            st.markdown("### 🛠️ Sistem Yönetim Paneli")
            
            with st.expander("📥 e-Okul Excel Listesi Yükle / Güncelle", expanded=True):
                uploaded_file = st.file_uploader("e-Okul'dan aldığınız Sinif_Listesi.xls/xlsx dosyasını seçin:", type=["xls", "xlsx"])
                if uploaded_file and st.button("Veritabanına İşle"):
                    toplam = parse_and_save_eokul(uploaded_file)
                    st.success(f"✅ {toplam} öğrenci başarıyla Google Sheets veritabanına aktarıldı!")
                    st.rerun()

            st.divider()
            st.markdown("### 📝 Form Sorularını Yönet (CRUD)")
            
            q_islem = st.radio("Yapmak istediğiniz işlem:", ["Yeni Soru Ekle", "Mevcut Soruyu Düzenle / Sil"], horizontal=True)
            
            if q_islem == "Yeni Soru Ekle":
                with st.form("yeni_soru_form"):
                    y_metin = st.text_input("Soru Metni:")
                    y_tip = st.selectbox("Soru Tipi:", ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"])
                    y_secenekler = st.text_input("Seçenekler (Çoktan seçmeli veya çoklu seçim ise virgülle ayırın):", help="Örn: EVET,HAYIR veya DİYABET,ASTIM,YOK")
                    y_sira = st.number_input("Soru Sırası:", min_value=1, value=len(df_q) + 1 if not df_q.empty else 1)
                    
                    parent_opts = {"Yok (Ana Soru)": "0"}
                    if not df_q.empty and 'id' in df_q.columns:
                        for _, q_item in df_q.iterrows():
                            c_id = clean_val(q_item['id'])
                            if c_id: parent_opts[f"ID:{c_id} - {q_item['soru_metni']}"] = c_id
                    
                    y_parent_sel = st.selectbox("Bağlı Olduğu Üst Soru (Şartlı Gösterim):", list(parent_opts.keys()))
                    y_parent_val = st.text_input("Üst Sorunun Hangi Seçeneğinde Bu Soru Görünsün?:", help="Örn: DİĞER")
                    
                    if st.form_submit_button("➕ Soruyu Kaydet"):
                        if y_metin:
                            try:
                                max_id = int(max([int(clean_val(x)) for x in df_q['id'] if clean_val(x).isdigit()]))
                            except:
                                max_id = 0
                            new_id = str(max_id + 1)
                            
                            new_q = {
                                "id": new_id,
                                "soru_metni": y_metin,
                                "soru_tipi": y_tip,
                                "secenekler": y_secenekler,
                                "sira": y_sira,
                                "bagli_parent_id": parent_opts[y_parent_sel],
                                "bagli_parent_deger": y_parent_val
                            }
                            df_q_updated = pd.concat([df_q, pd.DataFrame([new_q])], ignore_index=True)
                            save_data("sorular", df_q_updated)
                            st.success("✅ Yeni soru kaydedildi!")
                            st.rerun()
                        else:
                            st.error("Lütfen soru metnini boş bırakmayın.")
                            
            elif q_islem == "Mevcut Soruyu Düzenle / Sil":
                if not df_q.empty and 'id' in df_q.columns:
                    q_dict = {f"ID:{clean_val(r['id'])} - {r['soru_metni']}": clean_val(r['id']) for _, r in df_q.iterrows() if clean_val(r['id'])}
                    if q_dict:
                        secilen_q_label = st.selectbox("Düzenlenecek Soruyu Seçin:", list(q_dict.keys()))
                        secilen_q_id = q_dict[secilen_q_label]
                        q_row = df_q[df_q['id'].apply(lambda x: clean_val(x)) == secilen_q_id].iloc[0]
                        
                        with st.form("duzenle_soru_form"):
                            d_metin = st.text_input("Soru Metni:", value=str(q_row['soru_metni']))
                            d_tip_idx = ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"].index(q_row['soru_tipi']) if q_row['soru_tipi'] in ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"] else 0
                            d_tip = st.selectbox("Soru Tipi:", ["coktan_secmeli", "coklu_secim", "metin", "tc_no", "telefon"], index=d_tip_idx)
                            d_secenekler = st.text_input("Seçenekler:", value=clean_val(q_row.get('secenekler')))
                            d_sira = st.number_input("Soru Sırası:", value=int(clean_val(q_row.get('sira'), "1") or 1))
                            
                            btn_col1, btn_col2 = st.columns(2)
                            guncelle = btn_col1.form_submit_button("💾 Güncelle")
                            sil = btn_col2.form_submit_button("🗑️ Soruyu Sil", type="primary")
                            
                            if guncelle:
                                mask = df_q['id'].apply(lambda x: clean_val(x)) == secilen_q_id
                                df_q.loc[mask, ['soru_metni', 'soru_tipi', 'secenekler', 'sira']] = [d_metin, d_tip, d_secenekler, d_sira]
                                save_data("sorular", df_q)
                                st.success("✅ Soru güncellendi!")
                                st.rerun()
                                
                            if sil:
                                df_q_updated = df_q[df_q['id'].apply(lambda x: clean_val(x)) != secilen_q_id]
                                save_data("sorular", df_q_updated)
                                st.success("🗑️ Soru silindi!")
                                st.rerun()
                    else:
                        st.info("Düzenlenecek soru bulunamadı.")
