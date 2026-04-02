import io
import re
import os
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import urlparse

# ==========================================
# 1. SAHIFA SOZLAMALARI VA NEON DIZAYN
# ==========================================
st.set_page_config(page_title="Coursera Verifier Pro", layout="wide", page_icon="🎓")

st.markdown("""
    <style>
    /* Umumiy Neon Fon */
    .stApp {
        background-color: #0e1117;
        color: #e0e0e0;
    }
    
    /* Chap panel (Sidebar) dizayni */
    [data-testid="stSidebar"] {
        background-color: #161b22 !important;
        border-right: 2px solid #00f2ff;
        box-shadow: 5px 0 15px rgba(0, 242, 255, 0.2);
    }

    h1 {
        color: #00f2ff !important;
        text-shadow: 0 0 10px #00f2ff, 0 0 20px #00f2ff;
        text-align: center;
    }

    h2, h3 {
        color: #ff00ff !important;
        text-shadow: 0 0 5px #ff00ff;
    }

    /* Footer dizayni */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #000000;
        color: #00f2ff;
        text-align: center;
        padding: 10px;
        font-weight: bold;
        z-index: 1000;
        border-top: 1px solid #00f2ff;
    }

    /* Tugmalar */
    .stButton>button {
        background-color: transparent !important;
        color: #00ff00 !important;
        border: 2px solid #00ff00 !important;
        box-shadow: 0 0 10px #00ff00;
        border-radius: 20px;
        font-weight: bold;
        width: 100%;
    }
    
    .stButton>button:hover {
        background-color: #00ff00 !important;
        color: black !important;
    }

    /* DataFrame stili */
    .stDataFrame {
        border: 1px solid #00f2ff !important;
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SERTIFIKAT KODINI AJRATISH (O'ZGARISHSIZ)
# ==========================================
def extract_certificate_code(url):
    if pd.isna(url): return ""
    url = str(url).strip()
    if not url.startswith("http"): return ""
    try:
        parsed = urlparse(url)
        path = parsed.path.strip("/").lower()
        parts = path.split("/")
        if len(parts) >= 2 and parts[0] == "share": return parts[1].strip().lower()
        if "verify" in parts:
            idx = parts.index("verify")
            if idx + 1 < len(parts): return parts[idx + 1].strip().lower()
        match = re.search(r"(?:share|verify)/([^/?#]+)", url, re.IGNORECASE)
        return match.group(1).strip().lower() if match else ""
    except: return ""

# ==========================================
# 3. SERTIFIKAT SANASINI AJRATISH (O'ZGARISHSIZ)
# ==========================================
def extract_certificate_date(html):
    if not html: return ""
    try:
        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text(" ", strip=True)
        patterns = [
            r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}",
            r"\b(20\d{2}-\d{2}-\d{2})\b",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match: return match.group(0)
        return ""
    except: return ""

# ==========================================
# 4. NETWORK SESSIYASI (O'ZGARISHSIZ)
# ==========================================
@st.cache_resource
def get_pro_session():
    session = requests.Session()
    retry = Retry(total=5, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry, pool_connections=100, pool_maxsize=100)
    session.mount("https://", adapter)
    session.headers.update({"User-Agent": "Mozilla/5.0 Chrome/122.0.0.0"})
    return session

# ==========================================
# 5. VERIFIKATSIYA MANTIQI (O'ZGARISHSIZ)
# ==========================================
def verify_link(session, url, timeout):
    if pd.isna(url) or not str(url).startswith("http"):
        return "MAVJUD EMAS", "-", "Havola topilmadi", ""
    try:
        resp = session.get(url, timeout=timeout, allow_redirects=True)
        final_url = resp.url.lower()
        is_valid = any(x in final_url for x in ["/share/", "/verify/", "/accomplishments/"])
        cert_date = extract_certificate_date(resp.text) if resp.status_code == 200 else ""
        if resp.status_code == 200 and is_valid:
            return "MAVJUD", "200", "Tasdiqlandi ✅", cert_date
        return "MAVJUD EMAS", str(resp.status_code), "Sertifikat sahifasi emas", cert_date
    except:
        return "XATO", "Timeout", "Ulanish imkonsiz", ""

# ==========================================
# 6. ASOSIY ILOVA
# ==========================================
def main():
    # CHAP PANEL (SIDEBAR) - DOIM KO'RINADI
    with st.sidebar:
        st.markdown("### 🛠 Dastur haqida")
        st.info("Coursera sertifikatlarini avtomatik tekshirish tizimi.")
        st.markdown("---")
        st.markdown("### 📬 Muallifga murojaat")
        st.markdown("""
        <div style="line-height: 2; background: #1a1a1a; padding: 10px; border-radius: 10px; border: 1px solid #00f2ff;">
            <img src="https://img.icons8.com/color/20/gmail-new.png"/> 
            <a href="mailto:azamat3533141@gmail.com" style="color:#00f2ff; text-decoration:none;"> azamat3533141@gmail.com</a><br>
            <img src="https://img.icons8.com/color/20/telegram-app.png"/> 
            <a href="https://t.me/futurex_azamat" style="color:#00f2ff; text-decoration:none;"> @futurex_azamat</a>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        st.header("⚙️ Parametrlar")
        threads = st.slider("Parallel tekshiruvlar", 5, 50, 25)
        timeout = st.slider("Kutish vaqti (sekund)", 5, 30, 15)

    st.title("🎓 Coursera Certificate Verifier Pro")
    st.subheader("Maktab o'quvchilari sertifikatlarini avtomatik tekshirish tizimi")

    file = st.file_uploader("Excel (.xlsx) yoki CSV faylni yuklang", type=["xlsx", "csv"])

    if file:
        try:
            original_filename = file.name
            base_name = os.path.splitext(original_filename)[0]
            
            if file.name.endswith(".csv"):
                df = pd.read_csv(file, skiprows=2)
                selected_sheet = "CSV"
                uploaded_sheets = {selected_sheet: df}
            else:
                xls = pd.ExcelFile(file)
                available_sheets = xls.sheet_names
                selected_sheet = st.selectbox("Tekshirish uchun listni tanlang", available_sheets)
                df = pd.read_excel(file, sheet_name=selected_sheet, skiprows=2)
                uploaded_sheets = {selected_sheet: df}

            prepared_sheets = []
            for sheet_name, df_sheet in uploaded_sheets.items():
                df_sheet.columns = [str(c).replace("\n", " ").strip() for c in df_sheet.columns]
                all_cols = df_sheet.columns.tolist()
                fish_col = next((c for c in all_cols if "ФИШ" in c.upper() or "F.I.SH" in c.upper()), all_cols[4] if len(all_cols)>4 else all_cols[0])
                course_cols = [c for c in all_cols if df_sheet[c].astype(str).str.contains("coursera.org", na=False).any()]
                if course_cols:
                    prepared_sheets.append({"sheet_name": sheet_name, "df": df_sheet, "fish_col": fish_col, "course_cols": course_cols})

            st.success(f"Yuklandi: {selected_sheet}. Jami {len(df)} ta qator.")

            if st.button("🚀 TEKSHIRISHNI BOSHLASH"):
                all_entries = []
                unique_code_to_url = {}
                unique_fallback_to_url = {}

                for sheet_info in prepared_sheets:
                    for _, row in sheet_info["df"].iterrows():
                        for col in sheet_info["course_cols"]:
                            url = str(row[col]).strip()
                            if "http" in url:
                                code = extract_certificate_code(url)
                                all_entries.append({"sheet_name": sheet_info["sheet_name"], "name": row[sheet_info["fish_col"]], "course": col, "url": url, "cert_code": code})
                                if code: unique_code_to_url[code] = url
                                else: unique_fallback_to_url[url] = url

                if not all_entries:
                    st.warning("Linklar topilmadi.")
                    return

                results_cache = {}
                fallback_results_cache = {}
                session = get_pro_session()
                progress = st.progress(0)
                status_box = st.empty()

                tasks = list(unique_code_to_url.items()) + list(unique_fallback_to_url.items())
                total = len(tasks)

                with ThreadPoolExecutor(max_workers=threads) as executor:
                    futures = {executor.submit(verify_link, session, url, timeout): (k, v) for k, v in tasks}
                    for i, f in enumerate(as_completed(futures)):
                        key, url = futures[f]
                        res = f.result()
                        if key in unique_code_to_url: results_cache[key] = res
                        else: fallback_results_cache[key] = res
                        progress.progress((i + 1) / total)
                        status_box.text(f"Tekshirilmoqda: {i+1}/{total}")

                final_data = []
                seen_codes = set()
                for item in all_entries:
                    code, url = item["cert_code"], item["url"]
                    res = results_cache.get(code) if code else fallback_results_cache.get(url)
                    status, http_code, reason, date = res if res else ("XATO", "", "", "")
                    
                    if code and code in seen_codes: display_reason = "TAKRORLANUVCHI 🔄"
                    else: display_reason = reason; seen_codes.add(code) if code else None

                    final_data.append({"F.I.SH": item["name"], "Kurs": item["course"], "Holati": status, "Natija": display_reason, "Havola": url, "Sertifikat kodi": code, "Sana": date, "__sheet__": item["sheet_name"]})

                res_df = pd.DataFrame(final_data)
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Jami", len(res_df))
                c2.metric("Tasdiqlandi ✅", len(res_df[res_df["Holati"]=="MAVJUD"]))
                c3.metric("Xato ❌", len(res_df[res_df["Holati"]!="MAVJUD"]))

                st.dataframe(res_df.drop(columns=["__sheet__"]), use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for s_name in res_df["__sheet__"].unique():
                        res_df[res_df["__sheet__"]==s_name].drop(columns=["__sheet__"]).to_excel(writer, index=False, sheet_name=str(s_name)[:31])
                
                st.download_button("📥 Excelni yuklab olish", data=output.getvalue(), file_name=f"{base_name}_Verify.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        except Exception as e:
            st.error(f"Xatolik: {e}")

    st.markdown('<div class="footer">Developed by Azamat Madrimov | 2026</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
                                                                          
