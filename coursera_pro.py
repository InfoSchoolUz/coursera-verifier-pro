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
# 1. SAHIFA SOZLAMALARI VA DIZAYN
# ==========================================
st.set_page_config(page_title="Coursera Verifier Pro", layout="wide", page_icon="🎓")

st.markdown("""
<style>
    /* Global */
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(0, 255, 255, 0.12), transparent 28%),
            radial-gradient(circle at top right, rgba(255, 0, 204, 0.12), transparent 25%),
            linear-gradient(135deg, #050816 0%, #090d1f 45%, #04060f 100%);
        color: #ecf7ff;
    }

    [data-testid="stHeader"] {
        background: rgba(0, 0, 0, 0);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(8, 12, 28, 0.98) 0%, rgba(12, 18, 38, 0.96) 100%);
        border-right: 1px solid rgba(0, 255, 255, 0.18);
        box-shadow: 0 0 25px rgba(0, 255, 255, 0.08);
    }

    /* Main headings */
    h1, h2, h3 {
        color: #f5fbff !important;
        text-shadow: 0 0 8px rgba(0, 255, 255, 0.22);
    }

    .neon-title {
        padding: 20px 24px;
        border-radius: 18px;
        background: linear-gradient(135deg, rgba(0,255,255,0.12), rgba(255,0,204,0.10));
        border: 1px solid rgba(0,255,255,0.22);
        box-shadow:
            0 0 18px rgba(0,255,255,0.12),
            0 0 32px rgba(255,0,204,0.08),
            inset 0 0 14px rgba(255,255,255,0.03);
        margin-bottom: 16px;
    }

    .neon-subtitle {
        font-size: 1.02rem;
        color: #bfeeff;
        margin-top: 6px;
    }

    /* Cards / alerts */
    .stAlert, div[data-baseweb="notification"] {
        background: rgba(11, 20, 40, 0.82) !important;
        color: #e8faff !important;
        border: 1px solid rgba(0, 255, 255, 0.18) !important;
        border-radius: 16px !important;
        box-shadow: 0 0 18px rgba(0, 255, 255, 0.08);
    }

    /* Inputs */
    .stTextInput > div > div,
    .stNumberInput > div > div,
    .stSelectbox > div > div,
    .stMultiSelect > div > div,
    .stTextArea textarea,
    .stFileUploader,
    div[data-baseweb="select"] > div,
    div[data-testid="stFileUploaderDropzone"] {
        background: rgba(12, 20, 38, 0.88) !important;
        color: #ecf7ff !important;
        border: 1px solid rgba(0, 255, 255, 0.22) !important;
        border-radius: 14px !important;
        box-shadow: 0 0 12px rgba(0, 255, 255, 0.08);
    }

    label, .stMarkdown, .stCaption, .st-emotion-cache-10trblm, .st-emotion-cache-16idsys {
        color: #d7f7ff !important;
    }

    /* Buttons */
    .stButton > button,
    .stDownloadButton > button {
        width: 100%;
        border-radius: 14px;
        border: 1px solid rgba(0, 255, 255, 0.4);
        background: linear-gradient(90deg, rgba(0,255,255,0.18), rgba(255,0,204,0.18));
        color: #f8fdff;
        font-weight: 700;
        letter-spacing: 0.3px;
        box-shadow:
            0 0 14px rgba(0,255,255,0.18),
            0 0 20px rgba(255,0,204,0.10);
        transition: all 0.2s ease-in-out;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
        border-color: rgba(255, 0, 204, 0.5);
        box-shadow:
            0 0 20px rgba(0,255,255,0.25),
            0 0 24px rgba(255,0,204,0.16);
    }

    /* Slider */
    .stSlider [data-baseweb="slider"] > div div {
        background-color: #00f7ff !important;
    }

    /* Metric cards */
    div[data-testid="metric-container"] {
        background: linear-gradient(180deg, rgba(10,18,36,0.92), rgba(8,14,28,0.94));
        border: 1px solid rgba(0,255,255,0.18);
        padding: 16px 14px;
        border-radius: 18px;
        box-shadow:
            0 0 16px rgba(0,255,255,0.08),
            inset 0 0 12px rgba(255,255,255,0.02);
    }

    div[data-testid="metric-container"] label,
    div[data-testid="metric-container"] div {
        color: #effcff !important;
    }

    /* DataFrame */
    .stDataFrame {
        border: 1px solid rgba(0,255,255,0.2);
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 0 18px rgba(0,255,255,0.08);
    }

    /* Progress */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #00f7ff, #ff00cc) !important;
    }

    /* Divider */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0,255,255,0.45), transparent);
    }

    /* Footer */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background: rgba(4, 8, 18, 0.88);
        color: #c9f8ff;
        text-align: center;
        padding: 10px;
        font-weight: 700;
        border-top: 1px solid rgba(0, 255, 255, 0.2);
        box-shadow: 0 -4px 20px rgba(0,255,255,0.08);
        backdrop-filter: blur(8px);
        z-index: 1000;
    }

    a {
        color: #71f7ff !important;
        text-decoration: none !important;
    }

    a:hover {
        color: #ff75de !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SERTIFIKAT KODINI AJRATISH
# ==========================================
def extract_certificate_code(url):
    if pd.isna(url):
        return ""

    url = str(url).strip()
    if not url.startswith("http"):
        return ""

    try:
        parsed = urlparse(url)
        path = parsed.path.strip("/").lower()
        parts = path.split("/")

        if len(parts) >= 2 and parts[0] == "share":
            return parts[1].strip().lower()

        if "verify" in parts:
            idx = parts.index("verify")
            if idx + 1 < len(parts):
                return parts[idx + 1].strip().lower()

        if "accomplishments" in parts:
            for part in reversed(parts):
                part = part.strip().lower()
                if part and part not in [
                    "account", "accomplishments", "certificates",
                    "certificate", "verify", "share"
                ]:
                    return part

        match = re.search(r"(?:share|verify)/([^/?#]+)", url, re.IGNORECASE)
        if match:
            return match.group(1).strip().lower()

        return ""

    except Exception:
        return ""

# ==========================================
# 3. SERTIFIKAT SANASINI AJRATISH
# ==========================================
def extract_certificate_date(html):
    if not html:
        return ""

    try:
        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text(" ", strip=True)

        patterns = [
            r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}",
            r"\d{1,2}\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}",
            r"\b(20\d{2}-\d{2}-\d{2})\b",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(0)

        for script in soup.find_all("script"):
            script_text = script.get_text(" ", strip=True)
            for pattern in patterns:
                match = re.search(pattern, script_text, re.IGNORECASE)
                if match:
                    return match.group(0)

        return ""
    except Exception:
        return ""

# ==========================================
# 4. NETWORK SESSIYASI
# ==========================================
@st.cache_resource
def get_pro_session():
    session = requests.Session()
    retry = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=100, pool_maxsize=100)
    session.mount("https://", adapter)
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    })
    return session

# ==========================================
# 5. VERIFIKATSIYA MANTIQI
# ==========================================
def verify_link(session, url, timeout):
    if pd.isna(url) or not str(url).startswith("http"):
        return "MAVJUD EMAS", "-", "Havola topilmadi", ""

    url = str(url).strip()

    try:
        resp = session.get(url, timeout=timeout, allow_redirects=True)
        final_url = resp.url.lower()
        is_valid_path = any(x in final_url for x in ["/share/", "/verify/", "/accomplishments/"])
        cert_date = extract_certificate_date(resp.text) if resp.status_code == 200 else ""

        if resp.status_code == 200 and is_valid_path:
            return "MAVJUD", "200", "Tasdiqlandi ✅", cert_date
        elif "login" in final_url or "signup" in final_url:
            return "XATO", "Redirect", "Avtorizatsiya so'raldi (Xato link)", cert_date
        else:
            return "MAVJUD EMAS", str(resp.status_code), "Sertifikat sahifasi emas", cert_date

    except Exception:
        return "XATO", "Timeout/Error", "Ulanish imkonsiz", ""

# ==========================================
# 6. ASOSIY ILOVA
# ==========================================
def main():
    st.markdown(
        """
        <div class="neon-title">
            <h1 style="margin:0;">🎓 Coursera Certificate Verifier Pro</h1>
            <div class="neon-subtitle">Neon UI • Zamonaviy ko‘rinish • Bir oz cyberpunk, lekin ish baribir jiddiy 😎</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.sidebar:
        st.markdown("### 🛠 Dastur haqida")
        st.info("Coursera sertifikatlarini avtomatik tekshirish tizimi.")
        st.markdown("---")
        st.markdown("""
        <div class="footer">
            Developed by Azamat Madrimov 🚀 | 2026
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("### 📬 Muallifga murojaat")
        st.markdown("""
        <div style="line-height: 2;">
        <img src="https://img.icons8.com/color/20/gmail-new.png"/> 
        <a href="mailto:azamat3533141@gmail.com"> azamat3533141@gmail.com</a><br>

        <img src="https://img.icons8.com/color/20/telegram-app.png"/> 
        <a href="https://t.me/futurex_azamat"> @futurex_azamat</a>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        st.header("⚙️ Parametrlar")
        threads = st.slider("Parallel tekshiruvlar", 5, 50, 25)
        timeout = st.slider("Kutish vaqti (sekund)", 5, 30, 15)

    st.subheader("Maktab o'quvchilari sertifikatlarini avtomatik tekshirish tizimi")
    file = st.file_uploader(
        "Excel (.xlsx) yoki CSV faylni yuklang",
        type=["xlsx", "csv"],
        help="""
        Fayl quyidagi ustunlarda bo‘lishi kerak:

        • №  
        • Tuman/Shahar  
        • Maktab raqami  
        • Sinf  
        • F.I.SH  
        • Guvohnoma seriyasi va raqami  
        • Tug‘ilgan sana  
        • Sertifikat havolasi  
        • Elektron pochta
        """
    )

    if file:
        try:
            original_filename = file.name
            base_name, ext = os.path.splitext(original_filename)

            if file.name.endswith(".csv"):
                uploaded_sheets = {"CSV": pd.read_csv(file, skiprows=2)}
                selected_sheet = "CSV"
            else:
                uploaded_sheets = pd.read_excel(file, sheet_name=None, skiprows=2)

                available_sheet_names = list(uploaded_sheets.keys())
                selected_sheet = st.selectbox(
                    "Tekshirish uchun listni tanlang",
                    available_sheet_names,
                    index=0
                )

                uploaded_sheets = {selected_sheet: uploaded_sheets[selected_sheet]}

            prepared_sheets = []
            total_students = 0

            for sheet_name, df in uploaded_sheets.items():
                if df is None or df.empty:
                    continue

                df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
                all_cols = df.columns.tolist()

                if not all_cols:
                    continue

                fish_col = next(
                    (c for c in all_cols if "ФИШ" in c.upper() or "F.I.SH" in c.upper()),
                    all_cols[4] if len(all_cols) > 4 else all_cols[0]
                )

                course_cols = [
                    c for c in all_cols
                    if df[c].astype(str).str.contains("coursera.org", na=False).any()
                ]

                if not course_cols:
                    continue

                prepared_sheets.append({
                    "sheet_name": str(sheet_name)[:31],
                    "df": df,
                    "fish_col": fish_col,
                    "course_cols": course_cols
                })

                total_students += len(df)

            st.success(
                f"Ma'lumotlar yuklandi. Tanlangan list: {selected_sheet}. "
                f"Jami {total_students} ta o'quvchi aniqlandi."
            )

            if st.button("🚀 TEKSHIRISHNI BOSHLASH", type="primary", use_container_width=True):
                all_entries = []
                unique_code_to_url = {}
                unique_fallback_to_url = {}

                for sheet_info in prepared_sheets:
                    sheet_name = sheet_info["sheet_name"]
                    df = sheet_info["df"]
                    fish_col = sheet_info["fish_col"]
                    course_cols = sheet_info["course_cols"]

                    for _, row in df.iterrows():
                        for col in course_cols:
                            raw_value = row[col]
                            original_url = str(raw_value).strip()

                            if pd.notna(raw_value) and "http" in original_url:
                                cert_code = extract_certificate_code(original_url)

                                all_entries.append({
                                    "sheet_name": sheet_name,
                                    "name": row[fish_col],
                                    "course": col,
                                    "url": original_url,
                                    "cert_code": cert_code
                                })

                                if cert_code:
                                    if cert_code not in unique_code_to_url:
                                        unique_code_to_url[cert_code] = original_url
                                else:
                                    if original_url not in unique_fallback_to_url:
                                        unique_fallback_to_url[original_url] = original_url

                if not all_entries:
                    st.warning("Tekshirish uchun hech qanday sertifikat link topilmadi.")
                    return

                results_cache = {}
                fallback_results_cache = {}

                session = get_pro_session()
                progress = st.progress(0)
                status_box = st.empty()

                unique_items = list(unique_code_to_url.items())
                fallback_items = list(unique_fallback_to_url.items())
                all_unique_tasks = unique_items + fallback_items
                total_unique = len(all_unique_tasks)

                with ThreadPoolExecutor(max_workers=threads) as executor:
                    future_to_key = {}

                    for cert_code, original_url in unique_items:
                        future = executor.submit(verify_link, session, original_url, timeout)
                        future_to_key[future] = ("code", cert_code)

                    for fallback_key, original_url in fallback_items:
                        future = executor.submit(verify_link, session, original_url, timeout)
                        future_to_key[future] = ("url", fallback_key)

                    for i, future in enumerate(as_completed(future_to_key)):
                        key_type, key_value = future_to_key[future]
                        result = future.result()

                        if key_type == "code":
                            results_cache[key_value] = result
                        else:
                            fallback_results_cache[key_value] = result

                        progress.progress((i + 1) / total_unique)
                        status_box.text(f"Tekshirilmoqda: {i + 1}/{total_unique}")

                final_data = []
                seen_codes = set()
                seen_urls_without_code = set()

                for item in all_entries:
                    cert_code = item["cert_code"]
                    original_url = item["url"]

                    if cert_code and cert_code in results_cache:
                        status, code, reason, cert_date = results_cache[cert_code]
                    elif not cert_code and original_url in fallback_results_cache:
                        status, code, reason, cert_date = fallback_results_cache[original_url]
                    else:
                        status, code, reason, cert_date = "XATO", "CodeError", "Sertifikat kodi aniqlanmadi", ""

                    if cert_code:
                        if cert_code in seen_codes:
                            display_reason = "TAKRORLANUVCHI 🔄"
                        else:
                            display_reason = reason
                            seen_codes.add(cert_code)
                    else:
                        if original_url in seen_urls_without_code:
                            display_reason = "TAKRORLANUVCHI 🔄"
                        else:
                            display_reason = reason
                            seen_urls_without_code.add(original_url)

                    final_data.append({
                        "F.I.SH": item["name"],
                        "Kurs yo'nalishi": item["course"],
                        "Holati": status,
                        "Natija": display_reason,
                        "Havola": original_url,
                        "Sertifikat kodi": cert_code,
                        "Sertifikat olingan sana": cert_date,
                        "__sheet_name__": item["sheet_name"]
                    })

                res_df = pd.DataFrame(final_data)

                duplicate_count = int((res_df["Natija"] == "TAKRORLANUVCHI 🔄").sum())
                confirmed_count = int(
                    ((res_df["Holati"] == "MAVJUD") & (res_df["Natija"] != "TAKRORLANUVCHI 🔄")).sum()
                )
                error_count = int((res_df["Holati"] != "MAVJUD").sum())

                st.divider()
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Jami tekshirildi", len(res_df))
                c2.metric("Tasdiqlandi ✅", confirmed_count)
                c3.metric("Xato/Mavjud emas ❌", error_count)
                c4.metric("Takrorlanuvchi 🔄", duplicate_count)

                st.caption(
                    f"Unikal sertifikat kodlari: {res_df['Sertifikat kodi'].replace('', pd.NA).nunique()} | Takrorlar: {duplicate_count}"
                )

                st.subheader("📋 Batafsil hisobot")
                display_df = res_df.drop(columns=["__sheet_name__"])
                st.dataframe(
                    display_df.style.map(
                        lambda x: 'background-color: #0e3b4b; color: #d8fbff;' if x == 'MAVJUD'
                        else 'background-color: #4a1030; color: #ffe3f3;' if x == 'XATO'
                        else 'background-color: #4b3b0e; color: #fff6d8;' if x == 'MAVJUD EMAS'
                        else 'background-color: #18294f; color
