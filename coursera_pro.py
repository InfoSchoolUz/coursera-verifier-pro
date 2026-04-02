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

# Neon UI uchun CSS
st.markdown("""
    <style>
    /* Umumiy fon va matn rangi (Dark mode) */
    .stApp {
        background-color: #0e1117;
        color: #e0e0e0;
    }

    /* Asosiy Neon Sarlavha */
    h1 {
        color: #00f2ff !important;
        text-shadow: 0 0 10px #00f2ff, 0 0 20px #00f2ff, 0 0 30px #00f2ff;
        text-align: center;
        font-family: 'Courier New', monospace;
        padding-bottom: 20px;
    }

    /* Kichik sarlavhalar */
    h2, h3 {
        color: #ff00ff !important;
        text-shadow: 0 0 5px #ff00ff;
    }

    /* Sidebar dizayni */
    [data-testid="stSidebar"] {
        background-color: #161b22 !important;
        border-right: 2px solid #00f2ff;
        box-shadow: 5px 0 15px rgba(0, 242, 255, 0.2);
    }

    /* Sidebar ichidagi matnlar */
    [data-testid="stSidebar"] .stMarkdown p {
        color: #e0e0e0;
    }

    /* Input va Selectboxlar */
    .stTextInput>div>div>input, .stSelectbox>div>div>div {
        background-color: #1a1a1a !important;
        color: #00f2ff !important;
        border: 1px solid #00f2ff !important;
        box-shadow: 0 0 5px rgba(0, 242, 255, 0.5);
    }

    /* Slider rangi */
    .stSlider > div > div > div > div {
        background-color: #ff00ff !important;
    }

    /* Asosiy harakat tugmasi (🚀 TEKSHIRISHNI BOSHLASH) */
    .stButton>button {
        background-color: transparent !important;
        color: #00ff00 !important;
        border: 2px solid #00ff00 !important;
        border-radius: 20px !important;
        box-shadow: 0 0 10px #00ff00, inset 0 0 5px #00ff00;
        transition: all 0.3s ease;
        font-weight: bold !important;
        font-size: 18px !important;
        width: 100%;
    }

    .stButton>button:hover {
        background-color: #00ff00 !important;
        color: #000000 !important;
        box-shadow: 0 0 30px #00ff00, 0 0 50px #00ff00;
        transform: scale(1.02);
    }

    /* Yuklab olish tugmasi (📥 Excelni yuklab olish) */
    .stDownloadButton>button {
        background-color: transparent !important;
        color: #ff00ff !important;
        border: 2px solid #ff00ff !important;
        border-radius: 10px !important;
        box-shadow: 0 0 10px #ff00ff;
        font-weight: bold !important;
    }

    .stDownloadButton>button:hover {
        background-color: #ff00ff !important;
        color: #ffffff !important;
        box-shadow: 0 0 20px #ff00ff;
    }

    /* Progress bar */
    .stProgress > div > div > div > div {
        background-image: linear-gradient(to right, #00f2ff, #ff00ff) !important;
        box-shadow: 0 0 10px #00f2ff;
    }

    /* Metric kartalari (Natijalar) */
    [data-testid="stMetricValue"] {
        color: #00ff00 !important;
        text-shadow: 0 0 10px #00ff00;
        font-family: 'Orbitron', sans-serif;
    }
    
    [data-testid="stMetricLabel"] {
        color: #e0e0e0 !important;
    }

    /* DataFrame (Jadval) chegarasi */
    .stDataFrame {
        border: 1px solid #00f2ff !important;
        border-radius: 10px;
        box-shadow: 0 0 10px rgba(0, 242, 255, 0.3);
    }

    /* File Uploader */
    .stFileUploader > div > button {
        border: 1px dashed #00f2ff !important;
        color: #00f2ff !important;
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
        padding: 8px;
        font-weight: bold;
        z-index: 1000;
        border-top: 1px solid #00f2ff;
        box-shadow: 0 -5px 15px rgba(0, 242, 255, 0.2);
        font-family: 'Orbitron', sans-serif;
    }

    /* Sidebar info box */
    .stAlert {
        background-color: #1a1a1a !important;
        color: #e0e0e0 !important;
        border: 1px solid #ff00ff !important;
        box-shadow: 0 0 5px #ff00ff;
    }
    
    /* Hide default Streamlit elements */
    [data-testid="stToolbar"] { visibility: hidden; }
    footer { visibility: hidden; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SERTIFIKAT KODINI AJRATISH (O'ZGARISHSIZ)
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
# 3. SERTIFIKAT SANASINI AJRATISH (O'ZGARISHSIZ)
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
# 4. NETWORK SESSIYASI (O'ZGARISHSIZ)
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
# 5. VERIFIKATSIYA MANTIQI (O'ZGARISHSIZ)
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
    st.title("🎓 Coursera Certificate Verifier Pro")

    with st.sidebar:
        st.markdown("### 🛠 Dastur haqida")
        st.info("Coursera sertifikatlarini avtomatik tekshirish tizimi.")
        st.markdown("---")
        # Sidebar footer (neon rangda)
        st.markdown("""
        <div style="text-align:center; color:#00f2ff; font-weight:bold; text-shadow:0 0 5px #00f2ff;">
            Developed by Azamat Madrimov 🚀 | 2026
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("### 📬 Muallifga murojaat")
        st.markdown("""
        <div style="line-height: 2;">
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
                # Hisob-kitob mantiqi boshlanishi (O'ZGARISHSIZ)
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
                
                # Jadval stillari (Neon uslubida)
                st.dataframe(
                    display_df.style.map(
                        lambda x: 'background-color: #0b2e13; color: #00ff00; font-weight: bold;' if x == 'MAVJUD'
                        else 'background-color: #4c1111; color: #ff0000; font-weight: bold;' if x == 'XATO'
                        else 'background-color: #4a3b05; color: #ffc107; font-weight: bold;' if x == 'MAVJUD EMAS'
                        else 'color: #e0e0e0;',
                        subset=["Holati"]
                    ),
                    use_container_width=True
                )

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for sheet_name in res_df["__sheet_name__"].dropna().unique():
                        sheet_df = res_df[res_df["__sheet_name__"] == sheet_name].drop(columns=["__sheet_name__"])
                        if not sheet_df.empty:
                            safe_sheet_name = str(sheet_name)[:31]
                            sheet_df.to_excel(writer, index=False, sheet_name=safe_sheet_name)

                download_filename = f"{base_name}_Verify.xlsx"

                st.download_button(
                    label="📥 Excelni yuklab olish",
                    data=output.getvalue(),
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Xatolik: {e}")

    # Asosiy sahifa footeri
    st.markdown("""
        <div class="footer">
            Tuzuvchi: Azamat Madrimov | 2026
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
