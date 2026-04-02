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

st.set_page_config(page_title="Coursera Verifier Pro", layout="wide", page_icon="🎓")

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #0a0f1f 0%, #0e1117 45%, #12071f 100%);
    color: #e6f7ff;
}

[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #0a0f1f 0%, #0e1117 45%, #12071f 100%);
}

[data-testid="stHeader"] {
    background: rgba(0, 0, 0, 0);
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #111827 0%, #0b1220 100%) !important;
    border-right: 2px solid #00f2ff;
    box-shadow: 0 0 25px rgba(0, 242, 255, 0.20);
}

h1 {
    color: #00f2ff !important;
    text-shadow: 0 0 8px #00f2ff, 0 0 18px #00f2ff;
    text-align: center;
    font-weight: 800 !important;
}

h2, h3 {
    color: #ff4dff !important;
    text-shadow: 0 0 6px #ff4dff, 0 0 14px #ff4dff;
    font-weight: 700 !important;
}

p, label, div, span {
    color: #e6f7ff;
}

.stButton > button,
.stDownloadButton > button {
    background: transparent !important;
    color: #39ff14 !important;
    border: 2px solid #39ff14 !important;
    border-radius: 16px !important;
    font-weight: 700 !important;
    box-shadow: 0 0 10px rgba(57, 255, 20, 0.35);
    transition: all 0.25s ease-in-out;
    width: 100%;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #39ff14 !important;
    color: #000 !important;
    box-shadow: 0 0 20px #39ff14, 0 0 35px rgba(57, 255, 20, 0.55);
    transform: translateY(-1px);
}

.stTextInput > div > div > input,
.stNumberInput input,
.stTextArea textarea,
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div {
    background-color: #101827 !important;
    color: #e6f7ff !important;
    border: 1px solid #00f2ff !important;
    border-radius: 12px !important;
    box-shadow: 0 0 10px rgba(0, 242, 255, 0.12);
}

[data-testid="stFileUploader"] {
    background: rgba(0, 242, 255, 0.05);
    border: 1px solid #00f2ff;
    border-radius: 14px;
    padding: 10px;
    box-shadow: 0 0 15px rgba(0, 242, 255, 0.12);
}

[data-testid="stMetric"] {
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(0, 242, 255, 0.35);
    border-radius: 16px;
    padding: 10px;
    box-shadow: 0 0 12px rgba(0, 242, 255, 0.10);
}

.stDataFrame, div[data-testid="stDataFrame"] {
    border: 1px solid #00f2ff !important;
    border-radius: 12px !important;
    box-shadow: 0 0 14px rgba(0, 242, 255, 0.16);
    overflow: hidden;
}

div[data-testid="stAlert"] {
    border-radius: 14px;
    border: 1px solid rgba(0, 242, 255, 0.35);
    box-shadow: 0 0 12px rgba(0, 242, 255, 0.10);
}

hr {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, #00f2ff, transparent);
}

.contact-box {
    background: rgba(10, 16, 31, 0.85);
    padding: 15px;
    border-radius: 14px;
    border: 1px solid #00f2ff;
    box-shadow: 0 0 14px rgba(0, 242, 255, 0.18);
    margin-bottom: 18px;
}

.contact-box a {
    color: #7df9ff !important;
    text-decoration: none;
}

.contact-box a:hover {
    text-shadow: 0 0 8px #7df9ff;
}

.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background: rgba(0, 0, 0, 0.92);
    color: #00f2ff;
    text-align: center;
    padding: 10px;
    font-weight: 700;
    border-top: 1px solid #00f2ff;
    z-index: 999;
    box-shadow: 0 -2px 15px rgba(0, 242, 255, 0.20);
}

.block-container {
    padding-bottom: 80px !important;
}

[data-testid="stSidebar"] .stSlider label,
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] .stHeader {
    color: #e6f7ff !important;
}
</style>
""", unsafe_allow_html=True)


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


def main():
    st.title("🎓 Coursera Certificate Verifier Pro")

    with st.sidebar:
        st.markdown("""
        <div class="contact-box">
            <h3 style="margin-top:0;">📬 Aloqa</h3>
            <p style="margin-bottom:6px;">📧 Email: <a href="mailto:azamat3533141@gmail.com">azamat3533141@gmail.com</a></p>
            <p style="margin-bottom:0;">✈️ Telegram: <a href="https://t.me/futurex_azamat" target="_blank">@futurex_azamat</a></p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 🛠 Dastur haqida")
        st.info("Coursera sertifikatlarini avtomatik tekshirish tizimi.")
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

                styled_df = display_df.style.applymap(
                    lambda x: "background-color: #163d2a; color: #eaffea;" if x == "MAVJUD"
                    else "background-color: #4a1f2b; color: #ffecec;" if x == "XATO"
                    else "background-color: #4a3f14; color: #fff6d6;" if x == "MAVJUD EMAS"
                    else "",
                    subset=["Holati"]
                )

                st.dataframe(styled_df, use_container_width=True)

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

    st.markdown(
        '<div class="footer">Tuzuvchi: Azamat Madrimov | 2026</div>',
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
