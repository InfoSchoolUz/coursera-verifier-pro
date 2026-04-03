import io
import re
import os
import sys
import difflib
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
st.set_page_config(page_title="Coursera Certificate Verifier Pro", layout="wide", page_icon="🎓")
st.markdown("""
    <style>
    .reportview-container { background: #f0f2f6; }
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #0e1117;
        color: white;
        text-align: center;
        padding: 10px;
        font-weight: bold;
        z-index: 1000;
    }
    .sample-box {
        padding: 14px;
        border-radius: 14px;
        background: linear-gradient(135deg, rgba(0,242,254,0.10), rgba(79,172,254,0.10));
        border: 1px solid rgba(79,172,254,0.30);
        margin-bottom: 12px;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. NAMUNA FAYL SOZLAMALARI
# ==========================================
TEMPLATE_FILENAME = "coursera_template.xlsx"

def resolve_path(path: str) -> str:
    base_dir = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_dir, path)

def get_template_bytes():
    try:
        template_path = resolve_path(TEMPLATE_FILENAME)
        if os.path.exists(template_path):
            with open(template_path, "rb") as f:
                return f.read()
    except Exception:
        pass
    return None

def show_sample_download_section():
    st.markdown("### 📥 Namuna fayl")
    st.markdown("""
    <div class="sample-box">
        <b>Faylni xato formatda yuklamaslik uchun tayyor namuna ishlating.</b><br>
        Excel ustun nomlarini o'zgartirmang.
    </div>
    """, unsafe_allow_html=True)
    template_bytes = get_template_bytes()
    if template_bytes:
        st.download_button(
            label="📥 Namuna faylni yuklab olish",
            data=template_bytes,
            file_name=TEMPLATE_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.error("Namuna fayl topilmadi ❌")

# ==========================================
# 3. SERTIFIKAT KODINI AJRATISH
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
# 4. SERTIFIKAT SANASINI AJRATISH
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
# 5. ISM MOSLIGINI TEKSHIRISH — PRO
# ==========================================
def transliterate_cyrillic_to_latin(text: str) -> str:
    translit_map = {
        "а": "a", "б": "b", "в": "v", "г": "g", "д": "d",
        "е": "e", "ё": "yo", "ж": "j", "з": "z", "и": "i",
        "й": "y", "к": "k", "л": "l", "м": "m", "н": "n",
        "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
        "у": "u", "ф": "f", "х": "x", "ц": "s", "ч": "ch",
        "ш": "sh", "щ": "sh", "ъ": "", "ь": "", "э": "e",
        "ю": "yu", "я": "ya",
        "ў": "o'", "қ": "q", "ғ": "g'", "ҳ": "h",
        "ы": "i",
    }
    return "".join(translit_map.get(ch, ch) for ch in text)

def standardize_apostrophes(text: str) -> str:
    return (
        text.replace("’", "'")
            .replace("‘", "'")
            .replace("`", "'")
            .replace("ʻ", "'")
            .replace("ʼ", "'")
            .replace("ʹ", "'")
            .replace("´", "'")
            .replace("o‘", "o'")
            .replace("g‘", "g'")
    )

def simplify_token(token: str) -> str:
    """
    Fuzzy match uchun tokenni soddalashtiradi:
    - apostrofni olib tashlaydi
    - yo/yu/ya kabi yozuvlarni normalizatsiya qiladi
    """
    token = token.lower().strip()
    token = standardize_apostrophes(token)
    token = token.replace("'", "")
    token = token.replace("yo", "io")
    token = token.replace("yu", "iu")
    token = token.replace("ya", "ia")
    token = token.replace("ts", "s")
    token = token.replace("iy", "i")
    token = token.replace("yy", "y")
    token = re.sub(r"[^a-z0-9-]", "", token)
    return token

def normalize_name(name: str) -> list[str]:
    """
    Ismni normallashtirib, tokenlar ro'yxatini qaytaradi.
    Kuchli versiya:
    - kirill -> lotin
    - apostrof birxillashtirish
    - qizi/o'g'li kabi shovqinlarni chiqarib tashlash
    - keraksiz qismlar va bitta harfli bo'laklarni tozalash
    """
    if not name or not isinstance(name, str):
        return []

    name = str(name).strip().lower()
    name = standardize_apostrophes(name)
    name = transliterate_cyrillic_to_latin(name)

    # Ba'zi odatiy variantlarni birxillashtirish
    replacements = {
        "o g li": "ogli",
        "o g'li": "ogli",
        "o'g'li": "ogli",
        "o‘g‘li": "ogli",
        "ugli": "ogli",
        "u g li": "ogli",
        "qizi": "qizi",
        "kizi": "qizi",
    }
    for old, new in replacements.items():
        name = name.replace(old, new)

    name = re.sub(r"[^a-z0-9'\s-]", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()

    raw_tokens = [t.strip() for t in name.split() if t.strip()]

    stop_words = {
        "qizi", "qiz", "ogli", "ogli", "gizi",
        "mr", "mrs", "ms", "dr", "student",
        "certificate", "completion", "completed", "course",
        "coursera", "issued", "awarded", "earned", "verify"
    }

    tokens = []
    for t in raw_tokens:
        s = simplify_token(t)
        if not s:
            continue
        if s in stop_words:
            continue
        if len(s) == 1:
            continue
        tokens.append(s)

    return tokens

def token_similarity(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a, b).ratio()

def build_token_matches(excel_tokens: list[str], cert_tokens: list[str], threshold: float = 0.84):
    """
    Excel tokenlari uchun sertifikat tokenlari orasidan eng moslarini topadi.
    Har bir cert token bir marta ishlatiladi.
    """
    matches = []
    used_cert = set()

    for et in excel_tokens:
        best_idx = None
        best_score = 0.0
        best_token = None

        for idx, ct in enumerate(cert_tokens):
            if idx in used_cert:
                continue

            score = token_similarity(et, ct)

            # Prefix yordam: abdulaziz / abdulazizbek kabi holatlar
            if et.startswith(ct) or ct.startswith(et):
                score = max(score, 0.90)

            if score > best_score:
                best_score = score
                best_idx = idx
                best_token = ct

        if best_idx is not None and best_score >= threshold:
            used_cert.add(best_idx)
            matches.append((et, best_token, best_score))

    return matches

def canonical_name_string(tokens: list[str]) -> str:
    return " ".join(sorted(tokens))

def check_name_match(excel_name: str, cert_name: str) -> tuple[str, str]:
    """
    Excel dagi F.I.SH va sertifikatdagi ism familiyani pro darajada solishtiradi.
    """
    excel_tokens = normalize_name(str(excel_name))
    cert_tokens = normalize_name(str(cert_name))

    if not excel_tokens:
        return "TEKSHIRILMADI ⚠️", "Excel ismi bo'sh"
    if not cert_tokens:
        return "TEKSHIRILMADI ⚠️", "Sertifikatda ism topilmadi"

    excel_set = set(excel_tokens)
    cert_set = set(cert_tokens)

    # 1. To'liq set mosligi (tartibdan qat'i nazar)
    if excel_set == cert_set:
        return "MOS ✅", f"To'liq mos: '{cert_name}'"

    # 2. Excel tokenlari sertifikat ichida to'liq mavjud
    if excel_set.issubset(cert_set):
        extra = cert_set - excel_set
        if extra:
            return "MOS ✅", f"Mos (qo'shimcha tokenlar: {', '.join(sorted(extra))}): '{cert_name}'"
        return "MOS ✅", f"To'liq mos: '{cert_name}'"

    # 3. Sertifikat tokenlari Excel ichida subset
    if cert_set.issubset(excel_set):
        extra = excel_set - cert_set
        return "MOS ✅", f"Mos (Excel to'liqroq, ortiqcha: {', '.join(sorted(extra))}): '{cert_name}'"

    # 4. Fuzzy token match
    matches = build_token_matches(excel_tokens, cert_tokens, threshold=0.84)
    matched_excel_tokens = {m[0] for m in matches}
    matched_cert_tokens = {m[1] for m in matches}

    excel_count = len(excel_tokens)
    cert_count = len(cert_tokens)
    match_count = len(matches)

    excel_ratio = match_count / excel_count if excel_count else 0
    cert_ratio = match_count / cert_count if cert_count else 0

    # 5. Juda kuchli fuzzy moslik
    if match_count >= 2 and excel_ratio >= 0.80:
        detail_pairs = ", ".join([f"{a}~{b}" for a, b, _ in matches])
        return "MOS ✅", f"Fuzzy mos: {detail_pairs} | Sertifikat: '{cert_name}'"

    # 6. O'rtacha moslik
    if match_count >= 2 and (excel_ratio >= 0.60 or cert_ratio >= 0.60):
        missing_excel = [t for t in excel_tokens if t not in matched_excel_tokens]
        extra_cert = [t for t in cert_tokens if t not in matched_cert_tokens]
        detail_pairs = ", ".join([f"{a}~{b}" for a, b, _ in matches])
        return (
            "QISMAN MOS ⚠️",
            f"Qisman mos: {detail_pairs} | Yetishmaydi: {', '.join(missing_excel) if missing_excel else '-'} | "
            f"Qo'shimcha: {', '.join(extra_cert) if extra_cert else '-'} | Sertifikat: '{cert_name}'"
        )

    # 7. Umumiy canonical string yaqin bo'lsa
    excel_canon = canonical_name_string(excel_tokens)
    cert_canon = canonical_name_string(cert_tokens)
    whole_ratio = difflib.SequenceMatcher(None, excel_canon, cert_canon).ratio()

    if whole_ratio >= 0.88 and min(excel_count, cert_count) >= 2:
        return "MOS ✅", f"Canonically mos ({whole_ratio:.2f}): '{cert_name}'"

    if whole_ratio >= 0.72 and min(excel_count, cert_count) >= 2:
        return "QISMAN MOS ⚠️", f"Yaqin moslik ({whole_ratio:.2f}) | Excel: '{excel_name}' | Sertifikat: '{cert_name}'"

    return "MOS EMAS ❌", f"Excel: '{excel_name}' | Sertifikat: '{cert_name}'"

# ==========================================
# 6. SAHIFADAN ISM AJRATISH — KUCHAYTIRILGAN
# ==========================================
def clean_candidate_name(candidate: str) -> str:
    if not candidate:
        return ""

    candidate = re.sub(r"\s+", " ", candidate).strip()
    candidate = re.sub(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\b.*$",
        "",
        candidate,
        flags=re.IGNORECASE
    )
    candidate = re.sub(r"\b\d+\s+hours?.*$", "", candidate, flags=re.IGNORECASE)
    candidate = re.sub(r"\b(Coursera|Certificate|Completion|Course|Verify).*$", "", candidate, flags=re.IGNORECASE)
    candidate = re.sub(r"\s+", " ", candidate).strip()

    words = candidate.split()
    if not (2 <= len(words) <= 6):
        return ""

    bad_words = {
        "completion", "certificate", "course", "coursera",
        "google", "meta", "ibm", "scrimba", "openai", "build"
    }
    if any(w.lower() in bad_words for w in words):
        return ""

    return candidate

def extract_name_from_page(html: str) -> str:
    """
    Coursera sertifikat sahifasidan egasining ismini ajratib oladi.
    """
    if not html:
        return ""

    try:
        soup = BeautifulSoup(html, "html.parser")
        full_text = soup.get_text(" ", strip=True)
        full_text = re.sub(r"\s+", " ", full_text)

        direct_patterns = [
            r"Completed by\s+(.+?)(?=\s+(January|February|March|April|May|June|July|August|September|October|November|December)\b)",
            r"Completed by\s+(.+?)(?=\s+\d+\s+hours?)",
            r"Completed by\s+(.+?)(?=\s+Coursera\b)",
            r"Completed by\s+(.+?)(?=\s+certifies\b)",
            r"Completed by\s+(.+?)(?=\s+has\s+successfully\s+completed\b)",
            r"(?:awarded to|issued to|earned by|completed by)\s+(.+?)(?=\s+(January|February|March|April|May|June|July|August|September|October|November|December|\d+\s+hours?|Coursera|Certificate))",
        ]

        for pattern in direct_patterns:
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                candidate = clean_candidate_name(match.group(1))
                if candidate:
                    return candidate

        # title / meta
        og_title = soup.find("meta", property="og:title")
        if og_title and og_title.get("content"):
            title_text = og_title["content"]
            m = re.match(r"^(.+?)(?:'s)?\s+(?:Certificate|Certification|Course)", title_text, re.IGNORECASE)
            if m:
                candidate = clean_candidate_name(m.group(1))
                if candidate:
                    return candidate

        title_tag = soup.find("title")
        if title_tag and title_tag.string:
            title_text = title_tag.string.strip()
            m = re.match(r"^(.+?)(?:'s)?\s+(?:Certificate|Certification|Course)", title_text, re.IGNORECASE)
            if m:
                candidate = clean_candidate_name(m.group(1))
                if candidate:
                    return candidate

        # CSS selector fallback
        for selector in [
            "[data-e2e='certificate-name']",
            ".certificate-name",
            ".cert-name",
            "[class*='certificateName']",
            "[class*='recipient']",
        ]:
            element = soup.select_one(selector)
            if element and element.get_text(strip=True):
                candidate = clean_candidate_name(element.get_text(strip=True))
                if candidate:
                    return candidate

        # generic fallback
        fallback_patterns = [
            r"([A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+(?:\s+[A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+){1,5})\s+has\s+(?:successfully\s+)?completed",
            r"([A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+(?:\s+[A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+){1,5})\s+earned\s+this\s+certificate",
        ]

        for pattern in fallback_patterns:
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                candidate = clean_candidate_name(match.group(1))
                if candidate:
                    return candidate

        return ""
    except Exception:
        return ""

# ==========================================
# 7. NETWORK SESSIYASI
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
# 8. VERIFIKATSIYA MANTIQI
# ==========================================
def verify_link(session, url, timeout):
    if pd.isna(url) or not str(url).startswith("http"):
        return "MAVJUD EMAS", "-", "Havola topilmadi", "", ""

    url = str(url).strip()

    try:
        resp = session.get(url, timeout=timeout, allow_redirects=True)
        final_url = resp.url.lower()
        is_valid_path = any(x in final_url for x in ["/share/", "/verify/", "/accomplishments/"])
        html_content = resp.text if resp.status_code == 200 else ""

        cert_date = extract_certificate_date(html_content)
        cert_name = extract_name_from_page(html_content)

        if resp.status_code == 200 and is_valid_path:
            return "MAVJUD", "200", "Tasdiqlandi ✅", cert_date, cert_name
        elif "login" in final_url or "signup" in final_url:
            return "XATO", "Redirect", "Avtorizatsiya so'raldi (Xato link)", cert_date, cert_name
        else:
            return "MAVJUD EMAS", str(resp.status_code), "Sertifikat sahifasi emas", cert_date, cert_name

    except Exception:
        return "XATO", "Timeout/Error", "Ulanish imkonsiz", "", ""

# ==========================================
# 9. ASOSIY ILOVA
# ==========================================
def main():
    st.title("🎓 Coursera Certificate Verifier Pro")

    with st.sidebar:
        st.markdown("### 🛠 Dastur haqida")
        st.info("Coursera sertifikatlarini avtomatik tekshirish tizimi.")
        st.markdown("---")
        show_sample_download_section()
        st.markdown("---")
        st.markdown("### 📬 Muallifga murojaat")
        st.markdown("""
        <div style="line-height: 2;">
        <img src="https://img.icons8.com/color/20/gmail-new.png"/><a href="mailto:azamat3533141@gmail.com"> azamat3533141@gmail.com</a><br>
        <img src="https://img.icons8.com/color/20/telegram-app.png"/><a href="https://t.me/futurex_azamat"> @futurex_azamat</a>
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
Fayl quyidagi ustunlarga ega bo'lishi lozim:
• №
• Tuman/Shahar
• Maktab raqami
• Sinf
• F.I.SH
• Guvohnoma seriyasi va raqami
• Tug'ilgan sana
• Sertifikat havolasi
• Elektron pochta
⚠️ Ustun nomlari o'zgartirilsa yoki joyi almashsa, tizim noto'g'ri ishlashi mumkin.
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
                    excel_name = item["name"]

                    if cert_code and cert_code in results_cache:
                        status, code, reason, cert_date, cert_name = results_cache[cert_code]
                    elif not cert_code and original_url in fallback_results_cache:
                        status, code, reason, cert_date, cert_name = fallback_results_cache[original_url]
                    else:
                        status, code, reason, cert_date, cert_name = "XATO", "CodeError", "Sertifikat kodi aniqlanmadi", "", ""

                    if status == "MAVJUD" and reason != "TAKRORLANUVCHI 🔄":
                        name_match_status, name_match_detail = check_name_match(excel_name, cert_name)
                    else:
                        name_match_status = "TEKSHIRILMADI ⚠️"
                        name_match_detail = "Sertifikat mavjud emas yoki takrorlanuvchi"

                    if cert_code:
                        if cert_code in seen_codes:
                            display_reason = "TAKRORLANUVCHI 🔄"
                            name_match_status = "TEKSHIRILMADI ⚠️"
                            name_match_detail = "Takrorlanuvchi sertifikat"
                        else:
                            display_reason = reason
                            seen_codes.add(cert_code)
                    else:
                        if original_url in seen_urls_without_code:
                            display_reason = "TAKRORLANUVCHI 🔄"
                            name_match_status = "TEKSHIRILMADI ⚠️"
                            name_match_detail = "Takrorlanuvchi sertifikat"
                        else:
                            display_reason = reason
                            seen_urls_without_code.add(original_url)

                    final_data.append({
                        "F.I.SH": excel_name,
                        "Kurs yo'nalishi": item["course"],
                        "Holati": status,
                        "Natija": display_reason,
                        "Ism Moslik": name_match_status,
                        "Moslik Tafsiloti": name_match_detail,
                        "Sertifikatdagi Ism": cert_name,
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
                name_match_count = int((res_df["Ism Moslik"] == "MOS ✅").sum())
                name_partial_count = int((res_df["Ism Moslik"] == "QISMAN MOS ⚠️").sum())
                name_mismatch_count = int((res_df["Ism Moslik"] == "MOS EMAS ❌").sum())

                st.divider()
                c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
                c1.metric("Jami tekshirildi", len(res_df))
                c2.metric("Tasdiqlandi ✅", confirmed_count)
                c3.metric("Xato/Mavjud emas ❌", error_count)
                c4.metric("Takrorlanuvchi 🔄", duplicate_count)
                c5.metric("Ism mos ✅", name_match_count)
                c6.metric("Qisman mos ⚠️", name_partial_count)
                c7.metric("Ism mos emas ❌", name_mismatch_count)

                st.caption(
                    f"Unikal sertifikat kodlari: {res_df['Sertifikat kodi'].replace('', pd.NA).nunique()} | "
                    f"Takrorlar: {duplicate_count} | "
                    f"Ism mos emas: {name_mismatch_count}"
                )

                st.subheader("📋 Batafsil hisobot")
                display_df = res_df.drop(columns=["__sheet_name__"])

                def row_style(row):
                    styles = [""] * len(row)
                    holat_idx = row.index.get_loc("Holati") if "Holati" in row.index else None
                    moslik_idx = row.index.get_loc("Ism Moslik") if "Ism Moslik" in row.index else None

                    if holat_idx is not None:
                        if row["Holati"] == "MAVJUD":
                            styles[holat_idx] = "background-color: #d4edda"
                        elif row["Holati"] == "XATO":
                            styles[holat_idx] = "background-color: #f8d7da"
                        elif row["Holati"] == "MAVJUD EMAS":
                            styles[holat_idx] = "background-color: #fff3cd"
                        else:
                            styles[holat_idx] = "background-color: #cce5ff"

                    if moslik_idx is not None:
                        if row["Ism Moslik"] == "MOS ✅":
                            styles[moslik_idx] = "background-color: #d4edda"
                        elif row["Ism Moslik"] == "QISMAN MOS ⚠️":
                            styles[moslik_idx] = "background-color: #fff3cd"
                        elif row["Ism Moslik"] == "MOS EMAS ❌":
                            styles[moslik_idx] = "background-color: #f8d7da"
                        else:
                            styles[moslik_idx] = "background-color: #e2e3e5"

                    return styles

                st.dataframe(
                    display_df.style.apply(row_style, axis=1),
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

    st.markdown("""
        <div class="footer">
            Tuzuvchi: Azamat Madrimov | 2026
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
