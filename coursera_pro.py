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
# 1. KONSTANTALAR
# ==========================================
TEMPLATE_FILENAME = "coursera_template.xlsx"

MONTHS_PATTERN = (
    r"January|February|March|April|May|June|"
    r"July|August|September|October|November|December"
)

STOP_WORDS = frozenset({
    "qizi", "qiz", "ogli", "gizi",
    "mr", "mrs", "ms", "dr", "student",
    "certificate", "completion", "completed", "course",
    "coursera", "issued", "awarded", "earned", "verify"
})

BAD_WORDS = frozenset({
    "completion", "certificate", "course", "coursera",
    "google", "meta", "ibm", "scrimba", "openai", "build"
})

SKIP_URL_PARTS = frozenset({
    "account", "accomplishments", "certificates",
    "certificate", "verify", "share"
})

CYRILLIC_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d",
    "е": "e", "ё": "yo", "ж": "j", "з": "z", "и": "i",
    "й": "y", "к": "k", "л": "l", "м": "m", "н": "n",
    "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
    "у": "u", "ф": "f", "х": "x", "ц": "s", "ч": "ch",
    "ш": "sh", "щ": "sh", "ъ": "", "ь": "", "э": "e",
    "ю": "yu", "я": "ya", "ў": "o'", "қ": "q",
    "ғ": "g'", "ҳ": "h", "ы": "i",
}

NAME_REPLACEMENTS = {
    "o g li": "ogli", "o g'li": "ogli",
    "o'g'li": "ogli", "ugli": "ogli",
    "u g li": "ogli", "kizi": "qizi",
}

DATE_PATTERNS = [
    rf"({MONTHS_PATTERN})\s+\d{{1,2}},\s+\d{{4}}",
    rf"\d{{1,2}}\s+({MONTHS_PATTERN})\s+\d{{4}}",
    r"\b(20\d{2}-\d{2}-\d{2})\b",
]

CSS_SELECTORS = [
    "[data-e2e='certificate-name']",
    ".certificate-name",
    ".cert-name",
    "[class*='certificateName']",
    "[class*='recipient']",
]

# ==========================================
# 2. SAHIFA SOZLAMALARI
# ==========================================
st.set_page_config(
    page_title="Coursera Certificate Verifier Pro",
    layout="wide",
    page_icon="🛰️"
)

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Orbitron:wght@400;700;900&display=swap" rel="stylesheet">
    <style>
    /* ── BASE ── */
    html, body, [class*="css"] {
        font-family: 'Share Tech Mono', monospace !important;
        background-color: #060d18 !important;
        color: #e8f4ff !important;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #0a1628 !important;
        border-right: 1px solid #1a3a5c !important;
    }
    section[data-testid="stSidebar"] * { color: #e8f4ff !important; }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        font-family: 'Orbitron', sans-serif !important;
        color: #00d4ff !important;
        letter-spacing: 2px;
    }

    /* Main area */
    .main .block-container {
        background: #060d18 !important;
        padding-top: 24px;
    }

    /* Title */
    h1 {
        font-family: 'Orbitron', sans-serif !important;
        color: #00d4ff !important;
        letter-spacing: 4px !important;
        text-transform: uppercase;
        border-bottom: 1px solid #1a3a5c;
        padding-bottom: 12px;
    }
    h2, h3 {
        font-family: 'Orbitron', sans-serif !important;
        color: #00d4ff !important;
        letter-spacing: 2px !important;
    }

    /* Subheader */
    .stMarkdown p { color: #4a7fa5; font-size: 13px; letter-spacing: 1px; }

    /* Buttons */
    .stButton > button {
        font-family: 'Orbitron', sans-serif !important;
        background: linear-gradient(135deg, #0a1628, #0d2040) !important;
        color: #00d4ff !important;
        border: 1px solid #00d4ff !important;
        letter-spacing: 2px !important;
        border-radius: 2px !important;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        background: #00d4ff !important;
        color: #060d18 !important;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #ff6b2b, #cc4400) !important;
        color: #fff !important;
        border-color: #ff6b2b !important;
        font-size: 15px !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: #ff6b2b !important;
        box-shadow: 0 0 20px rgba(255,107,43,0.5) !important;
    }

    /* Download button */
    .stDownloadButton > button {
        font-family: 'Orbitron', sans-serif !important;
        background: #0a1628 !important;
        color: #39ff6e !important;
        border: 1px solid #39ff6e !important;
        letter-spacing: 2px !important;
        border-radius: 2px !important;
    }
    .stDownloadButton > button:hover {
        background: #39ff6e !important;
        color: #060d18 !important;
    }

    /* Metrics */
    [data-testid="metric-container"] {
        background: #0a1628 !important;
        border: 1px solid #1a3a5c !important;
        border-left: 3px solid #00d4ff !important;
        border-radius: 2px !important;
        padding: 12px !important;
    }
    [data-testid="metric-container"] label {
        font-family: 'Share Tech Mono', monospace !important;
        color: #4a7fa5 !important;
        font-size: 10px !important;
        letter-spacing: 2px !important;
    }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        font-family: 'Orbitron', sans-serif !important;
        color: #e8f4ff !important;
        font-size: 24px !important;
    }

    /* Dataframe */
    .stDataFrame {
        border: 1px solid #1a3a5c !important;
        border-radius: 2px !important;
        background: #0a1628 !important;
    }
    .stDataFrame thead th {
        background: #0d1e35 !important;
        color: #00d4ff !important;
        font-family: 'Orbitron', sans-serif !important;
        font-size: 11px !important;
        letter-spacing: 1px !important;
    }

    /* File uploader */
    [data-testid="stFileUploader"] {
        background: #0a1628 !important;
        border: 1px dashed #1a3a5c !important;
        border-radius: 2px !important;
        padding: 16px !important;
    }
    [data-testid="stFileUploader"] * { color: #4a7fa5 !important; }

    /* Selectbox */
    .stSelectbox > div > div {
        background: #0a1628 !important;
        border: 1px solid #1a3a5c !important;
        color: #e8f4ff !important;
        border-radius: 2px !important;
    }

    /* Sliders */
    .stSlider [data-baseweb="slider"] div[role="slider"] {
        background: #00d4ff !important;
    }
    .stSlider label { color: #4a7fa5 !important; letter-spacing: 1px; }

    /* Divider */
    hr { border-color: #1a3a5c !important; }

    /* Alerts */
    .stSuccess {
        background: rgba(57,255,110,0.08) !important;
        border: 1px solid #39ff6e !important;
        border-radius: 2px !important;
        color: #39ff6e !important;
    }
    .stWarning {
        background: rgba(255,215,0,0.08) !important;
        border: 1px solid #ffd700 !important;
        border-radius: 2px !important;
        color: #ffd700 !important;
    }
    .stError {
        background: rgba(255,59,59,0.08) !important;
        border: 1px solid #ff3b3b !important;
        border-radius: 2px !important;
        color: #ff3b3b !important;
    }

    /* Progress bar */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #00d4ff, #39ff6e) !important;
    }

    /* Footer */
    .footer {
        position: fixed; left: 0; bottom: 0; width: 100%;
        background: #0a1628;
        border-top: 1px solid #1a3a5c;
        color: #4a7fa5;
        text-align: center; padding: 10px;
        font-family: 'Share Tech Mono', monospace;
        font-size: 12px;
        letter-spacing: 2px;
        z-index: 1000;
    }

    /* Sample box */
    .sample-box {
        padding: 14px; border-radius: 2px;
        background: rgba(0,212,255,0.05);
        border: 1px solid rgba(0,212,255,0.25);
        border-left: 3px solid #00d4ff;
        margin-bottom: 12px;
        font-size: 12px;
        letter-spacing: 1px;
        color: #4a7fa5 !important;
    }

    /* Scanlines overlay */
    .main::before {
        content: '';
        position: fixed; top: 0; left: 0;
        width: 100%; height: 100%;
        background: repeating-linear-gradient(
            0deg, transparent, transparent 2px,
            rgba(0,212,255,0.012) 2px, rgba(0,212,255,0.012) 4px
        );
        pointer-events: none;
        z-index: 9999;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. YORDAMCHI FUNKSIYALAR
# ==========================================
def resolve_path(path: str) -> str:
    base_dir = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_dir, path)


def get_template_bytes() -> bytes | None:
    try:
        path = resolve_path(TEMPLATE_FILENAME)
        if os.path.exists(path):
            with open(path, "rb") as f:
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
# 4. SERTIFIKAT KODI AJRATISH
# ==========================================
def extract_certificate_code(url) -> str:
    if pd.isna(url):
        return ""
    url = str(url).strip()
    if not url.startswith("http"):
        return ""
    try:
        parts = urlparse(url).path.strip("/").lower().split("/")

        if len(parts) >= 2 and parts[0] == "share":
            return parts[1].strip().lower()

        if "verify" in parts:
            idx = parts.index("verify")
            if idx + 1 < len(parts):
                return parts[idx + 1].strip().lower()

        if "accomplishments" in parts:
            for part in reversed(parts):
                part = part.strip().lower()
                if part and part not in SKIP_URL_PARTS:
                    return part

        m = re.search(r"(?:share|verify)/([^/?#]+)", url, re.IGNORECASE)
        if m:
            return m.group(1).strip().lower()

    except Exception:
        pass
    return ""

# ==========================================
# 5. SERTIFIKAT SANASI AJRATISH
# ==========================================
def extract_certificate_date(html: str) -> str:
    if not html:
        return ""
    try:
        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text(" ", strip=True)

        for pattern in DATE_PATTERNS:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                return m.group(0)

        for script in soup.find_all("script"):
            script_text = script.get_text(" ", strip=True)
            for pattern in DATE_PATTERNS:
                m = re.search(pattern, script_text, re.IGNORECASE)
                if m:
                    return m.group(0)

    except Exception:
        pass
    return ""

# ==========================================
# 6. ISM NORMALIZATSIYASI VA MOSLIK
# ==========================================
def transliterate_cyrillic_to_latin(text: str) -> str:
    return "".join(CYRILLIC_MAP.get(ch, ch) for ch in text)


def standardize_apostrophes(text: str) -> str:
    for ch in ("'", "'", "`", "ʻ", "ʼ", "ʹ", "´"):
        text = text.replace(ch, "'")
    return text


def simplify_token(token: str) -> str:
    token = standardize_apostrophes(token.lower().strip())
    for old, new in (
        ("'", ""), ("yo", "io"), ("yu", "iu"),
        ("ya", "ia"), ("ts", "s"), ("iy", "i"), ("yy", "y"),
    ):
        token = token.replace(old, new)
    return re.sub(r"[^a-z0-9-]", "", token)


def normalize_name(name: str) -> list[str]:
    if not name or not isinstance(name, str):
        return []

    name = str(name).strip().lower()
    name = standardize_apostrophes(name)
    name = transliterate_cyrillic_to_latin(name)

    for old, new in NAME_REPLACEMENTS.items():
        name = name.replace(old, new)

    name = re.sub(r"[^a-z0-9'\s-]", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()

    tokens = []
    for t in name.split():
        s = simplify_token(t.strip())
        if s and s not in STOP_WORDS and len(s) > 1:
            tokens.append(s)
    return tokens


def token_similarity(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a, b).ratio()


def build_token_matches(
    excel_tokens: list[str],
    cert_tokens: list[str],
    threshold: float = 0.84
) -> list[tuple[str, str, float]]:
    matches = []
    used_cert: set[int] = set()

    for et in excel_tokens:
        best_idx, best_score, best_token = None, 0.0, None

        for idx, ct in enumerate(cert_tokens):
            if idx in used_cert:
                continue
            score = token_similarity(et, ct)
            if et.startswith(ct) or ct.startswith(et):
                score = max(score, 0.90)
            if score > best_score:
                best_score, best_idx, best_token = score, idx, ct

        if best_idx is not None and best_score >= threshold:
            used_cert.add(best_idx)
            matches.append((et, best_token, best_score))

    return matches


def check_name_match(excel_name: str, cert_name: str) -> tuple[str, str]:
    excel_tokens = normalize_name(str(excel_name))
    cert_tokens = normalize_name(str(cert_name))

    if not excel_tokens:
        return "TEKSHIRILMADI ⚠️", "Excel ismi bo'sh"
    if not cert_tokens:
        return "TEKSHIRILMADI ⚠️", "Sertifikatda ism topilmadi"

    excel_set = set(excel_tokens)
    cert_set = set(cert_tokens)

    if excel_set == cert_set:
        return "MOS ✅", f"To'liq mos: '{cert_name}'"

    if excel_set.issubset(cert_set):
        extra = cert_set - excel_set
        label = f"Mos (qo'shimcha tokenlar: {', '.join(sorted(extra))}): " if extra else ""
        return "MOS ✅", f"{label}'{cert_name}'"

    if cert_set.issubset(excel_set):
        extra = excel_set - cert_set
        return "MOS ✅", f"Mos (Excel to'liqroq, ortiqcha: {', '.join(sorted(extra))}): '{cert_name}'"

    matches = build_token_matches(excel_tokens, cert_tokens, threshold=0.84)
    matched_excel = {m[0] for m in matches}
    matched_cert = {m[1] for m in matches}
    n_e, n_c, n_m = len(excel_tokens), len(cert_tokens), len(matches)
    e_ratio = n_m / n_e if n_e else 0
    c_ratio = n_m / n_c if n_c else 0
    pairs_str = ", ".join(f"{a}~{b}" for a, b, _ in matches)

    if n_m >= 2 and e_ratio >= 0.80:
        return "MOS ✅", f"Fuzzy mos: {pairs_str} | Sertifikat: '{cert_name}'"

    if n_m >= 2 and (e_ratio >= 0.60 or c_ratio >= 0.60):
        missing = [t for t in excel_tokens if t not in matched_excel]
        extra = [t for t in cert_tokens if t not in matched_cert]
        return (
            "QISMAN MOS ⚠️",
            f"Qisman mos: {pairs_str} | Yetishmaydi: {', '.join(missing) or '-'} | "
            f"Qo'shimcha: {', '.join(extra) or '-'} | Sertifikat: '{cert_name}'"
        )

    excel_canon = " ".join(sorted(excel_tokens))
    cert_canon = " ".join(sorted(cert_tokens))
    ratio = difflib.SequenceMatcher(None, excel_canon, cert_canon).ratio()
    min_count = min(n_e, n_c)

    if ratio >= 0.88 and min_count >= 2:
        return "MOS ✅", f"Canonically mos ({ratio:.2f}): '{cert_name}'"
    if ratio >= 0.72 and min_count >= 2:
        return "QISMAN MOS ⚠️", f"Yaqin moslik ({ratio:.2f}) | Excel: '{excel_name}' | Sertifikat: '{cert_name}'"

    return "MOS EMAS ❌", f"Excel: '{excel_name}' | Sertifikat: '{cert_name}'"

# ==========================================
# 7. SAHIFADAN ISM AJRATISH
# ==========================================
_MONTH_RE = re.compile(MONTHS_PATTERN, re.IGNORECASE)
_HOURS_RE = re.compile(r"\b\d+\s+hours?.*$", re.IGNORECASE)
_CERT_RE = re.compile(r"\b(Coursera|Certificate|Completion|Course|Verify).*$", re.IGNORECASE)

DIRECT_NAME_PATTERNS = [
    rf"Completed by\s+(.+?)(?=\s+(?:{MONTHS_PATTERN})\b)",
    r"Completed by\s+(.+?)(?=\s+\d+\s+hours?)",
    r"Completed by\s+(.+?)(?=\s+Coursera\b)",
    r"Completed by\s+(.+?)(?=\s+certifies\b)",
    r"Completed by\s+(.+?)(?=\s+has\s+successfully\s+completed\b)",
    rf"(?:awarded to|issued to|earned by|completed by)\s+(.+?)(?=\s+(?:{MONTHS_PATTERN}|\d+\s+hours?|Coursera|Certificate))",
]

FALLBACK_NAME_PATTERNS = [
    r"([A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+(?:\s+[A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+){1,5})\s+has\s+(?:successfully\s+)?completed",
    r"([A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+(?:\s+[A-ZÀ-ÿА-ЯЁЎҚҒҲ][A-Za-zÀ-ÿА-Яа-яЁёЎўҚқҒғҲҳ'`\-]+){1,5})\s+earned\s+this\s+certificate",
]


def clean_candidate_name(candidate: str) -> str:
    if not candidate:
        return ""
    candidate = re.sub(r"\s+", " ", candidate).strip()
    candidate = _MONTH_RE.sub("", candidate)
    candidate = _HOURS_RE.sub("", candidate)
    candidate = _CERT_RE.sub("", candidate)
    candidate = re.sub(r"\s+", " ", candidate).strip()

    words = candidate.split()
    if not (2 <= len(words) <= 6):
        return ""
    if any(w.lower() in BAD_WORDS for w in words):
        return ""
    return candidate


def extract_name_from_page(html: str) -> str:
    if not html:
        return ""
    try:
        soup = BeautifulSoup(html, "html.parser")
        full_text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))

        for pattern in DIRECT_NAME_PATTERNS:
            m = re.search(pattern, full_text, re.IGNORECASE)
            if m:
                name = clean_candidate_name(m.group(1))
                if name:
                    return name

        for tag, attr in (
            (soup.find("meta", property="og:title"), "content"),
            (soup.find("title"), None),
        ):
            if tag:
                text = tag.get(attr) if attr else getattr(tag, "string", None)
                if text:
                    m = re.match(
                        r"^(.+?)(?:'s)?\s+(?:Certificate|Certification|Course)",
                        str(text).strip(), re.IGNORECASE
                    )
                    if m:
                        name = clean_candidate_name(m.group(1))
                        if name:
                            return name

        for selector in CSS_SELECTORS:
            el = soup.select_one(selector)
            if el:
                name = clean_candidate_name(el.get_text(strip=True))
                if name:
                    return name

        for pattern in FALLBACK_NAME_PATTERNS:
            m = re.search(pattern, full_text, re.IGNORECASE)
            if m:
                name = clean_candidate_name(m.group(1))
                if name:
                    return name

    except Exception:
        pass
    return ""

# ==========================================
# 8. NETWORK SESSIYASI
# ==========================================
@st.cache_resource
def get_pro_session() -> requests.Session:
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
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        )
    })
    return session

# ==========================================
# 9. VERIFIKATSIYA
# ==========================================
_VALID_PATHS = frozenset({"/share/", "/verify/", "/accomplishments/"})


def verify_link(session: requests.Session, url, timeout: int) -> tuple:
    if pd.isna(url) or not str(url).startswith("http"):
        return "MAVJUD EMAS", "-", "Havola topilmadi", "", ""

    url = str(url).strip()
    try:
        resp = session.get(url, timeout=timeout, allow_redirects=True)
        final_url = resp.url.lower()
        is_valid = any(p in final_url for p in _VALID_PATHS)
        html = resp.text if resp.status_code == 200 else ""

        cert_date = extract_certificate_date(html)
        cert_name = extract_name_from_page(html)

        if resp.status_code == 200 and is_valid:
            return "MAVJUD", "200", "Tasdiqlandi ✅", cert_date, cert_name
        if "login" in final_url or "signup" in final_url:
            return "XATO", "Redirect", "Avtorizatsiya so'raldi (Xato link)", cert_date, cert_name
        return "MAVJUD EMAS", str(resp.status_code), "Sertifikat sahifasi emas", cert_date, cert_name

    except Exception:
        return "XATO", "Timeout/Error", "Ulanish imkonsiz", "", ""

# ==========================================
# 10. FAYLNI O'QISH VA TAYYORLASH
# ==========================================
def load_sheets(file) -> tuple[dict, str]:
    """Fayl (xlsx yoki csv) dan listlarni o'qib qaytaradi."""
    if file.name.endswith(".csv"):
        return {"CSV": pd.read_csv(file, skiprows=2)}, "CSV"

    all_sheets = pd.read_excel(file, sheet_name=None, skiprows=2)
    names = list(all_sheets.keys())
    selected = st.selectbox("Tekshirish uchun listni tanlang", names, index=0)
    return {selected: all_sheets[selected]}, selected


def prepare_sheets(uploaded_sheets: dict) -> tuple[list[dict], int]:
    """Har bir listdan fish_col va course_cols ni aniqlab chiqadi."""
    prepared, total = [], 0
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

        prepared.append({
            "sheet_name": str(sheet_name)[:31],
            "df": df,
            "fish_col": fish_col,
            "course_cols": course_cols,
        })
        total += len(df)
    return prepared, total


def collect_entries(prepared_sheets: list[dict]) -> tuple[list[dict], dict, dict]:
    """Barcha sertifikat yozuvlarini to'plab, unikal kodlarni ajratadi."""
    all_entries: list[dict] = []
    unique_code_to_url: dict[str, str] = {}
    unique_fallback_to_url: dict[str, str] = {}

    for info in prepared_sheets:
        for _, row in info["df"].iterrows():
            for col in info["course_cols"]:
                raw = row[col]
                url = str(raw).strip()
                if pd.isna(raw) or "http" not in url:
                    continue

                code = extract_certificate_code(url)
                all_entries.append({
                    "sheet_name": info["sheet_name"],
                    "name": row[info["fish_col"]],
                    "course": col,
                    "url": url,
                    "cert_code": code,
                })

                if code:
                    unique_code_to_url.setdefault(code, url)
                else:
                    unique_fallback_to_url.setdefault(url, url)

    return all_entries, unique_code_to_url, unique_fallback_to_url


def run_verification(
    session: requests.Session,
    unique_code_to_url: dict,
    unique_fallback_to_url: dict,
    threads: int,
    timeout: int,
) -> tuple[dict, dict]:
    """Parallel tekshiruvni amalga oshirib natijalar lug'atini qaytaradi."""
    code_results: dict[str, tuple] = {}
    url_results: dict[str, tuple] = {}

    code_items = list(unique_code_to_url.items())
    url_items = list(unique_fallback_to_url.items())
    total = len(code_items) + len(url_items)

    progress = st.progress(0)
    status_box = st.empty()

    with ThreadPoolExecutor(max_workers=threads) as executor:
        future_map: dict = {}
        for code, url in code_items:
            future_map[executor.submit(verify_link, session, url, timeout)] = ("code", code)
        for key, url in url_items:
            future_map[executor.submit(verify_link, session, url, timeout)] = ("url", key)

        for i, future in enumerate(as_completed(future_map), 1):
            kind, key = future_map[future]
            result = future.result()
            (code_results if kind == "code" else url_results)[key] = result
            progress.progress(i / total)
            status_box.text(f"Tekshirilmoqda: {i}/{total}")

    return code_results, url_results


def build_final_data(
    all_entries: list[dict],
    code_results: dict,
    url_results: dict,
) -> list[dict]:
    """Har bir yozuv uchun natija va ism-moslikni yig'adi."""
    final_data = []
    seen_codes: set[str] = set()
    seen_urls: set[str] = set()

    for item in all_entries:
        code = item["cert_code"]
        url = item["url"]
        excel_name = item["name"]

        if code and code in code_results:
            status, http_code, reason, cert_date, cert_name = code_results[code]
        elif not code and url in url_results:
            status, http_code, reason, cert_date, cert_name = url_results[url]
        else:
            status, http_code, reason, cert_date, cert_name = (
                "XATO", "CodeError", "Sertifikat kodi aniqlanmadi", "", ""
            )

        # Takrorlanishni tekshirish
        is_dup = (code and code in seen_codes) or (not code and url in seen_urls)

        if is_dup:
            display_reason = "TAKRORLANUVCHI 🔄"
            name_status = "TEKSHIRILMADI ⚠️"
            name_detail = "Takrorlanuvchi sertifikat"
        else:
            display_reason = reason
            if code:
                seen_codes.add(code)
            else:
                seen_urls.add(url)

            if status == "MAVJUD":
                name_status, name_detail = check_name_match(excel_name, cert_name)
            else:
                name_status = "TEKSHIRILMADI ⚠️"
                name_detail = "Sertifikat mavjud emas"

        final_data.append({
            "F.I.SH": excel_name,
            "Kurs yo'nalishi": item["course"],
            "Holati": status,
            "Natija": display_reason,
            "Ism Moslik": name_status,
            "Moslik Tafsiloti": name_detail,
            "Sertifikatdagi Ism": cert_name,
            "Havola": url,
            "Sertifikat kodi": code,
            "Sertifikat olingan sana": cert_date,
            "__sheet_name__": item["sheet_name"],
        })

    return final_data

# ==========================================
# 11. NATIJALARNI KO'RSATISH
# ==========================================
def show_metrics(df: pd.DataFrame):
    dup = int((df["Natija"] == "TAKRORLANUVCHI 🔄").sum())
    confirmed = int(((df["Holati"] == "MAVJUD") & (df["Natija"] != "TAKRORLANUVCHI 🔄")).sum())
    errors = int((df["Holati"] != "MAVJUD").sum())
    mos = int((df["Ism Moslik"] == "MOS ✅").sum())
    partial = int((df["Ism Moslik"] == "QISMAN MOS ⚠️").sum())
    mismatch = int((df["Ism Moslik"] == "MOS EMAS ❌").sum())

    st.divider()
    cols = st.columns(7)
    for col, label, val in zip(cols, [
        "Jami tekshirildi", "Tasdiqlandi ✅", "Xato/Mavjud emas ❌",
        "Takrorlanuvchi 🔄", "Ism mos ✅", "Qisman mos ⚠️", "Ism mos emas ❌"
    ], [len(df), confirmed, errors, dup, mos, partial, mismatch]):
        col.metric(label, val)

    st.caption(
        f"Unikal sertifikat kodlari: {df['Sertifikat kodi'].replace('', pd.NA).nunique()} | "
        f"Takrorlar: {dup} | Ism mos emas: {mismatch}"
    )


def row_style(row):
    n = len(row)
    styles = [""] * n
    idx_map = {col: row.index.get_loc(col) for col in ("Holati", "Ism Moslik") if col in row.index}

    STATUS_COLORS = {
        "MAVJUD": "#d4edda", "XATO": "#f8d7da",
        "MAVJUD EMAS": "#fff3cd",
    }
    MATCH_COLORS = {
        "MOS ✅": "#d4edda", "QISMAN MOS ⚠️": "#fff3cd",
        "MOS EMAS ❌": "#f8d7da",
    }

    if "Holati" in idx_map:
        styles[idx_map["Holati"]] = (
            f"background-color: {STATUS_COLORS.get(row['Holati'], '#cce5ff')}"
        )
    if "Ism Moslik" in idx_map:
        styles[idx_map["Ism Moslik"]] = (
            f"background-color: {MATCH_COLORS.get(row['Ism Moslik'], '#e2e3e5')}"
        )
    return styles


def export_excel(res_df: pd.DataFrame, base_name: str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name in res_df["__sheet_name__"].dropna().unique():
            sheet_df = res_df[res_df["__sheet_name__"] == sheet_name].drop(columns=["__sheet_name__"])
            if not sheet_df.empty:
                sheet_df.to_excel(writer, index=False, sheet_name=str(sheet_name)[:31])

    st.download_button(
        label="📥 Excelni yuklab olish",
        data=output.getvalue(),
        file_name=f"{base_name}_Verify.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ==========================================
# 12. ASOSIY ILOVA
# ==========================================
def main():
    st.title("🎓 Coursera Certificate Verifier Pro")

    with st.sidebar:
        show_sample_download_section()
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

    st.subheader("Maktab o'quvchilari sertifikatlarini avtomatik tekshirish yordamchi tizimi")
    file = st.file_uploader(
        "Excel (.xlsx) yoki CSV faylni yuklang",
        type=["xlsx", "csv"],
        help=(
            "Fayl quyidagi ustunlarga ega bo'lishi lozim:\n"
            "• №\n• Tuman/Shahar\n• Maktab raqami\n• Sinf\n• F.I.SH\n"
            "• Guvohnoma seriyasi va raqami\n• Tug'ilgan sana\n"
            "• Sertifikat havolasi\n• Elektron pochta\n\n"
            "⚠️ Ustun nomlari o'zgartirilsa yoki joyi almashsa, "
            "tizim noto'g'ri ishlashi mumkin."
        )
    )

    if not file:
        st.markdown(
            '<div class="footer">Tuzuvchi: Azamat Madrimov | 2026</div>',
            unsafe_allow_html=True
        )
        return

    try:
        base_name = os.path.splitext(file.name)[0]
        uploaded_sheets, selected_sheet = load_sheets(file)
        prepared_sheets, total_students = prepare_sheets(uploaded_sheets)

        st.success(
            f"Ma'lumotlar yuklandi. Tanlangan list: **{selected_sheet}**. "
            f"Jami **{total_students}** ta o'quvchi aniqlandi."
        )

        if not st.button("🚀 TEKSHIRISHNI BOSHLASH", type="primary", use_container_width=True):
            st.markdown(
                '<div class="footer">Tuzuvchi: Azamat Madrimov | 2026</div>',
                unsafe_allow_html=True
            )
            return

        all_entries, code_urls, fallback_urls = collect_entries(prepared_sheets)

        if not all_entries:
            st.warning("Tekshirish uchun hech qanday sertifikat link topilmadi.")
            return

        session = get_pro_session()
        code_results, url_results = run_verification(
            session, code_urls, fallback_urls, threads, timeout
        )

        final_data = build_final_data(all_entries, code_results, url_results)
        res_df = pd.DataFrame(final_data)

        show_metrics(res_df)
        st.subheader("📋 Batafsil hisobot")
        display_df = res_df.drop(columns=["__sheet_name__"])
        st.dataframe(display_df.style.apply(row_style, axis=1), use_container_width=True)
        export_excel(res_df, base_name)

    except Exception as e:
        st.error(f"Xatolik: {e}")

    st.markdown(
        '<div class="footer">Tuzuvchi: Azamat Madrimov | 2026</div>',
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
