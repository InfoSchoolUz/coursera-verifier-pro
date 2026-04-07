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
    sessio
