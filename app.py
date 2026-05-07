import base64
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urlencode

import pandas as pd
import requests
import streamlit as st
from pypdf import PdfReader, PdfWriter


# -----------------------------------------------------------------------------
# Phoenix Contact PDF API settings
# -----------------------------------------------------------------------------
BASE_PDF_API_URL = "https://www.phoenixcontact.com/product/pdf/api/v1/{encoded_code}"
DEFAULT_REALM = "pc"
DEFAULT_LOCALE = "en-PC"
DEFAULT_TIMEOUT = (20, 60)
PDF_DOWNLOAD_RETRIES = 2

PDF_BLOCK_OPTIONS: List[Tuple[str, str, str]] = [
    ("commercial-data", "Commercial data", "Basic commercial and ordering information."),
    ("technical-data", "Technical data", "Electrical, mechanical, and product specifications."),
    ("drawings", "Drawings", "Dimensional drawings and product graphics."),
    ("classifications", "Classifications", "ETIM, eCl@ss, UNSPSC, and other classifications."),
    (
        "environmental-compliance-data",
        "Environmental compliance data",
        "RoHS, REACH, China RoHS, and related compliance information.",
    ),
    ("all-accessories", "Accessories", "Compatible accessories listed in the PDF."),
]

APP_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
DEFAULT_COVER_PATHS = [
    APP_DIR / "cover.pdf",
    APP_DIR / "cover" / "cover.pdf",
]


# -----------------------------------------------------------------------------
# Data preparation helpers
# -----------------------------------------------------------------------------
def clean_phoenix_code(value: str) -> str:
    """Return the numeric Phoenix Contact item number.

    Supported examples:
    - PHX-3201853 -> 3201853
    - 3201853 -> 3201853
    - PHX - 3201853 -> 3201853
    """
    value = str(value or "").strip()
    if not value:
        return ""

    value = value.replace("\u2013", "-").replace("\u2014", "-")
    value = re.sub(r"\s*-\s*", "-", value)

    if "-" in value:
        value = value.rsplit("-", 1)[1].strip()

    # Phoenix Contact order numbers used in this workflow are numeric.
    value = re.sub(r"[^0-9]", "", value)
    return value


def normalize_codes(raw_codes: Iterable[str]) -> List[str]:
    """Split, clean, and de-duplicate codes while preserving input order."""
    codes: List[str] = []

    for item in raw_codes:
        if item is None:
            continue

        item_str = str(item).strip()
        if not item_str:
            continue

        parts = re.split(r"[\s,;]+", item_str)
        for part in parts:
            code = clean_phoenix_code(part)
            if code:
                codes.append(code)

    seen = set()
    unique_codes: List[str] = []
    for code in codes:
        if code not in seen:
            seen.add(code)
            unique_codes.append(code)

    return unique_codes


def encode_item_number_for_phoenix(item_number: str) -> str:
    """Encode the Phoenix Contact numeric item number for the PDF API path.

    Example:
        3201853 -> MzIwMTg1Mw
    """
    item_number = clean_phoenix_code(item_number)
    if not item_number:
        raise ValueError("Phoenix Contact item number is empty.")

    encoded = base64.b64encode(item_number.encode("ascii")).decode("ascii")
    return encoded.rstrip("=")


def build_phoenix_pdf_url(
    item_number: str,
    selected_blocks: Sequence[str],
    realm: str = DEFAULT_REALM,
    locale: str = DEFAULT_LOCALE,
) -> str:
    """Build the Phoenix Contact PDF API download URL."""
    if not selected_blocks:
        raise ValueError("At least one PDF content block must be selected.")

    encoded_code = encode_item_number_for_phoenix(item_number)
    query = urlencode(
        [
            ("_realm", realm),
            ("_locale", locale),
            ("blocks", ",".join(selected_blocks)),
            ("action", "Downlaod"),
        ]
    )
    return f"{BASE_PDF_API_URL.format(encoded_code=encoded_code)}?{query}"


def ensure_pdf_filename(filename: str) -> str:
    filename = str(filename or "phoenix_contact_datasheet_pack.pdf").strip()
    filename = re.sub(r"[\\/:*?\"<>|]+", "-", filename)
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    return filename


def read_excel_file(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file)


def pick_default_excel_column(columns: List[str]) -> int:
    preferred_names = [
        "Item No.1",
        "Item No.",
        "Item No",
        "Order No.",
        "Order No",
        "Material",
        "Material Number",
        "Product Number",
        "Part Number",
        "Code",
    ]

    lower_to_index = {str(col).strip().lower(): idx for idx, col in enumerate(columns)}
    for name in preferred_names:
        idx = lower_to_index.get(name.lower())
        if idx is not None:
            return idx
    return 0


def extract_codes_from_selected_column(df: pd.DataFrame, selected_column: str) -> List[str]:
    if selected_column not in df.columns:
        return []

    values = df[selected_column].dropna().astype(str).tolist()
    return normalize_codes(values)


# -----------------------------------------------------------------------------
# PDF helpers
# -----------------------------------------------------------------------------
def looks_like_pdf(response: requests.Response) -> bool:
    content_type = (response.headers.get("Content-Type") or "").lower()
    content_disp = (response.headers.get("Content-Disposition") or "").lower()

    if "pdf" in content_type or ".pdf" in content_disp:
        return True

    return response.content[:5] == b"%PDF-"


def is_valid_pdf_bytes(pdf_bytes: bytes) -> bool:
    try:
        reader = PdfReader(BytesIO(pdf_bytes), strict=False)
        _ = len(reader.pages)
        return True
    except Exception:
        return False


def read_default_cover_pdf_bytes() -> Optional[bytes]:
    for cover_path in DEFAULT_COVER_PATHS:
        if cover_path.is_file():
            cover_pdf_bytes = cover_path.read_bytes()
            if is_valid_pdf_bytes(cover_pdf_bytes):
                return cover_pdf_bytes
    return None


def get_cover_pdf_bytes(uploaded_cover, include_cover: bool) -> Tuple[Optional[bytes], Optional[str], Optional[str]]:
    """Return cover bytes, error message, and warning message."""
    if not include_cover:
        return None, None, None

    if uploaded_cover is not None:
        cover_pdf_bytes = uploaded_cover.getvalue()
        if is_valid_pdf_bytes(cover_pdf_bytes):
            return cover_pdf_bytes, None, None
        return None, "The uploaded cover file is not a valid PDF.", None

    cover_pdf_bytes = read_default_cover_pdf_bytes()
    if cover_pdf_bytes is not None:
        return cover_pdf_bytes, None, None

    return (
        None,
        None,
        "No default cover.pdf was found. The pack will be created without a cover page.",
    )


def get_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (compatible; PhoenixContactPdfPackBuilder/1.0)",
            "Accept": "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8",
        }
    )
    return session


def download_pdf_bytes_for_code(
    session: requests.Session,
    code: str,
    selected_blocks: Sequence[str],
    realm: str,
    locale: str,
) -> Tuple[bool, Optional[bytes], str, str]:
    """Download one Phoenix Contact PDF.

    Returns:
        ok, pdf_bytes, used_url, error_message
    """
    try:
        url = build_phoenix_pdf_url(code, selected_blocks, realm=realm, locale=locale)
    except Exception as exc:
        return False, None, "", str(exc)

    last_error = "PDF was not downloaded."
    for _ in range(PDF_DOWNLOAD_RETRIES):
        try:
            response = session.get(url, timeout=DEFAULT_TIMEOUT, allow_redirects=True)

            if response.status_code != 200:
                last_error = f"HTTP {response.status_code}"
                continue

            if not looks_like_pdf(response):
                last_error = "The response did not look like a PDF."
                continue

            pdf_bytes = response.content
            if not is_valid_pdf_bytes(pdf_bytes):
                last_error = "The downloaded file was not a valid PDF."
                continue

            return True, pdf_bytes, url, ""

        except requests.RequestException as exc:
            last_error = str(exc)

    return False, None, url, last_error


def merge_pdf_bytes(pdf_byte_list: List[bytes], cover_pdf_bytes: Optional[bytes] = None) -> bytes:
    writer = PdfWriter()

    if cover_pdf_bytes:
        cover_reader = PdfReader(BytesIO(cover_pdf_bytes), strict=False)
        for page in cover_reader.pages:
            writer.add_page(page)

    for pdf_bytes in pdf_byte_list:
        reader = PdfReader(BytesIO(pdf_bytes), strict=False)
        for page in reader.pages:
            writer.add_page(page)

    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output.getvalue()


def process_code(
    index: int,
    code: str,
    selected_blocks: Sequence[str],
    realm: str,
    locale: str,
) -> Dict[str, object]:
    session = get_session()
    ok, pdf_bytes, used_url, error_message = download_pdf_bytes_for_code(
        session=session,
        code=code,
        selected_blocks=selected_blocks,
        realm=realm,
        locale=locale,
    )

    return {
        "index": index,
        "code": code,
        "encoded": encode_item_number_for_phoenix(code),
        "ok": ok,
        "pdf_bytes": pdf_bytes,
        "used_url": used_url,
        "error": error_message,
    }


def download_pdfs_parallel(
    codes: List[str],
    selected_blocks: Sequence[str],
    realm: str = DEFAULT_REALM,
    locale: str = DEFAULT_LOCALE,
    max_workers: int = 8,
):
    downloaded_pdfs: List[bytes] = []
    success_rows: List[Dict[str, object]] = []
    failed_rows: List[Dict[str, object]] = []
    results: List[Dict[str, object]] = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_code = {
            executor.submit(process_code, index, code, selected_blocks, realm, locale): code
            for index, code in enumerate(codes)
        }

        completed = 0
        progress_bar = st.progress(0)
        status_text = st.empty()

        for future in as_completed(future_to_code):
            code = future_to_code[future]
            completed += 1

            try:
                result = future.result()
            except Exception as exc:
                result = {
                    "index": codes.index(code),
                    "code": code,
                    "encoded": "",
                    "ok": False,
                    "pdf_bytes": None,
                    "used_url": "",
                    "error": str(exc),
                }

            results.append(result)
            status_text.info(f"Processed {completed} of {len(codes)} - {code}")
            progress_bar.progress(completed / len(codes))

        status_text.empty()

    results.sort(key=lambda x: int(x["index"]))

    for result in results:
        if result["ok"] and result["pdf_bytes"]:
            downloaded_pdfs.append(result["pdf_bytes"])
            success_rows.append(
                {
                    "Input code": f"PHX-{result['code']}",
                    "Item number": result["code"],
                    "Encoded API code": result["encoded"],
                    "Status": "Downloaded",
                    "Source URL": result["used_url"],
                }
            )
        else:
            failed_rows.append(
                {
                    "Input code": f"PHX-{result['code']}",
                    "Item number": result["code"],
                    "Encoded API code": result.get("encoded", ""),
                    "Status": "Failed",
                    "Error": result.get("error", "Unknown error"),
                    "Source URL": result.get("used_url", ""),
                }
            )

    return downloaded_pdfs, success_rows, failed_rows, results


# -----------------------------------------------------------------------------
# UI helpers
# -----------------------------------------------------------------------------
def render_step(number: str, title: str, text: str) -> None:
    st.markdown(
        f"""
        <div class="process-card">
            <div class="process-number">{number}</div>
            <div>
                <div class="process-title">{title}</div>
                <div class="process-text">{text}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_cards(submitted_count: int, downloaded_count: int, failed_count: int) -> None:
    st.markdown(
        f"""
        <div class="metric-grid">
            <div class="metric-card">
                <div class="metric-label">Submitted</div>
                <div class="metric-value">{submitted_count}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Downloaded</div>
                <div class="metric-value">{downloaded_count}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Failed</div>
                <div class="metric-value">{failed_count}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Page config
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Phoenix Contact Datasheet Pack Builder",
    page_icon="P",
    layout="wide",
)


# -----------------------------------------------------------------------------
# CSS - Phoenix Contact inspired, no official brand assets copied
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
        :root {
            --phx-yellow: #ffd200;
            --phx-yellow-soft: #fff4b8;
            --phx-black: #151515;
            --phx-ink: #232323;
            --phx-muted: #6b6b6b;
            --phx-line: #dedede;
            --phx-panel: #ffffff;
            --phx-warm: #f2f1ed;
            --phx-silver: #f6f6f4;
            --phx-shadow: rgba(10, 10, 10, 0.08);
        }

        #MainMenu,
        footer,
        header[data-testid="stHeader"] {
            visibility: hidden;
            height: 0;
        }

        .stApp {
            background:
                radial-gradient(circle at top right, rgba(255, 210, 0, 0.24), transparent 24rem),
                linear-gradient(180deg, #ffffff 0%, var(--phx-warm) 58%, #fafaf8 100%);
            color: var(--phx-ink);
            font-family: Arial, Helvetica, sans-serif;
        }

        .block-container {
            padding-top: 1.15rem;
            padding-bottom: 2.5rem;
            max-width: 1180px;
        }

        .phx-shell {
            background: rgba(255, 255, 255, 0.97);
            border: 1px solid var(--phx-line);
            box-shadow: 0 18px 44px var(--phx-shadow);
            margin-bottom: 1.25rem;
        }

        .utility-bar {
            display: flex;
            justify-content: flex-end;
            gap: 1.15rem;
            padding: 0.55rem 1.1rem;
            border-bottom: 1px solid var(--phx-line);
            color: var(--phx-muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .brand-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1.25rem;
            padding: 1rem 1.1rem 0.95rem 1.1rem;
        }

        .brand-lockup {
            display: flex;
            align-items: stretch;
            gap: 0;
            min-height: 48px;
            border: 1px solid var(--phx-black);
        }

        .brand-phoenix,
        .brand-contact {
            display: grid;
            place-items: center;
            padding: 0.45rem 0.85rem;
            font-size: clamp(1.4rem, 3vw, 2.25rem);
            line-height: 1;
            font-weight: 900;
            letter-spacing: -0.045em;
        }

        .brand-phoenix {
            background: #ffffff;
            color: var(--phx-black);
        }

        .brand-contact {
            background: var(--phx-yellow);
            color: var(--phx-black);
        }

        .search-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            border: 1px solid var(--phx-line);
            background: var(--phx-silver);
            color: var(--phx-muted);
            border-radius: 999px;
            padding: 0.62rem 0.95rem;
            font-size: 0.88rem;
            min-width: 245px;
            justify-content: space-between;
        }

        .search-dot {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            background: var(--phx-yellow);
            box-shadow: 0 0 0 5px rgba(255, 210, 0, 0.18);
        }

        .nav-row {
            display: flex;
            flex-wrap: wrap;
            border-top: 1px solid var(--phx-line);
        }

        .nav-item {
            padding: 0.82rem 1.05rem;
            border-right: 1px solid var(--phx-line);
            color: var(--phx-ink);
            font-size: 0.85rem;
            text-transform: uppercase;
            letter-spacing: 0.045em;
            font-weight: 800;
        }

        .nav-item:first-child {
            background: var(--phx-yellow);
        }

        .hero-section {
            position: relative;
            overflow: hidden;
            display: grid;
            grid-template-columns: 1.22fr 0.78fr;
            gap: 1.5rem;
            align-items: stretch;
            background: linear-gradient(135deg, #f7f5ee 0%, #ffffff 55%, #ebebe8 100%);
            border: 1px solid var(--phx-line);
            box-shadow: 0 18px 46px var(--phx-shadow);
            padding: 2rem;
            margin-bottom: 1rem;
        }

        .hero-kicker {
            display: inline-flex;
            align-items: center;
            gap: 0.55rem;
            color: var(--phx-muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            font-weight: 900;
        }

        .hero-kicker::before {
            content: "";
            display: block;
            width: 34px;
            height: 5px;
            background: var(--phx-yellow);
        }

        .hero-title {
            margin: 0.75rem 0 0.7rem 0;
            color: var(--phx-black);
            font-size: clamp(2.1rem, 4vw, 4rem);
            line-height: 0.98;
            letter-spacing: -0.045em;
            font-weight: 900;
        }

        .hero-copy {
            max-width: 690px;
            color: #505050;
            font-size: 1.03rem;
            line-height: 1.72;
            margin-bottom: 1.2rem;
        }

        .hero-tags {
            display: flex;
            flex-wrap: wrap;
            gap: 0.55rem;
        }

        .hero-tag {
            background: #ffffff;
            border: 1px solid var(--phx-line);
            border-left: 5px solid var(--phx-yellow);
            padding: 0.55rem 0.72rem;
            font-size: 0.82rem;
            color: var(--phx-ink);
            font-weight: 800;
        }

        .hero-visual {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 245px;
        }

        .terminal-card {
            width: min(100%, 350px);
            min-height: 210px;
            border-radius: 2px;
            background: linear-gradient(145deg, #ffffff 0%, #efefec 100%);
            border: 1px solid #d4d4d0;
            box-shadow: 0 24px 44px rgba(0, 0, 0, 0.13);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
            padding: 1.15rem;
            transform: rotate(-2deg);
        }

        .terminal-block {
            width: 50px;
            height: 140px;
            background: linear-gradient(180deg, #ffe66d 0%, var(--phx-yellow) 52%, #e5b900 100%);
            border: 1px solid #caa600;
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.7), 0 10px 18px rgba(0,0,0,0.07);
            position: relative;
        }

        .terminal-block::before,
        .terminal-block::after {
            content: "";
            position: absolute;
            left: 11px;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            background: #ffffff;
            border: 3px solid #222222;
            box-sizing: border-box;
        }

        .terminal-block::before { top: 22px; }
        .terminal-block::after { bottom: 22px; }
        .terminal-block:nth-child(2) { transform: translateY(-8px); }
        .terminal-block:nth-child(3) { transform: translateY(7px); }
        .terminal-block:nth-child(4) { transform: translateY(-3px); }

        .process-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0.85rem;
            margin: 0 0 1.25rem 0;
        }

        .process-card {
            display: flex;
            gap: 0.85rem;
            min-height: 98px;
            background: rgba(255,255,255,0.94);
            border: 1px solid var(--phx-line);
            border-bottom: 4px solid var(--phx-yellow);
            padding: 1rem;
            box-shadow: 0 12px 24px rgba(0,0,0,0.05);
        }

        .process-number {
            flex: 0 0 auto;
            width: 34px;
            height: 34px;
            display: grid;
            place-items: center;
            background: var(--phx-black);
            color: #ffffff;
            font-weight: 900;
            font-size: 0.88rem;
        }

        .process-title {
            color: var(--phx-black);
            font-weight: 900;
            font-size: 0.98rem;
            margin-bottom: 0.25rem;
        }

        .process-text {
            color: var(--phx-muted);
            font-size: 0.86rem;
            line-height: 1.46;
        }

        .section-heading {
            margin: 1.15rem 0 0.75rem 0;
            padding: 0 0 0.65rem 0;
            border-bottom: 1px solid var(--phx-line);
        }

        .section-eyebrow {
            color: var(--phx-muted);
            font-size: 0.75rem;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            font-weight: 900;
        }

        .section-title {
            color: var(--phx-black);
            font-size: 1.42rem;
            font-weight: 900;
            margin-top: 0.15rem;
            letter-spacing: -0.02em;
        }

        .section-subtitle {
            color: var(--phx-muted);
            font-size: 0.94rem;
            line-height: 1.6;
            margin-top: 0.2rem;
        }

        .panel-title {
            color: var(--phx-black);
            font-size: 1.02rem;
            font-weight: 900;
            margin-bottom: 0.25rem;
        }

        .panel-title::before {
            content: "";
            display: inline-block;
            width: 9px;
            height: 9px;
            background: var(--phx-yellow);
            margin-right: 0.45rem;
            transform: translateY(-1px);
        }

        .panel-subtitle {
            color: var(--phx-muted);
            font-size: 0.88rem;
            margin-bottom: 0.8rem;
            line-height: 1.55;
        }

        div[data-testid="stTextArea"] textarea {
            background-color: #ffffff !important;
            border: 1px solid var(--phx-line) !important;
            border-left: 5px solid var(--phx-yellow) !important;
            border-radius: 0 !important;
            color: var(--phx-ink) !important;
            font-size: 0.95rem !important;
            min-height: 232px !important;
            box-shadow: 0 14px 28px rgba(0,0,0,0.04) !important;
        }

        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] div[data-baseweb="select"],
        div[data-testid="stMultiselect"] div[data-baseweb="select"] {
            background-color: #ffffff !important;
            border: 1px solid var(--phx-line) !important;
            border-radius: 0 !important;
            color: var(--phx-ink) !important;
        }

        div[data-testid="stFileUploader"] {
            background: #ffffff !important;
            border: 1px solid var(--phx-line) !important;
            border-left: 5px solid var(--phx-yellow) !important;
            border-radius: 0 !important;
            padding: 22px !important;
            min-height: 190px !important;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 14px 28px rgba(0,0,0,0.04) !important;
        }

        div[data-testid="stFileUploader"] section {
            width: 100%;
        }

        div[data-testid="stTextArea"] label,
        div[data-testid="stTextInput"] label,
        div[data-testid="stCheckbox"] label,
        div[data-testid="stSelectbox"] label,
        div[data-testid="stFileUploader"] label,
        div[data-testid="stMultiselect"] label,
        div[data-testid="stSlider"] label {
            color: var(--phx-black) !important;
            font-weight: 800 !important;
        }

        .stButton > button,
        div[data-testid="stDownloadButton"] > button {
            background: var(--phx-black) !important;
            color: #ffffff !important;
            border: 1px solid var(--phx-black) !important;
            border-radius: 0 !important;
            font-weight: 900 !important;
            letter-spacing: 0.04em !important;
            text-transform: uppercase !important;
            padding: 0.86rem 1rem !important;
            box-shadow: 0 14px 28px rgba(0,0,0,0.14);
        }

        .stButton > button:hover,
        div[data-testid="stDownloadButton"] > button:hover {
            background: var(--phx-yellow) !important;
            border-color: var(--phx-yellow) !important;
            color: var(--phx-black) !important;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0.9rem;
            margin: 1.25rem 0 1rem 0;
        }

        .metric-card {
            background: #ffffff;
            border: 1px solid var(--phx-line);
            border-top: 6px solid var(--phx-yellow);
            padding: 1.1rem;
            box-shadow: 0 12px 24px rgba(0,0,0,0.05);
        }

        .metric-label {
            color: var(--phx-muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.09em;
            font-weight: 900;
            margin-bottom: 0.45rem;
        }

        .metric-value {
            color: var(--phx-black);
            font-size: 2rem;
            font-weight: 900;
            line-height: 1.1;
        }

        .info-note {
            background: var(--phx-yellow-soft);
            border: 1px solid #e8cc50;
            color: var(--phx-black);
            padding: 0.92rem 1rem;
            font-size: 0.93rem;
            margin-top: 0.9rem;
            line-height: 1.55;
        }

        .url-preview {
            background: #ffffff;
            border: 1px solid var(--phx-line);
            border-left: 5px solid var(--phx-yellow);
            padding: 0.9rem 1rem;
            color: var(--phx-muted);
            font-size: 0.88rem;
            line-height: 1.55;
            margin-top: 0.8rem;
        }

        .footer-note {
            text-align: center;
            color: var(--phx-muted);
            font-size: 0.82rem;
            margin-top: 1.4rem;
            letter-spacing: 0.04em;
            text-transform: uppercase;
        }

        div[data-testid="stExpander"] {
            background: #ffffff;
            border: 1px solid var(--phx-line);
            border-radius: 0;
        }

        @media (max-width: 900px) {
            .utility-bar,
            .brand-row,
            .nav-row {
                justify-content: flex-start;
            }

            .brand-row,
            .hero-section,
            .process-grid,
            .metric-grid {
                grid-template-columns: 1fr;
            }

            .hero-section {
                display: block;
                padding: 1.35rem;
            }

            .hero-visual {
                margin-top: 1rem;
            }

            .search-pill {
                display: none;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="phx-shell">
        <div class="utility-bar">
            <span>Products</span>
            <span>Solutions</span>
            <span>Support</span>
            <span>Downloads</span>
        </div>
        <div class="brand-row">
            <div class="brand-lockup" aria-label="Phoenix Contact inspired text mark">
                <div class="brand-phoenix">PHOENIX</div>
                <div class="brand-contact">CONTACT</div>
            </div>
            <div class="search-pill">
                <span>Search product documentation</span>
                <span class="search-dot"></span>
            </div>
        </div>
        <div class="nav-row">
            <div class="nav-item">Product documentation</div>
            <div class="nav-item">Technical data</div>
            <div class="nav-item">Drawings</div>
            <div class="nav-item">Classifications</div>
            <div class="nav-item">Accessories</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero-section">
        <div>
            <div class="hero-kicker">Phoenix Contact product documentation</div>
            <div class="hero-title">Datasheet pack builder</div>
            <div class="hero-copy">
                Enter Phoenix Contact item codes, choose the PDF sections
                to include, download the product PDFs automatically, and generate one consolidated pack.
            </div>
            <div class="hero-tags">
                <div class="hero-tag">PHX item codes</div>
                <div class="hero-tag">Excel import</div>
                <div class="hero-tag">Selectable PDF blocks</div>
                <div class="hero-tag">Merged pack</div>
            </div>
        </div>
        <div class="hero-visual" aria-hidden="true">
            <div class="terminal-card">
                <div class="terminal-block"></div>
                <div class="terminal-block"></div>
                <div class="terminal-block"></div>
                <div class="terminal-block"></div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

step_col1, step_col2, step_col3 = st.columns(3)
with step_col1:
    render_step("01", "Add codes", "Paste PHX codes manually or import them from an Excel column.")
with step_col2:
    render_step("02", "Choose sections", "Select which Phoenix Contact PDF blocks are included in every datasheet.")
with step_col3:
    render_step("03", "Build pack", "Download the PDFs and merge them in the original input order.")


# -----------------------------------------------------------------------------
# Input section
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Build your PDF pack</div>
        <div class="section-title">Codes and source file</div>
        <div class="section-subtitle">
            Manual codes and Excel codes are combined automatically. Duplicates are removed while keeping the first occurrence.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

manual_codes: List[str] = []
excel_codes: List[str] = []
excel_df: Optional[pd.DataFrame] = None

input_col1, input_col2 = st.columns(2)

with input_col1:
    st.markdown(
        """
        <div class="panel-title">Paste item codes</div>
        <div class="panel-subtitle">
            Enter one code per line, or separate them with commas, spaces, or semicolons.
            The app accepts PHX-3201853 or 3201853.
        </div>
        """,
        unsafe_allow_html=True,
    )

    codes_text = st.text_area(
        "Paste item codes",
        height=232,
        placeholder="Example:\nPHX-3201853\nPHX-1212318\n3201853",
        label_visibility="collapsed",
    )

    manual_codes = normalize_codes(codes_text.splitlines())

with input_col2:
    st.markdown(
        """
        <div class="panel-title">Upload Excel file</div>
        <div class="panel-subtitle">
            Drag and drop your Excel file here, then choose the column containing Phoenix Contact item codes.
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_excel = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

    if uploaded_excel is not None:
        try:
            excel_df = read_excel_file(uploaded_excel)

            if excel_df.empty:
                st.warning("The uploaded Excel file is empty.")
            else:
                column_options = excel_df.columns.tolist()
                default_index = pick_default_excel_column(column_options)

                selected_column = st.selectbox(
                    "Select the column containing item codes",
                    options=column_options,
                    index=default_index,
                )

                excel_codes = extract_codes_from_selected_column(excel_df, selected_column)
                st.caption(f"{len(excel_codes)} code(s) detected from Excel.")

        except Exception as exc:
            st.error(f"Could not read Excel file: {exc}")


# -----------------------------------------------------------------------------
# PDF block section
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">PDF content</div>
        <div class="section-title">Choose what to include</div>
        <div class="section-subtitle">
            These selected blocks are added to the Phoenix Contact PDF URL for every item in the current pack.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

block_label_to_key = {label: key for key, label, _ in PDF_BLOCK_OPTIONS}
block_key_to_label = {key: label for key, label, _ in PDF_BLOCK_OPTIONS}
all_block_labels = [label for _, label, _ in PDF_BLOCK_OPTIONS]

selected_block_labels = st.multiselect(
    "PDF sections",
    options=all_block_labels,
    default=all_block_labels,
    help="Changing this list edits the blocks=... part of the Phoenix Contact API URL.",
)
selected_blocks = [block_label_to_key[label] for label in all_block_labels if label in selected_block_labels]

with st.expander("Section details", expanded=False):
    for key, label, description in PDF_BLOCK_OPTIONS:
        st.markdown(f"**{label}** (`{key}`): {description}")

all_codes_preview = normalize_codes(manual_codes + excel_codes)
if all_codes_preview and selected_blocks:
    preview_code = all_codes_preview[0]
    preview_url = build_phoenix_pdf_url(preview_code, selected_blocks)
    st.markdown(
        f"""
        <div class="url-preview">
            Preview for <strong>PHX-{preview_code}</strong>: the item number is encoded as
            <strong>{encode_item_number_for_phoenix(preview_code)}</strong> and inserted into the PDF API URL.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.code(preview_url, language="text")


# -----------------------------------------------------------------------------
# Pack settings
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Pack settings</div>
        <div class="section-title">Cover and output</div>
        <div class="section-subtitle">
            Optionally add a cover PDF before the downloaded Phoenix Contact datasheets.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

settings_col1, settings_col2 = st.columns(2)
with settings_col1:
    keep_going = st.checkbox("Skip failed codes and continue", value=True)
    include_cover = st.checkbox("Add cover page if available", value=False)
    output_name = st.text_input("Output file name", value="phoenix_contact_datasheet_pack.pdf")

with settings_col2:
    uploaded_cover = st.file_uploader(
        "Use another cover page (optional)",
        type=["pdf"],
        help="Leave empty to use cover.pdf from the repository root when cover pages are enabled.",
    )

with st.expander("Advanced connection settings", expanded=False):
    advanced_col1, advanced_col2, advanced_col3 = st.columns(3)
    with advanced_col1:
        realm = st.text_input("Realm", value=DEFAULT_REALM)
    with advanced_col2:
        locale = st.text_input("Locale", value=DEFAULT_LOCALE)
    with advanced_col3:
        max_workers_setting = st.slider("Parallel downloads", min_value=1, max_value=16, value=8)

st.markdown(
    """
    <div class="info-note">
        Code conversion example: PHX-3201853 becomes 3201853, then Base64 becomes MzIwMTg1Mw,
        then padding is removed before building the Phoenix Contact PDF API URL.
    </div>
    """,
    unsafe_allow_html=True,
)

run_clicked = st.button("Build PDF Pack", type="primary", use_container_width=True)


# -----------------------------------------------------------------------------
# Action / processing
# -----------------------------------------------------------------------------
if run_clicked:
    codes = normalize_codes(manual_codes + excel_codes)

    if not codes:
        st.error("Please enter item codes manually or upload an Excel file.")
        st.stop()

    if not selected_blocks:
        st.error("Please select at least one PDF section.")
        st.stop()

    cover_pdf_bytes, cover_error, cover_warning = get_cover_pdf_bytes(uploaded_cover, include_cover)
    if cover_error:
        st.error(cover_error)
        st.stop()
    if cover_warning:
        st.warning(cover_warning)

    max_workers = min(max_workers_setting, max(1, len(codes)))

    downloaded_pdfs, success_rows, failed_rows, _ = download_pdfs_parallel(
        codes=codes,
        selected_blocks=selected_blocks,
        realm=realm.strip() or DEFAULT_REALM,
        locale=locale.strip() or DEFAULT_LOCALE,
        max_workers=max_workers,
    )

    submitted_count = len(codes)
    downloaded_count = len(downloaded_pdfs)
    failed_count = len(failed_rows)

    render_metric_cards(submitted_count, downloaded_count, failed_count)

    if failed_rows and not keep_going:
        st.error("At least one code failed and 'Skip failed codes and continue' is disabled.")
        with st.expander("Failed codes", expanded=True):
            st.dataframe(failed_rows, use_container_width=True)
        st.stop()

    if downloaded_count == 0:
        st.error("No PDFs were downloaded, so no merged file could be created.")
        if failed_rows:
            with st.expander("Failed codes", expanded=True):
                st.dataframe(failed_rows, use_container_width=True)
        st.stop()

    try:
        merged_pdf = merge_pdf_bytes(downloaded_pdfs, cover_pdf_bytes=cover_pdf_bytes)

        st.success("Your consolidated Phoenix Contact PDF pack is ready.")

        st.download_button(
            label="Download Merged PDF",
            data=merged_pdf,
            file_name=ensure_pdf_filename(output_name),
            mime="application/pdf",
            use_container_width=True,
        )

        with st.expander("Downloaded items", expanded=False):
            st.dataframe(success_rows, use_container_width=True)

        if failed_rows:
            with st.expander("Failed codes", expanded=True):
                st.dataframe(failed_rows, use_container_width=True)

        report_rows = success_rows + failed_rows
        if report_rows:
            report_df = pd.DataFrame(report_rows)
            csv_bytes = report_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download Download Report CSV",
                data=csv_bytes,
                file_name="phoenix_contact_download_report.csv",
                mime="text/csv",
                use_container_width=True,
            )

    except Exception as exc:
        st.error(f"Failed to merge PDFs: {exc}")

st.markdown(
    """
    <div class="footer-note">
        Built for fast retrieval and packaging of Phoenix Contact product documentation.
    </div>
    """,
    unsafe_allow_html=True,
)
