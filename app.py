import base64
import os
import re
import shutil
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlencode

import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service


# -----------------------------------------------------------------------------
# Phoenix Contact PDF API settings
# -----------------------------------------------------------------------------
BASE_PDF_API_URL = "https://www.phoenixcontact.com/product/pdf/api/v1/{encoded_code}"

DEFAULT_TIMEOUT_SECONDS = 60

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

DEFAULT_SELECTED_BLOCK_LABELS = [
    "Technical data",
    "Drawings",
]

PDF_URL_PROFILES = [
    {
        "name": "AE",
        "realm": "ae",
        "locale": "en-AE",
    },
    {
        "name": "PC",
        "realm": "pc",
        "locale": "en-PC",
    },
]

LOGIN_URL_MARKERS = [
    "login.phoenixcontact.com",
    "/oauth2/",
    "_oidc_auth",
]

NOT_FOUND_TITLE_MARKERS = [
    "page not found",
    "not found",
]

APP_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
DEFAULT_COVER_PATHS = [
    APP_DIR / "cover.pdf",
    APP_DIR / "cover" / "cover.pdf",
]


# -----------------------------------------------------------------------------
# Code helpers
# -----------------------------------------------------------------------------
def clean_phoenix_code(value: str) -> str:
    """Return the Phoenix Contact numeric item number safely.

    Handles:
    - PHX-3010110 -> 3010110
    - 3010110 -> 3010110
    - 3010110.0 -> 3010110
    - 3,010,110 -> 3010110
    """
    if value is None:
        return ""

    value = str(value).strip()

    if not value or value.lower() in {"nan", "none", "null"}:
        return ""

    value = value.replace("\u2013", "-").replace("\u2014", "-")
    value = re.sub(r"\s*-\s*", "-", value)

    if "-" in value:
        value = value.rsplit("-", 1)[1].strip()

    value = value.replace(",", "").replace(" ", "")

    if re.fullmatch(r"\d+\.0+", value):
        value = value.split(".", 1)[0]

    if re.fullmatch(r"\d+\.\d+", value):
        value = value.split(".", 1)[0]

    return re.sub(r"[^0-9]", "", value)


def normalize_codes(raw_codes: Iterable[str]) -> List[str]:
    codes: List[str] = []

    for item in raw_codes:
        if item is None:
            continue

        item_str = str(item).strip()
        if not item_str:
            continue

        for part in re.split(r"[\s,;]+", item_str):
            code = clean_phoenix_code(part)
            if code:
                codes.append(code)

    unique: List[str] = []
    seen = set()

    for code in codes:
        if code not in seen:
            seen.add(code)
            unique.append(code)

    return unique


def encode_item_number_for_phoenix(item_number: str) -> str:
    item_number = clean_phoenix_code(item_number)

    if not item_number:
        raise ValueError("Phoenix Contact item number is empty.")

    return base64.b64encode(item_number.encode("ascii")).decode("ascii").rstrip("=")


def build_phoenix_pdf_url(
    item_number: str,
    realm: str,
    locale: str,
    selected_blocks: Sequence[str],
) -> str:
    """Build Phoenix Contact PDF API URL without action parameter.

    Example:
    https://www.phoenixcontact.com/product/pdf/api/v1/{converted_code}
    ?_realm=ae&_locale=en-AE&blocks=technical-data%2Cdrawings
    """
    if not selected_blocks:
        raise ValueError("At least one PDF content block must be selected.")

    encoded_code = encode_item_number_for_phoenix(item_number)

    query = urlencode(
        [
            ("_realm", realm),
            ("_locale", locale),
            ("blocks", ",".join(selected_blocks)),
        ]
    )

    return f"{BASE_PDF_API_URL.format(encoded_code=encoded_code)}?{query}"
    

def build_candidate_pdf_urls(
    code: str,
    selected_blocks: Sequence[str],
) -> List[Tuple[str, str]]:
    """Return only the two allowed Phoenix Contact PDF URLs.

    Order:
    1. AE / en-AE
    2. PC / en-PC

    The blocks are chosen by the user, with Technical data and Drawings selected by default.
    """
    candidates: List[Tuple[str, str]] = []

    for profile in PDF_URL_PROFILES:
        url = build_phoenix_pdf_url(
            code,
            realm=profile["realm"],
            locale=profile["locale"],
            selected_blocks=selected_blocks,
        )
        candidates.append((url, profile["name"]))

    return candidates
    

def ensure_pdf_filename(filename: str) -> str:
    filename = str(filename or "phoenix_contact_datasheet_pack.pdf").strip()
    filename = re.sub(r"[\\/:*?\"<>|]+", "-", filename)

    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"

    return filename


def read_excel_file(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file, dtype=str)


def pick_default_excel_column(columns: List[str]) -> int:
    preferred = [
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

    lookup = {str(col).strip().lower(): idx for idx, col in enumerate(columns)}

    for name in preferred:
        if name.lower() in lookup:
            return lookup[name.lower()]

    return 0


def extract_codes_from_selected_column(df: pd.DataFrame, selected_column: str) -> List[str]:
    if selected_column not in df.columns:
        return []

    return normalize_codes(df[selected_column].dropna().astype(str).tolist())


# -----------------------------------------------------------------------------
# PDF helpers
# -----------------------------------------------------------------------------
def trim_to_pdf_start(data: bytes) -> bytes:
    marker = data.find(b"%PDF-")
    return data[marker:] if marker > 0 else data


def is_valid_pdf_bytes(pdf_bytes: bytes) -> bool:
    try:
        reader = PdfReader(BytesIO(trim_to_pdf_start(pdf_bytes)), strict=False)
        return len(reader.pages) > 0
    except Exception:
        return False


def read_default_cover_pdf_bytes() -> Optional[bytes]:
    for cover_path in DEFAULT_COVER_PATHS:
        if cover_path.is_file():
            data = cover_path.read_bytes()
            if is_valid_pdf_bytes(data):
                return data

    return None


def get_cover_pdf_bytes(
    uploaded_cover,
    include_cover: bool,
) -> Tuple[Optional[bytes], Optional[str], Optional[str]]:
    if not include_cover:
        return None, None, None

    if uploaded_cover is not None:
        data = uploaded_cover.getvalue()

        if is_valid_pdf_bytes(data):
            return data, None, None

        return None, "The uploaded cover file is not a valid PDF.", None

    data = read_default_cover_pdf_bytes()

    if data is not None:
        return data, None, None

    return None, None, "No default cover.pdf was found. The pack will be created without a cover page."


def merge_pdf_bytes(pdf_byte_list: List[bytes], cover_pdf_bytes: Optional[bytes] = None) -> bytes:
    writer = PdfWriter()

    if cover_pdf_bytes:
        cover_reader = PdfReader(BytesIO(trim_to_pdf_start(cover_pdf_bytes)), strict=False)
        for page in cover_reader.pages:
            writer.add_page(page)

    for pdf_bytes in pdf_byte_list:
        reader = PdfReader(BytesIO(trim_to_pdf_start(pdf_bytes)), strict=False)
        for page in reader.pages:
            writer.add_page(page)

    output = BytesIO()
    writer.write(output)
    output.seek(0)

    return output.getvalue()


# -----------------------------------------------------------------------------
# Chrome / Selenium helpers
# -----------------------------------------------------------------------------
def find_chrome_binary() -> Optional[str]:
    candidates = [
        os.environ.get("CHROME_BINARY"),
        shutil.which("google-chrome"),
        shutil.which("google-chrome-stable"),
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
    ]

    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)

    return None


def find_chromedriver_binary() -> Optional[str]:
    candidates = [
        os.environ.get("CHROMEDRIVER"),
        shutil.which("chromedriver"),
        "/usr/bin/chromedriver",
        "/usr/lib/chromium/chromedriver",
        "/usr/lib/chromium-browser/chromedriver",
    ]

    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)

    return None


def create_selenium_driver(download_dir: Path, headless: bool) -> webdriver.Chrome:
    chrome_options = Options()

    chrome_binary = find_chrome_binary()
    if chrome_binary:
        chrome_options.binary_location = chrome_binary

    if headless:
        chrome_options.add_argument("--headless=new")

    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-background-networking")
    chrome_options.add_argument("--disable-sync")
    chrome_options.add_argument("--disable-translate")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--window-size=1280,900")
    chrome_options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )

    chrome_options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(download_dir.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "safebrowsing.enabled": True,
        },
    )

    chromedriver = find_chromedriver_binary()

    if chromedriver:
        driver = webdriver.Chrome(service=Service(chromedriver), options=chrome_options)
    else:
        driver = webdriver.Chrome(options=chrome_options)

    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {
            "behavior": "allow",
            "downloadPath": str(download_dir.resolve()),
        },
    )

    return driver


def clear_download_dir(download_dir: Path) -> None:
    download_dir.mkdir(parents=True, exist_ok=True)

    for path in download_dir.iterdir():
        if path.is_file():
            try:
                path.unlink()
            except OSError:
                pass


def wait_for_pdf_download(
    driver: webdriver.Chrome,
    download_dir: Path,
    timeout_seconds: int,
) -> Tuple[Optional[Path], str]:
    start = time.time()
    last_seen = ""

    while time.time() - start < timeout_seconds:
        try:
            current_url = driver.current_url or ""
            title = driver.title or ""
        except Exception:
            current_url = ""
            title = ""

        current_url_lower = current_url.lower()
        title_lower = title.lower()

        if any(marker in current_url_lower for marker in LOGIN_URL_MARKERS):
            return (
                None,
                "Phoenix Contact redirected this item to the sign-in page. "
                "This PDF appears to require a logged-in Phoenix Contact session.",
            )

        if any(marker in title_lower for marker in NOT_FOUND_TITLE_MARKERS):
            return (
                None,
                "Phoenix Contact returned a page-not-found response for this generated PDF URL.",
            )

        pdf_files = list(download_dir.glob("*.pdf"))
        partial_files = list(download_dir.glob("*.crdownload"))
        all_files = list(download_dir.iterdir())

        last_seen = ", ".join(path.name for path in all_files) if all_files else "no files yet"

        if pdf_files and not partial_files:
            newest = max(pdf_files, key=lambda p: p.stat().st_mtime)
            time.sleep(0.1)
            return newest, ""

        time.sleep(0.2)

    return None, f"Download did not finish within {timeout_seconds} seconds. Files: {last_seen}"


def download_pdf_bytes_with_selenium(
    driver: webdriver.Chrome,
    download_dir: Path,
    url: str,
    timeout_seconds: int,
) -> Tuple[bool, Optional[bytes], str]:
    clear_download_dir(download_dir)

    try:
        driver.get(url)
    except Exception as exc:
        return False, None, f"Chrome could not open URL: {exc}"

    pdf_path, wait_error = wait_for_pdf_download(driver, download_dir, timeout_seconds)

    if pdf_path is None:
        current_url = ""
        title = ""

        try:
            current_url = driver.current_url
            title = driver.title
        except Exception:
            pass

        return False, None, f"{wait_error}. Browser title: {title}. Current URL: {current_url}"

    try:
        data = pdf_path.read_bytes()
    except OSError as exc:
        return False, None, f"Downloaded PDF could not be read: {exc}"

    if not is_valid_pdf_bytes(data):
        return False, None, f"Downloaded file is not a valid PDF: {pdf_path.name}"

    return True, trim_to_pdf_start(data), ""


def process_code_with_driver(
    driver: webdriver.Chrome,
    download_dir: Path,
    index: int,
    code: str,
    selected_blocks: Sequence[str],
    timeout_seconds: int,
) -> Dict[str, object]:
    encoded = encode_item_number_for_phoenix(code)
    candidate_urls = build_candidate_pdf_urls(code, selected_blocks)

    errors: List[str] = []

    for url, profile_name in candidate_urls:
        ok, data, error = download_pdf_bytes_with_selenium(
            driver=driver,
            download_dir=download_dir,
            url=url,
            timeout_seconds=timeout_seconds,
        )

        if ok and data:
            return {
                "index": index,
                "code": code,
                "encoded": encoded,
                "ok": True,
                "pdf_bytes": data,
                "used_url": url,
                "profile": profile_name,
                "status": "Downloaded",
                "error": "",
            }

        errors.append(f"[{profile_name}] {error}")

    first_url = candidate_urls[0][0] if candidate_urls else ""

    error_text = " | ".join(errors)

    if "sign-in page" in error_text.lower() or "login.phoenixcontact.com" in error_text.lower():
        status = "Auth required"
    elif "page-not-found" in error_text.lower() or "page not found" in error_text.lower():
        status = "Not found"
    else:
        status = "Failed"

    return {
        "index": index,
        "code": code,
        "encoded": encoded,
        "ok": False,
        "pdf_bytes": None,
        "used_url": first_url,
        "profile": "",
        "status": status,
        "error": error_text,
    }


def split_work_round_robin(codes: List[str], workers: int) -> List[List[Tuple[int, str]]]:
    chunks: List[List[Tuple[int, str]]] = [[] for _ in range(workers)]

    for index, code in enumerate(codes):
        chunks[index % workers].append((index, code))

    return [chunk for chunk in chunks if chunk]


def download_chunk_with_one_browser(
    chunk: List[Tuple[int, str]],
    selected_blocks: Sequence[str],
    headless: bool,
    timeout_seconds: int,
) -> List[Dict[str, object]]:
    results: List[Dict[str, object]] = []

    with tempfile.TemporaryDirectory(prefix="phoenix_contact_downloads_") as tmp:
        download_dir = Path(tmp)
        driver = create_selenium_driver(download_dir, headless=headless)

        try:
            for index, code in chunk:
                try:
                    result = process_code_with_driver(
                        driver=driver,
                        download_dir=download_dir,
                        index=index,
                        code=code,
                        selected_blocks=selected_blocks,
                        timeout_seconds=timeout_seconds,
                    )
                except Exception as exc:
                    fallback_url = build_candidate_pdf_urls(code, selected_blocks)[0][0]
                    result = {
                        "index": index,
                        "code": code,
                        "encoded": encode_item_number_for_phoenix(code),
                        "ok": False,
                        "pdf_bytes": None,
                        "used_url": fallback_url,
                        "profile": "",
                        "status": "Failed",
                        "error": str(exc),
                    }

                results.append(result)
        finally:
            driver.quit()

    return results


def download_pdfs(
    codes: List[str],
    selected_blocks: Sequence[str],
    headless: bool,
    timeout_seconds: int,
    browser_workers: int,
):
    downloaded_pdfs: List[bytes] = []
    success_rows: List[Dict[str, object]] = []
    failed_rows: List[Dict[str, object]] = []
    results: List[Dict[str, object]] = []

    workers = max(1, min(int(browser_workers), len(codes)))
    chunks = split_work_round_robin(codes, workers)

    progress_bar = st.progress(0)
    status_text = st.empty()
    completed = 0

    if workers == 1:
        status_text.info("Opening Chrome and downloading PDFs...")
        results = download_chunk_with_one_browser(
            chunk=chunks[0],
            selected_blocks=selected_blocks,
            headless=headless,
            timeout_seconds=timeout_seconds,
        )
        progress_bar.progress(1.0)
    else:
        status_text.info(f"Opening {workers} Chrome workers and downloading PDFs...")

        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_count = {
                executor.submit(
                    download_chunk_with_one_browser,
                    chunk,
                    selected_blocks,
                    headless,
                    timeout_seconds,
                ): len(chunk)
                for chunk in chunks
            }

            for future in as_completed(future_to_count):
                chunk_results = future.result()
                results.extend(chunk_results)
                completed += future_to_count[future]

                status_text.info(f"Processed {completed} of {len(codes)} items")
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
                    "Path used": result.get("profile", ""),
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
                    "Path used": result.get("profile", ""),
                    "Status": result.get("status", "Failed"),
                    "Error": result.get("error", "Unknown error"),
                    "Source URL": result.get("used_url", ""),
                }
            )

    return downloaded_pdfs, success_rows, failed_rows


# -----------------------------------------------------------------------------
# UI helpers
# -----------------------------------------------------------------------------
def phx_section(eyebrow: str, title: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div class="phx-section-title">
            <div class="phx-eyebrow">{eyebrow}</div>
            <h2>{title}</h2>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_step(number: str, title: str, text: str) -> None:
    st.markdown(
        f"""
        <div class="step-card">
            <div class="step-number">{number}</div>
            <div>
                <div class="step-title">{title}</div>
                <div class="step-text">{text}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_cards(submitted: int, downloaded: int, failed: int) -> None:
    st.markdown(
        f"""
        <div class="metric-grid">
            <div class="metric-card"><span>Submitted</span><strong>{submitted}</strong></div>
            <div class="metric-card"><span>Downloaded</span><strong>{downloaded}</strong></div>
            <div class="metric-card"><span>Failed</span><strong>{failed}</strong></div>
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
# CSS
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
        :root {
            --phx-green: #93c11c;
            --phx-teal: #009ba3;
            --phx-black: #111111;
            --phx-ink: #222222;
            --phx-muted: #666666;
            --phx-line: #d9d9d9;
            --phx-bg: #f5f6f3;
            --phx-shadow: rgba(0, 0, 0, 0.075);
        }

        #MainMenu,
        footer,
        header[data-testid="stHeader"] {
            visibility: hidden;
            height: 0;
        }

        .stApp {
            background:
                radial-gradient(circle at top right, rgba(0, 155, 163, 0.12), transparent 26rem),
                radial-gradient(circle at top left, rgba(147, 193, 28, 0.13), transparent 24rem),
                linear-gradient(180deg, #ffffff 0%, var(--phx-bg) 68%, #ffffff 100%);
            color: var(--phx-ink);
            font-family: Arial, Helvetica, sans-serif;
        }

        .block-container {
            max-width: 1180px;
            padding-top: 1rem;
            padding-bottom: 2rem;
        }

        .phx-topbar {
            background: #ffffff;
            border: 1px solid var(--phx-line);
            border-bottom: 0;
            padding: 0.55rem 1rem;
            color: var(--phx-muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.06em;
        }

        .phx-brandbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: #ffffff;
            border: 1px solid var(--phx-line);
            padding: 1rem;
        }

        .phx-wordmark {
            display: inline-flex;
            align-items: stretch;
            border: 1px solid var(--phx-black);
            line-height: 1;
            box-shadow: 0 8px 22px var(--phx-shadow);
        }

        .phx-wordmark span {
            display: grid;
            place-items: center;
            min-height: 48px;
            padding: 0.45rem 0.85rem;
            color: var(--phx-black);
            font-size: clamp(1.35rem, 2.6vw, 2.25rem);
            font-weight: 900;
            letter-spacing: -0.045em;
            background: #ffffff;
        }

        .phx-wordmark span:last-child {
            background: var(--phx-green);
            color: var(--phx-black);
        }

        .hero-card {
            position: relative;
            overflow: hidden;
            background: linear-gradient(135deg, #ffffff 0%, #f7faf8 58%, #edf4f2 100%);
            border: 1px solid var(--phx-line);
            padding: clamp(1.25rem, 3vw, 2.25rem);
            margin: 1rem 0;
            box-shadow: 0 18px 42px var(--phx-shadow);
        }

        .hero-card::before {
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 6px;
            background: linear-gradient(90deg, var(--phx-green), var(--phx-teal));
        }

        .hero-kicker {
            display: flex;
            align-items: center;
            gap: 0.65rem;
            color: var(--phx-muted);
            font-size: 0.78rem;
            font-weight: 900;
            letter-spacing: 0.12em;
            text-transform: uppercase;
        }

        .hero-kicker::before {
            content: "";
            width: 34px;
            height: 5px;
            background: var(--phx-green);
        }

        .hero-title {
            margin: 0.8rem 0 0.65rem;
            color: var(--phx-black);
            font-size: clamp(2.15rem, 4.3vw, 4rem);
            line-height: 0.98;
            font-weight: 900;
            letter-spacing: -0.05em;
        }

        .hero-copy {
            max-width: 760px;
            color: #505050;
            font-size: 1.02rem;
            line-height: 1.65;
            margin-bottom: 1.1rem;
        }

        .hero-pills {
            display: flex;
            flex-wrap: wrap;
            gap: 0.5rem;
        }

        .hero-pill {
            background: #ffffff;
            border: 1px solid var(--phx-line);
            border-left: 5px solid var(--phx-teal);
            padding: 0.55rem 0.7rem;
            font-size: 0.82rem;
            font-weight: 800;
        }

        .step-grid,
        .metric-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0.8rem;
            margin-bottom: 1.1rem;
        }

        .step-card,
        .metric-card {
            background: #ffffff;
            border: 1px solid var(--phx-line);
            box-shadow: 0 10px 24px rgba(0,0,0,0.045);
        }

        .step-card {
            display: flex;
            gap: 0.8rem;
            padding: 1rem;
            border-bottom: 4px solid var(--phx-green);
            min-height: 94px;
        }

        .step-number {
            width: 34px;
            height: 34px;
            flex: 0 0 auto;
            display: grid;
            place-items: center;
            background: var(--phx-black);
            color: #ffffff;
            font-weight: 900;
        }

        .step-title {
            color: var(--phx-black);
            font-size: 0.98rem;
            font-weight: 900;
            margin-bottom: 0.25rem;
        }

        .step-text {
            color: var(--phx-muted);
            font-size: 0.86rem;
            line-height: 1.45;
        }

        .phx-section-title {
            margin: 1.2rem 0 0.7rem;
            padding-bottom: 0.65rem;
            border-bottom: 1px solid var(--phx-line);
        }

        .phx-eyebrow {
            color: var(--phx-muted);
            font-size: 0.74rem;
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: 0.13em;
        }

        .phx-section-title h2 {
            margin: 0.15rem 0 0;
            color: var(--phx-black);
            font-size: 1.42rem;
            line-height: 1.2;
            font-weight: 900;
            letter-spacing: -0.025em;
        }

        .phx-section-title p {
            margin: 0.25rem 0 0;
            color: var(--phx-muted);
            font-size: 0.94rem;
            line-height: 1.55;
        }

        .panel-title {
            margin-bottom: 0.22rem;
            color: var(--phx-black);
            font-size: 1.02rem;
            font-weight: 900;
        }

        .panel-title::before {
            content: "";
            display: inline-block;
            width: 9px;
            height: 9px;
            margin-right: 0.45rem;
            background: var(--phx-green);
            transform: translateY(-1px);
        }

        .panel-subtitle {
            margin-bottom: 0.85rem;
            color: var(--phx-muted);
            font-size: 0.88rem;
            line-height: 1.5;
        }

        .metric-card {
            border-top: 6px solid var(--phx-teal);
            padding: 1rem;
        }

        .metric-card span {
            display: block;
            color: var(--phx-muted);
            font-size: 0.76rem;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            font-weight: 900;
            margin-bottom: 0.3rem;
        }

        .metric-card strong {
            color: var(--phx-black);
            font-size: 2rem;
            line-height: 1;
        }

        div[data-testid="stTextArea"] textarea,
        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] div[data-baseweb="select"],
        div[data-testid="stMultiselect"] div[data-baseweb="select"] {
            background-color: #ffffff !important;
            border: 1px solid var(--phx-line) !important;
            border-radius: 0 !important;
            color: var(--phx-ink) !important;
            box-shadow: none !important;
        }

        div[data-testid="stTextArea"] textarea {
            min-height: 232px !important;
            border-left: 5px solid var(--phx-teal) !important;
        }

        div[data-testid="stFileUploader"] {
            background: #ffffff !important;
            border: 1px solid var(--phx-line) !important;
            border-left: 5px solid var(--phx-teal) !important;
            border-radius: 0 !important;
            padding: 1rem !important;
            min-height: 190px !important;
        }

        div[data-testid="stTextArea"] label,
        div[data-testid="stTextInput"] label,
        div[data-testid="stCheckbox"] label,
        div[data-testid="stSelectbox"] label,
        div[data-testid="stFileUploader"] label,
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
            padding: 0.85rem 1rem !important;
            box-shadow: 0 12px 24px rgba(0,0,0,0.12) !important;
        }

        .stButton > button:hover,
        div[data-testid="stDownloadButton"] > button:hover {
            background: var(--phx-teal) !important;
            border-color: var(--phx-teal) !important;
            color: #ffffff !important;
        }

        div[data-testid="stExpander"] {
            border: 1px solid var(--phx-line);
            border-radius: 0;
            background: #ffffff;
        }

        .footer-note {
            text-align: center;
            color: var(--phx-muted);
            font-size: 0.8rem;
            margin-top: 1.4rem;
            letter-spacing: 0.05em;
            text-transform: uppercase;
        }

        @media (max-width: 900px) {
            .step-grid,
            .metric-grid {
                grid-template-columns: 1fr;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Header and hero
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="phx-topbar">
        <div>Product documentation</div>
    </div>

    <div class="phx-brandbar">
        <div class="phx-wordmark" aria-label="Phoenix Contact inspired wordmark">
            <span>PHOENIX</span>
            <span>CONTACT</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero-card">
        <div>
            <div class="hero-kicker">Phoenix Contact PDF automation</div>
            <div class="hero-title">Datasheet pack builder</div>
            <div class="hero-copy">
                Build one consolidated PDF from Phoenix Contact item codes.
                The app tries only the UAE and PC datasheet API paths using technical data and drawings.
            </div>
            <div class="hero-pills">
                <div class="hero-pill">PHX item codes</div>
                <div class="hero-pill">AE / en-AE path</div>
                <div class="hero-pill">PC / en-PC path</div>
                <div class="hero-pill">Technical data + drawings</div>
                <div class="hero-pill">Merged PDF output</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="step-grid">', unsafe_allow_html=True)

step_cols = st.columns(3)

with step_cols[0]:
    render_step("01", "Add codes", "Paste PHX codes manually or import an Excel column.")

with step_cols[1]:
    render_step("02", "Download PDFs", "The app tries AE first, then PC, without an action parameter.")

with step_cols[2]:
    render_step("03", "Build pack", "Downloaded datasheets are validated, merged, and exported.")

st.markdown("</div>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# Inputs
# -----------------------------------------------------------------------------
phx_section(
    "Build your PDF pack",
    "Codes and source file",
    "Manual codes and Excel codes are combined automatically. Duplicates are removed while keeping the first occurrence.",
)

manual_codes: List[str] = []
excel_codes: List[str] = []

input_col1, input_col2 = st.columns(2)

with input_col1:
    st.markdown(
        """
        <div class="panel-title">Paste item codes</div>
        <div class="panel-subtitle">
            Use one code per line, or separate codes with commas, spaces, or semicolons.
            Accepted examples: PHX-3010110, 3010110.
        </div>
        """,
        unsafe_allow_html=True,
    )

    codes_text = st.text_area(
        "Paste item codes",
        height=232,
        placeholder="Example:\nPHX-3010110\nPHX-3201853\nPHX-3213140",
        label_visibility="collapsed",
    )

    manual_codes = normalize_codes(codes_text.splitlines())

with input_col2:
    st.markdown(
        """
        <div class="panel-title">Upload Excel file</div>
        <div class="panel-subtitle">
            Upload an .xlsx or .xls file, then select the column that contains Phoenix Contact item codes.
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

                selected_column = st.selectbox(
                    "Select the column containing item codes",
                    options=column_options,
                    index=pick_default_excel_column(column_options),
                )

                excel_codes = extract_codes_from_selected_column(excel_df, selected_column)
                st.caption(f"{len(excel_codes)} code(s) detected from Excel.")

        except Exception as exc:
            st.error(f"Could not read Excel file: {exc}")


# -----------------------------------------------------------------------------
# PDF content
# -----------------------------------------------------------------------------
phx_section(
    "PDF content",
    "Choose what to include",
    "By default, Technical data and Drawings are selected. You can add other available Phoenix Contact PDF sections.",
)

block_label_to_key = {label: key for key, label, _ in PDF_BLOCK_OPTIONS}
all_block_labels = [label for _, label, _ in PDF_BLOCK_OPTIONS]

selected_block_labels = st.multiselect(
    "PDF sections",
    options=all_block_labels,
    default=DEFAULT_SELECTED_BLOCK_LABELS,
    help="These selected sections are added to the blocks=... part of both AE and PC PDF URLs.",
)

selected_blocks = [
    block_label_to_key[label]
    for label in all_block_labels
    if label in selected_block_labels
]

with st.expander("Section details", expanded=False):
    for key, label, description in PDF_BLOCK_OPTIONS:
        st.markdown(f"**{label}** (`{key}`): {description}")
        

# -----------------------------------------------------------------------------
# Settings
# -----------------------------------------------------------------------------
phx_section(
    "Pack settings",
    "Cover and output",
    "The datasheet URL is fixed to technical-data and drawings. No action parameter is used.",
)

settings_col1, settings_col2 = st.columns(2)

with settings_col1:
    include_cover = st.checkbox("Add cover page", value=True)
    output_name = st.text_input("Output file name", value="phoenix_contact_datasheet_pack.pdf")

with settings_col2:
    uploaded_cover = st.file_uploader(
        "Use another cover page (optional)",
        type=["pdf"],
        help="Leave empty to use cover.pdf from the repository root when cover pages are enabled.",
    )

with st.expander("Advanced Chrome settings", expanded=False):
    c1, c2 = st.columns(2)

    with c1:
        headless = st.checkbox(
            "Run Chrome headless",
            value=True,
            help="Use True on Streamlit Cloud. Use False locally to see Chrome.",
        )

        timeout_seconds = st.slider(
            "Timeout",
            min_value=30,
            max_value=180,
            value=DEFAULT_TIMEOUT_SECONDS,
            step=10,
        )

    with c2:
        browser_workers = st.slider(
            "Chrome workers",
            min_value=1,
            max_value=4,
            value=1,
            help="Use one worker on Streamlit Cloud. Use two only if your host has enough RAM.",
        )

run_clicked = st.button("Build PDF Pack", type="primary", use_container_width=True)


# -----------------------------------------------------------------------------
# Action
# -----------------------------------------------------------------------------
if run_clicked:
    codes = normalize_codes(manual_codes + excel_codes)

    if not codes:
        st.error("Please enter item codes manually or upload an Excel file.")
        st.stop()

    cover_pdf_bytes, cover_error, cover_warning = get_cover_pdf_bytes(uploaded_cover, include_cover)

    if cover_error:
        st.error(cover_error)
        st.stop()

    if cover_warning:
        st.warning(cover_warning)

    try:
        downloaded_pdfs, success_rows, failed_rows = download_pdfs(
            codes=codes,
            selected_blocks=selected_blocks,
            headless=headless,
            timeout_seconds=timeout_seconds,
            browser_workers=browser_workers,
        )
        )
    except Exception as exc:
        st.error(f"Chrome/Selenium failed before downloads could complete: {exc}")
        st.info(
            "On Streamlit Cloud, include packages.txt with chromium and chromium-driver. "
            "Locally, install Google Chrome or Chromium."
        )
        st.stop()

    render_metric_cards(len(codes), len(downloaded_pdfs), len(failed_rows))

    if failed_rows:
        with st.expander("Items not downloaded", expanded=True):
            st.dataframe(failed_rows, use_container_width=True)

    if not downloaded_pdfs:
        st.error("No PDFs were downloaded, so no merged file could be created.")
        st.stop()

    try:
        merged_pdf = merge_pdf_bytes(downloaded_pdfs, cover_pdf_bytes=cover_pdf_bytes)

        st.success("Your consolidated Phoenix Contact PDF pack is ready.")

        st.download_button(
            "Download Merged PDF",
            data=merged_pdf,
            file_name=ensure_pdf_filename(output_name),
            mime="application/pdf",
            use_container_width=True,
        )

        with st.expander("Downloaded items", expanded=False):
            st.dataframe(success_rows, use_container_width=True)

        report_rows = success_rows + failed_rows

        if report_rows:
            report_df = pd.DataFrame(report_rows)
            csv_bytes = report_df.to_csv(index=False).encode("utf-8")

            st.download_button(
                "Download Report CSV",
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
        Built for retrieval and packaging of Phoenix Contact product documentation.
    </div>
    """,
    unsafe_allow_html=True,
)
