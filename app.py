import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import base64
from copy import deepcopy
from datetime import date
from pathlib import Path
from urllib.parse import quote
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import pymupdf as fitz

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
except ImportError:
    service_account = None
    build = None
    MediaIoBaseUpload = None

try:
    from supabase import Client, create_client
    from postgrest.exceptions import APIError
except ImportError:
    Client = None
    create_client = None
    APIError = Exception


BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "sample.docx"
COMPANIES_DIR = BASE_DIR / "companies"
COMPANIES_DIR.mkdir(exist_ok=True)

RECIPIENT_ADDRESSES = {
    "Melaka": """PEJABAT KAWASAN NEGERI SEMBILAN & MELAKA
TINGKAT 3, WISMA PERKESO
JALAN PERSEKUTUAN, MITC
75450 AYER KEROH
MELAKA DARUL AZIM
TEL: 06 - 231 9594 / 9597
FAKS: 06 - 231 9620""",
    "Selangor": """PEJABAT KAWASAN N.SELANGOR & WP (KL&PUTRAJAYA)
TINGKAT 10A, MENARA PKNS-PJ
17, JALAN YONG SHOOK LIN
46050 PETALING JAYA
SELANGOR
TEL : 03-79558930
FAKS : 03-79558939""",
}


def get_secret(name):
    value = os.getenv(name)
    if value:
        return value.strip()
    try:
        value = st.secrets.get(name)
        return value.strip() if isinstance(value, str) else value
    except (FileNotFoundError, KeyError):
        return None


def company_name(company):
    return company.get("company_name") or company.get("klien", "")


def company_nickname(company):
    return company.get("nickname") or company_name(company)


def company_category(company):
    value = company.get("category") or company.get("category_name") or "Uncategorized"
    return str(value).strip() or "Uncategorized"


def category_options(companies):
    categories = {"uncategorized": "Uncategorized"}
    for company in companies:
        category = company_category(company)
        categories.setdefault(category.casefold(), category)
    return [categories[key] for key in sorted(categories)]


NEW_CATEGORY_OPTION = "Create new category..."


def slugify(value):
    return re.sub(r"[^a-z0-9]+", "-", str(value).strip().lower()).strip("-") or "uncategorized"


def get_default_remarks(company):
    raw = company.get("default_remarks")
    if raw is None:
        raw = company.get("default_remark")
    if isinstance(raw, str):
        values = [raw.strip() for _ in range(3)]
    elif isinstance(raw, (list, tuple)):
        values = [str(item).strip() for item in raw[:3]]
        values.extend(["", "", ""])
        values = values[:3]
    else:
        values = ["", "", ""]
    return values


def sanitized_filename(value):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", value)
    return value.rstrip(" .").strip() or "borang"


def upload_to_google_drive(file_bytes, filename, mime_type):
    credentials_json = get_secret("GOOGLE_DRIVE_SERVICE_ACCOUNT_JSON")
    if not all((service_account, build, MediaIoBaseUpload, credentials_json)):
        raise RuntimeError(
            "Google Drive upload is not configured. Add the Google Drive dependencies "
            "and GOOGLE_DRIVE_SERVICE_ACCOUNT_JSON to Streamlit secrets."
        )
    if isinstance(credentials_json, str):
        credentials_json = json.loads(credentials_json)
    credentials = service_account.Credentials.from_service_account_info(
        credentials_json,
        scopes=["https://www.googleapis.com/auth/drive.file"],
    )
    drive = build("drive", "v3", credentials=credentials)
    metadata = {"name": filename}
    folder_id = get_secret("GOOGLE_DRIVE_FOLDER_ID")
    if folder_id:
        metadata["parents"] = [folder_id]
    uploaded = drive.files().create(
        body=metadata,
        media_body=MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type),
        fields="id,webViewLink",
    ).execute()
    return uploaded.get("webViewLink") or f"https://drive.google.com/file/d/{uploaded['id']}/view"


def get_database():
    if create_client is None:
        return None
    url = get_secret("SUPABASE_URL")
    key = (
        get_secret("SUPABASE_KEY")
        or get_secret("SUPABASE_SERVICE_ROLE_KEY")
        or get_secret("SUPABASE_ANON_KEY")
    )
    return create_client(url, key) if url and key else None


def get_github_config():
    token = get_secret("GITHUB_TOKEN")
    repository = get_secret("GITHUB_REPOSITORY")
    branch = get_secret("GITHUB_BRANCH") or "main"
    return token, repository, branch


def github_request(method, path, payload=None):
    token, repository, branch = get_github_config()
    if not token or not repository:
        return None
    encoded_path = quote(path, safe="/")
    url = f"https://api.github.com/repos/{repository}/contents/{encoded_path}"
    if method == "GET" and branch:
        url = f"{url}?ref={quote(branch)}"
    request = Request(
        url,
        method=method,
        headers={
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {token}",
            "X-GitHub-Api-Version": "2022-11-28",
        },
        data=json.dumps(payload).encode("utf-8") if payload is not None else None,
    )
    try:
        with urlopen(request, timeout=20) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        if error.code == 404:
            return None
        detail = error.read().decode("utf-8", errors="replace")
        if error.code == 403:
            raise RuntimeError(
                "GitHub rejected the write request. Update the fine-grained token "
                "permission for this repository to Contents: Read and write, then "
                "replace GITHUB_TOKEN in Streamlit Secrets and reboot the app."
            ) from error
        raise RuntimeError(f"GitHub storage request failed ({error.code}): {detail}") from error
    except URLError as error:
        raise RuntimeError(f"GitHub storage network request failed: {error.reason}") from error


def github_file_path(company_name_value):
    return f"companies/{company_profile_filename(company_name_value)}"


def company_profile_filename(company_name_value):
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", str(company_name_value))
    return f"{safe_name.strip() or 'company'}.json"


@st.cache_data(ttl=300, show_spinner=False)
def load_github_companies():
    files = github_request("GET", "companies")
    if files is None:
        return []
    companies = []
    for file_info in files:
        if file_info.get("type") != "file" or not file_info.get("name", "").endswith(".json"):
            continue
        contents = github_request("GET", file_info["path"])
        if not contents or not contents.get("content"):
            continue
        try:
            decoded = base64.b64decode(contents["content"]).decode("utf-8")
            companies.append(json.loads(decoded))
        except (ValueError, UnicodeDecodeError, json.JSONDecodeError):
            continue
    return companies


@st.cache_data(ttl=300, show_spinner=False)
def load_local_companies():
    companies = []
    for path in sorted(COMPANIES_DIR.glob("*.json")):
        try:
            companies.append(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            continue
    return companies


def save_github_company(company):
    path = github_file_path(company["klien"])
    existing = github_request("GET", path)
    payload = {
        "message": f"Update company: {company['klien']}",
        "content": base64.b64encode(
            json.dumps(company, ensure_ascii=False, indent=2).encode("utf-8")
        ).decode("ascii"),
        "branch": get_github_config()[2],
    }
    if existing and existing.get("sha"):
        payload["sha"] = existing["sha"]
    github_request("PUT", path, payload)


def delete_github_company(company):
    path = github_file_path(company["klien"])
    existing = github_request("GET", path)
    if not existing or not existing.get("sha"):
        return False
    github_request(
        "DELETE",
        path,
        {
            "message": f"Delete company: {company['klien']}",
            "sha": existing["sha"],
            "branch": get_github_config()[2],
        },
    )
    return True

st.set_page_config(page_title="Inspection Certificate", layout="centered")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');

    :root {
        --ink: #17211b;
        --muted: #657268;
        --paper: #f5f3ed;
        --panel: #fffdf8;
        --line: #d9ddd3;
        --leaf: #2f6b4f;
        --leaf-dark: #214b38;
        --sun: #e7ad45;
    }

    .stApp {
        background:
            radial-gradient(circle at 12% 8%, rgba(231, 173, 69, 0.18), transparent 22rem),
            linear-gradient(135deg, #f5f3ed 0%, #eef2eb 100%);
        color: var(--ink);
        font-family: 'DM Sans', sans-serif;
    }

    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stMainBlockContainer"] { max-width: 760px; padding-top: 3.5rem; }

    h1, h2, h3 {
        color: var(--ink) !important;
        font-family: 'Space Grotesk', sans-serif !important;
        letter-spacing: 0 !important;
    }

    h1 {
        font-size: 2.65rem !important;
        line-height: 1.05 !important;
        margin-bottom: 0.35rem !important;
    }

    h1::after {
        content: '';
        display: block;
        width: 3.5rem;
        height: 0.28rem;
        margin-top: 0.9rem;
        border-radius: 99px;
        background: var(--sun);
    }

    h2 { margin-top: 2rem !important; }
    h3 { color: var(--leaf-dark) !important; }
    [data-testid="stCaptionContainer"] { color: var(--muted); }

    [data-testid="stVerticalBlockBorderWrapper"] {
        background: rgba(255, 253, 248, 0.72);
        border: 1px solid var(--line);
        border-radius: 14px;
        box-shadow: 0 14px 35px rgba(36, 57, 43, 0.07);
    }

    [data-testid="stTextInput"], [data-testid="stTextArea"], [data-testid="stDateInput"] {
        margin-bottom: 0.45rem;
    }

    input, textarea {
        background: var(--panel) !important;
        border-color: var(--line) !important;
        border-radius: 9px !important;
        color: var(--ink) !important;
    }

    input:focus, textarea:focus {
        border-color: var(--leaf) !important;
        box-shadow: 0 0 0 1px var(--leaf) !important;
    }

    [data-testid="stButton"] button, [data-testid="stFormSubmitButton"] button,
    [data-testid="stDownloadButton"] button {
        min-height: 2.7rem;
        border-radius: 9px;
        border: 1px solid var(--line);
        font-weight: 600;
        transition: transform 140ms ease, box-shadow 140ms ease;
    }

    [data-testid="stButton"] button:hover, [data-testid="stFormSubmitButton"] button:hover,
    [data-testid="stDownloadButton"] button:hover {
        border-color: var(--leaf);
        box-shadow: 0 5px 14px rgba(33, 75, 56, 0.14);
        transform: translateY(-1px);
    }

    [data-testid="stButton"] button[kind="primary"], [data-testid="stFormSubmitButton"] button[kind="primary"],
    [data-testid="stDownloadButton"] button[kind="primary"] {
        background: var(--leaf);
        border-color: var(--leaf);
    }

    [data-testid="stAlert"] { border-radius: 10px; }
    hr { border-color: rgba(101, 114, 104, 0.25); }

    @media (max-width: 640px) {
        [data-testid="stMainBlockContainer"] { padding: 2rem 1rem 3rem; }
        h1 { font-size: 2.15rem !important; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=300, show_spinner=False)
def load_companies():
    local_companies = load_local_companies()
    if local_companies:
        return local_companies

    github_token, github_repository, _ = get_github_config()
    if github_token and github_repository:
        try:
            return load_github_companies()
        except RuntimeError as error:
            st.warning(f"GitHub storage is unavailable. Showing bundled company data. ({error})")
            return []
    database = get_database()
    if database:
        try:
            response = database.table("companies").select("data").order("name").execute()
        except APIError as error:
            raise RuntimeError(
                "Supabase rejected the companies query. "
                "Check that the companies table exists and that the configured key "
                f"has access. Supabase says: {error}"
            ) from error
        if response.data:
            return [row["data"] for row in response.data]

    companies = load_local_companies()
    if database and companies:
        database.table("companies").upsert(
            [{"name": company["klien"], "data": company} for company in companies],
            on_conflict="name",
        ).execute()
    return companies


def save_company(company):
    github_token, github_repository, _ = get_github_config()
    if github_token and github_repository:
        save_github_company(company)
        load_github_companies.clear()
        load_companies.clear()
        return
    database = get_database()
    if database:
        database.table("companies").upsert(
            {"name": company["klien"], "data": company}, on_conflict="name"
        ).execute()
        load_companies.clear()
        return

    path = COMPANIES_DIR / company_profile_filename(company["klien"])
    path.write_text(json.dumps(company, ensure_ascii=False, indent=2), encoding="utf-8")
    load_companies.clear()


def update_company(original_name, company):
    if original_name.casefold() != company["klien"].casefold():
        delete_company({"klien": original_name})
    save_company(company)


def delete_company(company):
    github_token, github_repository, _ = get_github_config()
    if github_token and github_repository:
        deleted = delete_github_company(company)
        if deleted:
            load_github_companies.clear()
            load_companies.clear()
        return deleted
    database = get_database()
    if database:
        response = database.table("companies").delete().eq("name", company["klien"]).execute()
        if response.data:
            load_companies.clear()
        return bool(response.data)

    for path in COMPANIES_DIR.glob("*.json"):
        try:
            saved_company = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if saved_company.get("klien") == company.get("klien"):
            path.unlink()
            load_companies.clear()
            return True
    return False


def replace_runs(doc, refs, value):
    first = True
    for paragraph_index, run_index in refs:
        run = doc.paragraphs[paragraph_index].runs[run_index]
        if first:
            run.text = value
            first = False
        else:
            run.text = ""


def clear_highlights(doc):
    for element in doc.element.iter():
        for highlight in list(element.findall(qn("w:highlight"))):
            element.remove(highlight)


def format_points(value):
    points = []
    for line in value.splitlines():
        point = re.sub(r"^\s*(?:[-*]\s*)?(?:\d+[.)]\s*)?", "", line).strip()
        if point:
            points.append(point)
    return "\n".join(points) or "Tiada"


def set_paragraph_text(paragraph, text):
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def apply_document_font(doc):
    normal_font = doc.styles["Normal"].font
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = normal_font.name
            run.font.size = normal_font.size


def set_remark_spacing(paragraph):
    paragraph.paragraph_format.line_spacing = 1.15


def replace_remark_paragraph(paragraph, value):
    points = [line for line in format_points(value).splitlines() if line]
    set_remark_spacing(paragraph)
    set_paragraph_text(paragraph, points[0])
    for point in points[1:]:
        clone = deepcopy(paragraph._p)
        paragraph._p.addnext(clone)
        paragraph = Paragraph(clone, paragraph._parent)
        set_remark_spacing(paragraph)
        set_paragraph_text(paragraph, point)


def template_refs():
    return {
        "klien": [(14, 4), (14, 5), (14, 6)],
        "alamat": [(15, 3), (15, 4), (15, 5), (16, 0), (17, 0), (18, 0)],
        "giliran_no": [(19, 1)],
        "voltan": [(19, 3), (19, 4), (19, 5)],
        "ampere": [(19, 7), (19, 8)],
        "date": [(21, 1), (21, 2), (21, 3), (21, 4), (21, 5)],
        "remarks": [[(37, 0)], [(39, 0)], [(41, 0)]],
    }


def generate_docx(company, report_date, remarks):
    doc = Document(TEMPLATE_PATH)
    refs = template_refs()
    recipient_lines = RECIPIENT_ADDRESSES.get(
        company.get("kepada", "Melaka"), RECIPIENT_ADDRESSES["Melaka"]
    ).splitlines()
    recipient_paragraphs = doc.paragraphs[5:12]
    for paragraph in recipient_paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.left_indent = 0
        paragraph.paragraph_format.first_line_indent = 0
    set_paragraph_text(recipient_paragraphs[0], f"Kepada\t:\t {recipient_lines[0]}")
    for paragraph, line in zip(recipient_paragraphs[1:], recipient_lines[1:]):
        paragraph.paragraph_format.left_indent = Inches(0.8)
        set_paragraph_text(paragraph, line)
    replace_runs(doc, refs["klien"], company_name(company))
    for field in ("giliran_no", "voltan", "ampere"):
        replace_runs(doc, refs[field], company[field])
    address_lines = [line.strip().upper() for line in company["alamat"].splitlines() if line.strip()]
    address_paragraphs = [doc.paragraphs[index] for index in (15, 16, 17, 18)]
    replace_runs(doc, [(15, 3), (15, 4), (15, 5)], address_lines[0] if address_lines else "")
    continuation_indent = address_paragraphs[1].paragraph_format.left_indent
    for paragraph, line in zip(address_paragraphs[1:], address_lines[1:]):
        paragraph.paragraph_format.left_indent = continuation_indent
        paragraph.paragraph_format.first_line_indent = None
        set_paragraph_text(paragraph, line)
    for paragraph in address_paragraphs[len(address_lines):]:
        set_paragraph_text(paragraph, "")
    replace_runs(doc, refs["date"], report_date.strftime("%-d/%-m/%Y"))
    remark_paragraphs = [doc.paragraphs[37], doc.paragraphs[39], doc.paragraphs[41]]
    for index, paragraph in enumerate(remark_paragraphs):
        replace_remark_paragraph(paragraph, remarks[index] if index < len(remarks) else "")
    for paragraph in list(doc.paragraphs):
        if paragraph.text.strip() in {"1 / 2", "2 / 2"}:
            paragraph._element.getparent().remove(paragraph._element)
    clear_highlights(doc)
    apply_document_font(doc)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def convert_to_pdf(docx_bytes):
    converter = shutil.which("libreoffice") or shutil.which("soffice")
    if converter is None:
        return None
    with tempfile.TemporaryDirectory() as folder:
        source = Path(folder) / "filled.docx"
        source.write_bytes(docx_bytes)
        result = subprocess.run(
            [converter, "--headless", "--convert-to", "pdf", "--outdir", folder, str(source)],
            capture_output=True,
            text=True,
            timeout=60,
        )
        pdf_path = source.with_suffix(".pdf")
        if result.returncode != 0 or not pdf_path.exists():
            return None
        pdf_bytes = pdf_path.read_bytes()
        with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
            if pdf.page_count != 2:
                return None
        return pdf_bytes


def create_share_text(company, report_date, remarks):
    remark_text = "\n".join(
        f"{index + 1}. {format_points(value)}"
        for index, value in enumerate(remarks)
    )
    return (
        "Inspection Certificate\n"
        f"Company: {company_name(company)}\n"
        f"Kepada: {company.get('kepada', 'Melaka')}\n"
        f"Date: {report_date.strftime('%-d/%-m/%Y')}\n"
        f"Circuit No.: {company['giliran_no']}\n"
        f"Voltage: {company['voltan']}\n"
        f"Amperage: {company['ampere']}\n"
        f"Remarks:\n{remark_text}"
    )


def reset_form():
    for key in list(st.session_state):
        if key.startswith("remark_") or key in (
            "screen", "selected_company", "generated_docx", "generated_pdf",
            "generated_date", "drive_link", "remark_count"
        ):
            st.session_state.pop(key, None)


if "screen" not in st.session_state:
    st.session_state.screen = "home"

st.title("Inspection Certificate")
try:
    companies = load_companies()
except RuntimeError as error:
    st.error(str(error))
    st.stop()

if st.session_state.screen == "home":
    st.header("Select a company")
    if companies:
        grouped_companies = {}
        for company in companies:
            category = company_category(company)
            canonical_category = next(
                (option for option in category_options(companies) if option.casefold() == category.casefold()),
                "Uncategorized",
            )
            grouped_companies.setdefault(canonical_category, []).append(company)
        st.caption("Choose a company from the list below.")
        for category_index, category in enumerate(sorted(grouped_companies, key=str.casefold)):
            st.subheader(category)
            category_companies = grouped_companies[category]
            labels = [company_nickname(company) for company in category_companies]
            selected = st.radio(
                f"Companies in {category}",
                labels,
                key=f"company_radio_{category_index}",
                label_visibility="collapsed",
            )
            selected_company = category_companies[labels.index(selected)]
            pending_delete = st.session_state.get("pending_delete_company")
            if pending_delete == company_name(selected_company):
                st.warning(f'Are you sure you wish to delete "{selected}"? This cannot be undone.')
                confirm_delete = st.checkbox("Yes, delete this company profile")
                confirm_col, cancel_col = st.columns(2)
                with confirm_col:
                    if st.button(
                        "Delete permanently",
                        type="primary",
                        disabled=not confirm_delete,
                        key=f"delete_permanent_{slugify(company_name(selected_company))}",
                    ):
                        if delete_company(selected_company):
                            st.session_state.pop("pending_delete_company", None)
                            st.rerun()
                with cancel_col:
                    if st.button("Cancel deletion", key=f"cancel_delete_{slugify(company_name(selected_company))}"):
                        st.session_state.pop("pending_delete_company", None)
                        st.rerun()
            else:
                continue_col, edit_col, delete_col = st.columns(3)
                with continue_col:
                    if st.button("Continue", type="primary", key=f"continue_{slugify(company_name(selected_company))}"):
                        st.session_state.selected_company = selected_company
                        st.session_state.screen = "report"
                        st.rerun()
                with edit_col:
                    if st.button("Edit company", key=f"edit_{slugify(company_name(selected_company))}"):
                        st.session_state.edit_company = selected_company
                        st.session_state.screen = "edit_company"
                        st.rerun()
                with delete_col:
                    if st.button("Delete company", key=f"delete_{slugify(company_name(selected_company))}"):
                        st.session_state.pending_delete_company = company_name(selected_company)
                        st.rerun()
            st.divider()
    else:
        st.info("No companies have been registered yet.")

    st.divider()
    if st.button("Create a new company profile"):
        st.session_state.screen = "create_company"
        st.rerun()

elif st.session_state.screen == "create_company":
    st.header("Create a new company profile")
    if st.button("Back to companies"):
        st.session_state.screen = "home"
        st.rerun()
    with st.form("new_company"):
        klien = st.text_input("Client")
        nickname = st.text_input("Nickname", help="Used only in the company selection list.")
        categories = category_options(companies)
        category_choice = st.selectbox(
            "Category",
            categories + [NEW_CATEGORY_OPTION],
            index=categories.index("Uncategorized"),
        )
        new_category = st.text_input(
            "New category (used when selected above)",
            key="new_category",
        )
        category = new_category if category_choice == NEW_CATEGORY_OPTION else category_choice
        kepada = st.selectbox("Kepada", list(RECIPIENT_ADDRESSES))
        alamat = st.text_area("Address")
        st.caption("Default remarks (same as the three remark sections on the certificate)")
        default_remarks = [
            st.text_area("Default remark 1", key="new_default_remark_1", help="Prefilled in remark section 1.") ,
            st.text_area("Default remark 2", key="new_default_remark_2", help="Prefilled in remark section 2.") ,
            st.text_area("Default remark 3", key="new_default_remark_3", help="Prefilled in remark section 3.") ,
        ]
        col1, col2, col3 = st.columns(3)
        with col1:
            giliran_no = st.text_input("Circuit No.")
        with col2:
            voltan = st.text_input("Voltage")
        with col3:
            ampere = st.text_input("Amperage")
        submitted = st.form_submit_button("Save company", type="primary")
    if submitted:
        if not all((klien.strip(), alamat.strip(), giliran_no.strip(), voltan.strip(), ampere.strip())):
            st.error("Please complete all company details.")
        elif any(company_name(c).casefold() == klien.strip().casefold() for c in companies):
            st.error("That client is already registered.")
        else:
            values = [remark.strip() for remark in default_remarks]
            company = {
                "klien": klien.strip(),
                "company_name": klien.strip(),
                "nickname": nickname.strip(),
                "category": category.strip() or "Uncategorized",
                "default_remark": values[0],
                "default_remarks": values,
                "kepada": kepada,
                "alamat": alamat.strip(),
                "giliran_no": giliran_no.strip(),
                "voltan": voltan.strip(),
                "ampere": ampere.strip(),
            }
            try:
                save_company(company)
            except RuntimeError as error:
                st.error(str(error))
            else:
                st.session_state.selected_company = company
                st.session_state.screen = "report"
                st.rerun()

elif st.session_state.screen == "edit_company":
    original_company = st.session_state.edit_company
    st.header("Edit company profile")
    if st.button("Back to companies"):
        st.session_state.pop("edit_company", None)
        st.session_state.screen = "home"
        st.rerun()
    with st.form("edit_company_form"):
        klien = st.text_input("Client", value=company_name(original_company))
        nickname = st.text_input("Nickname", value=original_company.get("nickname", ""), help="Used only in the company selection list.")
        categories = category_options(companies)
        current_category = company_category(original_company)
        category_choice = st.selectbox(
            "Category",
            categories + [NEW_CATEGORY_OPTION],
            index=next(
                (index for index, option in enumerate(categories) if option.casefold() == current_category.casefold()),
                0,
            ),
        )
        new_category = st.text_input(
            "New category (used when selected above)",
            key="edit_new_category",
        )
        category = new_category if category_choice == NEW_CATEGORY_OPTION else category_choice
        kepada_options = list(RECIPIENT_ADDRESSES)
        current_kepada = original_company.get("kepada", "Melaka")
        kepada = st.selectbox(
            "Kepada",
            kepada_options,
            index=kepada_options.index(current_kepada)
            if current_kepada in kepada_options
            else 0,
        )
        alamat = st.text_area("Address", value=original_company["alamat"])
        default_remarks = get_default_remarks(original_company)
        st.caption("Default remarks (same as the three remark sections on the certificate)")
        default_remarks_inputs = [
            st.text_area("Default remark 1", value=default_remarks[0], help="Prefilled in remark section 1."),
            st.text_area("Default remark 2", value=default_remarks[1], help="Prefilled in remark section 2."),
            st.text_area("Default remark 3", value=default_remarks[2], help="Prefilled in remark section 3."),
        ]
        col1, col2, col3 = st.columns(3)
        with col1:
            giliran_no = st.text_input("Circuit No.", value=original_company["giliran_no"])
        with col2:
            voltan = st.text_input("Voltage", value=original_company["voltan"])
        with col3:
            ampere = st.text_input("Amperage", value=original_company["ampere"])
        submitted = st.form_submit_button("Save changes", type="primary")
    if submitted:
        values = [remark.strip() for remark in default_remarks_inputs]
        company = {
            "klien": klien.strip(),
            "company_name": klien.strip(),
            "nickname": nickname.strip(),
            "category": category.strip() or "Uncategorized",
            "default_remark": values[0],
            "default_remarks": values,
            "kepada": kepada,
            "alamat": alamat.strip(),
            "giliran_no": giliran_no.strip(),
            "voltan": voltan.strip(),
            "ampere": ampere.strip(),
        }
        duplicate = any(
            company_name(c).casefold() == company_name(company).casefold()
            and company_name(c).casefold() != company_name(original_company).casefold()
            for c in companies
        )
        if not all(company.values()):
            st.error("Please complete all company details.")
        elif duplicate:
            st.error("That client is already registered.")
        else:
            try:
                update_company(company_name(original_company), company)
            except RuntimeError as error:
                st.error(str(error))
            else:
                st.session_state.pop("edit_company", None)
                st.session_state.selected_company = company
                st.session_state.screen = "report"
                st.rerun()

elif st.session_state.screen == "report":
    company = st.session_state.selected_company
    st.header("New certificate")
    st.caption(f"Company: {company_name(company)}")
    report_date = st.date_input("Date", value=date.today())
    st.subheader("Remarks")
    default_remarks = get_default_remarks(company)
    remarks = []
    for index in range(3):
        key = f"remark_section_{index}"
        if key not in st.session_state:
            st.session_state[key] = default_remarks[index].strip() or "Tiada"
        remarks.append(
            st.text_area(
                f"Remark section {index + 1}",
                key=key,
                placeholder="Enter one point per line",
                height=120,
            )
        )
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Change company"):
            reset_form()
            st.rerun()
    with col2:
        if st.button("Generate document", type="primary"):
            st.session_state.remarks = remarks
            st.session_state.generated_docx = generate_docx(company, report_date, remarks)
            st.session_state.generated_pdf = convert_to_pdf(st.session_state.generated_docx)
            st.session_state.share_text = create_share_text(company, report_date, remarks)
            st.session_state.generated_date = report_date
            st.session_state.screen = "download"
            st.rerun()

elif st.session_state.screen == "download":
    company = st.session_state.selected_company
    st.header("Document ready")
    st.success(f"The document for {company_name(company)} has been generated.")
    share_text = st.session_state.share_text
    filename = f"{sanitized_filename(company_name(company))}_{st.session_state.generated_date:%Y-%m-%d}"

    col_primary_1, col_primary_2 = st.columns(2)
    with col_primary_1:
        st.download_button("Download Word (.docx)", st.session_state.generated_docx,
                           file_name=f"{filename}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
    with col_primary_2:
        if st.session_state.generated_pdf:
            st.download_button("Download PDF (.pdf)", st.session_state.generated_pdf,
                               file_name=f"{filename}.pdf", mime="application/pdf",
                               use_container_width=True)
        else:
            st.info("PDF export unavailable", icon="ℹ️")

    with st.popover("Export", use_container_width=True):
        if st.button("Upload Word to Google Drive", use_container_width=True):
            try:
                st.session_state.drive_link = upload_to_google_drive(
                    st.session_state.generated_docx, f"{filename}.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except (RuntimeError, ValueError, KeyError, json.JSONDecodeError) as error:
                st.error(str(error))
        if st.session_state.get("drive_link"):
            st.link_button("Open Google Drive file", st.session_state.drive_link, use_container_width=True)
        st.link_button("Share via WhatsApp", f"https://wa.me/?text={quote(share_text)}", use_container_width=True)
        st.download_button("Download text summary (.txt)", share_text,
                           file_name=f"{filename}.txt", mime="text/plain", use_container_width=True)
    if st.button("Create another certificate"):
        for index in range(3):
            st.session_state.pop(f"remark_section_{index}", None)
        st.session_state.screen = "report"
        st.rerun()
