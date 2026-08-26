import io
import json
import os
import re
import shutil
import subprocess
import tempfile
from copy import deepcopy
from datetime import date
from pathlib import Path

import streamlit as st
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import pymupdf as fitz

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


def get_secret(name):
    value = os.getenv(name)
    if value:
        return value.strip()
    try:
        value = st.secrets.get(name)
        return value.strip() if isinstance(value, str) else value
    except (FileNotFoundError, KeyError):
        return None


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


def load_companies():
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

    companies = []
    for path in sorted(COMPANIES_DIR.glob("*.json")):
        try:
            companies.append(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            continue
    if database and companies:
        database.table("companies").upsert(
            [{"name": company["klien"], "data": company} for company in companies],
            on_conflict="name",
        ).execute()
    return companies


def save_company(company):
    database = get_database()
    if database:
        database.table("companies").upsert(
            {"name": company["klien"], "data": company}, on_conflict="name"
        ).execute()
        return

    safe_name = "".join(c for c in company["klien"] if c.isalnum() or c in " _-").strip()
    path = COMPANIES_DIR / f"{safe_name or 'company'}.json"
    path.write_text(json.dumps(company, ensure_ascii=False, indent=2), encoding="utf-8")


def delete_company(company):
    database = get_database()
    if database:
        response = database.table("companies").delete().eq("name", company["klien"]).execute()
        return bool(response.data)

    for path in COMPANIES_DIR.glob("*.json"):
        try:
            saved_company = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if saved_company.get("klien") == company.get("klien"):
            path.unlink()
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
    for field in ("klien", "giliran_no", "voltan", "ampere"):
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


def reset_form():
    for key in list(st.session_state):
        if key.startswith("remark_") or key in (
            "screen", "selected_company", "generated_docx", "generated_pdf", "remark_count"
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
        labels = [company["klien"] for company in companies]
        st.caption("Choose a company from the list below.")
        with st.container(height=300, border=True):
            selected = st.radio("Registered companies", labels, label_visibility="collapsed")
        selected_company = next(c for c in companies if c["klien"] == selected)
        pending_delete = st.session_state.get("pending_delete_company")
        if pending_delete == selected:
            st.warning(f'Are you sure you wish to delete "{selected}"? This cannot be undone.')
            confirm_delete = st.checkbox("Yes, delete this company profile")
            confirm_col, cancel_col = st.columns(2)
            with confirm_col:
                if st.button("Delete permanently", type="primary", disabled=not confirm_delete):
                    if delete_company(selected_company):
                        st.session_state.pop("pending_delete_company", None)
                        st.rerun()
            with cancel_col:
                if st.button("Cancel deletion"):
                    st.session_state.pop("pending_delete_company", None)
                    st.rerun()
        else:
            continue_col, delete_col = st.columns(2)
            with continue_col:
                if st.button("Continue", type="primary"):
                    st.session_state.selected_company = selected_company
                    st.session_state.screen = "report"
                    st.rerun()
            with delete_col:
                if st.button("Delete company"):
                    st.session_state.pending_delete_company = selected
                    st.rerun()
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
        alamat = st.text_area("Address")
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
        elif any(c["klien"].casefold() == klien.strip().casefold() for c in companies):
            st.error("That client is already registered.")
        else:
            company = {
                "klien": klien.strip(),
                "alamat": alamat.strip(),
                "giliran_no": giliran_no.strip(),
                "voltan": voltan.strip(),
                "ampere": ampere.strip(),
            }
            save_company(company)
            st.session_state.selected_company = company
            st.session_state.screen = "report"
            st.rerun()

elif st.session_state.screen == "report":
    company = st.session_state.selected_company
    st.header("New certificate")
    st.caption(f"Company: {company['klien']}")
    report_date = st.date_input("Date", value=date.today())
    st.subheader("Remarks")
    remarks = []
    for index in range(3):
        key = f"remark_section_{index}"
        if key not in st.session_state or st.session_state[key].strip().casefold() == "1. tiada":
            st.session_state[key] = "Tiada"
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
            st.session_state.screen = "download"
            st.rerun()

elif st.session_state.screen == "download":
    company = st.session_state.selected_company
    st.header("Document ready")
    st.success(f"The document for {company['klien']} has been generated.")
    filename = "".join(c for c in company["klien"] if c.isalnum() or c in " _-").strip() or "borang"
    st.download_button(
        "Download Word (.docx)",
        st.session_state.generated_docx,
        file_name=f"{filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
    if st.session_state.generated_pdf:
        st.download_button(
            "Download PDF (.pdf)",
            st.session_state.generated_pdf,
            file_name=f"{filename}.pdf",
            mime="application/pdf",
        )
    else:
        st.button("Download PDF (.pdf)", disabled=True)
        st.warning("PDF export is unavailable because the generated Word document could not be converted. Install LibreOffice and try again.")
    if st.button("Create another certificate"):
        st.session_state.screen = "report"
        st.rerun()
