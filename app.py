import io
import json
import re
import shutil
import subprocess
import tempfile
from datetime import date
from pathlib import Path

import streamlit as st
from docx import Document
from docx.oxml.ns import qn
import fitz


BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "sample.docx"
COMPANIES_DIR = BASE_DIR / "companies"
COMPANIES_DIR.mkdir(exist_ok=True)

st.set_page_config(page_title="Inspection Certificate", layout="centered")


def load_companies():
    companies = []
    for path in sorted(COMPANIES_DIR.glob("*.json")):
        try:
            companies.append(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            continue
    return companies


def save_company(company):
    safe_name = "".join(c for c in company["klien"] if c.isalnum() or c in " _-").strip()
    path = COMPANIES_DIR / f"{safe_name or 'company'}.json"
    path.write_text(json.dumps(company, ensure_ascii=False, indent=2), encoding="utf-8")


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
            points.append(f"{len(points) + 1}. {point}")
    return "\n".join(points)


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
    for field in ("klien", "alamat", "giliran_no", "voltan", "ampere"):
        replace_runs(doc, refs[field], company[field])
    replace_runs(doc, refs["date"], report_date.strftime("%-d/%-m/%Y"))
    for index, field_refs in enumerate(refs["remarks"]):
        replace_runs(doc, field_refs, format_points(remarks[index] if index < len(remarks) else ""))
    clear_highlights(doc)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def convert_to_pdf(docx_bytes):
    converter = shutil.which("libreoffice") or shutil.which("soffice")
    if converter is None:
        return create_fallback_pdf(docx_bytes)
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
            return create_fallback_pdf(docx_bytes)
        return pdf_path.read_bytes()


def create_fallback_pdf(docx_bytes):
    docx = Document(io.BytesIO(docx_bytes))
    pdf = fitz.open()
    page = pdf.new_page(width=595, height=842)
    cursor = 48
    for paragraph in docx.paragraphs:
        text = paragraph.text.strip()
        if not text:
            cursor += 8
            continue
        if cursor > 780:
            page = pdf.new_page(width=595, height=842)
            cursor = 48
        page.insert_textbox(
            fitz.Rect(48, cursor, 547, cursor + 42),
            text,
            fontsize=9,
            fontname="helv",
            color=(0, 0, 0),
        )
        cursor += max(18, 12 * (text.count("\n") + 1))
    output = io.BytesIO()
    pdf.save(output)
    pdf.close()
    return output.getvalue()


def reset_form():
    for key in list(st.session_state):
        if key.startswith("remark_") or key in (
            "screen", "selected_company", "generated_docx", "generated_pdf", "remark_count"
        ):
            st.session_state.pop(key, None)


if "screen" not in st.session_state:
    st.session_state.screen = "home"

st.title("Inspection Certificate")
companies = load_companies()

if st.session_state.screen == "home":
    st.header("Select a company")
    if companies:
        labels = [company["klien"] for company in companies]
        st.caption("Choose a company from the list below.")
        with st.container(height=300, border=True):
            selected = st.radio("Registered companies", labels, label_visibility="collapsed")
        if st.button("Continue", type="primary"):
            st.session_state.selected_company = next(c for c in companies if c["klien"] == selected)
            st.session_state.screen = "report"
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
        if key not in st.session_state:
            st.session_state[key] = "No defect" if index == 0 else ""
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
        st.warning("PDF export requires LibreOffice to be installed on the host.")
    if st.button("Create another certificate"):
        st.session_state.screen = "report"
        st.rerun()
