import io
import json
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


def delete_company(company):
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
    address_lines = [line.strip() for line in company["alamat"].splitlines() if line.strip()]
    address_paragraphs = [doc.paragraphs[index] for index in (15, 16, 17, 18)]
    replace_runs(doc, [(15, 3), (15, 4), (15, 5)], address_lines[0] if address_lines else "")
    for paragraph, line in zip(address_paragraphs[1:], address_lines[1:]):
        leading_spaces = re.match(r"^\s*", paragraph.text).group(0)
        set_paragraph_text(paragraph, leading_spaces + line)
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
companies = load_companies()

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
