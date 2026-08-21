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
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def convert_to_pdf(docx_bytes):
    converter = shutil.which("libreoffice") or shutil.which("soffice")
    if converter is None:
        return create_reference_pdf(docx_bytes)
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
            return create_reference_pdf(docx_bytes)
        pdf_bytes = pdf_path.read_bytes()
        with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
            if pdf.page_count != 2:
                return create_reference_pdf(docx_bytes)
        return pdf_bytes


def create_reference_pdf(docx_bytes):
    reference_path = BASE_DIR / "PDF.sample.pdf"
    if not reference_path.exists():
        return create_fallback_pdf(docx_bytes)

    docx = Document(io.BytesIO(docx_bytes))
    pdf = fitz.open(reference_path)
    page_one = pdf[0]

    def replace_text(page, old_text, new_text):
        matches = page.search_for(old_text)
        for rect in matches:
            page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
            page.insert_text((rect.x0, rect.y1 - 2), new_text, fontsize=8, fontname="helv")

    replace_text(page_one, "Jonathan Joe star", docx.paragraphs[14].text.split("\t")[-1].strip())
    replace_text(page_one, "At MY HOUSE", docx.paragraphs[15].text.split("\t")[-1].strip())
    replace_text(page_one, "2", docx.paragraphs[19].text.split("\t")[0].split(":")[-1].strip())
    replace_text(page_one, "400k", docx.paragraphs[19].text.split("\t")[1].split(":")[-1].strip())
    replace_text(page_one, "50A", docx.paragraphs[19].text.split("\t")[2].split(":")[-1].strip())

    date_match = re.search(r"\b\d{1,2}/\d{1,2}/\d{4}\b", docx.paragraphs[21].text)
    if date_match:
        replace_text(page_one, "21/8/2026", date_match.group())

    remark_groups = []
    current_group = []
    for paragraph in docx.paragraphs:
        if paragraph.style.name == "List Paragraph":
            current_group.append(paragraph.text)
        elif current_group:
            remark_groups.append(current_group)
            current_group = []
    if current_group:
        remark_groups.append(current_group)

    page_two = pdf[1]
    for y, group in zip((100, 180, 260), remark_groups[:3]):
        page_two.draw_rect(fitz.Rect(60, y, 560, y + 48), color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
        for index, text in enumerate(group):
            page_two.insert_text((63, y + 16 + index * 12), f"{index + 1}. {text}", fontsize=8, fontname="helv")

    output = io.BytesIO()
    pdf.save(output)
    pdf.close()
    return output.getvalue()


def create_fallback_pdf(docx_bytes):
    docx = Document(io.BytesIO(docx_bytes))
    pdf = fitz.open()
    pdf.new_page(width=595, height=842)
    pdf.new_page(width=595, height=842)
    page_index = 0
    cursor = 48
    font_size = 8
    line_height = 11
    max_width = 499
    list_number = 0

    for paragraph in docx.paragraphs:
        if paragraph.text.strip() == "BORANG I":
            continue
        if paragraph.text.strip().startswith("Bahagian 3"):
            page_index = 1
            cursor = 48
        text = paragraph.text.replace("\t", "    ").rstrip()
        if paragraph.style.name == "List Paragraph":
            list_number += 1
            text = f"{list_number}. {text}"
        else:
            list_number = 0
        if not text:
            cursor += 6
            continue

        paragraph_format = paragraph.paragraph_format
        left_indent = paragraph_format.left_indent.inches * 72 if paragraph_format.left_indent else 0
        first_line_indent = (
            paragraph_format.first_line_indent.inches * 72
            if paragraph_format.first_line_indent
            else 0
        )
        paragraph_before = paragraph_format.space_before.pt if paragraph_format.space_before else 0
        paragraph_after = paragraph_format.space_after.pt if paragraph_format.space_after else 3
        cursor += paragraph_before
        text = text.lstrip() if left_indent or first_line_indent else text
        words = text.split()
        lines = []
        current = ""
        available_width = max_width - left_indent
        for word in words:
            candidate = f"{current} {word}".strip()
            line_width = available_width - first_line_indent if not lines else available_width
            if current and fitz.get_text_length(candidate, fontname="helv", fontsize=font_size) > line_width:
                lines.append(current)
                current = word
            else:
                current = candidate
        if current:
            lines.append(current)

        is_heading = paragraph.style.name.startswith("Heading")
        is_bold = is_heading or any(run.bold for run in paragraph.runs if run.text)
        paragraph_font_size = 9 if is_heading else font_size
        paragraph_line_height = 13 if is_heading else line_height
        for line_index, line in enumerate(lines):
            line_indent = left_indent + (first_line_indent if line_index == 0 else 0)
            pdf[page_index].insert_text(
                (48 + line_indent, cursor),
                line,
                fontsize=paragraph_font_size,
                fontname="hebo" if is_bold else "helv",
                color=(0, 0, 0),
            )
            cursor += paragraph_line_height
        cursor += paragraph_after

    for page_number, page in enumerate(pdf, start=1):
        page_number_text = f"{page_number} / 2"
        page_number_width = fitz.get_text_length(page_number_text, fontname="helv", fontsize=7)
        page.insert_text(
            ((595 - page_number_width) / 2, 810),
            page_number_text,
            fontsize=7,
            fontname="helv",
            color=(0, 0, 0),
        )
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
        st.warning("PDF export requires LibreOffice to be installed on the host.")
    if st.button("Create another certificate"):
        st.session_state.screen = "report"
        st.rerun()
