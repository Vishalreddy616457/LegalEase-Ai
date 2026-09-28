from io import BytesIO
from fpdf import FPDF
from docx import Document
from docx.shared import Pt

def format_docx(text, doc_type):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    doc.add_heading(doc_type, level=0)
    for line in text.splitlines():
        line = line.strip()
        if not line:
            doc.add_paragraph("")
        elif line[:2].isdigit() and "." in line[:4]:
            doc.add_heading(line, level=1)
        else:
            doc.add_paragraph(line)

    doc.sections[0].footer.paragraphs[0].text = (
        "LegalEase - Demo document. Not legal advice."
    )

    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def format_pdf(text, doc_type):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()
    pdf.set_font("Times", "B", 16)
    pdf.cell(0, 10, doc_type, ln=True, align="C")
    pdf.ln(5)
    pdf.set_font("Times", size=12)

    for line in text.splitlines():
        if not line.strip():
            pdf.ln(5)
        else:
            pdf.multi_cell(0, 7, line)

    pdf.set_y(-15)
    pdf.set_font("Times", size=8)
    pdf.cell(0, 10, "LegalEase - Demo document. Not legal advice.", align="C")
    return bytes(pdf.output(dest="S"))
