import os
import re
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import List, Optional

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

import latex2mathml.converter
import lxml.etree as LET

app = FastAPI(title="FGEI Multi-Format Exam Generator")

if os.path.exists("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")


# Pydantic Schemas
class MCQQuestion(BaseModel):
    text: str
    is_urdu: Optional[bool] = False
    options: List[str]


class QuestionItem(BaseModel):
    main_question: str
    or_question: Optional[str] = None
    is_urdu: Optional[bool] = False
    marks: Optional[int] = 3


class ExamData(BaseModel):
    school_name: str
    exam_title: str
    subject: str
    grade: str
    time_allowed: str
    total_marks: int
    is_board_pattern: Optional[bool] = False
    is_urdu_paper: Optional[bool] = False
    mcqs: Optional[List[MCQQuestion]] = []
    short_questions: Optional[List[QuestionItem]] = []
    long_questions: Optional[List[QuestionItem]] = []


# MathML to OMML XSL Transform
MML2OMML_XSLT = """<xsl:stylesheet version="1.0" 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <xsl:output method="xml" omit-xml-declaration="yes"/>
    <xsl:template match="*">
        <m:oMath>
            <m:r>
                <m:t><xsl:value-of select="."/></m:t>
            </m:r>
        </m:oMath>
    </xsl:template>
</xsl:stylesheet>"""

xslt_transform = LET.XSLT(LET.fromstring(MML2OMML_XSLT))


def latex_to_omml(latex_str: str):
    """Converts raw LaTeX string into a Word OMML math object."""
    try:
        clean_latex = latex_str.strip()
        mathml_str = latex2mathml.converter.convert(clean_latex)
        mathml_tree = LET.fromstring(mathml_str)
        omml_tree = xslt_transform(mathml_tree)
        xml_bytes = LET.tostring(omml_tree)
        return parse_xml(xml_bytes)
    except Exception as e:
        print(f"Math Convert Error: {e}")
        return None


def add_text_with_math(paragraph, full_text: str):
    """Parses inline LaTeX/Math tags and embeds native Word math elements."""
    if not full_text:
        return

    # Strips out [Math: ...] wrappers if sent from UI
    text = re.sub(r'\[Math:\s*(.*?)\s*\]', r'\1', full_text)
    pattern = r'(\\[a-zA-Z]+(?:\{[^}]*\}|_[^{\s]+|\^[^{\s]+|_[{][^}]+[}]|\^[^{][^}]+[}])*)'
    tokens = re.split(pattern, text)

    for token in tokens:
        if not token:
            continue
        if token.startswith('\\'):
            omml_elem = latex_to_omml(token)
            if omml_elem is not None:
                paragraph._element.append(omml_elem)
            else:
                paragraph.add_run(token)
        else:
            paragraph.add_run(token)


def set_paragraph_rtl(paragraph, font_name="Jameel Noori Nastaliq", font_size=11, align_right=True):
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)
    
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align_right else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.line_spacing = 1.3

    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        rPr = run._element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:cs'), font_name)
        rFonts.set(qn('w:ascii'), font_name)
        rPr.append(rFonts)


def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)


def set_cell_margins(cell, top=30, bottom=30, left=60, right=60):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m_name, m_val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m_name}')
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


# ==========================================
# FORMAT 1: JUNIOR / PRIMARY CLASS FORMAT
# ==========================================
def build_fgei_standard_doc(data: ExamData, output_filename="fgei_junior_paper.docx"):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    is_pure_urdu = data.is_urdu_paper or (data.subject.strip() in ["اُردو", "اردو", "Urdu"])

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    logo_file = next((os.path.join(BASE_DIR, f) for f in ["logo.png", "logo.jpg", "logo.jpeg"] if os.path.exists(os.path.join(BASE_DIR, f))), None)
    if logo_file:
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(2)
        p_logo.add_run().add_picture(logo_file, width=Inches(0.9))

    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.paragraph_format.space_after = Pt(2)
    p_head.add_run(f"{data.school_name}\n").bold = True
    p_head.add_run(f"{data.exam_title}\n").bold = True
    if is_pure_urdu:
        set_paragraph_rtl(p_head, font_size=12, align_right=False)

    meta_table = doc.add_table(rows=3, cols=3)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(meta_table)
    col_widths = [Inches(2.5), Inches(2.5), Inches(2.5)]

    if is_pure_urdu:
        meta_table.cell(0, 0).paragraphs[0].add_run(f"مضمون: {data.subject}")
        meta_table.cell(0, 1).paragraphs[0].add_run(f"نمبر: {data.total_marks}")
        meta_table.cell(0, 2).paragraphs[0].add_run("رولنمبر: _________________")
        meta_table.cell(1, 0).paragraphs[0].add_run(f"کلاس: {data.grade}")
        meta_table.cell(1, 1).paragraphs[0].add_run(f"وقت: {data.time_allowed}")
        meta_table.cell(1, 2).paragraphs[0].add_run("تاریخ: __________________")
        meta_table.cell(2, 0).paragraphs[0].add_run("دستخط نگران: ____________")
    else:
        meta_table.cell(0, 0).paragraphs[0].add_run(f"Subject: {data.subject}")
        meta_table.cell(0, 1).paragraphs[0].add_run(f"Total Marks: {data.total_marks}")
        meta_table.cell(0, 2).paragraphs[0].add_run("Roll No: _________________")
        meta_table.cell(1, 0).paragraphs[0].add_run(f"Class: {data.grade}")
        meta_table.cell(1, 1).paragraphs[0].add_run(f"Time: {data.time_allowed}")
        meta_table.cell(1, 2).paragraphs[0].add_run("Date: __________________")
        meta_table.cell(2, 0).paragraphs[0].add_run("Invigilator Sign: ____________")

    for row in meta_table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = col_widths[idx]
            set_cell_margins(cell, top=30, bottom=30)
            if is_pure_urdu:
                set_paragraph_rtl(cell.paragraphs[0], font_size=11, align_right=True)

    if data.mcqs:
        p_sec_a = doc.add_paragraph()
        p_sec_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sec_a.paragraph_format.space_before = Pt(8)
        p_sec_a.paragraph_format.space_after = Pt(4)
        sec_title = "حصّہ اوّل  (معروضی)" if is_pure_urdu else "Section – A (Objective)"
        p_sec_a.add_run(sec_title).bold = True

        p_q1 = doc.add_paragraph()
        p_q1.paragraph_format.space_after = Pt(4)
        q1_title = f"سوال نمبر 1۔ درست جواب پر (✔) کا نشان لگائیں۔  ({len(data.mcqs)})" if is_pure_urdu else f"Q.1  Mark the Correct Answer Using (✔)  ({len(data.mcqs)})"
        p_q1.add_run(q1_title).bold = True

        labels = ["(a)", "(b)", "(c)", "(d)"]
        for i, mcq in enumerate(data.mcqs, start=1):
            p_mcq = doc.add_paragraph()
            p_mcq.paragraph_format.space_after = Pt(3)
            p_mcq.add_run(f"{i}.  ")
            add_text_with_math(p_mcq, mcq.text)

            p_mcq.add_run("\n  ")
            for opt_idx, opt_text in enumerate(mcq.options[:4]):
                p_mcq.add_run(f"{labels[opt_idx]} ")
                add_text_with_math(p_mcq, opt_text)
                p_mcq.add_run("   ")

    if data.short_questions:
        p_sec_b = doc.add_paragraph()
        p_sec_b.paragraph_format.space_before = Pt(10)
        p_sec_b.add_run("Section – B (Short Questions)").bold = True

        for i, sq in enumerate(data.short_questions, start=1):
            p_sq = doc.add_paragraph()
            p_sq.add_run(f"Q{i}.  ").bold = True
            add_text_with_math(p_sq, sq.main_question)

    if data.long_questions:
        p_sec_c = doc.add_paragraph()
        p_sec_c.paragraph_format.space_before = Pt(10)
        p_sec_c.add_run("Section – C (Detailed Questions)").bold = True

        for i, lq in enumerate(data.long_questions, start=1):
            p_lq = doc.add_paragraph()
            p_lq.add_run(f"Q{i+len(data.short_questions or [])}.  ").bold = True
            add_text_with_math(p_lq, lq.main_question)

    doc.save(output_filename)
    return output_filename


# ==========================================
# FORMAT 2: HIGH SCHOOL & BOARD FORMAT
# ==========================================
def add_roll_number_grid(doc):
    roll_table = doc.add_table(rows=11, cols=6)
    roll_table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    digits = ["⓪", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨"]

    for col_idx in range(6):
        cell = roll_table.cell(0, col_idx)
        cell.width = Inches(0.22)

    for row_idx, symbol in enumerate(digits, start=1):
        for col_idx in range(6):
            cell = roll_table.cell(row_idx, col_idx)
            cell.width = Inches(0.22)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(symbol)
            run.font.size = Pt(7.5)


def build_fgei_board_omr_doc(data: ExamData, output_filename="fgei_board_paper.docx"):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)

    # --- SECTION A: OBJECTIVE PAPER ---
    if data.mcqs:
        p_head = doc.add_paragraph()
        p_head.paragraph_format.space_after = Pt(2)
        r_sub = p_head.add_run(f"{data.subject} {data.grade}\n")
        r_sub.font.bold = True
        r_sub.font.size = Pt(12)
        p_head.add_run(f"{data.school_name}\n{data.exam_title}").font.size = Pt(10)

        add_roll_number_grid(doc)

        total_mcq_marks = len(data.mcqs)
        p_sec_a = doc.add_paragraph()
        p_sec_a.paragraph_format.space_before = Pt(6)
        p_sec_a.paragraph_format.space_after = Pt(4)
        p_sec_a.add_run(f"Section A ({total_mcq_marks} Marks)\n").font.bold = True
        p_sec_a.add_run(f"Q1. Fill the relevant bubble, against each question according to curriculum: ({total_mcq_marks})").font.bold = True

        mcq_table = doc.add_table(rows=len(data.mcqs) + 1, cols=7)
        mcq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        mcq_table.style = 'Table Grid'

        headers = ["S #", "Question", "(A)", "(B)", "(C)", "(D)", "Bubble Choice"]
        widths = [Inches(0.4), Inches(2.7), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.8)]

        hdr_row = mcq_table.rows[0]
        for idx, text in enumerate(headers):
            cell = hdr_row.cells[idx]
            cell.width = widths[idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(text).font.bold = True
            set_cell_margins(cell, top=30, bottom=30)

        for i, mcq in enumerate(data.mcqs, start=1):
            row_cells = mcq_table.rows[i].cells
            row_cells[0].paragraphs[0].add_run(f"({i})")
            row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Question text with math support
            add_text_with_math(row_cells[1].paragraphs[0], mcq.text)

            # Option choices with math support
            for opt_idx in range(4):
                p_opt = row_cells[opt_idx + 2].paragraphs[0]
                p_opt.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if opt_idx < len(mcq.options):
                    add_text_with_math(p_opt, mcq.options[opt_idx])

            p_b = row_cells[6].paragraphs[0]
            p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_b.add_run("○  ○  ○  ○").font.size = Pt(8)

            for idx in range(7):
                row_cells[idx].width = widths[idx]
                set_cell_margins(row_cells[idx], top=20, bottom=20)

        doc.add_page_break()

    # --- SECTION B & C: SUBJECTIVE PAPER ---
    p_sub_head = doc.add_paragraph()
    p_sub_head.paragraph_format.space_after = Pt(4)
    r_sub2 = p_sub_head.add_run(f"{data.subject} {data.grade}\n")
    r_sub2.font.bold = True
    r_sub2.font.size = Pt(12)

    total_sub_marks = data.total_marks - (len(data.mcqs) if data.mcqs else 0)
    p_sub_head.add_run(f"Time allowed: {data.time_allowed}\t\tTotal Marks: {total_sub_marks}").font.bold = True

    col_widths_5x1 = [Inches(0.5), Inches(2.9), Inches(0.6), Inches(2.9), Inches(0.6)]

    # Section B - Short Questions (5x1 Table)
    if data.short_questions:
        total_sq_marks = sum(q.marks if q.marks else 3 for q in data.short_questions)
        p_sec_b = doc.add_paragraph()
        p_sec_b.paragraph_format.space_before = Pt(8)
        p_sec_b.paragraph_format.space_after = Pt(4)
        p_sec_b.add_run(f"Section-B (Marks {total_sq_marks})\n").font.bold = True
        p_sec_b.add_run(f"Q.2 Attempt the following questions. ({len(data.short_questions)}x3={total_sq_marks})").font.bold = True

        sq_table = doc.add_table(rows=0, cols=5)
        sq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sq_table.style = 'Table Grid'

        for i, sq in enumerate(data.short_questions, start=1):
            row = sq_table.add_row()
            cells = row.cells

            for idx, cell in enumerate(cells):
                cell.width = col_widths_5x1[idx]
                set_cell_margins(cell)

            p_s = cells[0].paragraphs[0]
            p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_s.add_run(str(i)).font.bold = True

            add_text_with_math(cells[1].paragraphs[0], sq.main_question)

            if sq.or_question:
                p_or = cells[2].paragraphs[0]
                p_or.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_or.add_run("OR").font.bold = True

                add_text_with_math(cells[3].paragraphs[0], sq.or_question)
            else:
                cells[1].merge(cells[3])

            p_m = cells[4].paragraphs[0]
            p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_m.add_run(f"({sq.marks or 3})")

    # Section C - Long Questions (5x1 Table)
    if data.long_questions:
        total_lq_marks = sum(q.marks if q.marks else 6 for q in data.long_questions)
        p_sec_c = doc.add_paragraph()
        p_sec_c.paragraph_format.space_before = Pt(10)
        p_sec_c.paragraph_format.space_after = Pt(4)
        p_sec_c.add_run(f"Section-C (Marks {total_lq_marks})\n").font.bold = True
        p_sec_c.add_run(f"Note: Attempt all questions. Marks of all questions are equal. ({len(data.long_questions)}x6={total_lq_marks})").font.bold = True

        lq_table = doc.add_table(rows=0, cols=5)
        lq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        lq_table.style = 'Table Grid'

        for i, lq in enumerate(data.long_questions, start=1):
            row = lq_table.add_row()
            cells = row.cells

            for idx, cell in enumerate(cells):
                cell.width = col_widths_5x1[idx]
                set_cell_margins(cell)

            p_s = cells[0].paragraphs[0]
            p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_s.add_run(f"Q{i+2}").font.bold = True

            add_text_with_math(cells[1].paragraphs[0], lq.main_question)

            if lq.or_question:
                p_or = cells[2].paragraphs[0]
                p_or.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_or.add_run("OR").font.bold = True

                add_text_with_math(cells[3].paragraphs[0], lq.or_question)
            else:
                cells[1].merge(cells[3])

            p_m = cells[4].paragraphs[0]
            p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_m.add_run(f"({lq.marks or 6})")

    doc.save(output_filename)
    return output_filename


def is_board_class(grade_str: str) -> bool:
    normalized = grade_str.lower()
    high_school_keywords = ["9", "10", "11", "12", "ssc", "hssc", "matric", "first year", "second year", "1st year", "2nd year"]
    return any(kw in normalized for kw in high_school_keywords)


@app.get("/")
def read_root():
    return FileResponse("static/index.html")


@app.post("/generate-paper")
def generate_paper(data: ExamData):
    try:
        if data.is_board_pattern or is_board_class(data.grade):
            output_file = build_fgei_board_omr_doc(data)
        else:
            output_file = build_fgei_standard_doc(data)

        return FileResponse(
            path=output_file,
            filename=f"{data.subject}_{data.grade}_Exam.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))