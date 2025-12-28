import docx
import os
from docx import Document
from docx.shared import Pt, Inches, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def create_final_custom_height_replica():
    # 1. Document Setup
    doc = Document()
    section = doc.sections[0]
    section.page_height = Mm(297) 
    section.page_width = Mm(210)  
    
    # Margins (0.25 inch)
    m = Inches(0.25)
    section.top_margin = m
    section.bottom_margin = m
    section.left_margin = m
    section.right_margin = m

    # 2. Helper Functions
    def set_font(run, size=10.5, bold=False, underline=False, color=None):
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.bold = bold
        run.underline = underline
        if color: run.font.color.rgb = color

    def add_centered_line(text, bold=True, size=11):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        set_font(p.add_run(text), size=size, bold=bold)
        return p

    def set_cell_borders(cell):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = tcPr.find(qn('w:tcBorders'))
        if tcBorders is None:
            tcBorders = OxmlElement('w:tcBorders')
            tcPr.append(tcBorders)
        for border in ['top', 'left', 'bottom', 'right']:
            element = OxmlElement(f'w:{border}')
            element.set(qn('w:val'), 'single')
            element.set(qn('w:sz'), '4') 
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), 'auto')
            tcBorders.append(element)

    def fill_cell(cell, text=None, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, color=None, underline=False, vertical_align=WD_CELL_VERTICAL_ALIGNMENT.TOP):
        cell.text = "" 
        cell.vertical_alignment = vertical_align
        p = cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        if text:
            set_font(p.add_run(text), bold=bold, color=color, underline=underline)
        set_cell_borders(cell)
        return p

    # 3. Header Content
    add_centered_line("FORM 'A'", size=12)
    add_centered_line("MEDIATION APPLICATION FORM", size=12)
    add_centered_line("[REFER RULE 3(1)]", size=11)
    add_centered_line("Mumbai District Legal Services Authority", bold=False, size=12)
    p = add_centered_line("City Civil Court, Mumbai", bold=False, size=12)
    p.paragraph_format.space_after = Pt(6)

    # 4. Table Setup
    table = doc.add_table(rows=0, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 
    col_widths = [Inches(0.37), Inches(1.2), Inches(5.68)]

    def add_row(height):
        row = table.add_row()
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.height = Inches(height)
        for i, width in enumerate(col_widths):
            row.cells[i].width = width
        return row

    # --- ROWS GENERATION ---

    # Row 1
    row = add_row(0.40)
    row.cells[0].merge(row.cells[2])
    fill_cell(row.cells[0], "DETAILS OF PARTIES:", bold=True)

    # Row 2
    row = add_row(0.40)
    fill_cell(row.cells[0], "1", align=WD_ALIGN_PARAGRAPH.CENTER)
    fill_cell(row.cells[1], "Name of\nApplicant", bold=True)
    fill_cell(row.cells[2], "{{client_name}}", bold=True)

    # Row 3
    row = add_row(0.35)
    row.cells[1].merge(row.cells[2])
    set_cell_borders(row.cells[0])
    fill_cell(row.cells[1], "Address and contact details of Applicant", bold=True)

    # Row 4 (Address 1)
    row = add_row(1.1)
    fill_cell(row.cells[0], "1", align=WD_ALIGN_PARAGRAPH.CENTER)
    fill_cell(row.cells[1], "Address", bold=True)
    
    cell = row.cells[2]
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("REGISTERED ADDRESS:"), bold=True)
    
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("{{branch_address}}"), bold=False)
    
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    p.add_run(" ")
    
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("CORRESPONDENCE BRANCH ADDRESS:"), bold=True)
    
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("{{branch_address}}"), bold=False)
    set_cell_borders(cell)

    # Rows 5, 6, 7
    row = add_row(0.35)
    set_cell_borders(row.cells[0])
    fill_cell(row.cells[1], "Telephone No.", bold=True)
    fill_cell(row.cells[2], "{{mobile}}", bold=True)

    row = add_row(0.35)
    set_cell_borders(row.cells[0])
    fill_cell(row.cells[1], "Mobile No.", bold=True)
    set_cell_borders(row.cells[2])

    row = add_row(0.35)
    set_cell_borders(row.cells[0])
    fill_cell(row.cells[1], "Email ID", bold=True)
    fill_cell(row.cells[2], "info@kslegal.co.in", color=RGBColor(0,0,255), underline=True)

    # Row 8
    row = add_row(0.35)
    fill_cell(row.cells[0], "2", align=WD_ALIGN_PARAGRAPH.CENTER)
    row.cells[1].merge(row.cells[2])
    fill_cell(row.cells[1], "Name, Address and Contact details of Opposite Party:", bold=True)

    # Row 9
    row = add_row(0.35)
    set_cell_borders(row.cells[0])
    row.cells[1].merge(row.cells[2])
    fill_cell(row.cells[1], "Address and contact details of Defendant/s", bold=True)

    # Row 10
    row = add_row(0.35)
    set_cell_borders(row.cells[0])
    fill_cell(row.cells[1], "Name", bold=True)
    fill_cell(row.cells[2], "{{customer_name}}", bold=True)

    # Row 11 (Address 2)
    row = add_row(1.5)
    set_cell_borders(row.cells[0])
    fill_cell(row.cells[1], "Address", bold=True)

    cell = row.cells[2]
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("REGISTERED ADDRESS:"), bold=True)
    
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("{% if address1 %}{{address1}}{% else %}____________{%\n endif %}"), bold=False)
    
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    p.add_run(" ")

    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("CORRESPONDENCE ADDRESS:"), bold=True)
    
    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run("{% if address1 %}{{address1}}{% else %}____________{%\n endif %}"), bold=False)
    set_cell_borders(cell)

    # Rows 12, 13, 14
    for label in ["Telephone No.", "Mobile No.", "Email ID"]:
        row = add_row(0.35)
        set_cell_borders(row.cells[0])
        fill_cell(row.cells[1], label, bold=True)
        set_cell_borders(row.cells[2])

    # Row 15
    row = add_row(0.35)
    row.cells[0].merge(row.cells[2])
    fill_cell(row.cells[0], "DETAILS OF DISPUTE:", bold=True)

    # Row 16
    row = add_row(0.40)
    row.cells[0].merge(row.cells[2])
    cell = row.cells[0]
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    set_font(p.add_run("THE COMM. COURTS (PRE-INSTITUTION.........SETTLEMENT) RULES,2018"), bold=True, underline=True)
    set_cell_borders(cell)

    # Row 17
    row = add_row(0.22)
    set_cell_borders(row.cells[0])
    row.cells[1].merge(row.cells[2])
    fill_cell(row.cells[1], "Nature of disputes as per section 2(1)(c) of the Commercial Courts Act, 2015 (4 of 2016):", bold=True)

    # 5. Save File
    file_name = "Form_A_Mediation_Replica.docx"
    file_path = os.path.join(os.getcwd(), file_name)
    doc.save(file_path)
    print(f"File created successfully: {file_path}")

if __name__ == "__main__":
    create_final_custom_height_replica()