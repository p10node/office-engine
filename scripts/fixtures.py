"""Generate sample + large Office files for the playground and tests.

Usage:  python3 -m venv .venv && .venv/bin/pip install python-docx openpyxl python-pptx Pillow
        .venv/bin/python scripts/fixtures.py [--large]
"""
import datetime
import io
import os
import random
import sys

from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(__file__), '..', 'playground', 'public', 'samples')
os.makedirs(OUT, exist_ok=True)
random.seed(7)

LOREM = ('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et '
         'dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex '
         'ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore.').split(' ')


def lorem(n):
    return ' '.join(random.choice(LOREM) for _ in range(n)).capitalize() + '.'


def png(w=640, h=320, hue=0):
    img = Image.new('RGB', (w, h))
    d = ImageDraw.Draw(img)
    for x in range(w):
        t = x / w
        d.line([(x, 0), (x, h)], fill=(int(40 + 180 * t), int(90 + 60 * (1 - t)), int(200 - 120 * t + hue) % 256))
    d.ellipse([w * 0.35, h * 0.2, w * 0.65, h * 0.8], fill=(255, 255, 255))
    d.text((20, 20), 'office-engine', fill=(255, 255, 255))
    b = io.BytesIO()
    img.save(b, 'PNG')
    b.seek(0)
    return b


# ---------------------------------------------------------------- DOCX
def add_page_field(paragraph):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    run = paragraph.add_run()
    for kind, text in (('begin', None), (None, ' PAGE '), ('separate', None), ('t', '1'), ('end', None)):
        if kind in ('begin', 'separate', 'end'):
            el = OxmlElement('w:fldChar')
            el.set(qn('w:fldCharType'), kind)
            run._r.append(el)
        elif kind == 't':
            t = OxmlElement('w:t')
            t.text = text
            run._r.append(t)
        else:
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = text
            run._r.append(it)


def add_hyperlink(paragraph, url, text):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    h = OxmlElement('w:hyperlink')
    h.set(qn('r:id'), r_id)
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    st = OxmlElement('w:rStyle')
    st.set(qn('w:val'), 'Hyperlink')
    rpr.append(st)
    col = OxmlElement('w:color')
    col.set(qn('w:val'), '0563C1')
    rpr.append(col)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rpr.append(u)
    r.append(rpr)
    t = OxmlElement('w:t')
    t.text = text
    r.append(t)
    h.append(r)
    paragraph._p.append(h)


def shade(cell, hex_fill):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_fill)
    tcPr.append(shd)


def make_docx(path, paragraphs=0):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.shared import Inches, Pt, RGBColor

    doc = Document()
    sec = doc.sections[0]
    hp = sec.header.paragraphs[0]
    hp.text = 'office-engine sample — header'
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.add_run('Page ')
    add_page_field(fp)

    doc.add_heading('Quarterly Report', 0)
    p = doc.add_paragraph('This document shows ')
    p.add_run('bold').bold = True
    p.add_run(', ')
    p.add_run('italic').italic = True
    p.add_run(', ')
    r = p.add_run('underlined red')
    r.underline = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    p.add_run(' and ')
    r = p.add_run('large')
    r.font.size = Pt(18)
    p.add_run(' text. Visit ')
    add_hyperlink(p, 'https://example.com', 'example.com')
    p.add_run('.')

    doc.add_heading('1. Highlights', level=1)
    for t in ['Revenue grew 24% year over year', 'Two new regions launched', 'Customer churn fell to 3.1%']:
        doc.add_paragraph(t, style='List Bullet')
    doc.add_heading('2. Next steps', level=1)
    for t in ['Hire the platform team', 'Ship the viewer engine', 'Expand to APAC']:
        doc.add_paragraph(t, style='List Number')

    doc.add_heading('Figures', level=2)
    doc.add_picture(png(), width=Inches(4.5))
    cap = doc.add_paragraph('Figure 1 — generated gradient image')
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading('Table', level=2)
    table = doc.add_table(rows=5, cols=4)
    table.style = 'Table Grid'
    hdr = ['Region', 'Q1', 'Q2', 'Total']
    for i, h in enumerate(hdr):
        table.rows[0].cells[i].text = h
        table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        shade(table.rows[0].cells[i], 'D9E2F3')
    data = [('North', 120, 140), ('South', 98, 110), ('East', 132, 150)]
    for ri, (name, a, b) in enumerate(data, start=1):
        cells = table.rows[ri].cells
        cells[0].text = name
        cells[1].text = str(a)
        cells[2].text = str(b)
        cells[3].text = str(a + b)
    merged = table.rows[4].cells[0].merge(table.rows[4].cells[2])
    merged.text = 'Merged footer cell (3 columns)'
    table.rows[4].cells[3].text = '750'

    para = doc.add_paragraph(lorem(60))
    para.paragraph_format.first_line_indent = Inches(0.4)
    q = doc.add_paragraph(lorem(30))
    q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    q.paragraph_format.left_indent = Inches(0.6)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_heading('Appendix', level=1)

    for i in range(paragraphs or 40):
        if i % 50 == 0 and i:
            doc.add_heading(f'Section {i // 50}', level=1)
        if i % 200 == 120:
            t = doc.add_table(rows=6, cols=3)
            t.style = 'Table Grid'
            for r in range(6):
                for c in range(3):
                    t.cell(r, c).text = f'R{r + 1}C{c + 1}'
            continue
        doc.add_paragraph(f'{i + 1}. ' + lorem(random.randint(20, 70)))
    doc.save(path)


# ---------------------------------------------------------------- XLSX
def make_xlsx(path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sales'
    ws.merge_cells('A1:F1')
    ws['A1'] = 'Sales by region — FY2026'
    ws['A1'].font = Font(size=16, bold=True, color='1F3864')
    ws['A1'].alignment = Alignment(horizontal='center')
    head = ['Region', 'Date', 'Units', 'Price', 'Revenue', 'Margin']
    thin = Side(style='thin', color='8EA9DB')
    for i, h in enumerate(head, start=1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='4472C4')
        c.alignment = Alignment(horizontal='center')
        c.border = Border(bottom=thin)
    regions = ['North', 'South', 'East', 'West', 'Central']
    for r in range(3, 63):
        ws.cell(row=r, column=1, value=random.choice(regions))
        d = ws.cell(row=r, column=2, value=datetime.date(2026, 1, 1) + datetime.timedelta(days=r * 5))
        d.number_format = 'yyyy-mm-dd'
        ws.cell(row=r, column=3, value=random.randint(5, 500))
        p = ws.cell(row=r, column=4, value=round(random.uniform(1, 99), 2))
        p.number_format = '"$"#,##0.00'
        rev = ws.cell(row=r, column=5, value=f'=C{r}*D{r}')
        rev.number_format = '#,##0.00'
        m = ws.cell(row=r, column=6, value=round(random.uniform(-0.1, 0.6), 4))
        m.number_format = '0.0%'
        if r % 2:
            for c in range(1, 7):
                ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor='D9E1F2')
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['E'].width = 14
    ws.freeze_panes = 'B3'
    ws['H2'] = 'Note: Revenue is a formula; cached values are absent in openpyxl output.'
    ws['H3'] = 1234567.891
    ws['H4'] = -42.5
    ws['H4'].number_format = '#,##0.00;[Red](#,##0.00)'
    ws['H5'] = 0.000123
    ws['H6'] = True
    ws['H7'] = 45291.75
    ws['H7'].number_format = 'dddd, mmmm d, yyyy h:mm AM/PM'

    ws2 = wb.create_sheet('Notes')
    ws2.column_dimensions['A'].width = 50
    ws2['A1'] = 'Wrapped text cell. ' + lorem(40)
    ws2['A1'].alignment = Alignment(wrap_text=True, vertical='top')
    ws2.row_dimensions[1].height = 120
    ws2['B1'] = 'Italic'
    ws2['B1'].font = Font(italic=True, underline='single')

    ws3 = wb.create_sheet('Data')
    ws3.append(['ID', 'Name', 'Score', 'Ratio'])
    for i in range(1, 5001):
        ws3.append([i, f'Item {i}', random.randint(0, 1000), random.random()])
    wb.save(path)


def make_large_xlsx(path, rows=300_000, cols=12):
    from openpyxl import Workbook
    wb = Workbook(write_only=True)
    ws = wb.create_sheet('Big')
    ws.append([f'Col {c + 1}' for c in range(cols)])
    words = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta']
    for r in range(rows):
        row = [r + 1]
        for c in range(1, cols):
            row.append(random.choice(words) + str(r % 97) if c % 3 == 0 else round(random.uniform(0, 10000), 2))
        ws.append(row)
    wb.save(path)


# ---------------------------------------------------------------- PPTX
def make_pptx(path, extra=0):
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = 'office-engine'
    s.placeholders[1].text = 'Progressive Office viewer for the browser'

    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = 'Why it is fast'
    body = s.placeholders[1].text_frame
    body.text = 'Lazy ZIP: only the central directory is read up front'
    for t, lvl in [('Entries are inflated on demand with DecompressionStream', 1), ('Streaming XML parser in a Web Worker', 0),
                   ('Blocks / rows / slides arrive progressively', 1), ('Virtualized React views', 0)]:
        p = body.add_paragraph()
        p.text = t
        p.level = lvl

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = 'Shapes & images'
    colors = [RGBColor(0x44, 0x72, 0xC4), RGBColor(0xED, 0x7D, 0x31), RGBColor(0x70, 0xAD, 0x47)]
    for i, kind in enumerate([MSO_SHAPE.ROUNDED_RECTANGLE, MSO_SHAPE.OVAL, MSO_SHAPE.RIGHT_ARROW]):
        sh = s.shapes.add_shape(kind, Inches(0.8 + i * 2.6), Inches(2), Inches(2.2), Inches(1.4))
        sh.fill.solid()
        sh.fill.fore_color.rgb = colors[i]
        sh.text_frame.text = ['Rounded', 'Oval', 'Arrow'][i]
    s.shapes.add_picture(png(800, 450, 60), Inches(8.6), Inches(1.8), width=Inches(4))
    tb = s.shapes.add_textbox(Inches(0.8), Inches(4.2), Inches(7), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = 'A free text box with '
    run.font.size = Pt(20)
    r2 = tf.paragraphs[0].add_run()
    r2.text = 'colored bold text'
    r2.font.bold = True
    r2.font.size = Pt(20)
    r2.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = 'Table'
    rows, cols = 5, 4
    t = s.shapes.add_table(rows, cols, Inches(0.8), Inches(1.8), Inches(11.5), Inches(3.5)).table
    for c, h in enumerate(['Metric', '2024', '2025', '2026']):
        t.cell(0, c).text = h
    for r in range(1, rows):
        t.cell(r, 0).text = ['Users', 'Revenue', 'NPS', 'Latency'][r - 1]
        for c in range(1, cols):
            t.cell(r, c).text = str(random.randint(10, 999))

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = 'Chart (placeholder)'
    cd = CategoryChartData()
    cd.categories = ['Q1', 'Q2', 'Q3', 'Q4']
    cd.add_series('Sales', (19.2, 21.4, 16.7, 25.0))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.8), Inches(8), Inches(5), cd)

    for i in range(extra):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = f'Slide {i + 6}'
        tf = s.placeholders[1].text_frame
        tf.text = lorem(8)
        for _ in range(3):
            tf.add_paragraph().text = lorem(10)
    prs.save(path)


if __name__ == '__main__':
    make_docx(os.path.join(OUT, 'sample.docx'))
    make_xlsx(os.path.join(OUT, 'sample.xlsx'))
    make_pptx(os.path.join(OUT, 'sample.pptx'))
    if '--large' in sys.argv:
        make_docx(os.path.join(OUT, 'large-doc.docx'), paragraphs=20000)
        make_large_xlsx(os.path.join(OUT, 'large-sheet.xlsx'))
        make_pptx(os.path.join(OUT, 'large-deck.pptx'), extra=300)
    print('ok')
