"""
Generates the OpenDocument presentation fixtures of test/fixtures/legacy/ with odfpy:
  sample.odp   the demo deck: master with a band, footer and slide numbers; title slide, outline with
               formatting and a link, shapes (preset, custom path with equations, rotated star, line),
               a picture, a table, numbered lists, a slide background, speaker notes, a hidden slide
  sample.otp   the same deck as a template
  sample.fodp  the same deck as flat XML (the picture inline as base64)

Usage (see odp.md): python -I scripts/legacy-samples/odp.py test/fixtures/legacy
"""
import base64
import re
import struct
import sys
import zlib
from pathlib import Path

from odf import draw, presentation, style, table, text
from odf.namespaces import PRESENTATIONNS
from odf.opendocument import OpenDocumentPresentation


def png(w, h, pixel):
    """A tiny PNG (RGB) from a pixel function."""
    raw = b''.join(b'\0' + b''.join(bytes(pixel(x, y)) for x in range(w)) for y in range(h))
    chunk = lambda t, d: struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xFFFFFFFF)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw, 9)) + chunk(b'IEND', b'')


def landscape(x, y):
    """sky, sun and hills"""
    if (x - 120) ** 2 + (y - 28) ** 2 < 16 ** 2:
        return (250, 204, 21)
    if y > 70 + 10 * ((x % 80) / 80 - 0.5) ** 2 * 4:
        return (22, 163, 74)
    return (96, 165 + y // 2, 250)


def setp(el, name, value):
    el.setAttrNS(PRESENTATIONNS, name, value)
    return el


def build(template=False):
    doc = OpenDocumentPresentation()
    if template:
        doc.mimetype = 'application/vnd.oasis.opendocument.presentation-template'
    S = doc.styles
    A = doc.automaticstyles
    MS = doc.masterstyles

    # ---- named styles, fill resources
    S.addElement(draw.Gradient(name='Band', style='linear', startcolor='#1d4ed8', endcolor='#60a5fa', angle='90deg', border='0%'))
    S.addElement(draw.Marker(name='Arrow', viewbox='0 0 20 30', d='M10 0l-10 30h20z'))
    d = style.DefaultStyle(family='graphic')
    d.addElement(style.GraphicProperties(strokecolor='#3465a4', fillcolor='#729fcf'))
    d.addElement(style.TextProperties(fontfamily='Liberation Sans', fontsize='18pt'))
    S.addElement(d)
    std = style.Style(name='standard', family='graphic')
    std.addElement(style.GraphicProperties(stroke='solid', strokewidth='0cm', strokecolor='#3465a4', fill='solid', fillcolor='#729fcf',
                                           paddingtop='0.125cm', paddingbottom='0.125cm', paddingleft='0.25cm', paddingright='0.25cm', wrapoption='wrap'))
    std.addElement(style.ParagraphProperties(marginleft='0cm', marginright='0cm', margintop='0cm', marginbottom='0cm', textindent='0cm'))
    std.addElement(style.TextProperties(fontfamily='Liberation Sans', fontsize='18pt', color='#1f2937'))
    S.addElement(std)

    def pres(name, parent=None, graphic=None, para=None, txt=None, liststyle=None):
        st = style.Style(name=name, family='presentation', **({'parentstylename': parent} if parent else {}))
        if graphic is not None or liststyle is not None:
            gp = style.GraphicProperties(**(graphic or {}))
            if liststyle is not None:
                gp.addElement(liststyle)
            st.addElement(gp)
        if para:
            st.addElement(style.ParagraphProperties(**para))
        if txt:
            st.addElement(style.TextProperties(**txt))
        S.addElement(st)
        return st

    outline_list = text.ListStyle(name='Default-outline1')
    for level, (char, before, size) in enumerate([('●', '0.3cm', '45%'), ('–', '1.5cm', '75%'), ('●', '2.7cm', '45%')], 1):
        b = text.ListLevelStyleBullet(level=level, bulletchar=char)
        b.addElement(style.ListLevelProperties(spacebefore=before, minlabelwidth='0.9cm'))
        b.addElement(style.TextProperties(fontfamily='OpenSymbol', fontsize=size, color='#2563eb'))
        outline_list.addElement(b)

    pres('Default-background', graphic={'stroke': 'none', 'fill': 'none'})
    pres('Default-backgroundobjects', graphic={'shadow': 'hidden'})
    title_style = pres('Default-title', graphic={'stroke': 'none', 'fill': 'none', 'textareaverticalalign': 'middle', 'autogrowheight': 'false'},
         para={'textalign': 'center'}, txt={'fontfamily': 'Liberation Sans', 'fontsize': '40pt', 'fontweight': 'bold', 'color': '#111827'})
    pres('Default-subtitle', graphic={'stroke': 'none', 'fill': 'none', 'textareaverticalalign': 'middle'},
         para={'textalign': 'center'}, txt={'fontfamily': 'Liberation Sans', 'fontsize': '24pt', 'color': '#4b5563'})
    outline_style = pres('Default-outline1', graphic={'stroke': 'none', 'fill': 'none', 'autogrowheight': 'false'}, liststyle=outline_list,
         para={'margintop': '0.4cm', 'marginbottom': '0cm'}, txt={'fontfamily': 'Liberation Sans', 'fontsize': '28pt', 'color': '#1f2937'})
    pres('Default-outline2', parent='Default-outline1', para={'margintop': '0.3cm'}, txt={'fontsize': '24pt'})
    pres('Default-outline3', parent='Default-outline2', para={'margintop': '0.2cm'}, txt={'fontsize': '20pt'})
    pres('Default-notes', graphic={'stroke': 'none', 'fill': 'none'}, txt={'fontfamily': 'Liberation Sans', 'fontsize': '12pt'})

    # ---- master page: band, title and outline areas, date, footer, slide number
    pl = style.PageLayout(name='PM1')
    pl.addElement(style.PageLayoutProperties(pagewidth='28cm', pageheight='15.75cm', printorientation='landscape'))
    A.addElement(pl)
    mdp = style.Style(name='Mdp1', family='drawing-page')
    mdp.addElement(style.DrawingPageProperties(fill='solid', fillcolor='#ffffff'))
    A.addElement(mdp)
    band = style.Style(name='Mgr1', family='graphic', parentstylename='standard')
    band.addElement(style.GraphicProperties(stroke='none', fill='gradient', fillgradientname='Band'))
    A.addElement(band)
    foot = style.Style(name='Mpr1', family='presentation', parentstylename='Default-backgroundobjects')
    foot.addElement(style.GraphicProperties(stroke='none', fill='none', textareaverticalalign='middle'))
    foot.addElement(style.TextProperties(fontfamily='Liberation Sans', fontsize='12pt', color='#ffffff'))
    A.addElement(foot)
    right = style.Style(name='MP1', family='paragraph')
    right.addElement(style.ParagraphProperties(textalign='end'))
    A.addElement(right)

    master = style.MasterPage(name='Default', pagelayoutname=pl, stylename=mdp)
    MS.addElement(master)
    master.addElement(draw.Rect(stylename=band, x='0cm', y='14.9cm', width='28cm', height='0.85cm'))
    title_ph = draw.Frame(stylename=title_style, x='1.4cm', y='0.6cm', width='25.2cm', height='2.6cm')
    master.addElement(setp(setp(title_ph, 'class', 'title'), 'placeholder', 'true'))
    title_ph.addElement(draw.TextBox())
    outline_ph = draw.Frame(stylename=outline_style, x='1.4cm', y='3.6cm', width='25.2cm', height='10.8cm')
    master.addElement(setp(setp(outline_ph, 'class', 'outline'), 'placeholder', 'true'))
    outline_ph.addElement(draw.TextBox())
    f = setp(draw.Frame(stylename=foot, x='9.5cm', y='14.9cm', width='9cm', height='0.85cm'), 'class', 'footer')
    tb = draw.TextBox()
    p = text.P()
    p.addElement(presentation.Footer())
    tb.addElement(p)
    f.addElement(tb)
    master.addElement(f)
    n = setp(draw.Frame(stylename=foot, x='22cm', y='14.9cm', width='5cm', height='0.85cm'), 'class', 'page-number')
    tb = draw.TextBox()
    p = text.P(stylename=right)
    p.addElement(text.PageNumber(selectpage='current'))
    tb.addElement(p)
    n.addElement(tb)
    master.addElement(n)

    # ---- automatic styles of the slides
    def auto(name, family, parent=None, graphic=None, para=None, txt=None, page=None):
        st = style.Style(name=name, family=family, **({'parentstylename': parent} if parent else {}))
        if graphic:
            st.addElement(style.GraphicProperties(**graphic))
        if page:
            st.addElement(style.DrawingPageProperties(**page))
        if para:
            st.addElement(style.ParagraphProperties(**para))
        if txt:
            st.addElement(style.TextProperties(**txt))
        A.addElement(st)
        return st

    dp1 = auto('dp1', 'drawing-page', page={'displayfooter': 'true', 'displaypagenumber': 'true', 'displaydatetime': 'false'})
    dp2 = auto('dp2', 'drawing-page', page={'fill': 'solid', 'fillcolor': '#f1f5f9', 'displayfooter': 'true', 'displaypagenumber': 'true', 'displaydatetime': 'false'})
    dp3 = auto('dp3', 'drawing-page', page={'visibility': 'hidden', 'displayfooter': 'false', 'displaypagenumber': 'false'})
    pr_title = auto('pr1', 'presentation', 'Default-title')
    pr_sub = auto('pr2', 'presentation', 'Default-subtitle')
    pr_out = auto('pr3', 'presentation', 'Default-outline1')
    pr_notes = auto('pr4', 'presentation', 'Default-notes')
    gr_blue = auto('gr1', 'graphic', 'standard', graphic={'fill': 'solid', 'fillcolor': '#2563eb', 'stroke': 'none', 'textareaverticalalign': 'middle'})
    gr_green = auto('gr2', 'graphic', 'standard', graphic={'fill': 'solid', 'fillcolor': '#16a34a', 'stroke': 'solid', 'strokecolor': '#14532d', 'strokewidth': '0.05cm', 'textareaverticalalign': 'middle'})
    gr_amber = auto('gr3', 'graphic', 'standard', graphic={'fill': 'solid', 'fillcolor': '#f59e0b', 'stroke': 'none', 'textareaverticalalign': 'middle'})
    gr_text = auto('gr4', 'graphic', 'standard', graphic={'fill': 'none', 'stroke': 'none', 'autogrowheight': 'true'})
    gr_pic = auto('gr5', 'graphic', graphic={'fill': 'none', 'stroke': 'solid', 'strokecolor': '#9ca3af', 'strokewidth': '0.03cm'})
    gr_star = auto('gr6', 'graphic', 'standard', graphic={'fill': 'gradient', 'fillgradientname': 'Band', 'stroke': 'none'})
    gr_line = auto('gr7', 'graphic', 'standard', graphic={'stroke': 'solid', 'strokecolor': '#dc2626', 'strokewidth': '0.08cm', 'markerend': 'Arrow', 'markerendwidth': '0.4cm'})
    p_center = auto('P1', 'paragraph', para={'textalign': 'center'}, txt={'color': '#ffffff', 'fontsize': '20pt', 'fontweight': 'bold'})
    t_bold = auto('T1', 'text', txt={'fontweight': 'bold'})
    t_italic = auto('T2', 'text', txt={'fontstyle': 'italic'})
    t_under = auto('T3', 'text', txt={'textunderlinestyle': 'solid', 'textunderlinewidth': 'auto', 'textunderlinecolor': 'font-color'})
    t_red = auto('T4', 'text', txt={'color': '#dc2626', 'fontweight': 'bold'})
    t_code = auto('T5', 'text', txt={'fontfamily': 'Liberation Mono', 'backgroundcolor': '#fef3c7'})
    t_sup = auto('T6', 'text', txt={'textposition': 'super 58%'})
    co1 = auto('co1', 'table-column')
    co1.addElement(style.TableColumnProperties(columnwidth='6cm'))
    co2 = auto('co2', 'table-column')
    co2.addElement(style.TableColumnProperties(columnwidth='4cm'))
    ro1 = auto('ro1', 'table-row')
    ro1.addElement(style.TableRowProperties(rowheight='1.4cm'))
    ce_head = style.Style(name='ce1', family='table-cell')
    ce_head.addElement(style.GraphicProperties(fill='solid', fillcolor='#1d4ed8', textareaverticalalign='middle'))
    ce_head.addElement(style.ParagraphProperties(border='0.03cm solid #ffffff'))
    ce_head.addElement(style.TextProperties(color='#ffffff', fontweight='bold', fontsize='18pt'))
    A.addElement(ce_head)
    ce_body = style.Style(name='ce2', family='table-cell')
    ce_body.addElement(style.GraphicProperties(fill='solid', fillcolor='#e5e7eb', textareaverticalalign='middle'))
    ce_body.addElement(style.ParagraphProperties(border='0.03cm solid #ffffff'))
    ce_body.addElement(style.TextProperties(color='#111827', fontsize='18pt'))
    A.addElement(ce_body)
    numbers = text.ListStyle(name='L1')
    for level, (fmt, suffix, before) in enumerate([('1', '.', '0.3cm'), ('a', ')', '1.5cm')], 1):
        b = text.ListLevelStyleNumber(level=level, numformat=fmt, numsuffix=suffix)
        b.addElement(style.ListLevelProperties(spacebefore=before, minlabelwidth='0.9cm'))
        numbers.addElement(b)
    A.addElement(numbers)

    footer = presentation.FooterDecl(name='ftr1')
    footer.addText('office-engine sample deck')
    doc.presentation.addElement(footer)

    def page(name, dp):
        pg = draw.Page(name=name, masterpagename=master, stylename=dp)
        pg.setAttrNS(PRESENTATIONNS, 'use-footer-name', 'ftr1')
        doc.presentation.addElement(pg)
        return pg

    def frame(pg, st, cls, x, y, w, h):
        fr = draw.Frame(stylename=st, x=x, y=y, width=w, height=h)
        if cls:
            setp(fr, 'class', cls)
        tb = draw.TextBox()
        fr.addElement(tb)
        pg.addElement(fr)
        return tb

    def para(parent, *parts, st=None):
        p = text.P(**({'stylename': st} if st else {}))
        for part in parts:
            if isinstance(part, str):
                p.addText(part)
            else:
                p.addElement(part)
        parent.addElement(p)
        return p

    def span(st, s):
        sp = text.Span(stylename=st)
        sp.addText(s)
        return sp

    def notes(pg, s):
        nt = presentation.Notes()
        fr = setp(draw.Frame(stylename=pr_notes, x='2cm', y='14cm', width='17cm', height='12cm'), 'class', 'notes')
        tb = draw.TextBox()
        para(tb, s)
        fr.addElement(tb)
        nt.addElement(fr)
        pg.addElement(nt)

    # 1. title slide
    pg = page('Title', dp1)
    para(frame(pg, pr_title, 'title', '2cm', '3.5cm', '24cm', '3.2cm'), 'office-engine')
    para(frame(pg, pr_sub, 'subtitle', '2cm', '7.2cm', '24cm', '2.4cm'), 'Progressive Office viewer for the browser')
    notes(pg, 'Welcome everyone and introduce the project.')

    # 2. outline with levels, formatting and a link
    pg = page('Why it is fast', dp1)
    para(frame(pg, pr_title, 'title', '1.4cm', '0.6cm', '25.2cm', '2.6cm'), 'Why it is fast')
    tb = frame(pg, pr_out, 'outline', '1.4cm', '3.6cm', '25.2cm', '10.8cm')
    lst = text.List(stylename=outline_list)
    item = text.ListItem()
    para(item, 'Lazy ZIP: only the ', span(t_bold, 'central directory'), ' is read up front')
    sub = text.List()
    sub_item = text.ListItem()
    para(sub_item, 'Entries are inflated ', span(t_italic, 'on demand'))
    sub.addElement(sub_item)
    item.addElement(sub)
    lst.addElement(item)
    item = text.ListItem()
    para(item, 'Streaming XML parser in a ', span(t_under, 'Web Worker'))
    lst.addElement(item)
    item = text.ListItem()
    link = text.A(href='https://example.com/docs', type='simple')
    link.addText('documentation')
    para(item, 'Read the ', link, ' for ', span(t_code, 'open()'), ' and E = mc', span(t_sup, '2'))
    lst.addElement(item)
    tb.addElement(lst)
    notes(pg, 'Mention the lazy ZIP reader first.')

    # 3. shapes, a picture, a line
    pg = page('Shapes', dp1)
    para(frame(pg, pr_title, 'title', '1.4cm', '0.6cm', '25.2cm', '2.6cm'), 'Shapes & images')
    cs = draw.CustomShape(stylename=gr_blue, x='1.5cm', y='4cm', width='5cm', height='3cm')
    para(cs, 'Rounded', st=p_center)
    cs.addElement(draw.EnhancedGeometry(type='round-rectangle', viewbox='0 0 21600 21600', modifiers='3600'))
    pg.addElement(cs)
    el = draw.Ellipse(stylename=gr_green, x='7.5cm', y='4cm', width='5cm', height='3cm')
    para(el, 'Oval', st=p_center)
    pg.addElement(el)
    cs = draw.CustomShape(stylename=gr_amber, x='13.5cm', y='4cm', width='5cm', height='3cm')
    para(cs, 'Arrow', st=p_center)
    geo = draw.EnhancedGeometry(type='right-arrow', viewbox='0 0 21600 21600', modifiers='16200 5400', enhancedpath='M 0 ?f0 L ?f1 ?f0 ?f1 0 21600 10800 ?f1 21600 ?f1 ?f2 0 ?f2 Z N')
    for name, formula in [('f0', '$1 '), ('f1', '$0 '), ('f2', '21600-$1 ')]:
        geo.addElement(draw.Equation(name=name, formula=formula))
    cs.addElement(geo)
    pg.addElement(cs)
    star = draw.CustomShape(stylename=gr_star, width='3cm', height='3cm', transform='rotate (0.523598775598299) translate (20cm 4.5cm)')
    star.addElement(draw.EnhancedGeometry(type='star5', viewbox='0 0 21600 21600', enhancedpath='M 10797 0 L 8278 8256 0 8256 6722 13405 4198 21600 10797 16580 17401 21600 14878 13405 21600 8256 13321 8256 10797 0 Z N'))
    pg.addElement(star)
    tb = frame(pg, gr_text, None, '1.5cm', '8.5cm', '11cm', '1.5cm')
    para(tb, 'A free text box with ', span(t_red, 'colored bold text'))
    pic = draw.Frame(stylename=gr_pic, x='14cm', y='8.2cm', width='5cm', height='3.2cm', name='Landscape')
    href = doc.addPicture('Pictures/landscape.png', 'image/png', png(160, 100, landscape))
    pic.addElement(draw.Image(href=href))
    pg.addElement(pic)
    pg.addElement(draw.Line(stylename=gr_line, x1='1.5cm', y1='12.5cm', x2='12cm', y2='12.5cm'))

    # 4. table
    pg = page('Table', dp1)
    para(frame(pg, pr_title, 'title', '1.4cm', '0.6cm', '25.2cm', '2.6cm'), 'Table')
    fr = draw.Frame(stylename=std, x='4cm', y='4cm', width='18cm', height='7cm')
    tbl = table.Table(templatename='default')
    tbl.addElement(table.TableColumn(stylename=co1))
    tbl.addElement(table.TableColumn(stylename=co2, numbercolumnsrepeated=3))
    rows = [['Metric', '2024', '2025', '2026'], ['Users', '380', '750', '424'], ['Revenue', '746', '55', '781'], ['NPS', '469', '441', '643'], ['Latency', '130', '209', '885']]
    for r, values in enumerate(rows):
        tr = table.TableRow(stylename=ro1)
        for v in values:
            tc = table.TableCell(stylename=ce_head if r == 0 else ce_body)
            para(tc, v)
            tr.addElement(tc)
        tbl.addElement(tr)
    fr.addElement(tbl)
    pg.addElement(fr)

    # 5. numbered list on a tinted background
    pg = page('Steps', dp2)
    para(frame(pg, pr_title, 'title', '1.4cm', '0.6cm', '25.2cm', '2.6cm'), 'Getting started')
    tb = frame(pg, pr_out, 'outline', '1.4cm', '3.6cm', '25.2cm', '10.8cm')
    lst = text.List(stylename=numbers)
    for s in ['Install the package', 'Open a document']:
        item = text.ListItem()
        para(item, s)
        lst.addElement(item)
    item = text.ListItem()
    para(item, 'Edit and save')
    sub = text.List()
    for s in ['Text and styles', 'Tables']:
        sub_item = text.ListItem()
        para(sub_item, s)
        sub.addElement(sub_item)
    item.addElement(sub)
    lst.addElement(item)
    tb.addElement(lst)

    # 6. a hidden backup slide
    pg = page('Backup', dp3)
    para(frame(pg, pr_title, 'title', '1.4cm', '0.6cm', '25.2cm', '2.6cm'), 'Backup')
    return doc


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else 'test/fixtures/legacy')
    out.mkdir(parents=True, exist_ok=True)
    build().save(str(out / 'sample.odp'))
    build(template=True).save(str(out / 'sample.otp'))
    # flat XML: the picture inline (office:binary-data) instead of a package path
    flat = build().xml().decode('utf-8')
    image = re.search(r'<draw:image xlink:href="Pictures/landscape.png"[^>]*/>', flat)
    data = base64.b64encode(png(160, 100, landscape)).decode('ascii')
    flat = flat.replace(image.group(0), f'<draw:image><office:binary-data>{data}</office:binary-data></draw:image>')
    (out / 'sample.fodp').write_text(flat, encoding='utf-8')


if __name__ == '__main__':
    main()
