"""
Builds the OpenDocument text fixtures of test/fixtures/legacy/ with odfpy (see odt.md):
- sample.odt: the richer sample (modeled on playground/public/samples/sample.docx), laid out the way LibreOffice
  writes documents: named styles with LibreOffice's names, automatic styles for direct formatting, a master page
  with a header and footer
- sample.fodt: the same document as flat XML (the picture inline as base64)
- template.ott: a short template (text-template mimetype)

Usage: python3 scripts/legacy-samples/odt.py <output dir> <picture.png>
"""
import sys

from odf.dc import Creator, Title as DcTitle
from odf.draw import Frame, Image, TextBox
from odf.office import BinaryData
from odf.opendocument import OpenDocumentText
from odf.style import (
    Column, Columns, DefaultStyle, FontFace, Footer, FooterStyle, GraphicProperties, Header, HeaderFooterProperties,
    HeaderStyle, ListLevelLabelAlignment, ListLevelProperties, MasterPage, PageLayout, PageLayoutProperties,
    ParagraphProperties, SectionProperties, Style, TableCellProperties, TableColumnProperties, TableProperties,
    TabStop, TabStops, TextProperties,
)
from odf.table import CoveredTableCell, Table, TableCell, TableColumn, TableHeaderRows, TableRow
from odf.text import (
    A, BookmarkEnd, BookmarkStart, H, IndexBody, IndexTitle, LineBreak, List, ListItem, ListLevelStyleBullet,
    ListLevelStyleNumber, ListStyle, Note, NoteBody, NoteCitation, NotesConfiguration, OutlineLevelStyle, OutlineStyle,
    P, PageCount, PageNumber, S, Section, Span, Tab, TableOfContent, TableOfContentSource,
)

out, picture = sys.argv[1], sys.argv[2]
png = open(picture, 'rb').read()


def build(flat=False, template=False):
    doc = OpenDocumentText()
    if template:
        doc.mimetype = 'application/vnd.oasis.opendocument.text-template'
    doc.meta.addElement(DcTitle(text='Quarterly Report'))
    doc.meta.addElement(Creator(text='Planning Team'))
    for name, family, generic in [('Liberation Serif', "'Liberation Serif'", 'roman'), ('Liberation Sans', "'Liberation Sans'", 'swiss'), ('Liberation Mono', "'Liberation Mono'", 'modern')]:
        doc.fontfacedecls.addElement(FontFace(name=name, fontfamily=family, fontfamilygeneric=generic, fontpitch='variable'))

    # ---- named styles (office:styles), as LibreOffice names them
    st = doc.styles
    d = DefaultStyle(family='paragraph')
    d.addElement(ParagraphProperties(tabstopdistance='1.25cm', orphans='2', widows='2'))
    d.addElement(TextProperties(fontname='Liberation Serif', fontsize='12pt', language='en', country='US'))
    st.addElement(d)

    def named(name, display=None, parent=None, nxt=None, family='paragraph', para=None, text=None, outline=None, cls='text'):
        s = Style(name=name, family=family, **({'displayname': display} if display else {}), **({'parentstylename': parent} if parent else {}),
                  **({'nextstylename': nxt} if nxt else {}), **({'defaultoutlinelevel': str(outline)} if outline else {}), **({'class': cls} if family == 'paragraph' else {}))
        if para:
            s.addElement(ParagraphProperties(**para))
        if text:
            s.addElement(TextProperties(**text))
        st.addElement(s)
        return s

    named('Standard')
    named('Heading', parent='Standard', nxt='Text_20_body', para={'margintop': '0.423cm', 'marginbottom': '0.212cm', 'keepwithnext': 'always'}, text={'fontname': 'Liberation Sans', 'fontsize': '14pt'})
    named('Text_20_body', 'Text body', 'Standard', para={'marginbottom': '0.247cm', 'lineheight': '115%'})
    named('Title', parent='Heading', nxt='Subtitle', para={'textalign': 'center'}, text={'fontsize': '28pt', 'fontweight': 'bold'}, cls='chapter')
    named('Subtitle', parent='Heading', nxt='Text_20_body', para={'margintop': '0.106cm', 'textalign': 'center'}, text={'fontsize': '18pt', 'color': '#595959'}, cls='chapter')
    named('Heading_20_1', 'Heading 1', 'Heading', 'Text_20_body', para={'margintop': '0.423cm', 'marginbottom': '0.212cm'}, text={'fontsize': '130%', 'fontweight': 'bold', 'color': '#1f3864'}, outline=1)
    named('Heading_20_2', 'Heading 2', 'Heading', 'Text_20_body', para={'margintop': '0.353cm', 'marginbottom': '0.212cm'}, text={'fontsize': '115%', 'fontweight': 'bold', 'fontstyle': 'italic'}, outline=2)
    named('List', parent='Text_20_body', cls='list')
    named('Table_20_Contents', 'Table Contents', 'Standard', cls='extra')
    named('Table_20_Heading', 'Table Heading', 'Table_20_Contents', para={'textalign': 'center'}, text={'fontweight': 'bold'}, cls='extra')
    named('Caption', parent='Standard', para={'margintop': '0.212cm', 'marginbottom': '0.212cm', 'textalign': 'center'}, text={'fontsize': '10pt', 'fontstyle': 'italic'}, cls='extra')
    named('Header', parent='Standard', para={'textalign': 'end'}, text={'fontsize': '9pt', 'color': '#808080'}, cls='extra')
    named('Footer', parent='Standard', para={'textalign': 'center'}, text={'fontsize': '9pt'}, cls='extra')
    named('Footnote', parent='Standard', para={'marginleft': '0.6cm', 'textindent': '-0.6cm'}, text={'fontsize': '10pt'}, cls='extra')
    named('Quotations', parent='Standard', para={'marginleft': '1cm', 'marginright': '1cm', 'marginbottom': '0.283cm'}, text={'fontstyle': 'italic'}, cls='html')
    named('Contents_20_Heading', 'Contents Heading', 'Heading', text={'fontsize': '16pt', 'fontweight': 'bold'}, cls='index')
    c1 = named('Contents_20_1', 'Contents 1', 'Standard', cls='index')
    c1pp = ParagraphProperties()
    tabs = TabStops()
    tabs.addElement(TabStop(position='17cm', type='right', leaderstyle='dotted', leadertext='.'))
    c1pp.addElement(tabs)
    c1.addElement(c1pp)
    named('Preformatted_20_Text', 'Preformatted Text', 'Standard', text={'fontname': 'Liberation Mono', 'fontsize': '10pt'}, cls='html')
    named('Internet_20_link', 'Internet link', family='text', text={'color': '#000080', 'textunderlinestyle': 'solid', 'textunderlinewidth': 'auto', 'textunderlinecolor': 'font-color'})
    named('Strong_20_Emphasis', 'Strong Emphasis', family='text', text={'fontweight': 'bold'})
    named('Footnote_20_anchor', 'Footnote anchor', family='text', text={'textposition': 'super 58%'})
    named('Footnote_20_Symbol', 'Footnote Symbol', family='text')

    # outline numbering of the headings: "1." on level 1, "1.1" on level 2
    outline = OutlineStyle(name='Outline')
    for lvl in range(1, 4):
        ls = OutlineLevelStyle(level=str(lvl), numformat='1', numsuffix='.' if lvl == 1 else '', displaylevels=str(lvl))
        lp = ListLevelProperties(listlevelpositionandspacemode='label-alignment')
        lp.addElement(ListLevelLabelAlignment(labelfollowedby='space'))
        ls.addElement(lp)
        outline.addElement(ls)
    st.addElement(outline)
    st.addElement(NotesConfiguration(noteclass='footnote', citationstylename='Footnote_20_Symbol', citationbodystylename='Footnote_20_anchor', numformat='1', startvalue='0', footnotesposition='page', startnumberingat='document'))

    # ---- automatic styles: direct formatting
    au = doc.automaticstyles

    def auto(name, family, parent=None, para=None, text=None, **kw):
        s = Style(name=name, family=family, **({'parentstylename': parent} if parent else {}), **kw)
        if para is not None:
            s.addElement(para if not isinstance(para, dict) else ParagraphProperties(**para))
        if text:
            s.addElement(TextProperties(**text))
        au.addElement(s)
        return s

    auto('P1', 'paragraph', 'Title', masterpagename='Standard')
    auto('P2', 'paragraph', 'Text_20_body', para={'textalign': 'justify', 'textindent': '0.5cm'})
    auto('P3', 'paragraph', 'Text_20_body', para={'textalign': 'center', 'backgroundcolor': '#fff2cc', 'border': '0.5pt solid #bf9000', 'padding': '0.1cm'})
    p4 = ParagraphProperties()
    t4 = TabStops()
    t4.addElement(TabStop(position='12cm', type='right', leaderstyle='dotted', leadertext='.'))
    p4.addElement(t4)
    auto('P4', 'paragraph', 'Text_20_body', para=p4)
    auto('P5', 'paragraph', 'Heading_20_1', para={'breakbefore': 'page'})
    auto('P6', 'paragraph', 'Standard', para={'textalign': 'center'})
    auto('P7', 'paragraph', 'Table_20_Contents', para={'textalign': 'end'})
    auto('T1', 'text', text={'fontweight': 'bold'})
    auto('T2', 'text', text={'fontstyle': 'italic'})
    auto('T3', 'text', text={'textunderlinestyle': 'solid', 'textunderlinewidth': 'auto', 'textunderlinecolor': 'font-color', 'color': '#c00000'})
    auto('T4', 'text', text={'fontsize': '16pt'})
    auto('T5', 'text', text={'textlinethroughstyle': 'solid', 'textlinethroughtype': 'single'})
    auto('T6', 'text', text={'backgroundcolor': '#ffff00'})
    auto('T7', 'text', text={'textposition': 'super 58%'})
    auto('T8', 'text', text={'textposition': 'sub 58%'})
    auto('T9', 'text', text={'fontvariant': 'small-caps', 'letterspacing': '0.05cm'})
    auto('T10', 'text', text={'fontname': 'Liberation Mono', 'fontsize': '10pt'})
    tb = Style(name='Table1', family='table')
    tb.addElement(TableProperties(width='17cm', align='margins'))
    au.addElement(tb)
    for col, w in zip('ABCD', ['5.5cm', '3.5cm', '3.5cm', '4.5cm']):
        c = Style(name=f'Table1.{col}', family='table-column')
        c.addElement(TableColumnProperties(columnwidth=w))
        au.addElement(c)
    cells = {
        'Table1.A1': {'backgroundcolor': '#1f3864', 'border': '0.5pt solid #000000', 'padding': '0.1cm', 'verticalalign': 'middle'},
        'Table1.B2': {'border': '0.5pt solid #000000', 'padding': '0.1cm'},
        'Table1.A2': {'backgroundcolor': '#d9e2f3', 'border': '0.5pt solid #000000', 'padding': '0.1cm', 'verticalalign': 'middle'},
    }
    for name, props in cells.items():
        c = Style(name=name, family='table-cell')
        c.addElement(TableCellProperties(**props))
        au.addElement(c)
    auto('T11', 'text', text={'color': '#ffffff', 'fontweight': 'bold'})
    fr = Style(name='fr1', family='graphic')
    fr.addElement(GraphicProperties(wrap='none', horizontalpos='center', horizontalrel='paragraph', verticalpos='top', verticalrel='paragraph'))
    au.addElement(fr)
    fr2 = Style(name='fr2', family='graphic')
    fr2.addElement(GraphicProperties(wrap='parallel', horizontalpos='right', horizontalrel='paragraph', verticalpos='top', verticalrel='paragraph',
                                     backgroundcolor='#e2efda', border='0.75pt solid #548235', padding='0.15cm', marginleft='0.2cm'))
    au.addElement(fr2)
    sect = Style(name='Sect1', family='section')
    sp = SectionProperties()
    cols = Columns(columncount='2', columngap='0.6cm')
    cols.addElement(Column(relwidth='1*', startindent='0cm', endindent='0.3cm'))
    cols.addElement(Column(relwidth='1*', startindent='0.3cm', endindent='0cm'))
    sp.addElement(cols)
    sect.addElement(sp)
    au.addElement(sect)

    # lists: bullets ("•", "◦") and numbers ("1.", "a)")
    bullets = ListStyle(name='L1')
    for lvl, ch in [(1, '•'), (2, '◦')]:
        b = ListLevelStyleBullet(level=str(lvl), bulletchar=ch)
        lp = ListLevelProperties(listlevelpositionandspacemode='label-alignment')
        lp.addElement(ListLevelLabelAlignment(labelfollowedby='listtab', listtabstopposition=f'{0.635 * lvl * 2:.3f}cm', textindent='-0.635cm', marginleft=f'{0.635 * lvl * 2:.3f}cm'))
        b.addElement(lp)
        bullets.addElement(b)
    au.addElement(bullets)
    numbers = ListStyle(name='L2')
    for lvl, fmt, suffix in [(1, '1', '.'), (2, 'a', ')')]:
        n = ListLevelStyleNumber(level=str(lvl), numformat=fmt, numsuffix=suffix)
        lp = ListLevelProperties(listlevelpositionandspacemode='label-alignment')
        lp.addElement(ListLevelLabelAlignment(labelfollowedby='listtab', listtabstopposition=f'{0.635 * lvl * 2:.3f}cm', textindent='-0.635cm', marginleft=f'{0.635 * lvl * 2:.3f}cm'))
        n.addElement(lp)
        numbers.addElement(n)
    au.addElement(numbers)

    # ---- page layout, header and footer
    pl = PageLayout(name='pm1')
    pl.addElement(PageLayoutProperties(pagewidth='21cm', pageheight='29.7cm', printorientation='portrait', margintop='1.5cm', marginbottom='1.5cm', marginleft='2cm', marginright='2cm', numformat='1'))
    hs = HeaderStyle()
    hs.addElement(HeaderFooterProperties(minheight='0.6cm', marginbottom='0.4cm'))
    pl.addElement(hs)
    fs = FooterStyle()
    fs.addElement(HeaderFooterProperties(minheight='0.6cm', margintop='0.4cm'))
    pl.addElement(fs)
    au.addElement(pl)
    mp = MasterPage(name='Standard', pagelayoutname=pl)
    h = Header()
    h.addElement(P(stylename='Header', text='Quarterly Report — Draft'))
    mp.addElement(h)
    f = Footer()
    fp = P(stylename='Footer', text='Page ')
    fp.addElement(PageNumber(selectpage='current', text='1'))
    fp.addText(' of ')
    fp.addElement(PageCount(text='2'))
    f.addElement(fp)
    mp.addElement(f)
    doc.masterstyles.addElement(mp)

    # ---- body
    body = doc.text

    def para(style, *parts):
        p = P(stylename=style)
        for x in parts:
            if isinstance(x, str):
                p.addText(x)
            elif isinstance(x, tuple):
                sp = Span(stylename=x[0])
                sp.addText(x[1]) if isinstance(x[1], str) else sp.addElement(x[1])
                p.addElement(sp)
            else:
                p.addElement(x)
        return p

    body.addElement(para('P1', 'Quarterly Report'))
    body.addElement(para('Subtitle', 'Prepared for the planning team'))

    toc = TableOfContent(name='Table of Contents1', protected='true')
    toc.addElement(TableOfContentSource(outlinelevel='2'))
    ib = IndexBody()
    it = IndexTitle(name='Table of Contents1_Head')
    it.addElement(P(stylename='Contents_20_Heading', text='Contents'))
    ib.addElement(it)
    for label, target, page in [('1. Highlights', 'highlights', '1'), ('2. Results', 'results', '1'), ('3. Appendix', 'appendix', '2')]:
        tp = P(stylename='Contents_20_1')
        a = A(type='simple', href='#' + target)
        a.addText(label)
        a.addElement(Tab())
        a.addText(page)
        tp.addElement(a)
        ib.addElement(tp)
    toc.addElement(ib)
    body.addElement(toc)

    h1 = H(outlinelevel='1', stylename='Heading_20_1')
    h1.addElement(BookmarkStart(name='highlights'))
    h1.addText('Highlights')
    h1.addElement(BookmarkEnd(name='highlights'))
    body.addElement(h1)

    note = Note(id='ftn1', noteclass='footnote')
    note.addElement(NoteCitation(text='1'))
    nb = NoteBody()
    nb.addElement(P(stylename='Footnote', text='Figures are rounded to the nearest thousand.'))
    note.addElement(nb)
    link = A(type='simple', href='https://example.com/report', stylename='Internet_20_link', visitedstylename='Internet_20_link')
    link.addText('example.com')
    body.addElement(para('P2', 'This document shows ', ('T1', 'bold'), ', ', ('T2', 'italic'), ', ', ('T3', 'underlined red'), ', ', ('T5', 'struck out'), ', ',
                         ('T6', 'highlighted'), ' and ', ('T4', 'large'), ' text, x', ('T7', '2'), ' and H', ('T8', '2'), 'O, ', ('T9', 'Small Caps'),
                         ' and ', ('T10', 'code()'), '. Revenue grew', note, ' this quarter. More at ', link, '.'))
    body.addElement(para('Quotations', 'Spacing is kept:', S(c=3), 'three spaces, a', Tab(), 'tab and a', LineBreak(), 'line break.'))

    lst = List(stylename='L1')
    for text, sub in [('Revenue grew 24% year over year', None), ('Two new regions launched', ['North', 'South']), ('Customer churn fell to 3.1%', None)]:
        li = ListItem()
        li.addElement(P(stylename='List', text=text))
        if sub:
            sl = List()
            for s in sub:
                si = ListItem()
                si.addElement(P(stylename='List', text=s))
                sl.addElement(si)
            li.addElement(sl)
        lst.addElement(li)
    body.addElement(lst)

    h2 = H(outlinelevel='1', stylename='Heading_20_1')
    h2.addElement(BookmarkStart(name='results'))
    h2.addText('Results')
    h2.addElement(BookmarkEnd(name='results'))
    body.addElement(h2)
    body.addElement(para('Text_20_body', 'The next steps are:'))
    nl = List(stylename='L2', id='list-steps')
    for text, sub in [('Hire the platform team', None), ('Ship the viewer engine', ['Text documents', 'Spreadsheets']), ('Expand to new markets', None)]:
        li = ListItem()
        li.addElement(P(stylename='List', text=text))
        if sub:
            sl = List()
            for s in sub:
                si = ListItem()
                si.addElement(P(stylename='List', text=s))
                sl.addElement(si)
            li.addElement(sl)
        nl.addElement(li)
    body.addElement(nl)
    body.addElement(para('Text_20_body', 'After a pause, the list goes on:'))
    nl2 = List(stylename='L2', continuenumbering='true')
    li = ListItem()
    li.addElement(P(stylename='List', text='Review the budget'))
    nl2.addElement(li)
    body.addElement(nl2)

    body.addElement(H(outlinelevel='2', stylename='Heading_20_2', text='Regions'))
    t = Table(name='Table1', stylename='Table1')
    for col in 'ABCD':
        t.addElement(TableColumn(stylename=f'Table1.{col}'))
    hr = TableHeaderRows()
    row = TableRow()
    for label in ['Region', 'Q1', 'Q2', 'Total']:
        c = TableCell(stylename='Table1.A1', valuetype='string')
        c.addElement(para('Table_20_Heading', ('T11', label)))
        row.addElement(c)
    hr.addElement(row)
    t.addElement(hr)
    data = [('North', '120', '140', '260'), ('South', '98', '110', '208'), ('East', '132', '150', '282')]
    for k, (region, *nums) in enumerate(data):
        row = TableRow()
        if k == 0:
            c = TableCell(stylename='Table1.A2', valuetype='string', numberrowsspanned='2')
            c.addElement(P(stylename='Table_20_Contents', text='North and South'))
            row.addElement(c)
        elif k == 1:
            row.addElement(CoveredTableCell())
        else:
            c = TableCell(stylename='Table1.B2', valuetype='string')
            c.addElement(P(stylename='Table_20_Contents', text=region))
            row.addElement(c)
        for v in nums:
            c = TableCell(stylename='Table1.B2', valuetype='float', value=v)
            c.addElement(P(stylename='P7', text=v))
            row.addElement(c)
        t.addElement(row)
    row = TableRow()
    c = TableCell(stylename='Table1.A2', valuetype='string', numbercolumnsspanned='3')
    c.addElement(para('Table_20_Contents', ('Strong_20_Emphasis', 'All regions')))
    row.addElement(c)
    row.addElement(CoveredTableCell())
    row.addElement(CoveredTableCell())
    c = TableCell(stylename='Table1.A2', valuetype='float', value='750')
    c.addElement(para('P7', ('Strong_20_Emphasis', '750')))
    row.addElement(c)
    t.addElement(row)
    body.addElement(t)
    body.addElement(para('P4', 'Total revenue', Tab(), '750'))

    body.addElement(H(outlinelevel='2', stylename='Heading_20_2', text='Figures'))
    pic = P(stylename='P6')
    frame = Frame(stylename=fr, name='Image1', anchortype='as-char', width='12cm', height='6cm', zindex='0')
    if flat:
        img = Image(type='simple', show='embed', actuate='onLoad')
        img.addElement(BinaryData(text=__import__('base64').b64encode(png).decode()))
    else:
        img = Image(href=doc.addPicture(picture), type='simple', show='embed', actuate='onLoad')
    frame.addElement(img)
    pic.addElement(frame)
    body.addElement(pic)
    body.addElement(P(stylename='Caption', text='Figure 1 — generated gradient image'))
    boxp = P(stylename='Text_20_body')
    box = Frame(stylename=fr2, name='Note', anchortype='paragraph', width='5cm', zindex='1')
    tbx = TextBox(minheight='1.5cm')
    tbx.addElement(para('Standard', ('T1', 'Note: '), 'figures are unaudited.'))
    box.addElement(tbx)
    boxp.addElement(box)
    boxp.addText('The chart above shows the trend; the box beside this paragraph is a text frame anchored to it.')
    body.addElement(boxp)
    body.addElement(para('P3', 'Back to the ', A(type='simple', href='#highlights', text='highlights'), '.'))

    h3 = H(outlinelevel='1', stylename='P5')
    h3.addElement(BookmarkStart(name='appendix'))
    h3.addText('Appendix')
    h3.addElement(BookmarkEnd(name='appendix'))
    body.addElement(h3)
    sec = Section(name='Columns', stylename='Sect1')
    words = 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua'.split()
    for k in range(6):
        text = ' '.join(words[(k * 3 + j) % len(words)] for j in range(40)).capitalize() + '.'
        sec.addElement(P(stylename='Text_20_body', text=text))
    body.addElement(sec)
    body.addElement(para('Preformatted_20_Text', 'const total = north + south'))
    return doc


build().save(f'{out}/sample.odt')
with open(f'{out}/sample.fodt', 'wb') as fh:
    fh.write(build(flat=True).xml())
small = OpenDocumentText()
small.mimetype = 'application/vnd.oasis.opendocument.text-template'
s = Style(name='Heading_20_1', displayname='Heading 1', family='paragraph', defaultoutlinelevel='1')
s.addElement(TextProperties(fontsize='18pt', fontweight='bold'))
small.styles.addElement(s)
small.text.addElement(H(outlinelevel='1', stylename='Heading_20_1', text='Meeting notes'))
small.text.addElement(P(text='Date:'))
small.text.addElement(P(text='Attendees:'))
small.save(f'{out}/template.ott')
print('wrote', f'{out}/sample.odt', f'{out}/sample.fodt', f'{out}/template.ott')
