"""
Builds the OpenDocument spreadsheet fixtures of test/fixtures/legacy/ with odfpy (see ods.md):
- sample.ods: the richer sample (modeled on playground/public/samples/sample.xlsx)
- sample.fods: the same document as flat XML

Usage: python3 scripts/legacy-samples/ods.py <output dir>
"""
import datetime
import random
import sys

from odf.config import ConfigItem, ConfigItemMapEntry, ConfigItemMapIndexed, ConfigItemMapNamed, ConfigItemSet
from odf.dc import Creator, Date as DcDate
from odf.number import (
    AmPm, CurrencyStyle, CurrencySymbol, DateStyle, Day, DayOfWeek, Hours, Minutes, Month, Number, NumberStyle,
    PercentageStyle, Seconds, Text as NText, TimeStyle, Year,
)
from odf.office import Annotation
from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import (
    FontFace, Footer, FooterStyle, Header, HeaderFooterProperties, HeaderStyle, Map as SMap, MasterPage, PageLayout,
    PageLayoutProperties, ParagraphProperties, Style, TableCellProperties, TableColumnProperties, TableProperties,
    TableRowProperties, TextProperties,
)
from odf.table import (
    CoveredTableCell, DatabaseRange, DatabaseRanges, NamedExpressions, NamedRange, Table, TableCell, TableColumn,
    TableRow,
)
from odf.text import A, P, PageCount, PageNumber, SheetName

out = sys.argv[1] if len(sys.argv) > 1 else '.'


def build(sales_rows, with_data):
    random.seed(7)
    doc = OpenDocumentSpreadsheet()

    # fonts and the default cell style
    doc.fontfacedecls.addElement(FontFace(name='Liberation Sans', fontfamily="'Liberation Sans'", fontfamilygeneric='swiss', fontpitch='variable'))
    default = Style(name='Default', family='table-cell')
    default.addElement(TextProperties(fontname='Liberation Sans', fontsize='10pt'))
    doc.styles.addElement(default)

    # page: landscape A4 with the sheet name as header and page numbers as footer
    layout = PageLayout(name='pm1')
    layout.addElement(PageLayoutProperties(printorientation='landscape', pagewidth='29.7cm', pageheight='21cm', margintop='2cm', marginbottom='2cm', marginleft='1.5cm', marginright='1.5cm'))
    hs = HeaderStyle()
    hs.addElement(HeaderFooterProperties(minheight='0.75cm', marginbottom='0.25cm'))
    layout.addElement(hs)
    fs = FooterStyle()
    fs.addElement(HeaderFooterProperties(minheight='0.75cm', margintop='0.25cm'))
    layout.addElement(fs)
    doc.automaticstyles.addElement(layout)
    master = MasterPage(name='Default', pagelayoutname=layout)
    hp = P()
    hp.addElement(SheetName())
    header = Header()
    header.addElement(hp)
    master.addElement(header)
    footer = Footer()
    fp = P(text='Page ')
    fp.addElement(PageNumber())
    fp.addText(' of ')
    fp.addElement(PageCount())
    footer.addElement(fp)
    master.addElement(footer)
    doc.masterstyles.addElement(master)


    def add(style):
        doc.automaticstyles.addElement(style)
        return style


    # number styles
    n_int = add(NumberStyle(name='N1'))
    n_int.addElement(Number(decimalplaces='0', minintegerdigits='1', grouping='true'))
    n_usd = add(CurrencyStyle(name='N2'))
    n_usd.addElement(CurrencySymbol(language='en', country='US', text='$'))
    n_usd.addElement(Number(decimalplaces='2', minintegerdigits='1', grouping='true'))
    n_pct = add(PercentageStyle(name='N3'))
    n_pct.addElement(Number(decimalplaces='1', minintegerdigits='1'))
    n_pct.addElement(NText(text='%'))
    n_date = add(DateStyle(name='N4'))
    for e in (Year(style='long'), NText(text='-'), Month(style='long'), NText(text='-'), Day(style='long')):
        n_date.addElement(e)
    n_pos = add(NumberStyle(name='N5P0', volatile='true'))
    n_pos.addElement(Number(decimalplaces='2', minintegerdigits='1', grouping='true'))
    n_neg = add(NumberStyle(name='N5'))
    n_neg.addElement(TextProperties(color='#ff0000'))
    n_neg.addElement(NText(text='('))
    n_neg.addElement(Number(decimalplaces='2', minintegerdigits='1', grouping='true'))
    n_neg.addElement(NText(text=')'))
    n_neg.addElement(SMap(condition='value()>=0', applystylename='N5P0'))
    n_long = add(DateStyle(name='N6'))
    for e in (DayOfWeek(style='long'), NText(text=', '), Month(style='long', textual='true'), NText(text=' '), Day(), NText(text=', '),
              Year(style='long'), NText(text=' '), Hours(), NText(text=':'), Minutes(style='long'), NText(text=' '), AmPm()):
        n_long.addElement(e)
    n_dur = add(TimeStyle(name='N7', truncateonoverflow='false'))
    for e in (Hours(), NText(text=':'), Minutes(style='long'), NText(text=':'), Seconds(style='long')):
        n_dur.addElement(e)
    n_dec = add(NumberStyle(name='N8'))
    n_dec.addElement(Number(decimalplaces='2', minintegerdigits='1', grouping='true'))


    def cell_style(name, data=None, cell=None, text=None, para=None):
        s = Style(name=name, family='table-cell', parentstylename='Default', **({'datastylename': data} if data else {}))
        if cell:
            s.addElement(TableCellProperties(**cell))
        if para:
            s.addElement(ParagraphProperties(**para))
        if text:
            s.addElement(TextProperties(**text))
        return add(s)


    band = {'backgroundcolor': '#d9e1f2'}
    cell_style('ceTitle', text={'fontweight': 'bold', 'fontsize': '16pt', 'color': '#1f3864'}, para={'textalign': 'center'}, cell={'textalignsource': 'fix', 'verticalalign': 'middle'})
    cell_style('ceHead', text={'fontweight': 'bold', 'color': '#ffffff'}, para={'textalign': 'center'},
               cell={'textalignsource': 'fix', 'backgroundcolor': '#4472c4', 'borderbottom': '0.75pt solid #8ea9db'})
    cell_style('ceBand', cell=band)
    cell_style('ceDate', data='N4')
    cell_style('ceDateBand', data='N4', cell=band)
    cell_style('ceInt', data='N1')
    cell_style('ceIntBand', data='N1', cell=band)
    cell_style('ceUsd', data='N2')
    cell_style('ceUsdBand', data='N2', cell=band)
    cell_style('ceDec', data='N8')
    cell_style('ceDecBand', data='N8', cell=band)
    cell_style('cePct', data='N3')
    cell_style('cePctBand', data='N3', cell=band)
    cell_style('ceNeg', data='N5')
    cell_style('ceLong', data='N6')
    cell_style('ceDur', data='N7')
    cell_style('ceTotal', text={'fontweight': 'bold'}, cell={'bordertop': '1.5pt solid #000000'})
    cell_style('ceTotalUsd', data='N8', text={'fontweight': 'bold'}, cell={'bordertop': '1.5pt solid #000000'})
    cell_style('ceWrap', cell={'wrapoption': 'wrap', 'verticalalign': 'top'})
    cell_style('ceItalic', text={'fontstyle': 'italic', 'textunderlinestyle': 'solid', 'textunderlinewidth': 'auto', 'textunderlinecolor': 'font-color'})
    cell_style('ceRotated', cell={'rotationangle': '90', 'verticalalign': 'middle', 'border': '0.75pt solid #000000'}, para={'textalign': 'center'})
    cell_style('ceStrike', text={'textlinethroughstyle': 'solid', 'color': '#c00000', 'fontname': 'Liberation Sans', 'fontsize': '12pt'})
    cell_style('ceIndent', para={'textalign': 'start', 'marginleft': '20pt'}, cell={'textalignsource': 'fix'})

    for name, width in (('coA', '3.5cm'), ('coB', '3.2cm'), ('coDef', '2.258cm'), ('coH', '7.5cm'), ('coWide', '9cm')):
        s = add(Style(name=name, family='table-column'))
        s.addElement(TableColumnProperties(columnwidth=width, breakbefore='auto'))
    for name, height, optimal in (('ro1', '0.452cm', 'true'), ('roTitle', '0.9cm', 'false'), ('roTall', '120pt', 'false'), ('roTurned', '1.6cm', 'true')):
        s = add(Style(name=name, family='table-row'))
        s.addElement(TableRowProperties(rowheight=height, useoptimalrowheight=optimal, breakbefore='auto'))
    ta = add(Style(name='ta1', family='table', masterpagename='Default'))
    ta.addElement(TableProperties(display='true', writingmode='lr-tb'))
    ta_hidden = add(Style(name='taHidden', family='table', masterpagename='Default'))
    ta_hidden.addElement(TableProperties(display='false', writingmode='lr-tb'))


    def cell(value=None, style=None, formula=None, kind=None, text=None, span=None, **extra):
        """a table cell: numbers, strings, dates (datetime.date), booleans, with the text a spreadsheet shows"""
        a = dict(extra)
        if style:
            a['stylename'] = style
        if formula:
            a['formula'] = formula
        if span:
            a['numbercolumnsspanned'] = span
            a['numberrowsspanned'] = 1
        shown = text
        if isinstance(value, bool):
            a.update(valuetype='boolean', booleanvalue='true' if value else 'false')
            shown = shown or ('TRUE' if value else 'FALSE')
        elif isinstance(value, datetime.datetime):
            a.update(valuetype='date', datevalue=value.strftime('%Y-%m-%dT%H:%M:%S'))
            shown = shown or value.strftime('%Y-%m-%d %H:%M')
        elif isinstance(value, datetime.date):
            a.update(valuetype='date', datevalue=value.isoformat())
            shown = shown or value.isoformat()
        elif isinstance(value, (int, float)):
            a.update(valuetype=kind or 'float', value=value)
            if kind == 'currency':
                a['currency'] = 'USD'
            shown = shown or str(value)
        elif value is not None:
            a['valuetype'] = 'string'
            shown = value
        c = TableCell(**a)
        if shown is not None:
            c.addElement(P(text=shown))
        return c


    def row(cells, style='ro1', **attrs):
        r = TableRow(stylename=style, **attrs)
        for c in cells:
            r.addElement(c)
        return r


    # ---- Sales
    sales = Table(name='Sales', stylename=ta)
    for st in ('coA', 'coB', 'coDef', 'coDef', 'coA', 'coDef', 'coDef', 'coH'):
        sales.addElement(TableColumn(stylename=st, defaultcellstylename='Default'))
    title = cell('Sales by region — FY2026', 'ceTitle', span=6)
    sales.addElement(row([title] + [CoveredTableCell() for _ in range(5)], 'roTitle'))
    heads = [cell(h, 'ceHead') for h in ('Region', 'Date', 'Units', 'Price', 'Revenue', 'Margin')]
    note = Annotation()
    note.addElement(Creator(text='Sample Author'))
    note.addElement(DcDate(text='2026-01-05T10:00:00'))
    note.addElement(P(text='Units × Price, computed by the spreadsheet.'))
    heads[4].insertBefore(note, heads[4].firstChild)
    heads.append(cell('Lookup'))
    heads.append(cell('Revenue is a formula; its cached value is stored in the file.'))
    sales.addElement(row(heads))
    regions = ['North', 'Central', 'South', 'East', 'West']
    extras = {
        0: cell(1234567.891, 'ceDec'),
        1: cell(-42.5, 'ceNeg', text='(42.50)'),
        2: cell(0.000123),
        3: cell(True),
        4: cell(datetime.datetime(2023, 12, 31, 18, 0), 'ceLong', text='Sunday, December 31, 2023 6:00 PM'),
        5: TableCell(stylename='ceDur', valuetype='time', timevalue='PT36H30M00S'),
    }
    extras[5].addElement(P(text='36:30:00'))
    start = datetime.date(2026, 1, 15)
    units_total = 0
    revenue_total = 0.0
    for i in range(sales_rows):
        r = i + 3
        b = 'Band' if i % 2 == 0 else ''
        units = random.randint(50, 500)
        price = round(random.uniform(3, 90), 2)
        revenue = round(units * price, 2)
        units_total += units
        revenue_total += revenue
        region = regions[(i * 3) % 5]
        cells = [
            cell(region, 'ceBand' if b else None),
            cell(start + datetime.timedelta(days=5 * i), 'ceDate' + b),
            cell(units, 'ceInt' + b),
            cell(price, 'ceUsd' + b, kind='currency', text=f'${price:,.2f}'),
            cell(revenue, 'ceDec' + b, formula=f'of:=[.C{r}]*[.D{r}]', text=f'{revenue:,.2f}'),
            cell(round(random.uniform(0.02, 0.6), 4), 'cePct' + b, kind='percentage'),
            cell(regions.index(region) + 1, formula=f'of:=VLOOKUP([.A{r}];[$Lookup.$A$1:.$B$5];2;0)'),
        ]
        if i in extras:
            cells.append(extras[i])
        sales.addElement(row(cells))
    sales.addElement(row([
        cell('Total', 'ceTotal'), cell(None, 'ceTotal'),
        cell(units_total, 'ceTotal', formula=f'of:=SUM([.C3:.C{sales_rows + 2}])'),
        cell(None, 'ceTotal'),
        cell(round(revenue_total, 2), 'ceTotalUsd', formula=f'of:=SUM([.E3:.E{sales_rows + 2}])', text=f'{revenue_total:,.2f}'),
        cell(None, 'ceTotal'),
    ]))
    doc.spreadsheet.addElement(sales)

    # ---- Notes
    notes = Table(name='Notes', stylename=ta)
    notes.addElement(TableColumn(stylename='coWide', defaultcellstylename='Default'))
    notes.addElement(TableColumn(stylename='coA', defaultcellstylename='Default', numbercolumnsrepeated=2))
    wrapped = ('Wrapped text cell. This paragraph is long enough to wrap inside the column, and the row is tall enough '
               'to show all of it. Spreadsheets keep the line breaks of the cell too.')
    w = cell(wrapped, 'ceWrap')
    w.addElement(P(text='A second paragraph in the same cell.'))
    notes.addElement(row([w, cell('Italic', 'ceItalic')], 'roTall'))
    link = TableCell(valuetype='string')
    lp = P(text='See ')
    lp.addElement(A(href='https://example.com/', type='simple', text='the project website'))
    link.addElement(lp)
    notes.addElement(row([link]))
    for item in ('• First item', '• Second item', '• Third item'):
        notes.addElement(row([cell(item, 'ceIndent')]))
    notes.addElement(row([cell('Struck through', 'ceStrike'), cell('Rotated', 'ceRotated')], 'roTurned'))
    doc.spreadsheet.addElement(notes)

    # ---- Data: a longer table with an autofilter
    data = Table(name='Data', stylename=ta)
    data.addElement(TableColumn(stylename='coDef', defaultcellstylename='Default', numbercolumnsrepeated=4))
    data.addElement(row([cell(h, 'ceHead') for h in ('ID', 'Name', 'Score', 'Ratio')]))
    for i in range(1, 401):
        data.addElement(row([cell(i), cell(f'Item {i}'), cell(random.randint(1, 999)), cell(round(random.random(), 6))]))
    if with_data:
        doc.spreadsheet.addElement(data)

    # ---- Lookup (hidden): the region codes the Sales sheet looks up
    lookup = Table(name='Lookup', stylename=ta_hidden)
    lookup.addElement(TableColumn(stylename='coDef', defaultcellstylename='Default', numbercolumnsrepeated=2))
    for k, region in enumerate(regions):
        lookup.addElement(row([cell(region), cell(k + 1)]))
    doc.spreadsheet.addElement(lookup)

    ne = NamedExpressions()
    ne.addElement(NamedRange(name='Revenue', cellrangeaddress=f'$Sales.$E$3:.$E${sales_rows + 2}', basecelladdress='$Sales.$E$3'))
    doc.spreadsheet.addElement(ne)
    dbr = DatabaseRanges()
    dbr.addElement(DatabaseRange(name='__Anonymous_Sheet_DB__2', targetrangeaddress='Data.A1:Data.D401', displayfilterbuttons='true'))
    if with_data:
        doc.spreadsheet.addElement(dbr)

    # ---- settings: Sales frozen at B3, Notes with a red tab and no grid, Sales active


    def item(name, kind, value):
        c = ConfigItem(name=name, type=kind)
        c.addText(str(value))
        return c


    vs = ConfigItemSet(name='ooo:view-settings')
    views = ConfigItemMapIndexed(name='Views')
    view = ConfigItemMapEntry()
    view.addElement(item('ViewId', 'string', 'view1'))
    tables = ConfigItemMapNamed(name='Tables')
    for name, items in (
        ('Sales', [('HorizontalSplitMode', 'short', 2), ('VerticalSplitMode', 'short', 2), ('HorizontalSplitPosition', 'int', 1),
                   ('VerticalSplitPosition', 'int', 2), ('ActiveSplitRange', 'short', 2), ('PositionLeft', 'int', 0),
                   ('PositionRight', 'int', 1), ('PositionTop', 'int', 0), ('PositionBottom', 'int', 2), ('ShowGrid', 'boolean', 'true')]),
        ('Notes', [('TabColor', 'int', 0xC00000), ('ShowGrid', 'boolean', 'false')]),
    ):
        e = ConfigItemMapEntry(name=name)
        for it in items:
            e.addElement(item(*it))
        tables.addElement(e)
    view.addElement(tables)
    view.addElement(item('ActiveTable', 'string', 'Sales'))
    view.addElement(item('ShowGrid', 'boolean', 'true'))
    views.addElement(view)
    vs.addElement(views)
    doc.settings.addElement(vs)
    return doc


build(60, True).save(f'{out}/sample.ods')
# the flat XML one: fewer sales rows and no Data table, to stay small
with open(f'{out}/sample.fods', 'wb') as f:
    f.write(build(12, False).xml())
print('wrote', f'{out}/sample.ods', f'{out}/sample.fods')
