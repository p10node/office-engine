"""Writes the committed .xls fixtures (test/fixtures/legacy/) with xlwt 1.3.0."""
import datetime, random, sys
import xlwt

out = sys.argv[1]

# ---------------------------------------------------------------- sample.xls: the playground sample
wb = xlwt.Workbook(encoding='utf-8')
title = xlwt.easyxf('font: bold on, height 280, colour dark_blue; align: horiz centre, vert centre')
head = xlwt.easyxf('font: bold on, colour white; pattern: pattern solid, fore_colour dark_blue; '
                   'borders: bottom medium, bottom_colour black; align: horiz centre')
text = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin')
date = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin', num_format_str='yyyy-mm-dd')
units = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin; align: horiz right')
price = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin', num_format_str='"$"#,##0.00')
money = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin', num_format_str='#,##0.00')
pct = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin', num_format_str='0.0%')
total = xlwt.easyxf('font: bold on; pattern: pattern solid, fore_colour light_yellow; borders: top double', num_format_str='#,##0.00')
neg = xlwt.easyxf(num_format_str='#,##0.00;[Red](#,##0.00)')
when = xlwt.easyxf(num_format_str='dddd, mmmm d, yyyy h:mm AM/PM')

ws = wb.add_sheet('Sales')
ws2 = wb.add_sheet('Notes')
ws3 = wb.add_sheet('Data')
ws.write_merge(0, 0, 0, 5, 'Sales by region — FY2026', title)
ws.row(0).height_mismatch = True
ws.row(0).height = 600
for c, h in enumerate(['Region', 'Date', 'Units', 'Price', 'Revenue', 'Margin']):
    ws.write(1, c, h, head)
random.seed(7)
regions = ['North', 'Central', 'South', 'East', 'West']
n = 24
for i in range(n):
    r = i + 2
    ws.write(r, 0, random.choice(regions), text)
    ws.write(r, 1, datetime.date(2026, 1, 16) + datetime.timedelta(days=5 * i), date)
    ws.write(r, 2, random.randint(10, 500), units)
    ws.write(r, 3, round(random.uniform(4, 99), 2), price)
    ws.write(r, 4, xlwt.Formula('C%d*D%d' % (r + 1, r + 1)), money)
    ws.write(r, 5, round(random.uniform(-0.05, 0.6), 4), pct)
ws.write(n + 2, 0, 'Total', total)
ws.write(n + 2, 2, xlwt.Formula('SUM(C3:C%d)' % (n + 2)), total)
ws.write(n + 2, 4, xlwt.Formula('SUM(E3:E%d)' % (n + 2)), total)
ws.write(n + 2, 5, xlwt.Formula('IF(E%d>0,AVERAGE(F3:F%d),0)' % (n + 3, n + 2)), pct)
ws.write(1, 7, 'Notes', head)
ws.write(2, 7, 1234567.891)
ws.write(3, 7, -42.5, neg)
ws.write(4, 7, True)
ws.write(5, 7, datetime.datetime(2023, 12, 31, 18, 0), when)
ws.write(6, 7, xlwt.Formula('ROUND(AVERAGE(D3:D%d);2)' % (n + 2)), price)
ws.write(7, 7, xlwt.Formula('Notes!B1&" "&Data!B2'))
ws.write(8, 7, xlwt.Formula('VLOOKUP("North";A3:F%d;3;FALSE)' % (n + 2)))
ws.write(9, 7, xlwt.Formula('UPPER(A3)&" / "&TEXT(B3;"mmm d")'))
ws.col(0).width = 14 * 256
ws.col(1).width = 13 * 256
ws.col(4).width = 14 * 256
ws.col(7).width = 30 * 256
ws.set_panes_frozen(True)
ws.set_horz_split_pos(2)
ws.set_vert_split_pos(1)
ws.set_remove_splits(True)
ws.set_portrait(False)
ws.set_paper_size_code(9)
ws.set_left_margin(0.5)
ws.set_right_margin(0.5)
ws.set_header_str(b'&C&BSales by region')
ws.set_footer_str(b'&LConfidential&RPage &P of &N')
ws.set_fit_num_pages(1)
ws.set_fit_width_to_pages(1)
ws.set_fit_height_to_pages(0)

wrap = xlwt.easyxf('align: wrap on, vert top')
ws2.write(0, 0, 'Wrapped text cell. A long note that wraps over several lines in a narrow column, '
                 'so the row grows to show all of it.', wrap)
ws2.write(0, 1, 'Italic', xlwt.easyxf('font: italic on'))
ws2.write(1, 0, 'Bold red', xlwt.easyxf('font: bold on, colour red'))
ws2.write(1, 1, 'Underlined', xlwt.easyxf('font: underline single'))
ws2.write(2, 0, 'Struck', xlwt.easyxf('font: struck_out on'))
ws2.write(2, 1, 'Centered on green', xlwt.easyxf('align: horiz centre; pattern: pattern solid, fore_colour light_green'))
ws2.write(3, 0, 'Rotated', xlwt.easyxf('align: rotation 45'))
ws2.write(3, 1, 'Indented', xlwt.easyxf('align: horiz left, indent 2'))
ws2.write_rich_text(4, 0, ('Rich: ', ('bold', xlwt.easyfont('bold true')), ' and ', ('blue italic', xlwt.easyfont('italic true, colour blue'))))
ws2.write(5, 0, 'Hidden row')
ws2.row(5).hidden = True
ws2.write(6, 0, 'Times New Roman 14', xlwt.easyxf('font: name Times New Roman, height 280'))
ws2.col(0).width = 50 * 256
ws2.col(2).hidden = True
ws2.write(0, 2, 'Hidden column')

for c, h in enumerate(['ID', 'Name', 'Score', 'Ratio']):
    ws3.write(0, c, h, xlwt.easyxf('font: bold on'))
random.seed(11)
for i in range(1, 201):
    ws3.write(i, 0, i)
    ws3.write(i, 1, 'Item %d' % i)
    ws3.write(i, 2, random.randint(0, 999))
    ws3.write(i, 3, random.random())
ws3.write(202, 1, 'Average')
ws3.write(202, 2, xlwt.Formula('AVERAGE(C2:C201)'), xlwt.easyxf(num_format_str='0.00'))
wb.set_active_sheet(0)
wb.save(out + '/sample.xls')

# ---------------------------------------------------------------- strings.xls: a shared string table over CONTINUE records
wb = xlwt.Workbook(encoding='utf-8')
ws = wb.add_sheet('Strings')
words = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta']
for i in range(300):
    # 8-bit and 16-bit strings mixed, so CONTINUE boundaries fall inside both kinds
    s = '%04d %s' % (i, ' '.join(words[(i + k) % 8] for k in range(4)))
    if i % 3 == 0:
        s += ' été 中文 Ω'
    ws.write(i, 0, s)
long = ''.join(words[k % 8] + ' ' for k in range(1500))
ws.write(0, 1, long)
ws.write(1, 1, '中文' * 2500)
wb.save(out + '/strings.xls')

# ---------------------------------------------------------------- dates1904.xls: the 1904 date system
wb = xlwt.Workbook()
wb.dates_1904 = True
ws = wb.add_sheet('Dates')
ws.write(0, 0, datetime.date(2024, 2, 29), xlwt.easyxf(num_format_str='yyyy-mm-dd'))
ws.write(1, 0, datetime.datetime(2001, 1, 1, 12, 30), xlwt.easyxf(num_format_str='d mmm yyyy hh:mm'))
wb.save(out + '/dates1904.xls')
