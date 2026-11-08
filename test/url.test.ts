import { describe, expect, test } from 'bun:test'
import { markdownToHtml } from '../src/core/convert'
import { htmlToParas } from '../src/core/html'
import { safeHref } from '../src/core/url'

describe('safeHref', () => {
  test('web, mail, phone, FTP, relative and in-document links are kept as they are', () => {
    for (const u of ['https://example.org/a?b=1#c', 'HTTP://X.Y', 'mailto:a@b.c', 'tel:+84123', 'sms:+84123', 'ftp://f.example', '#_Toc1', 'other.docx', '../a/b.html', '//cdn.example/x', '?q=1'])
      expect(safeHref(u)).toBe(u)
  })

  test('script, data and other schemes are dropped, however they are disguised', () => {
    for (const u of [
      'javascript:alert(1)',
      'JavaScript:alert(1)',
      '  javascript:alert(1)',
      '\u0000javascript:alert(1)',
      'java\tscript:alert(1)',
      'java\nscript:alert(1)',
      'jav\r\nascript:alert(1)',
      'vbscript:msgbox(1)',
      'data:text/html,<script>alert(1)</script>',
      'file:///etc/passwd',
      'C:\\Users\\a.docx',
      'blob:https://x/1',
    ])
      expect(safeHref(u)).toBeUndefined()
    expect(safeHref('')).toBeUndefined()
    expect(safeHref(undefined)).toBeUndefined()
    expect(safeHref(null)).toBeUndefined()
  })

  test('a control character inside the scheme makes it a relative path, as in a browser', () => {
    expect(safeHref('java\u0000script:alert(1)')).toBe('java\u0000script:alert(1)')
  })
})

describe('imported links use the same rule', () => {
  const hrefs = (html: string) => htmlToParas(html).flatMap((p) => p.runs.map((r) => (r.t === 'text' ? r.href : undefined)))

  test('HTML and Markdown keep the text of a script or data link, not the link', () => {
    expect(hrefs('<p><a href="https://ok.example">a</a><a href="data:text/html,x">b</a><a href="vbscript:x">c</a><a href="java&#9;script:x">d</a></p>')).toEqual([
      'https://ok.example',
      undefined,
      undefined,
      undefined,
    ])
    const md = markdownToHtml('[a](https://ok.example) [b](data:text/html,x) [c](javascript:x)')
    expect(md).toContain('<a href="https://ok.example">a</a>')
    expect(md).not.toContain('data:')
    expect(md).not.toContain('javascript:')
  })
})
