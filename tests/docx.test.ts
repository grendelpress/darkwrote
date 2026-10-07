import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import type { JSONContent } from '@tiptap/core'
import { importDocx } from '../src/docx/import'
import { exportDocx } from '../src/docx/export'

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

const find = (n: JSONContent, type: string, out: JSONContent[] = []): JSONContent[] => {
  if (n.type === type) out.push(n)
  n.content?.forEach((c) => find(c, type, out))
  return out
}

async function roundTrip(doc: JSONContent) {
  const blob = await exportDocx(doc)
  return importDocx(await blob.arrayBuffer())
}

describe('docx round trip', () => {
  it('keeps inline formatting', async () => {
    const doc: JSONContent = {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          attrs: { textAlign: 'center' },
          content: [
            { type: 'text', text: 'Bold', marks: [{ type: 'bold' }] },
            { type: 'text', text: 'Fancy', marks: [{ type: 'italic' }, { type: 'underline' }, { type: 'textStyle', attrs: { color: '#ff0000', fontFamily: 'Georgia', fontSize: '14pt' } }] },
            { type: 'text', text: 'Hi', marks: [{ type: 'highlight', attrs: { color: '#ffff00' } }, { type: 'strike' }] },
            { type: 'text', text: 'sup', marks: [{ type: 'superscript' }] },
            { type: 'text', text: 'site', marks: [{ type: 'link', attrs: { href: 'https://example.com/' } }] },
          ],
        },
        { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Title' }] },
      ],
    }
    const { doc: out } = await roundTrip(doc)
    const [p, h] = out.content!
    expect(p.attrs?.textAlign).toBe('center')
    const byText = (t: string) => p.content!.find((n) => n.text === t)!
    const types = (t: string) => byText(t).marks!.map((m) => m.type)
    expect(types('Bold')).toContain('bold')
    expect(types('Fancy')).toEqual(expect.arrayContaining(['italic', 'underline', 'textStyle']))
    const style = byText('Fancy').marks!.find((m) => m.type === 'textStyle')!.attrs!
    expect(style).toMatchObject({ color: '#ff0000', fontFamily: 'Georgia', fontSize: '14pt' })
    expect(byText('Hi').marks!.find((m) => m.type === 'highlight')!.attrs!.color).toBe('#ffff00')
    expect(types('Hi')).toContain('strike')
    expect(types('sup')).toContain('superscript')
    expect(byText('site').marks!.find((m) => m.type === 'link')!.attrs!.href).toBe('https://example.com/')
    expect(h.type).toBe('heading')
    expect(h.attrs?.level).toBe(2)
  })

  it('keeps nested lists, tables and images', async () => {
    const para = (text: string): JSONContent => ({ type: 'paragraph', content: [{ type: 'text', text }] })
    const doc: JSONContent = {
      type: 'doc',
      content: [
        {
          type: 'bulletList',
          content: [
            { type: 'listItem', content: [para('a'), { type: 'orderedList', content: [{ type: 'listItem', content: [para('a1')] }, { type: 'listItem', content: [para('a2')] }] }] },
            { type: 'listItem', content: [para('b')] },
          ],
        },
        {
          type: 'table',
          content: [
            { type: 'tableRow', content: [{ type: 'tableCell', attrs: { colspan: 2, rowspan: 1, backgroundColor: '#ccddee' }, content: [para('wide')] }] },
            { type: 'tableRow', content: [{ type: 'tableCell', attrs: { colspan: 1, rowspan: 1 }, content: [para('x')] }, { type: 'tableCell', attrs: { colspan: 1, rowspan: 1 }, content: [para('y')] }] },
          ],
        },
        { type: 'paragraph', content: [{ type: 'image', attrs: { src: PNG, width: 40, height: 20 } }] },
      ],
    }
    const { doc: out } = await roundTrip(doc)
    const lists = find(out, 'bulletList')
    expect(lists).toHaveLength(1)
    expect(lists[0].content).toHaveLength(2)
    expect(find(lists[0], 'orderedList')[0].content).toHaveLength(2)
    const cells = find(out, 'tableCell')
    expect(cells).toHaveLength(3)
    expect(cells[0].attrs).toMatchObject({ colspan: 2, backgroundColor: '#ccddee' })
    const img = find(out, 'image')[0]
    expect(img.attrs!.src).toMatch(/^data:image\/png;base64,/)
    expect(img.attrs).toMatchObject({ width: 40, height: 20 })
  })

  it('keeps page geometry', async () => {
    const blob = await exportDocx({ type: 'doc', content: [{ type: 'paragraph' }] }, {
      pageWidth: 11906, pageHeight: 16838, marginTop: 1000, marginRight: 900, marginBottom: 1000, marginLeft: 900,
    })
    const { meta } = await importDocx(await blob.arrayBuffer())
    expect(meta).toMatchObject({ pageWidth: 11906, pageHeight: 16838, marginLeft: 900 })
  })
})

describe('importing Word-authored features', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  async function build(body: string, styles = '', numbering = '') {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
    zip.file('word/document.xml', `<w:document ${W}><w:body>${body}</w:body></w:document>`)
    if (styles) zip.file('word/styles.xml', `<w:styles ${W}>${styles}</w:styles>`)
    if (numbering) zip.file('word/numbering.xml', `<w:numbering ${W}>${numbering}</w:numbering>`)
    return importDocx(await zip.generateAsync({ type: 'uint8array' }))
  }

  it('maps heading styles, style-level formatting and run overrides', async () => {
    const { doc } = await build(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Head</w:t></w:r></w:p>
       <w:p><w:pPr><w:pStyle w:val="Fancy"/><w:jc w:val="both"/></w:pPr><w:r><w:t>Styled</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:color w:val="auto"/></w:rPr><w:t>Plain</w:t></w:r></w:p>`,
      `<w:style w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
       <w:style w:styleId="Fancy"><w:name w:val="Fancy"/><w:rPr><w:b/><w:color w:val="1F3864"/></w:rPr></w:style>`,
    )
    expect(doc.content![0]).toMatchObject({ type: 'heading', attrs: { level: 1 } })
    const p = doc.content![1]
    expect(p.attrs?.textAlign).toBe('justify')
    const marks = (t: string) => p.content!.find((n) => n.text === t)!.marks ?? []
    expect(marks('Styled').map((m) => m.type)).toContain('bold')
    expect(marks('Styled').find((m) => m.type === 'textStyle')!.attrs!.color).toBe('#1f3864')
    expect(marks('Plain').map((m) => m.type)).not.toContain('bold')
    expect(marks('Plain').find((m) => m.type === 'textStyle')).toBeUndefined()
  })

  it('builds bullet vs numbered lists from numbering.xml', async () => {
    const item = (id: string, lvl: number, t: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${id}"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`
    const { doc } = await build(
      item('1', 0, 'one') + item('1', 1, 'nested') + item('1', 0, 'two') + '<w:p><w:r><w:t>gap</w:t></w:r></w:p>' + item('2', 0, 'bullet'),
      '',
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/></w:lvl></w:abstractNum>
       <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
       <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`,
    )
    expect(doc.content!.map((n) => n.type)).toEqual(['orderedList', 'paragraph', 'bulletList'])
    const ol = doc.content![0]
    expect(ol.content).toHaveLength(2)
    expect(find(ol.content![0], 'orderedList')[0]).toBeDefined()
  })

  it('handles vertical merges and cell shading', async () => {
    const tc = (inner: string, pr = '') => `<w:tc><w:tcPr>${pr}</w:tcPr><w:p><w:r><w:t>${inner}</w:t></w:r></w:p></w:tc>`
    const { doc } = await build(
      `<w:tbl><w:tr>${tc('a', '<w:vMerge w:val="restart"/><w:shd w:fill="FFFF00"/>')}${tc('b')}</w:tr><w:tr>${tc('', '<w:vMerge/>')}${tc('c')}</w:tr></w:tbl>`,
    )
    const cells = find(doc, 'tableCell')
    expect(cells).toHaveLength(3)
    expect(cells[0].attrs).toMatchObject({ rowspan: 2, backgroundColor: '#ffff00' })
  })

  it('rejects non-docx input with a friendly error', async () => {
    await expect(importDocx(new TextEncoder().encode('hello'))).rejects.toThrow(/not a \.docx/i)
  })
})
