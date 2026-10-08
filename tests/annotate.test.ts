import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { getSchema, type JSONContent } from '@tiptap/core'
import { EditorState } from '@tiptap/pm/state'
import { extensions } from '../src/extensions'
import { importDocx } from '../src/docx/import'
import { annotateDocx } from '../src/docx/annotate'
import { exportDocx } from '../src/docx/export'
import { annotationsToJson, annotationsToMarkdown, collectAnnotations, placeable } from '../src/annotations'

const schema = getSchema(extensions)
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'

const run = (t: string, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`

/** A document with features Darkwrote does not model: header, footnotes, text box, custom parts. */
async function sampleDocx(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://example/header" Target="header1.xml"/></Relationships>')
  zip.file('word/header1.xml', `<w:hdr ${NS}><w:p>${run('Running header')}</w:p></w:hdr>`)
  zip.file('word/footnotes.xml', `<w:footnotes ${NS}><w:footnote w:id="1"><w:p>${run('A footnote')}</w:p></w:footnote></w:footnotes>`)
  zip.file('customXml/item1.xml', '<root>keep me</root>')
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
<w:p><w:r><mc:AlternateContent><mc:Choice><w:drawing><w:txbxContent><w:p>${run('inside a text box')}</w:p></w:txbxContent></w:drawing></mc:Choice></mc:AlternateContent></w:r>${run('First paragraph')}</w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('Chapter One')}</w:p>
<w:p>${run('The quick ', '<w:b/>')}${run('brown fox', '<w:i/><w:color w:val="FF0000"/>')}<w:r><w:tab/></w:r>${run(' jumps over the lazy dog.')}<w:del w:id="9" w:author="x"><w:r><w:delText>deleted</w:delText></w:r></w:del></w:p>
<w:p>${run('Second paragraph, ')}<w:hyperlink w:anchor="x">${run('with a link')}</w:hyperlink>${run(' inside.')}</w:p>
<w:tbl><w:tr><w:tc><w:p>${run('cell text here')}</w:p></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'uint8array' })
}

function findText(doc: ReturnType<typeof schema.nodeFromJSON>, needle: string) {
  let hit: { from: number; to: number } | null = null
  doc.descendants((node, pos) => {
    if (hit || !node.isText) return
    const i = node.text!.indexOf(needle)
    if (i >= 0) hit = { from: pos + i, to: pos + i + needle.length }
  })
  if (!hit) throw new Error('text not found: ' + needle)
  return hit as { from: number; to: number }
}

function plainText(doc: JSONContent): string[] {
  const out: string[] = []
  const walk = (n: JSONContent) => {
    if (n.type === 'paragraph' || n.type === 'heading') out.push((n.content ?? []).map((c) => c.text ?? (c.type === 'hardBreak' ? '\n' : '')).join(''))
    n.content?.forEach(walk)
  }
  walk(doc)
  return out
}

describe('annotated export of the original docx', () => {
  it('adds comments and highlights without touching unsupported parts', async () => {
    const original = await sampleDocx()
    const { doc: json } = await importDocx(original)
    let state = EditorState.create({ schema, doc: schema.nodeFromJSON(json) })

    // a comment across two formatting runs + the tab, a highlight, and a comment across paragraphs
    const c1 = findText(state.doc, 'brown fox')
    const hl = findText(state.doc, 'lazy dog')
    const c2a = findText(state.doc, 'Second paragraph')
    const c2b = findText(state.doc, 'with a link')
    state = state.apply(state.tr.addMark(c1.from, c1.to, schema.marks.comment.create({ commentId: 'n1' })))
    state = state.apply(state.tr.addMark(hl.from, hl.to, schema.marks.highlight.create({ color: '#ffff00', user: true })))
    state = state.apply(state.tr.addMark(c2a.from, c2b.to, schema.marks.comment.create({ commentId: 'n2' })))

    const records = collectAnnotations(state.doc, [
      { id: 'n1', author: 'Reader One', date: '2026-01-02T03:04:05.000Z', text: 'Nice phrase.\nSecond line' },
      { id: 'n2', author: 'Reader One', date: '2026-01-02T03:04:05.000Z', text: 'Spans a hyperlink' },
    ])
    expect(records.map((r) => r.kind)).toEqual(['comment', 'highlight', 'comment'])
    expect(records[0].heading).toBe('Chapter One')
    expect(records.every((r) => r.anchor)).toBe(true)

    const result = await annotateDocx(original, placeable(records))
    expect(result.skipped).toEqual([])
    expect(result.placed).toHaveLength(3)

    // untouched parts are byte-for-byte identical
    const before = await JSZip.loadAsync(original)
    const after = await JSZip.loadAsync(await result.blob.arrayBuffer())
    for (const name of ['word/header1.xml', 'word/footnotes.xml', 'customXml/item1.xml', '_rels/.rels', 'word/_rels/document.xml.rels'].filter((n) => n !== 'word/_rels/document.xml.rels')) {
      expect(await after.file(name)!.async('string')).toBe(await before.file(name)!.async('string'))
    }
    expect(await after.file('word/_rels/document.xml.rels')!.async('string')).toContain('comments.xml')
    expect(await after.file('[Content_Types].xml')!.async('string')).toContain('/word/comments.xml')
    // the text box and tracked deletion survive in document.xml
    const docXml = await after.file('word/document.xml')!.async('string')
    expect(docXml).toContain('inside a text box')
    expect(docXml).toContain('deleted')
    expect(docXml).toContain('mc:Ignorable="w14"')

    // re-import: text unchanged, comments + highlight anchored on the right passages
    const reimport = await importDocx(await result.blob.arrayBuffer())
    expect(plainText(reimport.doc)).toEqual(plainText(json))
    expect(reimport.comments.map((c) => [c.text, c.author])).toEqual([
      ['Nice phrase.\nSecond line', 'Reader One'],
      ['Spans a hyperlink', 'Reader One'],
    ])
    const rstate = EditorState.create({ schema, doc: schema.nodeFromJSON(reimport.doc) })
    const rrecords = collectAnnotations(rstate.doc, reimport.comments)
    expect(rrecords.find((r) => r.text?.startsWith('Nice'))!.quote).toBe('brown fox')
    expect(rrecords.find((r) => r.text?.startsWith('Spans'))!.quote).toBe('Second paragraph, with a link')
    // formatting of the original runs is preserved on the split runs
    const texts: JSONContent[] = []
    const gather = (n: JSONContent) => (n.type === 'text' ? texts.push(n) : n.content?.forEach(gather))
    gather(reimport.doc)
    const fox = texts.find((t) => t.text === 'brown fox')!
    expect(fox.marks!.map((m) => m.type)).toEqual(expect.arrayContaining(['italic', 'comment', 'textStyle']))
    expect(fox.marks!.find((m) => m.type === 'textStyle')!.attrs!.color).toBe('#ff0000')
    const lazy = texts.find((t) => t.text === 'lazy dog')!
    expect(lazy.marks!.find((m) => m.type === 'highlight')!.attrs!.color).toBe('#ffff00')
    expect(docXml).toContain('w:highlight')
  })

  it('skips annotations whose passage no longer matches the original', async () => {
    const original = await sampleDocx()
    const result = await annotateDocx(original, [
      { kind: 'comment', id: 'x', anchor: { startPara: 2, startOffset: 0, endPara: 2, endOffset: 3 }, quote: 'NOT THE TEXT', text: 'hi', author: 'A' },
      { kind: 'comment', id: 'y', anchor: { startPara: 99, startOffset: 0, endPara: 99, endOffset: 3 }, quote: 'x', text: 'hi', author: 'A' },
    ])
    expect(result.placed).toHaveLength(0)
    expect(result.skipped).toHaveLength(2)
  })

  it('exports annotations as JSON and Markdown', async () => {
    const original = await sampleDocx()
    const { doc: json } = await importDocx(original)
    let state = EditorState.create({ schema, doc: schema.nodeFromJSON(json) })
    const r = findText(state.doc, 'quick')
    state = state.apply(state.tr.addMark(r.from, r.to, schema.marks.comment.create({ commentId: 'n1' })))
    const records = collectAnnotations(state.doc, [{ id: 'n1', author: 'Me', date: '2026-02-03T00:00:00Z', text: 'Why quick?' }])
    const info = { name: 'sample.docx', sha256: 'abc', size: original.length }
    const parsed = JSON.parse(annotationsToJson(records, info))
    expect(parsed).toMatchObject({ format: 'darkwrote-annotations', version: 1, document: info })
    expect(parsed.annotations[0]).toMatchObject({ kind: 'comment', quote: 'quick', text: 'Why quick?', heading: 'Chapter One' })
    expect(parsed.annotations[0].anchor).toEqual({ startPara: 3, startOffset: 4, endPara: 3, endOffset: 9 })
    const md = annotationsToMarkdown(records, info)
    expect(md).toContain('## Chapter One')
    expect(md).toContain('> quick')
    expect(md).toContain('**Me** (2026-02-03): Why quick?')
  })
})

describe('long manuscripts', () => {
  it('imports, annotates and re-imports a 3,000-paragraph book quickly', async () => {
    const para = (i: number): JSONContent => ({ type: 'paragraph', content: [{ type: 'text', text: `Paragraph ${i}: ` + 'lorem ipsum dolor sit amet '.repeat(12) }] })
    const doc: JSONContent = { type: 'doc', content: Array.from({ length: 3000 }, (_, i) => (i % 50 === 0 ? { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: `Chapter ${i / 50}` }] } : para(i))) }
    const original = await (await exportDocx(doc)).arrayBuffer()
    const t0 = performance.now()
    const { doc: json } = await importDocx(original)
    let state = EditorState.create({ schema, doc: schema.nodeFromJSON(json) })
    const comments: { id: string; author: string; date: string; text: string }[] = []
    for (let i = 0; i < 60; i++) {
      const r = findText(state.doc, `Paragraph ${i * 40 + 7}: lorem`)
      state = state.apply(state.tr.addMark(r.from, r.to, schema.marks.comment.create({ commentId: `n${i}` })))
      comments.push({ id: `n${i}`, author: 'R', date: '2026-01-01T00:00:00Z', text: `note ${i}` })
    }
    const result = await annotateDocx(original, placeable(collectAnnotations(state.doc, comments)))
    expect(result.placed).toHaveLength(60)
    expect(result.skipped).toHaveLength(0)
    const again = await importDocx(await result.blob.arrayBuffer())
    expect(again.comments).toHaveLength(60)
    expect(performance.now() - t0).toBeLessThan(15000)
  }, 30000)
})
