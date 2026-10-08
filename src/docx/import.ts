import JSZip from 'jszip'
import type { JSONContent } from '@tiptap/core'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** Page geometry in twips (1/20 pt). */
export interface DocMeta {
  pageWidth: number
  pageHeight: number
  marginTop: number
  marginRight: number
  marginBottom: number
  marginLeft: number
}

export const DEFAULT_META: DocMeta = {
  pageWidth: 12240,
  pageHeight: 15840,
  marginTop: 1440,
  marginRight: 1440,
  marginBottom: 1440,
  marginLeft: 1440,
}

export interface CommentData {
  id: string
  author: string
  date: string
  text: string
  /** True for comments that already exist in the source .docx (they are not re-added by annotated export). */
  imported?: boolean
}

export interface ImportResult {
  doc: JSONContent
  meta: DocMeta
  comments: CommentData[]
  warnings: string[]
}

const HIGHLIGHTS: Record<string, string> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
  white: '#ffffff',
}

interface RunProps {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  vertAlign?: 'superscript' | 'subscript' | null
  color?: string | null
  highlight?: string | null
  size?: number | null
  font?: string | null
}

interface StyleInfo {
  id: string
  name: string
  basedOn?: string
  rPr?: Element
  pPr?: Element
}

// ---------- XML helpers ----------

function parseXml(text: string): Document {
  return new DOMParser().parseFromString(text, 'application/xml')
}

function kids(el: Element | null | undefined, local: string): Element[] {
  if (!el) return []
  return Array.from(el.children).filter((c) => c.localName === local)
}

function kid(el: Element | null | undefined, local: string): Element | undefined {
  return kids(el, local)[0]
}

function wAttr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null
  return el.getAttributeNS(W, name) ?? el.getAttribute('w:' + name)
}

function onOff(el: Element | undefined): boolean | undefined {
  if (!el) return undefined
  const v = wAttr(el, 'val')
  return !(v === '0' || v === 'false' || v === 'off')
}

function normColor(v: string | null): string | null {
  if (!v || v === 'auto') return null
  return /^[0-9a-fA-F]{6}$/.test(v) ? '#' + v.toLowerCase() : null
}

// ---------- Importer ----------

class Importer {
  styles = new Map<string, StyleInfo>()
  numbering = new Map<string, boolean[]>() // numId -> [ordered per level]
  rels = new Map<string, { target: string; external: boolean }>()
  media = new Map<string, string>() // zip path -> data URL
  warnings: string[] = []
  comments = new Map<string, CommentData>()
  /** Comment ranges currently open (they may span paragraphs) */
  activeComments = new Set<string>()
  anchoredComments = new Set<string>()
  /** Index of every w:p in document order; paragraph nodes remember it so annotations can be mapped back to the source XML. */
  paraIndex = new Map<Element, number>()
  defaultRun: RunProps = {}

  constructor(private zip: JSZip) {}

  async load(): Promise<Document> {
    const docFile = this.zip.file('word/document.xml')
    if (!docFile) throw new Error('Not a valid .docx file (word/document.xml is missing).')
    const doc = parseXml(await docFile.async('string'))
    Array.from(doc.getElementsByTagNameNS(W, 'p')).forEach((p, i) => this.paraIndex.set(p, i))

    const stylesFile = this.zip.file('word/styles.xml')
    if (stylesFile) this.readStyles(parseXml(await stylesFile.async('string')))
    const numFile = this.zip.file('word/numbering.xml')
    if (numFile) this.readNumbering(parseXml(await numFile.async('string')))
    const relFile = this.zip.file('word/_rels/document.xml.rels')
    if (relFile) this.readRels(parseXml(await relFile.async('string')))
    const commentsFile = this.zip.file('word/comments.xml')
    if (commentsFile) this.readComments(parseXml(await commentsFile.async('string')))
    await this.loadMedia()
    return doc
  }

  readStyles(xml: Document) {
    const defaults = kid(kid(xml.documentElement, 'docDefaults'), 'rPrDefault')
    if (defaults) this.defaultRun = this.parseRPr(kid(defaults, 'rPr'))
    for (const s of kids(xml.documentElement, 'style')) {
      const id = wAttr(s, 'styleId')
      if (!id) continue
      this.styles.set(id, {
        id,
        name: wAttr(kid(s, 'name'), 'val') ?? id,
        basedOn: wAttr(kid(s, 'basedOn'), 'val') ?? undefined,
        rPr: kid(s, 'rPr'),
        pPr: kid(s, 'pPr'),
      })
    }
  }

  readNumbering(xml: Document) {
    const abstracts = new Map<string, boolean[]>()
    for (const an of kids(xml.documentElement, 'abstractNum')) {
      const id = wAttr(an, 'abstractNumId')
      if (id == null) continue
      const levels: boolean[] = []
      for (const lvl of kids(an, 'lvl')) {
        const i = Number(wAttr(lvl, 'ilvl') ?? 0)
        const fmt = wAttr(kid(lvl, 'numFmt'), 'val') ?? 'bullet'
        levels[i] = fmt !== 'bullet' && fmt !== 'none'
      }
      abstracts.set(id, levels)
    }
    for (const num of kids(xml.documentElement, 'num')) {
      const id = wAttr(num, 'numId')
      const abs = wAttr(kid(num, 'abstractNumId'), 'val')
      if (id != null && abs != null && abstracts.has(abs)) this.numbering.set(id, abstracts.get(abs)!)
    }
  }

  readComments(xml: Document) {
    for (const c of kids(xml.documentElement, 'comment')) {
      const id = wAttr(c, 'id')
      if (id == null) continue
      const text = kids(c, 'p')
        .map((p) =>
          Array.from(p.getElementsByTagNameNS(W, 't'))
            .map((t) => t.textContent ?? '')
            .join(''),
        )
        .join('\n')
      this.comments.set(id, {
        id,
        author: wAttr(c, 'author') ?? 'Unknown',
        date: wAttr(c, 'date') ?? '',
        text,
        imported: true,
      })
    }
  }

  readRels(xml: Document) {
    for (const r of Array.from(xml.getElementsByTagName('Relationship'))) {
      const id = r.getAttribute('Id')
      const target = r.getAttribute('Target')
      if (id && target) {
        this.rels.set(id, { target, external: r.getAttribute('TargetMode') === 'External' })
      }
    }
  }

  async loadMedia() {
    const mime: Record<string, string> = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      svg: 'image/svg+xml',
      webp: 'image/webp',
    }
    for (const rel of this.rels.values()) {
      if (rel.external) continue
      const path = rel.target.startsWith('/') ? rel.target.slice(1) : 'word/' + rel.target
      const ext = path.split('.').pop()?.toLowerCase() ?? ''
      if (!mime[ext]) continue
      const f = this.zip.file(path)
      if (!f) continue
      this.media.set(path, `data:${mime[ext]};base64,` + (await f.async('base64')))
    }
  }

  // ----- styles -----

  styleChain(id: string | null | undefined): StyleInfo[] {
    const chain: StyleInfo[] = []
    const seen = new Set<string>()
    let cur = id ? this.styles.get(id) : undefined
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id)
      chain.unshift(cur)
      cur = cur.basedOn ? this.styles.get(cur.basedOn) : undefined
    }
    return chain
  }

  headingLevel(styleId: string | null): number | null {
    // Built-in ids work even when styles.xml is missing or incomplete
    const direct = /^heading(\d)$/i.exec(styleId ?? '')
    if (direct) return Math.min(6, Math.max(1, Number(direct[1])))
    for (const s of this.styleChain(styleId).reverse()) {
      const m = /^heading\s*(\d)$/i.exec(s.name) ?? /^heading(\d)$/i.exec(s.id)
      if (m) return Math.min(6, Math.max(1, Number(m[1])))
      if (/^title$/i.test(s.name)) return 1
    }
    return null
  }

  parseRPr(rPr: Element | undefined): RunProps {
    const p: RunProps = {}
    if (!rPr) return p
    const b = onOff(kid(rPr, 'b'))
    if (b !== undefined) p.bold = b
    const i = onOff(kid(rPr, 'i'))
    if (i !== undefined) p.italic = i
    const s = onOff(kid(rPr, 'strike')) ?? onOff(kid(rPr, 'dstrike'))
    if (s !== undefined) p.strike = s
    const u = kid(rPr, 'u')
    if (u) p.underline = (wAttr(u, 'val') ?? 'single') !== 'none'
    const va = wAttr(kid(rPr, 'vertAlign'), 'val')
    if (va) p.vertAlign = va === 'superscript' || va === 'subscript' ? va : null
    const c = kid(rPr, 'color')
    if (c) p.color = normColor(wAttr(c, 'val'))
    const h = wAttr(kid(rPr, 'highlight'), 'val')
    if (h) p.highlight = HIGHLIGHTS[h] ?? null
    else {
      const shd = kid(rPr, 'shd')
      const fill = normColor(wAttr(shd, 'fill'))
      if (shd && fill) p.highlight = fill
    }
    const sz = wAttr(kid(rPr, 'sz'), 'val')
    if (sz) p.size = Number(sz) / 2
    const f = kid(rPr, 'rFonts')
    if (f) p.font = wAttr(f, 'ascii') ?? wAttr(f, 'hAnsi')
    return p
  }

  /** Effective run properties from the paragraph style chain (headings keep their CSS look). */
  styleRun(styleId: string | null, isHeading: boolean): RunProps {
    const merged: RunProps = isHeading ? {} : { ...this.defaultRun }
    if (isHeading) return merged
    for (const s of this.styleChain(styleId)) Object.assign(merged, this.parseRPr(s.rPr))
    return merged
  }

  styleParagraphProp<T>(styleId: string | null, pick: (pPr: Element) => T | undefined): T | undefined {
    let out: T | undefined
    for (const s of this.styleChain(styleId)) {
      if (!s.pPr) continue
      const v = pick(s.pPr)
      if (v !== undefined) out = v
    }
    return out
  }

  // ----- inline content -----

  marksFor(p: RunProps, href?: string): JSONContent['marks'] {
    const marks: NonNullable<JSONContent['marks']> = []
    if (p.bold) marks.push({ type: 'bold' })
    if (p.italic) marks.push({ type: 'italic' })
    if (p.underline) marks.push({ type: 'underline' })
    if (p.strike) marks.push({ type: 'strike' })
    if (p.vertAlign) marks.push({ type: p.vertAlign })
    if (p.highlight) marks.push({ type: 'highlight', attrs: { color: p.highlight } })
    if (href) marks.push({ type: 'link', attrs: { href } })
    for (const id of this.activeComments) {
      if (this.comments.has(id)) {
        marks.push({ type: 'comment', attrs: { commentId: id } })
        this.anchoredComments.add(id)
      }
    }
    const attrs: Record<string, string> = {}
    if (p.color) attrs.color = p.color
    if (p.font) attrs.fontFamily = p.font
    if (p.size) attrs.fontSize = `${p.size}pt`
    if (Object.keys(attrs).length) marks.push({ type: 'textStyle', attrs })
    return marks.length ? marks : undefined
  }

  pushText(out: JSONContent[], text: string, marks: JSONContent['marks']) {
    if (!text) return
    const last = out[out.length - 1]
    if (last && last.type === 'text' && JSON.stringify(last.marks) === JSON.stringify(marks)) {
      last.text += text
    } else {
      const node: JSONContent = { type: 'text', text }
      if (marks) node.marks = marks
      out.push(node)
    }
  }

  inline(parent: Element, base: RunProps, href: string | undefined, out: JSONContent[]) {
    for (const el of Array.from(parent.children)) {
      switch (el.localName) {
        case 'r':
          this.run(el, base, href, out)
          break
        case 'commentRangeStart':
          this.activeComments.add(wAttr(el, 'id') ?? '')
          break
        case 'commentRangeEnd':
          this.activeComments.delete(wAttr(el, 'id') ?? '')
          break
        case 'hyperlink': {
          const rid = el.getAttributeNS(R, 'id') ?? el.getAttribute('r:id')
          const rel = rid ? this.rels.get(rid) : undefined
          const anchor = wAttr(el, 'anchor')
          const link = rel?.external ? rel.target : anchor ? '#' + anchor : undefined
          this.inline(el, base, link, out)
          break
        }
        case 'ins':
        case 'smartTag':
        case 'sdt':
        case 'sdtContent':
        case 'fldSimple':
        case 'customXml':
          this.inline(el, base, href, out)
          break
        // w:del, bookmarks, proofErrors etc. are skipped
      }
    }
  }

  run(r: Element, base: RunProps, href: string | undefined, out: JSONContent[]) {
    const rStyleId = wAttr(kid(kid(r, 'rPr'), 'rStyle'), 'val')
    let props: RunProps = { ...base }
    for (const s of this.styleChain(rStyleId)) Object.assign(props, this.parseRPr(s.rPr))
    props = { ...props, ...this.parseRPr(kid(r, 'rPr')) }
    const marks = this.marksFor(props, href)

    for (const c of Array.from(r.children)) {
      switch (c.localName) {
        case 't':
        case 'delText':
          if (c.localName === 't') this.pushText(out, c.textContent ?? '', marks)
          break
        case 'tab':
          this.pushText(out, '\t', marks)
          break
        case 'noBreakHyphen':
          this.pushText(out, '‑', marks)
          break
        case 'br':
        case 'cr':
          out.push({ type: 'hardBreak' })
          break
        case 'drawing':
        case 'pict': {
          const img = this.image(c)
          if (img) out.push(img)
          break
        }
      }
    }
  }

  image(el: Element): JSONContent | null {
    let rid: string | null = null
    const blip = el.getElementsByTagNameNS(A, 'blip')[0]
    if (blip) rid = blip.getAttributeNS(R, 'embed') ?? blip.getAttribute('r:embed')
    if (!rid) {
      const vml = Array.from(el.getElementsByTagName('*')).find((e) => e.localName === 'imagedata')
      if (vml) rid = vml.getAttributeNS(R, 'id') ?? vml.getAttribute('r:id')
    }
    const rel = rid ? this.rels.get(rid) : undefined
    if (!rel) return null
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : 'word/' + rel.target
    const src = this.media.get(path)
    if (!src) {
      this.warnings.push(`Unsupported image skipped: ${rel.target}`)
      return null
    }
    const attrs: Record<string, unknown> = { src }
    const extent = el.getElementsByTagNameNS(WP, 'extent')[0]
    if (extent) {
      const cx = Number(extent.getAttribute('cx'))
      const cy = Number(extent.getAttribute('cy'))
      if (cx > 0 && cy > 0) {
        attrs.width = Math.round(cx / 9525)
        attrs.height = Math.round(cy / 9525)
      }
    }
    const docPr = el.getElementsByTagNameNS(WP, 'docPr')[0]
    if (docPr?.getAttribute('descr')) attrs.alt = docPr.getAttribute('descr')
    return { type: 'image', attrs }
  }

  // ----- blocks -----

  paragraph(p: Element): { node: JSONContent; list?: { level: number; ordered: boolean; numId: string } } {
    const pPr = kid(p, 'pPr')
    const styleId = wAttr(kid(pPr, 'pStyle'), 'val')
    const level = this.headingLevel(styleId)
    const base = this.styleRun(styleId, level != null)

    const content: JSONContent[] = []
    this.inline(p, base, undefined, content)

    const align = wAttr(kid(pPr, 'jc'), 'val') ?? this.styleParagraphProp(styleId, (pp) => wAttr(kid(pp, 'jc'), 'val') ?? undefined)
    const textAlign =
      align === 'center'
        ? 'center'
        : align === 'right' || align === 'end'
          ? 'right'
          : align === 'both' || align === 'distribute'
            ? 'justify'
            : null

    const node: JSONContent = level
      ? { type: 'heading', attrs: { level } }
      : { type: 'paragraph' }
    if (textAlign) node.attrs = { ...(node.attrs ?? {}), textAlign }
    node.attrs = { ...(node.attrs ?? {}), srcPara: this.paraIndex.get(p) ?? null }
    if (content.length) node.content = content

    // numbering (direct or via style)
    let numPr = kid(pPr, 'numPr')
    if (!numPr) {
      for (const s of this.styleChain(styleId).reverse()) {
        const np = kid(s.pPr, 'numPr')
        if (np) {
          numPr = np
          break
        }
      }
    }
    const numId = wAttr(kid(numPr, 'numId'), 'val')
    if (numPr && numId && numId !== '0' && !level) {
      const ilvl = Number(wAttr(kid(numPr, 'ilvl'), 'val') ?? 0)
      const ordered = this.numbering.get(numId)?.[ilvl] ?? false
      return { node, list: { level: ilvl, ordered, numId } }
    }
    return { node }
  }

  table(tbl: Element): JSONContent {
    const rows: JSONContent[] = []
    const origins: (JSONContent | undefined)[] = []
    for (const tr of kids(tbl, 'tr')) {
      const cells: JSONContent[] = []
      let col = 0
      for (const tc of kids(tr, 'tc')) {
        const tcPr = kid(tc, 'tcPr')
        const span = Number(wAttr(kid(tcPr, 'gridSpan'), 'val') ?? 1)
        const vm = kid(tcPr, 'vMerge')
        if (vm && wAttr(vm, 'val') !== 'restart') {
          const origin = origins[col]
          if (origin) origin.attrs!.rowspan = (origin.attrs!.rowspan as number) + 1
          col += span
          continue
        }
        const fill = normColor(wAttr(kid(tcPr, 'shd'), 'fill'))
        const cell: JSONContent = {
          type: 'tableCell',
          attrs: { colspan: span, rowspan: 1, backgroundColor: fill },
          content: this.blocks(tc),
        }
        if (!cell.content!.length) cell.content = [{ type: 'paragraph' }]
        cells.push(cell)
        for (let k = 0; k < span; k++) origins[col + k] = cell
        col += span
      }
      if (cells.length) rows.push({ type: 'tableRow', content: cells })
    }
    return { type: 'table', content: rows }
  }

  /** Convert block-level children (of body or a table cell), grouping list paragraphs. */
  blocks(parent: Element): JSONContent[] {
    const out: JSONContent[] = []
    type Frame = { list: JSONContent; level: number; numId: string; ordered: boolean }
    let stack: Frame[] = []

    const addListItem = (para: JSONContent, level: number, ordered: boolean, numId: string) => {
      const type = ordered ? 'orderedList' : 'bulletList'
      while (stack.length > level + 1) stack.pop()
      if (level === 0 && stack.length && stack[0].numId !== numId) stack = []
      if (stack.length === level + 1 && stack[level].ordered !== ordered) stack.length = level
      while (stack.length < level + 1) {
        const list: JSONContent = { type, content: [] }
        const parentFrame = stack[stack.length - 1]
        if (!parentFrame) out.push(list)
        else {
          const items = parentFrame.list.content!
          if (!items.length) items.push({ type: 'listItem', content: [{ type: 'paragraph' }] })
          items[items.length - 1].content!.push(list)
        }
        stack.push({ list, level: stack.length, numId, ordered })
      }
      stack[level].list.content!.push({ type: 'listItem', content: [para] })
    }

    const visit = (el: Element) => {
      switch (el.localName) {
        case 'commentRangeStart':
          this.activeComments.add(wAttr(el, 'id') ?? '')
          break
        case 'commentRangeEnd':
          this.activeComments.delete(wAttr(el, 'id') ?? '')
          break
        case 'p': {
          const { node, list } = this.paragraph(el)
          if (list) addListItem(node, list.level, list.ordered, list.numId)
          else {
            stack = []
            out.push(node)
          }
          break
        }
        case 'tbl':
          stack = []
          out.push(this.table(el))
          break
        case 'sdt': {
          const content = kid(el, 'sdtContent')
          if (content) for (const child of Array.from(content.children)) visit(child)
          break
        }
      }
    }
    for (const el of Array.from(parent.children)) visit(el)
    return out
  }

  meta(body: Element): DocMeta {
    const meta = { ...DEFAULT_META }
    const sect = kid(body, 'sectPr')
    const sz = kid(sect, 'pgSz')
    const mar = kid(sect, 'pgMar')
    const num = (v: string | null, fallback: number) => (v && Number(v) > 0 ? Number(v) : fallback)
    if (sz) {
      meta.pageWidth = num(wAttr(sz, 'w'), meta.pageWidth)
      meta.pageHeight = num(wAttr(sz, 'h'), meta.pageHeight)
    }
    if (mar) {
      meta.marginTop = num(wAttr(mar, 'top'), meta.marginTop)
      meta.marginRight = num(wAttr(mar, 'right'), meta.marginRight)
      meta.marginBottom = num(wAttr(mar, 'bottom'), meta.marginBottom)
      meta.marginLeft = num(wAttr(mar, 'left'), meta.marginLeft)
    }
    return meta
  }
}

export async function importDocx(data: ArrayBuffer | Uint8Array): Promise<ImportResult> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(data)
  } catch {
    throw new Error('This file is not a .docx document (legacy .doc files are not supported).')
  }
  const imp = new Importer(zip)
  const xml = await imp.load()
  const body = xml.getElementsByTagNameNS(W, 'body')[0]
  if (!body) throw new Error('The document has no body.')
  const content = imp.blocks(body)
  const doc: JSONContent = { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] }
  const comments = [...imp.comments.values()].filter((c) => imp.anchoredComments.has(c.id))
  const dropped = imp.comments.size - comments.length
  if (dropped) imp.warnings.push(`${dropped} comment(s) not attached to any text were skipped`)
  return { doc, meta: imp.meta(body), comments, warnings: imp.warnings }
}
