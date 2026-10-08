import JSZip from 'jszip'

/**
 * Annotated export: adds the reader's comments and highlights to the ORIGINAL .docx by editing only
 * word/document.xml (plus the comments part). Every other part of the package — styles, headers,
 * footers, footnotes, text boxes, tracked changes, custom XML — is carried over untouched, so nothing the
 * app can't model is lost.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

export interface SourceAnchor {
  /** Index of the w:p (document order) the range starts / ends in */
  startPara: number
  startOffset: number
  endPara: number
  endOffset: number
}

export interface Annotation {
  kind: 'comment' | 'highlight'
  id: string
  anchor: SourceAnchor
  /** The text the annotation covers, used to confirm the original file still matches */
  quote: string
  author?: string
  date?: string
  text?: string
  color?: string
}

export interface AnnotateResult {
  blob: Blob
  placed: Annotation[]
  skipped: { annotation: Annotation; reason: string }[]
}

// ---------- canonical paragraph text ----------
// Must stay in step with the importer: only these elements contribute characters.

const CONTAINERS = new Set(['hyperlink', 'ins', 'smartTag', 'sdt', 'sdtContent', 'fldSimple', 'customXml'])

interface Piece {
  run: Element
  /** The run child that holds the characters */
  child: Element
  start: number
  end: number
}

function isW(el: Element, local: string): boolean {
  return el.localName === local && el.namespaceURI === W
}

/** Characters of a run child: w:t is its text, tab / hyphen / break are one character each, everything else is zero-width. */
function childChars(child: Element): string {
  switch (child.localName) {
    case 't':
      return child.textContent ?? ''
    case 'tab':
      return '\t'
    case 'noBreakHyphen':
      return '‑'
    case 'br':
    case 'cr':
      return '\n'
    default:
      return ''
  }
}

function collectPieces(parent: Element, out: Piece[], pos: { n: number }) {
  for (const el of Array.from(parent.children)) {
    if (isW(el, 'r')) {
      for (const child of Array.from(el.children)) {
        const text = childChars(child)
        if (text) {
          out.push({ run: el, child, start: pos.n, end: pos.n + text.length })
          pos.n += text.length
        }
      }
    } else if (CONTAINERS.has(el.localName) && el.namespaceURI === W) {
      collectPieces(el, out, pos)
    }
  }
}

export function paragraphPieces(p: Element): Piece[] {
  const out: Piece[] = []
  collectPieces(p, out, { n: 0 })
  return out
}

export function paragraphText(p: Element): string {
  return paragraphPieces(p)
    .map((pc) => childChars(pc.child))
    .join('')
}

function norm(s: string): string {
  return s.replace(/\s+/g, ' ').trim()
}

// ---------- XML editing ----------

function w(doc: Document, local: string, attrs: Record<string, string> = {}): Element {
  const el = doc.createElementNS(W, 'w:' + local)
  for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(W, 'w:' + k, v)
  return el
}

/** Split a run in two at `charIndex` characters into the piece's child; returns the second run. */
function splitRun(doc: Document, piece: Piece, charIndex: number): Element {
  const { run, child } = piece
  const second = run.cloneNode(false) as Element
  const rPr = Array.from(run.children).find((c) => isW(c, 'rPr'))
  if (rPr) second.appendChild(rPr.cloneNode(true))

  if (isW(child, 't') && charIndex > 0 && charIndex < (child.textContent ?? '').length) {
    const text = child.textContent ?? ''
    child.textContent = text.slice(0, charIndex)
    child.setAttributeNS(XML_NS, 'xml:space', 'preserve')
    const tail = w(doc, 't')
    tail.setAttributeNS(XML_NS, 'xml:space', 'preserve')
    tail.textContent = text.slice(charIndex)
    second.appendChild(tail)
    // anything after the split text element moves to the second run
    let sib = child.nextSibling
    while (sib) {
      const next = sib.nextSibling
      second.appendChild(sib)
      sib = next
    }
  } else {
    // split between run children (before `child`, or after it when charIndex is at its end)
    const atEnd = charIndex > 0
    let sib: Node | null = atEnd ? child.nextSibling : child
    while (sib) {
      const next: Node | null = sib.nextSibling
      second.appendChild(sib)
      sib = next
    }
  }
  run.parentNode!.insertBefore(second, run.nextSibling)
  return second
}

/**
 * Ensure a run boundary exists at canonical `offset` of paragraph `p`, and return the node the caller should
 * insert BEFORE (a run element) together with its parent. Returns null when the offset is at/after the end.
 */
function boundaryAt(doc: Document, p: Element, offset: number): { parent: Node; before: Node | null } {
  const pieces = paragraphPieces(p)
  for (const pc of pieces) {
    if (offset <= pc.start) {
      // boundary lies before this piece: split its run before the child unless the child is first
      return placeBefore(doc, pc)
    }
    if (offset < pc.end) {
      const second = splitRun(doc, pc, offset - pc.start)
      return { parent: second.parentNode!, before: second }
    }
  }
  const last = pieces[pieces.length - 1]
  if (last) return { parent: last.run.parentNode!, before: last.run.nextSibling }
  return { parent: p, before: null }
}

function placeBefore(doc: Document, pc: Piece): { parent: Node; before: Node | null } {
  const { run, child } = pc
  const firstContent = Array.from(run.children).find((c) => !isW(c, 'rPr'))
  if (firstContent === child) return { parent: run.parentNode!, before: run }
  const second = splitRun(doc, pc, 0)
  return { parent: second.parentNode!, before: second }
}

function insertAt(doc: Document, p: Element, offset: number, node: Node) {
  const { parent, before } = boundaryAt(doc, p, offset)
  parent.insertBefore(node, before)
}

const HIGHLIGHT_COLORS: Record<string, [number, number, number]> = {
  yellow: [255, 255, 0],
  green: [0, 255, 0],
  cyan: [0, 255, 255],
  magenta: [255, 0, 255],
  blue: [0, 0, 255],
  red: [255, 0, 0],
  darkBlue: [0, 0, 128],
  darkCyan: [0, 128, 128],
  darkGreen: [0, 128, 0],
  darkMagenta: [128, 0, 128],
  darkRed: [128, 0, 0],
  darkYellow: [128, 128, 0],
  darkGray: [128, 128, 128],
  lightGray: [192, 192, 192],
}

/** Word only knows 16 highlight colours; pick the nearest one. */
export function nearestHighlight(hex: string | undefined): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex ?? '')
  if (!m) return 'yellow'
  const n = parseInt(m[1], 16)
  const rgb = [(n >> 16) & 255, (n >> 8) & 255, n & 255]
  let best = 'yellow'
  let bestD = Infinity
  for (const [name, c] of Object.entries(HIGHLIGHT_COLORS)) {
    const d = (c[0] - rgb[0]) ** 2 + (c[1] - rgb[1]) ** 2 + (c[2] - rgb[2]) ** 2
    if (d < bestD) {
      bestD = d
      best = name
    }
  }
  return best
}

const RPR_AFTER_HIGHLIGHT = new Set(['u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'])

function setRunHighlight(doc: Document, run: Element, color: string) {
  let rPr = Array.from(run.children).find((c) => isW(c, 'rPr'))
  if (!rPr) {
    rPr = w(doc, 'rPr')
    run.insertBefore(rPr, run.firstChild)
  }
  const existing = Array.from(rPr.children).find((c) => isW(c, 'highlight'))
  if (existing) {
    existing.setAttributeNS(W, 'w:val', color)
    return
  }
  const hl = w(doc, 'highlight', { val: color })
  const successor = Array.from(rPr.children).find((c) => RPR_AFTER_HIGHLIGHT.has(c.localName))
  rPr.insertBefore(hl, successor ?? null)
}

/** Highlight canonical [start, end) of a paragraph, splitting runs at both ends. */
function highlightRange(doc: Document, p: Element, start: number, end: number, color: string) {
  if (end <= start) return
  // Drop temporary markers at both boundaries (which splits runs there), then colour the runs between them.
  const endMark = doc.createComment('dw-end')
  const startMark = doc.createComment('dw-start')
  insertAt(doc, p, end, endMark)
  insertAt(doc, p, start, startMark)
  let inside = false
  const walk = (parent: Node) => {
    for (const n of Array.from(parent.childNodes)) {
      if (n === startMark) inside = true
      else if (n === endMark) inside = false
      else if (n.nodeType === 1) {
        const el = n as Element
        if (isW(el, 'r')) {
          if (inside) setRunHighlight(doc, el, color)
        } else if (CONTAINERS.has(el.localName) && el.namespaceURI === W) walk(el)
      }
    }
  }
  walk(p)
  startMark.parentNode?.removeChild(startMark)
  endMark.parentNode?.removeChild(endMark)
}

// ---------- package ----------

function serialize(doc: Document): string {
  const xml = new XMLSerializer().serializeToString(doc)
  return xml.startsWith('<?xml') ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml
}

function parse(text: string): Document {
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length) throw new Error('The document XML could not be read.')
  return doc
}

function initials(name: string): string {
  return (
    name
      .split(/\s+/)
      .map((p) => p[0] ?? '')
      .join('')
      .slice(0, 3)
      .toUpperCase() || 'A'
  )
}

export async function annotateDocx(original: ArrayBuffer | Uint8Array, annotations: Annotation[]): Promise<AnnotateResult> {
  const zip = await JSZip.loadAsync(original)
  const docFile = zip.file('word/document.xml')
  if (!docFile) throw new Error('The original file has no word/document.xml.')
  const doc = parse(await docFile.async('string'))
  const paras = Array.from(doc.getElementsByTagNameNS(W, 'p'))

  // existing comments part (so new ids don't collide and imported comments stay as they are)
  const commentsFile = zip.file('word/comments.xml')
  const commentsDoc = commentsFile
    ? parse(await commentsFile.async('string'))
    : parse(`<w:comments xmlns:w="${W}"/>`)
  let nextId =
    Math.max(-1, ...Array.from(commentsDoc.getElementsByTagNameNS(W, 'comment')).map((c) => Number(c.getAttributeNS(W, 'id')) || 0)) + 1
  // ids are also used by w:ins/w:del; Word only needs comment ids unique among comments.

  const placed: Annotation[] = []
  const skipped: AnnotateResult['skipped'] = []

  const textOf = (a: SourceAnchor): string | null => {
    const first = paras[a.startPara]
    const last = paras[a.endPara]
    if (!first || !last || a.endPara < a.startPara) return null
    const parts: string[] = []
    for (let i = a.startPara; i <= a.endPara; i++) {
      const t = paragraphText(paras[i])
      parts.push(t.slice(i === a.startPara ? a.startOffset : 0, i === a.endPara ? a.endOffset : undefined))
    }
    return parts.join('\n')
  }

  // Apply in reverse document order so earlier edits can't disturb later anchors' paragraph indexes.
  const ordered = [...annotations].sort(
    (a, b) => b.anchor.endPara - a.anchor.endPara || b.anchor.endOffset - a.anchor.endOffset,
  )

  // Number new comments in reading order, even though they are applied back to front.
  const newIds = new Map<Annotation, string>()
  for (const ann of [...ordered].reverse()) if (ann.kind === 'comment') newIds.set(ann, String(nextId++))
  const newComments: { id: number; el: Element }[] = []

  for (const ann of ordered) {
    const a = ann.anchor
    const found = textOf(a)
    if (found == null) {
      skipped.push({ annotation: ann, reason: 'The passage is not in the original file (it was added or moved).' })
      continue
    }
    if (norm(found) !== norm(ann.quote)) {
      skipped.push({ annotation: ann, reason: 'The text in the original file differs from what you annotated (it was edited).' })
      continue
    }

    if (ann.kind === 'comment') {
      const id = newIds.get(ann)!
      const author = ann.author || 'Reader'
      const date = ann.date ? new Date(ann.date) : new Date()

      const comment = w(commentsDoc, 'comment', {
        id,
        author,
        date: (Number.isNaN(date.getTime()) ? new Date() : date).toISOString().replace(/\.\d+Z$/, 'Z'),
        initials: initials(author),
      })
      for (const line of (ann.text ?? '').split('\n')) {
        const p = w(commentsDoc, 'p')
        const r = w(commentsDoc, 'r')
        const t = w(commentsDoc, 't')
        t.setAttributeNS(XML_NS, 'xml:space', 'preserve')
        t.textContent = line
        r.appendChild(t)
        p.appendChild(r)
        comment.appendChild(p)
      }
      newComments.push({ id: Number(id), el: comment })

      // end first, then start, so offsets in the same paragraph stay valid
      const endPara = paras[a.endPara]
      const ref = w(doc, 'r')
      ref.appendChild(w(doc, 'commentReference', { id }))
      const { parent, before } = boundaryAt(doc, endPara, a.endOffset)
      parent.insertBefore(w(doc, 'commentRangeEnd', { id }), before)
      parent.insertBefore(ref, before)
      insertAt(doc, paras[a.startPara], a.startOffset, w(doc, 'commentRangeStart', { id }))
    } else {
      const color = nearestHighlight(ann.color)
      for (let i = a.startPara; i <= a.endPara; i++) {
        const len = paragraphText(paras[i]).length
        highlightRange(doc, paras[i], i === a.startPara ? a.startOffset : 0, i === a.endPara ? a.endOffset : len, color)
      }
    }
    placed.push(ann)
  }

  for (const c of newComments.sort((x, y) => x.id - y.id)) commentsDoc.documentElement.appendChild(c.el)
  zip.file('word/document.xml', serialize(doc))

  if (commentsDoc.getElementsByTagNameNS(W, 'comment').length) {
    zip.file('word/comments.xml', serialize(commentsDoc))
    if (!commentsFile) await registerCommentsPart(zip)
  }

  const blob = await zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' })
  return { blob, placed: placed.reverse(), skipped }
}

async function registerCommentsPart(zip: JSZip) {
  const ctFile = zip.file('[Content_Types].xml')
  if (ctFile) {
    const ct = parse(await ctFile.async('string'))
    const has = Array.from(ct.documentElement.children).some((c) => c.getAttribute('PartName') === '/word/comments.xml')
    if (!has) {
      const o = ct.createElementNS(ct.documentElement.namespaceURI, 'Override')
      o.setAttribute('PartName', '/word/comments.xml')
      o.setAttribute('ContentType', COMMENTS_CT)
      ct.documentElement.appendChild(o)
      zip.file('[Content_Types].xml', serialize(ct))
    }
  }
  const relPath = 'word/_rels/document.xml.rels'
  const relFile = zip.file(relPath)
  const rels = relFile ? parse(await relFile.async('string')) : parse(`<Relationships xmlns="${REL_NS}"/>`)
  const ids = Array.from(rels.documentElement.children).map((r) => r.getAttribute('Id') ?? '')
  let n = ids.length + 1
  while (ids.includes('rId' + n)) n++
  const rel = rels.createElementNS(rels.documentElement.namespaceURI, 'Relationship')
  rel.setAttribute('Id', 'rId' + n)
  rel.setAttribute('Type', COMMENTS_REL)
  rel.setAttribute('Target', 'comments.xml')
  rels.documentElement.appendChild(rel)
  zip.file(relPath, serialize(rels))
}
