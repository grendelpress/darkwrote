import type { Node as PMNode } from '@tiptap/pm/model'
import type { Annotation, SourceAnchor } from './docx/annotate'
import type { CommentData } from './docx/import'

export interface AnnotationRecord {
  kind: 'comment' | 'highlight'
  id: string
  quote: string
  text?: string
  author?: string
  date?: string
  color?: string
  /** Came from the original .docx (so annotated export leaves it alone) */
  imported?: boolean
  /** Where the passage is in the source .docx, or null if it can't be traced (new/edited paragraphs) */
  anchor: SourceAnchor | null
  from: number
  to: number
  heading: string | null
  before: string
  after: string
}

function canonicalLength(node: PMNode): number {
  if (node.isText) return node.text!.length
  return node.type.name === 'hardBreak' ? 1 : 0
}

/** Position in the editor -> (source paragraph index, canonical character offset), if traceable. */
export function sourceLocation(doc: PMNode, pos: number): { para: number; offset: number } | null {
  const $pos = doc.resolve(pos)
  const parent = $pos.parent
  const para = parent.attrs?.srcPara
  if (!parent.isTextblock || typeof para !== 'number') return null
  const target = $pos.parentOffset
  let offset = 0
  parent.forEach((child, childOffset) => {
    if (childOffset >= target) return
    if (childOffset + child.nodeSize <= target) offset += canonicalLength(child)
    else if (child.isText) offset += target - childOffset
  })
  return { para, offset }
}

interface Range {
  from: number
  to: number
}

/** Gather every comment and reader highlight in the document, with source anchors where possible. */
export function collectAnnotations(doc: PMNode, comments: CommentData[]): AnnotationRecord[] {
  const byId = new Map(comments.map((c) => [c.id, c]))
  const commentRanges = new Map<string, Range>()
  const highlights: (Range & { color: string })[] = []

  const headings: { pos: number; text: string }[] = []
  doc.descendants((node, pos, parent) => {
    if (node.type.name === 'heading' && node.textContent.trim()) headings.push({ pos, text: node.textContent.trim() })
    if (!node.isText) return
    const end = pos + node.nodeSize
    for (const m of node.marks) {
      if (m.type.name === 'comment') {
        const id = String(m.attrs.commentId)
        const cur = commentRanges.get(id)
        if (cur) cur.to = end
        else commentRanges.set(id, { from: pos, to: end })
      } else if (m.type.name === 'highlight' && m.attrs.user) {
        // merge neighbours of the same colour inside one textblock
        const last = highlights[highlights.length - 1]
        if (last && last.to === pos && last.color === m.attrs.color && parent?.isTextblock) last.to = end
        else highlights.push({ from: pos, to: end, color: m.attrs.color ?? '#ffff00' })
      }
    }
  })

  const headingFor = (pos: number) => {
    let found: string | null = null
    for (const h of headings) {
      if (h.pos < pos) found = h.text
      else break
    }
    return found
  }
  const common = (from: number, to: number) => {
    const start = sourceLocation(doc, from)
    const end = sourceLocation(doc, to)
    return {
      from,
      to,
      quote: doc.textBetween(from, to, ' ').replace(/\s+/g, ' ').trim(),
      anchor:
        start && end
          ? ({ startPara: start.para, startOffset: start.offset, endPara: end.para, endOffset: end.offset } as SourceAnchor)
          : null,
      heading: headingFor(from),
      before: doc.textBetween(Math.max(0, from - 80), from, ' ').replace(/\s+/g, ' ').trim(),
      after: doc.textBetween(to, Math.min(doc.content.size, to + 80), ' ').replace(/\s+/g, ' ').trim(),
    }
  }

  const out: AnnotationRecord[] = []
  for (const [id, r] of commentRanges) {
    const c = byId.get(id)
    if (!c) continue
    out.push({ kind: 'comment', id, text: c.text, author: c.author, date: c.date, imported: c.imported, ...common(r.from, r.to) })
  }
  highlights.forEach((h, i) => out.push({ kind: 'highlight', id: `h${i + 1}`, color: h.color, ...common(h.from, h.to) }))
  return out.sort((a, b) => a.from - b.from)
}

/** Records that annotated export can place into the original file (reader's own additions only). */
export function placeable(records: AnnotationRecord[]): Annotation[] {
  return records
    .filter((r) => !r.imported && r.anchor)
    .map((r) => ({
      kind: r.kind,
      id: r.id,
      anchor: r.anchor!,
      quote: r.quote,
      author: r.author,
      date: r.date,
      text: r.text,
      color: r.color,
    }))
}

export interface DocumentInfo {
  name: string
  sha256: string | null
  size: number | null
}

export function annotationsToJson(records: AnnotationRecord[], info: DocumentInfo): string {
  return JSON.stringify(
    {
      format: 'darkwrote-annotations',
      version: 1,
      exportedAt: new Date().toISOString(),
      document: info,
      annotations: records.map(({ from: _from, to: _to, ...r }) => r),
    },
    null,
    2,
  )
}

export function annotationsToMarkdown(records: AnnotationRecord[], info: DocumentInfo): string {
  const comments = records.filter((r) => r.kind === 'comment').length
  const lines = [
    `# Annotations — ${info.name}`,
    '',
    `_Exported ${new Date().toISOString().slice(0, 10)} · ${comments} comment${comments === 1 ? '' : 's'} · ${records.length - comments} highlight${records.length - comments === 1 ? '' : 's'}_`,
    '',
  ]
  let currentHeading: string | null | undefined
  for (const r of records) {
    if (r.heading !== currentHeading) {
      currentHeading = r.heading
      lines.push(`## ${r.heading ?? 'Before the first heading'}`, '')
    }
    const quote = r.quote.replace(/\n/g, ' ')
    lines.push(`> ${r.kind === 'highlight' ? '🖍 ' : ''}${quote}`, '')
    if (r.kind === 'comment') {
      const when = r.date ? ` (${r.date.slice(0, 10)})` : ''
      const body = (r.text ?? '').split('\n').join('\n  ')
      lines.push(`- **${r.author ?? 'Reader'}**${when}: ${body}`, '')
    }
  }
  if (!records.length) lines.push('_No annotations yet._', '')
  return lines.join('\n')
}
