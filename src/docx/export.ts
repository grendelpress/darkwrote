import {
  AlignmentType,
  BorderStyle,
  Document,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  Paragraph,
  ShadingType,
  Tab,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  type ParagraphChild,
} from 'docx'
import type { JSONContent } from '@tiptap/core'
import { DEFAULT_META, type CommentData, type DocMeta } from './import'

type Mark = NonNullable<JSONContent['marks']>[number]
type Block = Paragraph | Table

const HEADINGS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
]

const ALIGN: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
}

const BULLETS = ['•', '◦', '▪']
const BULLET_REF = 'dw-bullet'
const ORDERED_REF = 'dw-ordered'
const MAX_LEVELS = 9

function hex(color: unknown): string | undefined {
  if (typeof color !== 'string') return undefined
  const m = /^#([0-9a-f]{6})$/i.exec(color.trim())
  if (m) return m[1].toUpperCase()
  const short = /^#([0-9a-f]{3})$/i.exec(color.trim())
  if (short) return short[1].split('').map((c) => c + c).join('').toUpperCase()
  const rgb = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/i.exec(color.trim())
  if (rgb) return [rgb[1], rgb[2], rgb[3]].map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase()
  return undefined
}

/** CSS font-size ("12pt", "16px") to half-points. */
function halfPoints(size: unknown): number | undefined {
  if (typeof size !== 'string') return undefined
  const m = /^([\d.]+)(pt|px)?$/.exec(size.trim())
  if (!m) return undefined
  const pt = m[2] === 'px' ? Number(m[1]) * 0.75 : Number(m[1])
  return pt > 0 ? Math.round(pt * 2) : undefined
}

function dataUrlToImage(src: string): { type: 'png' | 'jpg' | 'gif' | 'bmp'; data: Uint8Array } | null {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.*)$/i.exec(src)
  if (!m) return null
  const kind = m[1].toLowerCase()
  const type = kind === 'jpeg' || kind === 'jpg' ? 'jpg' : (kind as 'png' | 'gif' | 'bmp')
  const bin = atob(m[2])
  const data = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i)
  return { type, data }
}

class Exporter {
  /** Word comment ids are numeric; map from our string ids, only for comments that are anchored in the text. */
  commentIndex = new Map<string, number>()
  private remaining = new Map<string, number>()
  private started = new Set<string>()

  constructor(doc: JSONContent, known: Set<string>) {
    const count = (n: JSONContent) => {
      for (const m of n.marks ?? []) {
        const id = m.type === 'comment' ? String(m.attrs?.commentId) : null
        if (id && known.has(id)) {
          if (!this.commentIndex.has(id)) this.commentIndex.set(id, this.commentIndex.size)
          this.remaining.set(id, (this.remaining.get(id) ?? 0) + 1)
        }
      }
      n.content?.forEach(count)
    }
    count(doc)
  }

  inline(nodes: JSONContent[] | undefined, extra: { bold?: boolean; font?: string; size?: number } = {}): ParagraphChild[] {
    const out: ParagraphChild[] = []
    for (const n of nodes ?? []) {
      if (n.type === 'text') {
        out.push(...this.text(n, extra))
      } else if (n.type === 'hardBreak') {
        out.push(new TextRun({ break: 1 }))
      } else if (n.type === 'image') {
        const img = this.image(n)
        if (img) out.push(img)
      }
    }
    return out
  }

  text(n: JSONContent, extra: { bold?: boolean; font?: string; size?: number }): ParagraphChild[] {
    const marks: Mark[] = n.marks ?? []
    const has = (t: string) => marks.some((m) => m.type === t)
    const style = marks.find((m) => m.type === 'textStyle')?.attrs ?? {}
    const hl = marks.find((m) => m.type === 'highlight')?.attrs?.color
    const link = marks.find((m) => m.type === 'link')?.attrs?.href as string | undefined
    const hlHex = hex(hl)

    // TextRun can't contain raw tabs/newlines, so split them out.
    const parts = (n.text ?? '').split(/(\t)/)
    const runs = parts
      .filter((p) => p !== '')
      .map((part) => {
        const isTab = part === '\t'
        return new TextRun({
          text: isTab ? undefined : part,
          children: isTab ? [new Tab()] : undefined,
          bold: has('bold') || extra.bold || undefined,
          italics: has('italic') || undefined,
          underline: has('underline') || link ? {} : undefined,
          strike: has('strike') || undefined,
          superScript: has('superscript') || undefined,
          subScript: has('subscript') || undefined,
          color: hex(style.color) ?? (link ? '0563C1' : undefined),
          font: (style.fontFamily as string | undefined)?.split(',')[0].replace(/["']/g, '').trim() || extra.font,
          size: halfPoints(style.fontSize) ?? extra.size,
          shading: hlHex ? { type: ShadingType.CLEAR, fill: hlHex, color: 'auto' } : undefined,
        })
      })
    const body: ParagraphChild[] =
      link && !link.startsWith('#') ? [new ExternalHyperlink({ link, children: runs })] : runs

    // A comment range opens at its first text node and closes after its last, even across paragraphs.
    const ids = marks
      .filter((m) => m.type === 'comment')
      .map((m) => String(m.attrs?.commentId))
      .filter((id) => this.commentIndex.has(id))
    const before: ParagraphChild[] = []
    const after: ParagraphChild[] = []
    for (const id of ids) {
      const idx = this.commentIndex.get(id)!
      if (!this.started.has(id)) {
        this.started.add(id)
        before.push(new CommentRangeStart(idx))
      }
      const left = (this.remaining.get(id) ?? 1) - 1
      this.remaining.set(id, left)
      if (left === 0) after.push(new CommentRangeEnd(idx), new TextRun({ children: [new CommentReference(idx)] }))
    }
    return [...before, ...body, ...after]
  }

  image(n: JSONContent): ImageRun | null {
    const parsed = dataUrlToImage(String(n.attrs?.src ?? ''))
    if (!parsed) return null
    const width = Number(n.attrs?.width) || 320
    const height = Number(n.attrs?.height) || 240
    return new ImageRun({
      type: parsed.type,
      data: parsed.data,
      transformation: { width, height },
      altText: { name: 'image', description: String(n.attrs?.alt ?? ''), title: String(n.attrs?.title ?? '') },
    })
  }

  paragraph(n: JSONContent, opts: { listRef?: string; level?: number; indentLeft?: number } = {}): Paragraph {
    const attrs = n.attrs ?? {}
    const isHeading = n.type === 'heading'
    return new Paragraph({
      heading: isHeading ? HEADINGS[Math.min(5, Math.max(0, Number(attrs.level) - 1))] : undefined,
      alignment: attrs.textAlign ? ALIGN[attrs.textAlign as string] : undefined,
      numbering: opts.listRef ? { reference: opts.listRef, level: opts.level ?? 0 } : undefined,
      indent: opts.indentLeft ? { left: opts.indentLeft } : undefined,
      children: this.inline(n.content),
    })
  }

  list(n: JSONContent, level: number, out: Block[]) {
    const ref = n.type === 'orderedList' ? ORDERED_REF : BULLET_REF
    for (const item of n.content ?? []) {
      let first = true
      for (const child of item.content ?? []) {
        if (child.type === 'bulletList' || child.type === 'orderedList') {
          this.list(child, Math.min(level + 1, MAX_LEVELS - 1), out)
        } else if (child.type === 'paragraph' || child.type === 'heading') {
          // Only the first paragraph carries the bullet/number; later ones are indented continuations.
          out.push(
            first
              ? this.paragraph(child, { listRef: ref, level })
              : this.paragraph(child, { indentLeft: 720 * (level + 1) }),
          )
          first = false
        } else {
          out.push(...this.block(child))
        }
      }
    }
  }

  table(n: JSONContent): Table {
    const rows = (n.content ?? []).map(
      (row) =>
        new TableRow({
          children: (row.content ?? []).map((cell) => {
            const a = cell.attrs ?? {}
            const fill = hex(a.backgroundColor)
            const children = (cell.content ?? []).flatMap((c) => this.block(c))
            return new TableCell({
              columnSpan: Number(a.colspan) > 1 ? Number(a.colspan) : undefined,
              rowSpan: Number(a.rowspan) > 1 ? Number(a.rowspan) : undefined,
              shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
              children: children.length ? children : [new Paragraph({})],
            })
          }),
        }),
    )
    return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } })
  }

  block(n: JSONContent): Block[] {
    switch (n.type) {
      case 'paragraph':
      case 'heading':
        return [this.paragraph(n)]
      case 'bulletList':
      case 'orderedList': {
        const out: Block[] = []
        this.list(n, 0, out)
        return out
      }
      case 'blockquote':
        return (n.content ?? []).flatMap((c) =>
          c.type === 'paragraph' ? [this.paragraph(c, { indentLeft: 720 })] : this.block(c),
        )
      case 'codeBlock': {
        const lines = (n.content?.map((c) => c.text ?? '').join('') ?? '').split('\n')
        return lines.map(
          (line) =>
            new Paragraph({
              children: [new TextRun({ text: line, font: 'Courier New', size: 20 })],
              shading: { type: ShadingType.CLEAR, fill: 'F2F2F2', color: 'auto' },
            }),
        )
      }
      case 'horizontalRule':
        return [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '808080', space: 1 } },
          }),
        ]
      case 'table':
        return [this.table(n)]
      default:
        return []
    }
  }
}

function numberingConfig() {
  const levels = (kind: 'bullet' | 'ordered') =>
    Array.from({ length: MAX_LEVELS }, (_, level) => ({
      level,
      format: kind === 'bullet' ? LevelFormat.BULLET : [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN][level % 3],
      text: kind === 'bullet' ? BULLETS[level % BULLETS.length] : `%${level + 1}.`,
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
    }))
  return [
    { reference: BULLET_REF, levels: levels('bullet') },
    { reference: ORDERED_REF, levels: levels('ordered') },
  ]
}

export async function exportDocx(
  doc: JSONContent,
  meta: DocMeta = DEFAULT_META,
  comments: CommentData[] = [],
): Promise<Blob> {
  const byId = new Map(comments.map((c) => [c.id, c]))
  const ex = new Exporter(doc, new Set(byId.keys()))
  const children = (doc.content ?? []).flatMap((n) => ex.block(n))
  const document = new Document({
    creator: 'Darkwrote',
    numbering: { config: numberingConfig() },
    comments: {
      children: [...ex.commentIndex].map(([id, index]) => {
        const c = byId.get(id)!
        const date = new Date(c.date)
        return {
          id: index,
          author: c.author,
          initials: c.author
            .split(/\s+/)
            .map((w) => w[0] ?? '')
            .join('')
            .slice(0, 3)
            .toUpperCase(),
          date: Number.isNaN(date.getTime()) ? new Date() : date,
          children: c.text.split('\n').map((line) => new Paragraph({ children: [new TextRun(line)] })),
        }
      }),
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: meta.pageWidth, height: meta.pageHeight },
            margin: { top: meta.marginTop, right: meta.marginRight, bottom: meta.marginBottom, left: meta.marginLeft },
          },
        },
        children: children.length ? children : [new Paragraph({})],
      },
    ],
  })
  return Packer.toBlob(document)
}
