import type { AnnotateResult } from './docx/annotate'
import { annotationsToJson, annotationsToMarkdown, placeable, type AnnotationRecord, type DocumentInfo } from './annotations'
import { DOCX_MIME, baseName, canShareFile, downloadBlob, shareBlob } from './files'

export interface ExportContext {
  name: string
  hasOriginal: boolean
  /** The text was edited, so the original file can no longer be patched */
  contentEdited: boolean
  records: AnnotationRecord[]
  info: DocumentInfo
  buildAnnotated: () => Promise<AnnotateResult>
  buildRebuilt: () => Promise<Blob>
  onExported: () => void
  notify: (msg: string, error?: boolean) => void
}

function h<K extends keyof HTMLElementTagNameMap>(tag: K, props: Record<string, any> = {}, ...kids: (Node | string)[]): HTMLElementTagNameMap[K] {
  const n = document.createElement(tag)
  Object.assign(n, props)
  n.append(...kids)
  return n
}

interface Option {
  title: string
  detail: string
  disabledReason?: string
  filename: string
  produce: () => Promise<{ blob: Blob; warning?: string }>
}

export function openExportDialog(dialog: HTMLDialogElement, ctx: ExportContext) {
  const base = baseName(ctx.name)
  const comments = ctx.records.filter((r) => r.kind === 'comment' && !r.imported).length
  const highlights = ctx.records.filter((r) => r.kind === 'highlight').length
  const mine = comments + highlights
  const canPatch = ctx.hasOriginal && !ctx.contentEdited

  const options: Option[] = [
    {
      title: 'Original file + your annotations (.docx)',
      detail: `Adds your ${comments} comment${comments === 1 ? '' : 's'} and ${highlights} highlight${highlights === 1 ? '' : 's'} to a copy of the original. Headers, footnotes, styles, text boxes and everything else stay exactly as they were.`,
      disabledReason: !ctx.hasOriginal
        ? 'Only available for documents opened from a .docx file.'
        : ctx.contentEdited
          ? 'You edited the text, so the original can’t be patched. Export your annotations or the edited copy instead.'
          : mine === 0
            ? 'You haven’t added any comments or highlights yet.'
            : undefined,
      filename: `${base}.annotated.docx`,
      produce: async () => {
        const r = await ctx.buildAnnotated()
        const warning = r.skipped.length
          ? `${r.skipped.length} annotation${r.skipped.length === 1 ? ' was' : 's were'} left out because the passage couldn’t be matched in the original file. They’re still in the annotation exports.`
          : undefined
        return { blob: r.blob, warning }
      },
    },
    {
      title: 'Annotations only — Markdown (.md)',
      detail: 'A readable list of your comments and highlights, grouped by heading, with the passage each refers to.',
      filename: `${base}.annotations.md`,
      produce: async () => ({ blob: new Blob([annotationsToMarkdown(ctx.records, ctx.info)], { type: 'text/markdown' }) }),
    },
    {
      title: 'Annotations only — JSON (.json)',
      detail: 'Machine-readable, with exact locations in the source document, for other tools or later re-use.',
      filename: `${base}.annotations.json`,
      produce: async () => ({ blob: new Blob([annotationsToJson(ctx.records, ctx.info)], { type: 'application/json' }) }),
    },
    {
      title: 'Edited document, rebuilt (.docx)',
      detail: ctx.hasOriginal
        ? 'Rebuilds the whole document from what Darkwrote shows. Use this only if you edited the text; headers, footnotes, text boxes and other features Darkwrote can’t display are not carried over.'
        : 'Writes your document as a new .docx.',
      filename: ctx.hasOriginal ? `${base} (edited).docx` : `${base}.docx`,
      produce: async () => ({ blob: await ctx.buildRebuilt() }),
    },
  ]

  const warn = h('div', { className: 'export-warn', hidden: true })
  const run = async (opt: Option, how: 'download' | 'share', buttons: HTMLButtonElement[]) => {
    buttons.forEach((b) => (b.disabled = true))
    try {
      const { blob, warning } = await opt.produce()
      warn.hidden = !warning
      warn.textContent = warning ?? ''
      if (how === 'share') {
        if (await shareBlob(blob, opt.filename)) ctx.onExported()
      } else {
        downloadBlob(blob, opt.filename)
        ctx.notify(`Saved ${opt.filename} to your Downloads`)
        ctx.onExported()
      }
    } catch (e) {
      ctx.notify(`Export failed: ${(e as Error).message}`, true)
    } finally {
      buttons.forEach((b) => (b.disabled = false))
    }
  }

  const rows = options.map((opt) => {
    const disabled = !!opt.disabledReason
    const dl = h('button', { className: 'btn primary', type: 'button', textContent: 'Download', disabled })
    const sh = h('button', { className: 'btn', type: 'button', textContent: 'Share…', hidden: !('share' in navigator && 'canShare' in navigator), disabled })
    const all = [dl, sh]
    dl.addEventListener('click', () => void run(opt, 'download', all))
    sh.addEventListener('click', () => void run(opt, 'share', all))
    const info = h('div', { className: 'info' }, h('strong', {}, opt.title), h('small', {}, opt.disabledReason ?? opt.detail))
    return h('div', { className: 'export-opt' + (disabled ? ' disabled' : '') }, info, h('div', { className: 'btns' }, dl, sh))
  })

  // Hide Share when the platform can't share a file of this kind (checked lazily with a tiny sample).
  rows.forEach((row, i) => {
    const sh = row.querySelector<HTMLButtonElement>('.btns button:nth-child(2)')!
    if (!sh.hidden && !canShareFile(new Blob(['x'], { type: i === 0 || i === 3 ? DOCX_MIME : options[i].filename.endsWith('.md') ? 'text/markdown' : 'application/json' }), options[i].filename)) sh.hidden = true
  })

  const close = h('button', { className: 'btn', type: 'button', textContent: 'Close' })
  close.addEventListener('click', () => dialog.close())
  dialog.replaceChildren(
    h('h2', {}, 'Export'),
    h('p', { className: 'export-note' }, canPatch || !ctx.hasOriginal ? `${mine} of your annotations are included in the options below.` : 'Your edits change which options are available.'),
    ...rows,
    warn,
    h('div', { className: 'dialog-actions' }, h('span', { className: 'spacer' }), close),
  )
  if (!dialog.open) dialog.showModal()
}

export { placeable }
