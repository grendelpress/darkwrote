import type { Editor } from '@tiptap/core'

export interface Range {
  from: number
  to: number
}

const COLORS = ['#ffeb3b', '#69f0ae', '#40c4ff', '#ff80ab']

/** Selection range read from the browser selection (the editor has no caret of its own while read-only). */
export function domSelectionRange(editor: Editor): Range | null {
  const sel = window.getSelection()
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed || !sel.anchorNode || !sel.focusNode) return null
  const root = editor.view.dom
  if (!root.contains(sel.anchorNode) || !root.contains(sel.focusNode)) return null
  try {
    const a = editor.view.posAtDOM(sel.anchorNode, sel.anchorOffset)
    const b = editor.view.posAtDOM(sel.focusNode, sel.focusOffset)
    const [from, to] = a <= b ? [a, b] : [b, a]
    return from < to ? { from, to } : null
  } catch {
    return null
  }
}

export interface SelectionActions {
  highlight: (r: Range, color: string) => void
  clearHighlight: (r: Range) => void
  comment: (r: Range) => void
}

/**
 * Bottom action bar for read-only mode. It sits away from the text so it never collides with Android's own
 * selection handles / context menu, and rides above the on-screen keyboard through the --kb CSS variable.
 */
export class SelectionBar {
  private last: Range | null = null
  private timer = 0
  private pressing = false

  constructor(
    private host: HTMLElement,
    private editor: Editor,
    private active: () => boolean,
    private actions: SelectionActions,
  ) {
    const mk = (label: string, title: string, cls: string, fn: (r: Range) => void, style = '') => {
      const b = document.createElement('button')
      b.type = 'button'
      b.className = cls
      b.textContent = label
      b.title = title
      b.setAttribute('aria-label', title)
      if (style) b.style.cssText = style
      // keep the text selection alive while the button is pressed
      b.addEventListener('pointerdown', (e) => {
        e.preventDefault()
        this.pressing = true
        window.setTimeout(() => (this.pressing = false), 600)
      })
      b.addEventListener('click', () => {
        const r = domSelectionRange(this.editor) ?? this.last
        if (!r) return
        fn(r)
        this.finish()
      })
      return b
    }
    for (const c of COLORS) host.append(mk('', `Highlight`, 'swatch', (r) => this.actions.highlight(r, c), `background:${c}`))
    host.append(mk('⌫', 'Remove highlight', 'swatch', (r) => this.actions.clearHighlight(r)))
    const sep = document.createElement('span')
    sep.className = 'sep'
    host.append(sep)
    host.append(mk('💬 Comment', 'Comment on the selection', 'primary', (r) => this.actions.comment(r)))
    host.append(
      mk('Copy', 'Copy', '', (r) => {
        const text = this.editor.state.doc.textBetween(r.from, r.to, '\n')
        void navigator.clipboard?.writeText(text).catch(() => undefined)
      }),
    )
    document.addEventListener('selectionchange', () => {
      window.clearTimeout(this.timer)
      this.timer = window.setTimeout(() => this.update(), 120)
    })
  }

  private update() {
    const r = this.active() ? domSelectionRange(this.editor) : null
    if (!r && this.pressing) return
    if (r) this.last = r
    this.host.hidden = !r
  }

  private finish() {
    window.getSelection()?.removeAllRanges()
    this.last = null
    this.host.hidden = true
  }

  hide() {
    this.host.hidden = true
  }
}
