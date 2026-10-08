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

interface Pinned {
  pm: Range
  dom: globalThis.Range
}

const HIGHLIGHT_NAME = 'dw-pending'

/**
 * Action bar for read-only mode.
 *
 * Android (and Samsung's One UI in particular) shows its own floating Copy/Share/Translate menu on a long
 * press, and a web page cannot turn that menu off. So this bar is built to coexist with it:
 *  - it is always docked at the top of the screen, under the header;
 *  - the selected passage is "pinned": it keeps its own highlight and the bar stays up even after the
 *    system selection is dismissed, so you can tap away from Samsung's menu without losing your place.
 * It also rides above the on-screen keyboard through the --kb CSS variable.
 */
export class SelectionBar {
  private pinned: Pinned | null = null
  private timer = 0

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
      })
      b.addEventListener('click', () => {
        const r = domSelectionRange(this.editor) ?? this.pinned?.pm
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
    const close = document.createElement('button')
    close.type = 'button'
    close.className = 'close'
    close.textContent = '✕'
    close.title = 'Dismiss'
    close.setAttribute('aria-label', 'Dismiss')
    close.addEventListener('pointerdown', (e) => e.preventDefault())
    close.addEventListener('click', () => this.finish())
    host.append(close)
    document.addEventListener('selectionchange', () => {
      window.clearTimeout(this.timer)
      this.timer = window.setTimeout(() => this.update(), 120)
    })
  }

  private update() {
    if (!this.active()) return this.hide()
    const r = domSelectionRange(this.editor)
    if (r) {
      const sel = window.getSelection()!
      this.pinned = { pm: r, dom: sel.getRangeAt(0).cloneRange() }
      this.paintPinned()
      this.host.hidden = false
    }
    // No selection any more (e.g. the system menu was dismissed): keep the pinned passage and the bar.
  }

  /** Our own highlight for the pinned passage, since the system selection may be gone. */
  private paintPinned() {
    const api = (window as unknown as { CSS?: { highlights?: Map<string, unknown> }; Highlight?: new (...r: globalThis.Range[]) => unknown }).CSS?.highlights
    const Hl = (window as unknown as { Highlight?: new (...r: globalThis.Range[]) => unknown }).Highlight
    if (api && Hl) {
      if (this.pinned) api.set(HIGHLIGHT_NAME, new Hl(this.pinned.dom))
      else api.delete(HIGHLIGHT_NAME)
    }
  }

  private finish() {
    window.getSelection()?.removeAllRanges()
    this.hide()
  }

  hide() {
    this.pinned = null
    this.paintPinned()
    this.host.hidden = true
  }
}
