import type { Editor } from '@tiptap/core'

interface Heading {
  level: number
  text: string
  pos: number
}

/** Left-hand table of contents built from the document's headings; click to jump. */
export class Outline {
  private headings: Heading[] = []
  private items: HTMLElement[] = []
  private raf = 0
  private list: HTMLElement

  constructor(
    root: HTMLElement,
    private editor: Editor,
    private scroller: HTMLElement,
  ) {
    this.list = document.createElement('nav')
    this.list.className = 'outline-list'
    const head = document.createElement('div')
    head.className = 'side-head'
    head.innerHTML = '<strong>Outline</strong>'
    root.replaceChildren(head, this.list)

    editor.on('update', () => this.schedule())
    scroller.addEventListener('scroll', () => this.highlight(), { passive: true })
    this.rebuild()
  }

  private schedule() {
    cancelAnimationFrame(this.raf)
    this.raf = requestAnimationFrame(() => this.rebuild())
  }

  rebuild() {
    const found: Heading[] = []
    this.editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'heading' && node.textContent.trim()) {
        found.push({ level: node.attrs.level, text: node.textContent.trim(), pos })
      }
    })
    this.headings = found
    const minLevel = Math.min(...found.map((f) => f.level), 6)

    this.list.replaceChildren()
    this.items = []
    if (!found.length) {
      const empty = document.createElement('p')
      empty.className = 'side-empty'
      empty.textContent = 'No headings yet. Paragraphs formatted as Heading 1–6 appear here.'
      this.list.append(empty)
      return
    }
    for (const [i, hd] of found.entries()) {
      const a = document.createElement('button')
      a.type = 'button'
      a.className = 'outline-item'
      a.style.paddingLeft = `${10 + (hd.level - minLevel) * 14}px`
      a.textContent = hd.text
      a.title = hd.text
      a.addEventListener('click', () => this.jump(i))
      this.list.append(a)
      this.items.push(a)
    }
    this.highlight()
  }

  private domFor(i: number): HTMLElement | null {
    const hd = this.headings[i]
    const node = this.editor.view.nodeDOM(hd.pos)
    return node instanceof HTMLElement ? node : null
  }

  private jump(i: number) {
    const el = this.domFor(i)
    if (!el) return
    this.editor.commands.setTextSelection(this.headings[i].pos + 1)
    const top = el.getBoundingClientRect().top - this.scroller.getBoundingClientRect().top + this.scroller.scrollTop - 16
    this.scroller.scrollTo({ top, behavior: 'smooth' })
    this.editor.view.dom.focus({ preventScroll: true })
  }

  /** Mark the last heading at or above the top of the viewport. */
  private highlight() {
    if (!this.items.length) return
    const limit = this.scroller.getBoundingClientRect().top + 40
    let current = 0
    for (let i = 0; i < this.headings.length; i++) {
      const el = this.domFor(i)
      if (el && el.getBoundingClientRect().top <= limit) current = i
    }
    this.items.forEach((it, i) => it.classList.toggle('active', i === current))
  }
}
