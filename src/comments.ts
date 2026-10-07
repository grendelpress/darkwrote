import type { Editor } from '@tiptap/core'
import type { CommentData } from './docx/import'

const AUTHOR_KEY = 'darkwrote.author'

function h<K extends keyof HTMLElementTagNameMap>(tag: K, props: Record<string, any> = {}, ...children: (Node | string)[]): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag)
  for (const [k, v] of Object.entries(props)) (node as any)[k] = v
  node.append(...children)
  return node
}

interface Anchor {
  from: number
  to: number
  quote: string
}

export function getAuthor(): string {
  try {
    return localStorage.getItem(AUTHOR_KEY) || 'Author'
  } catch {
    return 'Author'
  }
}

function newId(): string {
  return 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
}

/** Sidebar listing the document's comments, with add / edit / delete and click-to-locate. */
export class CommentsPanel {
  private store = new Map<string, CommentData>()
  private draft: { from: number; to: number; quote: string } | null = null
  private editing: string | null = null
  private activeId: string | null = null
  private raf = 0
  private list: HTMLElement
  private nameInput: HTMLInputElement

  constructor(
    private root: HTMLElement,
    private editor: Editor,
    private onChange: () => void,
    private reveal: () => void,
  ) {
    this.nameInput = h('input', { type: 'text', value: getAuthor(), title: 'Name shown on your comments', placeholder: 'Your name' })
    this.nameInput.addEventListener('change', () => {
      try {
        localStorage.setItem(AUTHOR_KEY, this.nameInput.value.trim() || 'Author')
      } catch {
        /* ignore */
      }
    })
    this.list = h('div', { className: 'comment-list' })
    this.root.replaceChildren(h('div', { className: 'side-head' }, h('strong', {}, 'Comments'), this.nameInput), this.list)

    editor.on('update', () => this.schedule())
    editor.on('selectionUpdate', () => this.syncSelection())
    editor.view.dom.addEventListener('click', () => queueMicrotask(() => this.syncSelection()))
    this.render()
  }

  setComments(comments: CommentData[]) {
    this.store = new Map(comments.map((c) => [c.id, { ...c }]))
    this.draft = null
    this.editing = null
    this.activeId = null
    this.render()
  }

  /** Comments that are still anchored to text — the ones that should be saved. */
  anchoredComments(): CommentData[] {
    const anchors = this.anchors()
    return [...this.store.values()].filter((c) => anchors.has(c.id))
  }

  hasSelection(): boolean {
    return !this.editor.state.selection.empty
  }

  // ----- actions -----

  startDraft() {
    const { from, to } = this.editor.state.selection
    if (from === to) return
    this.draft = { from, to, quote: this.editor.state.doc.textBetween(from, to, ' ') }
    this.reveal()
    this.render()
    this.list.querySelector<HTMLTextAreaElement>('.draft textarea')?.focus()
  }

  private commitDraft(text: string) {
    const d = this.draft
    if (!d || !text.trim()) return
    const id = newId()
    this.store.set(id, { id, author: this.nameInput.value.trim() || 'Author', date: new Date().toISOString(), text: text.trim() })
    const markType = this.editor.schema.marks.comment
    this.editor.view.dispatch(this.editor.state.tr.addMark(d.from, d.to, markType.create({ commentId: id })))
    this.draft = null
    this.activeId = id
    this.onChange()
    this.render()
  }

  private remove(id: string) {
    const { state } = this.editor
    const tr = state.tr
    state.doc.descendants((node, pos) => {
      for (const m of node.marks) {
        if (m.type.name === 'comment' && m.attrs.commentId === id) tr.removeMark(pos, pos + node.nodeSize, m)
      }
    })
    // The store entry is kept so that Undo brings the comment back; it is dropped on save if unanchored.
    this.editor.view.dispatch(tr)
    this.onChange()
  }

  private locate(id: string) {
    const a = this.anchors().get(id)
    if (!a) return
    this.activeId = id
    this.editor.chain().focus().setTextSelection({ from: a.from, to: a.from }).run()
    const dom = this.editor.view.domAtPos(a.from).node
    ;(dom instanceof Element ? dom : dom.parentElement)?.scrollIntoView({ block: 'center', behavior: 'smooth' })
    this.render()
  }

  // ----- model -----

  private anchors(): Map<string, Anchor> {
    const map = new Map<string, Anchor>()
    this.editor.state.doc.descendants((node, pos) => {
      if (!node.isText) return
      for (const m of node.marks) {
        if (m.type.name !== 'comment') continue
        const id = String(m.attrs.commentId)
        const cur = map.get(id)
        const to = pos + node.nodeSize
        if (cur) {
          cur.to = to
          cur.quote += node.text
        } else map.set(id, { from: pos, to, quote: node.text ?? '' })
      }
    })
    return map
  }

  private syncSelection() {
    const marks = this.editor.state.selection.$from.marks().filter((m) => m.type.name === 'comment')
    const id = marks.length ? String(marks[marks.length - 1].attrs.commentId) : null
    if (id !== this.activeId) {
      this.activeId = id
      this.render()
      if (id) this.list.querySelector(`[data-id="${id}"]`)?.scrollIntoView({ block: 'nearest', behavior: 'smooth' })
    }
  }

  private schedule() {
    cancelAnimationFrame(this.raf)
    this.raf = requestAnimationFrame(() => {
      // don't clobber a text box that is being typed in
      if (!this.list.contains(document.activeElement)) this.render()
    })
  }

  // ----- view -----

  private render() {
    const anchors = this.anchors()
    const items = [...this.store.values()]
      .filter((c) => anchors.has(c.id))
      .sort((a, b) => anchors.get(a.id)!.from - anchors.get(b.id)!.from)

    const out: HTMLElement[] = []
    if (this.draft) out.push(this.draftCard(this.draft))
    for (const c of items) out.push(this.card(c, anchors.get(c.id)!))
    if (!out.length) out.push(h('p', { className: 'side-empty' }, 'No comments yet. Select some text and press the 💬 button to add one.'))
    this.list.replaceChildren(...out)
  }

  private draftCard(d: { quote: string }) {
    const ta = h('textarea', { rows: 3, placeholder: 'Write a comment…' })
    const save = h('button', { className: 'btn primary', type: 'button', textContent: 'Comment' })
    const cancel = h('button', { className: 'btn', type: 'button', textContent: 'Cancel' })
    save.addEventListener('click', () => this.commitDraft(ta.value))
    cancel.addEventListener('click', () => {
      this.draft = null
      this.render()
    })
    ta.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) this.commitDraft(ta.value)
      if (e.key === 'Escape') cancel.click()
    })
    return h('div', { className: 'comment draft' }, h('blockquote', {}, d.quote), ta, h('div', { className: 'row' }, save, cancel))
  }

  private card(c: CommentData, a: Anchor) {
    const card = h('div', { className: 'comment' + (c.id === this.activeId ? ' active' : '') })
    card.dataset.id = c.id
    const when = c.date ? new Date(c.date).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' }) : ''
    card.append(
      h('div', { className: 'meta' }, h('strong', {}, c.author), h('small', {}, when)),
      h('blockquote', {}, a.quote.length > 120 ? a.quote.slice(0, 117) + '…' : a.quote),
    )

    if (this.editing === c.id) {
      const ta = h('textarea', { rows: 3, value: c.text })
      const save = h('button', { className: 'btn primary', type: 'button', textContent: 'Save' })
      const cancel = h('button', { className: 'btn', type: 'button', textContent: 'Cancel' })
      save.addEventListener('click', () => {
        if (ta.value.trim()) {
          c.text = ta.value.trim()
          this.onChange()
        }
        this.editing = null
        this.render()
      })
      cancel.addEventListener('click', () => {
        this.editing = null
        this.render()
      })
      card.append(ta, h('div', { className: 'row' }, save, cancel))
      queueMicrotask(() => ta.focus())
    } else {
      const edit = h('button', { className: 'link-btn', type: 'button', textContent: 'Edit' })
      const del = h('button', { className: 'link-btn', type: 'button', textContent: 'Delete' })
      edit.addEventListener('click', (e) => {
        e.stopPropagation()
        this.editing = c.id
        this.render()
      })
      del.addEventListener('click', (e) => {
        e.stopPropagation()
        this.remove(c.id)
      })
      card.append(h('p', { className: 'body' }, c.text), h('div', { className: 'row' }, edit, del))
      card.addEventListener('click', () => this.locate(c.id))
    }
    return card
  }
}
