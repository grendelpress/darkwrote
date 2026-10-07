import type { Editor } from '@tiptap/core'

const FONTS = ['Calibri', 'Cambria', 'Arial', 'Verdana', 'Segoe UI', 'Georgia', 'Times New Roman', 'Garamond', 'Courier New', 'Consolas', 'Comic Sans MS']
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]

interface Control {
  update: () => void
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, props: Record<string, any> = {}, ...children: (Node | string)[]): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag)
  for (const [k, v] of Object.entries(props)) {
    if (k === 'style') node.setAttribute('style', String(v))
    else (node as any)[k] = v
  }
  node.append(...children)
  return node
}

export function buildToolbar(host: HTMLElement, editor: Editor) {
  const controls: Control[] = []
  const sep = () => host.append(el('span', { className: 'tb-sep' }))

  const button = (label: string, title: string, run: () => void, active?: () => boolean, enabled?: () => boolean) => {
    const b = el('button', { className: 'tb-btn', type: 'button', title, innerHTML: label })
    b.setAttribute('aria-label', title)
    // mousedown preventDefault keeps the editor selection while clicking toolbar buttons
    b.addEventListener('mousedown', (e) => e.preventDefault())
    b.addEventListener('click', run)
    host.append(b)
    controls.push({
      update() {
        b.classList.toggle('active', !!active?.())
        b.disabled = enabled ? !enabled() : false
      },
    })
    return b
  }

  const chain = () => editor.chain().focus()

  button('↶', 'Undo (Ctrl+Z)', () => chain().undo().run(), undefined, () => editor.can().undo())
  button('↷', 'Redo (Ctrl+Y)', () => chain().redo().run(), undefined, () => editor.can().redo())
  sep()

  // Paragraph style
  const style = el('select', { title: 'Paragraph style' })
  style.append(el('option', { value: 'p', textContent: 'Normal' }))
  for (let i = 1; i <= 6; i++) style.append(el('option', { value: String(i), textContent: `Heading ${i}` }))
  style.addEventListener('change', () => {
    const v = style.value
    if (v === 'p') chain().setParagraph().run()
    else chain().setHeading({ level: Number(v) as 1 | 2 | 3 | 4 | 5 | 6 }).run()
  })
  host.append(style)
  controls.push({
    update() {
      const lvl = [1, 2, 3, 4, 5, 6].find((l) => editor.isActive('heading', { level: l }))
      style.value = lvl ? String(lvl) : 'p'
    },
  })

  // Font family
  const font = el('select', { title: 'Font' })
  const fontOptions = (current: string | undefined) => {
    font.replaceChildren(el('option', { value: '', textContent: 'Default' }))
    const list = current && !FONTS.includes(current) ? [current, ...FONTS] : FONTS
    for (const f of list) font.append(el('option', { value: f, textContent: f }))
  }
  fontOptions(undefined)
  font.addEventListener('change', () => (font.value ? chain().setFontFamily(font.value).run() : chain().unsetFontFamily().run()))
  host.append(font)
  controls.push({
    update() {
      const cur = editor.getAttributes('textStyle').fontFamily as string | undefined
      const clean = cur?.split(',')[0].replace(/["']/g, '').trim()
      if (clean && !Array.from(font.options).some((o) => o.value === clean)) fontOptions(clean)
      font.value = clean ?? ''
    },
  })

  // Font size
  const size = el('select', { title: 'Font size (pt)' })
  const sizeOptions = (current: string | undefined) => {
    size.replaceChildren(el('option', { value: '', textContent: 'Size' }))
    const list = SIZES.map(String)
    if (current && !list.includes(current)) list.push(current)
    list.sort((a, b) => Number(a) - Number(b))
    for (const s of list) size.append(el('option', { value: s, textContent: s }))
  }
  sizeOptions(undefined)
  size.addEventListener('change', () => (size.value ? chain().setFontSize(`${size.value}pt`).run() : chain().unsetFontSize().run()))
  host.append(size)
  controls.push({
    update() {
      const cur = (editor.getAttributes('textStyle').fontSize as string | undefined)?.replace('pt', '')
      if (cur && !Array.from(size.options).some((o) => o.value === cur)) sizeOptions(cur)
      size.value = cur ?? ''
    },
  })
  sep()

  button('<b>B</b>', 'Bold (Ctrl+B)', () => chain().toggleBold().run(), () => editor.isActive('bold'))
  button('<i>I</i>', 'Italic (Ctrl+I)', () => chain().toggleItalic().run(), () => editor.isActive('italic'))
  button('<u>U</u>', 'Underline (Ctrl+U)', () => chain().toggleUnderline().run(), () => editor.isActive('underline'))
  button('<s>S</s>', 'Strikethrough', () => chain().toggleStrike().run(), () => editor.isActive('strike'))
  button('x<sub>2</sub>', 'Subscript', () => chain().toggleSubscript().run(), () => editor.isActive('subscript'))
  button('x<sup>2</sup>', 'Superscript', () => chain().toggleSuperscript().run(), () => editor.isActive('superscript'))
  sep()

  const colorPicker = (label: string, title: string, initial: string, apply: (c: string) => void, clear: () => void) => {
    const input = el('input', { type: 'color', value: initial })
    const wrap = el('label', { className: 'tb-color', title }, label, input)
    input.addEventListener('input', () => apply(input.value))
    const reset = el('button', { className: 'tb-btn', type: 'button', title: `Clear ${title.toLowerCase()}`, textContent: '✕' })
    reset.addEventListener('mousedown', (e) => e.preventDefault())
    reset.addEventListener('click', clear)
    host.append(wrap, reset)
  }
  colorPicker('A', 'Text color', '#cc0000', (c) => chain().setColor(c).run(), () => chain().unsetColor().run())
  colorPicker('🖍', 'Highlight', '#ffff00', (c) => chain().setHighlight({ color: c }).run(), () => chain().unsetHighlight().run())
  sep()

  for (const [label, name, value] of [
    ['⯇', 'Align left', 'left'],
    ['≡', 'Center', 'center'],
    ['⯈', 'Align right', 'right'],
    ['☰', 'Justify', 'justify'],
  ] as const) {
    button(label, name, () => chain().setTextAlign(value).run(), () => editor.isActive({ textAlign: value }))
  }
  sep()

  button('• List', 'Bulleted list', () => chain().toggleBulletList().run(), () => editor.isActive('bulletList'))
  button('1. List', 'Numbered list', () => chain().toggleOrderedList().run(), () => editor.isActive('orderedList'))
  button('❝', 'Quote', () => chain().toggleBlockquote().run(), () => editor.isActive('blockquote'))
  button('―', 'Horizontal rule', () => chain().setHorizontalRule().run())
  sep()

  button('🔗', 'Link', () => {
    const prev = editor.getAttributes('link').href as string | undefined
    const url = window.prompt('Link address (leave empty to remove)', prev ?? 'https://')
    if (url === null) return
    if (url.trim() === '') chain().extendMarkRange('link').unsetLink().run()
    else chain().extendMarkRange('link').setLink({ href: url.trim() }).run()
  }, () => editor.isActive('link'))

  const imageInput = el('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif', hidden: true })
  imageInput.addEventListener('change', () => {
    const file = imageInput.files?.[0]
    imageInput.value = ''
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      const src = String(reader.result)
      const probe = new Image()
      probe.onload = () => {
        const w = Math.min(probe.naturalWidth, 480)
        const h = Math.round((probe.naturalHeight * w) / probe.naturalWidth)
        chain().setImage({ src, width: w, height: h } as never).run()
      }
      probe.src = src
    }
    reader.readAsDataURL(file)
  })
  host.append(imageInput)
  button('🖼', 'Insert image', () => imageInput.click())
  sep()

  const inTable = () => editor.isActive('table')
  button('▦', 'Insert table', () => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: false }).run(), undefined, () => !inTable())
  button('+Row', 'Add row below', () => chain().addRowAfter().run(), undefined, inTable)
  button('+Col', 'Add column right', () => chain().addColumnAfter().run(), undefined, inTable)
  button('−Row', 'Delete row', () => chain().deleteRow().run(), undefined, inTable)
  button('−Col', 'Delete column', () => chain().deleteColumn().run(), undefined, inTable)
  button('Merge', 'Merge/split selected cells', () => chain().mergeOrSplit().run(), undefined, inTable)
  button('✕▦', 'Delete table', () => chain().deleteTable().run(), undefined, inTable)

  const update = () => controls.forEach((c) => c.update())
  editor.on('transaction', update)
  update()
}
