import {
  BUILTIN_THEMES,
  COLOR_LABELS,
  allThemes,
  getAdaptColors,
  isValidTheme,
  loadCustomThemes,
  newThemeId,
  saveCustomThemes,
  setAdaptColors,
  themeToJson,
  type Theme,
  type ThemeColors,
} from './themes'

type Notify = (msg: string, error?: boolean) => void

function h<K extends keyof HTMLElementTagNameMap>(tag: K, props: Record<string, any> = {}, ...children: (Node | string)[]): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag)
  for (const [k, v] of Object.entries(props)) {
    if (k === 'style') node.setAttribute('style', String(v))
    else (node as any)[k] = v
  }
  node.append(...children)
  return node
}

/** Populates the toolbar theme <select> and keeps it in sync with stored themes. */
export function refreshThemeSelect(select: HTMLSelectElement, activeId: string) {
  select.replaceChildren()
  const group = (label: string, themes: Theme[]) => {
    if (!themes.length) return
    const g = h('optgroup', { label })
    for (const t of themes) g.append(h('option', { value: t.id, textContent: t.name }))
    select.append(g)
  }
  const all = allThemes()
  group('Dark', all.filter((t) => t.mode === 'dark' && t.builtin))
  group('Light', all.filter((t) => t.mode === 'light' && t.builtin))
  group('Custom', all.filter((t) => !t.builtin))
  select.value = all.some((t) => t.id === activeId) ? activeId : BUILTIN_THEMES[0].id
}

export function openThemeEditor(dialog: HTMLDialogElement, activeId: string, onChange: (t: Theme) => void, notify: Notify) {
  let selectedId = activeId
  const current = () => allThemes().find((t) => t.id === selectedId) ?? BUILTIN_THEMES[0]

  const render = () => {
    const themes = allThemes()
    const theme = current()
    const editable = !theme.builtin

    // ---- list ----
    const list = h('ul', { className: 'theme-list' })
    for (const t of themes) {
      const li = h('li', { className: t.id === selectedId ? 'sel' : '' })
      li.append(h('span', {}, t.name), h('small', {}, t.mode))
      const sw = h('span', { className: 'swatches' })
      for (const k of ['bg', 'paper', 'text', 'accent'] as const) sw.append(h('i', { style: `background:${t.colors[k]}` }))
      li.append(sw)
      li.addEventListener('click', () => {
        selectedId = t.id
        onChange(t)
        render()
      })
      list.append(li)
    }

    const importInput = h('input', { type: 'file', accept: '.json,application/json', hidden: true })
    importInput.addEventListener('change', async () => {
      const file = importInput.files?.[0]
      if (!file) return
      try {
        const data = JSON.parse(await file.text())
        if (!isValidTheme(data)) throw new Error('missing name, mode or colours')
        const t: Theme = { id: newThemeId(), name: data.name, mode: data.mode, colors: data.colors, builtin: false }
        saveCustomThemes([...loadCustomThemes(), t])
        selectedId = t.id
        onChange(t)
        render()
        notify(`Imported theme “${t.name}”`)
      } catch (e) {
        notify(`Couldn't import theme: ${(e as Error).message}`, true)
      }
    })

    const duplicate = h('button', { className: 'btn', type: 'button', textContent: editable ? 'Duplicate' : 'Duplicate to edit' })
    duplicate.addEventListener('click', () => {
      const copy: Theme = { id: newThemeId(), name: `${theme.name} copy`, mode: theme.mode, colors: { ...theme.colors }, builtin: false }
      saveCustomThemes([...loadCustomThemes(), copy])
      selectedId = copy.id
      onChange(copy)
      render()
    })

    const left = h('div', {}, h('h3', {}, 'Themes'), list, h('div', { className: 'dialog-actions', style: 'padding:0' }, duplicate,
      (() => {
        const b = h('button', { className: 'btn', type: 'button', textContent: 'Import…' })
        b.addEventListener('click', () => importInput.click())
        return b
      })(), importInput))

    // ---- editor ----
    const right = h('div', {})
    right.append(h('h3', {}, editable ? 'Edit theme' : 'Built-in theme (read-only)'))

    const persist = (patch: Partial<Theme>) => {
      const customs = loadCustomThemes().map((t) => (t.id === theme.id ? { ...t, ...patch } : t))
      saveCustomThemes(customs)
      const updated = customs.find((t) => t.id === theme.id)!
      onChange(updated)
    }

    const nameInput = h('input', { type: 'text', value: theme.name, disabled: !editable })
    nameInput.addEventListener('change', () => persist({ name: nameInput.value.trim() || 'Untitled theme' }))
    const modeSelect = h('select', { disabled: !editable })
    modeSelect.append(h('option', { value: 'dark', textContent: 'Dark' }), h('option', { value: 'light', textContent: 'Light' }))
    modeSelect.value = theme.mode
    modeSelect.title = 'Dark themes brighten dark text from Word documents; light themes darken pale text'
    modeSelect.addEventListener('change', () => {
      persist({ mode: modeSelect.value as Theme['mode'] })
    })
    right.append(h('div', { className: 'name-row' }, nameInput, modeSelect))

    const preview = h('div', { className: 'preview' })
    const paintPreview = (t: Theme) => {
      preview.style.cssText = `background:${t.colors.paper};color:${t.colors.text};border-color:${t.colors.border}`
      preview.replaceChildren(
        h('strong', {}, 'The quick brown fox'),
        ' jumps over the ',
        h('a', { href: '#', style: `color:${t.colors.link}`, onclick: (e: Event) => e.preventDefault() }, 'lazy dog'),
        '. ',
        h('span', { style: `color:${t.colors.muted}` }, 'Secondary text. '),
        h('span', { style: `background:${t.colors.selection}` }, 'Selected text.'),
      )
    }

    for (const key of Object.keys(COLOR_LABELS) as (keyof ThemeColors)[]) {
      const picker = h('input', { type: 'color', value: theme.colors[key], disabled: !editable })
      const hexInput = h('input', { type: 'text', value: theme.colors[key], disabled: !editable, maxLength: 7 })
      const set = (v: string) => {
        const colors = { ...current().colors, [key]: v }
        persist({ colors })
        picker.value = v
        hexInput.value = v
        paintPreview({ ...current(), colors })
      }
      picker.addEventListener('input', () => set(picker.value))
      hexInput.addEventListener('change', () => {
        const v = hexInput.value.startsWith('#') ? hexInput.value : '#' + hexInput.value
        if (/^#[0-9a-fA-F]{6}$/.test(v)) set(v.toLowerCase())
        else hexInput.value = current().colors[key]
      })
      right.append(h('label', { className: 'field' }, COLOR_LABELS[key], picker, hexInput))
    }
    paintPreview(theme)
    right.append(preview)

    const body = h('div', { className: 'dialog-body' }, left, right)

    // ---- actions ----
    const adapt = h('input', { type: 'checkbox', checked: getAdaptColors() })
    adapt.addEventListener('change', () => setAdaptColors(adapt.checked))
    const adaptLabel = h('label', { className: 'check', title: 'Shift text and cell colours from the Word file so they stay readable (the file itself is never changed)' }, adapt, 'Adapt document colours to theme')

    const exportBtn = h('button', { className: 'btn', type: 'button', textContent: 'Export JSON' })
    exportBtn.addEventListener('click', () => {
      const url = URL.createObjectURL(new Blob([themeToJson(theme)], { type: 'application/json' }))
      const a = h('a', { href: url, download: `${theme.name.replace(/[^\w-]+/g, '-').toLowerCase()}.theme.json` })
      a.click()
      setTimeout(() => URL.revokeObjectURL(url), 1000)
    })
    const del = h('button', { className: 'btn danger', type: 'button', textContent: 'Delete theme', disabled: !editable })
    del.addEventListener('click', () => {
      if (!confirm(`Delete theme “${theme.name}”?`)) return
      saveCustomThemes(loadCustomThemes().filter((t) => t.id !== theme.id))
      selectedId = BUILTIN_THEMES[0].id
      onChange(BUILTIN_THEMES[0])
      render()
    })
    const close = h('button', { className: 'btn primary', type: 'button', textContent: 'Done' })
    close.addEventListener('click', () => dialog.close())

    const actions = h('div', { className: 'dialog-actions' }, adaptLabel, h('span', { className: 'spacer' }), exportBtn, del, close)
    dialog.replaceChildren(h('h2', {}, 'Themes'), body, actions)
  }

  render()
  if (!dialog.open) dialog.showModal()
}
