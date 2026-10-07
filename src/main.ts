import './style.css'
import { Editor } from '@tiptap/core'
import { EditorState } from '@tiptap/pm/state'
import { extensions } from './extensions'
import { buildToolbar } from './toolbar'
import { importDocx, DEFAULT_META, type DocMeta } from './docx/import'
import { exportDocx } from './docx/export'
import { allThemes, applyTheme, getActiveThemeId, getAdaptColors, setAdaptColors, BUILTIN_THEMES } from './themes'
import { openThemeEditor, refreshThemeSelect } from './themeEditor'

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

// File System Access API (Chromium) lets us save straight back to the opened file.
interface FsHandle {
  name: string
  getFile(): Promise<File>
  createWritable(): Promise<{ write(data: Blob): Promise<void>; close(): Promise<void> }>
}
type FsWindow = Window & {
  showOpenFilePicker?: (o: unknown) => Promise<FsHandle[]>
  showSaveFilePicker?: (o: unknown) => Promise<FsHandle>
}
const fsw = window as FsWindow
const pickerTypes = [{ description: 'Word document', accept: { [DOCX_MIME]: ['.docx'] } }]

let handle: FsHandle | null = null
let fileName = 'Untitled.docx'
let meta: DocMeta = { ...DEFAULT_META }
let dirty = false

// ---------- theme ----------
const themeSelect = $<HTMLSelectElement>('theme-select')
function switchTheme(id: string) {
  const theme = allThemes().find((t) => t.id === id) ?? BUILTIN_THEMES[0]
  applyTheme(theme)
  refreshThemeSelect(themeSelect, theme.id)
}
setAdaptColors(getAdaptColors())
switchTheme(getActiveThemeId())
themeSelect.addEventListener('change', () => switchTheme(themeSelect.value))
$('btn-themes').addEventListener('click', () =>
  openThemeEditor(
    $('theme-dialog') as HTMLDialogElement,
    themeSelect.value,
    (t) => {
      applyTheme(t)
      refreshThemeSelect(themeSelect, t.id)
    },
    toast,
  ),
)

// ---------- toast ----------
let toastTimer: number | undefined
function toast(msg: string, error = false) {
  const t = $('toast')
  t.textContent = msg
  t.classList.toggle('error', error)
  t.hidden = false
  clearTimeout(toastTimer)
  toastTimer = window.setTimeout(() => (t.hidden = true), error ? 6000 : 3000)
}

// ---------- editor ----------
const editor = new Editor({
  element: $('editor'),
  extensions,
  content: '<p></p>',
  autofocus: 'end',
  onUpdate: () => setDirty(true),
})
buildToolbar($('toolbar'), editor)

function setDirty(v: boolean) {
  dirty = v
  $('dirty').hidden = !v
  document.title = `${v ? '• ' : ''}${fileName} – Darkwrote`
}

function applyMeta(m: DocMeta) {
  meta = m
  const px = (twips: number) => `${twips / 15}px`
  const s = $('page').style
  s.setProperty('--page-width', px(m.pageWidth))
  s.setProperty('--pad-top', px(m.marginTop))
  s.setProperty('--pad-right', px(m.marginRight))
  s.setProperty('--pad-bottom', px(m.marginBottom))
  s.setProperty('--pad-left', px(m.marginLeft))
}

function setName(name: string) {
  fileName = name
  $('doc-name').textContent = name
  setDirty(dirty)
}

function confirmDiscard(): boolean {
  return !dirty || confirm('You have unsaved changes. Discard them?')
}

async function loadFile(file: File, newHandle: FsHandle | null) {
  try {
    const { doc, meta: m, warnings } = await importDocx(await file.arrayBuffer())
    editor.commands.setContent(doc, false)
    // a fresh undo history so Ctrl+Z can't undo the import itself
    editor.view.updateState(EditorState.create({ doc: editor.state.doc, plugins: editor.state.plugins }))
    handle = newHandle
    applyMeta(m)
    dirty = false
    setName(file.name)
    toast(warnings.length ? `Opened with ${warnings.length} warning(s): ${warnings[0]}` : `Opened ${file.name}`)
  } catch (e) {
    toast((e as Error).message, true)
  }
}

async function openFile() {
  if (!confirmDiscard()) return
  if (fsw.showOpenFilePicker) {
    try {
      const [h] = await fsw.showOpenFilePicker({ types: pickerTypes })
      await loadFile(await h.getFile(), h)
    } catch (e) {
      if ((e as Error).name !== 'AbortError') toast((e as Error).message, true)
    }
  } else {
    $<HTMLInputElement>('file-input').click()
  }
}

async function build(): Promise<Blob> {
  return exportDocx(editor.getJSON(), meta)
}

function download(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob)
  const a = Object.assign(document.createElement('a'), { href: url, download: name })
  document.body.append(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

async function save(forceAs = false) {
  try {
    const blob = await build()
    if (handle && !forceAs) {
      const w = await handle.createWritable()
      await w.write(blob)
      await w.close()
    } else if (fsw.showSaveFilePicker) {
      handle = await fsw.showSaveFilePicker({ suggestedName: fileName, types: pickerTypes })
      const w = await handle.createWritable()
      await w.write(blob)
      await w.close()
      setName(handle.name)
    } else {
      const name = fileName.toLowerCase().endsWith('.docx') ? fileName : fileName + '.docx'
      download(blob, name)
    }
    setDirty(false)
    toast(handle ? `Saved ${handle.name}` : 'Downloaded a copy')
  } catch (e) {
    if ((e as Error).name !== 'AbortError') toast(`Save failed: ${(e as Error).message}`, true)
  }
}

$('btn-open').addEventListener('click', openFile)
$('btn-save').addEventListener('click', () => save())
$('btn-saveas').addEventListener('click', () => save(true))
$('btn-new').addEventListener('click', () => {
  if (!confirmDiscard()) return
  editor.commands.setContent('<p></p>', false)
  handle = null
  applyMeta({ ...DEFAULT_META })
  dirty = false
  setName('Untitled.docx')
})
$<HTMLInputElement>('file-input').addEventListener('change', (e) => {
  const input = e.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (file && confirmDiscard()) void loadFile(file, null)
})

document.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return
  const k = e.key.toLowerCase()
  if (k === 's') {
    e.preventDefault()
    void save(e.shiftKey)
  } else if (k === 'o') {
    e.preventDefault()
    void openFile()
  }
})

window.addEventListener('beforeunload', (e) => {
  if (dirty) e.preventDefault()
})

// Drag & drop a .docx anywhere on the window
let dragDepth = 0
const hint = $('drop-hint')
window.addEventListener('dragenter', (e) => {
  if (e.dataTransfer?.types.includes('Files')) {
    dragDepth++
    hint.hidden = false
  }
})
window.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1)
  if (!dragDepth) hint.hidden = true
})
window.addEventListener('dragover', (e) => e.preventDefault())
window.addEventListener('drop', (e) => {
  e.preventDefault()
  dragDepth = 0
  hint.hidden = true
  const file = Array.from(e.dataTransfer?.files ?? []).find((f) => f.name.toLowerCase().endsWith('.docx'))
  if (file && confirmDiscard()) void loadFile(file, null)
})

setName(fileName)
