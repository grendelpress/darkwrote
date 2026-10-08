import './style.css'
import { Editor } from '@tiptap/core'
import { EditorState } from '@tiptap/pm/state'
import { AddMarkStep, RemoveMarkStep, type Step } from '@tiptap/pm/transform'
import { extensions } from './extensions'
import { buildToolbar } from './toolbar'
import { importDocx, DEFAULT_META, type CommentData, type DocMeta } from './docx/import'
import { exportDocx } from './docx/export'
import { annotateDocx } from './docx/annotate'
import { collectAnnotations, placeable } from './annotations'
import { allThemes, applyTheme, getActiveThemeId, getAdaptColors, setAdaptColors, BUILTIN_THEMES } from './themes'
import { openThemeEditor, refreshThemeSelect } from './themeEditor'
import { CommentsPanel } from './comments'
import { Outline } from './outline'
import { store, requestPersistence, sha256Hex, type StoredDoc } from './storage'
import { applyReadPrefs, buildReadSettings, loadReadPrefs, scrollToPos, topPosition } from './reading'
import { SelectionBar, domSelectionRange, type Range } from './selbar'
import { DOCX_MIME, baseName, downloadBlob } from './files'
import { openExportDialog } from './exportDialog'
import { renderRecent } from './recent'
import { onLaunchFiles, registerServiceWorker, setupInstallButton, takeSharedFile } from './pwa'

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T
const compactQuery = window.matchMedia('(max-width: 899px)')
const isCompact = () => compactQuery.matches
const isTouch = window.matchMedia('(pointer: coarse)').matches

// ---------- file access ----------
// The File System Access API (desktop Chromium) lets us save straight back to the opened file.
// Android Chrome does not have it, so there we open through the file picker and export through downloads / share.
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

interface Current {
  id: string
  name: string
  handle: FsHandle | null
  meta: DocMeta
  /** Bytes of the .docx as opened — kept so annotations can be added without rebuilding the file */
  original: ArrayBuffer | null
  sha256: string | null
  size: number | null
  /** Text changed beyond comments/highlights: the original can no longer be patched */
  contentEdited: boolean
}

let cur: Current | null = null
let notExported = false
let pendingSave = false
let saveFailed = false
let suppress = false

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
$('btn-themes').addEventListener('click', () => {
  closeMenu()
  openThemeEditor(
    $('theme-dialog') as HTMLDialogElement,
    themeSelect.value,
    (t) => {
      applyTheme(t)
      refreshThemeSelect(themeSelect, t.id)
    },
    toast,
  )
})

// ---------- toast ----------
let toastTimer: number | undefined
function toast(msg: string, error = false, action?: { label: string; run: () => void }) {
  const t = $('toast')
  t.replaceChildren(document.createTextNode(msg))
  if (action) {
    const b = document.createElement('button')
    b.className = 'act'
    b.textContent = action.label
    b.addEventListener('click', () => {
      action.run()
      t.hidden = true
    })
    t.append(b)
  }
  t.classList.toggle('error', error)
  t.hidden = false
  clearTimeout(toastTimer)
  toastTimer = window.setTimeout(() => (t.hidden = true), error ? 6000 : action ? 6000 : 3000)
}

// ---------- editor ----------
const editor = new Editor({
  element: $('editor'),
  extensions,
  content: '<p></p>',
  editable: false,
  editorProps: {
    scrollMargin: { top: 60, bottom: 120, left: 0, right: 0 },
    scrollThreshold: { top: 60, bottom: 120, left: 0, right: 0 },
  },
  onUpdate: () => markChanged(),
})

function isAnnotationStep(step: Step): boolean {
  if (!(step instanceof AddMarkStep || step instanceof RemoveMarkStep)) return false
  const m = step.mark
  return m.type.name === 'comment' || (m.type.name === 'highlight' && !!m.attrs.user)
}
editor.on('transaction', ({ transaction: tr }) => {
  if (suppress || !cur || !tr.docChanged) return
  if (tr.steps.some((s) => !isAnnotationStep(s))) cur.contentEdited = true
})

// ---------- sidebars (drawers on small screens) ----------
const SIDE_KEY = 'darkwrote.sidebars'
let sides: { outline: boolean; comments: boolean } = { outline: true, comments: true }
try {
  sides = { ...sides, ...JSON.parse(localStorage.getItem(SIDE_KEY) ?? '{}') }
} catch {
  /* defaults */
}
let drawer: 'outline' | 'comments' | null = null

function paintSides() {
  const compact = isCompact()
  for (const k of ['outline', 'comments'] as const) {
    const open = compact ? drawer === k : sides[k]
    $(k).classList.toggle('open', open)
    $(`btn-${k}`).classList.toggle('on', open)
  }
  $('scrim').hidden = !(compact && drawer)
  if (!compact) {
    try {
      localStorage.setItem(SIDE_KEY, JSON.stringify(sides))
    } catch {
      /* ignore */
    }
  }
}
for (const k of ['outline', 'comments'] as const) {
  $(`btn-${k}`).addEventListener('click', () => {
    if (isCompact()) drawer = drawer === k ? null : k
    else sides[k] = !sides[k]
    paintSides()
  })
}
$('scrim').addEventListener('click', () => {
  drawer = null
  paintSides()
})
compactQuery.addEventListener('change', () => {
  drawer = null
  closeMenu()
  paintSides()
})

const outline = new Outline($('outline'), editor, $('workspace'), () => {
  if (isCompact()) {
    drawer = null
    paintSides()
  }
})
const commentsPanel = new CommentsPanel(
  $('comments'),
  editor,
  () => markChanged(),
  () => {
    if (isCompact()) drawer = 'comments'
    else sides.comments = true
    paintSides()
  },
  $('composer'),
  isCompact,
)
for (const head of document.querySelectorAll('.side-head')) {
  const close = document.createElement('button')
  close.className = 'side-close'
  close.type = 'button'
  close.textContent = '✕'
  close.setAttribute('aria-label', 'Close panel')
  close.addEventListener('click', () => {
    drawer = null
    paintSides()
  })
  head.append(close)
}
paintSides()

// ---------- selection → range (works with or without a caret) ----------
function currentRange(): Range | null {
  if (editor.isEditable) {
    const { from, to } = editor.state.selection
    return from < to ? { from, to } : null
  }
  return domSelectionRange(editor)
}
buildToolbar($('toolbar'), editor, () => commentsPanel.startDraft(currentRange()))

function setUserHighlight(r: Range, color: string) {
  const { state } = editor
  editor.view.dispatch(state.tr.addMark(r.from, r.to, state.schema.marks.highlight.create({ color, user: true })))
}
function clearUserHighlight(r: Range) {
  const { state } = editor
  const tr = state.tr
  state.doc.nodesBetween(r.from, r.to, (node, pos) => {
    if (!node.isText) return
    for (const m of node.marks) {
      if (m.type.name === 'highlight' && m.attrs.user) tr.removeMark(Math.max(pos, r.from), Math.min(pos + node.nodeSize, r.to), m)
    }
  })
  editor.view.dispatch(tr)
}
const selectionBar = new SelectionBar($('selbar'), editor, () => !!cur && !editor.isEditable, {
  highlight(r, color) {
    setUserHighlight(r, color)
    toast('Highlighted', false, { label: 'Undo', run: () => editor.commands.undo() })
  },
  clearHighlight: clearUserHighlight,
  comment: (r) => commentsPanel.startDraft(r),
})

// ---------- on-screen keyboard: keep bars and sheets above it ----------
const vv = window.visualViewport
function updateKeyboardOffset() {
  const kb = vv ? Math.max(0, window.innerHeight - vv.height - vv.offsetTop) : 0
  document.documentElement.style.setProperty('--kb', `${Math.round(kb)}px`)
}
vv?.addEventListener('resize', updateKeyboardOffset)
vv?.addEventListener('scroll', updateKeyboardOffset)

// ---------- read / edit mode ----------
type Mode = 'read' | 'edit'
const MODE_KEY = 'darkwrote.mode'
let mode: Mode = (() => {
  try {
    const m = localStorage.getItem(MODE_KEY)
    if (m === 'read' || m === 'edit') return m
  } catch {
    /* default below */
  }
  return isTouch ? 'read' : 'edit'
})()

const workspace = $('workspace')

/** Run a layout-changing action and keep the reader on the same passage afterwards. */
function keepPlace(change: () => void) {
  const pos = cur ? topPosition(editor, workspace) : null
  change()
  if (pos != null) requestAnimationFrame(() => scrollToPos(editor, workspace, pos))
}

function applyMode() {
  document.body.classList.toggle('reading', mode === 'read')
  editor.setEditable(!!cur && mode === 'edit', false)
  const btn = $('btn-mode')
  btn.textContent = mode === 'read' ? '✎ Edit' : '📖 Read'
  btn.title = mode === 'read' ? 'Switch to editing' : 'Switch to reading (no accidental edits, text reflows)'
  if (mode === 'edit') selectionBar.hide()
  $('read-settings').hidden = true
}
$('btn-mode').addEventListener('click', () => {
  keepPlace(() => {
    mode = mode === 'read' ? 'edit' : 'read'
    try {
      localStorage.setItem(MODE_KEY, mode)
    } catch {
      /* ignore */
    }
    applyMode()
  })
  toast(mode === 'read' ? 'Reading mode — text reflows and can’t be changed by accident' : 'Editing mode')
})

const prefs = loadReadPrefs()
applyReadPrefs(prefs)
buildReadSettings($('read-settings'), prefs, keepPlace)
$('btn-readset').addEventListener('click', (e) => {
  e.stopPropagation()
  $('read-settings').hidden = !$('read-settings').hidden
})
document.addEventListener('click', (e) => {
  const pop = $('read-settings')
  if (!pop.hidden && !pop.contains(e.target as Node) && e.target !== $('btn-readset')) pop.hidden = true
  const menu = $('menu')
  if (document.body.classList.contains('menu-open') && !menu.contains(e.target as Node) && e.target !== $('btn-menu')) closeMenu()
})

// ---------- menu (collapses behind ☰ on small screens) ----------
function closeMenu() {
  document.body.classList.remove('menu-open')
  $('btn-menu').setAttribute('aria-expanded', 'false')
}
$('btn-menu').addEventListener('click', () => {
  const open = document.body.classList.toggle('menu-open')
  $('btn-menu').setAttribute('aria-expanded', String(open))
})
$('menu').addEventListener('click', (e) => {
  if ((e.target as HTMLElement).closest('button') && !(e.target as HTMLElement).closest('select')) closeMenu()
})

// ---------- document state ----------
function setDirty(v: boolean) {
  notExported = v
  $('dirty').hidden = !v
  updateTitle()
}
function updateTitle() {
  document.title = cur ? `${notExported ? '• ' : ''}${cur.name} – Darkwrote` : 'Darkwrote'
}
function setName(name: string) {
  if (cur) cur.name = name
  $('doc-name').textContent = cur ? name : 'Untitled.docx'
  updateTitle()
}

function applyMeta(m: DocMeta) {
  if (cur) cur.meta = m
  const px = (twips: number) => `${twips / 15}px`
  const s = $('page').style
  s.setProperty('--page-width', px(m.pageWidth))
  s.setProperty('--pad-top', px(m.marginTop))
  s.setProperty('--pad-right', px(m.marginRight))
  s.setProperty('--pad-bottom', px(m.marginBottom))
  s.setProperty('--pad-left', px(m.marginLeft))
}

function showDocument(open: boolean) {
  document.body.classList.toggle('no-doc', !open)
  for (const id of ['btn-save', 'btn-saveas', 'btn-export', 'btn-close', 'btn-revert']) $<HTMLButtonElement>(id).disabled = !open
  $<HTMLButtonElement>('btn-revert').disabled = !open || !cur?.original
  if (!open) {
    drawer = null
    paintSides()
    selectionBar.hide()
    $('composer').hidden = true
    void renderRecent($('recent'), (id) => void restoreById(id))
  }
  applyMode()
}

function newId(): string {
  return 'new-' + Date.now().toString(36)
}

interface Loaded {
  doc: ReturnType<Editor['getJSON']>
  meta: DocMeta
  comments: CommentData[]
}

function installDocument(next: Current, loaded: Loaded, opts: { viewPos?: number; edited?: boolean; exported?: boolean }) {
  suppress = true
  editor.commands.setContent(loaded.doc, false)
  // a fresh undo history so Undo can't unwind the import itself
  editor.view.updateState(EditorState.create({ doc: editor.state.doc, plugins: editor.state.plugins }))
  suppress = false
  cur = next
  cur.contentEdited = !!opts.edited
  applyMeta(loaded.meta)
  commentsPanel.setComments(loaded.comments)
  outline.rebuild()
  setName(next.name)
  showDocument(true)
  setDirty(false)
  workspace.scrollTop = 0
  if (opts.viewPos != null) requestAnimationFrame(() => scrollToPos(editor, workspace, opts.viewPos!))
  try {
    localStorage.setItem('darkwrote.last', next.id)
  } catch {
    /* ignore */
  }
}

// ---------- autosave & recovery ----------
let saveTimer: number | undefined

function markChanged() {
  if (!cur) return
  setDirty(true)
  scheduleSave()
}
function scheduleSave() {
  pendingSave = true
  clearTimeout(saveTimer)
  saveTimer = window.setTimeout(() => void flushSave(), 1500)
}

async function flushSave(): Promise<void> {
  clearTimeout(saveTimer)
  if (!cur) return
  const c = cur
  const comments = commentsPanel.allComments()
  const records = collectAnnotations(editor.state.doc, comments)
  const snapshot: StoredDoc = {
    id: c.id,
    name: c.name,
    updatedAt: Date.now(),
    doc: editor.getJSON(),
    meta: c.meta,
    comments,
    sha256: c.sha256,
    size: c.size,
    contentEdited: c.contentEdited,
    annotationCount: records.filter((r) => !r.imported).length,
  }
  pendingSave = false
  saveFailed = !(await store.saveDoc(snapshot))
  saveViewNow()
}

let viewTimer: number | undefined
function saveViewNow() {
  clearTimeout(viewTimer)
  if (!cur) return
  const pos = topPosition(editor, workspace)
  if (pos == null) return
  void store.saveView({ id: cur.id, pos, progress: pos / Math.max(1, editor.state.doc.content.size), updatedAt: Date.now() })
}
workspace.addEventListener(
  'scroll',
  () => {
    clearTimeout(viewTimer)
    viewTimer = window.setTimeout(saveViewNow, 500)
  },
  { passive: true },
)
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') void flushSave()
})
window.addEventListener('pagehide', () => void flushSave())

/** Make sure the working copy is on disk before the current document is replaced or closed. */
async function leaveDocument(): Promise<boolean> {
  if (!cur) return true
  await flushSave()
  return !saveFailed || confirm('Your notes could not be saved on this device. Leave anyway?')
}

// ---------- opening ----------
async function restoreById(id: string) {
  if (!(await leaveDocument())) return
  const stored = await store.getDoc(id)
  if (!stored) return toast('That document is no longer on this device', true)
  const original = (await store.getOriginal(id)) ?? null
  const view = await store.getView(id)
  installDocument(
    { id, name: stored.name, handle: null, meta: stored.meta, original, sha256: stored.sha256, size: stored.size, contentEdited: stored.contentEdited },
    { doc: stored.doc, meta: stored.meta, comments: stored.comments },
    { viewPos: view?.pos, edited: stored.contentEdited },
  )
  toast(`Restored ${stored.name}`)
}

async function loadFile(file: File, handle: FsHandle | null) {
  try {
    if (!(await leaveDocument())) return
    const bytes = await file.arrayBuffer()
    const sha = await sha256Hex(bytes)
    const id = sha ? `f-${sha.slice(0, 32)}` : `f-${file.name}-${file.size}`
    const stored = await store.getDoc(id)
    const next: Current = { id, name: file.name, handle, meta: DEFAULT_META, original: bytes, sha256: sha, size: file.size, contentEdited: false }

    if (stored && (stored.annotationCount > 0 || stored.contentEdited)) {
      // The same file again: bring back the notes and reading position from last time.
      const view = await store.getView(id)
      installDocument(next, { doc: stored.doc, meta: stored.meta, comments: stored.comments }, { viewPos: view?.pos, edited: stored.contentEdited })
      toast(`Restored your notes in ${file.name}`, false, { label: 'Start over', run: () => void revertToOriginal(true) })
    } else {
      const { doc, meta, comments, warnings } = await importDocx(bytes)
      installDocument(next, { doc, meta, comments }, {})
      toast(warnings.length ? `Opened with ${warnings.length} warning(s): ${warnings[0]}` : `Opened ${file.name}`)
    }
    await store.saveOriginal(id, bytes)
    void requestPersistence()
    scheduleSave()
    setDirty(false)
  } catch (e) {
    toast((e as Error).message, true)
  }
}

async function openFile() {
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

async function revertToOriginal(silent = false) {
  if (!cur?.original) return
  if (!silent && !confirm('Discard all your comments, highlights and edits in this document and reload the original file?')) return
  try {
    const { doc, meta, comments } = await importDocx(cur.original)
    installDocument({ ...cur }, { doc, meta, comments }, {})
    await flushSave()
    toast('Back to the original document')
  } catch (e) {
    toast((e as Error).message, true)
  }
}

function resetDocument(open: boolean) {
  suppress = true
  editor.commands.setContent('<p></p>', false)
  editor.view.updateState(EditorState.create({ doc: editor.state.doc, plugins: editor.state.plugins }))
  suppress = false
  cur = open ? { id: newId(), name: 'Untitled.docx', handle: null, meta: { ...DEFAULT_META }, original: null, sha256: null, size: null, contentEdited: true } : null
  applyMeta({ ...DEFAULT_META })
  commentsPanel.setComments([])
  outline.rebuild()
  setName('Untitled.docx')
  showDocument(open)
  setDirty(false)
  if (open) {
    mode = 'edit'
    applyMode()
    editor.commands.focus('start')
  } else {
    try {
      localStorage.removeItem('darkwrote.last')
    } catch {
      /* ignore */
    }
  }
}

$('btn-open').addEventListener('click', () => void openFile())
$('empty-open').addEventListener('click', () => void openFile())
const newDocument = async () => {
  if (await leaveDocument()) resetDocument(true)
}
$('btn-new').addEventListener('click', () => void newDocument())
$('empty-new').addEventListener('click', () => void newDocument())
$('btn-close').addEventListener('click', async () => {
  if (!(await leaveDocument())) return
  resetDocument(false)
  toast('Closed — your notes are kept under “Continue where you left off”')
})
$('btn-revert').addEventListener('click', () => void revertToOriginal())
$<HTMLInputElement>('file-input').addEventListener('change', (e) => {
  const input = e.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (file) void loadFile(file, null)
})

// ---------- saving & exporting ----------
function records() {
  return collectAnnotations(editor.state.doc, commentsPanel.allComments())
}

async function buildAnnotated() {
  if (!cur?.original) throw new Error('There is no original file to annotate.')
  return annotateDocx(cur.original, placeable(records()))
}
const buildRebuilt = () =>
  exportDocx(editor.getJSON(), cur?.meta ?? DEFAULT_META, commentsPanel.anchoredComments())

function documentInfo() {
  return { name: cur?.name ?? 'Untitled.docx', sha256: cur?.sha256 ?? null, size: cur?.size ?? null }
}

function openExport() {
  if (!cur) return
  const c = cur
  openExportDialog($('export-dialog') as HTMLDialogElement, {
    name: c.name,
    hasOriginal: !!c.original,
    contentEdited: c.contentEdited,
    records: records(),
    info: documentInfo(),
    buildAnnotated,
    buildRebuilt,
    onExported: () => setDirty(false),
    notify: toast,
  })
}
$('btn-export').addEventListener('click', openExport)

async function writeTo(handle: FsHandle, blob: Blob) {
  const w = await handle.createWritable()
  await w.write(blob)
  await w.close()
}

/**
 * Save keeps the original file intact whenever it can: comments and highlights are merged into the original
 * package instead of rebuilding it. A rebuild (which drops features Darkwrote can't display) only happens
 * when the text itself was edited, and only after confirmation if there is an original to lose.
 */
async function save(forceAs = false) {
  if (!cur) return
  const c = cur
  const patchable = !!c.original && !c.contentEdited
  const canPick = !!fsw.showSaveFilePicker

  if (!canPick && isCompact()) return openExport() // Android: no way to write back to a file, so offer the export choices

  try {
    let blob: Blob
    let patchedBytes: ArrayBuffer | null = null
    let hasAnnotations = false
    if (patchable) {
      const mine = placeable(records())
      hasAnnotations = mine.length > 0
      if (mine.length) {
        const r = await buildAnnotated()
        blob = r.blob
        patchedBytes = await blob.arrayBuffer()
        if (r.skipped.length) toast(`${r.skipped.length} annotation(s) couldn’t be placed in the original`, true)
      } else {
        blob = new Blob([c.original!], { type: DOCX_MIME })
      }
    } else {
      if (c.original && !confirm('You changed the text, so Darkwrote has to rebuild this document. Anything it can’t display (headers, footnotes, text boxes…) will be lost in the saved file.\n\nContinue? (Cancel, then use Export to keep your annotations separately.)')) return
      blob = await buildRebuilt()
    }

    let target = c.handle
    if (canPick && (!target || forceAs)) {
      target = await fsw.showSaveFilePicker!({ suggestedName: baseName(c.name) + '.docx', types: pickerTypes })
      c.handle = target
      setName(target.name)
    }
    if (target) {
      await writeTo(target, blob)
      if (patchedBytes) {
        // the comments are now part of the file on disk
        c.original = patchedBytes
        commentsPanel.markAllImported()
        void store.saveOriginal(c.id, patchedBytes)
      }
      toast(`Saved ${target.name}`)
    } else {
      downloadBlob(blob, baseName(c.name) + (hasAnnotations ? '.annotated.docx' : '.docx'))
      toast('Downloaded a copy')
    }
    setDirty(false)
  } catch (e) {
    if ((e as Error).name !== 'AbortError') toast(`Save failed: ${(e as Error).message}`, true)
  }
}
$('btn-save').addEventListener('click', () => void save())
$('btn-saveas').addEventListener('click', () => void save(true))

document.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return
  const k = e.key.toLowerCase()
  if (k === 's') {
    e.preventDefault()
    void save(e.shiftKey)
  } else if (k === 'o') {
    e.preventDefault()
    void openFile()
  } else if (k === 'm' && e.altKey && cur) {
    e.preventDefault()
    commentsPanel.startDraft(currentRange())
  }
})

window.addEventListener('beforeunload', (e) => {
  if (pendingSave || saveFailed) e.preventDefault()
})

// ---------- drag & drop ----------
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
  if (file) void loadFile(file, null)
})

// ---------- app install / offline / shared files ----------
registerServiceWorker(() => toast('A new version is ready — reload to update', false, { label: 'Reload', run: () => location.reload() }))
setupInstallButton($<HTMLButtonElement>('btn-install'), toast)
onLaunchFiles((file, handle) => void loadFile(file, handle as FsHandle))

async function boot() {
  let last: string | null = null
  try {
    last = localStorage.getItem('darkwrote.last')
  } catch {
    /* ignore */
  }
  resetDocument(false)
  const params = new URLSearchParams(location.search)
  if (params.has('shared')) {
    history.replaceState(null, '', location.pathname)
    const shared = await takeSharedFile()
    if (shared) return void loadFile(shared, null)
  }
  // Pick up where the last session ended (also covers Android discarding the tab in the background).
  if (last && (await store.getDoc(last))) await restoreById(last)
}
void boot()

// Handy for automated checks and debugging from the browser console.
;(window as unknown as { darkwrote: unknown }).darkwrote = { editor, getCurrent: () => cur, flushSave, topPosition: () => topPosition(editor, workspace) }
