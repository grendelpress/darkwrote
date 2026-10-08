import type { Editor } from '@tiptap/core'

export interface ReadPrefs {
  /** Text size multiplier, 0.7 – 2.2 */
  scale: number
  font: 'original' | 'serif' | 'sans'
  /** Line height */
  line: number
}

const KEY = 'darkwrote.read'
const DEFAULTS: ReadPrefs = { scale: 1, font: 'original', line: 1.55 }

export function loadReadPrefs(): ReadPrefs {
  try {
    const p = JSON.parse(localStorage.getItem(KEY) ?? '{}')
    return {
      scale: Math.min(2.2, Math.max(0.7, Number(p.scale) || DEFAULTS.scale)),
      font: ['original', 'serif', 'sans'].includes(p.font) ? p.font : DEFAULTS.font,
      line: [1.35, 1.55, 1.8].includes(Number(p.line)) ? Number(p.line) : DEFAULTS.line,
    }
  } catch {
    return { ...DEFAULTS }
  }
}

export function applyReadPrefs(p: ReadPrefs) {
  document.documentElement.style.setProperty('--read-scale', String(p.scale))
  document.documentElement.style.setProperty('--read-line', String(p.line))
  document.body.dataset.readFont = p.font
  try {
    localStorage.setItem(KEY, JSON.stringify(p))
  } catch {
    /* ignore */
  }
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, props: Record<string, any> = {}, ...kids: (Node | string)[]): HTMLElementTagNameMap[K] {
  const n = document.createElement(tag)
  Object.assign(n, props)
  n.append(...kids)
  return n
}

/** The "Aa" panel: text size, typeface and line spacing. `keepPlace` wraps changes so the reader stays at the same passage. */
export function buildReadSettings(host: HTMLElement, prefs: ReadPrefs, keepPlace: (change: () => void) => void) {
  const apply = () => keepPlace(() => applyReadPrefs(prefs))

  const readout = el('span', { className: 'readout' }, `${Math.round(prefs.scale * 100)}%`)
  const slider = el('input', { type: 'range', min: '70', max: '220', step: '5', value: String(Math.round(prefs.scale * 100)), ariaLabel: 'Text size' })
  const setScale = (pct: number) => {
    prefs.scale = Math.min(2.2, Math.max(0.7, pct / 100))
    slider.value = String(Math.round(prefs.scale * 100))
    readout.textContent = `${Math.round(prefs.scale * 100)}%`
    apply()
  }
  slider.addEventListener('input', () => setScale(Number(slider.value)))
  const smaller = el('button', { className: 'btn', type: 'button', textContent: 'A−', ariaLabel: 'Smaller text' })
  const larger = el('button', { className: 'btn', type: 'button', textContent: 'A+', ariaLabel: 'Larger text' })
  smaller.addEventListener('click', () => setScale(Math.round(prefs.scale * 100) - 10))
  larger.addEventListener('click', () => setScale(Math.round(prefs.scale * 100) + 10))

  const seg = <T extends string | number>(options: [T, string][], current: () => T, set: (v: T) => void) => {
    const wrap = el('div', { className: 'seg' })
    const paint = () => wrap.querySelectorAll('button').forEach((b, i) => b.classList.toggle('on', options[i][0] === current()))
    for (const [value, label] of options) {
      const b = el('button', { type: 'button', textContent: label })
      b.addEventListener('click', () => {
        set(value)
        paint()
        apply()
      })
      wrap.append(b)
    }
    paint()
    return wrap
  }

  host.replaceChildren(
    el('h3', {}, 'Reading'),
    el('div', { className: 'row' }, smaller, slider, larger, readout),
    el('div', { className: 'row' }, el('span', { className: 'label' }, 'Font'), seg<ReadPrefs['font']>([['original', 'Original'], ['serif', 'Serif'], ['sans', 'Sans']], () => prefs.font, (v) => (prefs.font = v))),
    el('div', { className: 'row' }, el('span', { className: 'label' }, 'Spacing'), seg<number>([[1.35, 'Tight'], [1.55, 'Normal'], [1.8, 'Roomy']], () => prefs.line, (v) => (prefs.line = v))),
  )
}

// ---------- reading position ----------

/** Editor position of the passage at the top of the visible area. */
export function topPosition(editor: Editor, scroller: HTMLElement): number | null {
  const box = scroller.getBoundingClientRect()
  const dom = editor.view.dom.getBoundingClientRect()
  const left = Math.min(Math.max(dom.left + Math.min(24, dom.width / 2), 0), window.innerWidth - 1)
  for (const dy of [24, 60, 110]) {
    const hit = editor.view.posAtCoords({ left, top: box.top + dy })
    if (hit) return hit.pos
  }
  return null
}

export function scrollToPos(editor: Editor, scroller: HTMLElement, pos: number) {
  const size = editor.state.doc.content.size
  const safe = Math.min(Math.max(pos, 0), size)
  const { node } = editor.view.domAtPos(safe)
  const target = node instanceof Element ? node : node.parentElement
  if (!target) return
  const top = target.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop - 12
  scroller.scrollTo({ top: Math.max(0, top), behavior: 'auto' })
}
