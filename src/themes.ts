export type ThemeMode = 'dark' | 'light'

export interface ThemeColors {
  /** App background behind the page */
  bg: string
  /** Toolbars and panels */
  surface: string
  /** The page itself */
  paper: string
  /** Body text */
  text: string
  /** Secondary text and icons */
  muted: string
  /** Borders and dividers */
  border: string
  /** Buttons, active states, caret */
  accent: string
  /** Text selection background */
  selection: string
  /** Hyperlinks */
  link: string
}

export interface Theme {
  id: string
  name: string
  mode: ThemeMode
  colors: ThemeColors
  builtin?: boolean
}

export const COLOR_LABELS: Record<keyof ThemeColors, string> = {
  bg: 'Background',
  surface: 'Toolbars & panels',
  paper: 'Page',
  text: 'Text',
  muted: 'Muted text',
  border: 'Borders',
  accent: 'Accent',
  selection: 'Selection',
  link: 'Links',
}

export const BUILTIN_THEMES: Theme[] = [
  {
    id: 'midnight', name: 'Midnight', mode: 'dark', builtin: true,
    colors: { bg: '#0d1117', surface: '#161b22', paper: '#1c2128', text: '#d6dde6', muted: '#8b96a5', border: '#30363d', accent: '#58a6ff', selection: '#264f78', link: '#79c0ff' },
  },
  {
    id: 'charcoal', name: 'Charcoal', mode: 'dark', builtin: true,
    colors: { bg: '#1a1a1a', surface: '#242424', paper: '#2b2b2b', text: '#e0e0e0', muted: '#9a9a9a', border: '#3d3d3d', accent: '#e5a45c', selection: '#5a4426', link: '#f0b878' },
  },
  {
    id: 'amoled', name: 'AMOLED Black', mode: 'dark', builtin: true,
    colors: { bg: '#000000', surface: '#0a0a0a', paper: '#000000', text: '#e8e8e8', muted: '#808080', border: '#262626', accent: '#00d4aa', selection: '#0b4a3f', link: '#4de3c3' },
  },
  {
    id: 'dracula', name: 'Dracula', mode: 'dark', builtin: true,
    colors: { bg: '#21222c', surface: '#282a36', paper: '#2d2f3d', text: '#f8f8f2', muted: '#8e93b5', border: '#44475a', accent: '#bd93f9', selection: '#44475a', link: '#8be9fd' },
  },
  {
    id: 'nord', name: 'Nord', mode: 'dark', builtin: true,
    colors: { bg: '#242933', surface: '#2e3440', paper: '#3b4252', text: '#eceff4', muted: '#9aa5b8', border: '#4c566a', accent: '#88c0d0', selection: '#506277', link: '#81a1c1' },
  },
  {
    id: 'solarized-dark', name: 'Solarized Dark', mode: 'dark', builtin: true,
    colors: { bg: '#00212b', surface: '#002b36', paper: '#073642', text: '#93a1a1', muted: '#657b83', border: '#0f4a58', accent: '#b58900', selection: '#124a58', link: '#268bd2' },
  },
  {
    id: 'paper', name: 'Paper', mode: 'light', builtin: true,
    colors: { bg: '#e9ecef', surface: '#f8f9fa', paper: '#ffffff', text: '#1f2328', muted: '#6a737d', border: '#d0d7de', accent: '#0969da', selection: '#b6d6fb', link: '#0969da' },
  },
  {
    id: 'sepia', name: 'Sepia', mode: 'light', builtin: true,
    colors: { bg: '#e6dcc6', surface: '#efe6d2', paper: '#f8f1e3', text: '#433422', muted: '#8a7658', border: '#d4c6a8', accent: '#a0522d', selection: '#e2cd9f', link: '#8b4513' },
  },
  {
    id: 'solarized-light', name: 'Solarized Light', mode: 'light', builtin: true,
    colors: { bg: '#eee8d5', surface: '#f5efdc', paper: '#fdf6e3', text: '#586e75', muted: '#93a1a1', border: '#ddd6c1', accent: '#268bd2', selection: '#d6dfd2', link: '#268bd2' },
  },
  {
    id: 'high-contrast', name: 'High Contrast', mode: 'light', builtin: true,
    colors: { bg: '#d8d8d8', surface: '#ffffff', paper: '#ffffff', text: '#000000', muted: '#3a3a3a', border: '#000000', accent: '#0000ee', selection: '#ffe58a', link: '#0000ee' },
  },
]

const CUSTOM_KEY = 'darkwrote.customThemes'
const ACTIVE_KEY = 'darkwrote.activeTheme'
const ADAPT_KEY = 'darkwrote.adaptColors'

function readStorage(key: string): string | null {
  try {
    return localStorage.getItem(key)
  } catch {
    return null
  }
}

function writeStorage(key: string, value: string) {
  try {
    localStorage.setItem(key, value)
  } catch {
    /* storage unavailable (private mode etc.) */
  }
}

const HEX = /^#[0-9a-fA-F]{6}$/

export function isValidTheme(t: unknown): t is Theme {
  if (!t || typeof t !== 'object') return false
  const o = t as Theme
  if (typeof o.name !== 'string' || (o.mode !== 'dark' && o.mode !== 'light')) return false
  return (Object.keys(COLOR_LABELS) as (keyof ThemeColors)[]).every((k) => HEX.test(o.colors?.[k] ?? ''))
}

export function loadCustomThemes(): Theme[] {
  try {
    const parsed = JSON.parse(readStorage(CUSTOM_KEY) ?? '[]')
    return Array.isArray(parsed) ? parsed.filter(isValidTheme).map((t) => ({ ...t, builtin: false })) : []
  } catch {
    return []
  }
}

export function saveCustomThemes(themes: Theme[]) {
  writeStorage(CUSTOM_KEY, JSON.stringify(themes.filter((t) => !t.builtin)))
}

export function allThemes(): Theme[] {
  return [...BUILTIN_THEMES, ...loadCustomThemes()]
}

export function getActiveThemeId(): string {
  return readStorage(ACTIVE_KEY) ?? 'midnight'
}

export function getAdaptColors(): boolean {
  return readStorage(ADAPT_KEY) !== 'false'
}

export function setAdaptColors(on: boolean) {
  writeStorage(ADAPT_KEY, String(on))
  document.documentElement.dataset.adapt = String(on)
}

export function newThemeId(): string {
  return 'custom-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
}

export function applyTheme(theme: Theme) {
  const root = document.documentElement
  for (const [k, v] of Object.entries(theme.colors)) root.style.setProperty(`--${k}`, v)
  root.dataset.mode = theme.mode
  root.style.colorScheme = theme.mode
  document.querySelector('meta[name="theme-color"]')?.setAttribute('content', theme.colors.surface)
  writeStorage(ACTIVE_KEY, theme.id)
}

export function themeToJson(theme: Theme): string {
  const { name, mode, colors } = theme
  return JSON.stringify({ name, mode, colors }, null, 2)
}
