import type { JSONContent } from '@tiptap/core'
import type { CommentData, DocMeta } from './docx/import'

/**
 * Local persistence (IndexedDB): the working copy of each document (notes, highlights, edits), the original
 * file bytes (needed for annotated export), and the reading position. Everything stays on the device.
 */

export interface StoredDoc {
  id: string
  name: string
  updatedAt: number
  doc: JSONContent
  meta: DocMeta
  comments: CommentData[]
  sha256: string | null
  size: number | null
  /** True once the text has been changed beyond comments/highlights, so the original can't be patched */
  contentEdited: boolean
  /** Number of the reader's comments + highlights, for the recent-documents list */
  annotationCount: number
}

export interface ViewState {
  id: string
  /** Editor position of the passage at the top of the screen */
  pos: number
  /** 0..1 through the document */
  progress: number
  updatedAt: number
}

const DB_NAME = 'darkwrote'
const STORES = ['docs', 'originals', 'view'] as const
type Store = (typeof STORES)[number]

let dbPromise: Promise<IDBDatabase> | null = null

function open(): Promise<IDBDatabase> {
  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      if (typeof indexedDB === 'undefined') return reject(new Error('IndexedDB unavailable'))
      const req = indexedDB.open(DB_NAME, 1)
      req.onupgradeneeded = () => {
        for (const s of STORES) if (!req.result.objectStoreNames.contains(s)) req.result.createObjectStore(s, { keyPath: 'id' })
      }
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error)
    })
    dbPromise.catch(() => (dbPromise = null))
  }
  return dbPromise
}

function tx<T>(store: Store, mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T> | void): Promise<T | undefined> {
  return open().then(
    (db) =>
      new Promise<T | undefined>((resolve, reject) => {
        const t = db.transaction(store, mode)
        const req = fn(t.objectStore(store))
        t.oncomplete = () => resolve(req ? req.result : undefined)
        t.onerror = () => reject(t.error)
        t.onabort = () => reject(t.error)
      }),
  )
}

export const store = {
  async saveDoc(d: StoredDoc): Promise<boolean> {
    try {
      await tx('docs', 'readwrite', (s) => s.put(d))
      return true
    } catch {
      return false
    }
  },
  async getDoc(id: string): Promise<StoredDoc | undefined> {
    try {
      return (await tx<StoredDoc>('docs', 'readonly', (s) => s.get(id))) ?? undefined
    } catch {
      return undefined
    }
  },
  async listDocs(): Promise<(Omit<StoredDoc, 'doc'> & { progress: number })[]> {
    try {
      const docs = (await tx<StoredDoc[]>('docs', 'readonly', (s) => s.getAll())) ?? []
      const views = (await tx<ViewState[]>('view', 'readonly', (s) => s.getAll())) ?? []
      const progress = new Map(views.map((v) => [v.id, v.progress]))
      return docs
        .map(({ doc: _doc, ...rest }) => ({ ...rest, progress: progress.get(rest.id) ?? 0 }))
        .sort((a, b) => b.updatedAt - a.updatedAt)
    } catch {
      return []
    }
  },
  async deleteDoc(id: string) {
    for (const s of STORES) await tx(s, 'readwrite', (st) => st.delete(id)).catch(() => undefined)
  },
  async saveOriginal(id: string, bytes: ArrayBuffer) {
    await tx('originals', 'readwrite', (s) => s.put({ id, bytes })).catch(() => undefined)
  },
  async hasOriginal(id: string): Promise<boolean> {
    const r = await tx<IDBValidKey | undefined>('originals', 'readonly', (s) => s.getKey(id)).catch(() => undefined)
    return r !== undefined
  },
  async getOriginal(id: string): Promise<ArrayBuffer | undefined> {
    const r = await tx<{ id: string; bytes: ArrayBuffer }>('originals', 'readonly', (s) => s.get(id)).catch(() => undefined)
    return r?.bytes
  },
  async saveView(v: ViewState) {
    await tx('view', 'readwrite', (s) => s.put(v)).catch(() => undefined)
  },
  async getView(id: string): Promise<ViewState | undefined> {
    return (await tx<ViewState>('view', 'readonly', (s) => s.get(id)).catch(() => undefined)) ?? undefined
  },
}

/** Ask the browser not to evict our data under storage pressure (best effort). */
export async function requestPersistence(): Promise<boolean> {
  try {
    return (await navigator.storage?.persist?.()) ?? false
  } catch {
    return false
  }
}

export async function sha256Hex(bytes: ArrayBuffer): Promise<string | null> {
  try {
    const digest = await crypto.subtle.digest('SHA-256', bytes)
    return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, '0')).join('')
  } catch {
    return null
  }
}
