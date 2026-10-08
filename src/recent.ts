import { store } from './storage'

function when(ts: number): string {
  const mins = Math.round((Date.now() - ts) / 60000)
  if (mins < 1) return 'just now'
  if (mins < 60) return `${mins} min ago`
  const hrs = Math.round(mins / 60)
  if (hrs < 24) return `${hrs} h ago`
  return new Date(ts).toLocaleDateString([], { month: 'short', day: 'numeric' })
}

/** The "Recent documents" list on the start screen: reopen where you left off, or forget a document. */
export async function renderRecent(host: HTMLElement, open: (id: string) => void) {
  const docs = await store.listDocs()
  host.hidden = docs.length === 0
  if (!docs.length) return
  const title = document.createElement('h2')
  title.textContent = 'Continue where you left off'
  host.replaceChildren(title)
  for (const d of docs.slice(0, 8)) {
    const row = document.createElement('div')
    row.className = 'recent-item'
    row.setAttribute('role', 'button')
    row.tabIndex = 0
    const info = document.createElement('div')
    info.className = 'name'
    const name = document.createElement('div')
    name.textContent = d.name
    name.style.cssText = 'overflow:hidden;text-overflow:ellipsis'
    const small = document.createElement('small')
    const notes = d.annotationCount ? `${d.annotationCount} note${d.annotationCount === 1 ? '' : 's'} · ` : ''
    small.textContent = `${notes}${Math.round(d.progress * 100)}% read · ${when(d.updatedAt)}`
    const bar = document.createElement('div')
    bar.className = 'bar'
    bar.innerHTML = `<i style="width:${Math.round(d.progress * 100)}%"></i>`
    info.append(name, small, bar)
    const x = document.createElement('button')
    x.className = 'x link-btn'
    x.type = 'button'
    x.textContent = '✕'
    x.title = 'Remove from this device'
    x.setAttribute('aria-label', `Remove ${d.name} from this device`)
    x.addEventListener('click', async (e) => {
      e.stopPropagation()
      if (!confirm(`Remove “${d.name}” and its notes from this device?\n\nThe original file is not affected.`)) return
      await store.deleteDoc(d.id)
      void renderRecent(host, open)
    })
    const go = () => open(d.id)
    row.addEventListener('click', go)
    row.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault()
        go()
      }
    })
    row.append(info, x)
    host.append(row)
  }
}
