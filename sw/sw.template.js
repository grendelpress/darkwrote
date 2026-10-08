/* Darkwrote service worker: offline app shell + Android "Share to Darkwrote" target.
   The version string and the precache list below are filled in at build time (see vite.config.ts). */
const VERSION = '__VERSION__'
const CACHE = 'darkwrote-app-' + VERSION
const SHARE_CACHE = 'darkwrote-share'
const PRECACHE = __PRECACHE__

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(PRECACHE)).then(() => self.skipWaiting()))
})

self.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      for (const key of await caches.keys()) {
        if (key.startsWith('darkwrote-app-') && key !== CACHE) await caches.delete(key)
      }
      await self.clients.claim()
    })(),
  )
})

self.addEventListener('fetch', (event) => {
  const req = event.request
  const url = new URL(req.url)
  if (req.method === 'POST' && url.pathname.endsWith('/share-target')) {
    event.respondWith(receiveShare(req))
    return
  }
  if (req.method !== 'GET' || url.origin !== self.location.origin) return
  event.respondWith(fromCache(req))
})

async function fromCache(req) {
  const cache = await caches.open(CACHE)
  const hit = await cache.match(req, { ignoreSearch: req.mode === 'navigate' })
  if (hit) return hit
  try {
    const res = await fetch(req)
    if (res.ok) cache.put(req, res.clone())
    return res
  } catch (err) {
    if (req.mode === 'navigate') {
      const shell = await cache.match(new URL('./', self.registration.scope).href)
      if (shell) return shell
    }
    throw err
  }
}

/** Another app shared a .docx to us: park it in a cache, then open the app, which picks it up. */
async function receiveShare(req) {
  try {
    const form = await req.formData()
    const file = form.getAll('file').find((f) => typeof f !== 'string')
    if (file) {
      const cache = await caches.open(SHARE_CACHE)
      await cache.put(
        new URL('shared/latest', self.registration.scope).href,
        new Response(file, { headers: { 'Content-Type': file.type || 'application/octet-stream', 'X-File-Name': encodeURIComponent(file.name || 'shared.docx') } }),
      )
    }
  } catch (e) {
    /* fall through: just open the app */
  }
  return Response.redirect(new URL('./?shared=1', self.registration.scope).href, 303)
}
