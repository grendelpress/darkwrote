/** Installable-app plumbing: offline service worker, "Install" prompt, files shared from other Android apps. */

interface InstallEvent extends Event {
  prompt(): Promise<void>
  userChoice: Promise<{ outcome: 'accepted' | 'dismissed' }>
}

export function registerServiceWorker(onUpdated: () => void) {
  if (!('serviceWorker' in navigator) || !import.meta.env.PROD) return
  window.addEventListener('load', () => {
    navigator.serviceWorker
      .register('./sw.js')
      .then((reg) => {
        // look for a new version whenever the app is brought back to the foreground
        document.addEventListener('visibilitychange', () => {
          if (document.visibilityState === 'visible') void reg.update().catch(() => undefined)
        })
        reg.addEventListener('updatefound', () => {
          const incoming = reg.installing
          incoming?.addEventListener('statechange', () => {
            if (incoming.state === 'installed' && navigator.serviceWorker.controller) onUpdated()
          })
        })
      })
      .catch(() => undefined)
  })
}

export function setupInstallButton(button: HTMLButtonElement, notify: (msg: string) => void) {
  let deferred: InstallEvent | null = null
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault()
    deferred = e as InstallEvent
    button.hidden = false
  })
  window.addEventListener('appinstalled', () => {
    deferred = null
    button.hidden = true
    notify('Darkwrote is installed')
  })
  button.addEventListener('click', async () => {
    if (!deferred) return
    await deferred.prompt()
    await deferred.userChoice
    deferred = null
    button.hidden = true
  })
}

/** A .docx handed over by Android's share sheet is parked by the service worker; fetch and clear it. */
export async function takeSharedFile(): Promise<File | null> {
  if (!('caches' in window)) return null
  try {
    const cache = await caches.open('darkwrote-share')
    const key = new URL('./shared/latest', location.href).href
    const res = await cache.match(key)
    if (!res) return null
    const name = decodeURIComponent(res.headers.get('X-File-Name') ?? 'shared.docx')
    const blob = await res.blob()
    await cache.delete(key)
    return new File([blob], name, { type: blob.type })
  } catch {
    return null
  }
}

interface LaunchParams {
  files: { getFile(): Promise<File> }[]
}

/** Desktop "Open with Darkwrote" for an installed app. */
export function onLaunchFiles(handler: (file: File, handle: unknown) => void) {
  const q = (window as unknown as { launchQueue?: { setConsumer(fn: (p: LaunchParams) => void): void } }).launchQueue
  q?.setConsumer(async (params) => {
    const first = params.files?.[0]
    if (first) handler(await first.getFile(), first)
  })
}
