import { execSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative } from 'node:path'
import { defineConfig, type Plugin } from 'vitest/config'

function listFiles(dir: string, root = dir): string[] {
  return readdirSync(dir).flatMap((name) => {
    const full = join(dir, name)
    return statSync(full).isDirectory() ? listFiles(full, root) : [relative(root, full).split('\\').join('/')]
  })
}

/** Emits sw.js with the list of built files to precache, so the app opens offline. */
function offlineServiceWorker(): Plugin {
  return {
    name: 'darkwrote-service-worker',
    apply: 'build',
    generateBundle(_options, bundle) {
      const built = Object.keys(bundle)
      const fromPublic = listFiles('public')
      const urls = ['./', ...built, ...fromPublic].map((f) => (f === './' ? f : './' + f))
      const html = bundle['index.html']
      const fingerprint = createHash('sha1')
        .update(built.join('|'))
        .update(html && 'source' in html ? String(html.source) : '')
        .update(fromPublic.map((f) => readFileSync(join('public', f)).length).join(','))
        .digest('hex')
        .slice(0, 12)
      const source = readFileSync('sw/sw.template.js', 'utf8')
        .replace("'__VERSION__'", JSON.stringify(fingerprint))
        .replace('__PRECACHE__', JSON.stringify(urls, null, 2))
      this.emitFile({ type: 'asset', fileName: 'sw.js', source })
    },
  }
}

function buildId(): string {
  try {
    return execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim()
  } catch {
    return new Date().toISOString().slice(0, 16)
  }
}

export default defineConfig({
  base: './',
  define: { __BUILD_ID__: JSON.stringify(buildId()) },
  plugins: [offlineServiceWorker()],
  test: { environment: 'jsdom' },
})
