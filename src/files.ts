export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

export function safeName(name: string): string {
  return name.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '_').trim() || 'document'
}

export function baseName(name: string): string {
  return safeName(name.replace(/\.docx$/i, ''))
}

/** Save through a normal browser download (the reliable route on Android, where Chrome puts it in Downloads). */
export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = Object.assign(document.createElement('a'), { href: url, download: filename })
  document.body.append(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 30_000)
}

export function canShareFile(blob: Blob, filename: string): boolean {
  try {
    const file = new File([blob], filename, { type: blob.type })
    return !!navigator.canShare?.({ files: [file] })
  } catch {
    return false
  }
}

/** Hand the file to Android's share sheet (Drive, Files, email, …). Returns false if the user dismissed it. */
export async function shareBlob(blob: Blob, filename: string): Promise<boolean> {
  const file = new File([blob], filename, { type: blob.type })
  try {
    await navigator.share({ files: [file], title: filename })
    return true
  } catch (e) {
    if ((e as Error).name === 'AbortError') return false
    throw e
  }
}
