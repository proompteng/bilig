export function downloadModelFile(title: string, extension: string, content: string): void {
  const url = URL.createObjectURL(new Blob([content], { type: 'application/json' }))
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = `${title.replace(/[^a-zA-Z0-9_-]+/gu, '-').slice(0, 80) || 'model'}.${extension}`
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
