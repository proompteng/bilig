export function collectLargeSimpleImportGarbage(): void {
  const bun = Reflect.get(globalThis, 'Bun')
  if (typeof bun === 'object' && bun !== null) {
    const bunGc = Reflect.get(bun, 'gc')
    if (typeof bunGc === 'function') {
      bunGc(true)
      return
    }
  }
  const gc = Reflect.get(globalThis, 'gc')
  if (typeof gc === 'function') {
    gc()
  }
}
