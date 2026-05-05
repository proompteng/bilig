const LOG_PREFIX = '[bilig]'

export function logError(...args: readonly unknown[]): void {
  process.stderr.write(`${LOG_PREFIX} ${args.map((arg) => (arg instanceof Error ? (arg.stack ?? arg.message) : String(arg))).join(' ')}\n`)
}
