export interface VitePreviewCliOptions {
  readonly port: number
  readonly host: string
}

export function parseVitePreviewCliArgs(argv: readonly string[]): VitePreviewCliOptions {
  return {
    port: parseVitePreviewPort(argv[0]),
    host: argv[1] ?? '127.0.0.1',
  }
}

export function parseVitePreviewPort(value: string | undefined): number {
  if (value === undefined || !/^(?:[1-9]\d*)$/u.test(value)) {
    throw new Error('Expected a decimal preview port.')
  }
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed) || parsed > 65_535) {
    throw new Error('Expected a preview port between 1 and 65535.')
  }
  return parsed
}
