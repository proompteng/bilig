import { describe, expect, it } from 'vitest'

import { parseVitePreviewCliArgs, parseVitePreviewPort } from '../vite-preview-cli.ts'

describe('vite preview CLI parser', () => {
  it('parses preview port and host without starting Vite', () => {
    expect(parseVitePreviewCliArgs(['4173', '0.0.0.0'])).toEqual({
      port: 4173,
      host: '0.0.0.0',
    })
  })

  it('defaults the preview host to loopback', () => {
    expect(parseVitePreviewCliArgs(['4173'])).toEqual({
      port: 4173,
      host: '127.0.0.1',
    })
  })

  it('rejects malformed preview ports instead of truncating them', () => {
    expect(() => parseVitePreviewPort('4173abc')).toThrow('Expected a decimal preview port.')
    expect(() => parseVitePreviewPort('0')).toThrow('Expected a decimal preview port.')
    expect(() => parseVitePreviewPort('70000')).toThrow('Expected a preview port between 1 and 65535.')
  })
})
