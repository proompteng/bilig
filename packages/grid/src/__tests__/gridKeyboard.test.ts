import { describe, expect, test } from 'vitest'
import {
  isClipboardShortcut,
  isHandledGridKey,
  isNavigationKey,
  isNumericEditorSeed,
  isPrintableKey,
  normalizeKeyboardKey,
} from '../gridKeyboard.js'

describe('gridKeyboard', () => {
  test('normalizes numpad keys into printable characters', () => {
    expect(normalizeKeyboardKey('1', 'Numpad1')).toBe('1')
    expect(normalizeKeyboardKey('.', 'NumpadDecimal')).toBe('.')
    expect(normalizeKeyboardKey('+', 'NumpadAdd')).toBe('+')
    expect(normalizeKeyboardKey('/', 'NumpadDivide')).toBe('/')
    expect(normalizeKeyboardKey('a', 'KeyA')).toBe('a')
  })

  test('classifies printable, navigation, clipboard, and handled keys', () => {
    expect(isPrintableKey({ altKey: false, ctrlKey: false, key: 'x', metaKey: false })).toBe(true)
    expect(isPrintableKey({ altKey: false, ctrlKey: true, key: 'x', metaKey: false })).toBe(false)
    expect(isNavigationKey('ArrowDown')).toBe(true)
    expect(isNavigationKey('Enter')).toBe(false)
    expect(isClipboardShortcut({ altKey: false, ctrlKey: true, key: 'c', metaKey: false })).toBe(true)
    expect(isHandledGridKey({ altKey: false, ctrlKey: false, key: 'F2', metaKey: false })).toBe(true)
    expect(isHandledGridKey({ altKey: false, ctrlKey: false, key: 'Escape', metaKey: false })).toBe(true)
    expect(
      isHandledGridKey({
        altKey: false,
        ctrlKey: true,
        key: 'a',
        metaKey: false,
        shiftKey: false,
      }),
    ).toBe(true)
    expect(
      isHandledGridKey({
        altKey: false,
        ctrlKey: false,
        key: ' ',
        metaKey: false,
        shiftKey: true,
      }),
    ).toBe(true)
    expect(
      isHandledGridKey({
        altKey: false,
        ctrlKey: true,
        key: ' ',
        metaKey: false,
        shiftKey: true,
      }),
    ).toBe(true)
    expect(
      isHandledGridKey({
        altKey: false,
        ctrlKey: false,
        key: 'Home',
        metaKey: false,
        shiftKey: false,
      }),
    ).toBe(true)
  })

  test('detects numeric editor seeds', () => {
    expect(isNumericEditorSeed('123')).toBe(true)
    expect(isNumericEditorSeed('-12.5')).toBe(true)
    expect(isNumericEditorSeed('=A1')).toBe(false)
    expect(isNumericEditorSeed(' hello ')).toBe(false)
  })
})
