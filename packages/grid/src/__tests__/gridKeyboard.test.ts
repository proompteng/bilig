import { describe, expect, test } from 'vitest'
import {
  isClearCellKey,
  isClipboardShortcut,
  isCurrentRegionSelectionShortcut,
  isDeleteKey,
  isFillSelectionShortcut,
  isFillShortcut,
  isHandledGridKey,
  isNavigationKey,
  isNumericEditorSeed,
  isPrintableKey,
  isScrollActiveCellShortcut,
  isSheetSelectionShortcut,
  isStructuralDeleteShortcut,
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
    expect(isClipboardShortcut({ altKey: false, ctrlKey: true, key: 'v', metaKey: false, shiftKey: true })).toBe(true)
    expect(isFillShortcut({ altKey: false, ctrlKey: true, key: 'd', metaKey: false, shiftKey: false })).toBe(true)
    expect(isFillShortcut({ altKey: false, ctrlKey: false, key: 'r', metaKey: true, shiftKey: false })).toBe(true)
    expect(isFillShortcut({ altKey: false, ctrlKey: true, key: 'd', metaKey: false, shiftKey: true })).toBe(false)
    expect(isFillSelectionShortcut({ altKey: false, ctrlKey: true, key: 'Enter', metaKey: false, shiftKey: false })).toBe(true)
    expect(isFillSelectionShortcut({ altKey: false, ctrlKey: false, key: 'Enter', metaKey: true, shiftKey: false })).toBe(true)
    expect(isFillSelectionShortcut({ altKey: false, ctrlKey: true, key: 'Enter', metaKey: false, shiftKey: true })).toBe(false)
    expect(isScrollActiveCellShortcut({ altKey: false, ctrlKey: true, key: 'Backspace', metaKey: false })).toBe(true)
    expect(isScrollActiveCellShortcut({ altKey: false, ctrlKey: false, key: 'Backspace', metaKey: true })).toBe(true)
    expect(isScrollActiveCellShortcut({ altKey: false, ctrlKey: true, key: 'Delete', metaKey: false })).toBe(false)
    expect(isStructuralDeleteShortcut({ altKey: true, ctrlKey: true, key: '-', metaKey: false })).toBe(true)
    expect(isStructuralDeleteShortcut({ altKey: true, ctrlKey: false, key: '-', metaKey: true })).toBe(true)
    expect(isStructuralDeleteShortcut({ altKey: false, ctrlKey: true, key: '-', metaKey: false })).toBe(false)
    expect(isSheetSelectionShortcut({ altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: true })).toBe(true)
    expect(isSheetSelectionShortcut({ altKey: false, ctrlKey: true, key: ' ', metaKey: false, shiftKey: false })).toBe(true)
    expect(isSheetSelectionShortcut({ altKey: false, ctrlKey: true, key: ' ', metaKey: false, shiftKey: true })).toBe(true)
    expect(isSheetSelectionShortcut({ altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: false })).toBe(false)
    expect(isCurrentRegionSelectionShortcut({ altKey: false, ctrlKey: true, key: '*', metaKey: false, shiftKey: true })).toBe(true)
    expect(isCurrentRegionSelectionShortcut({ altKey: false, ctrlKey: false, key: '*', metaKey: true, shiftKey: true })).toBe(true)
    expect(isCurrentRegionSelectionShortcut({ altKey: false, ctrlKey: true, key: '8', metaKey: false, shiftKey: true })).toBe(false)
    expect(isHandledGridKey({ altKey: false, ctrlKey: false, key: 'F2', metaKey: false })).toBe(true)
    expect(isHandledGridKey({ altKey: false, ctrlKey: false, key: 'Escape', metaKey: false })).toBe(true)
    expect(isHandledGridKey({ altKey: false, ctrlKey: false, key: 'Delete', metaKey: false })).toBe(true)
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
        ctrlKey: true,
        key: '*',
        metaKey: false,
        shiftKey: true,
      }),
    ).toBe(true)
    expect(
      isHandledGridKey({
        altKey: false,
        ctrlKey: true,
        key: 'Enter',
        metaKey: false,
        shiftKey: false,
      }),
    ).toBe(true)
    expect(
      isHandledGridKey({
        altKey: true,
        ctrlKey: true,
        key: '-',
        metaKey: false,
        shiftKey: false,
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

  test('only treats unmodified delete keys as grid clear commands', () => {
    expect(isDeleteKey('Delete')).toBe(true)
    expect(isDeleteKey('Backspace')).toBe(true)
    expect(isDeleteKey('Escape')).toBe(false)
    expect(isClearCellKey({ altKey: false, ctrlKey: false, key: 'Delete', metaKey: false })).toBe(true)
    expect(isClearCellKey({ altKey: false, ctrlKey: false, key: 'Backspace', metaKey: false })).toBe(true)
    expect(isClearCellKey({ altKey: true, ctrlKey: false, key: 'Delete', metaKey: false })).toBe(false)
    expect(isClearCellKey({ altKey: false, ctrlKey: true, key: 'Backspace', metaKey: false })).toBe(false)
    expect(isClearCellKey({ altKey: false, ctrlKey: false, key: 'Backspace', metaKey: true })).toBe(false)
    expect(isClearCellKey({ altKey: false, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: true })).toBe(false)
    expect(isHandledGridKey({ altKey: false, ctrlKey: false, key: 'Delete', metaKey: true })).toBe(false)
    expect(isHandledGridKey({ altKey: false, ctrlKey: false, key: 'Backspace', metaKey: true })).toBe(true)
  })

  test('detects numeric editor seeds', () => {
    expect(isNumericEditorSeed('123')).toBe(true)
    expect(isNumericEditorSeed('-12.5')).toBe(true)
    expect(isNumericEditorSeed('=A1')).toBe(false)
    expect(isNumericEditorSeed(' hello ')).toBe(false)
  })
})
