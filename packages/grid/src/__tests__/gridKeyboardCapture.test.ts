import { describe, expect, it, vi } from 'vitest'
import { handleWorkbookGridKeyDownCapture } from '../gridKeyboardCapture.js'

describe('gridKeyboardCapture', () => {
  it('should open the header context menu for the context-menu shortcut', () => {
    // Arrange
    const event = createKeyboardEvent({ key: 'ContextMenu', code: 'ContextMenu' })
    const handleGridKey = vi.fn()
    const openHeaderContextMenuFromKeyboard = vi.fn(() => true)
    const resetPointerInteraction = vi.fn()

    // Act
    handleWorkbookGridKeyDownCapture({
      event,
      handleGridKey,
      openHeaderContextMenuFromKeyboard,
      resetPointerInteraction,
    })

    // Assert
    expect(resetPointerInteraction).toHaveBeenCalledTimes(1)
    expect(openHeaderContextMenuFromKeyboard).toHaveBeenCalledTimes(1)
    expect(handleGridKey).not.toHaveBeenCalled()
    expect(event.preventDefault).toHaveBeenCalledTimes(1)
    expect(event.stopPropagation).toHaveBeenCalledTimes(1)
  })

  it('should ignore unhandled keys after resetting pointer state', () => {
    // Arrange
    const event = createKeyboardEvent({ key: 'F13', code: 'F13' })
    const handleGridKey = vi.fn()
    const openHeaderContextMenuFromKeyboard = vi.fn(() => false)
    const resetPointerInteraction = vi.fn()

    // Act
    handleWorkbookGridKeyDownCapture({
      event,
      handleGridKey,
      openHeaderContextMenuFromKeyboard,
      resetPointerInteraction,
    })

    // Assert
    expect(resetPointerInteraction).toHaveBeenCalledTimes(1)
    expect(openHeaderContextMenuFromKeyboard).not.toHaveBeenCalled()
    expect(handleGridKey).not.toHaveBeenCalled()
    expect(event.preventDefault).not.toHaveBeenCalled()
    expect(event.stopPropagation).not.toHaveBeenCalled()
  })

  it('should forward handled keys and stop propagation when default is prevented', () => {
    // Arrange
    const event = createKeyboardEvent({ key: 'Backspace', code: 'Backspace' })
    const handleGridKey = vi.fn(({ preventDefault }) => {
      preventDefault()
    })
    const openHeaderContextMenuFromKeyboard = vi.fn(() => false)
    const resetPointerInteraction = vi.fn()

    // Act
    handleWorkbookGridKeyDownCapture({
      event,
      handleGridKey,
      openHeaderContextMenuFromKeyboard,
      resetPointerInteraction,
    })

    // Assert
    expect(handleGridKey).toHaveBeenCalledTimes(1)
    expect(event.preventDefault).toHaveBeenCalledTimes(1)
    expect(event.stopPropagation).toHaveBeenCalledTimes(1)
  })
})

// Helpers

function createKeyboardEvent(input: { key: string; code: string; shiftKey?: boolean | undefined }): {
  altKey: boolean
  code: string
  ctrlKey: boolean
  defaultPrevented: boolean
  key: string
  metaKey: boolean
  preventDefault: ReturnType<typeof vi.fn>
  shiftKey: boolean
  stopPropagation: ReturnType<typeof vi.fn>
} {
  const { code, key, shiftKey = false } = input
  const event = {
    altKey: false,
    code,
    ctrlKey: false,
    defaultPrevented: false,
    key,
    metaKey: false,
    shiftKey,
    preventDefault: vi.fn(() => {
      event.defaultPrevented = true
    }),
    stopPropagation: vi.fn(),
  }
  return event
}
