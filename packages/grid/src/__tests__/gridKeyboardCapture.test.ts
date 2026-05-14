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

  it('should ignore unhandled keys without resetting pointer state', () => {
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
    expect(resetPointerInteraction).not.toHaveBeenCalled()
    expect(openHeaderContextMenuFromKeyboard).not.toHaveBeenCalled()
    expect(handleGridKey).not.toHaveBeenCalled()
    expect(event.preventDefault).not.toHaveBeenCalled()
    expect(event.stopPropagation).not.toHaveBeenCalled()
  })

  it('should ignore already-prevented keydown events', () => {
    // Arrange
    const event = createKeyboardEvent({ key: 'Backspace', code: 'Backspace' })
    event.preventDefault()
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
    expect(resetPointerInteraction).not.toHaveBeenCalled()
    expect(openHeaderContextMenuFromKeyboard).not.toHaveBeenCalled()
    expect(handleGridKey).not.toHaveBeenCalled()
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

  it('should not claim modified delete key combinations or mutate pointer state', () => {
    // Arrange
    const handleGridKey = vi.fn()
    const openHeaderContextMenuFromKeyboard = vi.fn(() => false)
    const resetPointerInteraction = vi.fn()
    const events = [
      createKeyboardEvent({ key: 'Backspace', code: 'Backspace', metaKey: true }),
      createKeyboardEvent({ key: 'Delete', code: 'Delete', ctrlKey: true }),
      createKeyboardEvent({ key: 'Backspace', code: 'Backspace', altKey: true }),
      createKeyboardEvent({ key: 'Delete', code: 'Delete', shiftKey: true }),
    ]

    // Act
    for (const event of events) {
      handleWorkbookGridKeyDownCapture({
        event,
        handleGridKey,
        openHeaderContextMenuFromKeyboard,
        resetPointerInteraction,
      })
    }

    // Assert
    expect(resetPointerInteraction).not.toHaveBeenCalled()
    expect(openHeaderContextMenuFromKeyboard).not.toHaveBeenCalled()
    expect(handleGridKey).not.toHaveBeenCalled()
    for (const event of events) {
      expect(event.preventDefault).not.toHaveBeenCalled()
      expect(event.stopPropagation).not.toHaveBeenCalled()
    }
  })

  it('should not claim alt-modified browser navigation shortcuts', () => {
    // Arrange
    const handleGridKey = vi.fn()
    const openHeaderContextMenuFromKeyboard = vi.fn(() => false)
    const resetPointerInteraction = vi.fn()
    const events = [
      createKeyboardEvent({ key: 'ArrowLeft', code: 'ArrowLeft', altKey: true }),
      createKeyboardEvent({ key: 'ArrowRight', code: 'ArrowRight', altKey: true }),
      createKeyboardEvent({ key: 'Home', code: 'Home', altKey: true }),
      createKeyboardEvent({ key: 'PageDown', code: 'PageDown', altKey: true }),
      createKeyboardEvent({ key: 'Enter', code: 'Enter', altKey: true }),
      createKeyboardEvent({ key: 'a', code: 'KeyA', ctrlKey: true, altKey: true }),
      createKeyboardEvent({ key: ' ', code: 'Space', ctrlKey: true, altKey: true }),
    ]

    // Act
    for (const event of events) {
      handleWorkbookGridKeyDownCapture({
        event,
        handleGridKey,
        openHeaderContextMenuFromKeyboard,
        resetPointerInteraction,
      })
    }

    // Assert
    expect(resetPointerInteraction).not.toHaveBeenCalled()
    expect(openHeaderContextMenuFromKeyboard).not.toHaveBeenCalled()
    expect(handleGridKey).not.toHaveBeenCalled()
    for (const event of events) {
      expect(event.preventDefault).not.toHaveBeenCalled()
      expect(event.stopPropagation).not.toHaveBeenCalled()
    }
  })
})

// Helpers

function createKeyboardEvent(input: {
  altKey?: boolean | undefined
  code: string
  ctrlKey?: boolean | undefined
  key: string
  metaKey?: boolean | undefined
  shiftKey?: boolean | undefined
}): {
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
  const { altKey = false, code, ctrlKey = false, key, metaKey = false, shiftKey = false } = input
  const event = {
    altKey,
    code,
    ctrlKey,
    defaultPrevented: false,
    key,
    metaKey,
    shiftKey,
    preventDefault: vi.fn(() => {
      event.defaultPrevented = true
    }),
    stopPropagation: vi.fn(),
  }
  return event
}
