import { isHandledGridKey } from "./gridKeyboard.js";
import { getNormalizedGridKeyboardKey } from "./gridClipboardKeyboardController.js";

interface KeyboardEventLike {
  altKey: boolean;
  code: string;
  ctrlKey: boolean;
  defaultPrevented: boolean;
  key: string;
  metaKey: boolean;
  shiftKey: boolean;
  preventDefault(): void;
  stopPropagation(): void;
}

interface HandledGridKeyEvent {
  altKey: boolean;
  cancel(): void;
  ctrlKey: boolean;
  key: string;
  metaKey: boolean;
  preventDefault(): void;
  shiftKey: boolean;
}

export function handleWorkbookGridKeyDownCapture(input: {
  event: KeyboardEventLike;
  handleGridKey: (event: HandledGridKeyEvent) => void;
  openHeaderContextMenuFromKeyboard: () => boolean;
  resetPointerInteraction: () => void;
}): void {
  const { event, handleGridKey, openHeaderContextMenuFromKeyboard, resetPointerInteraction } =
    input;
  const normalizedKey = getNormalizedGridKeyboardKey(event.key, event.code);
  resetPointerInteraction();
  if (normalizedKey === "ContextMenu" || (event.shiftKey && normalizedKey === "F10")) {
    if (openHeaderContextMenuFromKeyboard()) {
      event.preventDefault();
      event.stopPropagation();
    }
    return;
  }

  if (
    !isHandledGridKey({
      altKey: event.altKey,
      ctrlKey: event.ctrlKey,
      key: normalizedKey,
      metaKey: event.metaKey,
    })
  ) {
    return;
  }

  handleGridKey({
    altKey: event.altKey,
    cancel: () => event.stopPropagation(),
    ctrlKey: event.ctrlKey,
    key: normalizedKey,
    metaKey: event.metaKey,
    shiftKey: event.shiftKey,
    preventDefault: () => event.preventDefault(),
  });
  if (event.defaultPrevented) {
    event.stopPropagation();
  }
}
