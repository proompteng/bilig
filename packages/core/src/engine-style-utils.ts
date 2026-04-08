import type { CellStyleField, CellStylePatch, CellStyleRecord } from "@bilig/protocol";

export function cloneCellStyleRecord(style: CellStyleRecord): CellStyleRecord {
  const cloned: CellStyleRecord = { id: style.id };
  if (style.fill) {
    cloned.fill = { backgroundColor: style.fill.backgroundColor };
  }
  if (style.font) {
    cloned.font = { ...style.font };
  }
  if (style.alignment) {
    cloned.alignment = { ...style.alignment };
  }
  if (style.borders) {
    cloned.borders = {
      ...(style.borders.top ? { top: { ...style.borders.top } } : {}),
      ...(style.borders.right ? { right: { ...style.borders.right } } : {}),
      ...(style.borders.bottom ? { bottom: { ...style.borders.bottom } } : {}),
      ...(style.borders.left ? { left: { ...style.borders.left } } : {}),
    };
  }
  return cloned;
}

export function normalizeCellStylePatch(patch: CellStylePatch): CellStylePatch {
  const normalized: CellStylePatch = {};
  const fillColor = patch.fill?.backgroundColor;
  if (fillColor !== undefined) {
    normalized.fill =
      fillColor === null ? { backgroundColor: null } : { backgroundColor: fillColor };
  }
  if (patch.font) {
    normalized.font = {};
    if (patch.font.family !== undefined) {
      normalized.font.family = patch.font.family;
    }
    if (patch.font.size !== undefined) {
      normalized.font.size = patch.font.size;
    }
    if (patch.font.bold !== undefined) {
      normalized.font.bold = patch.font.bold;
    }
    if (patch.font.italic !== undefined) {
      normalized.font.italic = patch.font.italic;
    }
    if (patch.font.underline !== undefined) {
      normalized.font.underline = patch.font.underline;
    }
    if (patch.font.color !== undefined) {
      normalized.font.color = patch.font.color;
    }
  }
  if (patch.alignment) {
    normalized.alignment = {};
    if (patch.alignment.horizontal !== undefined) {
      normalized.alignment.horizontal = patch.alignment.horizontal;
    }
    if (patch.alignment.vertical !== undefined) {
      normalized.alignment.vertical = patch.alignment.vertical;
    }
    if (patch.alignment.wrap !== undefined) {
      normalized.alignment.wrap = patch.alignment.wrap;
    }
    if (patch.alignment.indent !== undefined) {
      normalized.alignment.indent = patch.alignment.indent;
    }
  }
  if (patch.borders) {
    normalized.borders = {};
    if (patch.borders.top !== undefined) {
      normalized.borders.top = patch.borders.top;
    }
    if (patch.borders.right !== undefined) {
      normalized.borders.right = patch.borders.right;
    }
    if (patch.borders.bottom !== undefined) {
      normalized.borders.bottom = patch.borders.bottom;
    }
    if (patch.borders.left !== undefined) {
      normalized.borders.left = patch.borders.left;
    }
  }
  return normalized;
}

export function applyStylePatch(
  baseStyle: Omit<CellStyleRecord, "id">,
  patch: CellStylePatch,
): Omit<CellStyleRecord, "id"> {
  const next = cloneStyleWithoutId(baseStyle);
  const backgroundColor = patch.fill?.backgroundColor;
  if (backgroundColor !== undefined) {
    if (backgroundColor === null) {
      delete next.fill;
    } else {
      next.fill = { backgroundColor };
    }
  }
  if (patch.font) {
    const font = { ...next.font };
    applyOptionalField(font, "family", patch.font.family);
    applyOptionalField(font, "size", patch.font.size);
    applyOptionalField(font, "bold", patch.font.bold);
    applyOptionalField(font, "italic", patch.font.italic);
    applyOptionalField(font, "underline", patch.font.underline);
    applyOptionalField(font, "color", patch.font.color);
    if (Object.keys(font).length > 0) {
      next.font = font;
    } else {
      delete next.font;
    }
  }
  if (patch.alignment) {
    const alignment = { ...next.alignment };
    applyOptionalField(alignment, "horizontal", patch.alignment.horizontal);
    applyOptionalField(alignment, "vertical", patch.alignment.vertical);
    applyOptionalField(alignment, "wrap", patch.alignment.wrap);
    applyOptionalField(alignment, "indent", patch.alignment.indent);
    if (Object.keys(alignment).length > 0) {
      next.alignment = alignment;
    } else {
      delete next.alignment;
    }
  }
  if (patch.borders) {
    const borders = { ...next.borders };
    applyOptionalField(borders, "top", normalizeBorderPatchSide(patch.borders.top));
    applyOptionalField(borders, "right", normalizeBorderPatchSide(patch.borders.right));
    applyOptionalField(borders, "bottom", normalizeBorderPatchSide(patch.borders.bottom));
    applyOptionalField(borders, "left", normalizeBorderPatchSide(patch.borders.left));
    if (Object.keys(borders).length > 0) {
      next.borders = borders;
    } else {
      delete next.borders;
    }
  }
  return next;
}

export function clearStyleFields(
  baseStyle: Omit<CellStyleRecord, "id">,
  fields: readonly CellStyleField[] | undefined,
): Omit<CellStyleRecord, "id"> {
  const clearAll = !fields || fields.length === 0;
  const cleared = new Set(fields ?? []);
  if (clearAll) {
    return {};
  }
  const next = cloneStyleWithoutId(baseStyle);
  if (cleared.has("backgroundColor")) {
    delete next.fill;
  }
  const font = filterStyleSection(
    next.font,
    [
      ["fontFamily", "family"],
      ["fontSize", "size"],
      ["fontBold", "bold"],
      ["fontItalic", "italic"],
      ["fontUnderline", "underline"],
      ["fontColor", "color"],
    ],
    cleared,
  );
  if (font) {
    next.font = font;
  } else {
    delete next.font;
  }
  const alignment = filterStyleSection(
    next.alignment,
    [
      ["alignmentHorizontal", "horizontal"],
      ["alignmentVertical", "vertical"],
      ["alignmentWrap", "wrap"],
      ["alignmentIndent", "indent"],
    ],
    cleared,
  );
  if (alignment) {
    next.alignment = alignment;
  } else {
    delete next.alignment;
  }
  const borders = filterStyleSection(
    next.borders,
    [
      ["borderTop", "top"],
      ["borderRight", "right"],
      ["borderBottom", "bottom"],
      ["borderLeft", "left"],
    ],
    cleared,
  );
  if (borders) {
    next.borders = borders;
  } else {
    delete next.borders;
  }
  return next;
}

function cloneStyleWithoutId(style: Omit<CellStyleRecord, "id">): Omit<CellStyleRecord, "id"> {
  const cloned: Omit<CellStyleRecord, "id"> = {};
  if (style.fill) {
    cloned.fill = { backgroundColor: style.fill.backgroundColor };
  }
  if (style.font) {
    cloned.font = { ...style.font };
  }
  if (style.alignment) {
    cloned.alignment = { ...style.alignment };
  }
  if (style.borders) {
    cloned.borders = {
      ...(style.borders.top ? { top: { ...style.borders.top } } : {}),
      ...(style.borders.right ? { right: { ...style.borders.right } } : {}),
      ...(style.borders.bottom ? { bottom: { ...style.borders.bottom } } : {}),
      ...(style.borders.left ? { left: { ...style.borders.left } } : {}),
    };
  }
  return cloned;
}

function applyOptionalField<T extends object, K extends keyof T>(
  target: T,
  key: K,
  value: T[K] | null | undefined,
): void {
  if (value === undefined) {
    return;
  }
  if (value === null) {
    delete target[key];
    return;
  }
  target[key] = value;
}

function normalizeBorderPatchSide(
  side: NonNullable<NonNullable<CellStylePatch["borders"]>["top"]> | null | undefined,
): NonNullable<NonNullable<CellStyleRecord["borders"]>["top"]> | null | undefined {
  if (side === undefined) {
    return undefined;
  }
  if (side === null) {
    return null;
  }
  if (!side.style || !side.weight || !side.color) {
    return null;
  }
  return {
    style: side.style,
    weight: side.weight,
    color: side.color,
  };
}

function filterStyleSection<T extends object>(
  section: T | undefined,
  keys: ReadonlyArray<[CellStyleField, keyof T]>,
  cleared: ReadonlySet<CellStyleField>,
): T | undefined {
  if (!section) {
    return undefined;
  }
  const next = { ...section };
  keys.forEach(([field, key]) => {
    if (cleared.has(field)) {
      delete next[key];
    }
  });
  return Object.keys(next).length > 0 ? next : undefined;
}
