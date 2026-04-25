export function resolveRequiresLiveViewportState(input: {
  readonly isEditingCell: boolean
  readonly fillPreviewActive: boolean
  readonly isFillHandleDragging: boolean
  readonly isRangeMoveDragging: boolean
  readonly hasActiveHeaderDrag: boolean
  readonly hasActiveResizeColumn: boolean
  readonly hasActiveResizeRow: boolean
  readonly hasColumnResizePreview: boolean
  readonly hasRowResizePreview: boolean
}): boolean {
  void input.isEditingCell
  return (
    input.fillPreviewActive ||
    input.isFillHandleDragging ||
    input.isRangeMoveDragging ||
    input.hasActiveHeaderDrag ||
    input.hasActiveResizeColumn ||
    input.hasActiveResizeRow ||
    input.hasColumnResizePreview ||
    input.hasRowResizePreview
  )
}
