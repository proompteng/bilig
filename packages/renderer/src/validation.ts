import type { WorkbookDescriptor } from "./descriptors.js";

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}

export function validateDescriptorTree(root: WorkbookDescriptor | null): void {
  if (!root) {
    return;
  }

  assert(root.kind === "Workbook", "Root descriptor must be a Workbook.");
  const sheetNames = new Set<string>();

  root.children.forEach((sheet) => {
    assert(sheet.kind === "Sheet", "Only <Sheet> nodes can exist under <Workbook>.");
    assert(Boolean(sheet.props.name), "<Sheet> requires a name prop.");
    assert(!sheetNames.has(sheet.props.name), `Duplicate sheet name '${sheet.props.name}'.`);
    sheetNames.add(sheet.props.name);

    sheet.children.forEach((cell) => {
      assert(cell.kind === "Cell", "Only <Cell> can be nested inside <Sheet>.");
      assert(Boolean(cell.props.addr), "<Cell> requires an addr prop.");
      assert(
        !(cell.props.value !== undefined && cell.props.formula !== undefined),
        "<Cell> cannot specify both value and formula."
      );
    });
  });
}
