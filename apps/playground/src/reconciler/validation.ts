import type { CellDescriptor, Descriptor, RenderModel, RenderSheetModel, SheetDescriptor, WorkbookDescriptor } from "./descriptors.js";
import { emptyRenderModel } from "./descriptors.js";

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function buildSheetModel(sheet: SheetDescriptor, order: number): RenderSheetModel {
  const cells = new Map<string, { addr: string; value?: string | number | boolean | null; formula?: string }>();
  for (const child of sheet.children) {
    assert(child.kind === "Cell", "Only <Cell> can be nested inside <Sheet>.");
    assert(Boolean(child.props.addr), "<Cell> requires an addr prop.");
    assert(
      !(child.props.value !== undefined && child.props.formula !== undefined),
      "<Cell> cannot specify both value and formula."
    );
    const cellModel: { addr: string; value?: string | number | boolean | null; formula?: string } = {
      addr: child.props.addr
    };
    if (child.props.value !== undefined) cellModel.value = child.props.value;
    if (child.props.formula !== undefined) cellModel.formula = child.props.formula;
    cells.set(child.props.addr, cellModel);
  }
  return { name: sheet.props.name, order, cells };
}

export function buildRenderModel(root: WorkbookDescriptor | null): RenderModel {
  if (!root) return emptyRenderModel();
  assert(root.kind === "Workbook", "Root descriptor must be a Workbook.");
  const model = emptyRenderModel();
  model.workbookName = root.props.name ?? "Workbook";

  root.children.forEach((sheet, order) => {
    assert(sheet.kind === "Sheet", "Only <Sheet> nodes can exist under <Workbook>.");
    assert(!model.sheets.has(sheet.props.name), `Duplicate sheet name '${sheet.props.name}'.`);
    model.sheets.set(sheet.props.name, buildSheetModel(sheet, order));
  });

  return model;
}
