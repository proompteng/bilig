import type { LiteralInput } from "@bilig/protocol";

export interface WorkbookProps {
  name?: string;
  children?: unknown;
}

export interface SheetProps {
  name: string;
  children?: unknown;
}

export interface CellProps {
  addr: string;
  value?: LiteralInput;
  formula?: string;
}

export interface BaseDescriptor {
  kind: "Workbook" | "Sheet" | "Cell";
  parent: Descriptor | null;
}

export interface WorkbookDescriptor extends BaseDescriptor {
  kind: "Workbook";
  props: WorkbookProps;
  children: SheetDescriptor[];
}

export interface SheetDescriptor extends BaseDescriptor {
  kind: "Sheet";
  props: SheetProps;
  children: CellDescriptor[];
}

export interface CellDescriptor extends BaseDescriptor {
  kind: "Cell";
  props: CellProps;
}

export type Descriptor = WorkbookDescriptor | SheetDescriptor | CellDescriptor;

export interface RenderCellModel {
  addr: string;
  value?: LiteralInput;
  formula?: string;
}

export interface RenderSheetModel {
  name: string;
  order: number;
  cells: Map<string, RenderCellModel>;
}

export interface RenderModel {
  workbookName: string;
  sheets: Map<string, RenderSheetModel>;
}

export function emptyRenderModel(): RenderModel {
  return {
    workbookName: "Workbook",
    sheets: new Map()
  };
}
