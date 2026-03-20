import type { ReactNode } from "react";
import type { LiteralInput } from "@bilig/protocol";

export interface WorkbookProps {
  name?: string;
  children?: ReactNode;
}

export interface SheetProps {
  name: string;
  children?: ReactNode;
}

export interface CellProps {
  addr: string;
  value?: LiteralInput;
  formula?: string;
  format?: string;
}

export interface BaseDescriptor {
  kind: "Workbook" | "Sheet" | "Cell";
  parent: Descriptor | null;
  container: unknown;
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
