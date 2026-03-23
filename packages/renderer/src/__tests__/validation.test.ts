import { describe, expect, it } from "vitest";
import { validateDescriptorTree } from "../validation.js";
import type {
  CellDescriptor,
  CellProps,
  SheetDescriptor,
  WorkbookDescriptor,
} from "../descriptors.js";

function cell(props: CellProps): CellDescriptor {
  return {
    kind: "Cell",
    props,
    parent: null,
    container: null,
  };
}

function sheet(name: string, children: CellDescriptor[]): SheetDescriptor {
  const descriptor: SheetDescriptor = {
    kind: "Sheet",
    props: { name },
    children,
    parent: null,
    container: null,
  };
  children.forEach((child) => {
    child.parent = descriptor;
  });
  return descriptor;
}

function workbook(children: SheetDescriptor[]): WorkbookDescriptor {
  const descriptor: WorkbookDescriptor = {
    kind: "Workbook",
    props: { name: "book" },
    children,
    parent: null,
    container: null,
  };
  children.forEach((child) => {
    child.parent = descriptor;
  });
  return descriptor;
}

describe("validateDescriptorTree", () => {
  it("accepts valid workbook trees and empty roots", () => {
    expect(() => validateDescriptorTree(null)).not.toThrow();
    expect(() =>
      validateDescriptorTree(workbook([sheet("Sheet1", [cell({ addr: "A1", value: 10 })])])),
    ).not.toThrow();
  });

  it("rejects duplicate sheets, missing addrs, and conflicting cell props", () => {
    expect(() =>
      validateDescriptorTree(
        workbook([sheet("Sheet1", [cell({ addr: "A1", value: 1 })]), sheet("Sheet1", [])]),
      ),
    ).toThrow("Duplicate sheet name 'Sheet1'.");

    expect(() => validateDescriptorTree(workbook([sheet("Sheet1", [cell({ value: 1 })])]))).toThrow(
      "<Cell> requires an addr prop.",
    );

    expect(() =>
      validateDescriptorTree(
        workbook([sheet("Sheet1", [cell({ addr: "A1", value: 1, formula: "B1" })])]),
      ),
    ).toThrow("<Cell> cannot specify both value and formula.");
  });
});
