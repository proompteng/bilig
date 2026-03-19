import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getLogicalBuiltin } from "../builtins/logical.js";

const empty = (): CellValue => ({ tag: ValueTag.Empty });
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value });
const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 });
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code });

describe("logical/info builtins", () => {
  it("supports IF with explicit condition coercion and error propagation", () => {
    const IF = getLogicalBuiltin("IF")!;

    expect(IF(bool(true), text("yes"), text("no"))).toEqual(text("yes"));
    expect(IF(num(0), text("yes"), text("no"))).toEqual(text("no"));
    expect(IF(empty(), num(1))).toEqual(empty());
    expect(IF(err(ErrorCode.Ref), num(1), num(2))).toEqual(err(ErrorCode.Ref));
    expect(IF(text("hello"), num(1), num(2))).toEqual(err(ErrorCode.Value));
    expect(IF(bool(true))).toEqual(err(ErrorCode.Value));
  });

  it("supports IFERROR and IFNA with selective error handling", () => {
    const IFERROR = getLogicalBuiltin("IFERROR")!;
    const IFNA = getLogicalBuiltin("IFNA")!;

    expect(IFERROR(num(7), text("fallback"))).toEqual(num(7));
    expect(IFERROR(err(ErrorCode.Div0), text("fallback"))).toEqual(text("fallback"));
    expect(IFERROR(err(ErrorCode.Value), empty())).toEqual(empty());
    expect(IFERROR()).toEqual(err(ErrorCode.Value));

    expect(IFNA(err(ErrorCode.NA), text("missing"))).toEqual(text("missing"));
    expect(IFNA(err(ErrorCode.Ref), text("missing"))).toEqual(err(ErrorCode.Ref));
    expect(IFNA(num(3), text("missing"))).toEqual(num(3));
    expect(IFNA()).toEqual(err(ErrorCode.Value));
  });

  it("supports AND, OR, and NOT with deterministic coercion", () => {
    const AND = getLogicalBuiltin("AND")!;
    const OR = getLogicalBuiltin("OR")!;
    const NOT = getLogicalBuiltin("NOT")!;

    expect(AND(bool(true), num(1), num(-2))).toEqual(bool(true));
    expect(AND(bool(true), empty())).toEqual(bool(false));
    expect(AND(err(ErrorCode.Name), bool(true))).toEqual(err(ErrorCode.Name));
    expect(AND(text("hello"), bool(true))).toEqual(err(ErrorCode.Value));
    expect(AND()).toEqual(err(ErrorCode.Value));

    expect(OR(empty(), bool(true))).toEqual(bool(true));
    expect(OR(num(0), empty())).toEqual(bool(false));
    expect(OR(err(ErrorCode.Ref), bool(true))).toEqual(err(ErrorCode.Ref));
    expect(OR(text("hello"), bool(false))).toEqual(err(ErrorCode.Value));
    expect(OR()).toEqual(err(ErrorCode.Value));

    expect(NOT(bool(false))).toEqual(bool(true));
    expect(NOT(empty())).toEqual(bool(true));
    expect(NOT(num(2))).toEqual(bool(false));
    expect(NOT(err(ErrorCode.Div0))).toEqual(err(ErrorCode.Div0));
    expect(NOT(text("hello"))).toEqual(err(ErrorCode.Value));
    expect(NOT()).toEqual(err(ErrorCode.Value));
  });

  it("supports ISBLANK, ISNUMBER, and ISTEXT without propagating errors", () => {
    const ISBLANK = getLogicalBuiltin("ISBLANK")!;
    const ISNUMBER = getLogicalBuiltin("ISNUMBER")!;
    const ISTEXT = getLogicalBuiltin("ISTEXT")!;

    expect(ISBLANK(empty())).toEqual(bool(true));
    expect(ISBLANK(text(""))).toEqual(bool(false));
    expect(ISBLANK(err(ErrorCode.NA))).toEqual(bool(false));

    expect(ISNUMBER(num(42))).toEqual(bool(true));
    expect(ISNUMBER(bool(true))).toEqual(bool(false));
    expect(ISNUMBER(text("42"))).toEqual(bool(false));

    expect(ISTEXT(text("hello"))).toEqual(bool(true));
    expect(ISTEXT(empty())).toEqual(bool(false));
    expect(ISTEXT(err(ErrorCode.Value))).toEqual(bool(false));
  });
});
