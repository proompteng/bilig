import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getBuiltin, getBuiltinId } from "../builtins.js";
import { evaluateAst, evaluateAstResult } from "../js-evaluator.js";
import { parseFormula } from "../parser.js";

describe("formula builtins and JS evaluator", () => {
  it("covers scalar builtins and builtin id lookup", () => {
    const SUM = getBuiltin("sum")!;
    const AVG = getBuiltin("AVG")!;
    const MOD = getBuiltin("MOD")!;
    const LEN = getBuiltin("LEN")!;
    const CONCAT = getBuiltin("CONCAT")!;
    const IF = getBuiltin("IF")!;
    const AND = getBuiltin("AND")!;
    const OR = getBuiltin("OR")!;
    const NOT = getBuiltin("NOT")!;

    expect(getBuiltinId("sum")).toBeDefined();
    expect(getBuiltinId("")).toBeUndefined();
    expect(getBuiltin("missing")).toBeUndefined();

    expect(
      SUM(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(
      AVG(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: "ignored", stringId: 0 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(MOD({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(LEN({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(
      CONCAT(
        { tag: ValueTag.String, value: "hello", stringId: 0 },
        { tag: ValueTag.Empty },
        { tag: ValueTag.Number, value: 7 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "hello7", stringId: 0 });
    expect(
      IF(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(AND({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: true })).toEqual(
      {
        tag: ValueTag.Boolean,
        value: true,
      },
    );
    expect(OR({ tag: ValueTag.Empty }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(NOT({ tag: ValueTag.Empty })).toEqual({ tag: ValueTag.Boolean, value: true });
  });

  it("evaluates range, concat, unary, comparison, builtin, and error paths", () => {
    const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
    const empty = (): CellValue => ({ tag: ValueTag.Empty });
    const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 });
    const context = {
      sheetName: "Sheet1",
      currentAddress: "C7",
      resolveCell: (sheetName: string, address: string): CellValue => {
        if (sheetName === "Sheet2" && address === "B1") {
          return num(9);
        }
        if (address === "A1") {
          return num(4);
        }
        if (address === "A2") {
          return text("x");
        }
        return empty();
      },
      resolveRange: (
        _sheetName: string,
        start: string,
        end: string,
        refKind: "cells" | "rows" | "cols",
      ): CellValue[] => {
        if (refKind === "cells" && start === "A1" && end === "B2") {
          return [num(1), num(2), num(3)];
        }
        if (refKind === "cells" && start === "A10" && end === "B12") {
          return [num(58), num(35), num(11), num(25), num(10), num(23)];
        }
        if (refKind === "cells" && start === "D10" && end === "E12") {
          return [num(45.35), num(47.65), num(17.56), num(18.44), num(16.09), num(16.91)];
        }
        if (refKind === "cells" && start === "A20" && end === "A24") {
          return [num(6), num(7), num(9), num(15), num(21)];
        }
        if (refKind === "cells" && start === "B20" && end === "B24") {
          return [num(20), num(28), num(31), num(38), num(40)];
        }
        if (refKind === "cells" && start === "D20" && end === "D24") {
          return [num(1), num(2), num(3), num(4), num(5)];
        }
        if (refKind === "cells" && start === "A30" && end === "A32") {
          return [num(5), num(8), num(11)];
        }
        if (refKind === "cells" && start === "B30" && end === "B32") {
          return [num(1), num(2), num(3)];
        }
        if (refKind === "cells" && start === "A40" && end === "A43") {
          return [num(10), num(20), num(20), num(30)];
        }
        if (refKind === "cells" && start === "A50" && end === "A57") {
          return [num(1), num(2), num(4), num(7), num(8), num(9), num(10), num(12)];
        }
        if (refKind === "cells" && start === "A60" && end === "A65") {
          return [num(79), num(85), num(78), num(85), num(50), num(81)];
        }
        if (refKind === "cells" && start === "B60" && end === "B62") {
          return [num(60), num(80), num(90)];
        }
        if (refKind === "cells" && start === "A70" && end === "A75") {
          return [num(1), num(2), num(2), num(3), num(3), num(4)];
        }
        if (refKind === "cells" && start === "A80" && end === "A85") {
          return [num(-70000), num(12000), num(15000), num(18000), num(21000), num(26000)];
        }
        if (refKind === "cells" && start === "A90" && end === "B91") {
          return [text("カタカナ"), { tag: ValueTag.Error, code: ErrorCode.Ref }];
        }
        if (refKind === "cells" && start === "B80" && end === "B85") {
          return [num(-120000), num(39000), num(30000), num(21000), num(37000), num(46000)];
        }
        if (refKind === "cells" && start === "C80" && end === "C84") {
          return [num(-10000), num(2750), num(4250), num(3250), num(2750)];
        }
        if (refKind === "cells" && start === "D80" && end === "D84") {
          return [num(39448), num(39508), num(39751), num(39859), num(39904)];
        }
        if (refKind === "rows") {
          return [num(6)];
        }
        return [];
      },
      resolveFormula: (sheetName: string, address: string): string | undefined =>
        sheetName === "Sheet2" && address === "B1" ? "A1*2" : undefined,
      listSheetNames: (): string[] => ["Sheet1", "Summary", "Sheet2"],
    };

    expect(evaluateAst(parseFormula('A1&"!"'), context)).toEqual({
      tag: ValueTag.String,
      value: "4!",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('A2="X"'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(evaluateAst(parseFormula('"b">"A"'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(evaluateAst(parseFormula("-A1"), context)).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(evaluateAst(parseFormula("A1>=Sheet2!B1"), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    });
    expect(evaluateAst(parseFormula("SUM(A1:B2)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 6,
    });
    expect(evaluateAst(parseFormula("SUM(1:1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 6,
    });
    expect(evaluateAst(parseFormula("A1/0"), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(evaluateAst(parseFormula("A2+1"), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluateAst(parseFormula("SIN(A1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: Math.sin(4),
    });
    expect(evaluateAst(parseFormula("POWER(2,3)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 8,
    });
    expect(evaluateAst(parseFormula("TRUNC(-3.98,1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: -3.9,
    });
    expect(evaluateAst(parseFormula("PRODUCT(2,3,4)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 24,
    });
    expect(evaluateAst(parseFormula("SUMSQ(2,3)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 13,
    });
    expect(evaluateAst(parseFormula("SIGN(-2)"), context)).toEqual({
      tag: ValueTag.Number,
      value: -1,
    });
    expect(evaluateAst(parseFormula("ACOSH(1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(evaluateAst(parseFormula("COMBIN(8,3)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 56,
    });
    expect(evaluateAst(parseFormula("FACTDOUBLE(6)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 48,
    });
    expect(
      evaluateAst(parseFormula("MDETERM(C1:D2)"), {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] =>
          start === "C1" && end === "D2"
            ? [
                { tag: ValueTag.Number, value: 1 },
                { tag: ValueTag.Number, value: 2 },
                { tag: ValueTag.Number, value: 3 },
                { tag: ValueTag.Number, value: 4 },
              ]
            : context.resolveRange(_sheetName, start, end, "cells"),
      }),
    ).toEqual({ tag: ValueTag.Number, value: -2 });
    expect(evaluateAst(parseFormula("DAYS(46101,46094)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 7,
    });
    expect(evaluateAst(parseFormula("WEEKNUM(46096)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 12,
    });
    expect(evaluateAst(parseFormula("WORKDAY(46094,1,46097)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 46098,
    });
    expect(evaluateAst(parseFormula("NETWORKDAYS(46094,46101,46097,46101)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(evaluateAst(parseFormula("COUNTBLANK(A1,A3,A4)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("DAYS360(DATE(2024,1,29),DATE(2024,3,31))"), context)).toEqual({
      tag: ValueTag.Number,
      value: 62,
    });
    expect(
      evaluateAst(parseFormula("DAYS360(DATE(2024,1,29),DATE(2024,3,31),TRUE)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 61,
    });
    expect(
      evaluateAst(parseFormula("YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(182 / 365, 12),
    });
    expect(evaluateAst(parseFormula("ISOWEEKNUM(DATE(2024,1,1))"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula('TIMEVALUE("1:30 PM")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 0.5625,
    });
    expect(evaluateAst(parseFormula('REPLACE("alphabet",3,2,"Z")'), context)).toEqual({
      tag: ValueTag.String,
      value: "alZabet",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('SUBSTITUTE("banana","an","oo",2)'), context)).toEqual({
      tag: ValueTag.String,
      value: "banooa",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('REPT("x",3)'), context)).toEqual({
      tag: ValueTag.String,
      value: "xxx",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("T.DIST(0.1,2,TRUE)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5352672807929295, 12),
    });
    expect(evaluateAst(parseFormula('TEXTJOIN(",",TRUE,"a","b")'), context)).toEqual({
      tag: ValueTag.String,
      value: "a,b",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('CHOOSE(2,"red","blue","green")'), context)).toEqual({
      tag: ValueTag.String,
      value: "blue",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("LET(x,2,x+3)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(evaluateAst(parseFormula("LAMBDA(x,x+1)(4)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(evaluateAst(parseFormula("LET(fn,LAMBDA(x,x*2),fn(4))"), context)).toEqual({
      tag: ValueTag.Number,
      value: 8,
    });
    expect(evaluateAst(parseFormula("LAMBDA(x,y,IF(ISOMITTED(y),x,y))(4)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(evaluateAst(parseFormula("LAMBDA(x,y,IF(ISOMITTED(y),x,y))(4,9)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 9,
    });
    expect(evaluateAst(parseFormula("LAMBDA(x,ISOMITTED(x))()"), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(evaluateAst(parseFormula("LAMBDA(x,IF(ISOMITTED(x),9,x))(4)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(evaluateAst(parseFormula("TRUE()"), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(evaluateAst(parseFormula("FALSE()"), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    });
    expect(evaluateAst(parseFormula('T("alpha")'), context)).toEqual({
      tag: ValueTag.String,
      value: "alpha",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("N(TRUE())"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula('TYPE("alpha")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("DELTA(4,4)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("GESTEP(-1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(evaluateAst(parseFormula("GAUSS(0)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    });
    expect(evaluateAst(parseFormula("PHI(0)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: 0.3989422804014327,
    });
    expect(evaluateAst(parseFormula('IFS(A1>3,"big",TRUE(),"small")'), context)).toEqual({
      tag: ValueTag.String,
      value: "big",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('SWITCH(A1,4,"four","other")'), context)).toEqual({
      tag: ValueTag.String,
      value: "four",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("XOR(TRUE(),FALSE(),TRUE())"), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    });
    expect(evaluateAst(parseFormula('TEXTBEFORE("alpha-beta","-")'), context)).toEqual({
      tag: ValueTag.String,
      value: "alpha",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('TEXTAFTER("alpha-beta","-")'), context)).toEqual({
      tag: ValueTag.String,
      value: "beta",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('NUMBERVALUE("2.500,27",",",".")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 2500.27,
    });
    expect(evaluateAst(parseFormula('TEXT(1234.567,"#,##0.00")'), context)).toEqual({
      tag: ValueTag.String,
      value: "1,234.57",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('TEXT(DATE(2024,3,5),"yyyy-mm-dd")'), context)).toEqual({
      tag: ValueTag.String,
      value: "2024-03-05",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("PHONETIC(A90:B91)"), context)).toEqual({
      tag: ValueTag.String,
      value: "カタカナ",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('REGEXTEST("Alpha-42","[a-z]+-[0-9]+",1)'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(
      evaluateAst(parseFormula('REGEXREPLACE("abc123","([a-z]+)([0-9]+)","$2-$1")'), context),
    ).toEqual({
      tag: ValueTag.String,
      value: "123-abc",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('VALUETOTEXT("alpha",1)'), context)).toEqual({
      tag: ValueTag.String,
      value: '"alpha"',
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('ASC("ＡＢＣ　１２３")'), context)).toEqual({
      tag: ValueTag.String,
      value: "ABC 123",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('JIS("ABC 123")'), context)).toEqual({
      tag: ValueTag.String,
      value: "ＡＢＣ　１２３",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('DBCS("ｶﾞｷﾞｸﾞｹﾞｺﾞ")'), context)).toEqual({
      tag: ValueTag.String,
      value: "ガギグゲゴ",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("BAHTTEXT(1234)"), context)).toEqual({
      tag: ValueTag.String,
      value: "หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("ROW()"), context)).toEqual({
      tag: ValueTag.Number,
      value: 7,
    });
    expect(evaluateAst(parseFormula("COLUMN()"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluateAst(parseFormula("ROW(A2:B4)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("COLUMN(B:D)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("FORMULATEXT(Sheet2!B1)"), context)).toEqual({
      tag: ValueTag.String,
      value: "=A1*2",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("FORMULA(Sheet2!B1)"), context)).toEqual({
      tag: ValueTag.String,
      value: "=A1*2",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula("SHEET()"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula('SHEET("Summary")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("SHEETS()"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluateAst(parseFormula("SHEETS(A1:B2)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula('CELL("address",B3)'), context)).toEqual({
      tag: ValueTag.String,
      value: "$B$3",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('CELL("row",B3)'), context)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluateAst(parseFormula('CELL("col",B3)'), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula('CELL("contents",A1)'), context)).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(evaluateAst(parseFormula('CELL("type",A2)'), context)).toEqual({
      tag: ValueTag.String,
      value: "l",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('BIN2DEC("1111111111")'), context)).toEqual({
      tag: ValueTag.Number,
      value: -1,
    });
    const besseli = evaluateAst(parseFormula("BESSELI(1.5,1)"), context);
    expect(besseli).toMatchObject({ tag: ValueTag.Number });
    if (besseli.tag !== ValueTag.Number) {
      throw new Error("Expected BESSELI result to be numeric");
    }
    expect(besseli.value).toBeCloseTo(0.981666428, 7);
    const besselj = evaluateAst(parseFormula("BESSELJ(1.9,2)"), context);
    expect(besselj).toMatchObject({ tag: ValueTag.Number });
    if (besselj.tag !== ValueTag.Number) {
      throw new Error("Expected BESSELJ result to be numeric");
    }
    expect(besselj.value).toBeCloseTo(0.329925728, 7);
    const besselk = evaluateAst(parseFormula("BESSELK(1.5,1)"), context);
    expect(besselk).toMatchObject({ tag: ValueTag.Number });
    if (besselk.tag !== ValueTag.Number) {
      throw new Error("Expected BESSELK result to be numeric");
    }
    expect(besselk.value).toBeCloseTo(0.277387804, 7);
    const bessely = evaluateAst(parseFormula("BESSELY(2.5,1)"), context);
    expect(bessely).toMatchObject({ tag: ValueTag.Number });
    if (bessely.tag !== ValueTag.Number) {
      throw new Error("Expected BESSELY result to be numeric");
    }
    expect(bessely.value).toBeCloseTo(0.145918138, 7);
    expect(evaluateAst(parseFormula('CONVERT(6,"mi","km")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 9.656064,
    });
    expect(evaluateAst(parseFormula('CONVERT(68,"F","C")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 20,
    });
    const euroconvert = evaluateAst(parseFormula('EUROCONVERT(1,"FRF","DEM",TRUE,3)'), context);
    expect(euroconvert).toMatchObject({ tag: ValueTag.Number });
    if (euroconvert.tag !== ValueTag.Number) {
      throw new Error("Expected EUROCONVERT result to be numeric");
    }
    expect(euroconvert.value).toBeCloseTo(0.29728616, 12);
    expect(evaluateAst(parseFormula('COMPLEX(3,-4,"j")'), context)).toEqual({
      tag: ValueTag.String,
      value: "3-4j",
      stringId: 0,
    });
    expect(evaluateAst(parseFormula('IMABS("3+4i")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(
      evaluateAst(parseFormula('DATEDIF(DATE(2020,1,15),DATE(2021,3,20),"YM")'), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("WORKDAY.INTL(DATE(2026,3,13),1,7)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 46097,
    });
    expect(
      evaluateAst(parseFormula("NETWORKDAYS.INTL(DATE(2026,3,13),DATE(2026,3,17),7)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluateAst(parseFormula("FVSCHEDULE(1000,0.09,0.11,0.1)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1330.89, 12),
    });
    expect(evaluateAst(parseFormula("DB(10000,1000,5,1)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3690, 12),
    });
    expect(evaluateAst(parseFormula("DDB(2400,300,10,2)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(384, 12),
    });
    expect(evaluateAst(parseFormula("VDB(2400,300,10,1,3)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(691.2, 12),
    });
    expect(evaluateAst(parseFormula("SLN(10000,1000,9)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1000,
    });
    expect(evaluateAst(parseFormula("SYD(10000,1000,9,1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1800,
    });
    expect(evaluateAst(parseFormula("PV(0.1,2,-576.1904761904761)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1000.0000000000006, 12),
    });
    expect(evaluateAst(parseFormula("PMT(0.1,2,1000)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-576.1904761904758, 12),
    });
    expect(evaluateAst(parseFormula("NPER(0.1,-576.1904761904761,1000)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.9999999999999982, 12),
    });
    expect(evaluateAst(parseFormula("RATE(48,-200,8000)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.007701472488246008, 12),
    });
    expect(evaluateAst(parseFormula("IPMT(0.1,1,2,1000)"), context)).toEqual({
      tag: ValueTag.Number,
      value: -100,
    });
    expect(evaluateAst(parseFormula("PPMT(0.1,1,2,1000)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-476.1904761904758, 12),
    });
    expect(evaluateAst(parseFormula("ISPMT(0.1,1,2,1000)"), context)).toEqual({
      tag: ValueTag.Number,
      value: -50,
    });
    expect(evaluateAst(parseFormula("CUMIPMT(9%/12,30*12,125000,13,24,0)"), context)).toMatchObject(
      {
        tag: ValueTag.Number,
        value: expect.closeTo(-11135.232130750845, 12),
      },
    );
    expect(
      evaluateAst(parseFormula("CUMPRINC(9%/12,30*12,125000,13,24,0)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-934.1071234208765, 12),
    });
    expect(
      evaluateAst(parseFormula("DISC(DATE(2023,1,1),DATE(2023,4,1),97,100,2)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    });
    expect(
      evaluateAst(parseFormula("INTRATE(DATE(2023,1,1),DATE(2023,4,1),1000,1030,2)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    });
    expect(
      evaluateAst(parseFormula("RECEIVED(DATE(2023,1,1),DATE(2023,4,1),1000,0.12,2)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1030.9278350515465, 12),
    });
    expect(
      evaluateAst(parseFormula("PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.79583333333333, 12),
    });
    expect(
      evaluateAst(parseFormula("YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.05282257198685834, 12),
    });
    expect(
      evaluateAst(
        parseFormula("PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)"),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.98449887555694, 12),
    });
    expect(
      evaluateAst(
        parseFormula("YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)"),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.060954333691538576, 12),
    });
    expect(
      evaluateAst(
        parseFormula(
          "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)",
        ),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(113.597717474079, 12),
    });
    expect(
      evaluateAst(
        parseFormula(
          "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)",
        ),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0772455415972989, 11),
    });
    expect(
      evaluateAst(
        parseFormula(
          "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)",
        ),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.8782860147213, 12),
    });
    expect(
      evaluateAst(
        parseFormula(
          "ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)",
        ),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0451922356291692, 12),
    });
    expect(
      evaluateAst(parseFormula("TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(98.45, 12),
    });
    expect(
      evaluateAst(parseFormula("TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09141696292534264, 12),
    });
    expect(
      evaluateAst(parseFormula("TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09415149356594302, 12),
    });
    expect(evaluateAst(parseFormula("IRR(A80:A85)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.08663094803653162, 12),
    });
    expect(evaluateAst(parseFormula("MIRR(B80:B85,10%,12%)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1260941303659051, 12),
    });
    expect(evaluateAst(parseFormula("XNPV(0.09,C80:C84,D80:D84)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2086.647602031535, 9),
    });
    expect(evaluateAst(parseFormula("XIRR(C80:C84,D80:D84)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.37336253351883136, 12),
    });
    expect(
      evaluateAst(parseFormula("COUPDAYBS(DATE(2007,1,25),DATE(2009,11,15),2,4)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 70,
    });
    expect(
      evaluateAst(parseFormula("COUPDAYS(DATE(2007,1,25),DATE(2009,11,15),2,4)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 180,
    });
    expect(
      evaluateAst(parseFormula("COUPDAYSNC(DATE(2007,1,25),DATE(2009,11,15),2,4)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 110,
    });
    expect(
      evaluateAst(parseFormula("COUPNCD(DATE(2007,1,25),DATE(2009,11,15),2,4)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 39217,
    });
    expect(
      evaluateAst(parseFormula("COUPNUM(DATE(2007,1,25),DATE(2009,11,15),2,4)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 6,
    });
    expect(
      evaluateAst(parseFormula("COUPPCD(DATE(2007,1,25),DATE(2009,11,15),2,4)"), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 39036,
    });
    expect(
      evaluateAst(
        parseFormula("PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)"),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(94.63436162132213, 12),
    });
    expect(
      evaluateAst(
        parseFormula("YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)"),
        context,
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.065, 7),
    });
    expect(
      evaluateAst(parseFormula("DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(10.919145281591925, 12),
    });
    expect(
      evaluateAst(parseFormula("MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)"), context),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5.735669813918838, 12),
    });
    expect(evaluateAst(parseFormula("CORREL(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("COVAR(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("COVARIANCE.P(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("COVARIANCE.S(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluateAst(parseFormula("PEARSON(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("INTERCEPT(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("SLOPE(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluateAst(parseFormula("RSQ(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("STEYX(A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(evaluateAst(parseFormula("FORECAST(4,A30:A32,B30:B32)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 14,
    });
    expect(evaluateAst(parseFormula("RANK(20,A40:A43)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("RANK.EQ(20,A40:A43)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("RANK.AVG(20,A40:A43)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2.5,
    });
    expect(evaluateAst(parseFormula("MEDIAN(A50:A57)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 7.5,
    });
    expect(evaluateAst(parseFormula("SMALL(A50:A57,3)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(evaluateAst(parseFormula("LARGE(A50:A57,2)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 10,
    });
    expect(evaluateAst(parseFormula("PERCENTILE(A50:A57,0.25)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3.5,
    });
    expect(evaluateAst(parseFormula("PERCENTILE.INC(A50:A57,0.25)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3.5,
    });
    expect(evaluateAst(parseFormula("PERCENTILE.EXC(A50:A57,0.25)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2.5,
    });
    expect(evaluateAst(parseFormula("PERCENTRANK(A50:A57,8)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0.571,
    });
    expect(evaluateAst(parseFormula("PERCENTRANK.INC(A50:A57,8)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0.571,
    });
    expect(evaluateAst(parseFormula("PERCENTRANK.EXC(A50:A57,8)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0.555,
    });
    expect(evaluateAst(parseFormula("QUARTILE(A50:A57,1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3.5,
    });
    expect(evaluateAst(parseFormula("QUARTILE.INC(A50:A57,1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 3.5,
    });
    expect(evaluateAst(parseFormula("QUARTILE.EXC(A50:A57,1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2.5,
    });
    expect(evaluateAstResult(parseFormula("MODE.MULT(A70:A75)"), context)).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ],
    });
    expect(evaluateAstResult(parseFormula("FREQUENCY(A60:A65,B60:B62)"), context)).toEqual({
      kind: "array",
      rows: 4,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
      ],
    });
    const chisqDist = evaluateAst(parseFormula("CHISQDIST(18.307,10)"), context);
    expect(chisqDist).toMatchObject({ tag: ValueTag.Number });
    expect(chisqDist.value).toBeCloseTo(0.05000058909139826, 12);

    const chiInv = evaluateAst(parseFormula("CHIINV(0.050001,10)"), context);
    expect(chiInv).toMatchObject({ tag: ValueTag.Number });
    expect(chiInv.value).toBeCloseTo(18.30697345696106, 12);

    const chisqInvRt = evaluateAst(parseFormula("CHISQ.INV.RT(0.050001,10)"), context);
    expect(chisqInvRt).toMatchObject({ tag: ValueTag.Number });
    expect(chisqInvRt.value).toBeCloseTo(18.30697345696106, 12);

    const chisqInvAlias = evaluateAst(parseFormula("CHISQINV(0.050001,10)"), context);
    expect(chisqInvAlias).toMatchObject({ tag: ValueTag.Number });
    expect(chisqInvAlias.value).toBeCloseTo(18.30697345696106, 12);

    const legacyChiInv = evaluateAst(parseFormula("LEGACY.CHIINV(0.050001,10)"), context);
    expect(legacyChiInv).toMatchObject({ tag: ValueTag.Number });
    expect(legacyChiInv.value).toBeCloseTo(18.30697345696106, 12);

    const chisqInv = evaluateAst(parseFormula("CHISQ.INV(0.93,1)"), context);
    expect(chisqInv).toMatchObject({ tag: ValueTag.Number });
    expect(chisqInv.value).toBeCloseTo(3.2830202867594993, 12);
    const chisqTest = evaluateAst(parseFormula("CHISQ.TEST(A10:B12,D10:E12)"), context);
    expect(chisqTest).toMatchObject({ tag: ValueTag.Number });
    expect(chisqTest.value).toBeCloseTo(0.0003082, 7);
    const chiTest = evaluateAst(parseFormula("CHITEST(A10:B12,D10:E12)"), context);
    expect(chiTest).toMatchObject({ tag: ValueTag.Number });
    expect(chiTest.value).toBeCloseTo(0.0003082, 7);
    const legacyChiTest = evaluateAst(parseFormula("LEGACY.CHITEST(A10:B12,D10:E12)"), context);
    expect(legacyChiTest).toMatchObject({ tag: ValueTag.Number });
    expect(legacyChiTest.value).toBeCloseTo(0.0003082, 7);
    const betaDist = evaluateAst(parseFormula("BETA.DIST(2,8,10,TRUE,1,3)"), context);
    expect(betaDist).toMatchObject({ tag: ValueTag.Number });
    expect(betaDist.value).toBeCloseTo(0.6854705810117458, 10);
    const betaInv = evaluateAst(parseFormula("BETA.INV(0.6854705810117458,8,10,1,3)"), context);
    expect(betaInv).toMatchObject({ tag: ValueTag.Number });
    expect(betaInv.value).toBeCloseTo(2, 10);
    const fDistRt = evaluateAst(parseFormula("F.DIST.RT(15.2068649,6,4)"), context);
    expect(fDistRt).toMatchObject({ tag: ValueTag.Number });
    expect(fDistRt.value).toBeCloseTo(0.01, 9);
    const fInvRt = evaluateAst(parseFormula("F.INV.RT(0.01,6,4)"), context);
    expect(fInvRt).toMatchObject({ tag: ValueTag.Number });
    expect(fInvRt.value).toBeCloseTo(15.206864870947697, 7);
    const tDist = evaluateAst(parseFormula("T.DIST(1,1,TRUE)"), context);
    expect(tDist).toMatchObject({ tag: ValueTag.Number });
    expect(tDist.value).toBeCloseTo(0.75, 12);
    const tDistRt = evaluateAst(parseFormula("T.DIST.RT(1,1)"), context);
    expect(tDistRt).toMatchObject({ tag: ValueTag.Number });
    expect(tDistRt.value).toBeCloseTo(0.25, 12);
    const tDist2T = evaluateAst(parseFormula("T.DIST.2T(1,1)"), context);
    expect(tDist2T).toMatchObject({ tag: ValueTag.Number });
    expect(tDist2T.value).toBeCloseTo(0.5, 12);
    const tInv = evaluateAst(parseFormula("T.INV(0.75,1)"), context);
    expect(tInv).toMatchObject({ tag: ValueTag.Number });
    expect(tInv.value).toBeCloseTo(1, 9);
    const tInv2T = evaluateAst(parseFormula("T.INV.2T(0.5,1)"), context);
    expect(tInv2T).toMatchObject({ tag: ValueTag.Number });
    expect(tInv2T.value).toBeCloseTo(1, 9);
    const confidenceT = evaluateAst(parseFormula("CONFIDENCE.T(0.5,2,4)"), context);
    expect(confidenceT).toMatchObject({ tag: ValueTag.Number });
    expect(confidenceT.value).toBeCloseTo(0.764892328404345, 12);
    expect(evaluateAst(parseFormula("STANDARDIZE(1,0,1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("CONFIDENCE.NORM(0.05,1,100)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1959963986120195, 12),
    });
    expect(evaluateAst(parseFormula("MODE(A70:A75)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("MODE.SNGL(A70:A75)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula("STDEV(1,2,3,4)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(5 / 3), 12),
    });
    expect(evaluateAst(parseFormula('STDEVA(2,TRUE(),"skip")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("VAR(1,2,3,4)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5 / 3, 12),
    });
    expect(evaluateAst(parseFormula('VARA(2,TRUE(),"skip")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("SKEW(2,3,4,5,6)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(evaluateAst(parseFormula("KURT(1,2,3,4,5)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1.2, 12),
    });
    expect(evaluateAst(parseFormula("NORMDIST(1,0,1,TRUE)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    });
    expect(evaluateAst(parseFormula("NORMINV(0.8413447460685429,0,1)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    });
    expect(evaluateAst(parseFormula("NORMSDIST(1)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    });
    expect(evaluateAst(parseFormula("NORMSINV(0.001)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.090232306167813, 8),
    });
    expect(evaluateAst(parseFormula("LOGINV(0.5,0,1)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluateAst(parseFormula("LOGNORMDIST(1,0,1)"), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 8),
    });
    const gammaInv = evaluateAst(parseFormula("GAMMA.INV(0.08030139707139418,3,2)"), context);
    expect(gammaInv).toMatchObject({ tag: ValueTag.Number });
    expect(gammaInv.value).toBeCloseTo(2, 10);
    const legacyGammaInv = evaluateAst(parseFormula("GAMMAINV(0.08030139707139418,3,2)"), context);
    expect(legacyGammaInv).toMatchObject({ tag: ValueTag.Number });
    expect(legacyGammaInv.value).toBeCloseTo(2, 10);
    const fTest = evaluateAst(parseFormula("F.TEST(A20:A24,B20:B24)"), context);
    expect(fTest).toMatchObject({ tag: ValueTag.Number });
    expect(fTest.value).toBeCloseTo(0.648317846786175, 12);
    const legacyFTest = evaluateAst(parseFormula("FTEST(A20:A24,B20:B24)"), context);
    expect(legacyFTest).toMatchObject({ tag: ValueTag.Number });
    expect(legacyFTest.value).toBeCloseTo(0.648317846786175, 12);
    const tTest = evaluateAst(parseFormula("T.TEST(A30:A32,B30:B32,2,1)"), context);
    expect(tTest).toMatchObject({ tag: ValueTag.Number });
    expect(tTest.value).toBeCloseTo(0.035098718645984794, 12);
    const legacyTTest = evaluateAst(parseFormula("TTEST(A30:A32,B30:B32,2,1)"), context);
    expect(legacyTTest).toMatchObject({ tag: ValueTag.Number });
    expect(legacyTTest.value).toBeCloseTo(0.035098718645984794, 12);
    const zTest = evaluateAst(parseFormula("Z.TEST(D20:D24,2,1)"), context);
    expect(zTest).toMatchObject({ tag: ValueTag.Number });
    expect(zTest.value).toBeCloseTo(0.012673617875446075, 12);
    const legacyZTest = evaluateAst(parseFormula("ZTEST(D20:D24,2,1)"), context);
    expect(legacyZTest).toMatchObject({ tag: ValueTag.Number });
    expect(legacyZTest.value).toBeCloseTo(0.012673617875446075, 12);
    expect(evaluateAst(parseFormula('LENB("é")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluateAst(parseFormula('INDIRECT("A1")+1'), context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(evaluateAst(parseFormula("SKEWP(1,2,3)"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    const legacyNormsDist = evaluateAst(parseFormula("LEGACY.NORMSDIST(0)"), context);
    expect(legacyNormsDist).toMatchObject({ tag: ValueTag.Number });
    if (legacyNormsDist.tag !== ValueTag.Number) {
      throw new Error("LEGACY.NORMSDIST should return a number");
    }
    expect(legacyNormsDist.value).toBeCloseTo(0.5, 8);
    expect(evaluateAst(parseFormula("MissingFn(A1)"), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });
    expect(evaluateAst(parseFormula("A:A"), context)).toEqual({ tag: ValueTag.Empty });
  });

  it("evaluates GETPIVOTDATA through the pivot-data context hook", () => {
    const result = evaluateAst(parseFormula('GETPIVOTDATA("Sales",B2,"Region","East")'), {
      sheetName: "Sheet1",
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
      resolvePivotData: ({ dataField, sheetName, address, filters }) =>
        dataField === "Sales" &&
        sheetName === "Sheet1" &&
        address === "B2" &&
        filters.length === 1 &&
        filters[0]?.field === "Region" &&
        filters[0]?.item.tag === ValueTag.String &&
        filters[0].item.value === "East"
          ? { tag: ValueTag.Number, value: 15 }
          : { tag: ValueTag.Error, code: ErrorCode.Ref },
    });

    expect(result).toEqual({ tag: ValueTag.Number, value: 15 });
  });

  it("evaluates GROUPBY, PIVOTBY, and MULTIPLE.OPERATIONS workbook-shape formulas", () => {
    const matrixContext = {
      sheetName: "Sheet1",
      resolveCell: (_sheetName: string, address: string): CellValue =>
        address === "B5" ? { tag: ValueTag.Number, value: 0 } : { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === "A1" && end === "A5") {
          return [
            { tag: ValueTag.String, value: "Region", stringId: 0 },
            { tag: ValueTag.String, value: "East", stringId: 0 },
            { tag: ValueTag.String, value: "West", stringId: 0 },
            { tag: ValueTag.String, value: "East", stringId: 0 },
            { tag: ValueTag.String, value: "West", stringId: 0 },
          ];
        }
        if (start === "B1" && end === "B5") {
          return [
            { tag: ValueTag.String, value: "Product", stringId: 0 },
            { tag: ValueTag.String, value: "Widget", stringId: 0 },
            { tag: ValueTag.String, value: "Widget", stringId: 0 },
            { tag: ValueTag.String, value: "Gizmo", stringId: 0 },
            { tag: ValueTag.String, value: "Gizmo", stringId: 0 },
          ];
        }
        if (start === "C1" && end === "C5") {
          return [
            { tag: ValueTag.String, value: "Sales", stringId: 0 },
            { tag: ValueTag.Number, value: 10 },
            { tag: ValueTag.Number, value: 7 },
            { tag: ValueTag.Number, value: 5 },
            { tag: ValueTag.Number, value: 4 },
          ];
        }
        return [];
      },
      resolveMultipleOperations: ({
        formulaSheetName,
        formulaAddress,
        rowCellAddress,
        rowReplacementAddress,
        columnCellAddress,
        columnReplacementAddress,
      }: {
        formulaSheetName: string;
        formulaAddress: string;
        rowCellSheetName: string;
        rowCellAddress: string;
        rowReplacementSheetName: string;
        rowReplacementAddress: string;
        columnCellSheetName?: string;
        columnCellAddress?: string;
        columnReplacementSheetName?: string;
        columnReplacementAddress?: string;
      }) =>
        formulaSheetName === "Sheet1" &&
        formulaAddress === "B5" &&
        rowCellAddress === "B3" &&
        rowReplacementAddress === "C4" &&
        columnCellAddress === "B2" &&
        columnReplacementAddress === "D2"
          ? { tag: ValueTag.Number, value: 5 }
          : { tag: ValueTag.Error, code: ErrorCode.Ref },
    };

    const groupBy = evaluateAstResult(parseFormula("GROUPBY(A1:A5,C1:C5,SUM,3,1)"), matrixContext);
    expect(groupBy).toEqual({
      kind: "array",
      rows: 4,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: "Region", stringId: 0 },
        { tag: ValueTag.String, value: "Sales", stringId: 0 },
        { tag: ValueTag.String, value: "East", stringId: 0 },
        { tag: ValueTag.Number, value: 15 },
        { tag: ValueTag.String, value: "West", stringId: 0 },
        { tag: ValueTag.Number, value: 11 },
        { tag: ValueTag.String, value: "Total", stringId: 0 },
        { tag: ValueTag.Number, value: 26 },
      ],
    });

    const pivotBy = evaluateAstResult(
      parseFormula("PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)"),
      matrixContext,
    );
    expect(pivotBy).toEqual({
      kind: "array",
      rows: 4,
      cols: 4,
      values: [
        { tag: ValueTag.String, value: "Region", stringId: 0 },
        { tag: ValueTag.String, value: "Widget", stringId: 0 },
        { tag: ValueTag.String, value: "Gizmo", stringId: 0 },
        { tag: ValueTag.String, value: "Total", stringId: 0 },
        { tag: ValueTag.String, value: "East", stringId: 0 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 15 },
        { tag: ValueTag.String, value: "West", stringId: 0 },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 11 },
        { tag: ValueTag.String, value: "Total", stringId: 0 },
        { tag: ValueTag.Number, value: 17 },
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 26 },
      ],
    });

    expect(evaluateAst(parseFormula("MULTIPLE.OPERATIONS(B5,B3,C4,B2,D2)"), matrixContext)).toEqual(
      { tag: ValueTag.Number, value: 5 },
    );
  });

  it("resolves scalar defined names through the JS evaluator", () => {
    const value = evaluateAst(parseFormula("TaxRate*A1"), {
      sheetName: "Sheet1",
      resolveCell: (_sheetName: string, address: string): CellValue =>
        address === "A1" ? { tag: ValueTag.Number, value: 100 } : { tag: ValueTag.Empty },
      resolveRange: (): CellValue[] => [],
      resolveName: (name: string): CellValue =>
        name.toUpperCase() === "TAXRATE"
          ? { tag: ValueTag.Number, value: 0.085 }
          : { tag: ValueTag.Error, code: ErrorCode.Name },
    });

    expect(value).toEqual({ tag: ValueTag.Number, value: 8.5 });
  });

  it("preserves sequence array results while flattening them for scalar consumers", () => {
    const context = {
      sheetName: "Sheet1",
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    };

    expect(evaluateAstResult(parseFormula("SEQUENCE(3,1,1,1)"), context)).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ],
    });
    expect(evaluateAst(parseFormula("SUM(SEQUENCE(3,1,1,1))"), context)).toEqual({
      tag: ValueTag.Number,
      value: 6,
    });

    expect(
      evaluateAstResult(parseFormula("FILTER(A1:A4,A1:A4>2)"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ],
    });

    expect(
      evaluateAstResult(parseFormula("UNIQUE(A1:A4)"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.String, value: "A", stringId: 0 },
          { tag: ValueTag.String, value: "B", stringId: 0 },
          { tag: ValueTag.String, value: "A", stringId: 0 },
          { tag: ValueTag.String, value: "C", stringId: 0 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: "A", stringId: 0 },
        { tag: ValueTag.String, value: "B", stringId: 0 },
        { tag: ValueTag.String, value: "C", stringId: 0 },
      ],
    });

    expect(
      evaluateAstResult(parseFormula('REGEXEXTRACT("a1 b2 c3","[a-z][0-9]",1)'), context),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: "a1", stringId: 0 },
        { tag: ValueTag.String, value: "b2", stringId: 0 },
        { tag: ValueTag.String, value: "c3", stringId: 0 },
      ],
    });

    expect(
      evaluateAstResult(parseFormula('REGEXEXTRACT("abc-123","([a-z]+)-([0-9]+)",2)'), context),
    ).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: "abc", stringId: 0 },
        { tag: ValueTag.String, value: "123", stringId: 0 },
      ],
    });

    expect(evaluateAstResult(parseFormula('TEXTSPLIT("red,blue|green",",","|")'), context)).toEqual(
      {
        kind: "array",
        rows: 2,
        cols: 2,
        values: [
          { tag: ValueTag.String, value: "red", stringId: 0 },
          { tag: ValueTag.String, value: "blue", stringId: 0 },
          { tag: ValueTag.String, value: "green", stringId: 0 },
          { tag: ValueTag.Error, code: ErrorCode.NA },
        ],
      },
    );

    expect(
      evaluateAstResult(parseFormula("EXPAND(A1:A3,4,2,0)"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 2 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 4,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ],
    });

    expect(
      evaluateAstResult(parseFormula("TRIMRANGE(A1:D4)"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Empty },
      ],
    });

    expect(
      evaluateAstResult(parseFormula('INDIRECT("A1:A3")'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 7 },
          { tag: ValueTag.Number, value: 8 },
          { tag: ValueTag.Number, value: 9 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 9 },
      ],
    });

    expect(evaluateAstResult(parseFormula("MAKEARRAY(2,2,LAMBDA(r,c,r+c))"), context)).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ],
    });

    expect(evaluateAstResult(parseFormula("MUNIT(3)"), context)).toEqual({
      kind: "array",
      rows: 3,
      cols: 3,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    });

    expect(
      evaluateAstResult(parseFormula("MMULT(A1:B2,C1:D2)"), {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
          if (start === "A1" && end === "B2") {
            return [
              { tag: ValueTag.Number, value: 1 },
              { tag: ValueTag.Number, value: 2 },
              { tag: ValueTag.Number, value: 3 },
              { tag: ValueTag.Number, value: 4 },
            ];
          }
          return [
            { tag: ValueTag.Number, value: 5 },
            { tag: ValueTag.Number, value: 6 },
            { tag: ValueTag.Number, value: 7 },
            { tag: ValueTag.Number, value: 8 },
          ];
        },
      }),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 19 },
        { tag: ValueTag.Number, value: 22 },
        { tag: ValueTag.Number, value: 43 },
        { tag: ValueTag.Number, value: 50 },
      ],
    });

    expect(
      evaluateAst(parseFormula("SUMXMY2(A1:A2,B1:B2)"), {
        ...context,
        resolveRange: (_sheetName: string, start: string, _end: string): CellValue[] =>
          start === "A1"
            ? [
                { tag: ValueTag.Number, value: 1 },
                { tag: ValueTag.Number, value: 2 },
              ]
            : [
                { tag: ValueTag.Number, value: 3 },
                { tag: ValueTag.Number, value: 4 },
              ],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 8 });

    const randomGrid = evaluateAstResult(parseFormula("RANDARRAY(2,2,3,7,TRUE())"), context);
    expect(randomGrid).toMatchObject({ kind: "array", rows: 2, cols: 2 });
    if (randomGrid.kind !== "array") {
      throw new Error("expected RANDARRAY to return an array");
    }
    for (const value of randomGrid.values) {
      expect(value.tag).toBe(ValueTag.Number);
      expect(value.value).toBeGreaterThanOrEqual(3);
      expect(value.value).toBeLessThanOrEqual(7);
    }

    expect(
      evaluateAstResult(parseFormula("MAP(A1:A3,LAMBDA(x,x*2))"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
      ],
    });

    expect(
      evaluateAst(parseFormula("REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
        ],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 6 });

    expect(
      evaluateAstResult(parseFormula("SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 6 },
      ],
    });

    expect(
      evaluateAstResult(parseFormula("BYROW(A1:B2,LAMBDA(r,SUM(r)))"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 7 },
      ],
    });

    expect(
      evaluateAstResult(parseFormula("BYCOL(A1:B2,LAMBDA(c,SUM(c)))"), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    ).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
      ],
    });
  });
});
