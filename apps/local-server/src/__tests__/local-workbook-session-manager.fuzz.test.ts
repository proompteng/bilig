import { describe, expect, it } from "vitest";
import fc, { type AsyncCommand } from "fast-check";
import { SpreadsheetEngine } from "@bilig/core";
import { formatAddress } from "@bilig/formula";
import type { AgentFrame, AgentRequest, AgentResponse } from "@bilig/agent-api";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
} from "@bilig/protocol";
import { LocalWorkbookSessionManager } from "../local-workbook-session-manager.js";
import { runModelProperty, shouldRunFuzzSuite } from "@bilig/test-fuzz";

const documentId = "local-fuzz-book";
const replicaId = "local-fuzz-agent";
const sheetName = "Sheet1";

interface LocalModel {
  engine: SpreadsheetEngine;
}

interface LocalReal {
  manager: LocalWorkbookSessionManager;
  sessionId: string;
}

type LocalCommand = AsyncCommand<LocalModel, LocalReal, boolean>;

function toRangeRef(
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function buildValueMatrix(
  height: number,
  width: number,
  values: readonly LiteralInput[],
): LiteralInput[][] {
  const rows: LiteralInput[][] = [];
  let offset = 0;
  for (let row = 0; row < height; row += 1) {
    const nextRow: LiteralInput[] = [];
    for (let col = 0; col < width; col += 1) {
      nextRow.push(values[offset] ?? null);
      offset += 1;
    }
    rows.push(nextRow);
  }
  return rows;
}

async function sendAgentRequest(
  manager: LocalWorkbookSessionManager,
  request: AgentRequest,
): Promise<AgentResponse> {
  const frame: AgentFrame = { kind: "request", request };
  const response = await manager.handleAgentFrame(frame);
  if (response.kind !== "response") {
    throw new Error(`Expected response frame, received ${response.kind}`);
  }
  if (response.response.kind === "error") {
    throw new Error(`${response.response.code}: ${response.response.message}`);
  }
  return response.response;
}

async function exportManagerSnapshot(real: LocalReal) {
  const response = await sendAgentRequest(real.manager, {
    kind: "exportSnapshot",
    id: "export-snapshot",
    sessionId: real.sessionId,
  });
  if (response.kind !== "snapshot") {
    throw new Error(`Expected snapshot response, received ${response.kind}`);
  }
  return response.snapshot;
}

async function expectSnapshotsToMatch(model: LocalModel, real: LocalReal) {
  expect(await exportManagerSnapshot(real)).toEqual(model.engine.exportSnapshot());
}

function writeRangeCommand(range: CellRangeRef, values: LiteralInput[][]): LocalCommand {
  return {
    check: () => true,
    async run(model, real) {
      model.engine.setRangeValues(range, values);
      await sendAgentRequest(real.manager, {
        kind: "writeRange",
        id: `write:${range.startAddress}:${range.endAddress}`,
        sessionId: real.sessionId,
        range,
        values,
      });
      await expectSnapshotsToMatch(model, real);
    },
    toString() {
      return `writeRange(${range.startAddress}:${range.endAddress})`;
    },
  };
}

function formulaRangeCommand(range: CellRangeRef, formulas: string[][]): LocalCommand {
  return {
    check: () => true,
    async run(model, real) {
      model.engine.setRangeFormulas(range, formulas);
      await sendAgentRequest(real.manager, {
        kind: "setRangeFormulas",
        id: `formula:${range.startAddress}:${range.endAddress}`,
        sessionId: real.sessionId,
        range,
        formulas,
      });
      await expectSnapshotsToMatch(model, real);
    },
    toString() {
      return `setRangeFormulas(${range.startAddress}:${range.endAddress})`;
    },
  };
}

function styleRangeCommand(range: CellRangeRef, patch: CellStylePatch): LocalCommand {
  return {
    check: () => true,
    async run(model, real) {
      model.engine.setRangeStyle(range, patch);
      await sendAgentRequest(real.manager, {
        kind: "setRangeStyle",
        id: `style:${range.startAddress}:${range.endAddress}`,
        sessionId: real.sessionId,
        range,
        patch,
      });
      await expectSnapshotsToMatch(model, real);
    },
    toString() {
      return `setRangeStyle(${range.startAddress}:${range.endAddress})`;
    },
  };
}

function formatRangeCommand(range: CellRangeRef, format: CellNumberFormatInput): LocalCommand {
  return {
    check: () => true,
    async run(model, real) {
      model.engine.setRangeNumberFormat(range, format);
      await sendAgentRequest(real.manager, {
        kind: "setRangeNumberFormat",
        id: `format:${range.startAddress}:${range.endAddress}`,
        sessionId: real.sessionId,
        range,
        format,
      });
      await expectSnapshotsToMatch(model, real);
    },
    toString() {
      return `setRangeNumberFormat(${range.startAddress}:${range.endAddress})`;
    },
  };
}

function clearRangeCommand(range: CellRangeRef): LocalCommand {
  return {
    check: () => true,
    async run(model, real) {
      model.engine.clearRange(range);
      await sendAgentRequest(real.manager, {
        kind: "clearRange",
        id: `clear:${range.startAddress}:${range.endAddress}`,
        sessionId: real.sessionId,
        range,
      });
      await expectSnapshotsToMatch(model, real);
    },
    toString() {
      return `clearRange(${range.startAddress}:${range.endAddress})`;
    },
  };
}

function fillRangeCommand(source: CellRangeRef, target: CellRangeRef): LocalCommand {
  return {
    check: () => true,
    async run(model, real) {
      model.engine.fillRange(source, target);
      await sendAgentRequest(real.manager, {
        kind: "fillRange",
        id: `fill:${source.startAddress}:${target.endAddress}`,
        sessionId: real.sessionId,
        source,
        target,
      });
      await expectSnapshotsToMatch(model, real);
    },
    toString() {
      return `fillRange(${source.startAddress}:${source.endAddress} -> ${target.startAddress}:${target.endAddress})`;
    },
  };
}

const literalInputArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -10_000, max: 10_000 }),
  fc.boolean(),
  fc.constantFrom("north", "ready", "done"),
  fc.constant(null),
);
const rangeSeedArbitrary = fc.record({
  startRow: fc.integer({ min: 0, max: 4 }),
  startCol: fc.integer({ min: 0, max: 4 }),
  height: fc.integer({ min: 1, max: 2 }),
  width: fc.integer({ min: 1, max: 2 }),
});
const rangeArbitrary = rangeSeedArbitrary.map((value) =>
  toRangeRef(
    value.startRow,
    value.startCol,
    value.startRow + value.height - 1,
    value.startCol + value.width - 1,
  ),
);
const writeRangeCommandArbitrary = rangeSeedArbitrary.chain((range) =>
  fc
    .array(literalInputArbitrary, {
      minLength: range.height * range.width,
      maxLength: range.height * range.width,
    })
    .map((values) =>
      writeRangeCommand(
        toRangeRef(
          range.startRow,
          range.startCol,
          range.startRow + range.height - 1,
          range.startCol + range.width - 1,
        ),
        buildValueMatrix(range.height, range.width, values),
      ),
    ),
);
const formulaCommandArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 4 }),
    col: fc.integer({ min: 0, max: 4 }),
    formula: fc
      .tuple(
        fc.constantFrom("C3", "C4", "D3", "D4", "E5"),
        fc.constantFrom("+", "-", "*"),
        fc.constantFrom("C3", "C4", "D3", "D4", "E5"),
      )
      .map(([left, operator, right]) => `${left}${operator}${right}`),
  })
  .map(({ row, col, formula }) => formulaRangeCommand(toRangeRef(row, col, row, col), [[formula]]));
const styleCommandArbitrary = fc
  .record({
    range: rangeArbitrary,
    patch: fc.constantFrom<CellStylePatch>(
      { fill: { backgroundColor: "#dbeafe" } },
      { font: { bold: true } },
      { alignment: { horizontal: "right", wrap: true } },
    ),
  })
  .map(({ range, patch }) => styleRangeCommand(range, patch));
const formatCommandArbitrary = fc
  .record({
    range: rangeArbitrary,
    format: fc.constantFrom<CellNumberFormatInput>(
      "0.00",
      { kind: "currency", currency: "USD", decimals: 2 },
      { kind: "percent", decimals: 1 },
      { kind: "text" },
    ),
  })
  .map(({ range, format }) => formatRangeCommand(range, format));
const clearCommandArbitrary = rangeArbitrary.map((range) => clearRangeCommand(range));
const fillCommandArbitrary = rangeSeedArbitrary.chain((source) =>
  fc
    .record({
      targetStartRow: fc.integer({ min: source.startRow, max: 4 }),
      targetStartCol: fc.integer({ min: source.startCol, max: 4 }),
    })
    .map(({ targetStartRow, targetStartCol }) =>
      fillRangeCommand(
        toRangeRef(
          source.startRow,
          source.startCol,
          source.startRow + source.height - 1,
          source.startCol + source.width - 1,
        ),
        toRangeRef(
          targetStartRow,
          targetStartCol,
          Math.min(4, targetStartRow + source.height - 1),
          Math.min(4, targetStartCol + source.width - 1),
        ),
      ),
    ),
);

describe("local workbook session manager fuzz", () => {
  it("keeps the manager snapshot converged with direct engine execution", async () => {
    const suite = "local-server/session-manager/differential";
    const executed = runModelProperty({
      suite,
      commands: fc.commands(
        [
          writeRangeCommandArbitrary,
          formulaCommandArbitrary,
          styleCommandArbitrary,
          formatCommandArbitrary,
          clearCommandArbitrary,
          fillCommandArbitrary,
        ],
        { maxCommands: 8 },
      ),
      createModel: () => {
        const engine = new SpreadsheetEngine({
          workbookName: documentId,
          replicaId: "local-fuzz-model",
        });
        engine.createSheet(sheetName);
        return { engine };
      },
      createReal: async () => {
        const manager = new LocalWorkbookSessionManager();
        const response = await sendAgentRequest(manager, {
          kind: "openWorkbookSession",
          id: "open-session",
          documentId,
          replicaId,
        });
        if (response.kind !== "ok" || !response.sessionId) {
          throw new Error("Expected openWorkbookSession to return a session id");
        }
        return {
          manager,
          sessionId: response.sessionId,
        };
      },
      teardown: async (real) => {
        await sendAgentRequest(real.manager, {
          kind: "closeWorkbookSession",
          id: "close-session",
          sessionId: real.sessionId,
        });
      },
    });

    await expect(executed).resolves.toBe(shouldRunFuzzSuite(suite, "model"));
  });
});
