import { describe, expect, it } from "vitest";
import {
  createSpreadsheetEngineFromDocument,
  exportSpreadsheetEngineDocument,
  importSpreadsheetEngineDocument,
  isPersistedSpreadsheetEngineDocument,
  parseSpreadsheetEngineDocument,
  SPREADSHEET_ENGINE_DOCUMENT_FORMAT,
  SpreadsheetEngine,
  serializeSpreadsheetEngineDocument,
} from "../index.js";

describe("spreadsheet engine persistence helpers", () => {
  it("roundtrips workbook and replica state through the persisted document format", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "persistence-roundtrip",
      replicaId: "persistence-primary",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 21);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setWorkbookMetadata("locale", "en-US");
    engine.setCalculationSettings({ mode: "manual", compatibilityMode: "excel-modern" });
    engine.setFreezePane("Sheet1", 1, 1);

    const document = exportSpreadsheetEngineDocument(engine);
    const serialized = serializeSpreadsheetEngineDocument(document);
    const parsed = parseSpreadsheetEngineDocument(serialized);
    const restored = await createSpreadsheetEngineFromDocument(parsed, {
      replicaId: "persistence-restored",
    });

    expect(parsed.format).toBe(SPREADSHEET_ENGINE_DOCUMENT_FORMAT);
    expect(isPersistedSpreadsheetEngineDocument(parsed)).toBe(true);
    expect(restored.exportSnapshot()).toEqual(engine.exportSnapshot());
    expect(restored.exportReplicaSnapshot()).toEqual(engine.exportReplicaSnapshot());
  });

  it("imports a persisted document into an existing engine instance", async () => {
    const source = new SpreadsheetEngine({
      workbookName: "persistence-import-source",
      replicaId: "persistence-import-source",
    });
    await source.ready();
    source.createSheet("Source");
    source.setCellValue("Source", "C3", 99);
    source.setCellFormula("Source", "D4", "C3+1");

    const target = new SpreadsheetEngine({
      workbookName: "persistence-import-target",
      replicaId: "persistence-import-target",
    });
    await target.ready();
    target.createSheet("Scratch");
    target.setCellValue("Scratch", "A1", 1);

    importSpreadsheetEngineDocument(target, exportSpreadsheetEngineDocument(source));

    expect(target.exportSnapshot()).toEqual(source.exportSnapshot());
    expect(target.exportReplicaSnapshot()).toEqual(source.exportReplicaSnapshot());
  });

  it("supports snapshot-only persistence documents", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "persistence-snapshot-only" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", "north");

    const document = exportSpreadsheetEngineDocument(engine, { includeReplica: false });
    const restored = await createSpreadsheetEngineFromDocument(document, {
      replicaId: "snapshot-only-restored",
    });

    expect(document.replica).toBeUndefined();
    expect(restored.exportSnapshot()).toEqual(engine.exportSnapshot());
  });

  it("rejects invalid persisted documents", () => {
    expect(isPersistedSpreadsheetEngineDocument({})).toBe(false);
    expect(() => parseSpreadsheetEngineDocument("{}")).toThrow(
      "Invalid persisted spreadsheet engine document",
    );
    expect(
      isPersistedSpreadsheetEngineDocument({
        format: SPREADSHEET_ENGINE_DOCUMENT_FORMAT,
        snapshot: { version: 2, workbook: { name: "Broken" }, sheets: [] },
      }),
    ).toBe(false);
  });
});
