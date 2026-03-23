import React from "react";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { buildDemoWorkbook } from "./demoWorkbook.js";

export type PlaygroundPresetId =
  | "starter"
  | "load-100k"
  | "load-250k"
  | "downstream-10k"
  | "range-aggregates"
  | "million-surface";

export interface PlaygroundPresetDefinition {
  id: PlaygroundPresetId;
  label: string;
  description: string;
  kind: "renderer" | "snapshot";
}

export type PlaygroundPresetPayload =
  | {
      kind: "renderer";
      element: React.ReactNode;
      defaultSheet: string;
      defaultAddress: string;
    }
  | {
      kind: "snapshot";
      snapshot: WorkbookSnapshot;
      defaultSheet: string;
      defaultAddress: string;
    };

export const PLAYGROUND_PRESETS: readonly PlaygroundPresetDefinition[] = [
  {
    id: "starter",
    label: "Starter Demo",
    description: "Small workbook mounted through the custom reconciler.",
    kind: "renderer",
  },
  {
    id: "load-100k",
    label: "100k Materialized",
    description: "Large load fixture with 100,000 materialized cells.",
    kind: "snapshot",
  },
  {
    id: "load-250k",
    label: "250k Materialized",
    description: "Stress fixture with 250,000 materialized cells.",
    kind: "snapshot",
  },
  {
    id: "downstream-10k",
    label: "10k Downstream Recalc",
    description: "One edit fans out into 10,000 dependent formulas.",
    kind: "snapshot",
  },
  {
    id: "range-aggregates",
    label: "Range Aggregates",
    description: "Repeated bounded aggregate formulas against shared ranges.",
    kind: "snapshot",
  },
  {
    id: "million-surface",
    label: "Million-Row Surface",
    description: "Excel-scale surface with landmarks at the far row and far column edges.",
    kind: "snapshot",
  },
] as const;

export async function loadPlaygroundPreset(
  id: PlaygroundPresetId,
): Promise<PlaygroundPresetPayload> {
  switch (id) {
    case "starter":
      return {
        kind: "renderer",
        element: buildDemoWorkbook(),
        defaultSheet: "Sheet1",
        defaultAddress: "A1",
      };
    case "load-100k": {
      const { buildWorkbookSnapshot } = await import("@bilig/benchmarks/generate-workbook");
      return {
        kind: "snapshot",
        snapshot: buildWorkbookSnapshot(100_000),
        defaultSheet: "Sheet1",
        defaultAddress: "A1",
      };
    }
    case "load-250k": {
      const { buildWorkbookSnapshot } = await import("@bilig/benchmarks/generate-workbook");
      return {
        kind: "snapshot",
        snapshot: buildWorkbookSnapshot(250_000),
        defaultSheet: "Sheet1",
        defaultAddress: "A1",
      };
    }
    case "downstream-10k": {
      const { buildDownstreamSnapshot } = await import("@bilig/benchmarks/generate-workbook");
      return {
        kind: "snapshot",
        snapshot: buildDownstreamSnapshot(10_000),
        defaultSheet: "Sheet1",
        defaultAddress: "A1",
      };
    }
    case "range-aggregates": {
      const { buildRangeAggregateSnapshot } = await import("@bilig/benchmarks/generate-workbook");
      return {
        kind: "snapshot",
        snapshot: buildRangeAggregateSnapshot(2_048, 10_000),
        defaultSheet: "Sheet1",
        defaultAddress: "B1",
      };
    }
    case "million-surface":
      return {
        kind: "snapshot",
        snapshot: buildMillionSurfaceSnapshot(),
        defaultSheet: "Sheet1",
        defaultAddress: "A1",
      };
  }
}

function buildMillionSurfaceSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: "million-row-surface" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: [
          { address: "A1", value: "Top left" },
          { address: "B1", formula: "1+1" },
          { address: "A1048576", value: 1048576 },
          { address: "B1048576", formula: "A1048576*2" },
          { address: "XFD1", value: "Far column" },
          { address: "XFD1048576", formula: "B1048576+1" },
        ],
      },
      {
        name: "Landmarks",
        order: 1,
        cells: [
          { address: "A1", value: "Jump targets" },
          { address: "A2", value: "A1048576" },
          { address: "A3", value: "XFD1" },
          { address: "A4", value: "XFD1048576" },
          { address: "B2", formula: "Sheet1!A1048576" },
          { address: "B4", formula: "Sheet1!XFD1048576" },
        ],
      },
    ],
  };
}
