import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "@bilig/core";
import { buildRenderCommitOps } from "./generate-workbook.js";

export interface RenderCommitBenchmarkResult {
  scenario: "render-commit";
  declaredCells: number;
  elapsedMs: number;
}

export async function runRenderCommitBenchmark(declaredCells = 1_000): Promise<RenderCommitBenchmarkResult> {
  const engine = new SpreadsheetEngine({ workbookName: "benchmark-render-commit" });
  await engine.ready();

  const started = performance.now();
  engine.renderCommit(buildRenderCommitOps(declaredCells));
  const elapsed = performance.now() - started;

  return {
    scenario: "render-commit",
    declaredCells,
    elapsedMs: elapsed
  };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const declaredCells = Number.parseInt(process.argv[2] ?? "1000", 10);
  console.log(JSON.stringify(await runRenderCommitBenchmark(declaredCells), null, 2));
}
