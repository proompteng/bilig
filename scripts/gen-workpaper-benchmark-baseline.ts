#!/usr/bin/env bun

import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import {
  runWorkPaperBenchmarkSuite,
  type WorkPaperBenchmarkResult,
} from "../packages/benchmarks/src/benchmark-workpaper.ts";

interface WorkPaperBenchmarkBaselineDocument {
  schemaVersion: 1;
  suite: "workpaper";
  generatedAt: string;
  host: {
    arch: string;
    nodeVersion: string;
    platform: string;
  };
  results: WorkPaperBenchmarkResult[];
}

interface WorkPaperBenchmarkBaselineShapeInput {
  schemaVersion: 1;
  suite: "workpaper";
  results: Array<{
    details: Record<string, unknown>;
    metrics?: Record<string, unknown>;
    scenario: string;
  }>;
}

const rootDir = resolve(new URL("..", import.meta.url).pathname);
const outputPath = join(rootDir, "packages", "benchmarks", "baselines", "workpaper-baseline.json");
const isCheckMode = process.argv.slice(2).includes("--check");

const baseline: WorkPaperBenchmarkBaselineDocument = {
  schemaVersion: 1,
  suite: "workpaper",
  generatedAt: new Date().toISOString(),
  host: {
    arch: process.arch,
    nodeVersion: process.version,
    platform: process.platform,
  },
  results: await runWorkPaperBenchmarkSuite(),
};

const serializedBaseline = formatJsonForRepo(`${JSON.stringify(baseline, null, 2)}\n`);

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error(
      `WorkPaper benchmark baseline is missing. Run: bun scripts/gen-workpaper-benchmark-baseline.ts`,
    );
  }

  const existing = parseBaselineForShape(readFileSync(outputPath, "utf8"));
  const expectedShape = normalizeBaselineShape(baseline);
  const actualShape = normalizeBaselineShape(existing);
  if (JSON.stringify(actualShape) !== JSON.stringify(expectedShape)) {
    throw new Error(
      `WorkPaper benchmark baseline shape is out of date. Run: bun scripts/gen-workpaper-benchmark-baseline.ts`,
    );
  }

  console.log(
    JSON.stringify(
      {
        outputPath,
        mode: "check",
        scenarios: actualShape.scenarios.map((scenario) => scenario.scenario),
      },
      null,
      2,
    ),
  );
  process.exit(0);
}

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, serializedBaseline);
console.log(
  JSON.stringify(
    {
      outputPath,
      mode: "write",
      scenarios: baseline.results.map((scenario) => scenario.scenario),
    },
    null,
    2,
  ),
);

function parseBaselineForShape(serialized: string): WorkPaperBenchmarkBaselineShapeInput {
  const candidate = toRecord(JSON.parse(serialized), `WorkPaper benchmark baseline ${outputPath}`);
  if (candidate.schemaVersion !== 1 || candidate.suite !== "workpaper") {
    throw new Error(`Unexpected WorkPaper benchmark baseline header: ${outputPath}`);
  }

  const results = candidate.results;
  if (!Array.isArray(results)) {
    throw new Error(`WorkPaper benchmark baseline is missing results: ${outputPath}`);
  }

  return {
    schemaVersion: 1,
    suite: "workpaper",
    results: results.map((result, index) => {
      const record = toRecord(result, `WorkPaper benchmark baseline result ${index}`);
      const scenario = record.scenario;
      if (typeof scenario !== "string") {
        throw new Error(`WorkPaper benchmark baseline result ${index} is missing scenario`);
      }

      return {
        scenario,
        details: toRecord(record.details, `WorkPaper benchmark baseline details ${index}`),
        metrics:
          record.metrics === undefined
            ? undefined
            : toRecord(record.metrics, `WorkPaper benchmark baseline metrics ${index}`),
      };
    }),
  };
}

function normalizeBaselineShape(
  baselineDocument: WorkPaperBenchmarkBaselineShapeInput | WorkPaperBenchmarkBaselineDocument,
): {
  schemaVersion: 1;
  suite: "workpaper";
  scenarios: Array<{
    detailKeys: string[];
    detailTypes: Record<string, string>;
    hasMetrics: boolean;
    metricKeys: string[];
    scenario: string;
  }>;
} {
  return {
    schemaVersion: 1,
    suite: "workpaper",
    scenarios: baselineDocument.results.map((result) => ({
      scenario: result.scenario,
      detailKeys: Object.keys(result.details).toSorted(),
      detailTypes: Object.fromEntries(
        Object.entries(result.details)
          .map(([key, value]) => [key, typeof value] as const)
          .toSorted(([left], [right]) => left.localeCompare(right)),
      ),
      hasMetrics: result.metrics !== undefined,
      metricKeys: result.metrics ? Object.keys(result.metrics).toSorted() : [],
    })),
  };
}

function toRecord(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`Expected ${context} to be an object`);
  }
  const record: Record<string, unknown> = {};
  for (const [key, entryValue] of Object.entries(value)) {
    record[key] = entryValue;
  }
  return record;
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), "workpaper-bench-baseline-"));
  const tempFilePath = join(tempDir, "baseline.json");
  writeFileSync(tempFilePath, serializedJson);
  const oxfmtPath = join(rootDir, "node_modules", ".bin", "oxfmt");

  const formatResult = Bun.spawnSync([oxfmtPath, "--write", tempFilePath], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
  });
  if (formatResult.exitCode !== 0) {
    const stderr = new TextDecoder().decode(formatResult.stderr).trim();
    rmSync(tempDir, { recursive: true, force: true });
    throw new Error(`Unable to format generated WorkPaper benchmark baseline: ${stderr}`);
  }

  const formattedJson = readFileSync(tempFilePath, "utf8");
  rmSync(tempDir, { recursive: true, force: true });
  return formattedJson;
}
