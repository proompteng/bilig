import { readFile } from "node:fs/promises";

const coveragePath = new URL("../coverage/coverage-final.json", import.meta.url);

const thresholds = [
  { label: "packages/core/src", prefix: "/packages/core/src/", lines: 90 },
  { label: "packages/formula/src", prefix: "/packages/formula/src/", lines: 90 },
  { label: "packages/renderer/src", prefix: "/packages/renderer/src/", lines: 90 }
];

const ignoredSuffixes = ["/index.ts", "/snapshot.ts", "/ast.ts"];
const coverage = JSON.parse(await readFile(coveragePath, "utf8"));

function lineStatsForFile(fileCoverage) {
  const totalLines = new Set();
  const coveredLines = new Set();

  for (const [statementId, location] of Object.entries(fileCoverage.statementMap)) {
    const hits = fileCoverage.s[statementId] ?? 0;
    for (let line = location.start.line; line <= location.end.line; line += 1) {
      totalLines.add(line);
      if (hits > 0) {
        coveredLines.add(line);
      }
    }
  }

  return {
    total: totalLines.size,
    covered: coveredLines.size
  };
}

function aggregatePrefix(prefix) {
  let total = 0;
  let covered = 0;

  for (const [filePath, fileCoverage] of Object.entries(coverage)) {
    if (!filePath.includes(prefix)) {
      continue;
    }
    if (ignoredSuffixes.some((suffix) => filePath.endsWith(suffix))) {
      continue;
    }
    const stats = lineStatsForFile(fileCoverage);
    total += stats.total;
    covered += stats.covered;
  }

  if (total === 0) {
    throw new Error(`No coverage entries found for ${prefix}`);
  }

  return (covered / total) * 100;
}

const results = thresholds.map((threshold) => ({
  label: threshold.label,
  linesPct: aggregatePrefix(threshold.prefix),
  requiredLinesPct: threshold.lines
}));

for (const result of results) {
  if (result.linesPct < result.requiredLinesPct) {
    throw new Error(
      `${result.label} line coverage is below target: ${result.linesPct.toFixed(2)}% < ${result.requiredLinesPct}%`
    );
  }
}

console.log(JSON.stringify({ results }, null, 2));
