import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import fc, {
  type AsyncCommand,
  type IAsyncProperty,
  type IProperty,
  type Parameters,
  type RunDetails,
  type Scheduler,
} from "fast-check";

export type FuzzProfile = "default" | "main" | "nightly" | "replay";
export type FuzzSuiteKind = "property" | "model" | "scheduled" | "browser";

interface BudgetProfile {
  property: { numRuns: number; maxMs: number };
  model: { numRuns: number; maxMs: number };
  scheduled: { numRuns: number; maxMs: number };
  browser: { numRuns: number; maxMs: number };
}

const BUDGETS: Record<Exclude<FuzzProfile, "replay">, BudgetProfile> = {
  default: {
    property: { numRuns: 50, maxMs: 10_000 },
    model: { numRuns: 10, maxMs: 15_000 },
    scheduled: { numRuns: 10, maxMs: 15_000 },
    browser: { numRuns: 5, maxMs: 20_000 },
  },
  main: {
    property: { numRuns: 200, maxMs: 30_000 },
    model: { numRuns: 40, maxMs: 30_000 },
    scheduled: { numRuns: 25, maxMs: 30_000 },
    browser: { numRuns: 15, maxMs: 30_000 },
  },
  nightly: {
    property: { numRuns: 2_000, maxMs: 120_000 },
    model: { numRuns: 200, maxMs: 120_000 },
    scheduled: { numRuns: 150, maxMs: 120_000 },
    browser: { numRuns: 75, maxMs: 120_000 },
  },
};

type FuzzParameters<Ts extends unknown[] = unknown[]> = Parameters<Ts>;

export interface ReplayFixture {
  suite: string;
  kind?: string;
  seed: number;
  path?: string;
  numRuns?: number;
  counterexample?: unknown;
  failures?: unknown;
  reproductionCommand?: string;
}

export interface ReplaySelector {
  enabled: boolean;
  suite: string | null;
  kind: string | null;
  filePath: string | null;
}

export interface PromoteCapturedArtifactOptions {
  artifactPath: string;
  fixturePath: string;
  metadata?: Record<string, unknown>;
}

export interface CaptureCounterexampleOptions<Ts extends unknown[] = unknown[]> {
  suite: string;
  kind: FuzzSuiteKind;
  details: RunDetails<Ts>;
}

export interface PropertySuiteOptions<T> {
  suite: string;
  arbitrary: fc.Arbitrary<T>;
  predicate: (value: T) => void | Promise<void>;
  kind?: FuzzSuiteKind;
  parameters?: FuzzParameters;
}

export interface ModelSuiteOptions<Model extends object, Real> {
  suite: string;
  commands: fc.Arbitrary<readonly AsyncCommand<Model, Real, boolean>[]>;
  createModel: () => Model;
  createReal: () => Real | Promise<Real>;
  teardown?: (real: Real) => void | Promise<void>;
  parameters?: FuzzParameters;
}

export interface ScheduledSuiteOptions<T> {
  suite: string;
  arbitrary: fc.Arbitrary<T>;
  predicate: (context: { scheduler: Scheduler; value: T }) => void | Promise<void>;
  kind?: FuzzSuiteKind;
  parameters?: FuzzParameters;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function parsePositiveInteger(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

function parseInteger(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function serializeArtifactValue(value: unknown): unknown {
  if (value instanceof Uint8Array) {
    return { type: "Uint8Array", values: [...value] };
  }
  if (value instanceof Error) {
    return {
      name: value.name,
      message: value.message,
      stack: value.stack ?? null,
    };
  }
  if (typeof value === "bigint") {
    return `${value}n`;
  }
  if (Array.isArray(value)) {
    return value.map((entry) => serializeArtifactValue(entry));
  }
  if (isRecord(value)) {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [key, serializeArtifactValue(entry)]),
    );
  }
  return value;
}

function resolveFuzzProfile(): FuzzProfile {
  const raw = process.env["BILIG_FUZZ_PROFILE"];
  if (raw === "main" || raw === "nightly" || raw === "replay") {
    return raw;
  }
  return "default";
}

function resolveBudget(kind: FuzzSuiteKind): { numRuns: number; maxMs: number } {
  const profile = resolveFuzzProfile();
  const baseProfile = profile === "replay" ? "main" : profile;
  return BUDGETS[baseProfile][kind === "browser" ? "browser" : kind];
}

function resolveReplayFixturePath(): string | null {
  return process.env["BILIG_FUZZ_REPLAY"] ? resolve(process.env["BILIG_FUZZ_REPLAY"]) : null;
}

export function loadReplayFixture(filePath: string): ReplayFixture {
  const resolvedPath = resolve(filePath);
  const raw = JSON.parse(readFileSync(resolvedPath, "utf8")) as unknown;
  if (!isRecord(raw) || typeof raw["suite"] !== "string" || typeof raw["seed"] !== "number") {
    throw new Error(`Invalid replay fixture: ${resolvedPath}`);
  }
  const fixture: ReplayFixture = {
    suite: raw["suite"],
    seed: raw["seed"],
  };
  if (typeof raw["kind"] === "string") {
    fixture.kind = raw["kind"];
  }
  if (typeof raw["path"] === "string") {
    fixture.path = raw["path"];
  }
  if (typeof raw["numRuns"] === "number") {
    fixture.numRuns = raw["numRuns"];
  }
  if ("counterexample" in raw) {
    fixture.counterexample = raw["counterexample"];
  }
  if ("failures" in raw) {
    fixture.failures = raw["failures"];
  }
  if (typeof raw["reproductionCommand"] === "string") {
    fixture.reproductionCommand = raw["reproductionCommand"];
  }
  return fixture;
}

export function promoteCapturedArtifact(options: PromoteCapturedArtifactOptions): string {
  const artifactPath = resolve(options.artifactPath);
  const fixturePath = resolve(options.fixturePath);
  const artifact = JSON.parse(readFileSync(artifactPath, "utf8")) as unknown;
  if (
    !isRecord(artifact) ||
    typeof artifact["suite"] !== "string" ||
    typeof artifact["seed"] !== "number"
  ) {
    throw new Error(`Invalid captured fuzz artifact: ${artifactPath}`);
  }
  const payload = {
    ...artifact,
    ...options.metadata,
    sourceArtifact: artifactPath,
    promotedAt: new Date().toISOString(),
  };
  mkdirSync(dirname(fixturePath), { recursive: true });
  writeFileSync(fixturePath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  return fixturePath;
}

function resolveReplayFixtureForSuite(suite: string): ReplayFixture | null {
  const replayPath = resolveReplayFixturePath();
  if (!replayPath) {
    return null;
  }
  const fixture = loadReplayFixture(replayPath);
  return fixture.suite === suite ? fixture : null;
}

export function resolveReplaySelector(): ReplaySelector {
  const replayPath = resolveReplayFixturePath();
  if (!replayPath) {
    return {
      enabled: false,
      suite: null,
      kind: null,
      filePath: null,
    };
  }
  const fixture = loadReplayFixture(replayPath);
  return {
    enabled: true,
    suite: fixture.suite,
    kind: fixture.kind ?? null,
    filePath: replayPath,
  };
}

export function shouldRunFuzzSuite(suite: string, kind?: string): boolean {
  const selector = resolveReplaySelector();
  if (!selector.enabled) {
    return true;
  }
  if (selector.suite !== suite) {
    return false;
  }
  if (kind && selector.kind && selector.kind !== kind) {
    return false;
  }
  return true;
}

function resolveParameters<Ts extends unknown[]>(
  suite: string,
  kind: FuzzSuiteKind,
  overrides?: FuzzParameters<Ts>,
): FuzzParameters<Ts> {
  const budget = resolveBudget(kind);
  const replayFixture = resolveReplayFixtureForSuite(suite);
  const seedOverride = parseInteger(process.env["BILIG_FUZZ_SEED"]);
  const runsOverride = parsePositiveInteger(process.env["BILIG_FUZZ_NUM_RUNS"]);
  const maxMsOverride = parsePositiveInteger(process.env["BILIG_FUZZ_MAX_MS"]);

  const resolved: FuzzParameters<Ts> = {
    numRuns: replayFixture?.numRuns ?? budget.numRuns,
    interruptAfterTimeLimit: budget.maxMs,
    ...overrides,
  };
  const seed = seedOverride ?? replayFixture?.seed;
  if (seed !== undefined) {
    resolved.seed = seed;
  }
  if (runsOverride !== undefined) {
    resolved.numRuns = runsOverride;
  }
  if (maxMsOverride !== undefined) {
    resolved.interruptAfterTimeLimit = maxMsOverride;
  }

  if (replayFixture?.path) {
    resolved.path = replayFixture.path;
    resolved.endOnFailure = true;
  }

  return resolved;
}

export function captureCounterexample<Ts extends unknown[]>(
  options: CaptureCounterexampleOptions<Ts>,
): string {
  const artifactDir = resolve(process.cwd(), "artifacts/fuzz", options.suite);
  mkdirSync(artifactDir, { recursive: true });
  const fingerprint = createHash("sha256")
    .update(
      JSON.stringify({
        suite: options.suite,
        kind: options.kind,
        seed: options.details.seed,
        path: options.details.counterexamplePath ?? null,
        counterexample: serializeArtifactValue(options.details.counterexample ?? null),
      }),
    )
    .digest("hex")
    .slice(0, 12);
  const artifactPath = resolve(artifactDir, `${fingerprint}.json`);
  const payload = {
    suite: options.suite,
    kind: options.kind,
    profile: resolveFuzzProfile(),
    seed: options.details.seed,
    path: options.details.counterexamplePath ?? null,
    numRuns: options.details.numRuns,
    numSkips: options.details.numSkips,
    numShrinks: options.details.numShrinks,
    counterexample: serializeArtifactValue(options.details.counterexample ?? null),
    failures: serializeArtifactValue(options.details.failures ?? null),
    error: serializeArtifactValue(options.details.errorInstance ?? null),
    reproductionCommand: `pnpm test:fuzz:replay -- ${artifactPath}`,
    createdAt: new Date().toISOString(),
  };
  mkdirSync(dirname(artifactPath), { recursive: true });
  writeFileSync(artifactPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  return artifactPath;
}

async function runChecked<Ts extends unknown[]>(
  suite: string,
  kind: FuzzSuiteKind,
  property: IProperty<Ts> | IAsyncProperty<Ts>,
  overrides?: FuzzParameters<Ts>,
): Promise<boolean> {
  if (!shouldRunFuzzSuite(suite, kind)) {
    return false;
  }
  const parameters = resolveParameters(suite, kind, overrides);
  const startedAt = Date.now();
  const details = await fc.check(property, parameters);
  const elapsedMs = Date.now() - startedAt;

  console.info(
    [
      `[fuzz] suite=${suite}`,
      `kind=${kind}`,
      `profile=${resolveFuzzProfile()}`,
      `seed=${details.seed}`,
      `runs=${details.numRuns}`,
      `shrinks=${details.numShrinks}`,
      `failed=${details.failed}`,
      `elapsedMs=${elapsedMs}`,
    ].join(" "),
  );

  if (details.failed) {
    let artifactPath: string | null = null;
    if (process.env["BILIG_FUZZ_CAPTURE"] === "1") {
      artifactPath = captureCounterexample({ suite, kind, details });
      console.info(`[fuzz] captured=${artifactPath}`);
    }
    if (details.errorInstance instanceof Error) {
      throw details.errorInstance;
    }
    throw new Error(
      `Fuzz suite ${suite} failed with seed=${details.seed} path=${details.counterexamplePath ?? "n/a"}${artifactPath ? ` artifact=${artifactPath}` : ""}`,
    );
  }

  return true;
}

export async function runProperty<T>(options: PropertySuiteOptions<T>): Promise<boolean> {
  const property = fc.asyncProperty(options.arbitrary, async (value) => {
    await options.predicate(value);
  });
  return runChecked(options.suite, options.kind ?? "property", property, options.parameters);
}

export async function runModelProperty<Model extends object, Real>(
  options: ModelSuiteOptions<Model, Real>,
): Promise<boolean> {
  const property = fc.asyncProperty(options.commands, async (commands) => {
    const model = options.createModel();
    const real = await options.createReal();
    try {
      await fc.asyncModelRun(() => ({ model, real }), commands);
    } finally {
      await options.teardown?.(real);
    }
  });
  return runChecked(options.suite, "model", property, options.parameters);
}

export async function runScheduledProperty<T>(options: ScheduledSuiteOptions<T>): Promise<boolean> {
  const property = fc.asyncProperty(fc.scheduler(), options.arbitrary, async (scheduler, value) => {
    await options.predicate({ scheduler, value });
    await scheduler.waitAll();
  });
  return runChecked(options.suite, options.kind ?? "scheduled", property, options.parameters);
}
