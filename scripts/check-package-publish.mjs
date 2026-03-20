import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { isAbsolute, join, resolve } from "node:path";

const rootDir = resolve(new URL("..", import.meta.url).pathname);
const packagesDir = join(rootDir, "packages");
const packDir = join(rootDir, "build", "npm-packages");

const packageDirs = readdirSync(packagesDir)
  .map((name) => join(packagesDir, name))
  .filter((dir) => existsSync(join(dir, "package.json")));

rmSync(packDir, { recursive: true, force: true });
mkdirSync(packDir, { recursive: true });

const failures = [];

for (const packageDir of packageDirs) {
  const packageJsonPath = join(packageDir, "package.json");
  const manifest = JSON.parse(readFileSync(packageJsonPath, "utf8"));
  const packageLabel = manifest.name ?? packageDir;

  validateManifestShape(packageLabel, manifest, failures);

  const tarballName = execFileSync("pnpm", ["pack", "--pack-destination", packDir], {
    cwd: packageDir,
    encoding: "utf8"
  })
    .trim()
    .split("\n")
    .pop();

  if (!tarballName) {
    failures.push(`${packageLabel}: pnpm pack did not return a tarball name`);
    continue;
  }

  const tarballPath = isAbsolute(tarballName) ? tarballName : join(packDir, tarballName);
  const tarEntries = execFileSync("tar", ["-tf", tarballPath], { encoding: "utf8" })
    .split("\n")
    .filter(Boolean);

  validateTarballContents(packageLabel, manifest, tarEntries, failures);

  const packedManifest = JSON.parse(execFileSync("tar", ["-xOf", tarballPath, "package/package.json"], { encoding: "utf8" }));
  validatePackedManifest(packageLabel, packedManifest, failures);
}

if (failures.length > 0) {
  throw new Error(`npm publish readiness check failed:\n- ${failures.join("\n- ")}`);
}

console.log(
  JSON.stringify(
    {
      packages: packageDirs.length,
      packDir
    },
    null,
    2
  )
);

function validateManifestShape(packageLabel, manifest, failureMessages) {
  const requiredFields = ["name", "version", "description", "license", "repository", "homepage", "bugs", "main", "types", "exports", "files"];
  for (const field of requiredFields) {
    if (!(field in manifest)) {
      failureMessages.push(`${packageLabel}: missing required manifest field "${field}"`);
    }
  }

  if (manifest.publishConfig?.access !== "public") {
    failureMessages.push(`${packageLabel}: publishConfig.access must be "public"`);
  }

  if (!Array.isArray(manifest.files) || manifest.files.length === 0) {
    failureMessages.push(`${packageLabel}: files list must be present and non-empty`);
  }

  if ((packageLabel === "@bilig/grid" || packageLabel === "@bilig/renderer") && !manifest.peerDependencies?.react) {
    failureMessages.push(`${packageLabel}: react must be declared as a peer dependency`);
  }
}

function validateTarballContents(packageLabel, manifest, tarEntries, failureMessages) {
  const requiredEntries = new Set(["package/package.json", "package/README.md", "package/LICENSE"]);

  if (typeof manifest.main === "string") {
    requiredEntries.add(`package/${stripDotSlash(manifest.main)}`);
  }
  if (typeof manifest.types === "string") {
    requiredEntries.add(`package/${stripDotSlash(manifest.types)}`);
  }

  collectExportTargets(manifest.exports).forEach((target) => requiredEntries.add(`package/${stripDotSlash(target)}`));
  if (packageLabel === "@bilig/wasm-kernel") {
    requiredEntries.add("package/build/release.wasm");
  }

  for (const requiredEntry of requiredEntries) {
    if (!tarEntries.includes(requiredEntry)) {
      failureMessages.push(`${packageLabel}: tarball is missing ${requiredEntry}`);
    }
  }

  for (const entry of tarEntries) {
    if (entry.includes("__tests__")) {
      failureMessages.push(`${packageLabel}: tarball must not contain test artifacts (${entry})`);
    }
    if (entry.endsWith(".tsbuildinfo")) {
      failureMessages.push(`${packageLabel}: tarball must not contain tsbuildinfo (${entry})`);
    }
    if (entry.startsWith("package/src/")) {
      failureMessages.push(`${packageLabel}: tarball must not contain source files (${entry})`);
    }
  }
}

function validatePackedManifest(packageLabel, packedManifest, failureMessages) {
  const serialized = JSON.stringify(packedManifest);
  if (serialized.includes("workspace:*")) {
    failureMessages.push(`${packageLabel}: packed manifest still contains workspace:* dependency ranges`);
  }
}

function collectExportTargets(exportsField) {
  const targets = new Set();
  visitExports(exportsField, targets);
  return [...targets];
}

function visitExports(node, targets) {
  if (typeof node === "string") {
    targets.add(node);
    return;
  }
  if (!node || typeof node !== "object") {
    return;
  }
  for (const value of Object.values(node)) {
    visitExports(value, targets);
  }
}

function stripDotSlash(value) {
  return value.startsWith("./") ? value.slice(2) : value;
}
