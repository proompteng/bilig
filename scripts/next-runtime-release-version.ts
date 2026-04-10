#!/usr/bin/env bun

import {
  assertAlignedVersions,
  determineRuntimeReleaseVersion,
  loadRuntimePackages,
} from "./runtime-package-set.ts";

const rootDir = new URL("..", import.meta.url).pathname;
const runtimePackages = loadRuntimePackages(rootDir);
const manifestVersion = assertAlignedVersions(runtimePackages);
const publishedVersions = runtimePackages.map((runtimePackage) => ({
  name: runtimePackage.name,
  version: getPublishedVersion(runtimePackage.name),
}));

const publishedVersionSet = new Set(
  publishedVersions
    .map((entry) => entry.version)
    .filter((version): version is string => typeof version === "string"),
);

if (publishedVersionSet.size > 1) {
  throw new Error(
    `Published runtime package versions are not aligned (${publishedVersions
      .map((entry) => `${entry.name}@${entry.version ?? "unpublished"}`)
      .join(", ")})`,
  );
}

const [publishedVersion] = publishedVersionSet;
const targetVersion = determineRuntimeReleaseVersion({
  autoIncrement: true,
  manifestVersion,
  publishedVersion: publishedVersion ?? null,
});

console.log(
  JSON.stringify(
    {
      manifestVersion,
      publishedVersion: publishedVersion ?? null,
      targetVersion,
    },
    null,
    2,
  ),
);

function getPublishedVersion(packageName: string): string | null {
  const result = Bun.spawnSync(["npm", "view", packageName, "dist-tags.latest", "--json"], {
    cwd: rootDir,
    env: process.env,
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
  });
  if (result.exitCode !== 0) {
    return null;
  }
  const output = new TextDecoder().decode(result.stdout).trim();
  if (output.length === 0 || output === "null") {
    return null;
  }
  return JSON.parse(output);
}
