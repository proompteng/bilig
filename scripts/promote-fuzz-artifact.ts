#!/usr/bin/env bun

import { promoteCapturedArtifact } from "@bilig/test-fuzz";

const args = process.argv.slice(2).filter((value) => value !== "--");
const [artifactPath, fixturePath] = args;

if (!artifactPath || !fixturePath) {
  console.error("Usage: pnpm test:fuzz:promote -- <artifact-path> <fixture-path>");
  process.exit(1);
}

const promotedPath = promoteCapturedArtifact({
  artifactPath,
  fixturePath,
});

console.info(`Promoted fuzz artifact to ${promotedPath}`);
