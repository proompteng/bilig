#!/usr/bin/env bun

import path from 'node:path'
import { runProtocolGenerator } from './protocol-generator.js'
import { builtinManifest, enumManifest } from './protocol-manifest.js'

const repoRoot = path.resolve(import.meta.dir, '..')

await runProtocolGenerator({ repoRoot, enumManifest, builtinManifest })
