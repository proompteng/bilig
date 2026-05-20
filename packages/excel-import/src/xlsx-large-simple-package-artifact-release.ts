import type { WorkbookPreservedPackagePartSnapshot } from '@bilig/protocol'
import { materializePreservedPackageParts } from './xlsx-preserved-package-parts.js'

interface PreservedPackageArtifacts {
  readonly parts?: readonly WorkbookPreservedPackagePartSnapshot[]
}

interface OpaquePackageArtifacts {
  readonly parts?: readonly unknown[]
}

export interface LargeSimplePackageArtifactReleasePlan {
  readonly retainZipSource: boolean
}

export function prepareLargeSimplePackageArtifactsForZipRelease(input: {
  readonly maxMaterializedBytes?: number
  readonly preservedArtifacts: readonly (PreservedPackageArtifacts | null | undefined)[]
  readonly opaqueArtifacts?: readonly (OpaquePackageArtifacts | null | undefined)[]
}): LargeSimplePackageArtifactReleasePlan {
  const preservedParts = input.preservedArtifacts.flatMap((artifacts) => [...(artifacts?.parts ?? [])])
  const opaquePartCount = input.opaqueArtifacts?.reduce((sum, artifacts) => sum + (artifacts?.parts?.length ?? 0), 0) ?? 0
  if (opaquePartCount > 0) {
    return { retainZipSource: true }
  }
  if (preservedParts.length === 0) {
    return { retainZipSource: false }
  }
  const preservedBytes = preservedParts.reduce((sum, part) => sum + part.byteLength, 0)
  if (input.maxMaterializedBytes === undefined || preservedBytes > input.maxMaterializedBytes) {
    return { retainZipSource: true }
  }
  materializePreservedPackageParts(preservedParts)
  return { retainZipSource: false }
}
