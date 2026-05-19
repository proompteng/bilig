import type { LargeSimpleXlsxImportStats } from '../packages/excel-import/src/xlsx-large-simple-import.js'
import { formatByteSize } from './public-workbook-corpus-process.ts'

export function largeSimpleImportPhaseTelemetryEvidence(stats: LargeSimpleXlsxImportStats): string[] {
  return stats.phaseTelemetry.map((entry) =>
    [
      `large-simple-import-phase=${entry.phase}`,
      `elapsed-ms=${String(entry.elapsedMs)}`,
      ...(entry.rssBytes !== undefined ? [`rss=${formatByteSize(entry.rssBytes)}`] : []),
      ...(entry.heapUsedBytes !== undefined ? [`heap-used=${formatByteSize(entry.heapUsedBytes)}`] : []),
      ...(entry.zipSourceBytesBeforeRelease !== undefined
        ? [`zip-source-before-release=${formatByteSize(entry.zipSourceBytesBeforeRelease)}`]
        : []),
      ...(entry.zipSourceBytesAfterRelease !== undefined
        ? [`zip-source-after-release=${formatByteSize(entry.zipSourceBytesAfterRelease)}`]
        : []),
      ...(entry.ownedSourceBytesBeforeRelease !== undefined
        ? [`owned-source-before-release=${formatByteSize(entry.ownedSourceBytesBeforeRelease)}`]
        : []),
      ...(entry.ownedSourceBytesAfterRelease !== undefined
        ? [`owned-source-after-release=${formatByteSize(entry.ownedSourceBytesAfterRelease)}`]
        : []),
    ].join(','),
  )
}
