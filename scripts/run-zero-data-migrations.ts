import { createZeroPool, resolveZeroDatabaseUrl } from '../apps/bilig/src/zero/db.js'
import { runPendingZeroDataMigrations } from '../apps/bilig/src/zero/data-migration-runner.js'

const connectionString = resolveZeroDatabaseUrl()

if (!connectionString) {
  throw new Error('Zero database URL is not configured. Set ZERO_UPSTREAM_DB, DATABASE_URL, or BILIG_DATABASE_URL.')
}

const pool = createZeroPool(connectionString)

try {
  const result = await runPendingZeroDataMigrations(pool)
  if (result.appliedThisRun.length === 0) {
    console.info('Zero data migrations are already up to date.')
  } else {
    console.info(`Applied Zero data migrations: ${result.appliedThisRun.join(', ')}`)
  }
} finally {
  await pool.end()
}
