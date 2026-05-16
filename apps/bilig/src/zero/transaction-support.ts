import type { Queryable } from './store.js'

interface TransactionClient extends Queryable {
  release(): void
}

interface TransactionalQueryable extends Queryable {
  connect(): Promise<TransactionClient>
}

function isTransactionalQueryable(db: Queryable): db is TransactionalQueryable {
  return 'connect' in db && typeof db.connect === 'function'
}

export async function runQueryableTransaction(db: Queryable, task: (transactionDb: Queryable) => Promise<void>): Promise<void> {
  if (!isTransactionalQueryable(db)) {
    await task(db)
    return
  }
  const client = await db.connect()
  try {
    await client.query('BEGIN')
    await task(client)
    await client.query('COMMIT')
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

export async function runSequentially<T>(items: readonly T[], task: (item: T, index: number) => Promise<void>): Promise<void> {
  await items.reduce<Promise<void>>(async (previous, item, index) => {
    await previous
    await task(item, index)
  }, Promise.resolve())
}
