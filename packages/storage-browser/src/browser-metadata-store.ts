const DEFAULT_DB_NAME = 'bilig-browser-state'
const DEFAULT_STORE_NAME = 'state'
const WRITE_THROUGH_LOCALSTORAGE_LIMIT_BYTES = 128 * 1024

function addOnceEventListener(target: EventTarget, type: string, listener: () => void): void {
  target.addEventListener(
    type,
    () => {
      listener()
    },
    { once: true },
  )
}

function getLocalStorage(): Storage | null {
  try {
    const scope = globalThis as typeof globalThis & { localStorage?: Storage }
    return scope.localStorage ?? null
  } catch {
    return null
  }
}

async function openDatabase(databaseName: string, storeName: string): Promise<IDBDatabase | null> {
  const scope = globalThis as typeof globalThis & { indexedDB?: IDBFactory }
  if (typeof scope.indexedDB === 'undefined') {
    return null
  }

  return await new Promise<IDBDatabase | null>((resolve) => {
    try {
      const request = scope.indexedDB.open(databaseName, 1)
      request.onupgradeneeded = () => {
        const database = request.result
        if (!database.objectStoreNames.contains(storeName)) {
          database.createObjectStore(storeName)
        }
      }
      addOnceEventListener(request, 'success', () => resolve(request.result))
      addOnceEventListener(request, 'error', () => resolve(null))
      addOnceEventListener(request, 'blocked', () => resolve(null))
    } catch {
      resolve(null)
    }
  })
}

async function readFromStore(databaseName: string, storeName: string, key: string): Promise<string | null> {
  const database = await openDatabase(databaseName, storeName)
  if (!database) {
    return null
  }

  return await new Promise<string | null>((resolve) => {
    const transaction = database.transaction(storeName, 'readonly')
    const store = transaction.objectStore(storeName)
    const request = store.get(key)

    addOnceEventListener(request, 'success', () => resolve(typeof request.result === 'string' ? request.result : null))
    addOnceEventListener(request, 'error', () => resolve(null))
    addOnceEventListener(transaction, 'complete', () => database.close())
    addOnceEventListener(transaction, 'error', () => database.close())
    addOnceEventListener(transaction, 'abort', () => database.close())
  })
}

async function writeToStore(databaseName: string, storeName: string, key: string, value: string): Promise<boolean> {
  const database = await openDatabase(databaseName, storeName)
  if (!database) {
    return false
  }

  return await new Promise<boolean>((resolve) => {
    let settled = false
    const finish = (result: boolean) => {
      if (settled) {
        return
      }
      settled = true
      resolve(result)
    }

    const transaction = database.transaction(storeName, 'readwrite')
    const request = transaction.objectStore(storeName).put(value, key)
    addOnceEventListener(request, 'success', () => finish(true))
    addOnceEventListener(request, 'error', () => finish(false))
    addOnceEventListener(transaction, 'complete', () => {
      database.close()
      finish(true)
    })
    addOnceEventListener(transaction, 'error', () => {
      database.close()
      finish(false)
    })
    addOnceEventListener(transaction, 'abort', () => {
      database.close()
      finish(false)
    })
  })
}

async function removeFromStore(databaseName: string, storeName: string, key: string): Promise<void> {
  const database = await openDatabase(databaseName, storeName)
  if (!database) {
    return
  }

  await new Promise<void>((resolve) => {
    const transaction = database.transaction(storeName, 'readwrite')
    const request = transaction.objectStore(storeName).delete(key)
    addOnceEventListener(request, 'success', () => resolve())
    addOnceEventListener(request, 'error', () => resolve())
    addOnceEventListener(transaction, 'complete', () => {
      database.close()
      resolve()
    })
    addOnceEventListener(transaction, 'error', () => {
      database.close()
      resolve()
    })
    addOnceEventListener(transaction, 'abort', () => {
      database.close()
      resolve()
    })
  })
}

export interface BrowserMetadataStoreOptions {
  databaseName?: string
  storeName?: string
}

export interface BrowserMetadataStore {
  loadJson<T>(key: string, parser: (value: unknown) => T | null): Promise<T | null>
  saveJson(key: string, value: unknown): Promise<void>
  remove(key: string): Promise<void>
}

export function createBrowserMetadataStore(options: BrowserMetadataStoreOptions = {}): BrowserMetadataStore {
  const databaseName = options.databaseName ?? DEFAULT_DB_NAME
  const storeName = options.storeName ?? DEFAULT_STORE_NAME

  return {
    async loadJson<T>(key: string, parser: (value: unknown) => T | null): Promise<T | null> {
      const storage = getLocalStorage()
      const cachedValue = storage?.getItem(key) ?? null
      if (cachedValue) {
        try {
          const parsed = parser(JSON.parse(cachedValue) as unknown)
          if (parsed !== null) {
            void writeToStore(databaseName, storeName, key, cachedValue)
            return parsed
          }
          storage?.removeItem(key)
        } catch {
          storage?.removeItem(key)
        }
      }

      const persisted = await readFromStore(databaseName, storeName, key)
      if (persisted) {
        try {
          return parser(JSON.parse(persisted) as unknown)
        } catch {
          return null
        }
      }

      return null
    },
    async saveJson(key: string, value: unknown): Promise<void> {
      const serialized = JSON.stringify(value)
      const storage = getLocalStorage()
      if (storage) {
        try {
          if (serialized.length <= WRITE_THROUGH_LOCALSTORAGE_LIMIT_BYTES) {
            storage.setItem(key, serialized)
          } else {
            storage.removeItem(key)
          }
        } catch {
          storage.removeItem(key)
        }
      }

      const persisted = await writeToStore(databaseName, storeName, key, serialized)
      if (persisted) {
        return
      }

      try {
        storage?.setItem(key, serialized)
      } catch (error) {
        console.warn(`Unable to persist ${key}`, error)
      }
    },
    async remove(key: string): Promise<void> {
      await removeFromStore(databaseName, storeName, key)
      getLocalStorage()?.removeItem(key)
    },
  }
}

const defaultMetadataStore = createBrowserMetadataStore()

export async function loadPersistedJson<T>(key: string, parser: (value: unknown) => T | null): Promise<T | null> {
  return defaultMetadataStore.loadJson<T>(key, parser)
}

export async function savePersistedJson(key: string, value: unknown): Promise<void> {
  return defaultMetadataStore.saveJson(key, value)
}

export async function removePersistedJson(key: string): Promise<void> {
  return defaultMetadataStore.remove(key)
}
