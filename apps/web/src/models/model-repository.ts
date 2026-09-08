import { parseModelDocument, type ModelDocument } from './model-document.js'

export class ModelConflictError extends Error {
  constructor() {
    super('This model changed in another tab. Export your edits or reload the saved version.')
    this.name = 'ModelConflictError'
  }
}

let database: Promise<IDBDatabase> | null = null

async function openDatabase(): Promise<IDBDatabase> {
  if (database) return database
  database = new Promise((resolve, reject) => {
    const request = indexedDB.open('bilig-model-workspace', 1)
    let blocked = false
    request.onupgradeneeded = () => request.result.createObjectStore('models', { keyPath: 'id' })
    request.addEventListener('error', () => {
      database = null
      reject(request.error ?? new Error('Browser storage is unavailable.'))
    })
    request.onblocked = () => {
      blocked = true
      database = null
      reject(new Error('Close other Bilig tabs, then retry opening browser storage.'))
    }
    request.onsuccess = () => {
      if (blocked) {
        request.result.close()
        return
      }
      request.result.onversionchange = () => {
        request.result.close()
        database = null
      }
      resolve(request.result)
    }
  })
  try {
    return await database
  } catch (error) {
    database = null
    throw error
  }
}

export async function listModels(): Promise<ModelDocument[]> {
  const db = await openDatabase()
  return new Promise((resolve, reject) => {
    const tx = db.transaction('models', 'readonly')
    const request = tx.objectStore('models').getAll()
    tx.addEventListener('abort', () => reject(tx.error ?? new Error('Could not read saved models.')))
    tx.oncomplete = () => {
      try {
        const records: unknown[] = request.result
        resolve(records.map(parseModelDocument).toSorted((a, b) => b.updatedAt.localeCompare(a.updatedAt)))
      } catch (error) {
        reject(error)
      }
    }
  })
}

export async function saveModel(model: ModelDocument, expectedRevision: number): Promise<ModelDocument> {
  const checked = parseModelDocument(model)
  const db = await openDatabase()
  return new Promise((resolve, reject) => {
    const saved = { ...checked, revision: expectedRevision + 1, updatedAt: new Date().toISOString() }
    const tx = db.transaction('models', 'readwrite')
    const store = tx.objectStore('models')
    const request = store.get(checked.id)
    let failure: unknown
    request.onsuccess = () => {
      try {
        const previous: unknown = request.result
        const revision = previous === undefined ? 0 : parseModelDocument(previous).revision
        if (revision !== expectedRevision) throw new ModelConflictError()
        store.put(saved)
      } catch (error) {
        failure = error
        tx.abort()
      }
    }
    tx.addEventListener('abort', () => reject(failure ?? tx.error ?? new Error('Could not save. Export a backup before closing this tab.')))
    tx.oncomplete = () => resolve(saved)
  })
}
