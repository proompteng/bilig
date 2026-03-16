const DB_NAME = "bilig-playground";
const STORE_NAME = "state";

function getLocalStorage(): Storage | null {
  try {
    return window.localStorage;
  } catch {
    return null;
  }
}

async function openDatabase(): Promise<IDBDatabase | null> {
  if (typeof window === "undefined" || typeof window.indexedDB === "undefined") {
    return null;
  }

  return await new Promise<IDBDatabase | null>((resolve) => {
    try {
      const request = window.indexedDB.open(DB_NAME, 1);

      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(STORE_NAME)) {
          database.createObjectStore(STORE_NAME);
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function readFromStore(key: string): Promise<string | null> {
  const database = await openDatabase();
  if (!database) {
    return null;
  }

  return await new Promise<string | null>((resolve) => {
    const transaction = database.transaction(STORE_NAME, "readonly");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.get(key);

    request.onsuccess = () => {
      const result = request.result;
      resolve(typeof result === "string" ? result : null);
    };
    request.onerror = () => resolve(null);
    transaction.oncomplete = () => database.close();
    transaction.onerror = () => database.close();
    transaction.onabort = () => database.close();
  });
}

async function writeToStore(key: string, value: string): Promise<boolean> {
  const database = await openDatabase();
  if (!database) {
    return false;
  }

  return await new Promise<boolean>((resolve) => {
    let settled = false;
    const finish = (result: boolean) => {
      if (settled) {
        return;
      }
      settled = true;
      resolve(result);
    };

    const transaction = database.transaction(STORE_NAME, "readwrite");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.put(value, key);

    request.onsuccess = () => finish(true);
    request.onerror = () => finish(false);
    transaction.oncomplete = () => {
      database.close();
      finish(true);
    };
    transaction.onerror = () => {
      database.close();
      finish(false);
    };
    transaction.onabort = () => {
      database.close();
      finish(false);
    };
  });
}

async function removeFromStore(key: string): Promise<void> {
  const database = await openDatabase();
  if (!database) {
    return;
  }

  await new Promise<void>((resolve) => {
    const transaction = database.transaction(STORE_NAME, "readwrite");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.delete(key);

    request.onsuccess = () => resolve();
    request.onerror = () => resolve();
    transaction.oncomplete = () => {
      database.close();
      resolve();
    };
    transaction.onerror = () => {
      database.close();
      resolve();
    };
    transaction.onabort = () => {
      database.close();
      resolve();
    };
  });
}

export async function loadPersistedJson<T>(key: string): Promise<T | null> {
  const persisted = await readFromStore(key);
  if (persisted) {
    try {
      return JSON.parse(persisted) as T;
    } catch {
      return null;
    }
  }

  const storage = getLocalStorage();
  const legacyValue = storage?.getItem(key) ?? null;
  if (!legacyValue) {
    return null;
  }

  try {
    const parsed = JSON.parse(legacyValue) as T;
    await writeToStore(key, legacyValue);
    storage?.removeItem(key);
    return parsed;
  } catch {
    return null;
  }
}

export async function savePersistedJson(key: string, value: unknown): Promise<void> {
  const serialized = JSON.stringify(value);
  const persistedToStore = await writeToStore(key, serialized);
  if (persistedToStore) {
    const storage = getLocalStorage();
    storage?.removeItem(key);
    return;
  }

  const storage = getLocalStorage();
  if (!storage) {
    return;
  }

  try {
    storage.setItem(key, serialized);
  } catch (error) {
    console.warn(`Unable to persist ${key}`, error);
  }
}

export async function removePersistedJson(key: string): Promise<void> {
  await removeFromStore(key);
  const storage = getLocalStorage();
  storage?.removeItem(key);
}
