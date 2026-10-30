// store.js: everything is kept in this browser's IndexedDB. Nothing is uploaded.

const DB_NAME = 'presentation-helper';
const DB_VERSION = 1;

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('projects')) db.createObjectStore('projects', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('files')) db.createObjectStore('files');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function run(storeName, mode, fn) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const store = tx.objectStore(storeName);
    const req = fn(store);
    tx.oncomplete = () => resolve(req ? req.result : undefined);
    tx.onerror = () => reject(tx.error);
  });
}

export const listProjects = () => run('projects', 'readonly', (s) => s.getAll());
export const getProject = (id) => run('projects', 'readonly', (s) => s.get(id));
export const putProject = (project) => run('projects', 'readwrite', (s) => s.put(project));
export const getFile = (key) => run('files', 'readonly', (s) => s.get(key));
export const putFile = (key, bytes) => run('files', 'readwrite', (s) => s.put(bytes, key));

export async function deleteProject(id) {
  await run('projects', 'readwrite', (s) => s.delete(id));
  await run('files', 'readwrite', (s) => s.delete(id + ':original'));
  await run('files', 'readwrite', (s) => s.delete(id + ':saved'));
}
