// Lớp truy cập dữ liệu — lưu SONG SONG vào IndexedDB + localStorage.
// Khi đọc: merge từ cả 2 nguồn để không bao giờ mất record.
// Khi ghi: ghi đồng thời vào cả 2; nếu 1 bên lỗi vẫn ghi bên kia.

const DB_NAME = 'booking-db'
const DB_VERSION = 1
const STORE = 'bookings'
const LS_KEY = 'booking-app-rows' // localStorage backup key

// ───────────────────────── IndexedDB helpers ─────────────────────────

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION)
    req.onupgradeneeded = () => {
      const db = req.result
      if (!db.objectStoreNames.contains(STORE)) {
        const store = db.createObjectStore(STORE, { keyPath: 'id' })
        store.createIndex('createdAt', 'createdAt', { unique: false })
      }
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error)
  })
}

async function idbGetAll() {
  try {
    const db = await openDB()
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly')
      const req = tx.objectStore(STORE).getAll()
      req.onsuccess = () => resolve(req.result || [])
      req.onerror = () => reject(req.error)
    })
  } catch (e) {
    console.warn('[db] IndexedDB read failed:', e)
    return null // null = lỗi, phân biệt với [] = trống
  }
}

async function idbPutRow(row) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    tx.objectStore(STORE).put(row)
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

async function idbPutRows(rows) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    const store = tx.objectStore(STORE)
    rows.forEach(r => store.put(r))
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

async function idbDeleteRow(id) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    tx.objectStore(STORE).delete(id)
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

async function idbClear() {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    tx.objectStore(STORE).clear()
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

async function idbDeleteBefore(isoDate) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    const index = tx.objectStore(STORE).index('createdAt')
    const range = IDBKeyRange.upperBound(isoDate, true)
    const req = index.openCursor(range)
    let count = 0
    req.onsuccess = () => {
      const cursor = req.result
      if (cursor) {
        cursor.delete()
        count++
        cursor.continue()
      }
    }
    tx.oncomplete = () => resolve(count)
    tx.onerror = () => reject(tx.error)
  })
}

// ───────────────────────── localStorage helpers ─────────────────────────

function lsGetAll() {
  try {
    const raw = localStorage.getItem(LS_KEY)
    if (!raw) return []
    const parsed = JSON.parse(raw)
    return Array.isArray(parsed) ? parsed : []
  } catch (e) {
    console.warn('[db] localStorage read failed:', e)
    return []
  }
}

function lsSaveAll(rows) {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify(rows))
    return true
  } catch (e) {
    console.warn('[db] localStorage write failed:', e)
    return false
  }
}

function lsClear() {
  try {
    localStorage.removeItem(LS_KEY)
  } catch (e) {
    console.warn('[db] localStorage clear failed:', e)
  }
}

// ───────────────────────── Merge logic ─────────────────────────
// Merge 2 mảng dữ liệu theo id, giữ tất cả record tồn tại ở ít nhất 1 bên.

function mergeRows(a, b) {
  const map = new Map()
  for (const r of a) {
    if (r && r.id) map.set(r.id, r)
  }
  for (const r of b) {
    if (r && r.id && !map.has(r.id)) {
      map.set(r.id, r)
    }
  }
  return [...map.values()]
}

// ───────────────────────── Public API ─────────────────────────

/**
 * Đọc tất cả rows — merge từ IndexedDB + localStorage.
 * Nếu phát hiện 1 bên thiếu dữ liệu, tự động ghi bổ sung để đồng bộ.
 */
export async function getAllRows() {
  const idbRows = await idbGetAll()   // null nếu lỗi, [] nếu trống
  const lsRows = lsGetAll()           // luôn trả []

  // Cả 2 đều đọc được
  if (idbRows !== null) {
    const merged = mergeRows(idbRows, lsRows)

    // Tự động sync lại nếu phát hiện lệch
    if (merged.length > idbRows.length) {
      // localStorage có record mà IndexedDB thiếu → ghi bổ sung vào IDB
      console.info(`[db] Khôi phục ${merged.length - idbRows.length} record từ localStorage vào IndexedDB`)
      try { await idbPutRows(merged) } catch { /* best effort */ }
    }
    if (merged.length > lsRows.length) {
      // IndexedDB có record mà localStorage thiếu → cập nhật LS
      console.info(`[db] Đồng bộ ${merged.length - lsRows.length} record từ IndexedDB vào localStorage`)
      lsSaveAll(merged)
    }

    return merged
  }

  // IndexedDB lỗi → dùng localStorage
  if (lsRows.length > 0) {
    console.warn('[db] IndexedDB không khả dụng, phục hồi từ localStorage:', lsRows.length, 'records')
    // Thử ghi lại vào IDB khi nó khả dụng lần sau
    try { await idbPutRows(lsRows) } catch { /* sẽ thử lại lần sau */ }
  }
  return lsRows
}

/** Ghi 1 record vào cả IndexedDB + localStorage. */
export async function putRow(row) {
  const errors = []

  // Ghi IDB
  try {
    await idbPutRow(row)
  } catch (e) {
    errors.push({ target: 'IndexedDB', error: e })
  }

  // Ghi LS — đọc ra, cập nhật, ghi lại
  try {
    const lsRows = lsGetAll()
    const idx = lsRows.findIndex(r => r.id === row.id)
    if (idx >= 0) {
      lsRows[idx] = row
    } else {
      lsRows.push(row)
    }
    if (!lsSaveAll(lsRows)) {
      errors.push({ target: 'localStorage', error: new Error('localStorage write returned false') })
    }
  } catch (e) {
    errors.push({ target: 'localStorage', error: e })
  }

  // Cả 2 đều lỗi → throw
  if (errors.length === 2) {
    throw errors[0].error
  }
}

/** Ghi toàn bộ rows vào cả IndexedDB + localStorage. */
export async function putRows(rows) {
  const errors = []

  try {
    await idbPutRows(rows)
  } catch (e) {
    errors.push({ target: 'IndexedDB', error: e })
  }

  if (!lsSaveAll(rows)) {
    errors.push({ target: 'localStorage', error: new Error('localStorage save failed') })
  }

  if (errors.length === 2) {
    throw errors[0].error
  }
}

/** Xóa 1 record theo id ở cả 2 nơi. */
export async function deleteRow(id) {
  try { await idbDeleteRow(id) } catch (e) { console.warn('[db] IDB deleteRow error:', e) }

  try {
    const lsRows = lsGetAll()
    lsSaveAll(lsRows.filter(r => r.id !== id))
  } catch (e) { console.warn('[db] LS deleteRow error:', e) }
}

/** Xóa toàn bộ dữ liệu ở cả 2 nơi. */
export async function clearRows() {
  try { await idbClear() } catch (e) { console.warn('[db] IDB clear error:', e) }
  lsClear()
}

/** Xóa record trước ngày iso ở cả 2 nơi. Trả về số record đã xóa. */
export async function deleteRowsBefore(isoDate) {
  let count = 0
  try { count = await idbDeleteBefore(isoDate) } catch (e) { console.warn('[db] IDB deleteBefore error:', e) }

  try {
    const lsRows = lsGetAll()
    const remaining = lsRows.filter(r => (r.createdAt || '') >= isoDate)
    const lsRemoved = lsRows.length - remaining.length
    lsSaveAll(remaining)
    count = Math.max(count, lsRemoved)
  } catch (e) { console.warn('[db] LS deleteBefore error:', e) }

  return count
}

/** Ước tính dung lượng storage. */
export async function estimateStorage() {
  if (navigator.storage && navigator.storage.estimate) {
    try {
      const { usage = 0, quota = 0 } = await navigator.storage.estimate()
      return { usage, quota }
    } catch {
      return { usage: 0, quota: 0 }
    }
  }
  return { usage: 0, quota: 0 }
}

/**
 * Yêu cầu trình duyệt giữ dữ liệu vĩnh viễn (persistent storage).
 * Trả về true nếu được grant, false nếu bị từ chối hoặc không hỗ trợ.
 */
export async function requestPersist() {
  if (navigator.storage && navigator.storage.persist) {
    try {
      return await navigator.storage.persist()
    } catch {
      return false
    }
  }
  return false
}

/**
 * Kiểm tra trạng thái persist hiện tại.
 * Trả về true/false hoặc null nếu không hỗ trợ.
 */
export async function checkPersisted() {
  if (navigator.storage && navigator.storage.persisted) {
    try {
      return await navigator.storage.persisted()
    } catch {
      return null
    }
  }
  return null
}
