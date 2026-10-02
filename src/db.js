// Lớp truy cập IndexedDB — lưu dữ liệu ngay trong trình duyệt, KHÔNG cần backend.
// Hạn mức lớn (hàng trăm MB tới vài GB tùy máy), phù hợp giữ nhiều tháng dữ liệu.

const DB_NAME = 'booking-db'
const DB_VERSION = 1
const STORE = 'bookings'

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

export async function getAllRows() {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly')
    const req = tx.objectStore(STORE).getAll()
    req.onsuccess = () => resolve(req.result || [])
    req.onerror = () => reject(req.error)
  })
}

// Ghi (thêm/cập nhật) 1 record.
export async function putRow(row) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    tx.objectStore(STORE).put(row)
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

// Ghi hàng loạt (dùng cho lưu thủ công toàn bộ & migrate).
export async function putRows(rows) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    const store = tx.objectStore(STORE)
    rows.forEach(r => store.put(r))
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

export async function deleteRow(id) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    tx.objectStore(STORE).delete(id)
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

export async function clearRows() {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    tx.objectStore(STORE).clear()
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

// Xóa tất cả record có createdAt < isoDate. Trả về số bản ghi đã xóa.
export async function deleteRowsBefore(isoDate) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    const index = tx.objectStore(STORE).index('createdAt')
    const range = IDBKeyRange.upperBound(isoDate, true) // < isoDate (loại trừ)
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

// Ước tính dung lượng đã dùng / tổng hạn mức của origin (bytes).
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
