import { useState, useRef, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'
import './App.css'
import {
  getAllRows,
  putRow,
  putRows,
  deleteRow as dbDeleteRow,
  clearRows,
  deleteRowsBefore,
  estimateStorage,
} from './db'

const PAGE_SIZE_OPTIONS = [20, 50, 100, 200]
const WARN_PERCENT = 80 // Nhắc xuất Excel khi dung lượng đạt mức này
const LEGACY_KEY = 'booking-app-rows' // Khóa localStorage cũ để migrate

function parseMoney(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val
  if (!val) return 0
  const digits = String(val).replace(/[^\d]/g, '')
  return digits ? parseInt(digits, 10) : 0
}

function parseBookingText(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean)
  const get = (key) => {
    const line = lines.find(l => l.toLowerCase().startsWith(key.toLowerCase()))
    if (!line) return ''
    return line.substring(line.indexOf(':') + 1).trim()
  }

  const checkIn = get('Check in')
  const checkOut = get('Check out')
  const gia = get('Giá')
  const daThanhToan = get('Đã thanh toán')
  const nguoiSale = get('Người sale')
  const tenCCCD = get('Tên CCCD')
  const maPhong = get('Mã Phòng Thuê') || get('Mã phòng thuê')

  let trangThai = 'Chưa thanh toán'
  const giaNum = parseFloat(gia) || 0
  const daTTNum = parseFloat(daThanhToan) || 0

  if (daTTNum > 0 && giaNum > 0 && daTTNum >= giaNum) {
    trangThai = 'Đã TT đủ'
  } else if (daTTNum > 0) {
    trangThai = 'Đã TT 1 phần'
  } else {
    trangThai = 'Chưa thanh toán'
  }

  const formatGia = (val) => {
    const num = parseFloat(val)
    if (isNaN(num)) return val
    return (num * 1000).toLocaleString('vi-VN')
  }

  const ngayGioCheckIn = checkIn

  return {
    tenKhach: tenCCCD,
    giaPhong: formatGia(gia),
    ngayGioCheckIn,
    trangThai,
    phong: maPhong,
    nguoiSale,
    daThanhToan: daThanhToan ? formatGia(daThanhToan) : '',
    note: checkOut ? `Check out: ${checkOut}` : '',
  }
}

function newId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8)
}

function fmtBytes(b) {
  if (!b) return '0 KB'
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(0)} KB`
  if (b < 1024 * 1024 * 1024) return `${(b / 1024 / 1024).toFixed(1)} MB`
  return `${(b / 1024 / 1024 / 1024).toFixed(2)} GB`
}

function fmtDate(iso) {
  if (!iso) return ''
  const d = new Date(iso)
  if (isNaN(d.getTime())) return ''
  const p = n => String(n).padStart(2, '0')
  return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`
}

function App() {
  const [text, setText] = useState('')
  const [rows, setRows] = useState([])
  const [editId, setEditId] = useState(null)
  const [error, setError] = useState('')
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(50)
  const [storage, setStorage] = useState({ usage: 0, quota: 0 })
  const [loading, setLoading] = useState(true)
  const [dirty, setDirty] = useState(false)
  const [lastSaved, setLastSaved] = useState(null)
  const [deleteBeforeDate, setDeleteBeforeDate] = useState('')
  const rowsRef = useRef(rows)

  const setRowsSynced = (newRows) => {
    rowsRef.current = newRows
    setRows(newRows)
  }

  const refreshStorage = useCallback(async () => {
    setStorage(await estimateStorage())
  }, [])

  const onQuota = (e) => {
    if (e && e.name === 'QuotaExceededError') {
      setError('Hết dung lượng lưu trữ trình duyệt. Hãy xuất Excel rồi xóa bớt dữ liệu cũ.')
    }
  }

  // Tải dữ liệu lần đầu: migrate từ localStorage cũ (nếu có) rồi đọc từ IndexedDB.
  useEffect(() => {
    (async () => {
      try {
        let data = await getAllRows()
        if (data.length === 0) {
          const raw = localStorage.getItem(LEGACY_KEY)
          if (raw) {
            try {
              const legacy = JSON.parse(raw)
              if (Array.isArray(legacy) && legacy.length) {
                const now = new Date().toISOString()
                data = legacy.map(r => ({
                  ...r,
                  id: r.id || newId(),
                  createdAt: r.createdAt || now,
                }))
                await putRows(data)
              }
            } catch { /* bỏ qua dữ liệu cũ hỏng */ }
          }
        }
        data.sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''))
        setRowsSynced(data)
      } catch (e) {
        setError('Không mở được cơ sở dữ liệu trình duyệt: ' + (e?.message || e))
      } finally {
        setLoading(false)
        refreshStorage()
      }
    })()
  }, [refreshStorage])

  const validateText = (text) => {
    if (!text.trim()) return 'Vui lòng dán form chốt khách vào ô nhập liệu.'

    const lines = text.split('\n').map(l => l.trim()).filter(Boolean)
    const get = (key) => {
      const line = lines.find(l => l.toLowerCase().startsWith(key.toLowerCase()))
      return line ? line.substring(line.indexOf(':') + 1).trim() : ''
    }

    // Chỉ cần ít nhất 1 trường có dữ liệu là thêm được; các trường thiếu để trống.
    const fields = [
      get('Check in'),
      get('Check out'),
      get('Giá'),
      get('Đã thanh toán'),
      get('Người sale'),
      get('Tên CCCD'),
      get('Mã Phòng Thuê') || get('Mã phòng thuê'),
    ]
    if (!fields.some(v => v)) {
      return 'Cần ít nhất 1 trường hợp lệ (Check in, Check out, Giá, Tên CCCD, Mã Phòng Thuê...).'
    }

    const gia = get('Giá')
    if (gia && isNaN(parseFloat(gia))) return `"Giá" phải là số (VD: 149), nhận được: "${gia}"`

    return ''
  }

  const handleAdd = () => {
    const err = validateText(text)
    if (err) {
      setError(err)
      return
    }
    setError('')
    const parsed = parseBookingText(text)
    const row = { ...parsed, id: newId(), createdAt: new Date().toISOString() }
    const newRows = [...rowsRef.current, row]
    setRowsSynced(newRows)
    setText('')
    setPage(Math.ceil(newRows.length / pageSize))
    putRow(row).then(() => { setDirty(false); setLastSaved(new Date()); refreshStorage() }).catch(onQuota)
  }

  const handleEdit = (id, field, value) => {
    const updated = rowsRef.current.map(r => (r.id === id ? { ...r, [field]: value } : r))
    setRowsSynced(updated)
    setDirty(true)
  }

  const handleRowSaveDone = (row) => {
    setEditId(null)
    putRow(row)
      .then(() => { setDirty(false); setLastSaved(new Date()); refreshStorage() })
      .catch(onQuota)
  }

  const handleDelete = (id) => {
    const row = rowsRef.current.find(r => r.id === id)
    if (!row) return
    setRowsSynced(rowsRef.current.filter(r => r.id !== id))
    if (editId === id) setEditId(null)
    if (row?.id) dbDeleteRow(row.id).then(refreshStorage)
  }

  const handleClearAll = async () => {
    if (!window.confirm('Xóa toàn bộ dữ liệu đã lưu? Hành động này không thể hoàn tác.')) return
    await clearRows()
    setRowsSynced([])
    setEditId(null)
    setPage(1)
    setDirty(false)
    refreshStorage()
  }

  const handleDeleteBefore = async () => {
    if (!deleteBeforeDate) {
      setError('Vui lòng chọn ngày trước khi xóa.')
      return
    }
    const iso = new Date(deleteBeforeDate + 'T00:00:00').toISOString()
    if (!window.confirm(`Xóa tất cả record được tạo TRƯỚC ngày ${deleteBeforeDate}?`)) return
    const removed = await deleteRowsBefore(iso)
    const remaining = rowsRef.current.filter(r => (r.createdAt || '') >= iso)
    setRowsSynced(remaining)
    setPage(1)
    setError('')
    refreshStorage()
    window.alert(`Đã xóa ${removed} record.`)
  }

  const handleManualSave = async () => {
    try {
      await putRows(rowsRef.current)
      setDirty(false)
      setLastSaved(new Date())
      refreshStorage()
    } catch (e) {
      onQuota(e)
    }
  }

  const handleExport = () => {
    setEditId(null)
    const currentRows = rowsRef.current
    if (currentRows.length === 0) return

    const data = currentRows.map((r, i) => ({
      'STT': i + 1,
      'Ngày tạo': fmtDate(r.createdAt),
      'Tên khách': r.tenKhach,
      'Giá phòng': r.giaPhong,
      'Ngày giờ check in': r.ngayGioCheckIn,
      'Trạng thái TT': r.trangThai,
      'Phòng': r.phong,
      'Người sale': r.nguoiSale,
      'Đã thanh toán': r.daThanhToan,
      'Note': r.note,
    }))

    // Thêm dòng tổng cộng vào cuối file Excel
    const exportRows = [
      ...data,
      {
        'STT': 'Tổng',
        'Ngày tạo': '',
        'Tên khách': `${currentRows.length} khách`,
        'Giá phòng': tongGiaPhong.toLocaleString('vi-VN'),
        'Ngày giờ check in': '',
        'Trạng thái TT': '',
        'Phòng': '',
        'Người sale': '',
        'Đã thanh toán': '',
        'Note': '',
      },
    ]

    const ws = XLSX.utils.json_to_sheet(exportRows)
    ws['!cols'] = [
      { wch: 5 },
      { wch: 18 },
      { wch: 22 },
      { wch: 15 },
      { wch: 30 },
      { wch: 18 },
      { wch: 18 },
      { wch: 15 },
      { wch: 15 },
      { wch: 20 },
    ]

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Danh sách khách')
    XLSX.writeFile(wb, 'danh_sach_khach.xlsx')
  }

  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize))
  const currentPage = Math.min(page, totalPages)
  const startIdx = (currentPage - 1) * pageSize
  const pagedRows = useMemo(
    () => rows.slice(startIdx, startIdx + pageSize).map((r, i) => ({ r, globalIdx: startIdx + i })),
    [rows, startIdx, pageSize]
  )

  const percent = storage.quota ? Math.min(100, (storage.usage / storage.quota) * 100) : 0
  const overThreshold = percent >= WARN_PERCENT
  const freeBytes = Math.max(0, storage.quota - storage.usage)

  const goToPage = (p) => setPage(Math.min(totalPages, Math.max(1, p)))

  // Tính tổng cột tiền ngoài cùng bên trái (cột Giá phòng)
  const tongGiaPhong = useMemo(() => {
    return rows.reduce((sum, r) => sum + parseMoney(r.giaPhong), 0)
  }, [rows])

  return (
    <div className="container">
      <h1>Quản lý chốt khách</h1>

      {overThreshold && (
        <div className="banner-warn">
          ⚠️ Dung lượng lưu trữ đã dùng <b>{percent.toFixed(0)}%</b>. Nên <b>Tải Excel</b> để
          sao lưu rồi xóa bớt dữ liệu cũ.
          <button className="btn-export sm" onClick={handleExport}>Tải Excel ngay</button>
        </div>
      )}

      <div className="input-section">
        <label>Dán form chốt khách vào đây:</label>
        <textarea
          value={text}
          onChange={e => { setText(e.target.value); if (error) setError('') }}
          rows={10}
          placeholder={`Form chốt khách: Thông báo chốt khách ngày\nCheck in: 20h 14/9\nCheck out: 22h 14/9\nGiá : 149\nĐã thanh toán: 149\nNgười sale: ly\nTên CCCD: hồng ngọc\nMã Phòng Thuê: 302 404 bưởi`}
        />
        <button className="btn-add" onClick={handleAdd}>+ Thêm khách</button>
        {error && <p className="error-msg">{error}</p>}
      </div>

      {loading ? (
        <p className="loading-msg">Đang tải dữ liệu từ cơ sở dữ liệu trình duyệt...</p>
      ) : rows.length > 0 ? (
        <div className="table-section">
          <div className="table-header">
            <h2>
              Danh sách khách ({rows.length})
              <span className="header-total"> · Tổng tiền: <b>{tongGiaPhong.toLocaleString('vi-VN')} đ</b></span>
            </h2>
            <div className="header-actions">
              <button className="btn-save-all" onClick={handleManualSave}>
                💾 Lưu{dirty ? ' *' : ''}
              </button>
              <button className="btn-export" onClick={handleExport}>
                Tổng hợp & Tải Excel
              </button>
              <button className="btn-clear" onClick={handleClearAll}>
                Xóa tất cả
              </button>
            </div>
          </div>

          <div className="status-line">
            {dirty
              ? <span className="dirty">● Có thay đổi chưa lưu</span>
              : <span className="saved">● Đã lưu{lastSaved ? ` lúc ${fmtDate(lastSaved.toISOString())}` : ''}</span>}
          </div>

          <div className="maintenance-bar">
            <label>Xóa record tạo trước ngày:</label>
            <input
              type="date"
              value={deleteBeforeDate}
              onChange={e => setDeleteBeforeDate(e.target.value)}
            />
            <button className="btn-delete-before" onClick={handleDeleteBefore}>
              Xóa theo ngày
            </button>
          </div>

          <div className="storage-bar">
            <div className="storage-track">
              <div
                className={`storage-fill ${overThreshold ? 'warn' : ''}`}
                style={{ width: `${percent}%` }}
              />
            </div>
            <span className="storage-text">
              Đã dùng {fmtBytes(storage.usage)}
              {storage.quota ? ` / ${fmtBytes(storage.quota)} (${percent.toFixed(1)}%)` : ''} ·
              còn trống ~{fmtBytes(freeBytes)} · {rows.length.toLocaleString('vi-VN')} record (IndexedDB)
            </span>
          </div>

          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>STT</th>
                  <th>Ngày tạo</th>
                  <th>Tên khách</th>
                  <th>Giá phòng</th>
                  <th>Ngày giờ check in</th>
                  <th>Trạng thái TT</th>
                  <th>Phòng</th>
                  <th>Người sale</th>
                  <th>Đã thanh toán</th>
                  <th>Note</th>
                  <th>Thao tác</th>
                </tr>
              </thead>
              <tbody>
                {pagedRows.map(({ r, globalIdx: i }) => (
                  <tr key={r.id || i}>
                    <td>{i + 1}</td>
                    <td className="created-cell">{fmtDate(r.createdAt)}</td>
                    {editId === r.id ? (
                      <>
                        <td><input value={r.tenKhach} onChange={e => handleEdit(r.id, 'tenKhach', e.target.value)} /></td>
                        <td><input value={r.giaPhong} onChange={e => handleEdit(r.id, 'giaPhong', e.target.value)} /></td>
                        <td><input value={r.ngayGioCheckIn} onChange={e => handleEdit(r.id, 'ngayGioCheckIn', e.target.value)} /></td>
                        <td>
                          <select
                            value={r.trangThai}
                            onChange={e => handleEdit(r.id, 'trangThai', e.target.value)}
                            style={{ padding: '4px', borderRadius: '4px', fontSize: '12px' }}
                          >
                            <option value="Đã TT đủ">Đã TT đủ</option>
                            <option value="Đã TT 1 phần">Đã TT 1 phần</option>
                            <option value="Chưa thanh toán">Chưa thanh toán</option>
                          </select>
                        </td>
                        <td><input value={r.phong} onChange={e => handleEdit(r.id, 'phong', e.target.value)} /></td>
                        <td><input value={r.nguoiSale} onChange={e => handleEdit(r.id, 'nguoiSale', e.target.value)} /></td>
                        <td><input value={r.daThanhToan} onChange={e => handleEdit(r.id, 'daThanhToan', e.target.value)} /></td>
                        <td><input value={r.note} onChange={e => handleEdit(r.id, 'note', e.target.value)} /></td>
                        <td>
                          <button className="btn-save" onClick={() => handleRowSaveDone(r)}>Lưu</button>
                        </td>
                      </>
                    ) : (
                      <>
                        <td>{r.tenKhach}</td>
                        <td>{r.giaPhong}</td>
                        <td>{r.ngayGioCheckIn}</td>
                        <td>
                          <span className={`status ${r.trangThai === 'Đã TT đủ' ? 'paid' : r.trangThai === 'Đã TT 1 phần' ? 'partial' : 'unpaid'}`}>
                            {r.trangThai}
                          </span>
                        </td>
                        <td>{r.phong}</td>
                        <td>{r.nguoiSale}</td>
                        <td>{r.daThanhToan}</td>
                        <td>{r.note}</td>
                        <td className="actions">
                          <button className="btn-edit" onClick={() => setEditId(r.id)}>Sửa</button>
                          <button className="btn-delete" onClick={() => handleDelete(r.id)}>Xóa</button>
                        </td>
                      </>
                    )}
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="summary-row">
                  <td colSpan={3} style={{ textAlign: 'right', fontWeight: 600 }}>
                    Tổng cộng:
                  </td>
                  <td className="total-cell">
                    {tongGiaPhong.toLocaleString('vi-VN')} đ
                  </td>
                  <td colSpan={7}></td>
                </tr>
              </tfoot>
            </table>
          </div>

          <div className="pagination">
            <div className="page-size">
              <label>Số dòng/trang:</label>
              <select
                value={pageSize}
                onChange={e => { setPageSize(Number(e.target.value)); setPage(1) }}
              >
                {PAGE_SIZE_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div className="page-nav">
              <button onClick={() => goToPage(1)} disabled={currentPage === 1}>«</button>
              <button onClick={() => goToPage(currentPage - 1)} disabled={currentPage === 1}>‹</button>
              <span>Trang {currentPage} / {totalPages}</span>
              <button onClick={() => goToPage(currentPage + 1)} disabled={currentPage === totalPages}>›</button>
              <button onClick={() => goToPage(totalPages)} disabled={currentPage === totalPages}>»</button>
            </div>
          </div>
        </div>
      ) : (
        <p className="loading-msg">Chưa có dữ liệu. Hãy dán form và bấm “+ Thêm khách”.</p>
      )}
    </div>
  )
}

export default App
