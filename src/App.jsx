import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

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
  if (daThanhToan && gia && daThanhToan === gia) {
    trangThai = 'Đã TT đủ'
  } else if (daThanhToan && parseFloat(daThanhToan) > 0) {
    trangThai = 'Đã TT 1 phần'
  }

  const formatGia = (val) => {
    const num = parseFloat(val)
    if (isNaN(num)) return val
    return (num * 1000).toLocaleString('vi-VN') + ' đ'
  }

  const ngayGioCheckIn = checkIn && checkOut ? `${checkIn} - ${checkOut}` : checkIn

  return {
    tenKhach: tenCCCD,
    giaPhong: formatGia(gia),
    ngayGioCheckIn,
    trangThai,
    phong: maPhong,
    nguoiSale,
    daThanhToan: daThanhToan ? formatGia(daThanhToan) : '',
    note: '',
  }
}

function App() {
  const [text, setText] = useState('')
  const [rows, setRows] = useState([])
  const [editIdx, setEditIdx] = useState(null)
  const [error, setError] = useState('')
  const rowsRef = useRef(rows)

  const updateRows = (newRows) => {
    rowsRef.current = newRows
    setRows(newRows)
  }

  const validateText = (text) => {
    if (!text.trim()) return 'Vui lòng dán form chốt khách vào ô nhập liệu.'

    const lines = text.split('\n').map(l => l.trim()).filter(Boolean)
    const has = (key) => lines.some(l => l.toLowerCase().startsWith(key.toLowerCase()))

    const missing = []
    if (!has('Check in')) missing.push('Check in')
    if (!has('Check out')) missing.push('Check out')
    if (!has('Giá')) missing.push('Giá')
    if (!has('Tên CCCD')) missing.push('Tên CCCD')
    if (!has('Mã Phòng Thuê') && !has('Mã phòng thuê')) missing.push('Mã Phòng Thuê')

    if (missing.length > 0) return `Thiếu trường bắt buộc: ${missing.join(', ')}`

    const get = (key) => {
      const line = lines.find(l => l.toLowerCase().startsWith(key.toLowerCase()))
      return line ? line.substring(line.indexOf(':') + 1).trim() : ''
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
    updateRows([...rowsRef.current, parsed])
    setText('')
  }

  const handleEdit = (idx, field, value) => {
    const updated = [...rowsRef.current]
    updated[idx] = { ...updated[idx], [field]: value }
    updateRows(updated)
  }

  const handleDelete = (idx) => {
    updateRows(rowsRef.current.filter((_, i) => i !== idx))
    if (editIdx === idx) setEditIdx(null)
  }

  const handleExport = () => {
    setEditIdx(null)
    const currentRows = rowsRef.current
    if (currentRows.length === 0) return

    const data = currentRows.map((r, i) => ({
      'STT': i + 1,
      'Tên khách': r.tenKhach,
      'Giá phòng': r.giaPhong,
      'Ngày giờ check in': r.ngayGioCheckIn,
      'Trạng thái TT': r.trangThai,
      'Phòng': r.phong,
      'Người sale': r.nguoiSale,
      'Đã thanh toán': r.daThanhToan,
      'Note': r.note,
    }))

    const ws = XLSX.utils.json_to_sheet(data)
    ws['!cols'] = [
      { wch: 5 },
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

  const trangThaiOptions = ['Chưa thanh toán', 'Đã TT đủ', 'Đã TT 1 phần']

  return (
    <div className="container">
      <h1>Quản lý chốt khách</h1>

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

      {rows.length > 0 && (
        <div className="table-section">
          <div className="table-header">
            <h2>Danh sách khách ({rows.length})</h2>
            <button className="btn-export" onClick={handleExport}>
              Tổng hợp & Tải Excel
            </button>
          </div>

          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>STT</th>
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
                {rows.map((r, i) => (
                  <tr key={i}>
                    <td>{i + 1}</td>
                    {editIdx === i ? (
                      <>
                        <td><input value={r.tenKhach} onChange={e => handleEdit(i, 'tenKhach', e.target.value)} /></td>
                        <td><input value={r.giaPhong} onChange={e => handleEdit(i, 'giaPhong', e.target.value)} /></td>
                        <td><input value={r.ngayGioCheckIn} onChange={e => handleEdit(i, 'ngayGioCheckIn', e.target.value)} /></td>
                        <td>
                          <select value={r.trangThai} onChange={e => handleEdit(i, 'trangThai', e.target.value)}>
                            {trangThaiOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        </td>
                        <td><input value={r.phong} onChange={e => handleEdit(i, 'phong', e.target.value)} /></td>
                        <td><input value={r.nguoiSale} onChange={e => handleEdit(i, 'nguoiSale', e.target.value)} /></td>
                        <td><input value={r.daThanhToan} onChange={e => handleEdit(i, 'daThanhToan', e.target.value)} /></td>
                        <td><input value={r.note} onChange={e => handleEdit(i, 'note', e.target.value)} /></td>
                        <td>
                          <button className="btn-save" onClick={() => setEditIdx(null)}>Lưu</button>
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
                          <button className="btn-edit" onClick={() => setEditIdx(i)}>Sửa</button>
                          <button className="btn-delete" onClick={() => handleDelete(i)}>Xóa</button>
                        </td>
                      </>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}

export default App
