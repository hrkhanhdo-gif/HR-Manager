// =============================================================================
// HRIS GIAI ĐOẠN 2 — UI/UX HIỆN ĐẠI + DASHBOARD
// Dán file này vào Apps Script Editor, BÊN DƯỚI file Giai Đoạn 1
// =============================================================================
// MODULE 10 — doGet() override (thay thế hàm doGet ở Giai đoạn 1)
// MODULE 11 — HTML Template chính (sidebar + multi-view)
// MODULE 12 — API handlers mới (ứng viên, phỏng vấn, tuyển dụng)
// MODULE 13 — Chart data builders
// =============================================================================

// GHI CHÚ: Xóa hàm doGet() và buildMainApp() ở file Giai đoạn 1,
// thay bằng phiên bản mới bên dưới


// =============================================================================
// MODULE 10 — doGet() PHIÊN BẢN MỚI
// =============================================================================

function doGet(e) {
  const user = getCurrentUser();

  if (!user.email || !user.role) {
    return HtmlService.createHtmlOutput(buildAccessDeniedPage(user.email))
      .setTitle('HR System — Không có quyền')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  writeAuditLog('LOGIN', 'SYSTEM', user.email, '', '', user.role,
    'Truy cập lúc ' + new Date().toLocaleString('vi-VN'));

  const page = (e && e.parameter && e.parameter.page) || 'dashboard';

  return HtmlService.createHtmlOutput(buildFullApp(user, page))
    .setTitle('HR Workspace')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// handleRequest mở rộng — thêm vào switch của Giai Đoạn 1
function handleRequest(request) {
  const user = getCurrentUser();
  if (!user.email || !user.role) {
    return { success: false, error: 'Không có quyền truy cập.' };
  }
  const { action, payload } = request;
  try {
    switch (action) {
      // === Từ Giai đoạn 1 ===
      case 'GET_NHAN_SU_LIST':
        if (!hasPermission(user.role, 'READ')) throw new Error('Không có quyền');
        return { success: true, data: getNhanSuList(user) };
      case 'CREATE_NHAN_SU':
        if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
        return createNhanSu(payload, user);
      case 'UPDATE_NHAN_SU':
        if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
        return updateNhanSu(payload, user);
      case 'DELETE_NHAN_SU':
        if (!hasPermission(user.role, 'DELETE')) throw new Error('Không có quyền');
        return deleteNhanSu(payload.id, user);
      case 'GET_DASHBOARD_STATS':
        if (!hasPermission(user.role, 'READ')) throw new Error('Không có quyền');
        return { success: true, data: getDashboardStats(user) };
      case 'GET_AUDIT_LOG':
        if (!hasPermission(user.role, 'ADMIN')) throw new Error('Chỉ Admin');
        return { success: true, data: getAuditLog(payload) };

      // === Mới — Giai đoạn 2 ===
      case 'GET_UNG_VIEN_LIST':
        return { success: true, data: getUngVienList(user) };
      case 'CREATE_UNG_VIEN':
        if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
        return createUngVien(payload, user);
      case 'UPDATE_UNG_VIEN_STATUS':
        if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
        return updateUngVienStatus(payload, user);
      case 'GET_PHONG_VAN_LIST':
        return { success: true, data: getPhongVanList(user) };
      case 'CREATE_PHONG_VAN':
        if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
        return createPhongVan(payload, user);
      case 'GET_CHART_DATA':
        return { success: true, data: getChartData(user) };
      case 'GET_SINH_NHAT':
        return { success: true, data: getSinhNhatTuanNay() };

      default:
        return { success: false, error: 'Action không hợp lệ: ' + action };
    }
  } catch (e) {
    writeAuditLog('ERROR', 'SYSTEM', '', action, '', '', e.message);
    return { success: false, error: e.message };
  }
}


// =============================================================================
// MODULE 12 — DATA ACCESS: ỨNG VIÊN, PHỎNG VẤN
// =============================================================================

function getUngVienList(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => h.toString());
  return data.slice(1).filter(r => r[0] !== '').map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]; });
    return obj;
  });
}

function createUngVien(data, user) {
  const validation = validateUngVien(data);
  if (!validation.valid) return { success: false, error: validation.errors.join('\n') };
  const id = generateId('UV');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN);
  if (!sheet) return { success: false, error: 'Không tìm thấy sheet UNG_VIEN' };
  sheet.appendRow([
    id, data.ho_ten, data.vi_tri_tuyen || '', data.nguon_cv || '',
    data.ngay_nop ? new Date(data.ngay_nop) : new Date(),
    data.trang_thai || 'Mới nộp',
    data.email || '', data.sdt || '', data.link_cv || '', data.ghi_chu || '',
    new Date(), user.email
  ]);
  writeAuditLog('CREATE', CONFIG.SHEETS.UNG_VIEN, id, '', '', data.ho_ten, 'Thêm ứng viên mới');
  return { success: true, id, message: 'Đã thêm ứng viên ' + data.ho_ten };
}

function updateUngVienStatus(data, user) {
  if (!data.id || !data.trang_thai) return { success: false, error: 'Thiếu ID hoặc trạng thái' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN);
  if (!sheet) return { success: false, error: 'Sheet không tồn tại' };
  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      const oldStatus = allData[i][5];
      sheet.getRange(i + 1, 6).setValue(data.trang_thai);
      if (data.ghi_chu) sheet.getRange(i + 1, 10).setValue(data.ghi_chu);
      writeAuditLog('UPDATE', CONFIG.SHEETS.UNG_VIEN, data.id, 'Trang_Thai', oldStatus, data.trang_thai, '');

      // Nếu chuyển sang Thử việc → tự động tạo bản ghi Nhân sự draft
      if (data.trang_thai === 'Thử việc') {
        autoPromoteToNhanSu(allData[i], user);
      }
      return { success: true, message: 'Đã cập nhật trạng thái' };
    }
  }
  return { success: false, error: 'Không tìm thấy ứng viên' };
}

function autoPromoteToNhanSu(uvRow, user) {
  try {
    createNhanSu({
      ho_ten: uvRow[1], bo_phan: '', chuc_vu: uvRow[2],
      ngay_vao: new Date(), email: uvRow[6], sdt: uvRow[7],
      trang_thai: 'Thử việc', link_ho_so: uvRow[8]
    }, user);
    writeAuditLog('CREATE', CONFIG.SHEETS.NHAN_SU, '', 'AutoPromote', uvRow[0], uvRow[1], 'Tự động tạo từ ứng viên thử việc');
  } catch (e) {}
}

function getPhongVanList(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PHONG_VAN);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => h.toString());
  return data.slice(1).filter(r => r[0] !== '').map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]; });
    return obj;
  });
}

function createPhongVan(data, user) {
  if (!data.id_ung_vien) return { success: false, error: 'Thiếu ID ứng viên' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PHONG_VAN);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PHONG_VAN);
    const headers = ['ID_PV','ID_UngVien','Ho_Ten_UV','Ngay_PV','Gio_PV','Nguoi_PV','Hinh_Thuc','Ket_Qua','Ghi_Chu','Created_By'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setBackground('#712B13').setFontColor('#FFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  const id = generateId('PV');
  sheet.appendRow([
    id, data.id_ung_vien, data.ho_ten_uv || '',
    data.ngay_pv ? new Date(data.ngay_pv) : '',
    data.gio_pv || '', data.nguoi_pv || '',
    data.hinh_thuc || 'Trực tiếp',
    data.ket_qua || 'Chờ kết quả',
    data.ghi_chu || '', user.email
  ]);
  // Cập nhật trạng thái ứng viên → Hẹn PV
  updateUngVienStatus({ id: data.id_ung_vien, trang_thai: 'Hẹn PV' }, user);
  writeAuditLog('CREATE', CONFIG.SHEETS.PHONG_VAN, id, '', '', data.id_ung_vien, 'Tạo lịch PV');
  return { success: true, id, message: 'Đã tạo lịch phỏng vấn' };
}


// =============================================================================
// MODULE 13 — CHART DATA
// =============================================================================

function getChartData(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  // 1. Nhân sự theo bộ phận
  try {
    const nsData = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU).getDataRange().getValues().slice(1)
      .filter(r => r[0] && r[8] !== 'Đã nghỉ');
    const byDept = {};
    nsData.forEach(r => { const d = r[2] || 'Chưa phân'; byDept[d] = (byDept[d] || 0) + 1; });
    result.nhan_su_theo_bo_phan = byDept;
  } catch(e) { result.nhan_su_theo_bo_phan = {}; }

  // 2. Ứng viên theo trạng thái (pipeline)
  try {
    const uvData = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN).getDataRange().getValues().slice(1)
      .filter(r => r[0]);
    const byStatus = { 'Mới nộp':0, 'Đang xét':0, 'Hẹn PV':0, 'Thử việc':0, 'Chính thức':0, 'Từ chối':0 };
    uvData.forEach(r => { const s = r[5] || 'Mới nộp'; if(byStatus[s] !== undefined) byStatus[s]++; else byStatus[s] = 1; });
    result.ung_vien_theo_trang_thai = byStatus;
  } catch(e) { result.ung_vien_theo_trang_thai = {}; }

  // 3. Tuyển dụng 6 tháng gần nhất
  try {
    const nsData = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU).getDataRange().getValues().slice(1).filter(r => r[0]);
    const monthly = {};
    const now = new Date();
    for (let m = 5; m >= 0; m--) {
      const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
      const key = (d.getMonth()+1) + '/' + d.getFullYear();
      monthly[key] = 0;
    }
    nsData.forEach(r => {
      if (r[4]) {
        const d = new Date(r[4]);
        const key = (d.getMonth()+1) + '/' + d.getFullYear();
        if (monthly[key] !== undefined) monthly[key]++;
      }
    });
    result.tuyen_dung_6_thang = monthly;
  } catch(e) { result.tuyen_dung_6_thang = {}; }

  return result;
}


// =============================================================================
// MODULE 11 — HTML FULL APP
// =============================================================================

function buildFullApp(user, activePage) {
  const canWrite = hasPermission(user.role, 'WRITE');
  const isAdmin  = hasPermission(user.role, 'ADMIN');

  return `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>HR Workspace</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600&family=Playfair+Display:wght@600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
  :root {
    --bg:       #F7F6F3;
    --surface:  #FFFFFF;
    --border:   rgba(0,0,0,0.08);
    --text-1:   #1A1A1A;
    --text-2:   #555550;
    --text-3:   #999;
    --accent:   #3D35A8;
    --accent-l: #EEEDFE;
    --green:    #1D9E75;
    --green-l:  #E1F5EE;
    --amber:    #BA7517;
    --amber-l:  #FAEEDA;
    --red:      #C0392B;
    --red-l:    #FCEBEB;
    --coral-l:  #FAECE7;
    --sidebar-w: 220px;
    --radius:   10px;
  }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Be Vietnam Pro',sans-serif; background:var(--bg); color:var(--text-1); font-size:14px; min-height:100vh; display:flex; }

  /* SIDEBAR */
  .sidebar {
    width: var(--sidebar-w); min-height:100vh; background:var(--text-1);
    display:flex; flex-direction:column; position:fixed; left:0; top:0; bottom:0; z-index:100;
  }
  .sidebar-logo {
    padding: 22px 20px 18px;
    border-bottom: 1px solid rgba(255,255,255,0.08);
  }
  .sidebar-logo .logo-text {
    font-family:'Playfair Display',serif; font-size:18px; color:#fff; letter-spacing:0.02em;
  }
  .sidebar-logo .logo-sub { font-size:10px; color:rgba(255,255,255,0.4); letter-spacing:0.1em; text-transform:uppercase; margin-top:2px; }
  .nav { flex:1; padding:12px 0; }
  .nav-item {
    display:flex; align-items:center; gap:10px;
    padding:10px 20px; color:rgba(255,255,255,0.55); font-size:13px;
    cursor:pointer; transition:all 0.15s; border-left:3px solid transparent;
    user-select:none;
  }
  .nav-item:hover { color:#fff; background:rgba(255,255,255,0.05); }
  .nav-item.active { color:#fff; background:rgba(255,255,255,0.1); border-left-color: #8B7FF0; }
  .nav-icon { font-size:15px; width:18px; text-align:center; }
  .nav-section { font-size:10px; color:rgba(255,255,255,0.25); letter-spacing:0.1em; text-transform:uppercase; padding:16px 20px 6px; }
  .sidebar-user {
    padding:16px 20px; border-top:1px solid rgba(255,255,255,0.08);
    color:rgba(255,255,255,0.6); font-size:12px;
  }
  .sidebar-user .uname { color:#fff; font-weight:500; font-size:13px; }
  .role-badge {
    display:inline-block; font-size:9px; padding:2px 7px; border-radius:20px; margin-top:4px;
    background:rgba(141,127,240,0.3); color:#c4beff; letter-spacing:0.05em; text-transform:uppercase;
  }

  /* MAIN */
  .main { margin-left:var(--sidebar-w); flex:1; min-height:100vh; display:flex; flex-direction:column; }
  .topbar {
    background:var(--surface); border-bottom:1px solid var(--border);
    padding:14px 28px; display:flex; align-items:center; justify-content:space-between;
    position:sticky; top:0; z-index:50;
  }
  .page-title { font-size:16px; font-weight:600; color:var(--text-1); }
  .topbar-actions { display:flex; gap:8px; align-items:center; }
  .content { padding:28px; flex:1; }

  /* Các style còn lại giữ nguyên từ GĐ2 (metrics, cards, tables, modals, toast, ...) */
  .metrics-row { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin-bottom:24px; }
  .metric-card {
    background:var(--surface); border-radius:var(--radius); padding:20px 22px;
    border:1px solid var(--border); position:relative; overflow:hidden;
  }
  .metric-card::before {
    content:''; position:absolute; top:0; left:0; right:0; height:3px;
  }
  .metric-card.purple::before { background:var(--accent); }
  .metric-card.green::before  { background:var(--green); }
  .metric-card.amber::before  { background:var(--amber); }
  .metric-card.red::before    { background:var(--red); }
  .metric-label { font-size:11px; color:var(--text-3); text-transform:uppercase; letter-spacing:0.07em; margin-bottom:10px; }
  .metric-value { font-size:32px; font-weight:300; color:var(--text-1); line-height:1; }
  .metric-delta { font-size:12px; margin-top:8px; color:var(--text-3); }
  .metric-delta.up   { color:var(--green); }
  .metric-delta.down { color:var(--red); }
  .metric-icon { position:absolute; right:18px; top:18px; font-size:22px; opacity:0.15; }

  .grid-2 { display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-bottom:20px; }
  .grid-3 { display:grid; grid-template-columns:2fr 1fr; gap:20px; margin-bottom:20px; }
  .panel {
    background:var(--surface); border-radius:var(--radius);
    border:1px solid var(--border); overflow:hidden;
  }
  .panel-header {
    padding:16px 20px; border-bottom:1px solid var(--border);
    display:flex; align-items:center; justify-content:space-between;
  }
  .panel-title { font-size:13px; font-weight:600; color:var(--text-1); }
  .panel-body { padding:16px 20px; }

  .data-table { width:100%; border-collapse:collapse; }
  .data-table th {
    text-align:left; font-size:11px; color:var(--text-3); font-weight:500;
    text-transform:uppercase; letter-spacing:0.07em;
    padding:10px 12px; border-bottom:1px solid var(--border); white-space:nowrap;
  }
  .data-table td { padding:10px 12px; font-size:13px; border-bottom:1px solid var(--border); vertical-align:middle; }
  .data-table tr:last-child td { border-bottom:none; }
  .data-table tr:hover td { background:var(--bg); }
  .data-table .name-cell { font-weight:500; }
  .avatar {
    width:30px; height:30px; border-radius:50%; display:inline-flex;
    align-items:center; justify-content:center; font-size:11px; font-weight:600;
    margin-right:8px; flex-shrink:0; vertical-align:middle;
  }

  .status {
    display:inline-block; font-size:11px; padding:3px 9px; border-radius:20px; font-weight:500; white-space:nowrap;
  }
  .s-chinh-thuc { background:var(--green-l);  color:#085041; }
  .s-thu-viec   { background:var(--accent-l); color:#3C3489; }
  .s-moi-nop    { background:#E6F1FB; color:#0C447C; }
  .s-dang-xet   { background:var(--amber-l);  color:#633806; }
  .s-hen-pv     { background:var(--accent-l); color:#3C3489; }
  .s-tu-choi    { background:var(--red-l);    color:#791F1F; }
  .s-da-nghi    { background:#F1EFE8; color:#444; }

  .btn {
    display:inline-flex; align-items:center; gap:6px; padding:7px 14px;
    border-radius:7px; font-size:13px; font-weight:500; cursor:pointer;
    border:none; font-family:inherit; transition:all 0.15s;
  }
  .btn-primary { background:var(--accent); color:#fff; }
  .btn-primary:hover { background:#2d279a; }
  .btn-ghost { background:transparent; color:var(--text-2); border:1px solid var(--border); }
  .btn-ghost:hover { background:var(--bg); }
  .btn-sm { padding:5px 10px; font-size:12px; }
  .btn-danger { background:var(--red-l); color:var(--red); border:1px solid rgba(192,57,43,0.2); }

  .overlay {
    position:fixed; inset:0; background:rgba(0,0,0,0.4); z-index:200;
    display:none; align-items:center; justify-content:center;
  }
  .overlay.show { display:flex; }
  .modal {
    background:var(--surface); border-radius:14px; width:500px; max-width:90vw;
    max-height:90vh; overflow-y:auto; box-shadow:0 20px 60px rgba(0,0,0,0.2);
  }
  .modal-header {
    padding:20px 24px; border-bottom:1px solid var(--border);
    display:flex; align-items:center; justify-content:space-between;
  }
  .modal-title { font-size:15px; font-weight:600; }
  .modal-close { cursor:pointer; font-size:20px; color:var(--text-3); line-height:1; }
  .modal-body { padding:24px; }
  .modal-footer { padding:16px 24px; border-top:1px solid var(--border); display:flex; gap:8px; justify-content:flex-end; }
  .form-row { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-bottom:14px; }
  .form-row.full { grid-template-columns:1fr; }
  .form-group { display:flex; flex-direction:column; gap:5px; }
  .form-label { font-size:11px; font-weight:600; color:var(--text-2); text-transform:uppercase; letter-spacing:0.06em; }
  .form-control {
    padding:8px 11px; border:1px solid var(--border); border-radius:7px;
    font-family:inherit; font-size:13px; color:var(--text-1); background:var(--bg);
    outline:none; transition:border 0.15s;
  }
  .form-control:focus { border-color:var(--accent); background:#fff; }
  select.form-control { cursor:pointer; }

  .pipeline { display:flex; align-items:center; gap:0; margin:12px 0; }
  .pipe-step {
    flex:1; text-align:center; padding:8px 4px; background:var(--bg);
    border-top:1px solid var(--border); border-bottom:1px solid var(--border);
    border-right:1px solid var(--border); font-size:11px; color:var(--text-2); position:relative;
  }
  .pipe-step:first-child { border-left:1px solid var(--border); border-radius:6px 0 0 6px; }
  .pipe-step:last-child  { border-radius:0 6px 6px 0; }
  .pipe-step .pipe-count { font-size:20px; font-weight:300; color:var(--text-1); display:block; }
  .pipe-step.active { background:var(--accent-l); border-color:var(--accent); }
  .pipe-step.active .pipe-count { color:var(--accent); }

  .bday-item { display:flex; align-items:center; gap:12px; padding:10px 0; border-bottom:1px solid var(--border); }
  .bday-item:last-child { border-bottom:none; }
  .bday-days { font-size:11px; color:var(--text-3); min-width:50px; text-align:right; }
  .bday-days.today { color:var(--amber); font-weight:600; }

  .audit-action {
    display:inline-block; font-size:10px; padding:2px 7px; border-radius:4px;
    font-weight:600; text-transform:uppercase; letter-spacing:0.05em;
  }
  .a-CREATE { background:var(--green-l); color:var(--green); }
  .a-UPDATE { background:var(--amber-l); color:var(--amber); }
  .a-DELETE { background:var(--red-l); color:var(--red); }
  .a-LOGIN  { background:#E6F1FB; color:#185FA5; }
  .a-ERROR  { background:var(--coral-l); color:#712B13; }

  .toast {
    position:fixed; bottom:24px; right:24px; z-index:999;
    background:var(--text-1); color:#fff; padding:12px 18px; border-radius:9px;
    font-size:13px; transform:translateY(80px); opacity:0;
    transition:all 0.3s; pointer-events:none; max-width:320px;
  }
  .toast.show { transform:translateY(0); opacity:1; }
  .toast.success { background:#1D9E75; }
  .toast.error   { background:#C0392B; }

  .view { display:none; animation:fadeIn 0.2s ease; }
  .view.active { display:block; }
  @keyframes fadeIn { from{opacity:0;transform:translateY(6px)} to{opacity:1;transform:none} }

  .skeleton { background:linear-gradient(90deg,var(--border) 25%,rgba(0,0,0,0.04) 50%,var(--border) 75%); background-size:200%; animation:shimmer 1.2s infinite; border-radius:4px; }
  @keyframes shimmer { 0%{background-position:200%} 100%{background-position:-200%} }

  @media(max-width:900px) {
    .metrics-row { grid-template-columns:1fr 1fr; }
    .grid-2, .grid-3 { grid-template-columns:1fr; }
    .sidebar { display:none; }
    .main { margin-left:0; }
  }

  .empty-state { text-align:center; padding:40px; color:var(--text-3); font-size:13px; }
  .divider { height:1px; background:var(--border); margin:16px 0; }
  .search-input { padding:7px 12px; border:1px solid var(--border); border-radius:7px; font-family:inherit; font-size:13px; outline:none; width:220px; }
  .search-input:focus { border-color:var(--accent); }
  .chart-wrap { position:relative; height:220px; }
</style>
</head>
<body>

<!-- SIDEBAR (ĐÃ CÓ GĐ4 + GĐ5) -->
<aside class="sidebar">
  <div class="sidebar-logo">
    <div class="logo-text">HR Workspace</div>
    <div class="logo-sub">Hệ thống nhân sự</div>
  </div>
  <nav class="nav">
    <div class="nav-section">Tổng quan</div>
    <div class="nav-item ${activePage==='dashboard'?'active':''}" onclick="showView('dashboard')">
      <span class="nav-icon">◈</span> Dashboard
    </div>
    <div class="nav-section">Nhân sự</div>
    <div class="nav-item ${activePage==='nhansu'?'active':''}" onclick="showView('nhansu')">
      <span class="nav-icon">◉</span> Nhân sự
    </div>
    <div class="nav-item ${activePage==='ungvien'?'active':''}" onclick="showView('ungvien')">
      <span class="nav-icon">◎</span> Ứng viên
    </div>
    <div class="nav-item ${activePage==='phongvan'?'active':''}" onclick="showView('phongvan')">
      <span class="nav-icon">◷</span> Phỏng vấn
    </div>

    <!-- GĐ4 -->
    <div class="nav-section">Hiệu suất</div>
    <div class="nav-item ${activePage==='onboarding'?'active':''}" onclick="showView('onboarding')">
      <span class="nav-icon">📋</span> Onboarding
    </div>
    <div class="nav-item ${activePage==='kpi'?'active':''}" onclick="showView('kpi')">
      <span class="nav-icon">🎯</span> KPI & Hiệu suất
    </div>

    <!-- GĐ5 -->
    <div class="nav-section">Tài chính & Báo cáo</div>
    <div class="nav-item ${activePage==='payroll'?'active':''}" onclick="showView('payroll')">
      <span class="nav-icon">💰</span> Bảng lương
    </div>
    <div class="nav-item ${activePage==='leave'?'active':''}" onclick="showView('leave')">
      <span class="nav-icon">🌴</span> Nghỉ phép
    </div>
    <div class="nav-item ${activePage==='reports'?'active':''}" onclick="showView('reports')">
      <span class="nav-icon">📊</span> Báo cáo HR
    </div>

    ${isAdmin ? `
    <div class="nav-section">Quản trị</div>
    <div class="nav-item ${activePage==='auditlog'?'active':''}" onclick="showView('auditlog')">
      <span class="nav-icon">◌</span> Audit Log
    </div>` : ''}
  </nav>
  <div class="sidebar-user">
    <div class="uname">${user.name}</div>
    <div class="role-badge">${user.role}</div>
  </div>
</aside>

<!-- MAIN -->
<main class="main">
  <div class="topbar">
    <span class="page-title" id="page-title">Dashboard</span>
    <div class="topbar-actions">
      ${canWrite ? `
      <button class="btn btn-ghost btn-sm" onclick="showView('ungvien');openModal('modal-uv')">+ Ứng viên</button>
      <button class="btn btn-primary btn-sm" onclick="showView('nhansu');openModal('modal-ns')">+ Nhân sự</button>
      ` : ''}
    </div>
  </div>
  <div class="content">

    <!-- VIEW DASHBOARD -->
    <div id="view-dashboard" class="view active">
      <div class="metrics-row" id="metrics-row">
        <div class="metric-card purple skeleton" style="height:100px"></div>
        <div class="metric-card green skeleton"  style="height:100px"></div>
        <div class="metric-card amber skeleton"  style="height:100px"></div>
        <div class="metric-card red skeleton"    style="height:100px"></div>
      </div>
      <div class="grid-2">
        <div class="panel">
          <div class="panel-header"><span class="panel-title">Nhân sự theo bộ phận</span></div>
          <div class="panel-body"><div class="chart-wrap"><canvas id="chart-dept" role="img" aria-label="Biểu đồ nhân sự theo bộ phận"></canvas></div></div>
        </div>
        <div class="panel">
          <div class="panel-header"><span class="panel-title">Tuyển dụng 6 tháng</span></div>
          <div class="panel-body"><div class="chart-wrap"><canvas id="chart-monthly" role="img" aria-label="Biểu đồ tuyển dụng 6 tháng"></canvas></div></div>
        </div>
      </div>
      <div class="grid-3">
        <div class="panel">
          <div class="panel-header"><span class="panel-title">Pipeline tuyển dụng</span></div>
          <div class="panel-body">
            <div class="pipeline" id="pipeline-bar"></div>
            <div class="divider"></div>
            <table class="data-table" id="uv-recent-table">
              <thead><tr><th>Ứng viên</th><th>Vị trí</th><th>Trạng thái</th></tr></thead>
              <tbody><tr><td colspan="3" class="empty-state">Đang tải...</td></tr></tbody>
            </table>
          </div>
        </div>
        <div class="panel">
          <div class="panel-header"><span class="panel-title">Sinh nhật sắp tới</span></div>
          <div class="panel-body" id="bday-list"><div class="empty-state">Đang tải...</div></div>
        </div>
      </div>
    </div>

    <!-- VIEW NHÂN SỰ -->
    <div id="view-nhansu" class="view">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
        <input type="text" class="search-input" placeholder="Tìm nhân sự..." oninput="filterTable('ns-table',this.value)">
        ${canWrite ? `<button class="btn btn-primary btn-sm" onclick="openModal('modal-ns')">+ Thêm nhân sự</button>` : ''}
      </div>
      <div class="panel">
        <table class="data-table" id="ns-table">
          <thead><tr><th>Nhân sự</th><th>Bộ phận</th><th>Chức vụ</th><th>Ngày vào</th><th>Trạng thái</th>${canWrite?'<th></th>':''}</tr></thead>
          <tbody><tr><td colspan="6" class="empty-state">Đang tải dữ liệu...</td></tr></tbody>
        </table>
      </div>
    </div>

    <!-- VIEW ỨNG VIÊN -->
    <div id="view-ungvien" class="view">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
        <input type="text" class="search-input" placeholder="Tìm ứng viên..." oninput="filterTable('uv-table',this.value)">
        ${canWrite ? `<button class="btn btn-primary btn-sm" onclick="openModal('modal-uv')">+ Thêm ứng viên</button>` : ''}
      </div>
      <div class="panel">
        <table class="data-table" id="uv-table">
          <thead><tr><th>Ứng viên</th><th>Vị trí</th><th>Nguồn</th><th>Ngày nộp</th><th>Trạng thái</th>${canWrite?'<th></th>':''}</tr></thead>
          <tbody><tr><td colspan="6" class="empty-state">Đang tải dữ liệu...</td></tr></tbody>
        </table>
      </div>
    </div>

    <!-- VIEW PHỎNG VẤN -->
    <div id="view-phongvan" class="view">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
        <span style="font-size:13px;color:var(--text-2);">Lịch phỏng vấn sắp tới</span>
        ${canWrite ? `<button class="btn btn-primary btn-sm" onclick="openModal('modal-pv')">+ Tạo lịch PV</button>` : ''}
      </div>
      <div class="panel">
        <table class="data-table" id="pv-table">
          <thead><tr><th>Ứng viên</th><th>Ngày PV</th><th>Giờ</th><th>Người PV</th><th>Hình thức</th><th>Kết quả</th></tr></thead>
          <tbody><tr><td colspan="6" class="empty-state">Đang tải...</td></tr></tbody>
        </table>
      </div>
    </div>

    <!-- VIEW AUDIT LOG -->
    ${isAdmin ? `
    <div id="view-auditlog" class="view">
      <div style="margin-bottom:16px;display:flex;align-items:center;gap:12px;">
        <span style="font-size:13px;color:var(--text-2);">200 bản ghi gần nhất</span>
        <button class="btn btn-ghost btn-sm" onclick="loadAuditLog()">↻ Làm mới</button>
      </div>
      <div class="panel">
        <table class="data-table" id="audit-table">
          <thead><tr><th>Thời gian</th><th>Người dùng</th><th>Action</th><th>Sheet</th><th>Thay đổi</th></tr></thead>
          <tbody><tr><td colspan="5" class="empty-state">Đang tải...</td></tr></tbody>
        </table>
      </div>
    </div>` : ''}

    <!-- GĐ4 TABS -->
    ${getOnboardingTabHtml()}
    ${getKpiTabHtml()}

    <!-- GĐ5 TABS (Payroll + Leave + Reports) -->
    ${getPayrollTabHtml()}
    ${getLeaveTabHtml()}
    ${getReportsTabHtml()}

  </div>
</main>


<!-- ═══════════════════════════════════════════
     MODALS
════════════════════════════════════════════ -->

<!-- Modal Nhân sự -->
<div class="overlay" id="modal-ns">
  <div class="modal">
    <div class="modal-header">
      <span class="modal-title">Thêm nhân sự mới</span>
      <span class="modal-close" onclick="closeModal('modal-ns')">×</span>
    </div>
    <div class="modal-body">
      <div class="form-row">
        <div class="form-group"><label class="form-label">Họ và tên *</label><input class="form-control" id="ns-hoten" placeholder="Nguyễn Văn A"></div>
        <div class="form-group"><label class="form-label">Bộ phận *</label><input class="form-control" id="ns-bophan" placeholder="Sales, Tech, HR..."></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label class="form-label">Chức vụ</label><input class="form-control" id="ns-chucvu" placeholder="Nhân viên, Trưởng nhóm..."></div>
        <div class="form-group"><label class="form-label">Ngày vào làm</label><input class="form-control" type="date" id="ns-ngayvao"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label class="form-label">Email</label><input class="form-control" type="email" id="ns-email" placeholder="email@company.com"></div>
        <div class="form-group"><label class="form-label">Số điện thoại</label><input class="form-control" id="ns-sdt" placeholder="0912345678"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label class="form-label">Ngày sinh</label><input class="form-control" type="date" id="ns-ngaysinh"></div>
        <div class="form-group"><label class="form-label">Trạng thái</label>
          <select class="form-control" id="ns-trangthai">
            <option>Chính thức</option><option>Thử việc</option><option>Thực tập</option>
          </select>
        </div>
      </div>
      <div class="form-row full">
        <div class="form-group"><label class="form-label">Link hồ sơ (Google Drive)</label><input class="form-control" id="ns-link" placeholder="https://drive.google.com/..."></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-ns')">Hủy</button>
      <button class="btn btn-primary" onclick="submitNhanSu()">Lưu nhân sự</button>
    </div>
  </div>
</div>

<!-- Modal Ứng viên -->
<div class="overlay" id="modal-uv">
  <div class="modal">
    <div class="modal-header">
      <span class="modal-title">Thêm ứng viên mới</span>
      <span class="modal-close" onclick="closeModal('modal-uv')">×</span>
    </div>
    <div class="modal-body">
      <div class="form-row">
        <div class="form-group"><label class="form-label">Họ và tên *</label><input class="form-control" id="uv-hoten" placeholder="Trần Thị B"></div>
        <div class="form-group"><label class="form-label">Vị trí ứng tuyển *</label><input class="form-control" id="uv-vitri" placeholder="Sales Executive..."></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label class="form-label">Email</label><input class="form-control" type="email" id="uv-email"></div>
        <div class="form-group"><label class="form-label">SĐT</label><input class="form-control" id="uv-sdt" placeholder="0912345678"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label class="form-label">Nguồn CV</label>
          <select class="form-control" id="uv-nguon">
            <option>LinkedIn</option><option>TopCV</option><option>Referral</option><option>Website</option><option>Khác</option>
          </select>
        </div>
        <div class="form-group"><label class="form-label">Ngày nộp</label><input class="form-control" type="date" id="uv-ngaynop"></div>
      </div>
      <div class="form-row full">
        <div class="form-group"><label class="form-label">Link CV (Google Drive)</label><input class="form-control" id="uv-link" placeholder="https://drive.google.com/..."></div>
      </div>
      <div class="form-row full">
        <div class="form-group"><label class="form-label">Ghi chú</label><textarea class="form-control" id="uv-ghichu" rows="2" style="resize:vertical"></textarea></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-uv')">Hủy</button>
      <button class="btn btn-primary" onclick="submitUngVien()">Lưu ứng viên</button>
    </div>
  </div>
</div>

<!-- Modal Phỏng vấn -->
<div class="overlay" id="modal-pv">
  <div class="modal">
    <div class="modal-header">
      <span class="modal-title">Tạo lịch phỏng vấn</span>
      <span class="modal-close" onclick="closeModal('modal-pv')">×</span>
    </div>
    <div class="modal-body">
      <div class="form-row full">
        <div class="form-group"><label class="form-label">ID Ứng viên *</label><input class="form-control" id="pv-uvid" placeholder="UV-20260427-xxxx"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label class="form-label">Ngày phỏng vấn</label><input class="form-control" type="date" id="pv-ngay"></div>
        <div class="form-group"><label class="form-label">Giờ</label><input class="form-control" type="time" id="pv-gio"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label class="form-label">Người phỏng vấn</label><input class="form-control" id="pv-nguoi"></div>
        <div class="form-group"><label class="form-label">Hình thức</label>
          <select class="form-control" id="pv-hinhthuc">
            <option>Trực tiếp</option><option>Online</option><option>Điện thoại</option>
          </select>
        </div>
      </div>
      <div class="form-row full">
        <div class="form-group"><label class="form-label">Ghi chú</label><textarea class="form-control" id="pv-ghichu" rows="2" style="resize:vertical"></textarea></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-pv')">Hủy</button>
      <button class="btn btn-primary" onclick="submitPhongVan()">Lưu lịch PV</button>
    </div>
  </div>
</div>

<!-- Toast -->
<div class="toast" id="toast"></div>


<!-- ═══════════════════════════════════════════
     JAVASCRIPT
════════════════════════════════════════════ -->
<script>
const PAGE_TITLES = { dashboard:'Dashboard', nhansu:'Nhân sự', ungvien:'Ứng viên', phongvan:'Lịch phỏng vấn', auditlog:'Audit Log' , onboarding: 'Onboarding', kpi: 'KPI & Hiệu suất'};
let _nsData=[], _uvData=[], _pvData=[];

// ... (các hàm cũ giữ nguyên) ...

// ── NAVIGATION (đã update) ──
function showView(id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const targetView = document.getElementById('view-'+id);
  if (targetView) targetView.classList.add('active');

  document.querySelectorAll('.nav-item').forEach(n => {
    if (n.getAttribute('onclick') && n.getAttribute('onclick').includes("'"+id+"'")) {
      n.classList.add('active');
    }
  });

  document.getElementById('page-title').textContent = PAGE_TITLES[id] || id;

  // Load data cho các tab mới
  if(id==='dashboard')  loadDashboard();
  if(id==='nhansu')     loadNhanSu();
  if(id==='ungvien')    loadUngVien();
  if(id==='phongvan')   loadPhongVan();
  if(id==='auditlog')   loadAuditLog();
  if(id==='onboarding') initOnboarding && initOnboarding();   // GĐ4
  if(id==='kpi')        initKpi && initKpi();                 // GĐ4
}

// ── MODALS ──
function openModal(id)  { document.getElementById(id).classList.add('show'); }
function closeModal(id) { document.getElementById(id).classList.remove('show'); }
document.querySelectorAll('.overlay').forEach(o => o.addEventListener('click', e => { if(e.target===o) o.classList.remove('show'); }));

// ── TOAST ──
function toast(msg, type='success') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast show ' + type;
  setTimeout(() => t.className='toast', 3000);
}

// ── FILTER TABLE ──
function filterTable(tableId, q) {
  const rows = document.querySelectorAll('#'+tableId+' tbody tr');
  const ql = q.toLowerCase();
  rows.forEach(r => r.style.display = r.textContent.toLowerCase().includes(ql) ? '' : 'none');
}

// ── AVATAR ──
const AVT_COLORS = ['#534AB7','#1D9E75','#BA7517','#C0392B','#0C447C','#712B13'];
function avatarBg(name, idx) { return AVT_COLORS[(idx||0) % AVT_COLORS.length]; }
function initials(name) {
  const parts = (name||'?').split(' ');
  return (parts[0][0] + (parts[parts.length>1?parts.length-1:0][0]||'')).toUpperCase();
}

// ── FORMAT DATE ──
function fmtDate(v) {
  if(!v) return '—';
  const d = new Date(v);
  if(isNaN(d)) return v;
  return d.toLocaleDateString('vi-VN');
}

// ── STATUS BADGE ──
function statusBadge(s) {
  const map = {
    'Chính thức':'s-chinh-thuc','Thử việc':'s-thu-viec','Thực tập':'s-thu-viec',
    'Mới nộp':'s-moi-nop','Đang xét':'s-dang-xet','Hẹn PV':'s-hen-pv',
    'Từ chối':'s-tu-choi','Đã nghỉ':'s-da-nghi'
  };
  return '<span class="status '+(map[s]||'s-da-nghi')+'">'+s+'</span>';
}

// ────────────────────────────────────────────
// LOAD DASHBOARD
// ────────────────────────────────────────────
function loadDashboard() {
  google.script.run.withSuccessHandler(res => {
    if(!res.success) return;
    const d = res.data;

    // Metrics
    document.getElementById('metrics-row').innerHTML = \`
      <div class="metric-card purple">
        <div class="metric-icon">◉</div>
        <div class="metric-label">Nhân sự hiện tại</div>
        <div class="metric-value">\${d.total_nhan_su||0}</div>
        <div class="metric-delta">Đang làm việc</div>
      </div>
      <div class="metric-card green">
        <div class="metric-icon">◎</div>
        <div class="metric-label">Ứng viên</div>
        <div class="metric-value">\${d.total_ung_vien||0}</div>
        <div class="metric-delta">\${d.thu_viec||0} đang thử việc</div>
      </div>
      <div class="metric-card amber">
        <div class="metric-icon">◷</div>
        <div class="metric-label">Sinh nhật tuần này</div>
        <div class="metric-value">\${(d.sinh_nhat_tuan_nay||[]).length}</div>
        <div class="metric-delta">trong 7 ngày tới</div>
      </div>
      <div class="metric-card red">
        <div class="metric-icon">◌</div>
        <div class="metric-label">Đã nghỉ</div>
        <div class="metric-value">\${d.nghi_viec||0}</div>
        <div class="metric-delta">tổng lịch sử</div>
      </div>\`;

    // Birthday
    const bdays = d.sinh_nhat_tuan_nay || [];
    const bdayEl = document.getElementById('bday-list');
    if(bdays.length===0) { bdayEl.innerHTML='<div class="empty-state">Không có sinh nhật trong 7 ngày tới</div>'; }
    else {
      bdayEl.innerHTML = bdays.map(b => \`
        <div class="bday-item">
          <div class="avatar" style="background:\${avatarBg(b.name,0)};color:#fff;font-size:11px;">\${initials(b.name)}</div>
          <div style="flex:1"><div style="font-weight:500;font-size:13px;">\${b.name}</div>
          <div style="font-size:11px;color:var(--text-3)">\${new Date(b.date).toLocaleDateString('vi-VN')}</div></div>
          <div class="bday-days \${b.days_until===0?'today':''}">\${b.days_until===0?'Hôm nay':b.days_until+' ngày'}</div>
        </div>\`).join('');
    }
  }).handleRequest({ action:'GET_DASHBOARD_STATS', payload:{} });

  // Charts + pipeline
  google.script.run.withSuccessHandler(res => {
    if(!res.success) return;
    renderCharts(res.data);
  }).handleRequest({ action:'GET_CHART_DATA', payload:{} });

  google.script.run.withSuccessHandler(res => {
    if(!res.success) return;
    _uvData = res.data;
    renderPipeline(res.data);
    renderUvRecent(res.data);
  }).handleRequest({ action:'GET_UNG_VIEN_LIST', payload:{} });
}

function renderCharts(data) {
  // Chart bộ phận
  const deptData = data.nhan_su_theo_bo_phan || {};
  const deptLabels = Object.keys(deptData);
  const deptVals   = Object.values(deptData);
  const COLORS = ['#534AB7','#1D9E75','#BA7517','#378ADD','#C0392B','#888780','#0F6E56','#3C3489'];
  if(document.getElementById('chart-dept')) {
    new Chart(document.getElementById('chart-dept'), {
      type:'doughnut',
      data:{ labels:deptLabels, datasets:[{ data:deptVals, backgroundColor:COLORS.slice(0,deptLabels.length), borderWidth:2, borderColor:'#fff' }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ position:'right', labels:{ font:{size:11}, boxWidth:12 } } } }
    });
  }

  // Chart tuyển dụng theo tháng
  const monthData = data.tuyen_dung_6_thang || {};
  if(document.getElementById('chart-monthly')) {
    new Chart(document.getElementById('chart-monthly'), {
      type:'bar',
      data:{ labels:Object.keys(monthData), datasets:[{ data:Object.values(monthData), backgroundColor:'#534AB7', borderRadius:5, borderSkipped:false }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false} }, scales:{ y:{ticks:{stepSize:1},grid:{color:'rgba(0,0,0,0.05)'}}, x:{grid:{display:false}} } }
    });
  }
}

function renderPipeline(uvData) {
  const stages = ['Mới nộp','Đang xét','Hẹn PV','Thử việc','Chính thức'];
  const counts = {};
  stages.forEach(s => counts[s]=0);
  uvData.forEach(r => { if(counts[r.Trang_Thai]!==undefined) counts[r.Trang_Thai]++; });
  const el = document.getElementById('pipeline-bar');
  if(el) el.innerHTML = stages.map(s=>\`
    <div class="pipe-step \${counts[s]>0?'active':''}">
      <span class="pipe-count">\${counts[s]}</span>\${s}
    </div>\`).join('');
}

function renderUvRecent(uvData) {
  const tbody = document.querySelector('#uv-recent-table tbody');
  if(!tbody) return;
  const recent = uvData.slice(-8).reverse();
  if(recent.length===0) { tbody.innerHTML='<tr><td colspan="3" class="empty-state">Chưa có ứng viên</td></tr>'; return; }
  tbody.innerHTML = recent.map((r,i) => \`
    <tr>
      <td><span class="avatar" style="background:\${avatarBg(r.Ho_Ten,i)};color:#fff;width:24px;height:24px;font-size:10px;">\${initials(r.Ho_Ten)}</span>\${r.Ho_Ten}</td>
      <td style="color:var(--text-2)">\${r.Vi_Tri_Tuyen||'—'}</td>
      <td>\${statusBadge(r.Trang_Thai)}</td>
    </tr>\`).join('');
}

// ────────────────────────────────────────────
// LOAD NHÂN SỰ
// ────────────────────────────────────────────
function loadNhanSu() {
  const tbody = document.querySelector('#ns-table tbody');
  tbody.innerHTML = '<tr><td colspan="6" class="empty-state">Đang tải...</td></tr>';
  google.script.run.withSuccessHandler(res => {
    if(!res.success) { toast(res.error,'error'); return; }
    _nsData = res.data;
    if(_nsData.length===0) { tbody.innerHTML='<tr><td colspan="6" class="empty-state">Chưa có nhân sự nào</td></tr>'; return; }
    tbody.innerHTML = _nsData.map((r,i) => \`
      <tr>
        <td><span class="avatar" style="background:\${avatarBg(r.Ho_Ten,i)};color:#fff;">\${initials(r.Ho_Ten)}</span><span class="name-cell">\${r.Ho_Ten}</span></td>
        <td>\${r.Bo_Phan||'—'}</td>
        <td style="color:var(--text-2)">\${r.Chuc_Vu||'—'}</td>
        <td style="color:var(--text-3)">\${fmtDate(r.Ngay_Vao)}</td>
        <td>\${statusBadge(r.Trang_Thai)}</td>
        <td><button class="btn btn-sm btn-danger" onclick="softDeleteNS('\${r.ID_NhanSu}')">Nghỉ</button></td>
      </tr>\`).join('');
  }).handleRequest({ action:'GET_NHAN_SU_LIST', payload:{} });
}

// ────────────────────────────────────────────
// LOAD ỨNG VIÊN
// ────────────────────────────────────────────
function loadUngVien() {
  const tbody = document.querySelector('#uv-table tbody');
  tbody.innerHTML = '<tr><td colspan="6" class="empty-state">Đang tải...</td></tr>';
  google.script.run.withSuccessHandler(res => {
    if(!res.success) { toast(res.error,'error'); return; }
    _uvData = res.data;
    if(_uvData.length===0) { tbody.innerHTML='<tr><td colspan="6" class="empty-state">Chưa có ứng viên nào</td></tr>'; return; }
    tbody.innerHTML = _uvData.map((r,i) => \`
      <tr>
        <td><span class="avatar" style="background:\${avatarBg(r.Ho_Ten,i)};color:#fff;">\${initials(r.Ho_Ten)}</span><span class="name-cell">\${r.Ho_Ten}</span></td>
        <td>\${r.Vi_Tri_Tuyen||'—'}</td>
        <td style="color:var(--text-3)">\${r.Nguon_CV||'—'}</td>
        <td style="color:var(--text-3)">\${fmtDate(r.Ngay_Nop)}</td>
        <td>
          <select class="form-control" style="padding:3px 7px;font-size:12px;width:auto" onchange="changeUVStatus('\${r.ID_UngVien}',this.value)">
            \${['Mới nộp','Đang xét','Hẹn PV','Thử việc','Chính thức','Từ chối'].map(s=>\`<option \${s===r.Trang_Thai?'selected':''}>\${s}</option>\`).join('')}
          </select>
        </td>
        <td><button class="btn btn-sm btn-ghost" onclick="scheduleInterview('\${r.ID_UngVien}','\${r.Ho_Ten}')">PV</button></td>
      </tr>\`).join('');
  }).handleRequest({ action:'GET_UNG_VIEN_LIST', payload:{} });
}

// ────────────────────────────────────────────
// LOAD PHỎNG VẤN
// ────────────────────────────────────────────
function loadPhongVan() {
  const tbody = document.querySelector('#pv-table tbody');
  tbody.innerHTML = '<tr><td colspan="6" class="empty-state">Đang tải...</td></tr>';
  google.script.run.withSuccessHandler(res => {
    if(!res.success) return;
    _pvData = res.data;
    if(_pvData.length===0) { tbody.innerHTML='<tr><td colspan="6" class="empty-state">Chưa có lịch phỏng vấn</td></tr>'; return; }
    tbody.innerHTML = _pvData.map(r => \`
      <tr>
        <td class="name-cell">\${r.Ho_Ten_UV||r.ID_UngVien}</td>
        <td>\${fmtDate(r.Ngay_PV)}</td>
        <td style="color:var(--text-2)">\${r.Gio_PV||'—'}</td>
        <td>\${r.Nguoi_PV||'—'}</td>
        <td><span class="status s-dang-xet">\${r.Hinh_Thuc||'—'}</span></td>
        <td>\${statusBadge(r.Ket_Qua||'Chờ kết quả')}</td>
      </tr>\`).join('');
  }).handleRequest({ action:'GET_PHONG_VAN_LIST', payload:{} });
}

// ────────────────────────────────────────────
// LOAD AUDIT LOG
// ────────────────────────────────────────────
function loadAuditLog() {
  const tbody = document.querySelector('#audit-table tbody');
  if(!tbody) return;
  tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Đang tải...</td></tr>';
  google.script.run.withSuccessHandler(res => {
    if(!res.success) { toast(res.error,'error'); return; }
    const data = res.data;
    if(data.length===0) { tbody.innerHTML='<tr><td colspan="5" class="empty-state">Chưa có dữ liệu</td></tr>'; return; }
    tbody.innerHTML = data.map(r => \`
      <tr>
        <td style="color:var(--text-3);font-size:12px;white-space:nowrap">\${r['Thời gian']?new Date(r['Thời gian']).toLocaleString('vi-VN'):'—'}</td>
        <td style="font-size:12px">\${r['Email']||'—'}</td>
        <td><span class="audit-action a-\${r['Hành động']||'UPDATE'}">\${r['Hành động']||'—'}</span></td>
        <td style="font-size:12px;color:var(--text-2)">\${r['Sheet']||'—'}</td>
        <td style="font-size:12px;color:var(--text-3);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">\${r['Ghi chú']||r['Giá trị mới']||'—'}</td>
      </tr>\`).join('');
  }).handleRequest({ action:'GET_AUDIT_LOG', payload:{ limit:200 } });
}

// ────────────────────────────────────────────
// FORM SUBMIT
// ────────────────────────────────────────────
function submitNhanSu() {
  const data = {
    ho_ten:     document.getElementById('ns-hoten').value.trim(),
    bo_phan:    document.getElementById('ns-bophan').value.trim(),
    chuc_vu:    document.getElementById('ns-chucvu').value.trim(),
    ngay_vao:   document.getElementById('ns-ngayvao').value,
    email:      document.getElementById('ns-email').value.trim(),
    sdt:        document.getElementById('ns-sdt').value.trim(),
    ngay_sinh:  document.getElementById('ns-ngaysinh').value,
    trang_thai: document.getElementById('ns-trangthai').value,
    link_ho_so: document.getElementById('ns-link').value.trim(),
  };
  if(!data.ho_ten) { toast('Vui lòng nhập họ tên','error'); return; }
  google.script.run.withSuccessHandler(res => {
    if(res.success) { toast(res.message,'success'); closeModal('modal-ns'); loadNhanSu(); }
    else toast(res.error,'error');
  }).handleRequest({ action:'CREATE_NHAN_SU', payload:data });
}

function submitUngVien() {
  const data = {
    ho_ten:      document.getElementById('uv-hoten').value.trim(),
    vi_tri_tuyen:document.getElementById('uv-vitri').value.trim(),
    email:       document.getElementById('uv-email').value.trim(),
    sdt:         document.getElementById('uv-sdt').value.trim(),
    nguon_cv:    document.getElementById('uv-nguon').value,
    ngay_nop:    document.getElementById('uv-ngaynop').value,
    link_cv:     document.getElementById('uv-link').value.trim(),
    ghi_chu:     document.getElementById('uv-ghichu').value.trim(),
  };
  if(!data.ho_ten || !data.vi_tri_tuyen) { toast('Vui lòng nhập họ tên và vị trí','error'); return; }
  google.script.run.withSuccessHandler(res => {
    if(res.success) { toast(res.message,'success'); closeModal('modal-uv'); loadUngVien(); }
    else toast(res.error,'error');
  }).handleRequest({ action:'CREATE_UNG_VIEN', payload:data });
}

function submitPhongVan() {
  const data = {
    id_ung_vien: document.getElementById('pv-uvid').value.trim(),
    ngay_pv:     document.getElementById('pv-ngay').value,
    gio_pv:      document.getElementById('pv-gio').value,
    nguoi_pv:    document.getElementById('pv-nguoi').value.trim(),
    hinh_thuc:   document.getElementById('pv-hinhthuc').value,
    ghi_chu:     document.getElementById('pv-ghichu').value.trim(),
  };
  if(!data.id_ung_vien) { toast('Vui lòng nhập ID ứng viên','error'); return; }
  google.script.run.withSuccessHandler(res => {
    if(res.success) { toast(res.message,'success'); closeModal('modal-pv'); loadPhongVan(); }
    else toast(res.error,'error');
  }).handleRequest({ action:'CREATE_PHONG_VAN', payload:data });
}

function changeUVStatus(id, status) {
  google.script.run.withSuccessHandler(res => {
    if(res.success) toast('Đã cập nhật: ' + status);
    else toast(res.error,'error');
  }).handleRequest({ action:'UPDATE_UNG_VIEN_STATUS', payload:{ id, trang_thai:status } });
}

function scheduleInterview(id, name) {
  document.getElementById('pv-uvid').value = id;
  closeModal('modal-uv');
  openModal('modal-pv');
}

function softDeleteNS(id) {
  if(!confirm('Xác nhận chuyển nhân sự này sang trạng thái "Đã nghỉ"?')) return;
  google.script.run.withSuccessHandler(res => {
    if(res.success) { toast(res.message); loadNhanSu(); }
    else toast(res.error,'error');
  }).handleRequest({ action:'DELETE_NHAN_SU', payload:{ id } });
}

// ── INIT ──
loadDashboard();
</script>
</body>
</html>`;
}
