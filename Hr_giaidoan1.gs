// =============================================================================
// HRIS GIAI ĐOẠN 1 — NỀN TẢNG & BẢO MẬT (ĐÃ UPDATE ĐẦY ĐỦ CHO GĐ4 + GĐ5)
// Google Apps Script — Dán toàn bộ file này vào Apps Script Editor
// Cập nhật: 2026 | Tác giả: HR Analyst System
// =============================================================================
// CẤU TRÚC FILE:
//   MODULE 1 — Cấu hình & Khởi tạo (ĐÃ BỔ SUNG TOÀN BỘ GĐ4 + GĐ5)
//   MODULE 2 — Xác thực & Phân quyền
//   MODULE 3 — Audit Log
//   MODULE 4 — Backup tự động
//   MODULE 5 — Validation dữ liệu
//   MODULE 6 — Web App Entry Point
// =============================================================================


// =============================================================================
// MODULE 1 — CẤU HÌNH TRUNG TÂM (ĐÃ UPDATE ĐẦY ĐỦ CHO TẤT CẢ GIAI ĐOẠN)
// =============================================================================

const CONFIG = {
  // --- Tên các Sheet trong Spreadsheet ---
  SHEETS: {
    // Giai đoạn 1
    NHAN_SU:    'NHAN_SU',
    UNG_VIEN:   'UNG_VIEN',
    PHONG_VAN:  'PHONG_VAN',
    TUYEN_DUNG: 'TUYEN_DUNG',
    AUDIT_LOG:  'AUDIT_LOG',
    PHAN_QUYEN: 'PHAN_QUYEN',

    // Giai đoạn 4 — Onboarding & KPI
    ONBOARDING_TASKS:  'ONBOARDING_TASKS',
    KPI:               'KPI',
    PERFORMANCE_REVIEW: 'PERFORMANCE_REVIEW',

    // Giai đoạn 5 — Payroll + Leave + Attendance + Reports
    PAYROLL:         'PAYROLL',
    LEAVE_REQUESTS:  'LEAVE_REQUESTS',
    ATTENDANCE:      'ATTENDANCE',
    LEAVE_BALANCE:   'LEAVE_BALANCE',
  },

  // --- Phân quyền theo role ---
  ROLES: {
    ADMIN:    ['hr.khanhdo@gmail.com'],       // Toàn quyền
    HR:       ['hr.trungkhanh@gmail.com'],    // Quản lý nhân sự, tuyển dụng, lương
    MANAGER:  ['hrm.khanhdo@gmail.com'],      // Xem báo cáo bộ phận
    VIEWER:   [],                             // Chỉ xem
  },

  // --- Cài đặt Backup ---
  BACKUP: {
    FOLDER_NAME: 'HR_Backup',
    RETENTION_DAYS: 30,
    TIME_HOUR: 23,
  },

  // --- Cài đặt Notification ---
  NOTIFY: {
    EMAIL_ALERT: 'hr.khanhdo@gmail.com',
    SLACK_WEBHOOK: '',
  },

  // --- Cấu hình Payroll (Giai đoạn 5) ---
  PAYROLL: {
    ANNUAL_LEAVE_DAYS: 12,           // Số ngày phép năm mặc định
    WORKING_DAYS_MONTH: 26,          // Ngày công chuẩn trong tháng
    BHXH_EMPLOYEE_RATE: 0.08,        // NLĐ đóng BHXH 8%
    BHXH_EMPLOYER_RATE: 0.175,       // NSDLĐ đóng BHXH 17.5%
    BHYT_EMPLOYEE_RATE: 0.015,       // BHYT NLĐ 1.5%
    BHTN_EMPLOYEE_RATE: 0.01,        // BHTN NLĐ 1%
    PERSONAL_DEDUCTION: 11000000,    // Giảm trừ bản thân 11 triệu
    DEPENDENT_DEDUCTION: 4400000,    // Giảm trừ người phụ thuộc 4.4 triệu/người
  },

  // --- Màu sắc cho status trong sheet ---
  COLORS: {
    MOI_NOP:    '#E6F1FB',
    DANG_XET:   '#FAEEDA',
    HEN_PV:     '#EEEDFE',
    THU_VIEC:   '#E1F5EE',
    CHINH_THUC: '#EAF3DE',
    TU_CHOI:    '#FCEBEB',
    NGHI:       '#F1EFE8',
  },
};


// =============================================================================
// MODULE 2 — XÁC THỰC & PHÂN QUYỀN
// =============================================================================

/**
 * Lấy thông tin user hiện tại đang truy cập Web App
 * @returns {Object} { email, role, name }
 */
function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { email: null, role: null, name: 'Khách' };

    const role = getUserRole(email);
    const name = email.split('@')[0];
    return { email, role, name };
  } catch (e) {
    Logger.log('getCurrentUser error: ' + e.message);
    return { email: null, role: null, name: 'Khách' };
  }
}

/**
 * Xác định role của user dựa trên email
 * Ưu tiên kiểm tra sheet PHAN_QUYEN trước, sau đó fallback sang CONFIG.ROLES
 * @param {string} email
 * @returns {string|null} 'ADMIN' | 'HR' | 'MANAGER' | 'VIEWER' | null
 */
function getUserRole(email) {
  if (!email) return null;

  // 1. Kiểm tra trong sheet PHAN_QUYEN (linh hoạt hơn, không cần sửa code)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PHAN_QUYEN);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      // Cột A = email, Cột B = role, Cột C = active (TRUE/FALSE)
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString().toLowerCase() === email.toLowerCase()
            && data[i][2] === true) {
          return data[i][1].toString().toUpperCase();
        }
      }
    }
  } catch (e) {}

  // 2. Fallback: kiểm tra CONFIG.ROLES
  for (const [role, emails] of Object.entries(CONFIG.ROLES)) {
    if (emails.map(e => e.toLowerCase()).includes(email.toLowerCase())) {
      return role;
    }
  }

  return null; // Không có quyền truy cập
}

/**
 * Kiểm tra xem user có quyền thực hiện action không
 * @param {string} role - role của user
 * @param {string} action - 'READ' | 'WRITE' | 'DELETE' | 'ADMIN'
 * @returns {boolean}
 */
function hasPermission(role, action) {
  const PERMISSION_MATRIX = {
    ADMIN:   ['READ', 'WRITE', 'DELETE', 'ADMIN'],
    HR:      ['READ', 'WRITE'],
    MANAGER: ['READ'],
    VIEWER:  ['READ'],
  };

  if (!role || !PERMISSION_MATRIX[role]) return false;
  return PERMISSION_MATRIX[role].includes(action);
}

/**
 * Tạo sheet PHAN_QUYEN với header mẫu nếu chưa tồn tại
 */
function setupPhanQuyenSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PHAN_QUYEN);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PHAN_QUYEN);
    const headers = ['Email', 'Role', 'Active', 'Ho_Ten', 'Bo_Phan', 'Ghi_Chu', 'Ngay_Cap'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#534AB7')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    // Thêm dữ liệu mẫu
    const sampleData = [
      ['admin@company.com', 'ADMIN', true, 'Quản trị viên', 'IT', 'Tài khoản admin', new Date()],
      ['hr@company.com',    'HR',    true, 'HR Manager',   'HR', '',                 new Date()],
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    sheet.autoResizeColumns(1, headers.length);

    Logger.log('✅ Đã tạo sheet PHAN_QUYEN');
  }
  return sheet;
}


// =============================================================================
// MODULE 3 — AUDIT LOG
// =============================================================================

/**
 * Ghi một bản ghi vào AUDIT_LOG
 * @param {string} action   - 'CREATE' | 'UPDATE' | 'DELETE' | 'LOGIN' | 'ERROR'
 * @param {string} sheetName - Sheet bị tác động
 * @param {string} rowId    - ID của bản ghi bị thay đổi
 * @param {string} fieldChanged - Tên trường bị thay đổi (nếu có)
 * @param {*} oldValue      - Giá trị cũ
 * @param {*} newValue      - Giá trị mới
 * @param {string} note     - Ghi chú thêm
 */
function writeAuditLog(action, sheetName, rowId, fieldChanged, oldValue, newValue, note) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG);
    if (!logSheet) return;

    const user = getCurrentUser();
    const timestamp = new Date();

    const logRow = [
      timestamp,                                    // A: Thời gian
      user.email || 'system',                       // B: Email người thực hiện
      user.role || 'SYSTEM',                        // C: Role
      action,                                       // D: Hành động
      sheetName,                                    // E: Sheet bị tác động
      rowId || '',                                  // F: ID bản ghi
      fieldChanged || '',                           // G: Trường bị thay đổi
      oldValue !== undefined ? String(oldValue) : '',  // H: Giá trị cũ
      newValue !== undefined ? String(newValue) : '',  // I: Giá trị mới
      note || '',                                   // J: Ghi chú
    ];

    logSheet.appendRow(logRow);

    // Tô màu theo loại action
    const lastRow = logSheet.getLastRow();
    const actionColors = {
      CREATE: '#EAF3DE',
      UPDATE: '#FAEEDA',
      DELETE: '#FCEBEB',
      LOGIN:  '#E6F1FB',
      ERROR:  '#FAECE7',
    };
    const bgColor = actionColors[action] || '#F1EFE8';
    logSheet.getRange(lastRow, 1, 1, logRow.length).setBackground(bgColor);

  } catch (e) {
    Logger.log('⚠️ writeAuditLog error: ' + e.message);
  }
}

/**
 * Tạo sheet AUDIT_LOG với header chuẩn nếu chưa tồn tại
 */
function setupAuditLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.AUDIT_LOG);

    const headers = [
      'Thời gian', 'Email', 'Role', 'Hành động',
      'Sheet', 'ID bản ghi', 'Trường thay đổi',
      'Giá trị cũ', 'Giá trị mới', 'Ghi chú'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2C2C2A')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    // Freeze header row và cột thời gian
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);

    // Định dạng cột thời gian
    sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');

    // Set độ rộng cột
    const colWidths = [150, 200, 80, 80, 120, 100, 150, 200, 200, 200];
    colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    // Bảo vệ sheet AUDIT_LOG — chỉ xem, không cho sửa thủ công
    const protection = sheet.protect();
    protection.setDescription('Audit Log — chỉ được ghi bởi hệ thống');
    protection.setWarningOnly(true); // Cảnh báo khi cố sửa

    Logger.log('✅ Đã tạo sheet AUDIT_LOG');
  }
  return sheet;
}

/**
 * Trigger: Tự động ghi audit log khi có thay đổi trong bất kỳ sheet nào
 * Cài đặt: Triggers > onEdit > Spreadsheet > On edit
 */
function onEditAuditTrigger(e) {
  // Bỏ qua nếu đang sửa trong AUDIT_LOG hoặc PHAN_QUYEN
  const ignoredSheets = [CONFIG.SHEETS.AUDIT_LOG, CONFIG.SHEETS.PHAN_QUYEN];
  const sheetName = e.range.getSheet().getName();
  if (ignoredSheets.includes(sheetName)) return;

  // Bỏ qua nếu chỉ sửa 1 ô đơn và không có giá trị cũ (tránh log quá nhiều)
  const oldValue = e.oldValue;
  const newValue = e.value;
  if (oldValue === newValue) return;

  // Lấy header của cột bị sửa
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  let fieldName = '';
  try {
    fieldName = sheet.getRange(1, col).getValue();
  } catch (_) {}

  // Cố gắng lấy ID của bản ghi (giả sử cột A là ID)
  let rowId = '';
  try {
    rowId = sheet.getRange(e.range.getRow(), 1).getValue();
  } catch (_) {}

  writeAuditLog(
    'UPDATE',
    sheetName,
    rowId,
    fieldName,
    oldValue,
    newValue,
    `Sửa tại ô ${e.range.getA1Notation()}`
  );
}


// =============================================================================
// MODULE 4 — BACKUP TỰ ĐỘNG
// =============================================================================

/**
 * Tạo backup toàn bộ Spreadsheet vào Google Drive
 * Đặt Trigger: Time-driven > Day timer > 11pm to midnight
 */
function runDailyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileName = ss.getName();
  const today = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  const backupName = `[BACKUP] ${fileName} — ${today}`;

  try {
    // 1. Tìm hoặc tạo thư mục backup
    const folder = getOrCreateBackupFolder();

    // 2. Copy file hiện tại vào thư mục
    const ssFile = DriveApp.getFileById(ss.getId());
    const backupFile = ssFile.makeCopy(backupName, folder);

    // 3. Xóa các backup cũ hơn RETENTION_DAYS ngày
    cleanOldBackups(folder);

    // 4. Ghi log thành công
    writeAuditLog('CREATE', 'SYSTEM', 'BACKUP', 'DailyBackup', '', backupFile.getId(),
      `Backup thành công: ${backupName}`);

    Logger.log(`✅ Backup xong: ${backupName} (ID: ${backupFile.getId()})`);

    // 5. Gửi email xác nhận (tùy chọn)
    if (CONFIG.NOTIFY.EMAIL_ALERT) {
      sendBackupNotification(backupName, backupFile.getUrl(), true);
    }

    return { success: true, fileName: backupName, url: backupFile.getUrl() };

  } catch (e) {
    Logger.log('❌ Backup thất bại: ' + e.message);
    writeAuditLog('ERROR', 'SYSTEM', 'BACKUP', 'DailyBackup', '', '', 'Lỗi: ' + e.message);

    if (CONFIG.NOTIFY.EMAIL_ALERT) {
      sendBackupNotification(backupName, '', false, e.message);
    }
    return { success: false, error: e.message };
  }
}

/**
 * Lấy hoặc tạo thư mục backup trong Google Drive
 * @returns {Folder}
 */
function getOrCreateBackupFolder() {
  const folderName = CONFIG.BACKUP.FOLDER_NAME;
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();

  const folder = DriveApp.createFolder(folderName);
  Logger.log(`📁 Đã tạo thư mục backup: ${folderName}`);
  return folder;
}

/**
 * Xóa các file backup cũ hơn RETENTION_DAYS ngày
 * @param {Folder} folder
 */
function cleanOldBackups(folder) {
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - CONFIG.BACKUP.RETENTION_DAYS);

  const files = folder.getFiles();
  let deletedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoffDate) {
      file.setTrashed(true);
      deletedCount++;
      Logger.log(`🗑️ Đã xóa backup cũ: ${file.getName()}`);
    }
  }

  if (deletedCount > 0) {
    Logger.log(`✅ Đã dọn ${deletedCount} file backup cũ`);
  }
}

/**
 * Gửi email thông báo kết quả backup
 */
function sendBackupNotification(fileName, fileUrl, success, errorMsg) {
  const recipient = CONFIG.NOTIFY.EMAIL_ALERT;
  const today = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');

  const subject = success
    ? `✅ [HR System] Backup thành công — ${today}`
    : `❌ [HR System] Backup THẤT BẠI — ${today}`;

  const body = success
    ? `Backup hệ thống HR đã hoàn thành lúc ${today}.\n\nFile: ${fileName}\nLink: ${fileUrl}\n\nHệ thống giữ backup ${CONFIG.BACKUP.RETENTION_DAYS} ngày gần nhất.`
    : `Backup hệ thống HR THẤT BẠI lúc ${today}.\n\nLỗi: ${errorMsg}\n\nVui lòng kiểm tra lại hệ thống.`;

  try {
    MailApp.sendEmail(recipient, subject, body);
  } catch (e) {
    Logger.log('⚠️ Không gửi được email notification: ' + e.message);
  }
}


// =============================================================================
// MODULE 5 — VALIDATION DỮ LIỆU ĐẦU VÀO
// =============================================================================

/**
 * Validate một bản ghi nhân sự trước khi lưu
 * @param {Object} data - dữ liệu từ form
 * @returns {Object} { valid: boolean, errors: string[] }
 */
function validateNhanSu(data) {
  const errors = [];

  // Họ tên bắt buộc
  if (!data.ho_ten || data.ho_ten.trim().length < 2) {
    errors.push('Họ tên phải có ít nhất 2 ký tự');
  }

  // Email hợp lệ
  if (data.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) {
    errors.push('Email không hợp lệ');
  }

  // Số điện thoại (VN format)
  if (data.sdt && !/^(0|\+84)[3-9]\d{8}$/.test(data.sdt.replace(/\s/g, ''))) {
    errors.push('Số điện thoại không hợp lệ (phải là số VN, ví dụ: 0912345678)');
  }

  // Ngày sinh hợp lý (phải > 18 tuổi, < 65 tuổi)
  if (data.ngay_sinh) {
    const dob = new Date(data.ngay_sinh);
    const age = (new Date() - dob) / (365.25 * 24 * 3600 * 1000);
    if (age < 18) errors.push('Nhân sự phải từ 18 tuổi trở lên');
    if (age > 65) errors.push('Tuổi nhân sự không hợp lý (> 65)');
  }

  // Ngày vào làm không được trong tương lai
  if (data.ngay_vao) {
    const startDate = new Date(data.ngay_vao);
    if (startDate > new Date()) {
      errors.push('Ngày vào làm không thể là ngày trong tương lai');
    }
  }

  // Bộ phận bắt buộc
  if (!data.bo_phan || data.bo_phan.trim() === '') {
    errors.push('Bộ phận là thông tin bắt buộc');
  }

  return { valid: errors.length === 0, errors };
}

/**
 * Validate bản ghi ứng viên
 * @param {Object} data
 * @returns {Object} { valid: boolean, errors: string[] }
 */
function validateUngVien(data) {
  const errors = [];

  if (!data.ho_ten || data.ho_ten.trim().length < 2) {
    errors.push('Họ tên ứng viên là bắt buộc');
  }

  if (!data.vi_tri_tuyen || data.vi_tri_tuyen.trim() === '') {
    errors.push('Vị trí ứng tuyển là bắt buộc');
  }

  if (data.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) {
    errors.push('Email không hợp lệ');
  }

  // Kiểm tra trùng email trong sheet ỨNG VIÊN
  if (data.email) {
    const isDuplicate = checkDuplicateEmail(data.email, CONFIG.SHEETS.UNG_VIEN, data.id);
    if (isDuplicate) {
      errors.push(`Email ${data.email} đã tồn tại trong danh sách ứng viên`);
    }
  }

  return { valid: errors.length === 0, errors };
}

/**
 * Kiểm tra email trùng trong một sheet
 * @param {string} email
 * @param {string} sheetName
 * @param {string} excludeId - ID bản ghi hiện tại (khi update, bỏ qua bản ghi này)
 * @returns {boolean}
 */
function checkDuplicateEmail(email, sheetName, excludeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const emailCol = headers.indexOf('email');
    const idCol = 0; // Giả sử cột A là ID

    if (emailCol === -1) return false;

    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][idCol];
      const rowEmail = data[i][emailCol];
      if (rowEmail.toString().toLowerCase() === email.toLowerCase()
          && rowId !== excludeId) {
        return true;
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

/**
 * Tạo ID tự động theo format: PREFIX-YYYYMMDD-XXXX
 * @param {string} prefix - 'NS' cho nhân sự, 'UV' cho ứng viên, 'PV' cho phỏng vấn
 * @returns {string} Ví dụ: 'NS-20260427-0001'
 */
function generateId(prefix) {
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd');
  const random = Math.floor(Math.random() * 9000) + 1000;
  return `${prefix}-${dateStr}-${random}`;
}


// =============================================================================
// MODULE 6 — WEB APP ENTRY POINT
// =============================================================================

/**
 * Điểm vào của Web App — chạy khi user truy cập link
 * Deploy > New deployment > Web app > Execute as: Me > Who: Anyone with Google Account
 */
function doGet(e) {
  // 1. Xác thực user
  const user = getCurrentUser();

  // 2. Nếu không có quyền → trả về trang thông báo
  if (!user.email || !user.role) {
    return HtmlService.createHtmlOutput(buildAccessDeniedPage(user.email))
      .setTitle('HR System — Không có quyền truy cập')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 3. Ghi log đăng nhập
  writeAuditLog('LOGIN', 'SYSTEM', user.email, '', '', user.role,
    `Truy cập lúc ${new Date().toLocaleString('vi-VN')}`);

  // 4. Render giao diện chính (sẽ phát triển trong Giai đoạn 2)
  const template = HtmlService.createTemplateFromString(buildMainApp(user));
  return template.evaluate()
    .setTitle('HR Workspace — ' + user.name)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Xử lý các API call từ frontend (gọi qua google.script.run)
 * @param {Object} request - { action, payload }
 * @returns {Object} { success, data, error }
 */
function handleRequest(request) {
  const user = getCurrentUser();

  // Guard: phải đăng nhập
  if (!user.email || !user.role) {
    return { success: false, error: 'Không có quyền truy cập. Vui lòng đăng nhập bằng tài khoản công ty.' };
  }

  const { action, payload } = request;

  try {
    switch (action) {

      case 'GET_NHAN_SU_LIST':
        if (!hasPermission(user.role, 'READ')) throw new Error('Không có quyền xem danh sách nhân sự');
        return { success: true, data: getNhanSuList(user) };

      case 'CREATE_NHAN_SU':
        if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền thêm nhân sự');
        return createNhanSu(payload, user);

      case 'UPDATE_NHAN_SU':
        if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền sửa nhân sự');
        return updateNhanSu(payload, user);

      case 'DELETE_NHAN_SU':
        if (!hasPermission(user.role, 'DELETE')) throw new Error('Không có quyền xóa nhân sự');
        return deleteNhanSu(payload.id, user);

      case 'GET_AUDIT_LOG':
        if (!hasPermission(user.role, 'ADMIN')) throw new Error('Chỉ Admin mới xem được Audit Log');
        return { success: true, data: getAuditLog(payload) };

      case 'GET_DASHBOARD_STATS':
        if (!hasPermission(user.role, 'READ')) throw new Error('Không có quyền xem báo cáo');
        return { success: true, data: getDashboardStats(user) };

      default:
        return { success: false, error: `Action không hợp lệ: ${action}` };
    }
  } catch (e) {
    writeAuditLog('ERROR', 'SYSTEM', '', action, '', '', e.message);
    return { success: false, error: e.message };
  }
}


// =============================================================================
// MODULE 7 — DATA ACCESS LAYER (CRUD cơ bản)
// =============================================================================

/**
 * Lấy danh sách nhân sự
 * MANAGER chỉ thấy bộ phận của mình
 */
function getNhanSuList(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => h.toString());
  const rows = data.slice(1).filter(row => row[0] !== '');

  let result = rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  // MANAGER chỉ xem bộ phận của mình
  if (user.role === 'MANAGER') {
    const managerDept = getManagerDepartment(user.email);
    if (managerDept) {
      result = result.filter(r => r['Bo_Phan'] === managerDept);
    }
  }

  return result;
}

/**
 * Tạo mới nhân sự
 */
function createNhanSu(data, user) {
  // 1. Validate
  const validation = validateNhanSu(data);
  if (!validation.valid) {
    return { success: false, error: validation.errors.join('\n') };
  }

  // 2. Tạo ID
  const id = generateId('NS');

  // 3. Ghi vào sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
  if (!sheet) return { success: false, error: 'Không tìm thấy sheet NHAN_SU' };

  const newRow = [
    id,
    data.ho_ten,
    data.bo_phan || '',
    data.chuc_vu || '',
    data.ngay_vao ? new Date(data.ngay_vao) : '',
    data.ngay_sinh ? new Date(data.ngay_sinh) : '',
    data.sdt || '',
    data.email || '',
    data.trang_thai || 'Chính thức',
    data.link_ho_so || '',
    new Date(),       // created_at
    user.email,       // created_by
  ];
  sheet.appendRow(newRow);

  // 4. Tô màu theo trạng thái
  const lastRow = sheet.getLastRow();
  const statusColor = CONFIG.COLORS[data.trang_thai] || '#FFFFFF';
  sheet.getRange(lastRow, 1, 1, newRow.length).setBackground(statusColor);

  // 5. Ghi audit log
  writeAuditLog('CREATE', CONFIG.SHEETS.NHAN_SU, id, '', '', JSON.stringify(data),
    `Thêm nhân sự mới: ${data.ho_ten}`);

  return { success: true, id, message: `Đã thêm nhân sự ${data.ho_ten} thành công` };
}

/**
 * Cập nhật thông tin nhân sự
 */
function updateNhanSu(data, user) {
  if (!data.id) return { success: false, error: 'Thiếu ID nhân sự' };

  const validation = validateNhanSu(data);
  if (!validation.valid) {
    return { success: false, error: validation.errors.join('\n') };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
  if (!sheet) return { success: false, error: 'Không tìm thấy sheet NHAN_SU' };

  const allData = sheet.getDataRange().getValues();
  const idCol = 0;

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][idCol] === data.id) {
      // Ghi nhận giá trị cũ để audit
      const oldRow = allData[i];

      // Cập nhật từng trường
      const updates = {
        1: data.ho_ten,
        2: data.bo_phan,
        3: data.chuc_vu,
        4: data.ngay_vao ? new Date(data.ngay_vao) : allData[i][4],
        5: data.ngay_sinh ? new Date(data.ngay_sinh) : allData[i][5],
        6: data.sdt,
        7: data.email,
        8: data.trang_thai,
        9: data.link_ho_so,
        11: user.email, // updated_by (cột L)
        12: new Date(), // updated_at (cột M) — thêm nếu có
      };

      Object.entries(updates).forEach(([col, val]) => {
        if (val !== undefined) sheet.getRange(i + 1, parseInt(col) + 1).setValue(val);
      });

      writeAuditLog('UPDATE', CONFIG.SHEETS.NHAN_SU, data.id, 'multiple',
        JSON.stringify(oldRow), JSON.stringify(data), `Cập nhật: ${data.ho_ten}`);

      return { success: true, message: `Đã cập nhật thông tin ${data.ho_ten}` };
    }
  }

  return { success: false, error: `Không tìm thấy nhân sự với ID: ${data.id}` };
}

/**
 * Xóa mềm nhân sự (đặt trạng thái = 'Đã nghỉ', không xóa hàng)
 */
function deleteNhanSu(id, user) {
  if (!id) return { success: false, error: 'Thiếu ID' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
  if (!sheet) return { success: false, error: 'Không tìm thấy sheet' };

  const allData = sheet.getDataRange().getValues();

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === id) {
      // Xóa mềm: set trạng thái = 'Đã nghỉ'
      sheet.getRange(i + 1, 9).setValue('Đã nghỉ');
      sheet.getRange(i + 1, 1, 1, allData[0].length).setBackground(CONFIG.COLORS.NGHI);

      writeAuditLog('DELETE', CONFIG.SHEETS.NHAN_SU, id, 'Trang_Thai',
        allData[i][8], 'Đã nghỉ', `Xóa mềm bởi ${user.email}`);

      return { success: true, message: 'Đã cập nhật trạng thái nhân sự thành "Đã nghỉ"' };
    }
  }

  return { success: false, error: 'Không tìm thấy nhân sự' };
}

/**
 * Lấy thống kê Dashboard
 */
function getDashboardStats(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stats = {};

  // Đếm nhân sự
  try {
    const nsSheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
    if (nsSheet) {
      const nsData = nsSheet.getDataRange().getValues().slice(1);
      stats.total_nhan_su = nsData.filter(r => r[0] && r[8] !== 'Đã nghỉ').length;
      stats.nghi_viec = nsData.filter(r => r[8] === 'Đã nghỉ').length;
    }
  } catch (e) {}

  // Đếm ứng viên
  try {
    const uvSheet = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN);
    if (uvSheet) {
      const uvData = uvSheet.getDataRange().getValues().slice(1);
      stats.total_ung_vien = uvData.filter(r => r[0]).length;
      stats.thu_viec = uvData.filter(r => r[5] === 'Thử việc').length;
    }
  } catch (e) {}

  // Sinh nhật tuần này
  try {
    stats.sinh_nhat_tuan_nay = getSinhNhatTuanNay();
  } catch (e) {
    stats.sinh_nhat_tuan_nay = [];
  }

  return stats;
}

/**
 * Lấy danh sách nhân sự có sinh nhật trong 7 ngày tới
 */
function getSinhNhatTuanNay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues().slice(1);
  const today = new Date();
  const result = [];

  data.forEach(row => {
    if (!row[5]) return; // ngay_sinh ở cột F (index 5)
    const dob = new Date(row[5]);
    const thisYearBirthday = new Date(today.getFullYear(), dob.getMonth(), dob.getDate());

    const diffDays = Math.ceil((thisYearBirthday - today) / (1000 * 3600 * 24));

    if (diffDays >= 0 && diffDays <= 7) {
      result.push({ name: row[1], date: thisYearBirthday, days_until: diffDays });
    }
  });

  return result.sort((a, b) => a.days_until - b.days_until);
}

/**
 * Lấy audit log (Admin only)
 */
function getAuditLog(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const rows = data.slice(1).reverse(); // Mới nhất trước

  // Giới hạn 200 bản ghi gần nhất
  const limit = (options && options.limit) || 200;
  return rows.slice(0, limit).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

/**
 * Helper: Lấy bộ phận của Manager
 */
function getManagerDepartment(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PHAN_QUYEN);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        return data[i][4] || null; // Cột E = Bo_Phan
      }
    }
  } catch (e) {}
  return null;
}


// =============================================================================
// MODULE 8 — SETUP & KHỞI TẠO HỆ THỐNG
// Chạy hàm này 1 lần duy nhất để cài đặt toàn bộ
// =============================================================================

/**
 * CHẠY HÀM NÀY ĐẦU TIÊN để khởi tạo toàn bộ hệ thống
 * Tạo tất cả sheet cần thiết, đặt trigger tự động
 */
function setupSystem() {
  Logger.log('🚀 Bắt đầu khởi tạo hệ thống HR...');

  // 1. Tạo các sheet cần thiết
  setupAuditLogSheet();
  setupPhanQuyenSheet();
  setupNhanSuSheet();
  setupUngVienSheet();

  // 2. Đặt trigger tự động
  setupTriggers();

  // 3. Ghi log
  writeAuditLog('CREATE', 'SYSTEM', 'SETUP', '', '', 'v1.0',
    'Khởi tạo hệ thống HR Giai đoạn 1');

  Logger.log('✅ Khởi tạo hệ thống hoàn tất!');
  SpreadsheetApp.getUi().alert(
    '✅ Hệ thống HR đã được khởi tạo thành công!\n\n' +
    'Đã tạo:\n• Sheet AUDIT_LOG\n• Sheet PHAN_QUYEN\n• Sheet NHAN_SU (nếu chưa có)\n• Sheet UNG_VIEN (nếu chưa có)\n\n' +
    'Trigger tự động:\n• Backup hàng ngày lúc 11pm\n• Nhắc sinh nhật 8am hàng ngày\n• Audit log mọi thay đổi\n\n' +
    'Bước tiếp theo: Thêm email người dùng vào sheet PHAN_QUYEN'
  );
}

/**
 * Tạo sheet NHAN_SU với header chuẩn
 */
function setupNhanSuSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.NHAN_SU);
    const headers = [
      'ID_NhanSu', 'Ho_Ten', 'Bo_Phan', 'Chuc_Vu',
      'Ngay_Vao', 'Ngay_Sinh', 'SDT', 'Email',
      'Trang_Thai', 'Link_HoSo', 'Created_At', 'Created_By'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#085041')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Data validation cho Trang_Thai
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Chính thức', 'Thử việc', 'Thực tập', 'Đã nghỉ'], true)
      .build();
    sheet.getRange('I2:I1000').setDataValidation(statusRule);

    Logger.log('✅ Đã tạo sheet NHAN_SU');
  }
}

/**
 * Tạo sheet UNG_VIEN với header chuẩn
 */
function setupUngVienSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.UNG_VIEN);
    const headers = [
      'ID_UngVien', 'Ho_Ten', 'Vi_Tri_Tuyen', 'Nguon_CV',
      'Ngay_Nop', 'Trang_Thai', 'Email', 'SDT',
      'Link_CV', 'Ghi_Chu', 'Created_At', 'Created_By'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#0C447C')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Data validation Trang_Thai ứng viên
    const uvStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Mới nộp', 'Đang xét', 'Hẹn PV', 'Thử việc', 'Chính thức', 'Từ chối'], true)
      .build();
    sheet.getRange('F2:F1000').setDataValidation(uvStatusRule);

    Logger.log('✅ Đã tạo sheet UNG_VIEN');
  }
}

/**
 * Cài đặt tất cả trigger tự động
 * Xóa trigger cũ trước khi tạo mới để tránh duplicate
 */
function setupTriggers() {
  // Xóa tất cả trigger cũ của project này
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 1. Trigger Audit Log — chạy khi có edit
  ScriptApp.newTrigger('onEditAuditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  // 2. Trigger Backup hàng ngày lúc 11pm
  ScriptApp.newTrigger('runDailyBackup')
    .timeBased()
    .atHour(CONFIG.BACKUP.TIME_HOUR)
    .everyDays(1)
    .create();

  // 3. Trigger nhắc sinh nhật lúc 8am
  ScriptApp.newTrigger('runBirthdayReminder')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  Logger.log('✅ Đã cài đặt 3 triggers: onEdit, Daily Backup 11pm, Birthday Reminder 8am');
}

/**
 * Gửi nhắc sinh nhật hàng ngày lúc 8am
 */
function runBirthdayReminder() {
  const birthdays = getSinhNhatTuanNay().filter(b => b.days_until === 0);

  if (birthdays.length === 0) return;

  const names = birthdays.map(b => b.name).join(', ');
  const subject = `🎂 [HR] Hôm nay có ${birthdays.length} nhân sự sinh nhật!`;
  const body = `Xin chào,\n\nHôm nay là sinh nhật của:\n${birthdays.map(b => `• ${b.name}`).join('\n')}\n\nĐừng quên gửi lời chúc nhé! 🎉`;

  try {
    MailApp.sendEmail(CONFIG.NOTIFY.EMAIL_ALERT, subject, body);
    writeAuditLog('CREATE', 'SYSTEM', 'BIRTHDAY', '', '', names, 'Đã gửi nhắc sinh nhật');
  } catch (e) {
    Logger.log('⚠️ Không gửi được nhắc sinh nhật: ' + e.message);
  }
}


// =============================================================================
// MODULE 9 — HTML TEMPLATES (tạm thời cho Giai đoạn 1)
// Giai đoạn 2 sẽ rebuild toàn bộ UI
// =============================================================================

function buildMainApp(user) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>HR Workspace</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Google Sans', sans-serif; }
    body { background: #f5f5f5; color: #333; }
    .topbar { background: #534AB7; color: white; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; }
    .topbar h1 { font-size: 16px; font-weight: 500; }
    .user-info { font-size: 13px; opacity: 0.85; }
    .badge { display: inline-block; background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 12px; font-size: 11px; margin-left: 8px; }
    .content { max-width: 900px; margin: 32px auto; padding: 0 16px; }
    .card { background: white; border-radius: 8px; padding: 20px; margin-bottom: 16px; border: 0.5px solid #e0e0e0; }
    .card h2 { font-size: 14px; color: #666; font-weight: 500; margin-bottom: 12px; }
    .stat-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
    .stat { background: #f8f8f8; border-radius: 6px; padding: 12px; text-align: center; }
    .stat .number { font-size: 28px; font-weight: 500; color: #534AB7; }
    .stat .label { font-size: 12px; color: #888; margin-top: 4px; }
    .status-ok { color: #3B6D11; background: #EAF3DE; padding: 3px 8px; border-radius: 4px; font-size: 12px; }
  </style>
</head>
<body>
  <div class="topbar">
    <h1>HR Workspace</h1>
    <div class="user-info">
      ${user.name} <span class="badge">${user.role}</span>
    </div>
  </div>
  <div class="content">
    <div class="card">
      <h2>Trạng thái hệ thống</h2>
      <p style="font-size:13px;color:#555;">Giai đoạn 1 đã hoàn thành ✅ — Bảo mật & Nền tảng đã được thiết lập.</p>
      <ul style="margin-top:10px;font-size:13px;color:#555;padding-left:18px;line-height:2;">
        <li><span class="status-ok">Active</span> &nbsp;Audit Log — ghi lại mọi thay đổi</li>
        <li><span class="status-ok">Active</span> &nbsp;Backup tự động lúc 11pm hàng ngày</li>
        <li><span class="status-ok">Active</span> &nbsp;Phân quyền theo Google Account</li>
        <li><span class="status-ok">Active</span> &nbsp;Validation dữ liệu đầu vào</li>
      </ul>
    </div>
    <div class="card">
      <h2>Xem nhanh</h2>
      <div class="stat-grid" id="stats">
        <div class="stat"><div class="number" id="ns">...</div><div class="label">Nhân sự hiện tại</div></div>
        <div class="stat"><div class="number" id="uv">...</div><div class="label">Ứng viên</div></div>
        <div class="stat"><div class="number" id="bday">...</div><div class="label">Sinh nhật tuần này</div></div>
      </div>
    </div>
  </div>
  <script>
    google.script.run.withSuccessHandler(function(res) {
      if(res.success) {
        document.getElementById('ns').textContent = res.data.total_nhan_su || 0;
        document.getElementById('uv').textContent = res.data.total_ung_vien || 0;
        document.getElementById('bday').textContent = (res.data.sinh_nhat_tuan_nay||[]).length;
      }
    }).handleRequest({ action: 'GET_DASHBOARD_STATS', payload: {} });
  </script>
</body>
</html>`;
}

function buildAccessDeniedPage(email) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;background:#f5f5f5;}
  .box{background:white;padding:40px;border-radius:12px;text-align:center;max-width:400px;}
  h2{color:#A32D2D;margin-bottom:12px;}p{color:#666;font-size:14px;line-height:1.6;}</style>
  </head><body><div class="box">
  <h2>Không có quyền truy cập</h2>
  <p>Tài khoản <strong>${email || 'của bạn'}</strong> chưa được cấp quyền truy cập hệ thống HR.</p>
  <p style="margin-top:12px;">Vui lòng liên hệ Admin để được cấp quyền.</p>
  </div></body></html>`;
}
