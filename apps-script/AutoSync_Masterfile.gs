/**
 * AUTO SYNC cho SC_Masterfile_Onboarding_Training
 * ------------------------------------------------------------
 * 1) syncUpcomingBD : chụp (snapshot) DS BD upcoming từ
 *    [SPX BD]- Headcount Tracker › "[BD] Headcount_upcoming" cột AN:AZ
 *    → nối thêm vào sheet "Danh sách OB (draw)" (A:M). Chỉ THÊM người mới
 *    (so theo Staff ID / Email), không xoá ai → qua ngày HC Tracker mất DS
 *    thì Masterfile vẫn giữ.
 * 1b) updateOnboardStatus: cột S "Trang_Thai_OB" tự chuyển
 *    Chờ onboard → Đã onboard   (đã điểm danh "Có", hoặc HC Tracker báo Active và đã tới ngày OB)
 *               → Không onboard (nằm trong sheet Inactive / Status = Inactive)
 *               → Cần kiểm tra  (quá ngày OB + GRACE_DAYS mà chưa có tín hiệu nào)
 *    Sửa tay thành "Đã onboard"/"Không onboard" thì script không ghi đè nữa.
 * 2) syncQuizResults: lấy Data_Topic1 / Data_Topic2 từ file
 *    "Kết quả bài quiz sau training" → nối các lượt làm mới vào sheet cùng tên
 *    trong Masterfile (khoá = Email + Attempt_No).
 *
 * CÀI ĐẶT (làm 1 lần):
 *   Masterfile › Tiện ích mở rộng › Apps Script › dán file này › Lưu
 *   › chọn hàm setupTriggers › Run › cấp quyền.
 *   Sau đó script tự chạy mỗi giờ. Có menu "🔄 Auto Sync" để chạy tay.
 *
 * Lưu ý: script chạy bằng quyền của người bấm setupTriggers, người đó chỉ cần
 * quyền XEM file Headcount Tracker (không cần quyền sửa).
 */

const CFG = {
  HC_ID: '122JM8NSS7FdBkz5n10I2zYtRz97-HPYpDiknw5YHRW4',
  HC_SHEET: '[BD] Headcount_upcoming',
  HC_FIRST_COL: 'AN',          // AN..AZ = 13 cột → A..M của "Danh sách OB (draw)"
  HC_NUM_COLS: 13,

  QUIZ_ID: '1LrzFY81FirN7hC3-yQwbnTt7pMyc0NDl8InQN1H4nIU',
  QUIZ_SHEETS: [               // [sheet nguồn, sheet đích, số cột cần copy]
    ['Data_Topic1', 'Data_Topic1', 7],   // Email..Attempt_No
    ['Data_Topic2', 'Data_Topic2', 25],  // Email..N6_TongThe
  ],

  OB_SHEET: 'Danh sách OB (draw)',
  OB_NOTE_COL: 18,             // R = Ghi_chu
  OB_STATUS_COL: 19,           // S = Trang_Thai_OB (cột mới)

  // Nguồn xác nhận BD đã thực sự onboard
  HC_ACTIVE_SHEET: 'Headcount List - Update',   // C=Staff ID, D=Mail, P=Status
  HC_INACTIVE_SHEET: 'Inactive',                // B=Staff ID, C=Mail
  ATTEND_SHEET: 'Diem_Danh',                    // A=Email, F=Chang1_Buoi1, G=Chang1_Buoi2
  GRACE_DAYS: 3,               // quá ngày OB bấy nhiêu ngày mà chưa thấy → "Cần kiểm tra"
  LOG_SHEET: 'Sync_Log',
  TZ: 'Asia/Ho_Chi_Minh',
};

/* ====================== MENU & TRIGGER ====================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Auto Sync')
    .addItem('Đồng bộ ngay (DS onboard + Quiz)', 'syncAll')
    .addItem('Chỉ lấy DS BD upcoming', 'syncUpcomingBD')
    .addItem('Chỉ cập nhật trạng thái onboard', 'updateOnboardStatus')
    .addItem('Chỉ lấy kết quả Quiz', 'syncQuizResults')
    .addSeparator()
    .addItem('Bật chạy tự động mỗi giờ', 'setupTriggers')
    .addToUi();
}

function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncAll')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  syncAll();
}

function syncAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;
  try {
    syncUpcomingBD();
    updateOnboardStatus();
    syncQuizResults();
  } finally {
    lock.releaseLock();
  }
}

/* ====================== 1. DS BD UPCOMING ====================== */

function syncUpcomingBD() {
  const master = SpreadsheetApp.getActive();
  const ob = master.getSheetByName(CFG.OB_SHEET);
  if (!ob) throw new Error('Không thấy sheet ' + CFG.OB_SHEET);

  // --- Đọc nguồn ---
  const hc = SpreadsheetApp.openById(CFG.HC_ID).getSheetByName(CFG.HC_SHEET);
  const hcLast = hc.getLastRow();
  if (hcLast < 2) return log_('BD upcoming', 0, 'Nguồn trống');
  const src = hc.getRange(CFG.HC_FIRST_COL + '2:' + colLetter_(colIndex_(CFG.HC_FIRST_COL) + CFG.HC_NUM_COLS - 1) + hcLast)
    .getValues()
    .filter(r => isValid_(r[0]) && isValid_(r[1]));   // có Full Name + Staff ID

  // --- Đọc đích: khoá theo Staff ID & Email ---
  const obLast = lastDataRow_(ob, 2);                  // dò theo cột B (Staff_ID)
  const existing = obLast >= 2 ? ob.getRange(2, 1, obLast - 1, 6).getValues() : [];
  const ids = new Set(), mails = new Set();
  const rowById = {};
  existing.forEach((r, i) => {
    const id = norm_(r[1]), mail = norm_(r[2]);
    if (id) { ids.add(id); rowById[id] = i + 2; }
    if (mail) mails.add(mail);
  });

  const today = Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy');
  const toAdd = [];
  let updated = 0;

  src.forEach(r => {
    const id = norm_(r[1]), mail = norm_(r[2]);
    if (ids.has(id) || (mail && mails.has(mail))) {
      // Đã có → chỉ cập nhật Start_Date nếu HR đổi ngày onboard
      const row = rowById[id];
      if (row && r[5] instanceof Date) {
        const cell = ob.getRange(row, 6);
        const cur = cell.getValue();
        if (!(cur instanceof Date) || cur.getTime() !== r[5].getTime()) {
          cell.setValue(r[5]);
          updated++;
        }
      }
      return;
    }
    ids.add(id); if (mail) mails.add(mail);
    toAdd.push(r);
  });

  if (toAdd.length) {
    const start = obLast + 1;
    ob.getRange(start, 1, toAdd.length, CFG.HC_NUM_COLS).setValues(toAdd);
    ob.getRange(start, CFG.OB_NOTE_COL, toAdd.length, 1)
      .setValues(toAdd.map(() => ['Auto từ HC Tracker ' + today]));
    ob.getRange(start, CFG.OB_STATUS_COL, toAdd.length, 1)
      .setValues(toAdd.map(() => [ST.WAIT]));
  }
  log_('BD upcoming', toAdd.length, 'Thêm mới ' + toAdd.length + ', đổi ngày OB ' + updated);
}

/* ====================== 1b. XÁC NHẬN ĐÃ ONBOARD ====================== */

const ST = {
  WAIT: 'Chờ onboard',
  DONE: 'Đã onboard',
  NO: 'Không onboard',
  CHECK: 'Cần kiểm tra',
};
// Chỉ các trạng thái này mới được script tự đổi. Bạn sửa tay thành
// "Đã onboard" / "Không onboard" thì script sẽ KHÔNG đụng vào nữa.
const AUTO_ST = new Set([ST.WAIT, ST.CHECK]);

function updateOnboardStatus() {
  const master = SpreadsheetApp.getActive();
  const ob = master.getSheetByName(CFG.OB_SHEET);
  ensureStatusColumn_(ob);

  const obLast = lastDataRow_(ob, 2);
  if (obLast < 2) return;
  const rows = ob.getRange(2, 1, obLast - 1, CFG.OB_STATUS_COL).getValues();

  // --- Tín hiệu từ HC Tracker ---
  const hc = SpreadsheetApp.openById(CFG.HC_ID);
  const active = new Set(), inactive = new Set();
  const act = hc.getSheetByName(CFG.HC_ACTIVE_SHEET);
  if (act && act.getLastRow() > 1) {
    act.getRange(2, 3, act.getLastRow() - 1, 14).getValues()      // C..P
      .forEach(r => {
        const st = norm_(r[13]);                                    // P = Status
        const set = st === 'active' ? active : st === 'inactive' ? inactive : null;
        if (set) { if (norm_(r[0])) set.add(norm_(r[0])); if (norm_(r[1])) set.add(norm_(r[1])); }
      });
  }
  const ina = hc.getSheetByName(CFG.HC_INACTIVE_SHEET);
  if (ina && ina.getLastRow() > 1) {
    ina.getRange(2, 2, ina.getLastRow() - 1, 2).getValues()       // B..C
      .forEach(r => { if (norm_(r[0])) inactive.add(norm_(r[0])); if (norm_(r[1])) inactive.add(norm_(r[1])); });
  }

  // --- Tín hiệu từ điểm danh trong Masterfile ---
  const attended = new Set();
  const dd = master.getSheetByName(CFG.ATTEND_SHEET);
  if (dd && dd.getLastRow() > 1) {
    dd.getRange(2, 1, dd.getLastRow() - 1, 7).getValues()
      .forEach(r => { if (norm_(r[5]) === 'có' || norm_(r[6]) === 'có') attended.add(norm_(r[0])); });
  }

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const DAY = 24 * 3600 * 1000;
  const changes = { [ST.DONE]: 0, [ST.NO]: 0, [ST.CHECK]: 0 };

  const out = rows.map(r => {
    const cur = String(r[CFG.OB_STATUS_COL - 1]).trim();
    if (!AUTO_ST.has(cur)) return [cur];                 // trống hoặc đã chốt tay → giữ nguyên
    const id = norm_(r[1]), mail = norm_(r[2]);
    const start = r[5] instanceof Date ? r[5] : null;
    const has = set => (id && set.has(id)) || (mail && set.has(mail));

    let next = ST.WAIT;
    if (attended.has(mail) || (has(active) && start && start <= today)) next = ST.DONE;
    else if (has(inactive)) next = ST.NO;
    else if (start && today - start > CFG.GRACE_DAYS * DAY) next = ST.CHECK;

    if (next !== cur && changes[next] !== undefined) changes[next]++;
    return [next];
  });

  ob.getRange(2, CFG.OB_STATUS_COL, out.length, 1).setValues(out);
  log_('Trạng thái OB', changes[ST.DONE] + changes[ST.NO] + changes[ST.CHECK],
    'Đã onboard +' + changes[ST.DONE] + ', Không onboard +' + changes[ST.NO] + ', Cần kiểm tra +' + changes[ST.CHECK]);
}

function ensureStatusColumn_(ob) {
  const head = ob.getRange(1, CFG.OB_STATUS_COL);
  if (!String(head.getValue()).trim()) head.setValue('Trang_Thai_OB');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([ST.WAIT, ST.DONE, ST.NO, ST.CHECK], true)
    .setAllowInvalid(true).build();
  ob.getRange(2, CFG.OB_STATUS_COL, Math.max(ob.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

/* ====================== 2. KẾT QUẢ QUIZ ====================== */

function syncQuizResults() {
  const master = SpreadsheetApp.getActive();
  const quiz = SpreadsheetApp.openById(CFG.QUIZ_ID);

  CFG.QUIZ_SHEETS.forEach(([srcName, dstName, nCols]) => {
    const s = quiz.getSheetByName(srcName);
    const d = master.getSheetByName(dstName);
    if (!s || !d) return log_(dstName, 0, 'Thiếu sheet');

    // Nếu sheet đích đang chạy bằng công thức thì không ghi đè
    if (d.getRange('A2').getFormula()) return log_(dstName, 0, 'A2 là công thức → bỏ qua');

    const sLast = s.getLastRow();
    if (sLast < 2) return log_(dstName, 0, 'Nguồn trống');
    const src = s.getRange(2, 1, sLast - 1, nCols).getValues().filter(r => isValid_(r[0]));

    const dLast = lastDataRow_(d, 1);
    const keys = new Set(
      (dLast >= 2 ? d.getRange(2, 1, dLast - 1, 7).getValues() : [])
        .map(r => norm_(r[0]) + '|' + r[6])            // Email | Attempt_No
    );

    const toAdd = src.filter(r => {
      const k = norm_(r[0]) + '|' + r[6];
      if (keys.has(k)) return false;
      keys.add(k);
      return true;
    });

    if (toAdd.length) d.getRange(dLast + 1, 1, toAdd.length, nCols).setValues(toAdd);
    log_(dstName, toAdd.length, 'Thêm ' + toAdd.length + ' lượt làm mới');
  });
}

/* ====================== HELPERS ====================== */

function isValid_(v) {
  const s = String(v == null ? '' : v).trim();
  return s !== '' && s !== '-' && s.charAt(0) !== '#';   // loại trống, "-", #N/A, #REF!
}

function norm_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }

function lastDataRow_(sheet, col) {
  const n = sheet.getLastRow();
  if (n < 2) return 1;
  const vals = sheet.getRange(1, col, n, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim() !== '') return i + 1;
  return 1;
}

function colIndex_(letters) {
  return letters.toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function log_(what, count, msg) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.LOG_SHEET) || ss.insertSheet(CFG.LOG_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(['Thời gian', 'Hạng mục', 'Số dòng', 'Ghi chú']);
  sh.appendRow([new Date(), what, count, msg]);
}
