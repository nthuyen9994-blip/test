/**
 * AUTO SYNC cho SC_Masterfile_Onboarding_Training
 * ------------------------------------------------------------
 * 1) syncUpcomingBD : lấy DS BD upcoming từ
 *    [SPX BD]- Headcount Tracker › "[BD] Headcount_upcoming" cột AN:AZ
 *    → thêm vào "Danh sách OB (draw)" (A:M). Chỉ thêm người chưa từng lấy về
 *    (nhớ trong sheet ẩn "Sync_Seen"), nên người đã xoá sẽ không bị thêm lại.
 * 2) syncTracking  : BD được lấy tự động (đã có email) → tự thêm vào
 *    "Diem_Danh" và "Master_Tracking" (kèm công thức). Track / Trainer / Team /
 *    Start_Date sửa ở "Danh sách OB (draw)" sẽ tự cập nhật sang 2 sheet kia.
 *    XOÁ người không đến ở "Danh sách OB (draw)" → tự xoá ở 2 sheet kia.
 * 3) syncQuizResults: đọc thẳng các sheet câu trả lời Form trong file
 *    "Kết quả bài quiz sau training" → thêm lượt làm còn thiếu vào
 *    Data_Topic1 / Data_Topic2 của Masterfile.
 *
 * CÀI ĐẶT: Masterfile › Tiện ích mở rộng › Apps Script › dán đè toàn bộ › Lưu
 *   › chạy syncAll để thử › chạy setupTriggers để bật tự động mỗi giờ.
 */

const CFG = {
  HC_ID: '122JM8NSS7FdBkz5n10I2zYtRz97-HPYpDiknw5YHRW4',
  HC_SHEET: '[BD] Headcount_upcoming',
  HC_FIRST_COL: 'AN',          // AN..AZ = 13 cột → A..M của "Danh sách OB (draw)"
  HC_NUM_COLS: 13,

  QUIZ_ID: '1LrzFY81FirN7hC3-yQwbnTt7pMyc0NDl8InQN1H4nIU',
  // Vị trí cột tính từ 0 (A=0, B=1, ...)
  T1: {
    dest: 'Data_Topic1', max: 13, passMin: 11,          // Pass khi Score > 10
    sources: [
      { sheet: 'Topic 1',           ts: 1, email: 2, score: 3, team: 4 },
      { sheet: 'Thi lại topic 1',   ts: 0, email: 1, score: 2, team: 3 },
      { sheet: 'Thi lại topic 1.1', ts: 0, email: 1, score: 2, team: 3 },
    ],
  },
  T2: {
    dest: 'Data_Topic2', max: '', passMin: 11,          // Pass khi Score >= 11
    sources: [
      { sheet: 'Topic 2 + survey', ts: 1, email: 2, score: 3, team: 4, gv: 22, csat: 23, text: 29 },
      { sheet: 'Thi lại topic 2',  ts: 0, email: 1, score: 2, team: 3, gv: 21, csat: 22, text: 28 },
    ],
  },

  OB_SHEET: 'Danh sách OB (draw)',
  OB_NOTE_COL: 18,             // R = Ghi_chu
  MT_SHEET: 'Master_Tracking',
  DD_SHEET: 'Diem_Danh',
  SEEN_SHEET: 'Sync_Seen',     // sheet ẩn: A Staff_ID | B Email | C Ngày lấy về | D Email đã thêm vào MT/DD
  LOG_SHEET: 'Sync_Log',
  TZ: 'Asia/Ho_Chi_Minh',
};

/* ====================== MENU & TRIGGER ====================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Auto Sync')
    .addItem('Đồng bộ ngay (DS onboard + Quiz)', 'syncAll')
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
    syncTracking();
    syncQuizResults();
    widenTrackingRanges_();
  } finally {
    lock.releaseLock();
  }
}

/* ====================== 1. DS BD UPCOMING ====================== */

function syncUpcomingBD() {
  const ob = sheet_(CFG.OB_SHEET);
  const hcFile = SpreadsheetApp.openById(CFG.HC_ID);
  const hc = hcFile.getSheetByName(CFG.HC_SHEET);
  if (!hc) throw new Error('Không thấy sheet ' + CFG.HC_SHEET + ' trong Headcount Tracker');
  const hcLast = hc.getLastRow();
  if (hcLast < 2) return log_('BD upcoming', 0, 'Nguồn trống');
  const fromTz = hcFile.getSpreadsheetTimeZone(), toTz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const src = hc.getRange(2, colIndex_(CFG.HC_FIRST_COL), hcLast - 1, CFG.HC_NUM_COLS)
    .getValues()
    .filter(r => isValid_(r[0]) && isValid_(r[1]))    // có Full Name + Staff ID
    .map(r => r.map(v => sameDay_(v, fromTz, toTz))); // giữ đúng ngày dù 2 file khác múi giờ

  const obLast = lastDataRow_(ob, 2);
  const existing = obLast >= 2 ? ob.getRange(2, 1, obLast - 1, 6).getValues() : [];
  const rowById = {};
  existing.forEach((r, i) => { const id = norm_(r[1]); if (id) rowById[id] = i + 2; });

  const seen = readSeen_();
  const known = new Set();
  seen.rows.forEach(s => { if (s.id) known.add(s.id); if (s.email) known.add(s.email); });
  if (seen.rows.length === 0) {                        // lần chạy đầu: coi DS đang có là "có sẵn"
    existing.forEach(r => {
      const id = norm_(r[1]), mail = norm_(r[2]);
      if (id || mail) { seen.add([id, mail, 'có sẵn', '']); known.add(id); known.add(mail); }
    });
  }

  const today = Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy');
  const toAdd = [];
  let updated = 0;

  src.forEach(r => {
    const id = norm_(r[1]), mail = norm_(r[2]);
    const row = rowById[id];
    if (row) {                                          // đã có → bổ sung email / đổi ngày OB
      const cur = ob.getRange(row, 3, 1, 4).getValues()[0];   // C..F
      if (!norm_(cur[0]) && mail) { ob.getRange(row, 3).setValue(r[2]); updated++; }
      if (r[5] instanceof Date && (!(cur[3] instanceof Date) || cur[3].getTime() !== r[5].getTime())) {
        ob.getRange(row, 6).setValue(r[5]); updated++;
      }
      return;
    }
    if (known.has(id) || (mail && known.has(mail))) return;   // đã từng lấy (đã bị xoá) → bỏ qua
    known.add(id); if (mail) known.add(mail);
    seen.add([id, mail, today, '']);
    toAdd.push(r);
  });

  if (toAdd.length) {
    const start = obLast + 1;
    ob.getRange(start, 1, toAdd.length, CFG.HC_NUM_COLS).setValues(toAdd);
    ob.getRange(start, CFG.OB_NOTE_COL, toAdd.length, 1)
      .setValues(toAdd.map(() => ['Auto từ HC Tracker ' + today]));
  }
  seen.save();
  log_('BD upcoming', toAdd.length, 'Thêm mới ' + toAdd.length + ', cập nhật ' + updated);
}

/* ====================== 2. DIEM_DANH + MASTER_TRACKING ====================== */

function syncTracking() {
  const ob = sheet_(CFG.OB_SHEET), mt = sheet_(CFG.MT_SHEET), dd = sheet_(CFG.DD_SHEET);

  const obLast = lastDataRow_(ob, 2);
  const obById = {};
  (obLast >= 2 ? ob.getRange(2, 1, obLast - 1, 15).getValues() : [])
    .forEach(r => { const id = norm_(r[1]); if (id) obById[id] = r; });

  const mtIdx = indexByEmail_(mt), ddIdx = indexByEmail_(dd);
  const seen = readSeen_();
  const del = { mt: [], dd: [] };
  const mtNew = [], ddNew = [];
  let changed = 0;

  seen.rows.forEach(s => {
    if (s.date === 'có sẵn' || !s.id) return;          // chỉ quản lý người do script lấy về
    const r = obById[s.id];

    if (!r) {                                           // trainer đã xoá khỏi DS OB → xoá ở MT/DD
      if (s.added) {
        if (mtIdx[s.added]) del.mt.push(mtIdx[s.added]);
        if (ddIdx[s.added]) del.dd.push(ddIdx[s.added]);
        s.added = ''; s.dirty = true;
      }
      return;
    }
    const email = norm_(r[2]);
    if (!email) return;                                 // chưa có email → đợi HR cập nhật

    // r: 0 Name, 2 Email, 5 Start, 8 Team, 13 Track, 14 Trainer
    if (!mtIdx[email]) mtNew.push(r); else changed += patch_(mt, mtIdx[email], { 2: r[0], 3: r[8], 4: r[13], 5: r[14], 6: r[5] });
    if (!ddIdx[email]) ddNew.push(r); else changed += patch_(dd, ddIdx[email], { 2: r[0], 3: r[8], 4: r[14], 5: r[5] });
    if (s.added !== email) { s.added = email; s.dirty = true; }
  });

  // Xoá từ dưới lên
  del.mt.sort((a, b) => b - a).forEach(row => mt.deleteRow(row));
  del.dd.sort((a, b) => b - a).forEach(row => dd.deleteRow(row));

  if (ddNew.length) {
    const start = lastDataRow_(dd, 1) + 1;
    dd.getRange(start, 1, ddNew.length, 12).setValues(ddNew.map(r =>
      [r[2], r[0], r[8], r[14], r[5], 'Chưa', 'Chưa', 'Chưa', 'Chưa', 'Chưa', 'Chưa', 'Chưa']));
  }
  if (mtNew.length) {
    const start = lastDataRow_(mt, 1) + 1;
    mt.getRange(start, 1, mtNew.length, 17).setValues(mtNew.map((r, i) => mtRow_(r, start + i)));
  }
  seen.save();
  log_('Diem_Danh / Master_Tracking', mtNew.length + ddNew.length + del.mt.length,
    'MT +' + mtNew.length + ', DD +' + ddNew.length + ', xoá ' + del.mt.length + ', cập nhật ' + changed + ' ô');
}

function mtRow_(r, n) {
  const T1 = 'Data_Topic1!$A:$A', T1F = 'Data_Topic1!$F:$F', T2 = 'Data_Topic2!$A:$A', T2F = 'Data_Topic2!$F:$F';
  return [
    r[2], r[0], r[8], r[13], r[14], r[5],                                   // A..F
    `=IF(F${n}="","",TODAY()-F${n})`,                                       // G
    'Chưa onboard',                                                         // H (trainer chọn)
    `=COUNTIF(${T1},A${n})`,                                                // I
    `=IF(I${n}=0,"Chưa làm",IF(COUNTIFS(${T1},A${n},${T1F},"Pass")>0,"Pass","Fail"))`,
    `=COUNTIF(${T2},A${n})`,                                                // K
    `=IF(K${n}=0,"Chưa làm",IF(COUNTIFS(${T2},A${n},${T2F},"Pass")>0,"Pass","Fail"))`,
    `=IF(F${n}="","Chưa có ngày bắt đầu",IF(AND(G${n}>=3,H${n}<>"Hoàn thành"),"⚠ Quá 3 ngày chưa xong Chặng 1",IF(AND(G${n}>=7,OR(J${n}="Fail",J${n}="Chưa làm")),"⚠ Quá 7 ngày chưa đạt Topic 1",IF(AND(G${n}>=14,OR(L${n}="Fail",L${n}="Chưa làm")),"⚠ Quá 14 ngày chưa đạt Topic 2","OK"))))`,
    `=IFERROR(VLOOKUP(A${n},Data_Topic2!$A:$O,15,0),"")`,                   // N
    `=IF(F${n}="","Chưa cập nhật",IF(AND(H${n}="Hoàn thành",J${n}="Pass",L${n}="Pass"),"Hoàn thành onboarding",IF(M${n}<>"OK","Cảnh báo","Đang onboarding")))`,
    '',                                                                     // P
    `=IF(F${n}="","",IF(G${n}<=2,"1. Orientation (D1-D2)",IF(G${n}<=15,"2. Acquisition (D3-D15)",IF(G${n}<=20,"3. Integration (D15-D20)",IF(G${n}<=30,"4. Acceleration (D20-D30)","5. Mastery")))))`,
  ];
}

// Ghi đè các ô nếu khác giá trị (chỉ khi OB có giá trị). cols: {cột: giá trị}
function patch_(sh, row, cols) {
  let n = 0;
  Object.keys(cols).forEach(c => {
    const v = cols[c];
    if (v === '' || v == null) return;
    const cell = sh.getRange(row, +c);
    if (cell.getFormula()) return;                      // không đụng ô công thức
    const cur = cell.getValue();
    const same = (v instanceof Date && cur instanceof Date) ? v.getTime() === cur.getTime() : String(cur) === String(v);
    if (!same) { cell.setValue(v); n++; }
  });
  return n;
}

// Mở rộng vùng công thức cũ của Master_Tracking ($A$2:$A$316 → $A$2:$A) để không bỏ sót dữ liệu mới
function widenTrackingRanges_() {
  const mt = sheet_(CFG.MT_SHEET);
  const last = lastDataRow_(mt, 1);
  if (last < 2) return;
  const rng = mt.getRange(2, 9, last - 1, 6);          // I..N
  const f = rng.getFormulas(), v = rng.getValues();
  let changed = 0;
  const out = f.map((row, i) => row.map((x, j) => {
    if (!x) return v[i][j];
    const y = x.replace(/(Data_Topic[12]!\$[A-Z])\$2:(\$[A-Z])\$\d+/g, '$1$$2:$2');
    if (y !== x) changed++;
    return y;
  }));
  if (changed) { rng.setValues(out); log_('Master_Tracking', changed, 'Mở rộng vùng công thức cũ'); }
}

/* ====================== 3. KẾT QUẢ QUIZ ====================== */

function syncQuizResults() {
  const quiz = SpreadsheetApp.openById(CFG.QUIZ_ID);
  syncTopic_(quiz, CFG.T1, false);
  syncTopic_(quiz, CFG.T2, true);
}

function syncTopic_(quiz, T, isT2) {
  const d = sheet_(T.dest);
  const fromTz = quiz.getSpreadsheetTimeZone(), toTz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  // Gom tất cả lượt làm từ các sheet Form
  const raw = [];
  T.sources.forEach(S => {
    const s = quiz.getSheetByName(S.sheet);
    if (!s || s.getLastRow() < 2) return;
    s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues().forEach(v => {
      const email = String(v[S.email] || '').trim();
      if (!email || !(v[S.ts] instanceof Date)) return;
      raw.push({ S, v, email, key: email.toLowerCase(), ts: sameClock_(v[S.ts], fromTz, toTz) });
    });
  });
  raw.sort((a, b) => a.ts - b.ts);

  // Số lượt đã có trong Masterfile theo email
  const dLast = lastDataRow_(d, 1);
  const have = {};
  (dLast >= 2 ? d.getRange(2, 1, dLast - 1, 1).getValues() : [])
    .forEach(r => { const k = norm_(r[0]); if (k) have[k] = (have[k] || 0) + 1; });

  // Mỗi email: thiếu bao nhiêu lượt thì thêm bấy nhiêu lượt MỚI NHẤT
  const byEmail = {};
  raw.forEach(x => (byEmail[x.key] = byEmail[x.key] || []).push(x));
  let toAdd = [];
  Object.keys(byEmail).forEach(k => {
    const list = byEmail[k], miss = list.length - (have[k] || 0);
    if (miss > 0) toAdd = toAdd.concat(list.slice(-miss));
  });
  toAdd.sort((a, b) => a.ts - b.ts);

  if (toAdd.length) {
    const start = dLast + 1;
    const rows = toAdd.map((x, i) => {
      const n = start + i, S = x.S, v = x.v, score = Number(v[S.score]) || 0;
      const base = [x.email, v[S.team], x.ts, score, T.max, score >= T.passMin ? 'Pass' : 'Fail', `=COUNTIF($A$2:A${n},A${n})`];
      if (!isT2) return base;
      const csat = v.slice(S.csat, S.csat + 6), text = v.slice(S.text, S.text + 4);
      return base.concat(
        [v[S.gv]], csat,
        [`=IFERROR(AVERAGE(VALUE(LEFT(I${n},1)),VALUE(LEFT(J${n},1)),VALUE(LEFT(K${n},1)),VALUE(LEFT(L${n},1)),VALUE(LEFT(M${n},1)),VALUE(LEFT(N${n},1))),"")`],
        text,
        ['I', 'J', 'K', 'L', 'M', 'N'].map(c => `=IFERROR(VALUE(LEFT(${c}${n},1)),"")`)
      );
    });
    d.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  }
  log_(T.dest, toAdd.length, 'Thêm ' + toAdd.length + ' lượt làm (nguồn có ' + raw.length + ')');
}

/* ====================== HELPERS ====================== */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Không thấy sheet ' + name);
  return sh;
}

function readSeen_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.SEEN_SHEET);
  const rows = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r =>
      rows.push({ id: norm_(r[0]), email: norm_(r[1]), date: String(r[2]), added: norm_(r[3]) }));
  }
  const fresh = [];
  return {
    rows,
    add(r) { const o = { id: r[0], email: r[1], date: r[2], added: r[3] }; rows.push(o); fresh.push(o); },
    save() {
      if (!sh) { sh = ss.insertSheet(CFG.SEEN_SHEET); sh.hideSheet(); }
      sh.getRange(1, 1, 1, 4).setValues([['Staff_ID', 'Email', 'Ngay_lay_ve', 'Da_them_MT_DD']]);
      if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows.map(o => [o.id, o.email, o.date, o.added]));
    },
  };
}

function indexByEmail_(sh) {
  const last = lastDataRow_(sh, 1), idx = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues()
    .forEach((r, i) => { const k = norm_(r[0]); if (k && !idx[k]) idx[k] = i + 2; });
  return idx;
}

// Ngày (không giờ): 18/11 ở file nguồn → 18/11 ở Masterfile, bất kể múi giờ mỗi file
function sameDay_(v, fromTz, toTz) {
  if (!(v instanceof Date) || fromTz === toTz) return v;
  return Utilities.parseDate(Utilities.formatDate(v, fromTz, 'yyyy-MM-dd'), toTz, 'yyyy-MM-dd');
}

// Ngày + giờ: 11:54 ở file nguồn → hiển thị 11:54 ở Masterfile
function sameClock_(v, fromTz, toTz) {
  if (!(v instanceof Date) || fromTz === toTz) return v;
  const f = 'yyyy-MM-dd HH:mm:ss';
  return Utilities.parseDate(Utilities.formatDate(v, fromTz, f), toTz, f);
}

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

function log_(what, count, msg) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.LOG_SHEET) || ss.insertSheet(CFG.LOG_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(['Thời gian', 'Hạng mục', 'Số dòng', 'Ghi chú']);
  sh.appendRow([new Date(), what, count, msg]);
}
