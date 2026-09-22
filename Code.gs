/**
 * Claude Team 帳號登記 — Google Apps Script 後台
 * 資料存放在綁定的 Google 試算表中（工作表「登記名單」）
 *
 * 使用方式：
 *   1. 在 Google 試算表 → 擴充功能 → Apps Script，貼上本檔內容
 *   2. 先執行一次 setup()（會建立工作表與標題列，並要求授權）
 *   3. 部署 → 新增部署作業 → 網頁應用程式
 *        執行身分：我    存取權：所有人
 *   4. 把產生的網址貼到 index.html 的 CONFIG.API_URL
 */

// ================== 設定區 ==================
const SHEET_NAME   = '登記名單';
const EMAIL_DOMAIN = 'gm.student.ncut.edu.tw';
const SID_PATTERN  = /^[A-Z0-9]{6,12}$/;   // 學號格式，請與前端一致
const CAPACITY     = 0;                     // 名額上限，0 = 不限
const DEADLINE     = '2026-09-28T23:59:59+08:00';                    // 例：'2026-10-15T23:59:00+08:00'，留空 = 不限
const SEND_CONFIRM_EMAIL = false;           // true = 登記成功後寄確認信給學生
const NOTIFY_ADMIN_EMAIL = '';              // 填入信箱則每筆新登記通知管理者，留空 = 不通知
// ===========================================

const HEADERS = ['登記時間', '姓名', '學號', 'E-mail', '狀態', '最後更新', '邀請已寄出', '備註'];
const COL = { time: 1, name: 2, sid: 3, email: 4, status: 5, updated: 6, invited: 7, note: 8 };

/** 第一次使用請手動執行：建立工作表與格式 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#f6e4dc');
  sh.setFrozenRows(1);
  sh.getRange('C:C').setNumberFormat('@');   // 學號以文字儲存，避免前導 0 消失
  sh.setColumnWidths(1, HEADERS.length, 150);
  SpreadsheetApp.flush();
  Logger.log('setup 完成：' + ss.getUrl());
}

/* ---------------- GET：統計 / 查詢 ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    const sh = getSheet_();
    if (p.action === 'stats') {
      return json_({ ok: true, count: countActive_(sh) });
    }
    if (p.action === 'check') {
      const sid = normSid_(p.stuid || p.id);
      const row = findRowBySid_(sh, sid);
      if (!row) return json_({ ok: true, found: false });
      const v = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      return json_({
        ok: true, found: true,
        name: maskName_(String(v[COL.name - 1])),
        email: maskEmail_(String(v[COL.email - 1])),
        status: v[COL.status - 1]
      });
    }
    return json_({ ok: true, service: 'Claude Team 登記系統運作中' });
  } catch (err) {
    return json_({ ok: false, message: '伺服器錯誤：' + err.message });
  }
}

/* ---------------- POST：登記 / 取消 ---------------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);   // 避免多人同時送出寫入衝突
    const d = JSON.parse(e.postData.contents || '{}');
    const sh = getSheet_();
    if (d.action === 'register') return json_(register_(sh, d));
    if (d.action === 'cancel')   return json_(cancel_(sh, d));
    return json_({ ok: false, message: '未知的操作' });
  } catch (err) {
    return json_({ ok: false, message: '伺服器忙碌或錯誤，請稍後再試。' });
  } finally {
    lock.releaseLock();
  }
}

function register_(sh, d) {
  if (DEADLINE && new Date() > new Date(DEADLINE)) return { ok: false, message: '登記已截止。' };

  const name  = String(d.name || '').trim().slice(0, 30);
  const sid   = normSid_(d.sid);
  const email = String(d.email || '').trim().toLowerCase();

  if (!name) return { ok: false, message: '請填寫姓名。' };
  if (!SID_PATTERN.test(sid)) return { ok: false, message: '學號格式不正確。' };
  const re = new RegExp('^[a-z0-9._-]{2,40}@' + EMAIL_DOMAIN.replace(/\./g, '\\.') + '$');
  if (!re.test(email)) return { ok: false, message: '僅接受 @' + EMAIL_DOMAIN + ' 信箱。' };

  // 同一 Email 不可被不同學號使用
  const emailRow = findRowByEmail_(sh, email);
  const sidRow   = findRowBySid_(sh, sid);
  if (emailRow && emailRow !== sidRow) return { ok: false, message: '此 E-mail 已被其他學號登記，請洽管理者。' };

  const now = new Date();
  if (sidRow) {
    // 已存在 → 更新資料
    sh.getRange(sidRow, COL.name, 1, 4).setValues([[name, sid, email, '已登記']]);
    sh.getRange(sidRow, COL.updated).setValue(now);
    return { ok: true, updated: true, message: '你的登記資料已更新！' };
  }

  if (CAPACITY > 0 && countActive_(sh) >= CAPACITY) return { ok: false, message: '名額已滿。' };

  sh.appendRow([now, name, sid, email, '已登記', now, false, '']);
  const last = sh.getLastRow();
  sh.getRange(last, COL.sid).setNumberFormat('@').setValue(sid);
  sh.getRange(last, COL.invited).insertCheckboxes();   // 「邀請已寄出」核取方塊，方便管理者勾選

  if (SEND_CONFIRM_EMAIL) {
    try {
      MailApp.sendEmail(email, '【Claude Team】登記成功通知',
        name + ' 同學您好：\n\n您已完成 Claude Team 帳號登記。\n學號：' + sid +
        '\n\n後續邀請信將寄至此信箱，請留意收件。\n\n（本信由系統自動寄出）');
    } catch (err) { /* 寄信失敗不影響登記 */ }
  }
  if (NOTIFY_ADMIN_EMAIL) {
    try { MailApp.sendEmail(NOTIFY_ADMIN_EMAIL, '[Claude Team] 新登記：' + name, name + ' / ' + sid + ' / ' + email); } catch (err) {}
  }
  return { ok: true, updated: false, message: '登記成功！' };
}

function cancel_(sh, d) {
  const sid = normSid_(d.sid);
  const email = String(d.email || '').trim().toLowerCase();
  const row = findRowBySid_(sh, sid);
  if (!row) return { ok: false, message: '查無此學號的登記資料。' };
  const saved = String(sh.getRange(row, COL.email).getValue()).toLowerCase();
  if (saved !== email) return { ok: false, message: '學號與 E-mail 不相符。' };
  sh.getRange(row, COL.status).setValue('已取消');
  sh.getRange(row, COL.updated).setValue(new Date());
  return { ok: true, message: '已取消登記。' };
}

/* ---------------- 工具函式 ---------------- */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { setup(); sh = ss.getSheetByName(SHEET_NAME); }
  return sh;
}
function colValues_(sh, col) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, col, n, 1).getValues().map(r => String(r[0]));
}
function findRowBySid_(sh, sid) {
  const i = colValues_(sh, COL.sid).map(normSid_).indexOf(sid);
  return i < 0 ? 0 : i + 2;
}
function findRowByEmail_(sh, email) {
  const i = colValues_(sh, COL.email).map(s => s.toLowerCase()).indexOf(email);
  return i < 0 ? 0 : i + 2;
}
function countActive_(sh) {
  return colValues_(sh, COL.status).filter(s => s === '已登記').length;
}
function normSid_(v) { return String(v || '').trim().toUpperCase().replace(/\s+/g, ''); }
function maskName_(n) {
  if (n.length <= 1) return n;
  if (n.length === 2) return n[0] + '○';
  return n[0] + '○'.repeat(n.length - 2) + n.slice(-1);
}
function maskEmail_(e) { const p = e.split('@'); return p[0].slice(0, 2) + '***@' + (p[1] || ''); }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 管理用：把「已登記」的 Email 列成一欄，方便複製貼到 Claude Team 邀請 */
function exportActiveEmails() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rows = sh.getRange(2, 1, n, HEADERS.length).getValues()
    .filter(r => r[COL.status - 1] === '已登記').map(r => [r[COL.email - 1]]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = ss.getSheetByName('邀請清單') || ss.insertSheet('邀請清單');
  out.clear();
  out.getRange(1, 1).setValue('E-mail（已登記，共 ' + rows.length + ' 人）').setFontWeight('bold');
  if (rows.length) out.getRange(2, 1, rows.length, 1).setValues(rows);
}

/** 在試算表上方加入選單 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Claude Team')
    .addItem('初始化工作表', 'setup')
    .addItem('產生邀請 Email 清單', 'exportActiveEmails')
    .addToUi();
}
