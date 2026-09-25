/***** ====== Bank Snapshot: จดยอดบัญชี ณ สิ้นวันที่ 24 (ก่อนส่งบิล) ======
 * เปิดผ่าน web app: ?page=snapshot&t=<token>[&asOf=yyyy-MM-dd]
 * n8n เช็กว่าจดครบหรือยัง: ?page=snapshot-status&t=<token>[&asOf=yyyy-MM-dd]
 *
 * ความหมายของ OpeningDate ในโซน BANK SNAPSHOT INPUT = "ยอด ณ สิ้นวัน"
 * Bank_Reconciliation_Calculations นับเงินเข้า-ออกที่ CashDate > OpeningDate และ <= EndingDate
 * จึงจดเช้าวันที่ 25 ได้ แต่ต้องใช้ยอดหลังรายการสุดท้ายของวันที่ 24 ไม่ใช่ยอดคงเหลือปัจจุบัน
 *****/
const SNAPSHOT_SHEET_NAME = 'Bank Reconciliation Dashboard';
const SNAPSHOT_FIRST_ROW = 41;
const SNAPSHOT_LAST_ROW = 238;
const SNAPSHOT_DAY = 24;
const SNAPSHOT_STATUS = 'OPENING_CAPTURED';
const SNAPSHOT_TOKEN_SHA256 = '0dd7580e44334dc5baf779aa078db7b32547f71d7d3a8a6b67bc58d4e424b7e8';

function snapshotTokenOk_(t) {
  if (!t) return false;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(t), Utilities.Charset.UTF_8);
  const hex = bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  return hex === SNAPSHOT_TOKEN_SHA256;
}

function snapshotTz_() {
  return Session.getScriptTimeZone() || 'Asia/Bangkok';
}

function snapshotYmd_(d) {
  return Utilities.formatDate(d, snapshotTz_(), 'yyyy-MM-dd');
}

// วันที่ 24 ล่าสุดที่ผ่านมาแล้ว (รวมวันนี้) — วันที่ 1–23 จะได้ 24 ของเดือนก่อน
function defaultSnapshotAsOf_() {
  const parts = snapshotYmd_(new Date()).split('-').map(Number);
  let y = parts[0], m = parts[1];
  if (parts[2] < SNAPSHOT_DAY) {
    m -= 1;
    if (m < 1) { m = 12; y -= 1; }
  }
  return y + '-' + String(m).padStart(2, '0') + '-' + SNAPSHOT_DAY;
}

function normalizeSnapshotAsOf_(asOf) {
  const s = String(asOf || '').trim();
  if (!s) return defaultSnapshotAsOf_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('asOf ต้องเป็น yyyy-MM-dd: ' + s);
  return s;
}

function snapshotDateFromYmd_(ymd) {
  const p = ymd.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]); // เที่ยงคืนตาม timezone ของสคริปต์ (Asia/Bangkok)
}

function snapshotCellYmd_(v) {
  if (v instanceof Date) return snapshotYmd_(v);
  return String(v || '').trim();
}

function readSnapshotRows_(sh) {
  const n = SNAPSHOT_LAST_ROW - SNAPSHOT_FIRST_ROW + 1;
  return sh.getRange(SNAPSHOT_FIRST_ROW, 1, n, 9).getValues().map(function (r, i) {
    return {
      row: SNAPSHOT_FIRST_ROW + i,
      month: r[0] instanceof Date ? Utilities.formatDate(r[0], snapshotTz_(), 'yyyy-MM') : String(r[0] || '').trim(),
      account: String(r[1] || '').trim(),
      open: r[2],
      close: r[3],
      openingDate: snapshotCellYmd_(r[5]),
      endingDate: snapshotCellYmd_(r[6]),
      capturedAt: r[7] instanceof Date ? Utilities.formatDate(r[7], snapshotTz_(), 'yyyy-MM-dd HH:mm') : String(r[7] || ''),
      status: String(r[8] || '').trim()
    };
  });
}

// บัญชีที่ต้องจด = บัญชีในรอบล่าสุดก่อนรอบนี้ (ตามลำดับเดิม) + บัญชีที่จดในรอบนี้แล้วแต่ไม่อยู่ในรอบก่อน
function snapshotAccountsFor_(rows, cycle) {
  const prevMonths = rows
    .filter(function (r) { return r.month && r.account && r.month < cycle; })
    .map(function (r) { return r.month; })
    .sort();
  const prevCycle = prevMonths.length ? prevMonths[prevMonths.length - 1] : '';
  const codes = [];
  rows.forEach(function (r) {
    if (r.month === prevCycle && r.account && codes.indexOf(r.account) < 0) codes.push(r.account);
  });
  rows.forEach(function (r) {
    if (r.month === cycle && r.account && codes.indexOf(r.account) < 0) codes.push(r.account);
  });
  return { prevCycle: prevCycle, codes: codes };
}

function loadAccountConfigMap_(ss) {
  const sh = ss.getSheetByName('_Config_Accounts');
  const map = {};
  if (!sh) return map;
  const v = sh.getDataRange().getValues();
  const h = v[0].map(function (x) { return String(x).trim(); });
  const ix = function (k) { return h.indexOf(k); };
  v.slice(1).forEach(function (r) {
    const code = String(r[ix('Code')] || '').trim();
    if (!code) return;
    map[code] = {
      accountNo: String(r[ix('AccountNo')] || '').trim(),
      bankLabel: String(r[ix('BankLabel')] || '').trim(),
      purpose: String(r[ix('Purpose')] || '').trim()
    };
  });
  return map;
}

// เงินเข้าที่ระบบเห็นแล้วหลังสิ้นวัน asOf (ใช้เตือนตอนจดเช้าวันที่ 25 ว่ายอดปัจจุบันจะเกินไปเท่าไร)
function ledgerInAfter_(ss, asOfYmd) {
  const sh = ss.getSheetByName('Bank_Reconciliation_Calculations');
  const out = {};
  if (!sh) return out;
  const p = asOfYmd.split('-').map(Number);
  const cutoffSerial = (Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
  const v = sh.getRange('A2:L5003').getValues();
  v.forEach(function (r) {
    const d = r[0], acct = String(r[2] || '').trim();
    if (!acct || r[3] !== 'IN' || r[11] !== true) return;
    let ymd;
    if (d instanceof Date) ymd = snapshotYmd_(d);
    else if (typeof d === 'number') ymd = d > cutoffSerial ? 'after' : 'before';
    else return;
    if (ymd === 'before' || (ymd !== 'after' && ymd <= asOfYmd)) return;
    const o = out[acct] || (out[acct] = { count: 0, amount: 0 });
    o.count += 1;
    o.amount += Number(r[4]) || 0;
  });
  return out;
}

function buildSnapshotState_(asOf) {
  const asOfYmd = normalizeSnapshotAsOf_(asOf);
  const cycle = asOfYmd.slice(0, 7);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SNAPSHOT_SHEET_NAME);
  const rows = readSnapshotRows_(sh);
  const acc = snapshotAccountsFor_(rows, cycle);
  const cfg = loadAccountConfigMap_(ss);
  const nowYmd = snapshotYmd_(new Date());
  const afterAsOf = nowYmd > asOfYmd;
  const ledgerAfter = afterAsOf ? ledgerInAfter_(ss, asOfYmd) : {};

  const accounts = acc.codes.map(function (code) {
    const c = cfg[code] || {};
    const cur = rows.filter(function (r) { return r.month === cycle && r.account === code; })[0];
    const last = rows.filter(function (r) {
      return r.account === code && r.month < cycle && r.close !== '' && r.close !== null;
    }).pop();
    return {
      code: code,
      bankLabel: c.bankLabel || '',
      purpose: c.purpose || '',
      last4: c.accountNo ? c.accountNo.slice(-4) : '',
      saved: cur && cur.open !== '' ? Number(cur.open) : null,
      savedAt: cur ? cur.capturedAt : '',
      lastClose: last ? Number(last.close) : null,
      lastCloseDate: last ? (last.endingDate || last.month) : '',
      ledgerAfter: ledgerAfter[code] || null
    };
  });
  const missing = accounts.filter(function (a) { return a.saved === null; }).map(function (a) { return a.code; });
  return {
    asOf: asOfYmd,
    cycle: cycle,
    prevCycle: acc.prevCycle,
    now: Utilities.formatDate(new Date(), snapshotTz_(), 'yyyy-MM-dd HH:mm'),
    afterAsOf: afterAsOf,
    accounts: accounts,
    expected: accounts.length,
    savedCount: accounts.length - missing.length,
    missing: missing,
    complete: accounts.length > 0 && missing.length === 0
  };
}

/***** web app routes (เรียกจาก doGet) *****/
function bankSnapshotPage_(e) {
  const p = (e && e.parameter) || {};
  if (!snapshotTokenOk_(p.t)) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">ลิงก์ไม่ถูกต้อง</p>');
  }
  const tpl = HtmlService.createTemplateFromFile('BankSnapshotPage');
  tpl.token = p.t;
  tpl.asOf = p.asOf || '';
  return tpl.evaluate()
    .setTitle('จดยอดบัญชี')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function bankSnapshotStatus_(e) {
  const p = (e && e.parameter) || {};
  if (!snapshotTokenOk_(p.t)) return jsonResponseRM_({ ok: false, error: 'unauthorized' });
  try {
    const s = buildSnapshotState_(p.asOf);
    // จดภายในวัน asOf เอง = อาจมีเงินเข้า-ออกหลังเวลาที่จดก่อนเที่ยงคืน ต้องให้ยืนยันยอดสิ้นวันอีกรอบ
    const sameDay = s.accounts
      .filter(function (a) { return a.saved !== null && String(a.savedAt).slice(0, 10) === s.asOf; })
      .map(function (a) { return { code: a.code, at: String(a.savedAt).slice(11) }; });
    return jsonResponseRM_({
      ok: true, asOf: s.asOf, cycle: s.cycle, expected: s.expected,
      savedCount: s.savedCount, missing: s.missing, complete: s.complete,
      capturedSameDay: sameDay
    });
  } catch (err) {
    return jsonResponseRM_({ ok: false, error: String(err && err.message || err) });
  }
}

/***** เรียกจากหน้าเว็บผ่าน google.script.run *****/
function getBankSnapshotForm(token, asOf) {
  if (!snapshotTokenOk_(token)) throw new Error('unauthorized');
  return buildSnapshotState_(asOf);
}

// entries: [{code, amount}] — amount ว่าง = ข้าม (ยังไม่จด), จดซ้ำ = แก้แถวเดิมของรอบนี้
function saveBankSnapshot(token, asOf, entries) {
  if (!snapshotTokenOk_(token)) throw new Error('unauthorized');
  const asOfYmd = normalizeSnapshotAsOf_(asOf);
  const cycle = asOfYmd.slice(0, 7);
  const clean = (entries || []).map(function (x) {
    const raw = String(x && x.amount != null ? x.amount : '').replace(/[,\s฿]/g, '');
    return { code: String(x && x.code || '').trim(), raw: raw, amount: Number(raw) };
  }).filter(function (x) { return x.code && x.raw !== ''; });
  clean.forEach(function (x) {
    if (!isFinite(x.amount)) throw new Error('ยอดของ ' + x.code + ' ไม่ใช่ตัวเลข: ' + x.raw);
  });
  if (!clean.length) throw new Error('ยังไม่ได้กรอกยอดสักบัญชี');

  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SNAPSHOT_SHEET_NAME);
    const rows = readSnapshotRows_(sh);
    const allowed = snapshotAccountsFor_(rows, cycle).codes;
    const openingDate = snapshotDateFromYmd_(asOfYmd);
    const now = new Date();
    let lastUsed = SNAPSHOT_FIRST_ROW - 1;
    rows.forEach(function (r) { if (r.month || r.account) lastUsed = r.row; });
    let nextRow = lastUsed + 1;

    clean.forEach(function (x) {
      if (allowed.indexOf(x.code) < 0) throw new Error('ไม่รู้จักบัญชี ' + x.code);
      const cur = rows.filter(function (r) { return r.month === cycle && r.account === x.code; })[0];
      let row;
      if (cur) {
        row = cur.row;
      } else {
        if (nextRow > SNAPSHOT_LAST_ROW) throw new Error('โซน BANK SNAPSHOT INPUT เต็มแล้ว (แถว ' + SNAPSHOT_LAST_ROW + ')');
        row = nextRow++;
        // Month ต้องเป็นข้อความ ไม่งั้น FILTER ที่เทียบกับ B4 คืน #N/A ทั้งหน้า
        sh.getRange(row, 1).setNumberFormat('@').setValue(cycle);
        sh.getRange(row, 2).setValue(x.code);
        sh.getRange(row, 5).setFormula('=IF(D' + row + '="","",D' + row + '-C' + row + ')');
      }
      sh.getRange(row, 3).setValue(x.amount);
      sh.getRange(row, 6).setValue(openingDate).setNumberFormat('yyyy-mm-dd');
      sh.getRange(row, 8).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
      sh.getRange(row, 9).setValue(SNAPSHOT_STATUS);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return buildSnapshotState_(asOfYmd);
}
