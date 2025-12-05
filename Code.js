/***** CONFIG: โฟลเดอร์รายงาน Horganice และโฟลเดอร์ Statement CSV *****/
const HORG_FOLDER_ID = "1aFxmXNgQQKt3gl2Yk-FsQedGTUBnqPMo"; // your folder (Horganice XLS/XLSX)
const BANK_FOLDER_ID = '1KRfvhgw1Xw26arN_yvj9-_BUKfO-XfJu';  // folder for bank CSVs

/***** เมนูบน Google Sheets *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Mama Mansion")
    .addItem("📥 Import Horganice Report (XLS)", "importHorganice")
    .addItem('📥 Import Bank CSV (3 บัญชี)', 'importBankCsv')
    .addToUi();
}

/***** BILLING CYCLE HELPER (match PAY_RENT: 24th onward = next month) *****/
function getBillingYmForDate_(d){
  const tz  = Session.getScriptTimeZone() || 'Asia/Bangkok';
  const y   = Number(Utilities.formatDate(d, tz, 'yyyy'));
  const m   = Number(Utilities.formatDate(d, tz, 'MM')) - 1; // 0-based
  const day = Number(Utilities.formatDate(d, tz, 'dd'));

  let targetY = y;
  let targetM = m;
  if (day >= 24) {
    targetM += 1;
    if (targetM > 11) { targetM = 0; targetY += 1; }
  }
  return targetY + '-' + String(targetM + 1).padStart(2, '0'); // YYYY-MM
}

/***** ====== Horganice → Horga_Bills ====== *****/
function importHorganice() {
  const ui = SpreadsheetApp.getUi();

  // 1) newest .xls/.xlsx in folder
  const folder = DriveApp.getFolderById(HORG_FOLDER_ID);
  const files = folder.getFiles();
  let latest = null, latestTs = 0;
  while (files.hasNext()) {
    const f = files.next();
    const n = f.getName().toLowerCase();
    if (!(n.endsWith(".xls") || n.endsWith(".xlsx"))) continue;
    const ts = f.getLastUpdated().getTime();
    if (ts > latestTs) { latestTs = ts; latest = f; }
  }
  if (!latest) { ui.alert("No XLS/XLSX report found in the folder."); return; }

  // 2) convert Excel -> temp Google Sheet (Advanced Drive service must be ON)
  const blob = latest.getBlob();
  const temp = Drive.Files.insert(
    { title: `TEMP_${new Date().toISOString()}`, mimeType: MimeType.GOOGLE_SHEETS },
    blob
  );
  const tempId = temp.id;

  try {
    const tempSS = SpreadsheetApp.openById(tempId);

    // 3) choose the sheet that actually has a table
    const sheets = tempSS.getSheets();
    let chosen = sheets[0], bestScore = -1;
    sheets.forEach(sh => {
      const r = sh.getDataRange().getValues();
      if (!r || r.length < 2) return;
      const score = r.length * (r[0] ? r[0].length : 0);
      if (score > bestScore) { bestScore = score; chosen = sh; }
    });

    const all = chosen.getDataRange().getValues();

    // 4) find header row by looking for "ห้อง" or "Room"
    let headerRow = -1;
    for (let i = 0; i < all.length; i++) {
      const row = all[i].map(x => String(x || "").trim());
      if (row.some(h => /^(room|ห้อง)$/i.test(h))) { headerRow = i; break; }
    }
    if (headerRow === -1) { ui.alert("Cannot find a header row (no 'ห้อง' / 'Room')."); return; }

    const header = all[headerRow].map(h => String(h || "").trim());

    // 5) locate key columns and charge columns
    const idxRoom   = findHeaderIndex(header, [/^room$/i, /^ห้อง$/i]);
    const idxTenant = findHeaderIndex(header, [/tenant|name/i, /ผู้เช่า|ชื่อ/i]);
    const idxDue    = findHeaderIndex(header, [/due|date/i, /ครบกำหนด|กำหนดชำระ|วันที่/i]); // optional
    const idxTotal  = findHeaderIndex(header, [/รวมสุทธิ|ยอดรวม|ต้องชำระ|^รวม$/i]); // prefer explicit total column

    const chargeMatchers = [
      /amount|total/i,
      /ค่า[เชเ]่?า/i,            // ค่าเช่า/ค่าเช่าห้อง
      /ค่าเช่าห้อง/i,
      /ค่าน้ำ/i,
      /ค่าไฟฟ้า|ไฟฟ้า/i,
      /ค่าบริการ|ค่าดูแล|service/i,
      /ค่าปรับ|ปรับ/i,
      /อินเทอร์เน็ต|internet/i,
      /ที่จอด|parking/i,
      /อื่นๆ|misc/i,
      /รวมสุทธิ|ยอดรวม|ต้องชำระ/i
    ];

    const chargeColIdx = [];
    header.forEach((h, i) => {
      const hit = chargeMatchers.some(re => re.test(h));
      if (hit && i !== idxRoom) chargeColIdx.push(i);
    });

    if (idxRoom < 0 || chargeColIdx.length === 0) {
      ui.alert("Missing required columns: 'ห้อง/Room' and at least one charge column.");
      return;
    }

    // 6) build output rows (NO clearing — we will upsert)
    const monthStr = getBillingYmForDate_(new Date(latestTs)); // align with PAY_RENT billing window
    const rowsToUpsert = []; // each is an array in the schema below

    // schema
    const SCHEMA = ['BillID','Room','Tenant','Month','Type','AmountDue','DueDate',
                    'Status','PaidAt','SlipID','Account','BankMatchStatus','ChargeItems','Notes'];

    for (let r = headerRow + 1; r < all.length; r++) {
      const row = all[r];
      const room = toStr(row[idxRoom]);
      if (!room) continue;

      // Skip subtotal rows
      if (/^รวม|total|summary/i.test(room)) continue;

      const tenant = idxTenant >= 0 ? toStr(row[idxTenant]) : "";

      let amountDue = 0;
      let hasAny = false;
      const chargeParts = [];

      // Prefer explicit "รวม/ยอดรวม" column if present
      if (idxTotal >= 0) {
        const num = toNumber(row[idxTotal]);
        if (num != null && !isNaN(num) && num !== 0) {
          hasAny = true;
          amountDue = num;
          chargeParts.push(`${header[idxTotal]} ${num}`);
        }
      }

      // Fallback to summing charge columns if no usable total
      if (!hasAny) {
        chargeColIdx.forEach(i => {
          const val = row[i];
          const num = toNumber(val);
          if (num != null && !isNaN(num) && num !== 0) {
            hasAny = true;
            amountDue += num;
            chargeParts.push(`${header[i]} ${num}`);
          }
        });
      }

      if (!hasAny) continue;

      const dueStr  = idxDue >= 0 ? formatAsDateString(row[idxDue]) : "";
      const account = getAccountFromRoom_(room);  // keep your original logic
      const billId  = `${monthStr}-${room}`;

      rowsToUpsert.push([
        billId,
        room,
        tenant,
        monthStr,
        'Rent',
        amountDue,
        dueStr,
        'Unpaid',
        '',
        '',
        account,
        '',
        chargeParts.join('; '),
        `Imported: ${latest.getName()}`
      ]);
    }

    // 7) Upsert into Horga_Bills
    const master = SpreadsheetApp.getActiveSpreadsheet();
    const sh = master.getSheetByName("Horga_Bills") || master.insertSheet("Horga_Bills");

    // ensure header present and exact order
    const existingLastRow = sh.getLastRow();
    let headerOk = false;
    if (existingLastRow >= 1) {
      const firstRow = sh.getRange(1,1,1,SCHEMA.length).getValues()[0].map(x => String(x||''));
      headerOk = SCHEMA.every((h, i) => (firstRow[i] || '') === h);
    }
    if (!headerOk) {
      sh.clear(); // just clear this sheet once to place the correct header (not every import)
      sh.getRange(1,1,1,SCHEMA.length).setValues([SCHEMA]);
    }

    // build BillID -> rowIndex map (existing)
    const lastRow = sh.getLastRow();
    const map = new Map();
    if (lastRow > 1) {
      const existing = sh.getRange(2,1,lastRow-1,SCHEMA.length).getValues();
      const cBillId = 1; // column A in the sheet = BillID
      for (let i=0;i<existing.length;i++){
        const id = String(existing[i][cBillId-1]||'').trim();
        if (id) map.set(id, i + 2); // store sheet row index
      }
    }

    let inserted = 0, updated = 0;
    rowsToUpsert.forEach(arr => {
      const billId = String(arr[0]||'').trim();
      const hitRow = map.get(billId);
      if (hitRow) {
        // update in place (full row in schema)
        sh.getRange(hitRow, 1, 1, SCHEMA.length).setValues([arr]);
        updated++;
      } else {
        sh.appendRow(arr);
        inserted++;
      }
    });

    ui.alert(`Imported ${rowsToUpsert.length} bills from "${latest.getName()}".\n` +
             `Upserts → inserted: ${inserted}, updated: ${updated}`);

  } catch (e) {
    SpreadsheetApp.getUi().alert(`Import failed: ${e}`);
  } finally {
    try { DriveApp.getFileById(tempId).setTrashed(true); } catch (_) {}
  }
}

/** ===== helpers you already have in your file (kept for clarity) =====
 *  - findHeaderIndex(header, patterns)
 *  - toStr(v)
 *  - toNumber(v)
 *  - formatAsDateString(v)
 *  - getAccountFromRoom_(room)
 *  Keep using your existing versions; no changes needed.
 */


/***** NEW HELPER FUNCTION *****/
/**
 * Maps a room number (e.g., "A101", "B305") to an account code based on the floor.
 * Assumes room format is [BuildingLetter(s)][FloorNumber][RoomNumber] e.g., "A101", "B305"
 * @param {string} roomStr - The room number.
 * @returns {string} The corresponding account code (KKK+, TMK+, KGSI, KBIZ) or "".
 */
function getAccountFromRoom_(roomStr) {
  if (!roomStr) return "";
  
  const roomUpper = String(roomStr).toUpperCase().trim();
  
  // This regex looks for optional letters at the start, followed by ONE digit.
  // This digit is assumed to be the floor.
  // ^[A-Z]* -> Optional letters (A, B, AB, etc.) at the start.
  // (\d)      -> Captures the single digit that follows.
  const floorMatch = roomUpper.match(/^[A-Z]*(\d)/); 

  if (floorMatch && floorMatch[1]) {
    const floorDigit = floorMatch[1]; // This will be '1', '2', '3', '4', or '5'
    
    switch (floorDigit) {
      case '1':
        return "KKK+";
      case '2':
        return "MAK+";
      case '3':
        return "KGSI";
      case '4':
        return "GSB";
      case '5':
        return "GSB";
      default:
        return ""; // Floor 0, 6, 7, etc. get no account
    }
  }
  
  // Log if we can't figure out the floor
  Logger.log(`Could not determine floor for room: ${roomStr}`);
  return ""; // No digit found, or format is unexpected
}


/******** helpers (Horganice) ********/
function findHeaderIndex(headerArr, patterns) {
  for (let i = 0; i < headerArr.length; i++) {
    const h = headerArr[i];
    for (const re of patterns) if (re.test(h)) return i;
  }
  return -1;
}
function toStr(v){ return String(v == null ? "" : v).trim(); }
function toNumber(v){
  if (v == null || v === "") return null;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/[^\d.-]/g, ""));
  return isNaN(n) ? null : n;
}
function formatAsDateString(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (typeof v === "number") {
    // Handle Excel serial date format
    const excelEpoch = new Date(Date.UTC(1899,11,30));
    const jsDate = new Date(excelEpoch.getTime() + v * 86400000);
    return Utilities.formatDate(jsDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = String(v);
  // Try to parse common date strings
  const m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  // ----- FIX: Corrected m3] to m[3] -----
  if (m) return `${m[1]}-${("0"+m[2]).slice(-2)}-${("0"+m[3]).slice(-2)}`;
  // ----------------------------------------
  return s; // return as-is if unparseable
}

/***** อ่านแผ่น Rooms → map room → account code *****/
// This function is NO LONGER USED by importHorganice, 
// but left here in case other parts of your script use it.
function loadRoomAccountMap_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Rooms');
  if (!sh) return {};

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return {};

  const head = values[0].map(v => String(v||'').trim().toLowerCase());
  const cRoom = head.findIndex(h => /^room$/.test(h) || h.includes('ห้อง'));
  const cAcct = head.findIndex(h => /^account$/.test(h) || h.includes('บัญชี'));
  if (cRoom < 0 || cAcct < 0) return {};

  const map = {};
  for (let i = 1; i < values.length; i++){
    const r = values[i];
    const room = String(r[cRoom]||'').toUpperCase().trim();
    if (!room) continue;
    const code = String(r[cAcct]||'').toUpperCase().trim(); // KKK+ / KBIZ / KGSI
    map[room] = code;
  }
  return map;
}

/***** ====== Bank CSV → Bank_Transactions ====== *****/

/** สร้าง/ดึงชีต Bank_Transactions **/
function ensureBankTxnSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Bank_Transactions');
  if (!sh) {
    sh = ss.insertSheet('Bank_Transactions');
    sh.getRange(1,1,1,10).setValues([[
      'TxnId','Date','Account','Amount','Type','Ref','Description','LinkedBillId','LinkedAt','Notes'
    ]]);
  }
  return sh;
}

/** MD5 → hex (ทำ TxnId เสถียร แม้ไม่มี Ref) **/
function md5Hex_(s){
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
function toYMD2_(v){
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  let m = s.match(/\b([0-3]?\d)[\/\-]([01]?\d)[\/\-](\d{4})\b/); // d/m/y
  if (m) return `${m[3]}-${('0'+m[2]).slice(-2)}-${('0'+m[1]).slice(-2)}`;
  m = s.match(/\b(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\b/); // y/m/d
  if (m) return `${m[1]}-${('0'+m[2]).slice(-2)}-${('0'+m[3]).slice(-2)}`;
  return s;
}
function toNumberLoose_(v){
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[\u00A0\s,]/g,''); // ตัดช่องว่าง/คอมมา
  const n = Number(s);
  return isNaN(n) ? null : n;
}

/** เลือกไฟล์ CSV ล่าสุดในโฟลเดอร์ **/
function pickLatestCsv_(folderId){
  const folder = DriveApp.getFolderById(folderId);
  const it = folder.getFiles();
  let latest=null, latestTs=0;
  while (it.hasNext()){
    const f = it.next();
    if (!String(f.getName()).toLowerCase().endsWith('.csv')) continue;
    const ts = f.getLastUpdated().getTime();
    if (ts > latestTs){ latest = f; latestTs = ts; }
  }
  return latest;
}

function detectDelimiter_(text){
  const first = (text.split(/\r?\n/)[0] || '');
  const cands = [',',';','\t','|'];
  let best=',', score=-1;
  for (const d of cands){
    const cnt = (first.match(new RegExp('\\' + d,'g')) || []).length;
    if (cnt > score){ score = cnt; best = d; }
  }
  return best;
}

/** อ่าน CSV เป็น array of rows (ลอง UTF-8 ก่อน, เพี้ยนค่อยลอง windows-874) **/
function readCsv_(file){
  let txt = file.getBlob().getDataAsString('UTF-8');
  const bad = (txt.match(/\uFFFD/g) || []).length;
  // If many replacement characters, try a common Thai encoding
  if (bad > 5) { 
    try { txt = file.getBlob().getDataAsString('windows-874'); } catch(e) {}
  }
  
  const delim = detectDelimiter_(txt);
  return { rows: Utilities.parseCsv(txt, delim), delim: delim };
}

/** เดาโครง header ธนาคาร (รองรับไฟล์ที่ไม่มี Ref) **/
function headerMap_(hdr, sampleRows){
  const low = hdr.map(h => String(h||'').trim().toLowerCase());
  const pos = (cands)=> {
    for (const c of cands) {
      const i = low.findIndex(x => x.includes(c));
      if (i >= 0) return i;
    }
    return -1;
  };
  let idxDate = pos(['date','วันที่','transaction date','วัน-เวลา']);
  let idxTime = pos(['time','เวลา']); // อาจไม่ได้ใช้
  let idxCredit = pos(['credit','ฝาก']);
  let idxDebit  = pos(['debit','ถอน']);
  let idxAmount = pos(['amount','จำนวนเงิน','ยอดเงิน']);
  let idxType   = pos(['type','ประเภทรายการ','db/cr','credit/debit','cr/db','code']);
  let idxDesc   = pos(['description','details','transaction','รายการ','คำอธิบาย','รายละเอียด']);
  let idxRef    = pos(['ref','reference','reference no','เลขอ้างอิง','หมายเลขอ้างอิง']);

  // If credit/debit not found, try to infer from sample data
  if (idxCredit < 0 && idxDebit < 0 && idxAmount >= 0 && idxType < 0) {
      let hasPos = false, hasNeg = false;
      for (const r of sampleRows) {
          if (!r[idxAmount]) continue;
          const n = toNumberLoose_(r[idxAmount]);
          if (n > 0) hasPos = true;
          if (n < 0) hasNeg = true;
      }
      // If data has both positive and negative, we assume it's a single amount column
      // and type must be inferred by sign.
      if (!hasPos || !hasNeg) {
          // If only positive, maybe it's KBank style with separate Cr/Db columns
          // that just aren't named well. Let's guess.
          const guessCr = low.findIndex(h => h.includes('เครดิต'));
          if (guessCr >= 0) idxCredit = guessCr;
          const guessDb = low.findIndex(h => h.includes('เดบิต'));
          if (guessDb >= 0) idxDebit = guessDb;
      }
  }

  return { idxDate, idxTime, idxCredit, idxDebit, idxAmount, idxType, idxDesc, idxRef };
}

/** แปลงหนึ่งแถว CSV → ระเบียนมาตรฐาน (เอาเฉพาะ CREDIT) **/
/***** แปลงหนึ่งแถว → ระเบียนเครดิตมาตรฐาน *****/
function normalizeTxnRow_(row, map, accountCode){
  const get = (i)=> (i>=0 && i < row.length) ? row[i] : '';
  const dateRaw = get(map.idxDate);
  if (!dateRaw) return null;

  const dateYMD = toYMD2_(dateRaw);
  let desc = String(get(map.idxDesc)||'').trim();
  let ref  = String(get(map.idxRef)||'').trim();
  const typeRaw = String(get(map.idxType)||'').trim().toUpperCase();

  // ตีความยอดเงิน/ประเภท
  let amount = null, type = 'CREDIT';

  if (map.idxCredit>=0 || map.idxDebit>=0){
    // Case 1: Separate Credit and Debit columns
    const cr = toNumberLoose_(get(map.idxCredit));
    const db = toNumberLoose_(get(map.idxDebit));
    if (cr != null && cr > 0){ amount = cr; type = 'CREDIT'; }
    else if (db != null && db > 0){ amount = db; type = 'DEBIT'; }
  } else if (map.idxAmount>=0){
    // Case 2: Single Amount column
    const amt = toNumberLoose_(get(map.idxAmount));
    if (amt == null) return null;
    amount = Math.abs(amt);
    
    if (/DB|DEBIT|ถอน/i.test(typeRaw)) type = 'DEBIT'; // Type column says DEBIT
    else if (/CR|CREDIT|ฝาก/i.test(typeRaw)) type = 'CREDIT'; // Type column says CREDIT
    else if (amt < 0) type = 'DEBIT'; // Negative amount means DEBIT
    else type = 'CREDIT'; // Positive amount means CREDIT
  }

  if (!dateYMD || !amount || isNaN(amount) || amount === 0) return null;
  if (type !== 'CREDIT') return null; // ใช้เฉพาะเงินเข้า

  const descKey = desc.replace(/\s+/g,' ').toLowerCase();
  
  // Create a stable ID based on key fields
  const txnId = md5Hex_([accountCode, dateYMD, amount, descKey].join('|'));

  return {
    TxnId: txnId,
    Date: dateYMD,
    Account: accountCode,
    Amount: amount,
    Type: 'CREDIT',
    Ref: ref || '',
    Description: desc || '',
    LinkedBillId: '',
    LinkedAt: '',
    Notes: ''
  };
}

function importBankCsv(){
  const ui = SpreadsheetApp.getUi();
  const ans = ui.prompt('Import Bank CSV', 'ใส่รหัสบัญชี: KKK+ / KBIZ / KGSI', ui.ButtonSet.OK_CANCEL);
  if (ans.getSelectedButton() !== ui.Button.OK) return;
  const accountCode = (ans.getResponseText()||'').trim().toUpperCase();
  // ----- CHANGE: Added TMK+ as a valid account code -----
  if (!/^(KKK\+|KBIZ|KGSI|TMK\+)$/.test(accountCode)) { 
    ui.alert('รหัสบัญชีไม่ถูกต้อง (ต้องเป็น KKK+, KBIZ, KGSI, หรือ TMK+)'); 
    return; 
  }
  // -----------------------------------------------------

  const file = pickLatestCsv_(BANK_FOLDER_ID);
  if (!file){ ui.alert('ไม่พบไฟล์ .csv ในโฟลเดอร์'); return; }

  const { rows, delim } = readCsv_(file);
  if (!rows || rows.length < 2){ ui.alert('ไฟล์ว่างหรืออ่านไม่ได้'); return; }

  // Find the first row that looks like a header (has non-numeric values)
  let headerRowIndex = 0;
  let header = [];
  for(let i=0; i<rows.length; i++){
      const row = rows[i];
      if (row.some(cell => isNaN(Number(String(cell||'').replace(/[,]/g, ''))) && String(cell||'').trim() !== "" )) {
          header = row.map(x => String(x||'').trim());
          headerRowIndex = i;
          break;
      }
  }

  if (header.length === 0) { ui.alert('ไม่สามารถหาแถว Header ได้'); return; }
  
  const dataRows = rows.slice(headerRowIndex + 1);
  const map = headerMap_(header, dataRows.slice(0, 200)); // ใช้ตัวอย่าง 200 แถวช่วยเดา

  // Check if essential columns were found
  if (map.idxDate < 0 || (map.idxCredit < 0 && map.idxDebit < 0 && map.idxAmount < 0) || map.idxDesc < 0) {
      ui.alert(
          `ไม่พบคอลัมน์ที่จำเป็น:\n` +
          `Date: ${map.idxDate >= 0 ? '✔️' : '❌'}\n` +
          `Amount (Credit/Debit/Amount): ${(map.idxCredit >= 0 || map.idxDebit >= 0 || map.idxAmount >= 0) ? '✔️' : '❌'}\n` +
          `Description: ${map.idxDesc >= 0 ? '✔️' : '❌'}\n\n` +
          `แมปที่ได้: Date:${map.idxDate}, Cr:${map.idxCredit}, Db:${map.idxDebit}, Amt:${map.idxAmount}, Desc:${map.idxDesc}`
      );
      return;
  }


  const sh = ensureBankTxnSheet_();
  const existing = sh.getDataRange().getValues();
  const H = {};
  if (existing.length > 0) {
    existing[0].forEach((h,i)=> H[String(h).trim()] = i);
  } else {
    ui.alert('Sheet "Bank_Transactions" ไม่มี Header'); return;
  }
  
  // Build a set of existing transactions to prevent duplicates
  const setTxnId = new Set();
  const setCombo = new Set(); // Fallback check
  for (let i=1;i<existing.length;i++){
    const r = existing[i];
    if (!r[H['TxnId']]) continue; // Skip if no TxnId

    const tid = String(r[H['TxnId']]||'').trim();
    if (tid) setTxnId.add(tid);
    
    const combo = [
      String(r[H['Account']]||'').trim().toUpperCase(),
      toYMD2_(r[H['Date']]), // Normalize date for comparison
      Number(r[H['Amount']]||0).toFixed(2),
      String(r[H['Description']]||'').trim().replace(/\s+/g,' ').toLowerCase()
    ].join('|');
    setCombo.add(combo);
  }

  let parsed = 0, creditable = 0, appended = 0;
  const append = [];
  for (const row of dataRows){
    if (row.every(cell => String(cell||'').trim() === '')) continue; // Skip empty rows
    parsed++;
    const rec = normalizeTxnRow_(row, map, accountCode);
    if (!rec) continue;
    creditable++;

    if (setTxnId.has(rec.TxnId)) continue;
    const combo = [
        rec.Account, 
        rec.Date, 
        rec.Amount.toFixed(2), 
        rec.Description.replace(/\s+/g,' ').toLowerCase()
    ].join('|');
    if (setCombo.has(combo)) continue;

    append.push([
      rec.TxnId, rec.Date, rec.Account, rec.Amount, rec.Type,
      rec.Ref, rec.Description, rec.LinkedBillId, rec.LinkedAt, rec.Notes
    ]);
    // Add to sets to prevent duplicates *within the same file*
    setTxnId.add(rec.TxnId); 
    setCombo.add(combo);
    appended++;
  }

  if (appended > 0){
    sh.getRange(sh.getLastRow()+1, 1, append.length, append[0].length).setValues(append);
  }

  ui.alert(
    [
      `ไฟล์: ${file.getName()} (delimiter: "${delim}")`,
      `แถวข้อมูลที่อ่าน: ${dataRows.length}`,
      `ตีความเป็น "เงินเข้า" ได้: ${creditable} แถว`,
      `เพิ่มใหม่: ${appended} แถว`,
      ``,
      `แมปคอลัมน์ → Date:${map.idxDate}  Credit:${map.idxCredit}  Debit:${map.idxDebit}  Amount:${map.idxAmount}  Type:${map.idxType}  Ref:${map.idxRef}  Desc:${map.idxDesc}`
    ].join('\n')
  );
}
