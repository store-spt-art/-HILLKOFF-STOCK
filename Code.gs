/**
 * Dead Stock Ledger — Apps Script backend
 * ------------------------------------------------------------
 * วิธีติดตั้ง:
 * 1. เปิดชีต "สินค้า Dead Stock 2569 ต้องผลักดัน" ที่ต้องการเชื่อมต่อ
 * 2. เมนู Extensions > Apps Script
 * 3. ลบโค้ดตัวอย่างเดิม แล้ววางไฟล์นี้ทั้งหมด
 * 4. แก้ค่า SHEET_NAME ด้านล่างให้ตรงกับชื่อแท็บจริงในชีตของคุณ
 * 5. กด Deploy > New deployment > เลือกประเภท "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone (หรือ Anyone within [organization] ถ้าต้องการจำกัด)
 * 6. คัดลอกลิงก์ Web app ที่ได้ (ลงท้ายด้วย /exec) ไปใส่ในตัวแปร API_URL ของไฟล์ index.html
 *
 * ⚠️ สำคัญที่สุด: ทุกครั้งที่แก้โค้ดไฟล์นี้ ต้องกด
 *    Deploy > Manage deployments > (ไอคอนดินสอ) > Version: New version > Deploy
 *    ใหม่เสมอ ไม่งั้นลิงก์ /exec เดิมจะยังรันโค้ดเวอร์ชันเก่าอยู่ (สาเหตุที่พบบ่อยที่สุด
 *    ที่ทำให้ "แก้โค้ดแล้วแต่เว็บยังพังเหมือนเดิม")
 * ------------------------------------------------------------
 */

// ⚠️ ชื่อแท็บข้อมูลหลักจริงในสเปรดชีตนี้คือ "ENGSPTอะไหล่ไหม่"
const SHEET_NAME = 'ENGSPTอะไหล่ไหม่';

// แถวที่ข้อมูลเริ่มต้น (แถว 1-2 เป็นหัวตาราง, ข้อมูลเริ่มแถว 3)
const DATA_START_ROW = 3;

// ค่าตำแหน่งคอลัมน์เริ่มต้น (fallback) — ใช้เฉพาะกรณีหาหัวตารางไม่เจอเท่านั้น
// อ้างอิงจากไฟล์ต้นฉบับ: A=รหัสสินค้า ... U=ราคาขาย
const COLS_FALLBACK = {
  SKU: 1,             // A
  NAME: 2,             // B
  UNIT: 3,             // C
  STATUS: 4,            // D สถานะเฝ้าระวัง
  INITIAL_QTY: 5,        // E จำนวนคงค้าง
  MONTH_START: 6,        // F = ม.ค.
  MONTH_END: 17,         // Q = ธ.ค.
  REMAIN: 18,           // R คงเหลือ
  CHANNEL_FB: 19,        // S ช่องทางการลงขาย 1
  CHANNEL_SHOPEE: 20,    // T ช่องทางการลงขาย 2
  PRICE: 21             // U ราคาขาย
};

const MONTH_LABELS_TH = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];

const MOVEMENT_LOG_SHEET = 'MovementLog';
const LISTING_LOG_SHEET = 'ListingLog';

/* ------------------------------------------------------------------ */
/* Entry points                                                        */
/* ------------------------------------------------------------------ */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getData';
  let result;
  try {
    if (action === 'getData') {
      result = getData();
    } else {
      result = { success: false, error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return jsonOut(result);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'addMovement') {
      result = addMovement(body);
    } else if (body.action === 'addListing') {
      result = addListing(body);
    } else {
      result = { success: false, error: 'unknown action: ' + body.action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Sheet helpers                                                       */
/* ------------------------------------------------------------------ */

function getMainSheet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('ไม่พบแท็บชื่อ "' + SHEET_NAME + '" กรุณาแก้ค่า SHEET_NAME ในโค้ด');
  return sh;
}

function getOrCreateLogSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === MOVEMENT_LOG_SHEET) {
      sh.appendRow(['เวลาบันทึก', 'รหัสสินค้า', 'ชื่อสินค้า', 'เดือน', 'จำนวน', 'หมายเหตุ']);
    } else if (name === LISTING_LOG_SHEET) {
      sh.appendRow(['เวลาบันทึก', 'รหัสสินค้า', 'ชื่อสินค้า', 'ช่องทาง', 'วันที่ลงขาย', 'หมายเหตุ']);
    }
  }
  return sh;
}

/**
 * หาตำแหน่งคอลัมน์จริงในชีตหลัก โดย "อ่านจากชื่อหัวตาราง" (แถว 1 และแถว 2)
 * แทนการอ้างตำแหน่งคอลัมน์ตายตัว เพื่อไม่ให้พังถ้ามีการแทรก/ลบ/สลับคอลัมน์ในชีตจริง
 * ถ้าหาหัวตารางไม่เจอ จะ fallback ไปใช้ตำแหน่งเดิมใน COLS_FALLBACK
 */
function resolveMainColumns(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), COLS_FALLBACK.PRICE);
  const header1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const header2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

  function col(candidates, fallbackKey) {
    const idx = findColByHeader(header1, candidates);
    return idx !== -1 ? idx + 1 : COLS_FALLBACK[fallbackKey];
  }

  const C = {
    SKU: col(['รหัสสินค้า'], 'SKU'),
    NAME: col(['ชื่อสินค้า'], 'NAME'),
    UNIT: col(['หน่วยนับ'], 'UNIT'),
    STATUS: col(['สถานะ'], 'STATUS'),
    INITIAL_QTY: col(['จำนวนคงค้าง'], 'INITIAL_QTY'),
    REMAIN: col(['คงเหลือ'], 'REMAIN'),
    CHANNEL_FB: col(['ช่องทางการลงขาย 1'], 'CHANNEL_FB'),
    CHANNEL_SHOPEE: col(['ช่องทางการลงขาย 2'], 'CHANNEL_SHOPEE'),
    PRICE: col(['ราคาขาย'], 'PRICE')
  };

  // เดือน: หาจากแถวหัวที่ 2 (ม.ค., ก.พ., ... ธ.ค.) เพราะแถว 1 เป็นหัวข้อรวม (merge cell)
  const monthCols = {};
  MONTH_LABELS_TH.forEach((label, i) => {
    const idx = header2.findIndex(h => h === label);
    monthCols[i + 1] = idx !== -1 ? idx + 1 : (COLS_FALLBACK.MONTH_START + i);
  });
  C.MONTH_COLS = monthCols;
  C.MAX_COL = Math.max(
    C.SKU, C.NAME, C.UNIT, C.STATUS, C.INITIAL_QTY, C.REMAIN,
    C.CHANNEL_FB, C.CHANNEL_SHOPEE, C.PRICE,
    ...Object.values(monthCols)
  );
  return C;
}

function findRowBySku(sheet, sku, skuCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return -1;
  const values = sheet.getRange(DATA_START_ROW, skuCol, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(sku).trim()) return DATA_START_ROW + i;
  }
  return -1;
}

/* ------------------------------------------------------------------ */
/* Read                                                                 */
/* ------------------------------------------------------------------ */

function getData() {
  const sheet = getMainSheet();
  const C = resolveMainColumns(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    return { success: true, products: [], movementLog: [], listingLog: [] };
  }
  const numRows = lastRow - DATA_START_ROW + 1;
  const values = sheet.getRange(DATA_START_ROW, 1, numRows, C.MAX_COL).getValues();

  const products = values
    .filter(r => r[C.SKU - 1])
    .map(r => {
      const months = {};
      for (let m = 1; m <= 12; m++) {
        const raw = r[C.MONTH_COLS[m] - 1];
        months[m] = raw === '' || raw === null ? 0 : Number(raw) || 0;
      }
      // "คงเหลือ" อ้างอิงจากคอลัมน์ที่ AppSheet ดูแลโดยตรง (ไม่ได้คำนวณจากยอดเดือนในเว็บ)
      const remainRaw = r[C.REMAIN - 1];
      const remaining = (remainRaw === '-' || remainRaw === '' || remainRaw === null)
        ? 0
        : (Number(remainRaw) || 0);

      return {
        sku: r[C.SKU - 1],
        name: r[C.NAME - 1],
        unit: r[C.UNIT - 1],
        status: r[C.STATUS - 1],
        initialQty: Number(r[C.INITIAL_QTY - 1]) || 0,
        months: months,
        remaining: remaining,
        // "ช่องทางการลงขาย 1" = ลงขาย Facebook, เขียนว่า "ลงขาย FB DD-MM-YY"
        listedFB: r[C.CHANNEL_FB - 1] || '',
        // "ช่องทางการลงขาย 2" = ลงขาย Shopee, เขียนว่า "ลงขาย shopee DD-MM-YY"
        listedShopee: r[C.CHANNEL_SHOPEE - 1] || '',
        price: Number(r[C.PRICE - 1]) || 0
      };
    });

  return {
    success: true,
    products: products,
    // "ประวัติการเคลื่อนไหว" อ้างอิงคอลัมน์เดือน (ม.ค.–ธ.ค.) ในชีตหลักโดยตรง ต่อ 1 สินค้า
    movementLog: buildMovementHistoryFromMain(products),
    // "ประวัติการลงขาย" อ้างอิงคอลัมน์ "ช่องทางการลงขาย 1" (FB) และ "ช่องทางการลงขาย 2" (Shopee) ในชีตหลักโดยตรง
    listingLog: buildListingHistoryFromMain(products)
  };
}

/**
 * สร้างประวัติการเคลื่อนไหว จากคอลัมน์เดือน (ม.ค.–ธ.ค.) ของแต่ละสินค้าในชีตหลักโดยตรง
 * (ไม่มีเวลาบันทึกที่แท้จริงเพราะชีตหลักเก็บเป็นยอดรวมต่อเดือน ไม่ใช่ทีละรายการ)
 */
function buildMovementHistoryFromMain(products) {
  const rows = [];
  products.forEach(p => {
    for (let m = 1; m <= 12; m++) {
      const qty = Number(p.months[m] || 0);
      if (qty !== 0) {
        rows.push({ ts: '', sku: p.sku, name: p.name, month: m, qty: qty, note: '' });
      }
    }
  });
  rows.sort((a, b) => b.month - a.month); // เดือนล่าสุดขึ้นก่อน
  return rows;
}

/**
 * สร้างประวัติการลงขาย จากคอลัมน์ "ช่องทางการลงขาย 1" (FB) และ "ช่องทางการลงขาย 2" (Shopee)
 * ของแต่ละสินค้าในชีตหลักโดยตรง ข้อความในคอลัมน์มีรูปแบบ "ลงขาย FB DD-MM-YY" / "ลงขาย shopee DD-MM-YY"
 */
function buildListingHistoryFromMain(products) {
  const rows = [];
  products.forEach(p => {
    if (p.listedFB) rows.push(parseListingEntry(p, 'FB', p.listedFB));
    if (p.listedShopee) rows.push(parseListingEntry(p, 'SHOPEE', p.listedShopee));
  });
  rows.sort((a, b) => (b._sortKey || 0) - (a._sortKey || 0)); // วันที่ล่าสุดขึ้นก่อน
  rows.forEach(r => { delete r._sortKey; });
  return rows;
}

// แปลงข้อความ เช่น "ลงขาย FB 03-03-69" ให้เป็น entry ของประวัติการลงขาย พร้อม key สำหรับเรียงลำดับ
function parseListingEntry(p, channel, text) {
  const m = String(text).match(/(\d{1,2})-(\d{1,2})-(\d{2,4})/);
  let dateStr = String(text).replace(/^ลงขาย\s*(FB|shopee)\s*/i, '').trim();
  let sortKey = 0;
  if (m) {
    const dd = m[1].padStart(2, '0');
    const mm = m[2].padStart(2, '0');
    const yy = m[3];
    dateStr = dd + '-' + mm + '-' + yy;
    const yyNum = Number(yy.length >= 2 ? yy.slice(-2) : yy);
    sortKey = yyNum * 10000 + Number(mm) * 100 + Number(dd); // พ.ศ. 2 หลัก ใช้เรียงลำดับพอเพียง
  }
  return { ts: '', sku: p.sku, name: p.name, channel: channel, date: dateStr, note: '', _sortKey: sortKey };
}

// หา index (0-based) ของคอลัมน์จากชื่อหัวตาราง: จับคู่แบบตรงเป๊ะก่อน แล้วค่อย fallback แบบมีคำนั้นอยู่ในชื่อ
function findColByHeader(headers, candidates) {
  for (const kw of candidates) {
    const idx = headers.findIndex(h => h === kw);
    if (idx !== -1) return idx;
  }
  for (const kw of candidates) {
    const idx = headers.findIndex(h => h.indexOf(kw) !== -1);
    if (idx !== -1) return idx;
  }
  return -1;
}

/* ------------------------------------------------------------------ */
/* Write (ปัจจุบันหน้าเว็บไม่ได้เรียกใช้ 2 ฟังก์ชันนี้แล้ว เพราะการบันทึก   */
/* ข้อมูลย้ายไปทำผ่าน AppSheet ทั้งหมด — คงไว้เผื่อใช้งานในอนาคต)         */
/* ------------------------------------------------------------------ */

function addMovement(body) {
  const sku = body.sku;
  const month = Number(body.month); // 1-12
  const qty = Number(body.qty);
  const note = body.note || '';

  if (!sku) return { success: false, error: 'ไม่ได้ระบุรหัสสินค้า' };
  if (!month || month < 1 || month > 12) return { success: false, error: 'เดือนไม่ถูกต้อง' };
  if (isNaN(qty)) return { success: false, error: 'จำนวนไม่ถูกต้อง' };

  const sheet = getMainSheet();
  const C = resolveMainColumns(sheet);
  const row = findRowBySku(sheet, sku, C.SKU);
  if (row === -1) return { success: false, error: 'ไม่พบรหัสสินค้า: ' + sku };

  const col = C.MONTH_COLS[month];
  const current = Number(sheet.getRange(row, col).getValue()) || 0;
  sheet.getRange(row, col).setValue(current + qty);

  // recompute "คงเหลือ" = คงค้างเริ่มต้น - ผลรวมที่โอนออกทุกเดือน
  const initialQty = Number(sheet.getRange(row, C.INITIAL_QTY).getValue()) || 0;
  let sumMonths = 0;
  for (let m = 1; m <= 12; m++) {
    sumMonths += Number(sheet.getRange(row, C.MONTH_COLS[m]).getValue()) || 0;
  }
  const remaining = initialQty - sumMonths;
  sheet.getRange(row, C.REMAIN).setValue(remaining <= 0 ? '-' : remaining);

  const name = sheet.getRange(row, C.NAME).getValue();
  const log = getOrCreateLogSheet(MOVEMENT_LOG_SHEET);
  log.appendRow([new Date(), sku, name, month, qty, note]);

  return { success: true, remaining: remaining };
}

function addListing(body) {
  const sku = body.sku;
  const channel = body.channel; // 'FB' or 'SHOPEE'
  const dateInput = body.date;  // 'YYYY-MM-DD'
  const note = body.note || '';

  if (!sku) return { success: false, error: 'ไม่ได้ระบุรหัสสินค้า' };
  if (channel !== 'FB' && channel !== 'SHOPEE') return { success: false, error: 'ช่องทางไม่ถูกต้อง' };
  if (!dateInput) return { success: false, error: 'ไม่ได้ระบุวันที่' };

  const sheet = getMainSheet();
  const C = resolveMainColumns(sheet);
  const row = findRowBySku(sheet, sku, C.SKU);
  if (row === -1) return { success: false, error: 'ไม่พบรหัสสินค้า: ' + sku };

  const dateStr = toBuddhistDateStr(dateInput);
  const label = channel === 'FB' ? ('ลงขาย FB ' + dateStr) : ('ลงขาย shopee ' + dateStr);
  const col = channel === 'FB' ? C.CHANNEL_FB : C.CHANNEL_SHOPEE;
  sheet.getRange(row, col).setValue(label);

  const name = sheet.getRange(row, C.NAME).getValue();
  const log = getOrCreateLogSheet(LISTING_LOG_SHEET);
  log.appendRow([new Date(), sku, name, channel, dateStr, note]);

  return { success: true, label: label };
}

/* ------------------------------------------------------------------ */
/* Utils                                                                */
/* ------------------------------------------------------------------ */

// แปลงวันที่ 'YYYY-MM-DD' (ค.ศ.) เป็นรูปแบบ 'DD-MM-YY' แบบ พ.ศ. สองหลัก
// ให้ตรงกับรูปแบบเดิมในชีต เช่น "03-03-69" (2569)
function toBuddhistDateStr(isoDateStr) {
  const parts = String(isoDateStr).split('-'); // [YYYY, MM, DD]
  const y = Number(parts[0]);
  const m = parts[1];
  const d = parts[2];
  const buddhistYearShort = String((y + 543) % 100).padStart(2, '0');
  return d + '-' + m + '-' + buddhistYearShort;
}
