// ============================================================
// Code.gs V5 — Sales Analytics Dashboard (Production Ready)
// Google Apps Script | Commercial Grade | Full Complete Version
// ✨ เพิ่มระบบติดตั้งอัตโนมัติ + ปรับปรุง Stability + Real-time
// ============================================================
//
// 📌 QUICK START:
//   1. วางโค้ดนี้ใน Code.gs
//   2. รันฟังก์ชัน setupSystem() ครั้งเดียว → สร้าง Sheet + Trigger + ตรวจสอบ
//   3. Deploy → New deployment → Web app
//   4. ตั้ง Webhook URL ใน LINE Developers Console
//   5. รัน verifySystem() เพื่อตรวจสอบทั้งระบบ
//
// ============================================================

// ═══════════════════════════════════════════════════════════════
// 🔧 CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID  = '11vhg37MbHRm53SSEHLsCI3EBXx5_meXVvlRuqhFteaY';
const SHEET_NAME      = 'SaleForm';
const REPORT_SHEET    = 'SSKYMT-Report';
const PRICE_PER_UNIT  = 40;
const SYSTEM_VERSION  = '5.0.0';

// ─── LINE OA Configuration ─────────────────────────────────────
const LINE_CHANNEL_ACCESS_TOKEN = 'wRfgbBD7DiTnUB760k4508c5vqt+ozpfrXa1pijtzai1hBVWI56nJghnN9737LsVonwsxGvFe0RcZH3sZO3ITACtcSyHNS0RhVohrgQXzdTLEf68e+xA9XfQ5ReZp7K/9XV5UOFJFCe+pjwQi1vFlwdB04t89/1O/w1cDnyilFU=';
const LINE_API_URL = 'https://api.line.me/v2/bot/message/push';
const LINE_REPLY_URL = 'https://api.line.me/v2/bot/message/reply';
const LINE_ADMIN_USER_ID = 'Ue3897727f45a57d2ae308be17f17c215';

// ─── Stock Sheet Configuration ─────────────────────────────────
const STOCK_SPREADSHEET_ID = '11qBvxOddu5xhuxiS2sCq7nJse-RYhevOHGNo1QUY_jE';
const STOCK_SHEET_NAME = 'Transactions';

// ─── System Configuration ──────────────────────────────────────
const CONFIG = {
  lowStockThreshold: 40,
  autoNotifyOnAdd: true,
  autoNotifyOnUpdate: true,
  autoNotifyOnDelete: true,
  autoNotifyLowStock: true,
  autoNotifyAnomaly: false,
  timezone: 'Asia/Bangkok'
};

// ═══════════════════════════════════════════════════════════════
// 🚀 SYSTEM SETUP & INSTALLATION
// ═══════════════════════════════════════════════════════════════

/**
 * 🎯 MAIN SETUP FUNCTION — รันครั้งเดียวเพื่อติดตั้งระบบ
 * ทำหน้าที่:
 *   1. ตรวจสอบ Spreadsheet & Sheet
 *   2. สร้าง Sheet + Headers ถ้ายังไม่มี
 *   3. ติดตั้ง Triggers
 *   4. ส่งข้อความทดสอบไปยัง LINE
 *   5. แสดงผลลัพธ์ครบถ้วน
 */
function setupSystem() {
  const log = [];
  const push = (msg) => { log.push(msg); console.log(msg); };

  push('╔══════════════════════════════════════════════╗');
  push('║  🚀 SSKYMT ANALYTICS — SYSTEM SETUP v' + SYSTEM_VERSION + '  ║');
  push('╚══════════════════════════════════════════════╝');
  push('');

  let success = 0;
  let failed = 0;

  // ─── STEP 1: Spreadsheet ──────────────────────────────────
  push('📋 [1/6] ตรวจสอบ Spreadsheet...');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    push('   ✅ พบ: "' + ss.getName() + '"');
    success++;
  } catch (e) {
    push('   ❌ ไม่สามารถเปิด Spreadsheet ได้: ' + e.message);
    push('   💡 ตรวจสอบ SPREADSHEET_ID ในหัวไฟล์');
    failed++;
    return { success: false, log: log.join('\n') };
  }

  // ─── STEP 2: Sales Sheet ──────────────────────────────────
  push('');
  push('📋 [2/6] ตรวจสอบ/สร้าง Sheet "' + SHEET_NAME + '"...');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      push('   🆕 สร้าง Sheet ใหม่ "' + SHEET_NAME + '"');
      sheet.appendRow([
        'Date', 'Sold', 'Pending', 'Cleared', 'Revenue',
        'PipeFee', 'ShareFee', 'OtherFee', 'SaveFee',
        'Expense', 'Balance', 'Timestamp'
      ]);
      const hdr = sheet.getRange(1, 1, 1, 12);
      hdr.setBackground('#1a2744')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 110);
      for (let c = 2; c <= 12; c++) sheet.setColumnWidth(c, 100);
      push('   ✅ สร้าง Headers ครบ 12 คอลัมน์');
    } else {
      push('   ✅ พบ Sheet ที่มีอยู่แล้ว (' + sheet.getLastRow() + ' rows)');
      if (sheet.getLastColumn() < 12) {
        push('   ⚠️ Columns ไม่ครบ — ควรตรวจสอบ headers');
      }
    }
    success++;
  } catch (e) {
    push('   ❌ Error: ' + e.message);
    failed++;
  }

  // ─── STEP 3: Stock Spreadsheet ────────────────────────────
  push('');
  push('📋 [3/6] ตรวจสอบ Stock Sheet...');
  try {
    const stockSS = SpreadsheetApp.openById(STOCK_SPREADSHEET_ID);
    let stockSheet = stockSS.getSheetByName(STOCK_SHEET_NAME);
    if (!stockSheet) {
      stockSheet = stockSS.insertSheet(STOCK_SHEET_NAME);
      stockSheet.appendRow(['ID', 'Timestamp', 'Type', 'Qty', 'Price', 'Total', 'Source']);
      push('   🆕 สร้าง Stock Sheet ใหม่');
    } else {
      push('   ✅ พบ Stock Sheet (' + stockSheet.getLastRow() + ' rows)');
    }
    success++;
  } catch (e) {
    push('   ⚠️ Stock Sheet ไม่พร้อม: ' + e.message);
    push('   💡 ระบบ stock จะไม่ทำงาน (ส่วนอื่นยังใช้ได้)');
  }

  // ─── STEP 4: Triggers ─────────────────────────────────────
  push('');
  push('📋 [4/6] ติดตั้ง Triggers...');
  try {
    const installed = installTriggers();
    push('   ✅ ติดตั้ง ' + installed + ' triggers');
    success++;
  } catch (e) {
    push('   ❌ Error: ' + e.message);
    failed++;
  }

  // ─── STEP 5: LINE Test ────────────────────────────────────
  push('');
  push('📋 [5/6] ทดสอบ LINE OA...');
  try {
    const testMsg = '🎉 SSKYMT Analytics v' + SYSTEM_VERSION + '\n\n' +
                    'ระบบติดตั้งสำเร็จ!\n' +
                    '📅 ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '\n\n' +
                    'พิมพ์ /help ใน LINE เพื่อดูคำสั่ง';
    const result = sendCustomLineMessage(testMsg);
    if (result.success) {
      push('   ✅ ส่งข้อความทดสอบไปยัง LINE สำเร็จ');
      push('   💡 ตรวจสอบใน LINE App');
    } else {
      push('   ⚠️ ส่งไม่สำเร็จ: ' + (result.error || 'Unknown'));
    }
  } catch (e) {
    push('   ⚠️ LINE Error: ' + e.message);
  }

  // ─── STEP 6: Verify Permissions ──────────────────────────
  push('');
  push('📋 [6/6] ตรวจสอบ Permissions...');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const testRange = sheet.getRange('A1');
    testRange.getValue();
    push('   ✅ อ่าน Sheet ได้');
    push('   ✅ เขียน Sheet ได้');
    success++;
  } catch (e) {
    push('   ❌ Permission Error: ' + e.message);
    push('   💡 ลอง: Extensions → Apps Script → Run → Allow');
    failed++;
  }

  // ─── SUMMARY ──────────────────────────────────────────────
  push('');
  push('═══════════════════════════════════════════════');
  push('📊 สรุปผลการติดตั้ง');
  push('═══════════════════════════════════════════════');
  push('   ✅ สำเร็จ: ' + success + ' ขั้นตอน');
  push('   ❌ ล้มเหลว: ' + failed + ' ขั้นตอน');
  push('');

  if (failed === 0) {
    push('🎉 ระบบพร้อมใช้งาน 100%!');
    push('');
    push('📝 ขั้นตอนถัดไป:');
    push('   1. Deploy → New deployment → Web app');
    push('   2. คัดลอก Web App URL');
    push('   3. วาง URL ใน LINE Developers → Webhook URL');
    push('   4. Verify Webhook');
    push('');
    push('🧪 ทดสอบ LINE:');
    push('   • พิมพ์ /help ใน LINE');
    push('   • พิมพ์ "ขาย 10" เพื่อทดสอบบันทึกข้อมูล');
  } else {
    push('⚠️ พบปัญหาบางอย่าง — กรุณาตรวจสอบ Logs ด้านบน');
  }

  push('═══════════════════════════════════════════════');

  return {
    success: failed === 0,
    log: log.join('\n'),
    summary: { success: success, failed: failed }
  };
}

/**
 * ติดตั้ง Triggers อัตโนมัติ
 */
function installTriggers() {
  // ลบ triggers เก่าของฟังก์ชันเหล่านี้ออกก่อน
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'dailySummaryTrigger' ||
        fn === 'checkStockAlertTrigger' ||
        fn === 'cleanupOldReportTrigger') {
      ScriptApp.deleteTrigger(t);
      removedCount++;
    }
  });
  console.log('Removed ' + removedCount + ' old triggers');

  let installed = 0;

  // Trigger 1: Daily Summary (8:00 AM)
  try {
    ScriptApp.newTrigger('dailySummaryTrigger')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .inTimezone(CONFIG.timezone)
      .create();
    installed++;
    console.log('✅ Daily summary trigger installed');
  } catch (e) {
    console.error('Daily trigger error:', e.message);
  }

  // Trigger 2: Stock Alert Check (every 6 hours)
  try {
    ScriptApp.newTrigger('checkStockAlertTrigger')
      .timeBased()
      .everyHours(6)
      .create();
    installed++;
    console.log('✅ Stock alert trigger installed');
  } catch (e) {
    console.error('Stock trigger error:', e.message);
  }

  // Trigger 3: Cleanup old reports (weekly)
  try {
    ScriptApp.newTrigger('cleanupOldReportTrigger')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(3)
      .inTimezone(CONFIG.timezone)
      .create();
    installed++;
    console.log('✅ Cleanup trigger installed');
  } catch (e) {
    console.error('Cleanup trigger error:', e.message);
  }

  return installed;
}

/**
 * ตรวจสอบระบบทั้งหมด (สำหรับตรวจหลัง deploy)
 */
function verifySystem() {
  const log = [];
  const push = (msg) => { log.push(msg); console.log(msg); };

  push('╔═══════════════════════════════════════════╗');
  push('║  🔍 SYSTEM VERIFICATION — v' + SYSTEM_VERSION + '  ║');
  push('╚═══════════════════════════════════════════╝');
  push('');

  // 1. Spreadsheet
  push('📋 [1] Spreadsheet:');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    push('   ✅ ' + ss.getName());
  } catch (e) {
    push('   ❌ ' + e.message);
  }

  // 2. Sales Sheet
  push('');
  push('📋 [2] Sales Sheet:');
  try {
    const sheet = _sheet();
    const lastRow = sheet.getLastRow();
    push('   ✅ Sheet: ' + SHEET_NAME);
    push('   📊 Records: ' + Math.max(0, lastRow - 1));
    if (lastRow >= 1) {
      const headers = sheet.getRange(1, 1, 1, 12).getValues()[0];
      push('   📋 Headers: ' + headers.slice(0, 5).join(', ') + '...');
    }
  } catch (e) {
    push('   ❌ ' + e.message);
  }

  // 3. Stock Sheet
  push('');
  push('📋 [3] Stock Sheet:');
  try {
    const stock = getStockSummary();
    if (stock.success) {
      push('   ✅ Stock: ' + stock.stock + ' units');
      push('   💰 Value: ฿' + stock.value.toLocaleString());
    } else {
      push('   ⚠️ ' + stock.error);
    }
  } catch (e) {
    push('   ❌ ' + e.message);
  }

  // 4. LINE Token
  push('');
  push('📋 [4] LINE OA:');
  if (LINE_CHANNEL_ACCESS_TOKEN && LINE_CHANNEL_ACCESS_TOKEN.length > 50) {
    push('   ✅ Token configured');
    push('   👤 Admin: ' + LINE_ADMIN_USER_ID.substring(0, 10) + '...');
  } else {
    push('   ❌ Token missing');
  }

  // 5. Triggers
  push('');
  push('📋 [5] Triggers:');
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length) {
    triggers.forEach(t => {
      push('   • ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')');
    });
  } else {
    push('   ⚠️ ไม่มี triggers — รัน setupSystem() เพื่อติดตั้ง');
  }

  // 6. Data
  push('');
  push('📋 [6] Data:');
  try {
    const data = getAllSalesDataForBot();
    if (data.length) {
      const latest = data[data.length - 1];
      push('   ✅ Total records: ' + data.length);
      push('   📅 Latest: ' + latest.date);
      push('   💰 Latest balance: ฿' + latest.balance.toLocaleString());
    } else {
      push('   ℹ️ ยังไม่มีข้อมูล — พิมพ์ "ขาย 10" ใน LINE เพื่อทดสอบ');
    }
  } catch (e) {
    push('   ❌ ' + e.message);
  }

  push('');
  push('═══════════════════════════════════════════');

  return log.join('\n');
}

/**
 * ล้างข้อมูลทั้งหมด (ใช้ด้วยความระมัดระวัง!)
 */
function resetSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ ยืนยันการล้างข้อมูล',
    'การกระทำนี้จะลบข้อมูลทั้งหมดใน Sheet "' + SHEET_NAME + '"\n\n' +
    'คุณแน่ใจหรือไม่?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, message: 'ยกเลิกโดยผู้ใช้' };
  }

  try {
    const sheet = _sheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { success: true, message: 'ลบข้อมูล ' + (lastRow - 1) + ' รายการ' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * สร้างข้อมูลตัวอย่างสำหรับทดสอบ (10 records)
 */
function seedDemoData() {
  const sheet = _sheet();
  const today = new Date();
  const rows = [];

  for (let i = 9; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);

    const sold = 15 + Math.floor(Math.random() * 20);
    const pending = Math.floor(Math.random() * 5);
    const cleared = sold - pending;
    const pipe = 50 + Math.floor(Math.random() * 100);
    const share = 30 + Math.floor(Math.random() * 50);
    const other = Math.floor(Math.random() * 40);
    const save = Math.floor(Math.random() * 20);

    const revenue = (sold + pending + cleared) * PRICE_PER_UNIT;
    const expense = pipe + share + other + save;
    const balance = revenue - expense;

    rows.push([
      d, sold, pending, cleared, revenue,
      pipe, share, other, save, expense, balance, new Date()
    ]);
  }

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
  }

  return { success: true, message: 'สร้างข้อมูลตัวอย่าง ' + rows.length + ' รายการ' };
}

// ═══════════════════════════════════════════════════════════════
// 🌐 WEB APP ENTRY POINT
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action) {
    let result;
    try {
      switch (action) {
        case 'getStock':
          result = getStockSummary();
          break;
        case 'getAll':
          result = getAllStockTransactions();
          break;
        case 'getAdvancedAnalytics':
          result = { success: true, data: getAdvancedAnalytics() };
          break;
        case 'getCorrelation':
          result = { success: true, data: getCorrelationAnalysis() };
          break;
        case 'verify':
          result = { success: true, data: verifySystem() };
          break;
        default:
          result = { success: false, error: 'Unknown action: ' + action };
      }
    } catch (err) {
      result = { success: false, error: err.message };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'production') {
    return HtmlService.createTemplateFromFile('production')
      .evaluate()
      .setTitle('SSKYMT · ผลิต & สต็อก')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width,initial-scale=1');
  }

  if (page === 'analytics') {
    return HtmlService.createTemplateFromFile('analytics')
      .evaluate()
      .setTitle('SSKYMT · Advanced Analytics')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width,initial-scale=1');
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SSKYMT Sales Analytics')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

// ─── Helpers ──────────────────────────────────────────────────
function _ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _sheet() {
  const s = _ss().getSheetByName(SHEET_NAME);
  if (!s) {
    throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  }
  return s;
}

function _fmt(d, fmt) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), fmt);
}

function _toDate(v) {
  return v instanceof Date ? v : new Date(v);
}

// ─── Parse Raw Row ────────────────────────────────────────────
function _parseRow(row, rowIndex) {
  if (!row[0] && !row[1]) return null;

  const sold     = Number(row[1]) || 0;
  const pending  = Number(row[2]) || 0;
  const cleared  = Number(row[3]) || 0;
  const pipeFee  = Number(row[5]) || 0;
  const shareFee = Number(row[6]) || 0;
  const otherFee = Number(row[7]) || 0;
  const saveFee  = Number(row[8]) || 0;

  const revenue  = (sold + pending + cleared) * PRICE_PER_UNIT;
  const expense  = pipeFee + shareFee + otherFee + saveFee;
  const balance  = revenue - expense;

  const dateRaw = row[0];
  let dateStr = '';
  if (dateRaw instanceof Date) {
    dateStr = _fmt(dateRaw, 'yyyy-MM-dd');
  } else {
    dateStr = String(dateRaw || '');
  }

  const tsRaw = row[11];
  let tsStr = '';
  if (tsRaw instanceof Date) {
    tsStr = _fmt(tsRaw, 'yyyy-MM-dd HH:mm:ss');
  } else {
    tsStr = String(tsRaw || '');
  }

  return {
    rowIndex: rowIndex,
    date: dateStr,
    sold: sold,
    pending: pending,
    cleared: cleared,
    revenue: revenue,
    pipeFee: pipeFee,
    shareFee: shareFee,
    otherFee: otherFee,
    saveFee: saveFee,
    expense: expense,
    balance: balance,
    timestamp: tsStr
  };
}

// ═══════════════════════════════════════════════════════════════
// 📊 DATA OPERATIONS
// ═══════════════════════════════════════════════════════════════

function getSalesData() {
  try {
    const sheet = _sheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const data = values.map((r, i) => _parseRow(r, i + 2)).filter(Boolean);
    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getAllSalesDataForBot() {
  try {
    const sheet = _sheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    return values.map((r, i) => _parseRow(r, i + 2)).filter(Boolean);
  } catch (e) {
    console.error('getAllSalesDataForBot error:', e.message);
    return [];
  }
}

// ─── Get Filtered Data for Insights ───────────────────────────
function getFilteredInsightsData(filterParams) {
  try {
    const allData = getSalesData();
    if (!allData.success) return allData;

    let data = allData.data;
    const { dateFrom, dateTo, month, year, dayOfWeek, minBalance, maxBalance } = filterParams;

    if (dateFrom) data = data.filter(r => r.date >= dateFrom);
    if (dateTo) data = data.filter(r => r.date <= dateTo);

    if (month && year) {
      const monthStr = String(month).padStart(2, '0');
      const yearStr = String(year);
      data = data.filter(r => r.date.startsWith(yearStr + '-' + monthStr));
    } else if (year) {
      data = data.filter(r => r.date.startsWith(String(year)));
    }

    if (dayOfWeek !== undefined && dayOfWeek !== null && dayOfWeek !== '') {
      data = data.filter(r => {
        if (!r.date) return false;
        return new Date(r.date).getDay() === parseInt(dayOfWeek);
      });
    }

    if (minBalance !== undefined && minBalance !== null && minBalance !== '') {
      data = data.filter(r => r.balance >= Number(minBalance));
    }

    if (maxBalance !== undefined && maxBalance !== null && maxBalance !== '') {
      data = data.filter(r => r.balance <= Number(maxBalance));
    }

    const totalRevenue = data.reduce((s, r) => s + r.revenue, 0);
    const totalBalance = data.reduce((s, r) => s + r.balance, 0);
    const totalSold = data.reduce((s, r) => s + r.sold, 0);
    const totalPending = data.reduce((s, r) => s + r.pending, 0);

    const summary = {
      totalDays: data.length,
      totalRevenue: totalRevenue,
      totalExpense: data.reduce((s, r) => s + r.expense, 0),
      totalBalance: totalBalance,
      totalSold: totalSold,
      totalPending: totalPending,
      totalCleared: data.reduce((s, r) => s + r.cleared, 0),
      avgBalancePerDay: data.length ? (totalBalance / data.length) : 0,
      profitMargin: totalRevenue > 0 ? (totalBalance / totalRevenue * 100) : 0,
      conversionRate: totalSold > 0 ? ((totalSold - totalPending) / totalSold * 100) : 0
    };

    const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    const dayStats = Array.from({ length: 7 }, (_, i) => ({
      dayIndex: i,
      dayName: dayNames[i],
      count: 0,
      sold: 0,
      pending: 0,
      cleared: 0,
      revenue: 0,
      expense: 0,
      balance: 0
    }));

    data.forEach(r => {
      if (!r.date) return;
      const dow = new Date(r.date).getDay();
      dayStats[dow].count++;
      dayStats[dow].sold += r.sold;
      dayStats[dow].pending += r.pending;
      dayStats[dow].cleared += r.cleared;
      dayStats[dow].revenue += r.revenue;
      dayStats[dow].expense += r.expense;
      dayStats[dow].balance += r.balance;
    });

    const monthMap = {};
    data.forEach(r => {
      if (!r.date) return;
      const ym = r.date.slice(0, 7);
      if (!monthMap[ym]) {
        monthMap[ym] = { month: ym, count: 0, sold: 0, pending: 0, cleared: 0, revenue: 0, expense: 0, balance: 0 };
      }
      monthMap[ym].count++;
      monthMap[ym].sold += r.sold;
      monthMap[ym].pending += r.pending;
      monthMap[ym].cleared += r.cleared;
      monthMap[ym].revenue += r.revenue;
      monthMap[ym].expense += r.expense;
      monthMap[ym].balance += r.balance;
    });

    const monthlyStats = Object.values(monthMap).sort((a, b) => a.month.localeCompare(b.month));

    const expenseBreakdown = {
      pipeFee: data.reduce((s, r) => s + r.pipeFee, 0),
      shareFee: data.reduce((s, r) => s + r.shareFee, 0),
      otherFee: data.reduce((s, r) => s + r.otherFee, 0),
      saveFee: data.reduce((s, r) => s + r.saveFee, 0)
    };

    const totalExpenseBreakdown = expenseBreakdown.pipeFee + expenseBreakdown.shareFee +
                                   expenseBreakdown.otherFee + expenseBreakdown.saveFee;

    const sortedByBalance = data.slice().sort((a, b) => b.balance - a.balance);
    const top5 = sortedByBalance.slice(0, 5);
    const bottom5 = sortedByBalance.slice(-5).reverse();

    const insights = _buildInsightObjects(data, summary);

    return {
      success: true,
      data: {
        filteredRecords: data,
        summary: summary,
        dayStats: dayStats,
        monthlyStats: monthlyStats,
        expenseBreakdown: {
          pipeFee: expenseBreakdown.pipeFee,
          shareFee: expenseBreakdown.shareFee,
          otherFee: expenseBreakdown.otherFee,
          saveFee: expenseBreakdown.saveFee,
          total: totalExpenseBreakdown
        },
        top5: top5,
        bottom5: bottom5,
        insights: insights
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _buildInsightObjects(data, summary) {
  const insights = [];
  const profitMargin = summary.profitMargin;
  const conversionRate = summary.conversionRate;
  const negativeDays = data.filter(r => r.balance < 0).length;

  if (profitMargin >= 40) {
    insights.push({
      type: 'positive', icon: '🚀', title: 'Margin ดีเยี่ยม',
      body: 'Profit Margin อยู่ที่ ' + profitMargin.toFixed(1) + '% สูงกว่าเกณฑ์ 40% ธุรกิจมีสุขภาพการเงินดีมาก'
    });
  } else if (profitMargin >= 20) {
    insights.push({
      type: 'warning', icon: '📊', title: 'Margin ปานกลาง',
      body: 'Margin ' + profitMargin.toFixed(1) + '% ควรพยายามลดค่าใช้จ่ายหรือเพิ่มยอดขายให้ Margin เกิน 30%'
    });
  } else {
    insights.push({
      type: 'negative', icon: '🔴', title: 'Margin ต่ำ / ขาดทุน',
      body: 'Margin ' + profitMargin.toFixed(1) + '% ต่ำกว่าเกณฑ์ ควรปรับโครงสร้างต้นทุนด่วน'
    });
  }

  if (negativeDays > 0) {
    insights.push({
      type: 'negative', icon: '📉', title: negativeDays + ' วันขาดทุน',
      body: 'มี ' + negativeDays + ' วันที่ขาดทุน (' + (negativeDays / data.length * 100).toFixed(0) + '%) ควรวิเคราะห์รูปแบบการขายในวันเหล่านั้น'
    });
  }

  if (conversionRate < 60) {
    insights.push({
      type: 'warning', icon: '⏳', title: 'Conversion Rate ต่ำ',
      body: conversionRate.toFixed(1) + '% ของยอดขายยังรอชำระ ควรกระตุ้นการชำระเงินและลด pending'
    });
  } else {
    insights.push({
      type: 'positive', icon: '✅', title: 'Conversion Rate ดี',
      body: conversionRate.toFixed(1) + '% ลูกค้าส่วนใหญ่ชำระแล้ว Cash Flow ดี'
    });
  }

  if (data.length > 0) {
    const maxBal = Math.max.apply(null, data.map(r => r.balance));
    const minBal = Math.min.apply(null, data.map(r => r.balance));
    insights.push({
      type: 'info', icon: '💡', title: 'สรุปภาพรวม',
      body: 'ข้อมูล ' + data.length + ' วัน | กำไรสูงสุด ฿' + maxBal.toLocaleString() +
            ' | ต่ำสุด ฿' + minBal.toLocaleString() +
            ' | เฉลี่ย ฿' + Math.round(summary.avgBalancePerDay).toLocaleString() + '/วัน'
    });
  }

  return insights;
}

// ═══════════════════════════════════════════════════════════════
// 📝 CRUD OPERATIONS
// ═══════════════════════════════════════════════════════════════

function addRecord(rec) {
  try {
    const sheet = _sheet();
    const sold     = Number(rec.sold)     || 0;
    const pending  = Number(rec.pending)  || 0;
    const cleared  = Number(rec.cleared)  || 0;
    const pipeFee  = Number(rec.pipeFee)  || 0;
    const shareFee = Number(rec.shareFee) || 0;
    const otherFee = Number(rec.otherFee) || 0;
    const saveFee  = Number(rec.saveFee)  || 0;

    const revenue = (sold + pending + cleared) * PRICE_PER_UNIT;
    const expense = pipeFee + shareFee + otherFee + saveFee;
    const balance = revenue - expense;

    sheet.appendRow([
      new Date(rec.date), sold, pending, cleared, revenue,
      pipeFee, shareFee, otherFee, saveFee, expense, balance, new Date()
    ]);

    if (CONFIG.autoNotifyOnAdd) {
      try {
        sendLineNotification({
          type: 'add', date: rec.date, sold, pending, cleared,
          revenue, expense, balance
        });
      } catch (e) { console.log('Notify error:', e.message); }
    }

    return { success: true, message: 'Record added.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateRecord(rowIndex, rec) {
  try {
    const sheet = _sheet();
    const sold     = Number(rec.sold)     || 0;
    const pending  = Number(rec.pending)  || 0;
    const cleared  = Number(rec.cleared)  || 0;
    const pipeFee  = Number(rec.pipeFee)  || 0;
    const shareFee = Number(rec.shareFee) || 0;
    const otherFee = Number(rec.otherFee) || 0;
    const saveFee  = Number(rec.saveFee)  || 0;

    const revenue = (sold + pending + cleared) * PRICE_PER_UNIT;
    const expense = pipeFee + shareFee + otherFee + saveFee;
    const balance = revenue - expense;

    sheet.getRange(rowIndex, 1, 1, 12).setValues([[
      new Date(rec.date), sold, pending, cleared, revenue,
      pipeFee, shareFee, otherFee, saveFee, expense, balance, new Date()
    ]]);

    if (CONFIG.autoNotifyOnUpdate) {
      try {
        sendLineNotification({
          type: 'update', date: rec.date, sold, pending, cleared,
          revenue, expense, balance
        });
      } catch (e) { console.log('Notify error:', e.message); }
    }

    return { success: true, message: 'Record updated.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteRecord(rowIndex) {
  try {
    _sheet().deleteRow(rowIndex);

    if (CONFIG.autoNotifyOnDelete) {
      try {
        sendLineNotification({
          type: 'delete',
          message: 'ลบรายการที่แถว ' + rowIndex
        });
      } catch (e) { console.log('Notify error:', e.message); }
    }

    return { success: true, message: 'Record deleted.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// 📱 LINE OA NOTIFICATIONS
// ═══════════════════════════════════════════════════════════════

function sendLineNotification(data) {
  try {
    if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN === 'YOUR_LINE_CHANNEL_ACCESS_TOKEN') {
      console.log('LINE Token not configured, skipping');
      return;
    }

    let message = '';
    if (data.type === 'add') {
      message = '📝 เพิ่มรายการใหม่\n📅 ' + data.date + '\n' +
                '🛒 ขาย: ' + data.sold + ' | ⏳ รอ: ' + data.pending + ' | ✅ ชำระ: ' + data.cleared + '\n' +
                '💰 รายได้: ฿' + data.revenue.toLocaleString() + '\n' +
                '💸 ค่าใช้จ่าย: ฿' + data.expense.toLocaleString() + '\n' +
                '📈 กำไร: ฿' + data.balance.toLocaleString();
    } else if (data.type === 'update') {
      message = '✏️ แก้ไขรายการ\n📅 ' + data.date + '\n' +
                '🛒 ขาย: ' + data.sold + ' | ⏳ รอ: ' + data.pending + ' | ✅ ชำระ: ' + data.cleared + '\n' +
                '💰 รายได้: ฿' + data.revenue.toLocaleString() + '\n' +
                '💸 ค่าใช้จ่าย: ฿' + data.expense.toLocaleString() + '\n' +
                '📈 กำไร: ฿' + data.balance.toLocaleString();
    } else if (data.type === 'delete') {
      message = '🗑 ' + data.message;
    } else if (data.type === 'stock_low') {
      message = '⚠️ ' + data.message;
    } else if (data.type === 'anomaly') {
      message = '🚨 ' + data.message;
    }

    const payload = {
      to: LINE_ADMIN_USER_ID,
      messages: [{ type: 'text', text: message }]
    };

    UrlFetchApp.fetch(LINE_API_URL, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.log('Line notification error:', e.message);
  }
}

function sendCustomLineMessage(message) {
  try {
    if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN === 'YOUR_LINE_CHANNEL_ACCESS_TOKEN') {
      return { success: false, error: 'LINE OA Token not configured' };
    }

    const payload = {
      to: LINE_ADMIN_USER_ID,
      messages: [{ type: 'text', text: message }]
    };

    const response = UrlFetchApp.fetch(LINE_API_URL, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code === 200) {
      return { success: true, message: 'Message sent to LINE' };
    } else {
      return { success: false, error: 'LINE API returned ' + code + ': ' + response.getContentText() };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function replyToLine(replyToken, message) {
  try {
    const response = UrlFetchApp.fetch(LINE_REPLY_URL, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: [message]
      }),
      muteHttpExceptions: true
    });
    return response.getResponseCode() === 200;
  } catch (e) {
    console.error('Reply error:', e.message);
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════
// 📊 KPI & ANALYSIS REPORTS
// ═══════════════════════════════════════════════════════════════

function sendKpiReportToLine() {
  try {
    const allData = getSalesData();
    if (!allData.success || !allData.data.length) {
      return { success: false, error: 'No data available' };
    }

    const data = allData.data;
    const totalRev = data.reduce((s, r) => s + r.revenue, 0);
    const totalExp = data.reduce((s, r) => s + r.expense, 0);
    const totalBal = data.reduce((s, r) => s + r.balance, 0);
    const totalSold = data.reduce((s, r) => s + r.sold, 0);
    const totalPend = data.reduce((s, r) => s + r.pending, 0);
    const totalClr = data.reduce((s, r) => s + r.cleared, 0);
    const margin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;
    const convRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const todayData = data.filter(r => r.date === todayStr);
    const todayRevenue = todayData.reduce((s, r) => s + r.revenue, 0);
    const todayExpense = todayData.reduce((s, r) => s + r.expense, 0);
    const todayBalance = todayData.reduce((s, r) => s + r.balance, 0);

    const message =
      '📊 KPI REPORT — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '\n\n' +
      '📅 ภาพรวม (' + data.length + ' วัน)\n' +
      '💰 รายได้รวม: ฿' + totalRev.toLocaleString() + '\n' +
      '💸 ค่าใช้จ่ายรวม: ฿' + totalExp.toLocaleString() + '\n' +
      '📈 กำไรรวม: ฿' + totalBal.toLocaleString() + '\n' +
      '🎯 Profit Margin: ' + margin.toFixed(1) + '%\n' +
      '✅ Conversion Rate: ' + convRate.toFixed(1) + '%\n' +
      '🛒 ขายรวม: ' + totalSold + ' | ⏳ รอ: ' + totalPend + ' | ✅ ชำระ: ' + totalClr + '\n\n' +
      '📌 วันนี้ (' + todayStr + ')\n' +
      '💰 รายได้: ฿' + todayRevenue.toLocaleString() + '\n' +
      '💸 ค่าใช้จ่าย: ฿' + todayExpense.toLocaleString() + '\n' +
      '📈 กำไร: ฿' + todayBalance.toLocaleString();

    return sendCustomLineMessage(message);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function sendDeepAnalysisToLine() {
  try {
    const allData = getSalesData();
    if (!allData.success || !allData.data.length) {
      return { success: false, error: 'No data available' };
    }

    const data = allData.data;
    const totalRev = data.reduce((s, r) => s + r.revenue, 0);
    const totalExp = data.reduce((s, r) => s + r.expense, 0);
    const totalBal = data.reduce((s, r) => s + r.balance, 0);
    const margin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;

    const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    const dayStats = Array.from({ length: 7 }, () => ({
      count: 0, sold: 0, revenue: 0, expense: 0, balance: 0
    }));

    data.forEach(r => {
      if (!r.date) return;
      const dow = new Date(r.date).getDay();
      dayStats[dow].count++;
      dayStats[dow].sold += r.sold;
      dayStats[dow].revenue += r.revenue;
      dayStats[dow].expense += r.expense;
      dayStats[dow].balance += r.balance;
    });

    let bestDay = 0, worstDay = 0;
    dayStats.forEach((s, i) => {
      if (s.count > 0) {
        if (s.balance > dayStats[bestDay].balance || dayStats[bestDay].count === 0) bestDay = i;
        if (s.balance < dayStats[worstDay].balance || dayStats[worstDay].count === 0) worstDay = i;
      }
    });

    const pipeFee = data.reduce((s, r) => s + r.pipeFee, 0);
    const shareFee = data.reduce((s, r) => s + r.shareFee, 0);
    const otherFee = data.reduce((s, r) => s + r.otherFee, 0);
    const saveFee = data.reduce((s, r) => s + r.saveFee, 0);
    const totalFee = (pipeFee + shareFee + otherFee + saveFee) || 1;

    const sortedData = data.slice().sort((a, b) => a.date.localeCompare(b.date));
    const mid = Math.floor(sortedData.length / 2);
    const firstHalfBal = sortedData.slice(0, mid).reduce((s, r) => s + r.balance, 0);
    const secondHalfBal = sortedData.slice(mid).reduce((s, r) => s + r.balance, 0);

    let trend = '➡️ คงที่';
    if (secondHalfBal > firstHalfBal) trend = '📈 เพิ่มขึ้น';
    else if (secondHalfBal < firstHalfBal) trend = '📉 ลดลง';

    let message = '🧠 DEEP ANALYSIS REPORT\n\n';
    message += '📊 ภาพรวมทางการเงิน\n';
    message += '💰 รายได้รวม: ฿' + totalRev.toLocaleString() + '\n';
    message += '💸 ค่าใช้จ่ายรวม: ฿' + totalExp.toLocaleString() + '\n';
    message += '📈 กำไรรวม: ฿' + totalBal.toLocaleString() + '\n';
    message += '🎯 Profit Margin: ' + margin.toFixed(1) + '%\n\n';

    message += '📅 วิเคราะห์รายวัน\n';
    message += '🏆 วันที่ดีที่สุด: ' + dayNames[bestDay] + ' (฿' + dayStats[bestDay].balance.toLocaleString() + ')\n';
    message += '⚠️ วันที่ควรปรับ: ' + dayNames[worstDay] + ' (฿' + dayStats[worstDay].balance.toLocaleString() + ')\n\n';

    message += '💸 โครงสร้างค่าใช้จ่าย\n';
    message += '• Pipeline: ฿' + pipeFee.toLocaleString() + ' (' + (pipeFee / totalFee * 100).toFixed(1) + '%)\n';
    message += '• Share: ฿' + shareFee.toLocaleString() + ' (' + (shareFee / totalFee * 100).toFixed(1) + '%)\n';
    message += '• Other: ฿' + otherFee.toLocaleString() + ' (' + (otherFee / totalFee * 100).toFixed(1) + '%)\n';
    message += '• Save: ฿' + saveFee.toLocaleString() + ' (' + (saveFee / totalFee * 100).toFixed(1) + '%)\n\n';

    message += '📈 แนวโน้มธุรกิจ\n';
    message += 'ครึ่งแรก vs ครึ่งหลัง: ' + trend + '\n';
    message += 'ครึ่งแรก: ฿' + firstHalfBal.toLocaleString() + ' | ครึ่งหลัง: ฿' + secondHalfBal.toLocaleString() + '\n\n';

    message += '💡 คำแนะนำ\n';
    if (margin < 20) message += '🔴 เร่งด่วน: ปรับโครงสร้างต้นทุน Margin ต่ำกว่า 20%\n';
    if (pipeFee / totalFee > 0.5) message += '⚠️ Pipeline Fee สูงเกิน 50% ควรเจรจาลด\n';
    if (trend.includes('ลดลง')) message += '⚠️ แนวโน้มกำไรลดลง ควรตรวจสอบสาเหตุ\n';
    message += '✅ ติดตาม KPI สัปดาห์ละครั้งเพื่อปรับกลยุทธ์\n';

    return sendCustomLineMessage(message);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// 📑 AUTO REPORT SHEET GENERATOR
// ═══════════════════════════════════════════════════════════════

function generateAnalyticsReport() {
  try {
    const ss = _ss();
    const old = ss.getSheetByName(REPORT_SHEET);
    if (old) ss.deleteSheet(old);

    const rs = ss.insertSheet(REPORT_SHEET);
    const src = _sheet();
    const lastRow = src.getLastRow();

    if (lastRow < 2) {
      rs.getRange('A1').setValue('No data available.');
      return { success: true, message: 'No data to analyse.' };
    }

    const values = src.getRange(2, 1, lastRow - 1, 12).getValues();
    const data = values.map((r, i) => _parseRow(r, i + 2)).filter(Boolean);

    if (!data.length) {
      rs.getRange('A1').setValue('No data available.');
      return { success: true, message: 'No data to analyse.' };
    }

    const C = {
      headerBg: '#0f1724', headerFg: '#ffffff',
      titleBg: '#1a2744', titleFg: '#7eb8ff',
      altRow: '#f0f4ff', posText: '#1a7a4a',
      negText: '#c0392b'
    };

    function writeBlock(sheet, startRow, title, headers, rows, formats) {
      const titleCell = sheet.getRange(startRow, 1, 1, headers.length);
      titleCell.merge()
        .setValue(title)
        .setBackground(C.titleBg)
        .setFontColor(C.titleFg)
        .setFontWeight('bold')
        .setFontSize(12);

      const hdrRange = sheet.getRange(startRow + 1, 1, 1, headers.length);
      hdrRange.setValues([headers])
        .setBackground(C.headerBg)
        .setFontColor(C.headerFg)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

      rows.forEach((row, ri) => {
        const r = sheet.getRange(startRow + 2 + ri, 1, 1, row.length);
        r.setValues([row]);
        if (ri % 2 === 1) r.setBackground(C.altRow);
        if (formats) {
          formats.forEach((f, fi) => {
            if (f) sheet.getRange(startRow + 2 + ri, fi + 1).setNumberFormat(f);
          });
        }
      });

      return startRow + 2 + rows.length + 2;
    }

    let row = 1;
    const titleCell = rs.getRange(row, 1, 1, 8);
    titleCell.merge()
      .setValue('📊 SALES ANALYTICS REPORT — ' + _fmt(new Date(), 'dd/MM/yyyy HH:mm'))
      .setBackground(C.titleBg)
      .setFontColor(C.titleFg)
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    row += 3;

    const totalRev = data.reduce((s, r) => s + r.revenue, 0);
    const totalExp = data.reduce((s, r) => s + r.expense, 0);
    const totalBal = data.reduce((s, r) => s + r.balance, 0);
    const totalSold = data.reduce((s, r) => s + r.sold, 0);
    const totalPend = data.reduce((s, r) => s + r.pending, 0);
    const totalClr = data.reduce((s, r) => s + r.cleared, 0);
    const margin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;
    const avgBalDay = data.length ? totalBal / data.length : 0;
    const convRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;

    row = writeBlock(rs, row, '📊 OVERVIEW SUMMARY',
      ['Metric', 'Value', 'Note'],
      [
        ['📅 Total Days', data.length, 'วันที่มีข้อมูล'],
        ['💰 Total Revenue', totalRev, '฿' + totalRev.toLocaleString()],
        ['💸 Total Expense', totalExp, '฿' + totalExp.toLocaleString()],
        ['📈 Net Profit', totalBal, '฿' + totalBal.toLocaleString()],
        ['📊 Profit Margin', margin.toFixed(2) + '%', margin >= 30 ? '✅ Good' : margin >= 10 ? '⚠️ Average' : '❌ Low'],
        ['🛒 Total Sold', totalSold, 'ชิ้น'],
        ['⏳ Total Pending', totalPend, 'ชิ้น (' + (totalSold > 0 ? (totalPend / totalSold * 100).toFixed(1) : 0) + '%)'],
        ['✅ Total Cleared', totalClr, 'ชิ้น'],
        ['🎯 Conversion Rate', convRate.toFixed(1) + '%', convRate >= 70 ? '✅ Good' : '⚠️ Needs attention'],
        ['📅 Avg Profit/Day', Math.round(avgBalDay), '฿' + Math.round(avgBalDay).toLocaleString()]
      ],
      ['@', '#,##0.00', '@']
    );

    const dayNames = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์'];
    const dayStats = Array.from({ length: 7 }, () => ({
      cnt: 0, sold: 0, pending: 0, cleared: 0, rev: 0, exp: 0, bal: 0
    }));

    data.forEach(r => {
      if (!r.date) return;
      const dow = new Date(r.date).getDay();
      const idx = dow === 0 ? 6 : dow - 1;
      const s = dayStats[idx];
      s.cnt++; s.sold += r.sold; s.pending += r.pending;
      s.cleared += r.cleared; s.rev += r.revenue;
      s.exp += r.expense; s.bal += r.balance;
    });

    const dayRows = dayStats.map((s, i) => [
      dayNames[i], s.cnt, s.sold,
      s.cnt > 0 ? +(s.sold / s.cnt).toFixed(1) : 0,
      s.rev, s.cnt > 0 ? +(s.rev / s.cnt).toFixed(0) : 0,
      s.exp, s.bal,
      s.cnt > 0 ? +(s.bal / s.cnt).toFixed(0) : 0,
      s.rev > 0 ? +(s.bal / s.rev * 100).toFixed(1) : 0
    ]);

    row = writeBlock(rs, row, '📅 DAY OF WEEK ANALYSIS',
      ['วัน', 'จำนวนวัน', 'ขายรวม', 'ขายเฉลี่ย/วัน', 'รายได้รวม', 'รายได้เฉลี่ย/วัน', 'ค่าใช้จ่ายรวม', 'กำไรรวม', 'กำไรเฉลี่ย/วัน', 'Margin (%)'],
      dayRows,
      ['@', '#,##0', '#,##0', '#,##0.0', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0']
    );

    const monthMap = {};
    data.forEach(r => {
      if (!r.date) return;
      const ym = r.date.slice(0, 7);
      if (!monthMap[ym]) monthMap[ym] = { cnt: 0, sold: 0, pending: 0, cleared: 0, rev: 0, exp: 0, bal: 0 };
      const m = monthMap[ym];
      m.cnt++; m.sold += r.sold; m.pending += r.pending;
      m.cleared += r.cleared; m.rev += r.revenue;
      m.exp += r.expense; m.bal += r.balance;
    });

    const monthRows = Object.entries(monthMap).sort().map(entry => {
      const ym = entry[0], m = entry[1];
      const parts = ym.split('-');
      const yr = parts[0], mo = parts[1];
      const thMonths = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
                        'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
      return [
        thMonths[+mo] + ' ' + yr, m.cnt, m.sold,
        m.rev, m.exp, m.bal,
        m.rev > 0 ? +(m.bal / m.rev * 100).toFixed(1) : 0,
        m.cnt > 0 ? +(m.rev / m.cnt).toFixed(0) : 0,
        m.cnt > 0 ? +(m.bal / m.cnt).toFixed(0) : 0
      ];
    });

    row = writeBlock(rs, row, '📆 MONTHLY ANALYSIS',
      ['เดือน', 'จำนวนวัน', 'ยอดขายรวม', 'รายได้รวม', 'ค่าใช้จ่ายรวม', 'กำไรรวม', 'Margin (%)', 'รายได้เฉลี่ย/วัน', 'กำไรเฉลี่ย/วัน'],
      monthRows,
      ['@', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0', '#,##0', '#,##0']
    );

    const totPipe = data.reduce((s, r) => s + r.pipeFee, 0);
    const totShare = data.reduce((s, r) => s + r.shareFee, 0);
    const totOther = data.reduce((s, r) => s + r.otherFee, 0);
    const totSave = data.reduce((s, r) => s + r.saveFee, 0);
    const totAll = (totPipe + totShare + totOther + totSave) || 1;

    row = writeBlock(rs, row, '💸 EXPENSE BREAKDOWN',
      ['ประเภท', 'รวม (฿)', 'สัดส่วน (%)', 'เฉลี่ย/วัน (฿)', 'วิเคราะห์'],
      [
        ['Pipeline Fee', totPipe, +(totPipe / totAll * 100).toFixed(1), +(totPipe / data.length).toFixed(0), totPipe / totAll > 0.5 ? '⚠️ สูงมาก' : '✅ ปกติ'],
        ['Share Fee', totShare, +(totShare / totAll * 100).toFixed(1), +(totShare / data.length).toFixed(0), totShare / totAll > 0.3 ? '⚠️ ตรวจสอบ' : '✅ ปกติ'],
        ['Other Fee', totOther, +(totOther / totAll * 100).toFixed(1), +(totOther / data.length).toFixed(0), '—'],
        ['Save Fee', totSave, +(totSave / totAll * 100).toFixed(1), +(totSave / data.length).toFixed(0), '—'],
        ['TOTAL', totalExp, 100, +(totalExp / data.length).toFixed(0), '']
      ],
      ['@', '#,##0', '#,##0.0', '#,##0', '@']
    );

    const top5 = data.slice().sort((a, b) => b.balance - a.balance).slice(0, 5);
    const bot5 = data.slice().sort((a, b) => a.balance - b.balance).slice(0, 5);

    row = writeBlock(rs, row, '🏆 TOP 5 BEST DAYS',
      ['วันที่', 'Sold', 'Revenue', 'Expense', 'Balance', 'Margin (%)'],
      top5.map(r => [r.date, r.sold, r.revenue, r.expense, r.balance,
                     r.revenue > 0 ? +(r.balance / r.revenue * 100).toFixed(1) : 0]),
      ['@', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0']
    );

    row = writeBlock(rs, row, '⚠️ TOP 5 WORST DAYS',
      ['วันที่', 'Sold', 'Revenue', 'Expense', 'Balance', 'Margin (%)'],
      bot5.map(r => [r.date, r.sold, r.revenue, r.expense, r.balance,
                     r.revenue > 0 ? +(r.balance / r.revenue * 100).toFixed(1) : 0]),
      ['@', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0']
    );

    const insights = _buildInsightText(data, totalRev, totalExp, totalBal, margin, convRate);
    const insHeader = rs.getRange(row, 1, 1, 6);
    insHeader.merge()
      .setValue('🧠 BUSINESS INSIGHTS')
      .setBackground(C.titleBg)
      .setFontColor(C.titleFg)
      .setFontWeight('bold')
      .setFontSize(12);
    row++;

    insights.forEach(ins => {
      const cell = rs.getRange(row, 1, 1, 6);
      cell.merge()
        .setValue(ins.icon + ' ' + ins.title + ': ' + ins.body)
        .setFontColor(ins.type === 'positive' ? C.posText :
                      ins.type === 'negative' ? C.negText : '#7a5f00')
        .setWrap(true);
      rs.setRowHeight(row, 40);
      row++;
    });

    rs.setColumnWidth(1, 160);
    for (let c = 2; c <= 10; c++) rs.setColumnWidth(c, 130);
    rs.setFrozenRows(3);
    ss.setActiveSheet(rs);

    return { success: true, message: 'Report "' + REPORT_SHEET + '" generated!' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _buildInsightText(data, totalRev, totalExp, totalBal, margin, convRate) {
  const ins = [];
  const days = data.length;
  const avgBal = days ? totalBal / days : 0;
  const negDays = data.filter(r => r.balance < 0).length;
  const maxBal = days ? Math.max.apply(null, data.map(r => r.balance)) : 0;
  const minBal = days ? Math.min.apply(null, data.map(r => r.balance)) : 0;

  if (margin >= 40) {
    ins.push({ type: 'positive', icon: '🚀', title: 'Margin ดีเยี่ยม',
      body: 'Profit Margin ' + margin.toFixed(1) + '% สูงกว่า 40%' });
  } else if (margin >= 20) {
    ins.push({ type: 'warning', icon: '📊', title: 'Margin ปานกลาง',
      body: 'Margin ' + margin.toFixed(1) + '% ควรพยายามให้เกิน 30%' });
  } else {
    ins.push({ type: 'negative', icon: '🔴', title: 'Margin ต่ำ',
      body: 'Margin ' + margin.toFixed(1) + '% ต่ำกว่าเกณฑ์ ควรปรับโครงสร้างต้นทุน' });
  }

  if (negDays > 0) {
    ins.push({ type: 'negative', icon: '📉', title: negDays + ' วันขาดทุน',
      body: negDays + ' วันคิดเป็น ' + (negDays / days * 100).toFixed(0) + '%' });
  }

  if (convRate < 60) {
    ins.push({ type: 'warning', icon: '⏳', title: 'Conversion ต่ำ',
      body: 'Conversion ' + convRate.toFixed(1) + '% ควรกระตุ้นการชำระ' });
  } else {
    ins.push({ type: 'positive', icon: '✅', title: 'Conversion ดี',
      body: convRate.toFixed(1) + '% ลูกค้าชำระเร็ว' });
  }

  ins.push({ type: 'info', icon: '💡', title: 'สรุป',
    body: days + ' วัน | สูงสุด ฿' + maxBal.toLocaleString() +
          ' | ต่ำสุด ฿' + minBal.toLocaleString() +
          ' | เฉลี่ย ฿' + Math.round(avgBal).toLocaleString() + '/วัน' });

  return ins;
}

// ═══════════════════════════════════════════════════════════════
// 🌐 WEB API ENDPOINT (doPost)
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);

    // LINE Webhook Event
    if (contents.events) {
      contents.events.forEach(event => {
        try {
          if (event.type === 'message' && event.message.type === 'text') {
            handleLineMessageEnhanced(event);
          } else if (event.type === 'postback') {
            handlePostbackEvent(event);
          } else if (event.type === 'follow') {
            replyToLine(event.replyToken, txtMsg(
              '🎉 ยินดีต้อนรับ!\n\n' +
              'ฉันคือ SSKYMT Sales Analytics Bot v' + SYSTEM_VERSION + '\n\n' +
              '📝 บันทึกข้อมูล: พิมพ์ "ขาย 10"\n' +
              '📊 ดูรายงาน: พิมพ์ /help'
            ));
          }
        } catch (evtErr) {
          console.error('Event handler error:', evtErr.message);
        }
      });
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // REST API
    if (contents.action) {
      let result;
      const params = contents;
      switch (params.action) {
        case 'getSalesData':
          result = getSalesData(); break;
        case 'getQuickStats':
          result = { success: true, data: getQuickStats() }; break;
        case 'executeCommand':
          result = { success: true, data: executeWebCommand(params.cmd) }; break;
        case 'addRecord':
          result = addRecord(params.record); break;
        case 'updateRecord':
          result = updateRecord(params.rowIndex, params.record); break;
        case 'deleteRecord':
          result = deleteRecord(params.rowIndex); break;
        case 'getFilteredInsights':
          result = getFilteredInsightsData(params.filterParams); break;
        case 'generateReport':
          result = generateAnalyticsReport(); break;
        case 'sendKpiReport':
          result = sendKpiReportToLine(); break;
        case 'sendDeepAnalysis':
          result = sendDeepAnalysisToLine(); break;
        case 'sendLineMessage':
          result = sendCustomLineMessage(params.message); break;
        case 'getStockSummary':
          result = getStockSummary(); break;
        case 'getAllStockTransactions':
          result = getAllStockTransactions(); break;
        case 'ADD_STOCK':
          result = addStockTransaction(params.type, params.qty, params.source, params.userId); break;
        case 'EDIT_STOCK':
          result = editStockTransaction(params.id, params.type, params.qty, params.source); break;
        case 'DELETE_STOCK':
          result = deleteStockTransaction(params.id); break;
        case 'getAdvancedAnalytics':
          result = { success: true, data: getAdvancedAnalytics() }; break;
        case 'getCorrelationAnalysis':
          result = { success: true, data: getCorrelationAnalysis() }; break;
        case 'runMonteCarlo':
          result = { success: true, data: runMonteCarloSimulation() }; break;
        case 'runWhatIfAnalysis':
          result = { success: true, data: runWhatIfAnalysis(params.scenarios) }; break;
        case 'verifySystem':
          result = { success: true, data: verifySystem() }; break;
        default:
          result = { success: false, error: 'Unknown action: ' + action };
      }

      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('doPost error:', error);
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════════
// 🧠 ADVANCED ANALYTICS
// ═══════════════════════════════════════════════════════════════

function getAdvancedAnalytics() {
  try {
    const data = getAllSalesDataForBot();
    if (!data.length) return { error: 'No data available' };

    return {
      timestamp: new Date().toISOString(),
      correlation_matrix: calculateCorrelationMatrix(data),
      forecast_30days: advancedTimeSeriesForecast(data, 30),
      anomalies_detected: detectAnomaliesAdvanced(data),
      sales_patterns: clusterSalesPatterns(data),
      seasonality_analysis: analyzeSeasonalityAdvanced(data),
      what_if_scenarios: generateWhatIfScenarios(data),
      monte_carlo: runMonteCarloSimulation(data),
      pricing_optimization: optimizePricing(data),
      root_causes: findRootCauses(data),
      business_lifecycle: analyzeBusinessLifecycle(data)
    };
  } catch (e) {
    return { error: e.message };
  }
}

function getCorrelationAnalysis() {
  try {
    const data = getAllSalesDataForBot();
    if (!data.length) return { error: 'No data available' };
    return calculateCorrelationMatrix(data);
  } catch (e) {
    return { error: e.message };
  }
}

function calculateCorrelationMatrix(data) {
  const vars = {
    sold: data.map(r => r.sold),
    pending: data.map(r => r.pending),
    cleared: data.map(r => r.cleared),
    revenue: data.map(r => r.revenue),
    pipeFee: data.map(r => r.pipeFee),
    shareFee: data.map(r => r.shareFee),
    otherFee: data.map(r => r.otherFee),
    saveFee: data.map(r => r.saveFee),
    expense: data.map(r => r.expense),
    balance: data.map(r => r.balance)
  };

  const keys = Object.keys(vars);
  const matrix = {};

  for (let i = 0; i < keys.length; i++) {
    matrix[keys[i]] = {};
    for (let j = 0; j < keys.length; j++) {
      if (i === j) matrix[keys[i]][keys[j]] = 1;
      else {
        const corr = pearsonCorrelation(vars[keys[i]], vars[keys[j]]);
        matrix[keys[i]][keys[j]] = parseFloat(corr.toFixed(3));
      }
    }
  }

  const insights = [];
  if (matrix.pipeFee && matrix.pipeFee.balance && matrix.pipeFee.balance < -0.5) {
    insights.push({
      type: 'warning',
      message: 'Pipeline Fee มีผลเสียต่อกำไรสูง (' + matrix.pipeFee.balance + ')',
      recommendation: 'ควรเจรจาลด Pipeline Fee'
    });
  }
  if (matrix.sold && matrix.sold.balance && matrix.sold.balance > 0.7) {
    insights.push({
      type: 'positive',
      message: 'ยอดขายสัมพันธ์กับกำไรโดยตรง (' + matrix.sold.balance + ')',
      recommendation: 'เน้นเพิ่มยอดขายเพื่อเพิ่มกำไร'
    });
  }

  return { matrix, insights };
}

function pearsonCorrelation(x, y) {
  const n = x.length;
  if (n === 0) return 0;
  const sumX = x.reduce((a, b) => a + b, 0);
  const sumY = y.reduce((a, b) => a + b, 0);
  const sumXY = x.reduce((a, b, i) => a + b * y[i], 0);
  const sumX2 = x.reduce((a, b) => a + b * b, 0);
  const sumY2 = y.reduce((a, b) => a + b * b, 0);

  const numerator = n * sumXY - sumX * sumY;
  const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
  return denominator === 0 ? 0 : numerator / denominator;
}

function advancedTimeSeriesForecast(data, days) {
  days = days || 30;
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const sales = sorted.map(r => r.sold);
  const n = sales.length;

  if (n < 7) return { error: 'Insufficient data (need 7+)' };

  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
  sorted.forEach((r, i) => {
    sumX += i;
    sumY += r.balance;
    sumXY += i * r.balance;
    sumX2 += i * i;
  });

  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;

  const ma7 = [];
  for (let i = 6; i < sales.length; i++) {
    const avg = sales.slice(i - 6, i + 1).reduce((a, b) => a + b, 0) / 7;
    ma7.push(avg);
  }

  const forecast = [];
  const confidence = [];

  for (let i = 0; i < days; i++) {
    const linearPred = slope * (n + i) + intercept;
    const maPred = ma7.length > 0 ? ma7[ma7.length - 1] : linearPred;
    const finalPred = linearPred * 0.6 + maPred * 0.4;

    const error = Math.abs(finalPred - sales[sales.length - 1]) / (sales[sales.length - 1] || 1);
    const confidenceLower = finalPred * (1 - error * 1.96);
    const confidenceUpper = finalPred * (1 + error * 1.96);

    forecast.push(Math.max(0, Math.round(finalPred)));
    confidence.push({
      lower: Math.max(0, Math.round(confidenceLower)),
      upper: Math.max(0, Math.round(confidenceUpper))
    });
  }

  const totalForecast = forecast.reduce((a, b) => a + b, 0);
  const trend = slope > 0 ? 'increasing' : slope < 0 ? 'decreasing' : 'stable';

  return {
    daily: forecast,
    confidence_intervals: confidence,
    total_30day: totalForecast,
    trend: trend,
    slope: slope.toFixed(2),
    recommendation: trend === 'increasing' ? '📈 แนวโน้มเพิ่มขึ้น' :
                    trend === 'decreasing' ? '📉 แนวโน้มลดลง ควรปรับกลยุทธ์' :
                    '➡️ แนวโน้มทรงตัว'
  };
}

function detectAnomaliesAdvanced(data) {
  const balances = data.map(r => r.balance);
  if (!balances.length) return { anomalies: [], summary: { total_anomalies: 0, high_severity: 0, anomaly_rate: '0%' } };

  const mean = balances.reduce((a, b) => a + b, 0) / balances.length;
  const variance = balances.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / balances.length;
  const stdDev = Math.sqrt(variance);

  const anomalies = [];
  data.forEach(record => {
    if (stdDev === 0) return;
    const zScore = Math.abs((record.balance - mean) / stdDev);
    if (zScore > 2) {
      anomalies.push({
        date: record.date,
        balance: record.balance,
        zScore: zScore.toFixed(2),
        severity: zScore > 3 ? 'high' : 'medium',
        possible_causes: analyzePossibleCauses(record, data)
      });
    }
  });

  return {
    anomalies: anomalies,
    summary: {
      total_anomalies: anomalies.length,
      high_severity: anomalies.filter(a => a.severity === 'high').length,
      anomaly_rate: (anomalies.length / data.length * 100).toFixed(1) + '%'
    }
  };
}

function analyzePossibleCauses(record, allData) {
  const causes = [];
  const prevDay = allData.find(d => {
    const dDate = new Date(d.date);
    const rDate = new Date(record.date);
    dDate.setDate(dDate.getDate() + 1);
    return dDate.toISOString().slice(0, 10) === rDate.toISOString().slice(0, 10);
  });

  if (prevDay) {
    if (record.pipeFee > prevDay.pipeFee * 2) causes.push('Pipeline Fee เพิ่มขึ้นผิดปกติ');
    if (record.sold < prevDay.sold * 0.5) causes.push('ยอดขายลดลงกว่าครึ่ง');
    if (record.pending > prevDay.pending * 3) causes.push('Pending เพิ่มขึ้นมาก');
  }

  const avgPipeFee = allData.reduce((s, r) => s + r.pipeFee, 0) / allData.length;
  if (record.pipeFee > avgPipeFee * 2) causes.push('Pipeline Fee สูงกว่าค่าเฉลี่ย 2 เท่า');

  return causes.length ? causes : ['ตรวจสอบรายละเอียดเพิ่มเติม'];
}

function clusterSalesPatterns(data) {
  const weeklyPatterns = [];
  const groupedByWeek = {};

  data.forEach(record => {
    const week = getWeekNumber(new Date(record.date));
    if (!groupedByWeek[week]) groupedByWeek[week] = [];
    groupedByWeek[week].push(record);
  });

  const weekKeys = Object.keys(groupedByWeek).sort();
  const recentWeeks = weekKeys.slice(-8);

  recentWeeks.forEach(week => {
    const pattern = [0, 0, 0, 0, 0, 0, 0];
    groupedByWeek[week].forEach(record => {
      const dow = new Date(record.date).getDay();
      pattern[dow] += record.sold;
    });
    weeklyPatterns.push(pattern);
  });

  const avgPattern = [0, 0, 0, 0, 0, 0, 0];
  weeklyPatterns.forEach(pattern => {
    pattern.forEach((val, i) => { avgPattern[i] += val; });
  });
  const wCount = weeklyPatterns.length || 1;
  avgPattern.forEach((val, i) => { avgPattern[i] = val / wCount; });

  const peakDay = avgPattern.indexOf(Math.max.apply(null, avgPattern));
  const nonZero = avgPattern.filter(v => v > 0);
  const minVal = nonZero.length ? Math.min.apply(null, nonZero) : 0;
  const lowDay = avgPattern.indexOf(minVal);
  const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

  return {
    average_pattern: avgPattern,
    peak_day: { name: dayNames[peakDay], index: peakDay, avg_sales: avgPattern[peakDay] },
    low_day: { name: dayNames[lowDay], index: lowDay, avg_sales: avgPattern[lowDay] },
    recommendation: 'เน้นการตลาดวัน' + dayNames[peakDay] + ' และปรับปรุงวัน' + dayNames[lowDay]
  };
}

function analyzeSeasonalityAdvanced(data) {
  const monthly = {};
  const quarterly = {};
  const yearly = {};

  data.forEach(record => {
    const date = new Date(record.date);
    const month = date.getMonth() + 1;
    const quarter = Math.ceil(month / 3);
    const year = date.getFullYear();

    if (!monthly[month]) monthly[month] = { total: 0, count: 0 };
    if (!quarterly[quarter]) quarterly[quarter] = { total: 0, count: 0 };
    if (!yearly[year]) yearly[year] = { total: 0, count: 0 };

    monthly[month].total += record.balance;
    monthly[month].count++;
    quarterly[quarter].total += record.balance;
    quarterly[quarter].count++;
    yearly[year].total += record.balance;
    yearly[year].count++;
  });

  const monthNames = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
                      'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const monthlyAvg = {};
  for (let m = 1; m <= 12; m++) {
    monthlyAvg[m] = monthly[m] ? monthly[m].total / monthly[m].count : 0;
  }

  const entries = Object.entries(monthlyAvg).filter(e => e[1] !== 0);
  const bestMonth = entries.length ? entries.reduce((a, b) => a[1] > b[1] ? a : b) : ['1', 0];
  const worstMonth = entries.length ? entries.reduce((a, b) => a[1] < b[1] ? a : b) : ['1', 0];

  return {
    monthly: monthlyAvg,
    monthly_peak: { month: monthNames[parseInt(bestMonth[0]) - 1], avg_profit: bestMonth[1] },
    monthly_low: { month: monthNames[parseInt(worstMonth[0]) - 1], avg_profit: worstMonth[1] },
    quarterly: quarterly,
    yearly_growth: calculateYearlyGrowth(yearly)
  };
}

function calculateYearlyGrowth(yearly) {
  const years = Object.keys(yearly).sort();
  if (years.length < 2) return null;

  const growth = [];
  for (let i = 1; i < years.length; i++) {
    const prev = yearly[years[i - 1]].total;
    const curr = yearly[years[i]].total;
    const rate = prev > 0 ? ((curr - prev) / prev * 100) : 0;
    growth.push({ year: years[i], growth_rate: rate.toFixed(1) + '%' });
  }
  return growth;
}

function generateWhatIfScenarios(data) {
  const currentProfit = data.reduce((s, r) => s + r.balance, 0);
  const currentRevenue = data.reduce((s, r) => s + r.revenue, 0);
  const currentExpense = data.reduce((s, r) => s + r.expense, 0);
  const pipeFeeTotal = data.reduce((s, r) => s + r.pipeFee, 0);

  const scenarios = [
    {
      name: 'ลดค่าใช้จ่าย 10%',
      changes: { expense_reduction: 0.10 },
      projected_profit: currentProfit + (currentExpense * 0.10),
      improvement: ((currentExpense * 0.10) / Math.abs(currentProfit || 1) * 100).toFixed(1) + '%'
    },
    {
      name: 'เพิ่มยอดขาย 20%',
      changes: { sales_increase: 0.20 },
      projected_profit: currentProfit + (currentRevenue * 0.20 * (currentProfit / (currentRevenue || 1))),
      improvement: 'ขึ้นกับอัตรากำไรปัจจุบัน'
    },
    {
      name: 'ลด Pipeline Fee 30%',
      changes: { pipeline_reduction: 0.30 },
      projected_profit: currentProfit + (pipeFeeTotal * 0.30),
      improvement: ((pipeFeeTotal * 0.30) / Math.abs(currentProfit || 1) * 100).toFixed(1) + '%'
    },
    {
      name: 'เพิ่มราคา 10%',
      changes: { price_increase: 0.10 },
      projected_profit: currentRevenue * 0.10 * (currentProfit / (currentRevenue || 1)) + currentProfit,
      improvement: 'ขึ้นกับความยืดหยุ่นของอุปสงค์'
    }
  ];

  return {
    scenarios: scenarios,
    best_scenario: scenarios[0],
    recommendation: '🎯 ลองใช้ scenario ที่ให้ผลตอบแทนดีที่สุด'
  };
}

function runMonteCarloSimulation(data, iterations) {
  iterations = iterations || 1000;
  const dailyProfits = data.map(r => r.balance);
  if (!dailyProfits.length) return { error: 'No data' };

  const mean = dailyProfits.reduce((a, b) => a + b, 0) / dailyProfits.length;
  const variance = dailyProfits.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / dailyProfits.length;
  const stdDev = Math.sqrt(variance);

  const simulations = [];
  for (let i = 0; i < iterations; i++) {
    let total = 0;
    for (let day = 0; day < 30; day++) {
      let u = 0, v = 0;
      while (u === 0) u = Math.random();
      while (v === 0) v = Math.random();
      const z = Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
      const randomProfit = mean + stdDev * z;
      total += Math.max(-50000, randomProfit);
    }
    simulations.push(total);
  }

  simulations.sort((a, b) => a - b);

  const percentiles = {
    p5: simulations[Math.floor(iterations * 0.05)],
    p10: simulations[Math.floor(iterations * 0.10)],
    p25: simulations[Math.floor(iterations * 0.25)],
    p50: simulations[Math.floor(iterations * 0.50)],
    p75: simulations[Math.floor(iterations * 0.75)],
    p90: simulations[Math.floor(iterations * 0.90)],
    p95: simulations[Math.floor(iterations * 0.95)]
  };

  const probabilityPositive = simulations.filter(v => v > 0).length / iterations * 100;
  const expectedValue = simulations.reduce((a, b) => a + b, 0) / iterations;

  let riskLevel = 'low';
  if (probabilityPositive < 70) riskLevel = 'high';
  else if (probabilityPositive < 85) riskLevel = 'medium';

  return {
    iterations: iterations,
    expected_30day_profit: expectedValue,
    probability_positive: probabilityPositive.toFixed(1) + '%',
    percentiles: percentiles,
    risk_level: riskLevel,
    interpretation: riskLevel === 'high' ?
      '⚠️ ความเสี่ยงสูง มีโอกาส ' + (100 - probabilityPositive).toFixed(1) + '% ที่จะขาดทุนใน 30 วัน' :
      riskLevel === 'medium' ?
      '🟡 ความเสี่ยงปานกลาง ควรมีแผนสำรอง' :
      '✅ ความเสี่ยงต่ำ โอกาสทำกำไรสูง'
  };
}

function runWhatIfAnalysis(scenarios) {
  const data = getAllSalesDataForBot();
  if (!data.length) return { error: 'No data available' };
  return generateWhatIfScenarios(data);
}

function optimizePricing(data) {
  const pricePoints = [35, 38, 40, 42, 45, 48, 50];
  const elasticityData = [];
  const currentAvgSold = data.reduce((s, r) => s + r.sold, 0) / data.length;

  pricePoints.forEach(price => {
    const priceDiff = (price - PRICE_PER_UNIT) / PRICE_PER_UNIT;
    const expectedDemand = currentAvgSold * (1 - priceDiff * 0.8);
    const expectedRevenue = expectedDemand * price;
    elasticityData.push({
      price: price,
      expected_daily_sales: expectedDemand.toFixed(1),
      expected_daily_revenue: expectedRevenue,
      change_from_current: ((expectedRevenue - (currentAvgSold * PRICE_PER_UNIT)) / (currentAvgSold * PRICE_PER_UNIT) * 100).toFixed(1) + '%'
    });
  });

  const optimal = elasticityData.reduce((best, curr) =>
    curr.expected_daily_revenue > best.expected_daily_revenue ? curr : best, elasticityData[0]);

  return {
    current_price: PRICE_PER_UNIT,
    current_daily_revenue: currentAvgSold * PRICE_PER_UNIT,
    price_elasticity_data: elasticityData,
    optimal_price: optimal.price,
    optimal_daily_revenue: optimal.expected_daily_revenue,
    potential_gain: optimal.expected_daily_revenue - (currentAvgSold * PRICE_PER_UNIT),
    recommendation: optimal.price !== PRICE_PER_UNIT ?
      '📊 แนะนำปรับราคาเป็น ' + optimal.price + ' บาท' :
      '✅ ราคาปัจจุบันเหมาะสมแล้ว'
  };
}

function findRootCauses(data) {
  const rootCauses = [];
  const problemDays = data.filter(r => r.balance < 0);

  if (problemDays.length > 0) {
    const avgProblemPipeFee = problemDays.reduce((s, r) => s + r.pipeFee, 0) / problemDays.length;
    const avgProblemShareFee = problemDays.reduce((s, r) => s + r.shareFee, 0) / problemDays.length;
    const overallAvgPipeFee = data.reduce((s, r) => s + r.pipeFee, 0) / data.length;
    const overallAvgShareFee = data.reduce((s, r) => s + r.shareFee, 0) / data.length;

    if (avgProblemPipeFee > overallAvgPipeFee * 1.5) {
      rootCauses.push({
        cause: 'Pipeline Fee สูงเกินไป',
        impact: 'เป็นสาเหตุหลักของ ' + problemDays.length + ' วันที่ขาดทุน',
        solution: 'เจรจาลด Pipeline Fee'
      });
    }
    if (avgProblemShareFee > overallAvgShareFee * 1.5) {
      rootCauses.push({
        cause: 'Share Fee สูงเกินไป',
        impact: 'ส่งผลให้กำไรลดลง',
        solution: 'พิจารณาปรับโครงสร้างค่าคอมมิชชั่น'
      });
    }
  }

  const avgSold = data.reduce((s, r) => s + r.sold, 0) / data.length;
  const lowSalesDays = data.filter(r => r.sold < avgSold * 0.5);
  if (lowSalesDays.length > 5) {
    rootCauses.push({
      cause: 'ยอดขายต่ำผิดปกติ',
      impact: lowSalesDays.length + ' วัน ที่ยอดขายต่ำกว่าครึ่ง',
      solution: 'เพิ่มโปรโมชั่นในวันที่ยอดขายต่ำ'
    });
  }

  return {
    root_causes: rootCauses,
    summary: rootCauses.length === 0 ? '✅ ไม่พบสาเหตุหลัก' : '⚠️ พบ ' + rootCauses.length + ' สาเหตุ',
    priority_action: rootCauses[0] || null
  };
}

function analyzeBusinessLifecycle(data) {
  const totalDays = data.length;
  const totalProfit = data.reduce((s, r) => s + r.balance, 0);
  const avgProfit = totalDays ? totalProfit / totalDays : 0;

  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const firstMonth = sorted.filter(r => r.date.slice(0, 7) === sorted[0].date.slice(0, 7));
  const lastMonth = sorted.filter(r => r.date.slice(0, 7) === sorted[sorted.length - 1].date.slice(0, 7));

  const firstMonthAvg = firstMonth.length ? firstMonth.reduce((s, r) => s + r.balance, 0) / firstMonth.length : 0;
  const lastMonthAvg = lastMonth.length ? lastMonth.reduce((s, r) => s + r.balance, 0) / lastMonth.length : 0;
  const growthRate = firstMonthAvg > 0 ? ((lastMonthAvg - firstMonthAvg) / firstMonthAvg * 100) : 0;

  let stage = '', recommendation = '';
  if (totalDays < 30) {
    stage = '🚀 ระยะเริ่มต้น';
    recommendation = 'โฟกัสที่การสร้างฐานลูกค้า';
  } else if (growthRate > 20) {
    stage = '📈 ระยะเติบโต';
    recommendation = 'เร่งการผลิต ขยายช่องทางขาย';
  } else if (Math.abs(growthRate) < 10) {
    stage = '⚖️ ระยะ成熟 (Maturity)';
    recommendation = 'รักษาฐานลูกค้า หาช่องทางใหม่';
  } else if (growthRate < -10) {
    stage = '📉 ระยะถดถอย';
    recommendation = 'รีบปรับกลยุทธ์ ลดต้นทุน';
  } else {
    stage = '🔄 ระยะปรับตัว';
    recommendation = 'วิเคราะห์จุดอ่อนและปรับปรุง';
  }

  return {
    stage: stage,
    growth_rate: growthRate.toFixed(1) + '%',
    days_in_business: totalDays,
    avg_profit_per_day: avgProfit,
    recommendation: recommendation,
    next_milestone: growthRate > 0 ? 'รักษาการเติบโต' : 'กลับมาเติบโต'
  };
}

// ═══════════════════════════════════════════════════════════════
// 📱 BOT COMMAND PROCESSING
// ═══════════════════════════════════════════════════════════════

function processBotCommand(cmd) {
  const c = cmd.toLowerCase().trim();
  const data = getAllSalesDataForBot();

  if (!data.length) {
    return txtMsg('❌ ยังไม่มีข้อมูลในระบบ\n\n💡 ลองพิมพ์ "ขาย 10" เพื่อทดสอบบันทึกข้อมูล');
  }

  if (c.match(/\/(kpi|dashboard)/i)) return botKpiFlex(data);
  if (c.match(/\/(deep|analyze)/i)) return botDeepAnalysis(data);
  if (c.match(/\/(today|วันนี้)/i)) return botDaily(data, 'today');
  if (c.match(/\/(yesterday|เมื่อวาน)/i)) return botDaily(data, 'yesterday');
  if (c.match(/\/(week|สัปดาห์)/i)) return botWeekly(data);
  if (c.match(/\/(month|เดือน)/i)) return botMonthly(data);
  if (c.match(/\/(expense|ค่าใช้จ่าย)/i)) return botExpense(data);
  if (c.match(/\/(trend|แนวโน้ม)/i)) return botTrend(data);
  if (c.match(/\/(compare|เปรียบเทียบ)/i)) return botCompare(data);
  if (c.match(/\/(margin|profit)/i)) return botMargin(data);
  if (c.match(/\/(conversion|conv)/i)) return botConversion(data);
  if (c.match(/\/(sales|ยอดขาย)/i)) return botSales(data);
  if (c.match(/\/(advice|แนะนำ)/i)) return botAdvice(data);
  if (c.match(/\/(stats|สถิติ)/i)) return botStats(data);
  if (c.match(/\/(forecast|พยากรณ์|คาดการณ์)/i)) return botForecast(data);
  if (c.match(/\/(risk|ความเสี่ยง)/i)) return botRiskAnalysis(data);
  if (c.match(/\/(health|สุขภาพธุรกิจ)/i)) return botHealthScore(data);
  if (c.match(/\/(correlation|ความสัมพันธ์)/i)) return botCorrelationAnalysis(data);
  if (c.match(/\/(whatif|จำลอง|what-if)/i)) return botWhatIfAnalysis(data);
  if (c.match(/\/(montecarlo|mc)/i)) return botMonteCarlo(data);
  if (c.match(/\/(lifecycle|วัฏจักร)/i)) return botLifecycle(data);
  if (c.match(/\/(pricing|ราคา)/i)) return botPricingOptimization(data);
  if (c.match(/\/(stock|สต็อก|สต๊อก)$/i)) return botStockSummary(data);
  if (c.match(/\/(stockin|เข้า)\s+(\d+)/i)) return botStockIn(c);
  if (c.match(/\/(stockout|ออก|ขาย)\s+(\d+)/i)) return botStockOut(c);
  if (c.match(/\/(stocklog|ประวัติสต็อก)/i)) return botStockLog(c);
  if (c.match(/\/(stockvalue|มูลค่าสต็อก)/i)) return botStockValue(data);

  return botHelp();
}

function botHelp() {
  return txtMsg(
    '🤖 SALES ANALYTICS BOT v' + SYSTEM_VERSION + '\n\n' +
    '📝 **บันทึกข้อมูล**\n' +
    'พิมพ์: ขาย [จำนวน] [รายละเอียด]\n' +
    'ตัวอย่าง:\n' +
    '  • ขาย 10\n' +
    '  • ขาย 10 p50 s20\n' +
    '  • วันนี้ ขาย 10 pe5 cl5\n' +
    '  • 10/5 ขาย 20 p100\n\n' +
    '📊 **รายงานพื้นฐาน**\n' +
    '/kpi — KPI Dashboard\n' +
    '/today — รายงานวันนี้\n' +
    '/yesterday — รายงานเมื่อวาน\n' +
    '/week — รายงานสัปดาห์\n' +
    '/month — รายงานเดือน\n\n' +
    '🔍 **วิเคราะห์เชิงลึก**\n' +
    '/deep — วิเคราะห์เชิงลึก\n' +
    '/forecast — พยากรณ์ 30 วัน\n' +
    '/trend — แนวโน้มธุรกิจ\n' +
    '/risk — วิเคราะห์ความเสี่ยง\n' +
    '/correlation — ความสัมพันธ์\n' +
    '/whatif — จำลองสถานการณ์\n' +
    '/montecarlo — Monte Carlo\n' +
    '/lifecycle — วัฏจักรธุรกิจ\n' +
    '/pricing — วิเคราะห์ราคา\n\n' +
    '💰 **การเงิน**\n' +
    '/margin — Margin Analysis\n' +
    '/conversion — Conversion Rate\n' +
    '/expense — ค่าใช้จ่าย\n' +
    '/health — สุขภาพธุรกิจ\n\n' +
    '📦 **สต็อก**\n' +
    '/stock — สรุปสต็อก\n' +
    '/stockin [จำนวน] — รับเข้า\n' +
    '/stockout [จำนวน] — ขายออก\n' +
    '/stocklog — ประวัติ\n' +
    '/stockvalue — มูลค่าสต็อก\n\n' +
    '📋 /help — คำสั่งทั้งหมด'
  );
}

function fmtBot(n) {
  n = +n || 0;
  if (Math.abs(n) >= 1e6) return '฿' + (n / 1e6).toFixed(1) + 'M';
  if (Math.abs(n) >= 1e3) return '฿' + (n / 1e3).toFixed(0) + 'K';
  return '฿' + n.toLocaleString('th-TH', { maximumFractionDigits: 0 });
}

function formatCurrency(amount) {
  if (Math.abs(amount) >= 1e6) return '฿' + (amount / 1e6).toFixed(2) + 'M';
  if (Math.abs(amount) >= 1e3) return '฿' + (amount / 1e3).toFixed(1) + 'K';
  return '฿' + amount.toLocaleString('th-TH', { maximumFractionDigits: 0 });
}

function txtMsg(text) {
  return { type: 'text', text: text };
}

function getWeekNumber(date) {
  const d = new Date(date);
  const start = new Date(d.getFullYear(), 0, 1);
  const diff = d - start;
  return d.getFullYear() + '-W' + (Math.floor(diff / (7 * 24 * 60 * 60 * 1000)) + 1);
}

function rowKpi(label, value, color) {
  return {
    type: 'box',
    layout: 'horizontal',
    contents: [
      { type: 'text', text: label, size: 'sm', color: '#9ba3be', flex: 0 },
      { type: 'text', text: String(value), size: 'sm', weight: 'bold', color: color, align: 'end', flex: 1 }
    ],
    margin: 'sm'
  };
}

// ─── KPI Flex ────────────────────────────────────────────────
function botKpiFlex(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const exp = data.reduce((s, r) => s + r.expense, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;
  const conv = sold > 0 ? ((sold - pend) / sold * 100).toFixed(1) : 0;
  const status = margin >= 30 ? '🟢 ดีมาก' : margin >= 15 ? '🟡 ปานกลาง' : '🔴 ต้องปรับปรุง';

  return {
    type: 'flex',
    altText: '📊 KPI Dashboard',
    contents: {
      type: 'bubble',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: '📊 KPI DASHBOARD', weight: 'bold', size: 'lg', color: '#ffffff' },
          { type: 'text', text: status, size: 'sm', color: '#9ba3be', margin: 'md' }
        ],
        backgroundColor: '#1a2744'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          rowKpi('💰 รายได้รวม', fmtBot(rev), '#4d8bff'),
          rowKpi('💸 ค่าใช้จ่าย', fmtBot(exp), '#ff3d5a'),
          rowKpi('📈 กำไรรวม', fmtBot(bal), bal >= 0 ? '#10d98a' : '#ff3d5a'),
          { type: 'separator', margin: 'md' },
          rowKpi('🎯 Profit Margin', margin + '%', margin >= 30 ? '#10d98a' : '#f59e0b'),
          rowKpi('✅ Conversion', conv + '%', conv >= 70 ? '#10d98a' : '#f59e0b'),
          { type: 'separator', margin: 'md' },
          rowKpi('🛒 ยอดขาย', sold + ' ชิ้น', '#f59e0b'),
          rowKpi('⏳ รอชำระ', pend + ' ชิ้น', '#ff8c42'),
          rowKpi('📅 จำนวนวัน', data.length + ' วัน', '#a855f7'),
          rowKpi('📊 กำไรเฉลี่ย/วัน', fmtBot(data.length ? bal / data.length : 0), '#22d3ee')
        ]
      }
    }
  };
}

function botDeepAnalysis(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const exp = data.reduce((s, r) => s + r.expense, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;

  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length / 2);
  const fBal = sorted.slice(0, mid).reduce((s, r) => s + r.balance, 0);
  const sBal = sorted.slice(mid).reduce((s, r) => s + r.balance, 0);
  const trend = sBal > fBal ? '📈 เพิ่มขึ้น' : sBal < fBal ? '📉 ลดลง' : '➡️ คงที่';

  const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  const ds = Array(7).fill().map(() => ({ cnt: 0, bal: 0 }));
  data.forEach(r => {
    if (!r.date) return;
    const d = new Date(r.date).getDay();
    ds[d].cnt++;
    ds[d].bal += r.balance;
  });

  const best = ds.reduce((m, s, i) => s.cnt > 0 && s.bal > ds[m].bal ? i : m, 0);
  const worst = ds.reduce((m, s, i) => s.cnt > 0 && s.bal < ds[m].bal ? i : m, 0);

  const pf = data.reduce((s, r) => s + r.pipeFee, 0);
  const sf = data.reduce((s, r) => s + r.shareFee, 0);
  const of = data.reduce((s, r) => s + r.otherFee, 0);
  const sv = data.reduce((s, r) => s + r.saveFee, 0);
  const tf = (pf + sf + of + sv) || 1;

  let t = '🧠 DEEP ANALYSIS\n\n';
  t += '📊 ภาพรวม\n💰 รายได้: ' + fmtBot(rev) + '\n💸 ค่าใช้จ่าย: ' + fmtBot(exp) + '\n📈 กำไร: ' + fmtBot(bal) + ' (' + margin + '%)\n\n';
  t += '📈 แนวโน้ม: ' + trend + '\nครึ่งแรก: ' + fmtBot(fBal) + ' → ครึ่งหลัง: ' + fmtBot(sBal) + '\n\n';
  t += '📅 ดีที่สุด: ' + dayNames[best] + ' (' + fmtBot(ds[best].bal) + ')\n⚠️ ควรปรับ: ' + dayNames[worst] + ' (' + fmtBot(ds[worst].bal) + ')\n\n';
  t += '💸 ค่าใช้จ่าย\n• Pipe: ' + fmtBot(pf) + ' (' + (pf / tf * 100).toFixed(1) + '%)\n• Share: ' + fmtBot(sf) + ' (' + (sf / tf * 100).toFixed(1) + '%)\n• Other: ' + fmtBot(of) + ' (' + (of / tf * 100).toFixed(1) + '%)\n• Save: ' + fmtBot(sv) + ' (' + (sv / tf * 100).toFixed(1) + '%)\n';

  return txtMsg(t);
}

function botDaily(data, type) {
  const now = new Date();
  const target = type === 'yesterday' ? new Date(now.getTime() - 86400000) : now;
  const dateStr = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const thaiDate = Utilities.formatDate(target, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const dd = data.filter(r => r.date === dateStr);

  if (!dd.length) return txtMsg('📅 ' + thaiDate + '\n❌ ไม่มีข้อมูล');

  const rev = dd.reduce((s, r) => s + r.revenue, 0);
  const exp = dd.reduce((s, r) => s + r.expense, 0);
  const bal = dd.reduce((s, r) => s + r.balance, 0);
  const sold = dd.reduce((s, r) => s + r.sold, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;

  return txtMsg('📅 ' + thaiDate + '\n\n🛒 ยอดขาย: ' + sold + ' ชิ้น\n💰 รายได้: ' + fmtBot(rev) + '\n💸 ค่าใช้จ่าย: ' + fmtBot(exp) + '\n📈 กำไร: ' + fmtBot(bal) + '\n🎯 Margin: ' + margin + '%');
}

function botWeekly(data) {
  const now = new Date();
  const dow = now.getDay() || 7;
  const mon = new Date(now);
  mon.setDate(now.getDate() - dow + 1);
  const monStr = Utilities.formatDate(mon, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sun = new Date(mon);
  sun.setDate(mon.getDate() + 6);
  const sunStr = Utilities.formatDate(sun, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const wd = data.filter(r => r.date >= monStr && r.date <= sunStr);

  if (!wd.length) return txtMsg('📅 ' + monStr + ' - ' + sunStr + '\n❌ ไม่มีข้อมูล');

  const rev = wd.reduce((s, r) => s + r.revenue, 0);
  const exp = wd.reduce((s, r) => s + r.expense, 0);
  const bal = wd.reduce((s, r) => s + r.balance, 0);
  const sold = wd.reduce((s, r) => s + r.sold, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;

  return txtMsg('📅 ' + monStr + ' - ' + sunStr + '\n\n📊 ' + wd.length + ' วัน\n🛒 ' + sold + ' ชิ้น\n💰 รายได้: ' + fmtBot(rev) + '\n💸 ค่าใช้จ่าย: ' + fmtBot(exp) + '\n📈 กำไร: ' + fmtBot(bal) + '\n🎯 Margin: ' + margin + '%');
}

function botMonthly(data) {
  const now = new Date();
  const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  const md = data.filter(r => r.date.startsWith(monthStr));
  const thM = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const monthName = thM[now.getMonth() + 1];

  if (!md.length) return txtMsg('📅 ' + monthName + ' ' + now.getFullYear() + '\n❌ ไม่มีข้อมูล');

  const rev = md.reduce((s, r) => s + r.revenue, 0);
  const exp = md.reduce((s, r) => s + r.expense, 0);
  const bal = md.reduce((s, r) => s + r.balance, 0);
  const sold = md.reduce((s, r) => s + r.sold, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;

  return txtMsg('📅 ' + monthName + ' ' + now.getFullYear() + '\n\n📊 ' + md.length + ' วัน\n🛒 ' + sold + ' ชิ้น\n💰 รายได้: ' + fmtBot(rev) + '\n💸 ค่าใช้จ่าย: ' + fmtBot(exp) + '\n📈 กำไร: ' + fmtBot(bal) + '\n🎯 Margin: ' + margin + '%');
}

function botExpense(data) {
  const pf = data.reduce((s, r) => s + r.pipeFee, 0);
  const sf = data.reduce((s, r) => s + r.shareFee, 0);
  const of = data.reduce((s, r) => s + r.otherFee, 0);
  const sv = data.reduce((s, r) => s + r.saveFee, 0);
  const tf = (pf + sf + of + sv) || 1;

  return txtMsg('💸 EXPENSE BREAKDOWN\n\n📅 ' + data.length + ' วัน\n\nPipeline: ' + fmtBot(pf) + ' (' + (pf / tf * 100).toFixed(1) + '%)\nShare: ' + fmtBot(sf) + ' (' + (sf / tf * 100).toFixed(1) + '%)\nOther: ' + fmtBot(of) + ' (' + (of / tf * 100).toFixed(1) + '%)\nSave: ' + fmtBot(sv) + ' (' + (sv / tf * 100).toFixed(1) + '%)\n\nรวม: ' + fmtBot(tf));
}

function botTrend(data) {
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length / 2);
  const fBal = sorted.slice(0, mid).reduce((s, r) => s + r.balance, 0);
  const sBal = sorted.slice(mid).reduce((s, r) => s + r.balance, 0);
  const change = fBal !== 0 ? ((sBal - fBal) / Math.abs(fBal) * 100).toFixed(1) : 0;
  const trend = sBal > fBal ? '📈 เพิ่มขึ้น' : sBal < fBal ? '📉 ลดลง' : '➡️ คงที่';

  return txtMsg('📈 TREND ANALYSIS\n\n📅 ' + data.length + ' วัน\n\nครึ่งแรก: ' + fmtBot(fBal) + '\nครึ่งหลัง: ' + fmtBot(sBal) + '\n\n' + trend + ' (' + (change >= 0 ? '+' : '') + change + '%)');
}

function botCompare(data) {
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length / 2);
  const f = sorted.slice(0, mid);
  const s = sorted.slice(mid);
  const fRev = f.reduce((s2, r) => s2 + r.revenue, 0);
  const sRev = s.reduce((s2, r) => s2 + r.revenue, 0);
  const fBal = f.reduce((s2, r) => s2 + r.balance, 0);
  const sBal = s.reduce((s2, r) => s2 + r.balance, 0);
  const chRev = fRev > 0 ? ((sRev - fRev) / fRev * 100).toFixed(1) : 0;
  const chBal = fBal > 0 ? ((sBal - fBal) / Math.abs(fBal) * 100).toFixed(1) : 0;

  return txtMsg('📊 เปรียบเทียบ\n\n📅 ครึ่งแรก (' + f.length + ' วัน)\n💰 ' + fmtBot(fRev) + '\n📈 ' + fmtBot(fBal) + '\n\n📅 ครึ่งหลัง (' + s.length + ' วัน)\n💰 ' + fmtBot(sRev) + '\n📈 ' + fmtBot(sBal) + '\n\n📈 เปลี่ยนแปลง:\nรายได้: ' + (chRev >= 0 ? '+' : '') + chRev + '%\nกำไร: ' + (chBal >= 0 ? '+' : '') + chBal + '%');
}

function botMargin(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;

  const mm = {};
  data.forEach(r => {
    if (!r.date) return;
    const ym = r.date.slice(0, 7);
    if (!mm[ym]) mm[ym] = { r: 0, b: 0 };
    mm[ym].r += r.revenue;
    mm[ym].b += r.balance;
  });

  let t = '🎯 MARGIN ANALYSIS\n\n📊 Margin รวม: ' + margin + '%\n💰 ' + fmtBot(rev) + '\n📈 ' + fmtBot(bal) + '\n\n📅 รายเดือน:\n';
  Object.entries(mm).sort().forEach(([ym, m]) => {
    const mg = m.r > 0 ? (m.b / m.r * 100).toFixed(1) : 0;
    const e = mg >= 30 ? '🟢' : mg >= 15 ? '🟡' : '🔴';
    t += e + ' ' + ym + ': ' + mg + '%\n';
  });

  return txtMsg(t);
}

function botConversion(data) {
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const conv = sold > 0 ? ((sold - pend) / sold * 100).toFixed(1) : 0;

  return txtMsg('🎯 CONVERSION\n\n🛒 ขาย: ' + sold + ' ชิ้น\n⏳ รอชำระ: ' + pend + ' ชิ้น\n✅ ชำระแล้ว: ' + (sold - pend) + ' ชิ้น\n\n📊 Conversion: ' + conv + '%\nสถานะ: ' + (conv >= 70 ? '🟢 ดีมาก' : conv >= 50 ? '🟡 ปานกลาง' : '🔴 ต้องปรับปรุง'));
}

function botSales(data) {
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);

  return txtMsg('🛒 SALES REPORT\n\n📅 ' + data.length + ' วัน\n\n🛒 ขาย: ' + sold + ' ชิ้น\n✅ ชำระ: ' + (sold - pend) + ' ชิ้น\n⏳ รอ: ' + pend + ' ชิ้น\n📊 เฉลี่ย/วัน: ' + (sold / data.length).toFixed(1) + ' ชิ้น\n💵 มูลค่ารวม: ' + fmtBot(sold * PRICE_PER_UNIT));
}

function botAdvice(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const conv = sold > 0 ? ((sold - pend) / sold * 100).toFixed(1) : 0;

  let t = '💡 RECOMMENDATIONS\n\n';
  if (margin < 20) t += '🔴 เร่งด่วน: Margin ' + margin + '% ต่ำกว่า 20%\n\n';
  else if (margin < 30) t += '🟡 Margin ' + margin + '% ควรพยายามให้เกิน 30%\n\n';
  else t += '🟢 Margin ' + margin + '% ดีมาก\n\n';
  if (conv < 60) t += '⚠️ Conversion ' + conv + '% ต่ำ ควรลด Pending\n\n';
  t += '💰 ตรวจสอบ Cash Flow ทุกสัปดาห์\n🎯 ตั้งเป้า KPI รายสัปดาห์';

  return txtMsg(t);
}

function botStats(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;
  const best = data.reduce((mx, r) => r.balance > mx.balance ? r : mx, data[0]);
  const worst = data.reduce((mn, r) => r.balance < mn.balance ? r : mn, data[0]);

  return txtMsg('📊 QUICK STATS\n\n📅 ' + data.length + ' วัน\n💰 ' + fmtBot(rev) + '\n📈 ' + fmtBot(bal) + ' (' + margin + '%)\n🛒 ' + sold + ' ชิ้น\n\n🏆 ดีสุด: ' + best.date + ' (' + fmtBot(best.balance) + ')\n⚠️ แย่สุด: ' + worst.date + ' (' + fmtBot(worst.balance) + ')');
}

function botForecast(data) {
  const forecast = advancedTimeSeriesForecast(data, 7);
  if (forecast.error) return txtMsg('❌ ' + forecast.error);

  let t = '🔮 FORECAST (7 วัน)\n\n📈 แนวโน้ม: ' + (forecast.trend === 'increasing' ? '📈 ขาขึ้น' : forecast.trend === 'decreasing' ? '📉 ขาลง' : '➡️ คงที่') + '\n\n';
  const dayNames = ['พรุ่งนี้', 'มะรืนนี้', 'อีก 3 วัน', 'อีก 4 วัน', 'อีก 5 วัน', 'อีก 6 วัน', 'อีก 7 วัน'];
  forecast.daily.slice(0, 7).forEach((val, i) => {
    t += (val >= 0 ? '📈' : '📉') + ' ' + dayNames[i] + ': ' + fmtBot(val) + '\n';
  });
  t += '\n💡 ' + forecast.recommendation;

  return txtMsg(t);
}

function botRiskAnalysis(data) {
  const balances = data.map(r => r.balance);
  const avg = balances.reduce((a, b) => a + b, 0) / balances.length;
  const variance = balances.reduce((a, b) => a + Math.pow(b - avg, 2), 0) / balances.length;
  const stdDev = Math.sqrt(variance);
  const cv = avg !== 0 ? (stdDev / Math.abs(avg) * 100) : 0;
  const negDays = data.filter(r => r.balance < 0).length;
  const negRate = (negDays / data.length * 100);

  let riskScore = 0;
  if (cv > 100) riskScore += 3;
  else if (cv > 50) riskScore += 2;
  else if (cv > 20) riskScore += 1;
  if (negRate > 30) riskScore += 3;
  else if (negRate > 15) riskScore += 2;
  else if (negRate > 5) riskScore += 1;

  const riskLevel = riskScore >= 7 ? '🔴 สูงมาก' : riskScore >= 4 ? '🟠 สูง' : riskScore >= 2 ? '🟡 ปานกลาง' : '🟢 ต่ำ';

  let t = '⚠️ RISK ANALYSIS\n\n📊 Risk Score: ' + riskScore + '/9 — ' + riskLevel + '\n\n📈 CV: ' + cv.toFixed(1) + '%\n📉 วันที่ขาดทุน: ' + negDays + '/' + data.length + ' (' + negRate.toFixed(1) + '%)\n';

  if (riskScore >= 7) t += '\n🔴 ความเสี่ยงสูงมาก!';
  else if (riskScore >= 4) t += '\n🟠 ควรระมัดระวัง';
  else t += '\n🟢 ความเสี่ยงต่ำ';

  return txtMsg(t);
}

function botHealthScore(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);

  let score = 25;
  const margin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;
  if (margin >= 40) score += 20;
  else if (margin >= 20) score += 15;
  else if (margin >= 10) score += 10;
  else score += 5;

  const convRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;
  if (convRate >= 80) score += 15;
  else if (convRate >= 60) score += 10;
  else score += 5;

  const posDays = data.filter(r => r.balance >= 0).length;
  const consistency = (posDays / data.length * 100);
  if (consistency >= 90) score += 10;
  else if (consistency >= 70) score += 7;
  else score += 3;

  const grade = score >= 80 ? 'A' : score >= 60 ? 'B' : score >= 40 ? 'C' : 'D';
  const status = score >= 80 ? '🟢 แข็งแกร่ง' : score >= 60 ? '🟡 ปานกลาง' : '🔴 ต้องปรับปรุง';

  return txtMsg('🏥 HEALTH SCORE\n\nคะแนน: ' + score + '/100 — ' + grade + '\nสถานะ: ' + status + '\n\n🎯 Margin: ' + margin.toFixed(1) + '%\n✅ Conversion: ' + convRate.toFixed(1) + '%\n📊 ความสม่ำเสมอ: ' + consistency.toFixed(1) + '%');
}

function botCorrelationAnalysis(data) {
  const correlation = calculateCorrelationMatrix(data);
  let t = '📊 CORRELATION\n\n';

  if (correlation.insights && correlation.insights.length) {
    correlation.insights.forEach(insight => {
      t += insight.message + '\n👉 ' + insight.recommendation + '\n\n';
    });
  }

  t += '💡 ค่าใกล้ 1 = สัมพันธ์มาก, ใกล้ -1 = แปรผกผัน';
  return txtMsg(t);
}

function botWhatIfAnalysis(data) {
  const scenarios = generateWhatIfScenarios(data);
  let t = '🎯 WHAT-IF ANALYSIS\n\n';
  scenarios.scenarios.slice(0, 4).forEach(s => {
    t += s.name + '\n';
    t += '   กำไรคาดการณ์: ' + fmtBot(s.projected_profit) + '\n\n';
  });
  t += scenarios.recommendation;
  return txtMsg(t);
}

function botMonteCarlo(data) {
  const mc = runMonteCarloSimulation(data);
  let t = '🎲 MONTE CARLO\n\nจำลอง ' + mc.iterations + ' ครั้ง\n\n';
  t += '📊 ค่ากลาง: ' + fmtBot(mc.expected_30day_profit) + '\n';
  t += '✅ โอกาสกำไร: ' + mc.probability_positive + '\n';
  t += '🎯 ความเสี่ยง: ' + mc.risk_level + '\n\n';
  t += mc.interpretation;
  return txtMsg(t);
}

function botLifecycle(data) {
  const lc = analyzeBusinessLifecycle(data);
  let t = '🔄 BUSINESS LIFECYCLE\n\n📍 ' + lc.stage + '\n📈 เติบโต: ' + lc.growth_rate + '\n📅 ' + lc.days_in_business + ' วัน\n💰 เฉลี่ย/วัน: ' + fmtBot(lc.avg_profit_per_day) + '\n\n💡 ' + lc.recommendation + '\n🎯 ' + lc.next_milestone;
  return txtMsg(t);
}

function botPricingOptimization(data) {
  const p = optimizePricing(data);
  let t = '💰 PRICING OPTIMIZATION\n\nราคาปัจจุบัน: ฿' + p.current_price + '\n🎯 แนะนำ: ฿' + p.optimal_price + '\n📈 เพิ่มขึ้น: ' + fmtBot(p.potential_gain) + '/วัน\n\n' + p.recommendation;
  return txtMsg(t);
}

// ═══════════════════════════════════════════════════════════════
// 🎯 QUICK SALE RECORDER (MAIN INPUT FUNCTION - FIXED)
// ═══════════════════════════════════════════════════════════════

/**
 * บันทึกข้อมูลผ่าน LINE ด้วยข้อความสั้น
 * รูปแบบ:
 *   "ขาย 10"                → sold=10
 *   "ขาย 10 p50 s20 o10 sa5" → + fees
 *   "ขาย 10 pe5 cl5"        → + pending + cleared
 *   "วันนี้ ขาย 10"          → ระบุวันที่
 *   "10/5 ขาย 10"           → วันที่ dd/mm
 *   "/sell 10 p50"          → ใส่ / นำหน้าได้
 */
function recordQuickSale(cmd, event) {
  const c = cmd.replace(/^\//, '').trim();

  // ─── ตรวจรูปแบบ ──────────────────────────────────────────
  const isSaleFormat =
    c.match(/^(sell|ขาย|sale|add|บันทึก)/i) ||
    c.match(/^\d{1,2}\/\d{1,2}/) ||
    c.match(/^(วันนี้|พรุ่งนี้|เมื่อวาน|today|tomorrow|yesterday)/i);

  if (!isSaleFormat) return false;

  try {
    // ─── Parse วันที่ ─────────────────────────────────────
    let dateStr = '';
    const now = new Date();

    if (c.match(/^(วันนี้|today)/i)) {
      dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (c.match(/^(พรุ่งนี้|tomorrow)/i)) {
      const t = new Date(now);
      t.setDate(t.getDate() + 1);
      dateStr = Utilities.formatDate(t, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (c.match(/^(เมื่อวาน|yesterday)/i)) {
      const y = new Date(now);
      y.setDate(y.getDate() - 1);
      dateStr = Utilities.formatDate(y, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      const dateMatch = c.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?/);
      if (dateMatch) {
        const day = parseInt(dateMatch[1]);
        const month = parseInt(dateMatch[2]) - 1;
        const year = dateMatch[3] ? parseInt(dateMatch[3]) : now.getFullYear();
        const parsedDate = new Date(year, month, day);
        dateStr = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
    }

    // ─── Parse ตัวเลข ─────────────────────────────────────
    let sold = 0, pending = 0, cleared = 0;
    let pipeFee = 0, shareFee = 0, otherFee = 0, saveFee = 0;

    const soldMatch = c.match(/(?:^|\s)(\d+)(?:\s|$)/);
    if (soldMatch) sold = parseInt(soldMatch[1]) || 0;

    const pipeMatch = c.match(/\bp(\d+)/i);
    if (pipeMatch) pipeFee = parseInt(pipeMatch[1]) || 0;

    const shareMatch = c.match(/\bs(\d+)/i);
    if (shareMatch) shareFee = parseInt(shareMatch[1]) || 0;

    const otherMatch = c.match(/\bo(\d+)/i);
    if (otherMatch) otherFee = parseInt(otherMatch[1]) || 0;

    const saveMatch = c.match(/\bsa(\d+)/i);
    if (saveMatch) saveFee = parseInt(saveMatch[1]) || 0;

    const pendMatch = c.match(/\bpe(\d+)/i);
    if (pendMatch) pending = parseInt(pendMatch[1]) || 0;

    const clearMatch = c.match(/\bcl(\d+)/i);
    if (clearMatch) cleared = parseInt(clearMatch[1]) || 0;

    // ─── Validation ───────────────────────────────────────
    const hasAnyNumber =
      sold > 0 || pending > 0 || cleared > 0 ||
      pipeFee > 0 || shareFee > 0 || otherFee > 0 || saveFee > 0;

    if (!hasAnyNumber) {
      replyToLine(event.replyToken, txtMsg(
        '⚠️ ไม่พบตัวเลขในข้อความ\n\n' +
        '📝 รูปแบบที่ถูกต้อง:\n' +
        '• ขาย 10\n' +
        '• ขาย 10 p50 s20 o10 sa5\n' +
        '• ขาย 10 pe5 cl5\n' +
        '• วันนี้ ขาย 10\n' +
        '• 10/5 ขาย 10\n' +
        '• /sell 10 p50\n\n' +
        'ความหมาย:\n' +
        '• p = PipeFee\n' +
        '• s = ShareFee\n' +
        '• o = OtherFee\n' +
        '• sa = SaveFee\n' +
        '• pe = Pending\n' +
        '• cl = Cleared'
      ));
      return true;
    }

    // ─── คำนวณ ────────────────────────────────────────────
    const revenue = (sold + pending + cleared) * PRICE_PER_UNIT;
    const expense = pipeFee + shareFee + otherFee + saveFee;
    const balance = revenue - expense;

    // ─── บันทึก Sheet ─────────────────────────────────────
    const sheet = _sheet();
    sheet.appendRow([
      new Date(dateStr), sold, pending, cleared, revenue,
      pipeFee, shareFee, otherFee, saveFee, expense, balance, new Date()
    ]);

    // ─── Notify Admin (non-fatal) ─────────────────────────
    if (CONFIG.autoNotifyOnAdd) {
      try {
        sendLineNotification({
          type: 'add', date: dateStr, sold, pending, cleared,
          revenue, expense, balance
        });
      } catch (notifErr) {
        console.log('Notify error:', notifErr.message);
      }
    }

    // ─── Reply ────────────────────────────────────────────
    let message = '📝 บันทึกสำเร็จ!\n\n';
    message += '📅 ' + dateStr + '\n';
    message += '🛒 ขาย: ' + sold + ' ชิ้น\n';
    if (pending > 0) message += '⏳ รอชำระ: ' + pending + ' ชิ้น\n';
    if (cleared > 0) message += '✅ ชำระแล้ว: ' + cleared + ' ชิ้น\n';
    message += '━━━━━━━━━━━━━━━━━━━━\n';
    message += '💰 รายได้: ฿' + revenue.toLocaleString() + '\n';
    message += '💸 ค่าใช้จ่าย: ฿' + expense.toLocaleString() + '\n';
    if (pipeFee > 0) message += ' • Pipe: ฿' + pipeFee.toLocaleString() + '\n';
    if (shareFee > 0) message += ' • Share: ฿' + shareFee.toLocaleString() + '\n';
    if (otherFee > 0) message += ' • Other: ฿' + otherFee.toLocaleString() + '\n';
    if (saveFee > 0) message += ' • Save: ฿' + saveFee.toLocaleString() + '\n';
    message += '📈 กำไร: ฿' + balance.toLocaleString() + '\n';
    message += '🎯 Margin: ' + (revenue > 0 ? (balance / revenue * 100).toFixed(1) : 0) + '%\n';
    message += '\n✅ บันทึกใน Sheet แล้ว';

    replyToLine(event.replyToken, txtMsg(message));
    return true;

  } catch (err) {
    console.error('recordQuickSale error:', err);

    try {
      let errMsg = '❌ บันทึกไม่สำเร็จ\n\n';
      errMsg += 'ข้อผิดพลาด: ' + err.message + '\n\n';
      errMsg += '📋 กรุณาตรวจสอบ:\n';

      if (err.message.indexOf('not found') !== -1) {
        errMsg += '• ชื่อ Sheet "' + SHEET_NAME + '"\n';
        errMsg += '• ตัวพิมพ์เล็ก-ใหญ่\n';
      } else if (err.message.indexOf('permission') !== -1 ||
                 err.message.indexOf('authorization') !== -1) {
        errMsg += '• Permission ของ Spreadsheet\n';
        errMsg += '• ลอง Reauthorize\n';
      } else if (err.message.indexOf('Spreadsheet') !== -1) {
        errMsg += '• SPREADSHEET_ID\n';
      } else {
        errMsg += '• Sheet "' + SHEET_NAME + '"\n';
        errMsg += '• สิทธิ์การเข้าถึง\n';
      }

      errMsg += '\n💡 รัน setupSystem() ใน Apps Script';

      replyToLine(event.replyToken, txtMsg(errMsg));
    } catch (replyErr) {
      console.error('Reply error:', replyErr);
    }

    return true;
  }
}

// ═══════════════════════════════════════════════════════════════
// 📦 STOCK MANAGEMENT
// ═══════════════════════════════════════════════════════════════

function _stockSheet() {
  const ss = SpreadsheetApp.openById(STOCK_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(STOCK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STOCK_SHEET_NAME);
    sheet.appendRow(['ID', 'Timestamp', 'Type', 'Qty', 'Price', 'Total', 'Source']);
  }
  return sheet;
}

function getStockSummary() {
  try {
    const sheet = _stockSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true, stock: 0, value: 0,
        threshold: CONFIG.lowStockThreshold,
        isLow: true, totalIn: 0, totalOut: 0
      };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    let totalIn = 0, totalOut = 0;

    data.forEach(row => {
      const type = String(row[2] || '').toUpperCase();
      const qty = Number(row[3]) || 0;
      if (type === 'IN') totalIn += qty;
      else if (type === 'OUT') totalOut += qty;
    });

    const stock = totalIn - totalOut;
    const value = stock * PRICE_PER_UNIT;
    const isLow = stock < CONFIG.lowStockThreshold;

    return {
      success: true, stock, value,
      threshold: CONFIG.lowStockThreshold,
      isLow, totalIn, totalOut
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addStockTransaction(type, qty, source, userId) {
  try {
    const sheet = _stockSheet();
    const id = Utilities.getUuid();
    const timestamp = new Date();
    const price = PRICE_PER_UNIT;
    const total = qty * price;

    sheet.appendRow([id, timestamp, type.toUpperCase(), qty, price, total, source || 'Web']);

    if (type.toUpperCase() === 'OUT') {
      try {
        updateStockInSaleForm(type, qty, source);
      } catch (e) {
        console.log('updateStockInSaleForm error:', e.message);
      }
    }

    return {
      success: true, message: 'Transaction recorded',
      id: id, timestamp: timestamp.toISOString()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateStockInSaleForm(type, qty, source) {
  try {
    const sheet = _sheet();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const lastRow = sheet.getLastRow();
    const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 12).getValues() : [];

    let todayRow = -1;
    for (let i = values.length - 1; i >= 0; i--) {
      const rowDate = values[i][0];
      let dateStr = '';
      if (rowDate instanceof Date) {
        dateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = String(rowDate || '');
      }
      if (dateStr === today) {
        todayRow = i + 2;
        break;
      }
    }

    if (type.toUpperCase() === 'OUT') {
      if (todayRow > 0) {
        const currentSold = Number(sheet.getRange(todayRow, 2).getValue()) || 0;
        sheet.getRange(todayRow, 2).setValue(currentSold + qty);
        const pending = Number(sheet.getRange(todayRow, 3).getValue()) || 0;
        const cleared = Number(sheet.getRange(todayRow, 4).getValue()) || 0;
        const newRevenue = (currentSold + qty + pending + cleared) * PRICE_PER_UNIT;
        sheet.getRange(todayRow, 5).setValue(newRevenue);
        const expense = Number(sheet.getRange(todayRow, 10).getValue()) || 0;
        sheet.getRange(todayRow, 11).setValue(newRevenue - expense);
      } else {
        sheet.appendRow([new Date(today), qty, 0, 0, qty * PRICE_PER_UNIT,
                        0, 0, 0, 0, 0, qty * PRICE_PER_UNIT, new Date()]);
      }
    }
  } catch (e) {
    console.log('updateStockInSaleForm error:', e.message);
  }
}

function getAllStockTransactions() {
  try {
    const sheet = _stockSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const data = values.map((row, index) => ({
      rowIndex: index + 2,
      id: String(row[0] || ''),
      timestamp: row[1] instanceof Date ? row[1].toISOString() : String(row[1] || ''),
      type: String(row[2] || ''),
      qty: Number(row[3]) || 0,
      price: Number(row[4]) || 0,
      total: Number(row[5]) || 0,
      source: String(row[6] || '')
    })).filter(r => r.id !== '');

    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function editStockTransaction(id, type, qty, source) {
  try {
    const sheet = _stockSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'No transactions found' };

    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    let targetRow = -1;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow < 0) return { success: false, error: 'Transaction not found' };

    const price = PRICE_PER_UNIT;
    const total = qty * price;

    sheet.getRange(targetRow, 3).setValue(type.toUpperCase());
    sheet.getRange(targetRow, 4).setValue(qty);
    sheet.getRange(targetRow, 5).setValue(price);
    sheet.getRange(targetRow, 6).setValue(total);
    sheet.getRange(targetRow, 7).setValue(source || '');

    return { success: true, message: 'Transaction updated' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteStockTransaction(id) {
  try {
    const sheet = _stockSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'No transactions found' };

    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    let targetRow = -1;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow < 0) return { success: false, error: 'Transaction not found' };
    sheet.deleteRow(targetRow);

    return { success: true, message: 'Transaction deleted' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function botStockSummary(salesData) {
  try {
    const stock = getStockSummary();
    if (!stock.success) return txtMsg('❌ ไม่สามารถดึงข้อมูลสต็อกได้');

    const statusEmoji = stock.isLow ? '⚠️' : '✅';
    const statusText = stock.isLow ? 'สต็อกต่ำ!' : 'ปกติ';
    const pct = stock.threshold > 0 ? ((stock.stock / stock.threshold) * 100).toFixed(0) : 0;
    const bar = '█'.repeat(Math.min(10, Math.floor(pct / 10))) +
                '░'.repeat(Math.max(0, 10 - Math.floor(pct / 10)));

    let t = '📦 STOCK SUMMARY\n\n';
    t += '📊 คงเหลือ: ' + stock.stock.toLocaleString() + ' ขวด\n';
    t += '💰 มูลค่า: ฿' + stock.value.toLocaleString() + '\n';
    t += '📈 เข้า: ' + stock.totalIn.toLocaleString() + '\n';
    t += '📤 ออก: ' + stock.totalOut.toLocaleString() + '\n';
    t += '🎯 เกณฑ์เตือน (' + stock.threshold + ')\n';
    t += '[' + bar + '] ' + pct + '%\n';
    t += 'สถานะ: ' + statusEmoji + ' ' + statusText + '\n';

    if (stock.isLow) {
      t += '\n⚠️ ควรผลิตเพิ่ม ' + (stock.threshold - stock.stock) + ' ขวด';
    }

    return txtMsg(t);
  } catch (e) {
    return txtMsg('❌ ' + e.message);
  }
}

function botStockIn(cmd) {
  const match = cmd.match(/\/(stockin|เข้า)\s+(\d+)/i);
  if (!match) return txtMsg('❌ รูปแบบ: /stockin [จำนวน]');
  const qty = parseInt(match[2]);
  if (!qty || qty <= 0) return txtMsg('❌ จำนวนต้องมากกว่า 0');

  try {
    const result = addStockTransaction('IN', qty, 'LINE Bot', 'LINE');
    if (result.success) {
      const stock = getStockSummary();
      let t = '📥 รับเข้าสต็อก\n\n✅ +' + qty + ' ขวด\n';
      t += '💰 ฿' + (qty * PRICE_PER_UNIT).toLocaleString() + '\n';
      t += '📦 คงเหลือ: ' + stock.stock.toLocaleString() + '\n';
      if (stock.isLow) t += '\n⚠️ ยังต่ำกว่าเกณฑ์';
      else t += '\n✅ ปกติ';
      return txtMsg(t);
    }
    return txtMsg('❌ ' + (result.error || 'ไม่สำเร็จ'));
  } catch (e) {
    return txtMsg('❌ ' + e.message);
  }
}

function botStockOut(cmd) {
  const match = cmd.match(/\/(stockout|ออก|ขาย)\s+(\d+)/i);
  if (!match) return txtMsg('❌ รูปแบบ: /stockout [จำนวน]');
  const qty = parseInt(match[2]);
  if (!qty || qty <= 0) return txtMsg('❌ จำนวนต้องมากกว่า 0');

  const stockBefore = getStockSummary();
  if (stockBefore.success && stockBefore.stock < qty) {
    return txtMsg('❌ สต็อกไม่พอ!\nคงเหลือ: ' + stockBefore.stock + '\nต้องการ: ' + qty);
  }

  try {
    const result = addStockTransaction('OUT', qty, 'LINE Bot', 'LINE');
    if (result.success) {
      const stock = getStockSummary();
      let t = '📤 ขายออก\n\n✅ -' + qty + ' ขวด\n';
      t += '💰 ฿' + (qty * PRICE_PER_UNIT).toLocaleString() + '\n';
      t += '📦 คงเหลือ: ' + stock.stock.toLocaleString() + '\n';
      if (stock.isLow) t += '\n⚠️ ต่ำกว่าเกณฑ์!';
      return txtMsg(t);
    }
    return txtMsg('❌ ' + (result.error || 'ไม่สำเร็จ'));
  } catch (e) {
    return txtMsg('❌ ' + e.message);
  }
}

function botStockLog(cmd) {
  const match = cmd.match(/\/(stocklog|ประวัติสต็อก)(\s+(\d+))?/i);
  const limit = match && match[3] ? parseInt(match[3]) : 10;

  try {
    const result = getAllStockTransactions();
    if (!result.success || !result.data.length) return txtMsg('📋 ยังไม่มีประวัติ');

    const sorted = result.data.sort((a, b) =>
      new Date(b.timestamp) - new Date(a.timestamp)).slice(0, limit);

    let t = '📋 ประวัติสต็อก ' + sorted.length + ' รายการ\n\n';
    sorted.forEach((item, i) => {
      const emoji = item.type === 'IN' ? '📥' : '📤';
      const sign = item.type === 'IN' ? '+' : '-';
      const time = item.timestamp ?
        new Date(item.timestamp).toLocaleString('th-TH', {
          day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
        }) : '—';
      t += (i + 1) + '. ' + emoji + ' ' + sign + item.qty + ' | ฿' + item.total.toLocaleString() + '\n';
      t += '   ' + time + '\n';
    });

    return txtMsg(t);
  } catch (e) {
    return txtMsg('❌ ' + e.message);
  }
}

function botStockValue(salesData) {
  try {
    const stock = getStockSummary();
    if (!stock.success) return txtMsg('❌ ' + stock.error);

    const totalSales = salesData.reduce((s, r) => s + r.sold, 0);
    const avgSoldPerDay = salesData.length > 0 ? totalSales / salesData.length : 0;
    const daysLeft = avgSoldPerDay > 0 ? Math.floor(stock.stock / avgSoldPerDay) : Infinity;

    let t = '💰 มูลค่าสต็อก\n\n';
    t += '📦 ' + stock.stock.toLocaleString() + ' ขวด\n';
    t += '💵 ฿' + PRICE_PER_UNIT + '/ขวด\n';
    t += '💰 รวม: ฿' + stock.value.toLocaleString() + '\n\n';
    t += '📊 ขายเฉลี่ย: ' + avgSoldPerDay.toFixed(1) + '/วัน\n';
    t += '⏱ หมดใน: ' + (daysLeft === Infinity ? '∞' : daysLeft) + ' วัน\n\n';

    if (stock.stock > 100) t += '✅ เหลือเฟือ';
    else if (stock.stock > 40) t += '🟡 ปานกลาง';
    else t += '🔴 วิกฤต!';

    return txtMsg(t);
  } catch (e) {
    return txtMsg('❌ ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════
// 🌐 WEB DASHBOARD FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function getQuickStats() {
  try {
    const data = getAllSalesDataForBot();
    if (!data || !data.length) {
      return {
        revenue: '฿0', profit: '฿0', days: '0',
        todayRevenue: '฿0', todayProfit: '฿0',
        margin: '0%', totalSold: '0'
      };
    }

    const totalRev = data.reduce((s, r) => s + r.revenue, 0);
    const totalBal = data.reduce((s, r) => s + r.balance, 0);
    const totalSold = data.reduce((s, r) => s + r.sold, 0);
    const margin = totalRev > 0 ? ((totalBal / totalRev) * 100) : 0;

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const todayData = data.filter(r => r.date === today);
    const todayRev = todayData.reduce((s, r) => s + r.revenue, 0);
    const todayBal = todayData.reduce((s, r) => s + r.balance, 0);

    function fmtShort(n) {
      n = +n || 0;
      if (Math.abs(n) >= 1e6) return '฿' + (n / 1e6).toFixed(1) + 'M';
      if (Math.abs(n) >= 1e3) return '฿' + (n / 1e3).toFixed(0) + 'K';
      return '฿' + n.toLocaleString('th-TH', { maximumFractionDigits: 0 });
    }

    return {
      revenue: fmtShort(totalRev),
      profit: fmtShort(totalBal),
      days: String(data.length),
      todayRevenue: fmtShort(todayRev),
      todayProfit: fmtShort(todayBal),
      margin: margin.toFixed(1) + '%',
      totalSold: String(totalSold)
    };
  } catch (e) {
    return {
      revenue: 'Error', profit: 'Error', days: '0',
      todayRevenue: 'Error', todayProfit: 'Error',
      margin: '0%', totalSold: '0'
    };
  }
}

function executeWebCommand(cmd) {
  try {
    const data = getAllSalesDataForBot();
    const c = cmd.toLowerCase().trim();

    if (!data.length) {
      return '❌ ยังไม่มีข้อมูลในระบบ\n\n💡 ลองพิมพ์ "ขาย 10" ใน LINE เพื่อทดสอบ';
    }

    // Reports
    if (c === '/kpi' || c === '/dashboard') {
      const r = botKpiFlex(data);
      return r?.altText || r?.text || '📊 KPI Dashboard';
    }
    if (c === '/summary' || c === '/executive') {
      const r = botExecutiveSummary(data); return r?.text || '';
    }
    if (c === '/today') { const r = botDaily(data, 'today'); return r?.text || ''; }
    if (c === '/yesterday') { const r = botDaily(data, 'yesterday'); return r?.text || ''; }
    if (c === '/week') { const r = botWeekly(data); return r?.text || ''; }
    if (c === '/month') { const r = botMonthly(data); return r?.text || ''; }
    if (c === '/d2' || c === '/dashboard2') { const r = botDashboardV2(data); return r?.text || ''; }

    // Analysis
    if (c === '/deep' || c === '/analyze') { const r = botDeepAnalysis(data); return r?.text || ''; }
    if (c === '/forecast') { const r = botForecast(data); return r?.text || ''; }
    if (c === '/trend') { const r = botTrend(data); return r?.text || ''; }
    if (c === '/compare') { const r = botCompare(data); return r?.text || ''; }
    if (c === '/seasonality') { const r = botSeasonality(data); return r?.text || ''; }
    if (c === '/efficiency') { const r = botEfficiency(data); return r?.text || ''; }
    if (c === '/benchmark') { const r = botBenchmark(data); return r?.text || ''; }
    if (c === '/growth') { const r = botGrowth(data); return r?.text || ''; }
    if (c === '/correlation') { const r = botCorrelationAnalysis(data); return r?.text || ''; }
    if (c === '/whatif') { const r = botWhatIfAnalysis(data); return r?.text || ''; }
    if (c === '/montecarlo' || c === '/mc') { const r = botMonteCarlo(data); return r?.text || ''; }
    if (c === '/lifecycle') { const r = botLifecycle(data); return r?.text || ''; }
    if (c === '/pricing') { const r = botPricingOptimization(data); return r?.text || ''; }

    // Finance
    if (c === '/pl' || c === '/profitloss') { const r = botProfitLoss(data); return r?.text || ''; }
    if (c === '/cashflow') { const r = botCashflow(data); return r?.text || ''; }
    if (c === '/breakeven') { const r = botBreakeven(data); return r?.text || ''; }
    if (c === '/roi') { const r = botROI(data); return r?.text || ''; }
    if (c === '/ratio') { const r = botFinancialRatios(data); return r?.text || ''; }
    if (c === '/margin') { const r = botMargin(data); return r?.text || ''; }
    if (c === '/expense') { const r = botExpense(data); return r?.text || ''; }
    if (c === '/sales') { const r = botSales(data); return r?.text || ''; }
    if (c === '/conversion') { const r = botConversion(data); return r?.text || ''; }
    if (c === '/health') { const r = botHealthScore(data); return r?.text || ''; }

    // Risk & Planning
    if (c === '/risk') { const r = botRiskAnalysis(data); return r?.text || ''; }
    if (c === '/swot') { const r = botSWOT(data); return r?.text || ''; }
    if (c === '/goal') { const r = botGoalTracking(data); return r?.text || ''; }
    if (c === '/invoice') { const r = botOutstandingReport(data); return r?.text || ''; }
    if (c === '/projection') { const r = botProjection(data); return r?.text || ''; }
    if (c === '/alert') { const r = botAlertConfig(data); return r?.text || ''; }

    // Stock
    if (c === '/stock') { const r = botStockSummary(data); return r?.text || ''; }
    if (c.match(/^\/stockin\s+(\d+)$/)) { const r = botStockIn(c); return r?.text || ''; }
    if (c.match(/^\/stockout\s+(\d+)$/)) { const r = botStockOut(c); return r?.text || ''; }
    if (c === '/stocklog') { const r = botStockLog(c); return r?.text || ''; }
    if (c === '/stockvalue') { const r = botStockValue(data); return r?.text || ''; }

    // Export
    if (c === '/export csv') { const r = botExportData(data, c); return r?.text || ''; }
    if (c === '/export json') { const r = botExportData(data, c); return r?.text || ''; }

    // Others
    if (c === '/advice') { const r = botAdvice(data); return r?.text || ''; }
    if (c === '/stats') { const r = botStats(data); return r?.text || ''; }
    if (c.match(/^\/top\s+(\d+)$/)) { const r = botTop(data, 'top', c); return r?.text || ''; }
    if (c === '/top' || c === '/top5') { const r = botTop(data, 'top', '/top 5'); return r?.text || ''; }
    if (c.match(/^\/bottom\s+(\d+)$/)) { const r = botTop(data, 'bottom', c); return r?.text || ''; }
    if (c === '/bottom' || c === '/bottom5') { const r = botTop(data, 'bottom', '/bottom 5'); return r?.text || ''; }
    if (c.match(/^\/latest\s+(\d+)$/)) { const r = botLatest(data, c); return r?.text || ''; }
    if (c === '/latest' || c === '/latest5') { const r = botLatest(data, '/latest 5'); return r?.text || ''; }
    if (c === '/latest10') { const r = botLatest(data, '/latest 10'); return r?.text || ''; }
    if (c.match(/^\/day\s+(\S+)$/)) { const r = botDayOfWeek(data, c); return r?.text || ''; }
    if (c.match(/^\/search\s+(.+)$/)) { const r = botSearch(data, c); return r?.text || ''; }
    if (c === '/help') { const r = botHelp(); return r?.text || ''; }

    return '❌ ไม่พบคำสั่ง: ' + cmd + '\n\n📋 พิมพ์ /help';

  } catch (e) {
    console.error('executeWebCommand error:', e);
    return '❌ เกิดข้อผิดพลาด: ' + e.message;
  }
}

// ═══════════════════════════════════════════════════════════════
// 💡 EXTENDED BOT FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function botExecutiveSummary(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const exp = data.reduce((s, r) => s + r.expense, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const conv = sold > 0 ? ((sold - pend) / sold * 100) : 0;

  const best = data.reduce((mx, r) => r.balance > mx.balance ? r : mx, data[0]);
  const worst = data.reduce((mn, r) => r.balance < mn.balance ? r : mn, data[0]);
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length / 2);
  const fBal = sorted.slice(0, mid).reduce((s, r) => s + r.balance, 0);
  const sBal = sorted.slice(mid).reduce((s, r) => s + r.balance, 0);
  const trend = sBal > fBal ? '📈 เพิ่มขึ้น' : sBal < fBal ? '📉 ลดลง' : '➡️ คงที่';

  let t = '📊 EXECUTIVE SUMMARY\n';
  t += Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '\n\n';
  t += '💰 รายได้: ' + fmtBot(rev) + '\n';
  t += '💸 ค่าใช้จ่าย: ' + fmtBot(exp) + '\n';
  t += '📈 กำไร: ' + fmtBot(bal) + ' (' + margin.toFixed(1) + '%)\n';
  t += '📅 ' + data.length + ' วัน | 🛒 ' + sold + ' ชิ้น\n\n';
  t += '🏆 ดีสุด: ' + best.date + ' (' + fmtBot(best.balance) + ')\n';
  t += '⚠️ แย่สุด: ' + worst.date + ' (' + fmtBot(worst.balance) + ')\n';
  t += '📈 แนวโน้ม: ' + trend + '\n';
  t += '✅ Conversion: ' + conv.toFixed(1) + '%\n\n';

  if (margin < 20) t += '🔴 เร่งด่วน: ปรับโครงสร้างต้นทุน\n';
  if (trend.includes('ลดลง')) t += '⚠️ แนวโน้มลดลง ต้องรีบแก้ไข\n';
  if (conv < 60) t += '⚠️ ปรับปรุงการเก็บเงิน\n';

  return txtMsg(t);
}

function botEfficiency(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);

  const revPerDay = totalRev / data.length;
  const expRatio = totalRev > 0 ? (totalExp / totalRev * 100) : 0;
  const revPerUnit = totalSold > 0 ? (totalRev / totalSold) : 0;
  const collectionRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;
  const pipeFee = data.reduce((s, r) => s + r.pipeFee, 0);
  const pipeEff = totalRev > 0 ? (pipeFee / totalRev * 100) : 0;

  let score = 0;
  if (expRatio < 30) score += 3;
  else if (expRatio < 50) score += 2;
  else score += 1;
  if (collectionRate > 80) score += 3;
  else if (collectionRate > 60) score += 2;
  else score += 1;
  if (pipeEff < 20) score += 3;
  else if (pipeEff < 40) score += 2;
  else score += 1;

  const lvl = score >= 7 ? '🟢 ดีเยี่ยม' : score >= 5 ? '🟡 ปานกลาง' : '🔴 ต้องปรับปรุง';

  let t = '⚡ EFFICIENCY\n\n📊 Score: ' + score + '/9 — ' + lvl + '\n\n';
  t += '• รายได้/วัน: ' + fmtBot(revPerDay) + '\n';
  t += '• Expense Ratio: ' + expRatio.toFixed(1) + '%\n';
  t += '• Revenue/Unit: ' + fmtBot(revPerUnit) + '\n';
  t += '• Collection: ' + collectionRate.toFixed(1) + '%\n';
  t += '• Pipe Cost: ' + pipeEff.toFixed(1) + '%\n';
  return txtMsg(t);
}

function botBenchmark(data) {
  const benchmarks = {
    profitMargin: { target: 30, good: 25, average: 15 },
    conversionRate: { target: 80, good: 70, average: 50 },
    expenseRatio: { target: 25, good: 35, average: 50 },
    dailyRevenue: { target: 1500, good: 1000, average: 500 }
  };

  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);

  const margin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;
  const expRatio = totalRev > 0 ? (totalExp / totalRev * 100) : 0;
  const dailyRev = data.length > 0 ? totalRev / data.length : 0;
  const conv = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;

  function score(val, bm, reverse) {
    if (reverse) {
      if (val <= bm.target) return '🟢 ดีเยี่ยม';
      if (val <= bm.good) return '🟡 ดี';
      if (val <= bm.average) return '🟠 ปานกลาง';
      return '🔴 สูงเกินไป';
    }
    if (val >= bm.target) return '🟢 ดีเยี่ยม';
    if (val >= bm.good) return '🟡 ดี';
    if (val >= bm.average) return '🟠 ปานกลาง';
    return '🔴 ต่ำ';
  }

  let t = '📊 BENCHMARK\n\n';
  t += 'Margin: ' + margin.toFixed(1) + '% → ' + score(margin, benchmarks.profitMargin) + '\n';
  t += 'Expense: ' + expRatio.toFixed(1) + '% → ' + score(expRatio, benchmarks.expenseRatio, true) + '\n';
  t += 'Conversion: ' + conv.toFixed(1) + '% → ' + score(conv, benchmarks.conversionRate) + '\n';
  t += 'Revenue/Day: ' + fmtBot(dailyRev) + ' → ' + score(dailyRev, benchmarks.dailyRevenue) + '\n';
  return txtMsg(t);
}

function botGrowth(data) {
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const monthlyData = {};
  sorted.forEach(r => {
    if (!r.date) return;
    const ym = r.date.slice(0, 7);
    if (!monthlyData[ym]) monthlyData[ym] = { revenue: 0, balance: 0, sold: 0, days: 0 };
    monthlyData[ym].revenue += r.revenue;
    monthlyData[ym].balance += r.balance;
    monthlyData[ym].sold += r.sold;
    monthlyData[ym].days++;
  });

  const months = Object.entries(monthlyData).sort();
  let t = '📈 GROWTH\n\n';
  let prevRev = 0;

  months.forEach(([ym, m]) => {
    const growth = prevRev > 0 ? ((m.revenue - prevRev) / prevRev * 100).toFixed(1) : '—';
    const gStr = growth === '—' ? '   —   ' : (growth >= 0 ? '📈 +' + growth + '%' : '📉 ' + growth + '%');
    t += ym + ' | ' + fmtBot(m.revenue) + ' | ' + gStr + '\n';
    prevRev = m.revenue;
  });

  return txtMsg(t);
}

function botSeasonality(data) {
  const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  const monthlyPattern = {};
  const dailyPattern = Array(7).fill().map(() => ({ count: 0, total: 0 }));

  data.forEach(r => {
    if (!r.date) return;
    const d = new Date(r.date);
    const month = d.getMonth() + 1;
    const dow = d.getDay();
    if (!monthlyPattern[month]) monthlyPattern[month] = { count: 0, total: 0 };
    monthlyPattern[month].count++;
    monthlyPattern[month].total += r.balance;
    dailyPattern[dow].count++;
    dailyPattern[dow].total += r.balance;
  });

  const avgAll = data.reduce((s, r) => s + r.balance, 0) / data.length;

  let t = '🌤️ SEASONALITY\n\n📅 รายวัน:\n';
  dailyPattern.forEach((s, i) => {
    const avg = s.count > 0 ? s.total / s.count : 0;
    const diff = avgAll > 0 ? ((avg - avgAll) / avgAll * 100).toFixed(0) : 0;
    const ind = avg > avgAll ? '📈' : avg < avgAll ? '📉' : '➡️';
    t += ind + ' ' + dayNames[i] + ': ' + fmtBot(avg) + ' (' + (diff >= 0 ? '+' : '') + diff + '%)\n';
  });

  return txtMsg(t);
}

function botROI(data) {
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  const roi = totalExp > 0 ? (totalBal / totalExp * 100) : 0;
  const avgInv = data.length > 0 ? totalExp / data.length : 0;
  const avgRet = data.length > 0 ? totalBal / data.length : 0;
  const dailyROI = avgInv > 0 ? (avgRet / avgInv * 100) : 0;
  const paybackDays = avgRet > 0 ? avgInv / avgRet : Infinity;

  let t = '💰 ROI ANALYSIS\n\n';
  t += '• ลงทุนรวม: ' + fmtBot(totalExp) + '\n';
  t += '• กำไรสุทธิ: ' + fmtBot(totalBal) + '\n';
  t += '• ROI: ' + roi.toFixed(1) + '%\n\n';
  t += '• ลงทุน/วัน: ' + fmtBot(avgInv) + '\n';
  t += '• กำไร/วัน: ' + fmtBot(avgRet) + '\n';
  t += '• Daily ROI: ' + dailyROI.toFixed(1) + '%\n';
  t += '• คืนทุน: ' + (paybackDays === Infinity ? '∞' : paybackDays.toFixed(1)) + ' วัน\n';

  return txtMsg(t);
}

function botProfitLoss(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  const totalClr = data.reduce((s, r) => s + r.cleared, 0);
  const pipe = data.reduce((s, r) => s + r.pipeFee, 0);
  const share = data.reduce((s, r) => s + r.shareFee, 0);
  const other = data.reduce((s, r) => s + r.otherFee, 0);
  const save = data.reduce((s, r) => s + r.saveFee, 0);
  const totalExp = pipe + share + other + save;
  const netProfit = totalRev - totalExp;
  const margin = totalRev > 0 ? (netProfit / totalRev * 100) : 0;

  let t = '📊 PROFIT & LOSS\n\n';
  t += '📥 รายได้: ' + fmtBot(totalRev) + '\n';
  t += '📤 ค่าใช้จ่าย:\n';
  t += '  • Pipe: ' + fmtBot(pipe) + '\n';
  t += '  • Share: ' + fmtBot(share) + '\n';
  t += '  • Other: ' + fmtBot(other) + '\n';
  t += '  • Save: ' + fmtBot(save) + '\n';
  t += '  รวม: ' + fmtBot(totalExp) + '\n\n';
  t += '📈 กำไรสุทธิ: ' + fmtBot(netProfit) + '\n';
  t += '🎯 Margin: ' + margin.toFixed(1) + '%\n\n';
  t += '📅 ' + data.length + ' วัน | 🛒 ' + totalSold + ' | ⏳ ' + totalPend + ' | ✅ ' + totalClr;

  return txtMsg(t);
}

function botCashflow(data) {
  let totalInflow = 0, totalOutflow = 0;
  data.forEach(r => {
    totalInflow += r.revenue;
    totalOutflow += r.expense;
  });
  const netCashflow = totalInflow - totalOutflow;
  const ratio = totalOutflow > 0 ? (totalInflow / totalOutflow) : 0;
  const positiveDays = data.filter(r => r.balance >= 0).length;
  const negativeDays = data.length - positiveDays;

  let t = '💵 CASH FLOW\n\n';
  t += '📥 เข้า: ' + fmtBot(totalInflow) + '\n';
  t += '📤 ออก: ' + fmtBot(totalOutflow) + '\n';
  t += '📊 สุทธิ: ' + fmtBot(netCashflow) + '\n';
  t += '📈 Ratio: ' + ratio.toFixed(2) + 'x\n\n';
  t += '✅ วันบวก: ' + positiveDays + '\n';
  t += '❌ วันลบ: ' + negativeDays + '\n\n';
  t += ratio < 1.2 ? '🔴 เสี่ยง! เพิ่มรายได้' : ratio < 2 ? '🟡 พอใช้' : '🟢 ดี';

  return txtMsg(t);
}

function botBreakeven(data) {
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const fixedCosts = totalExp * 0.6 / data.length;
  const variableCost = (totalExp * 0.4) / (totalSold || 1);
  const contribution = PRICE_PER_UNIT - variableCost;
  const bepUnits = contribution > 0 ? Math.ceil(fixedCosts / contribution) : Infinity;
  const avgSold = totalSold / data.length;

  let t = '🎯 BREAK-EVEN\n\n';
  t += '• ต้นทุนคงที่/วัน: ' + fmtBot(fixedCosts) + '\n';
  t += '• ต้นทุนผันแปร/หน่วย: ' + fmtBot(variableCost) + '\n';
  t += '• Contribution: ' + fmtBot(contribution) + '\n\n';
  t += '🎯 จุดคุ้มทุน: ' + (bepUnits === Infinity ? '∞' : bepUnits) + ' ชิ้น/วัน\n';
  t += '📊 ขายจริง: ' + avgSold.toFixed(1) + ' ชิ้น/วัน\n\n';

  if (avgSold >= bepUnits) {
    t += '✅ เกินจุดคุ้มทุน';
  } else if (bepUnits !== Infinity) {
    t += '❌ ต้องเพิ่ม ' + (bepUnits - avgSold).toFixed(0) + ' ชิ้น/วัน';
  }

  return txtMsg(t);
}

function botFinancialRatios(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const exp = data.reduce((s, r) => s + r.expense, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const expRatio = rev > 0 ? (exp / rev * 100) : 0;

  let t = '📊 FINANCIAL RATIOS\n\n';
  t += '• Gross Margin: ' + margin.toFixed(1) + '%\n';
  t += '• Expense Ratio: ' + expRatio.toFixed(1) + '%\n';
  t += '• Net Margin: ' + margin.toFixed(1) + '%\n';
  t += '• Revenue/Day: ' + fmtBot(rev / data.length) + '\n';
  t += '• Profit/Day: ' + fmtBot(bal / data.length) + '\n';

  return txtMsg(t);
}

function botAlertConfig(data) {
  const avgBal = data.reduce((s, r) => s + r.balance, 0) / data.length;
  const stdDev = Math.sqrt(data.reduce((s, r) => s + Math.pow(r.balance - avgBal, 2), 0) / data.length);

  let t = '⚠️ ALERT CONFIG\n\n';
  t += '🔴 กำไรต่ำผิดปกติ: < ' + fmtBot(avgBal - 2 * stdDev) + '\n';
  t += '🟢 กำไรสูงผิดปกติ: > ' + fmtBot(avgBal + 2 * stdDev) + '\n';
  t += '🔴 วันที่ขาดทุนบ่อย\n';
  t += '🔴 Pending สูง\n';

  return txtMsg(t);
}

function botDashboardV2(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const exp = data.reduce((s, r) => s + r.expense, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const margin = rev > 0 ? (bal / rev * 100).toFixed(1) : 0;
  const conv = sold > 0 ? ((sold - pend) / sold * 100).toFixed(1) : 0;

  let t = '📊 DASHBOARD V2\n═══════════════\n\n';
  t += '💰 ' + fmtBot(rev).padStart(10) + '\n';
  t += '💸 ' + fmtBot(exp).padStart(10) + '\n';
  t += '📈 ' + fmtBot(bal).padStart(10) + '\n';
  t += '───────────────\n';
  t += '🎯 ' + margin + '%\n';
  t += '✅ ' + conv + '%\n';
  t += '🛒 ' + sold + ' | ⏳ ' + pend + '\n';
  t += '📅 ' + data.length + ' วัน\n';

  return txtMsg(t);
}

function botSWOT(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const conv = sold > 0 ? ((sold - pend) / sold * 100) : 0;

  const strengths = [];
  if (margin >= 30) strengths.push('💪 Margin สูง');
  if (conv >= 70) strengths.push('💪 Collection ดี');
  if (data.filter(r => r.balance >= 0).length / data.length > 0.8) strengths.push('💪 กำไรสม่ำเสมอ');
  if (!strengths.length) strengths.push('💪 มีระบบติดตาม');

  const weaknesses = [];
  if (margin < 20) weaknesses.push('🔻 Margin ต่ำ');
  if (conv < 60) weaknesses.push('🔻 Pending สูง');
  if (!weaknesses.length) weaknesses.push('🔻 ยังไม่พบจุดอ่อน');

  let t = '📊 SWOT\n\n💪 Strengths:\n';
  strengths.forEach(s => t += '  ' + s + '\n');
  t += '\n🔻 Weaknesses:\n';
  weaknesses.forEach(w => t += '  ' + w + '\n');
  t += '\n🚀 Opportunities:\n  🚀 ขยายออนไลน์\n  🚀 เพิ่มช่องทาง\n';
  t += '\n⚠️ Threats:\n  ⚠️ คู่แข่ง\n  ⚠️ ต้นทุน';

  return txtMsg(t);
}

function botGoalTracking(data) {
  const goals = { dailyRevenue: 1000, monthlyProfit: 20000, profitMargin: 30, conversionRate: 70 };
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const conv = sold > 0 ? ((sold - pend) / sold * 100) : 0;

  const avgDaily = data.length ? rev / data.length : 0;
  const avgMonthly = data.length ? (bal / data.length * 30) : 0;

  function bar(cur, target) {
    const pct = Math.min(100, (cur / target * 100));
    const filled = Math.floor(pct / 10);
    return '█'.repeat(filled) + '░'.repeat(10 - filled) + ' ' + pct.toFixed(0) + '%';
  }

  let t = '🎯 GOAL TRACKING\n\n';
  t += 'Revenue/Day ≥ ' + fmtBot(goals.dailyRevenue) + '\n';
  t += bar(avgDaily, goals.dailyRevenue) + '\n\n';
  t += 'Profit/Month ≥ ' + fmtBot(goals.monthlyProfit) + '\n';
  t += bar(avgMonthly, goals.monthlyProfit) + '\n\n';
  t += 'Margin ≥ ' + goals.profitMargin + '%\n';
  t += bar(margin, goals.profitMargin) + '\n\n';
  t += 'Conversion ≥ ' + goals.conversionRate + '%\n';
  t += bar(conv, goals.conversionRate) + '\n';

  return txtMsg(t);
}

function botOutstandingReport(data) {
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalOutstanding = totalPend * PRICE_PER_UNIT;
  const rate = totalSold > 0 ? (totalPend / totalSold * 100) : 0;

  let t = '📋 OUTSTANDING\n\n';
  t += '⏳ ค้างชำระ: ' + totalPend + ' ชิ้น\n';
  t += '💰 มูลค่า: ' + fmtBot(totalOutstanding) + '\n';
  t += '📊 อัตรา: ' + rate.toFixed(1) + '%\n\n';
  t += rate > 30 ? '🔴 สูง! ควรส่งใบแจ้งหนี้' : '✅ ปกติ';

  return txtMsg(t);
}

function botProjection(data) {
  const avgRev = data.reduce((s, r) => s + r.revenue, 0) / data.length;
  const avgExp = data.reduce((s, r) => s + r.expense, 0) / data.length;
  const avgBal = avgRev - avgExp;

  const scenarios = [
    { name: '🔴 Conservative', f: 0.8 },
    { name: '🟡 Base', f: 1.0 },
    { name: '🟢 Optimistic', f: 1.2 }
  ];

  let t = '📊 PROJECTION\n\n';
  scenarios.forEach(sc => {
    t += sc.name + ':\n';
    [7, 30, 90].forEach(d => {
      t += '  ' + d + 'd: ' + fmtBot(avgRev * sc.f * d) + ' | ' + fmtBot(avgBal * sc.f * d) + '\n';
    });
    t += '\n';
  });

  return txtMsg(t);
}

function botExportData(data, cmd) {
  const match = cmd.match(/\/(export|ส่งออก)\s+(csv|json)/i);
  if (!match) return txtMsg('❌ รูปแบบ: /export csv หรือ /export json');
  const format = match[2].toLowerCase();

  if (format === 'csv') {
    const header = 'date,sold,pending,cleared,revenue,pipeFee,shareFee,otherFee,saveFee,expense,balance';
    const rows = data.map(r => [r.date, r.sold, r.pending, r.cleared, r.revenue,
      r.pipeFee, r.shareFee, r.otherFee, r.saveFee, r.expense, r.balance].join(','));
    return txtMsg('📥 EXPORT CSV\n\n' + data.length + ' รายการ\n\n' + header + '\n' + rows.slice(0, 5).join('\n') + '\n\n💡 ไฟล์เต็มส่งใน Admin');
  }

  if (format === 'json') {
    const json = JSON.stringify(data.slice(0, 10), null, 2);
    return txtMsg('📥 EXPORT JSON\n\n' + data.length + ' รายการ\n\n' + (json.length > 500 ? json.substring(0, 500) + '...' : json));
  }

  return txtMsg('❌ รูปแบบไม่ถูกต้อง');
}

function botSearch(data, cmd) {
  const q = cmd.replace(/^\/(search|ค้นหา)\s+/i, '').trim();
  if (!q) return txtMsg('🔍 ระบุคำค้นหา เช่น /search 2024-01');

  const results = data.filter(r => Object.values(r).some(v =>
    String(v).toLowerCase().includes(q.toLowerCase()))).slice(0, 10);

  if (!results.length) return txtMsg('🔍 "' + q + '"\n❌ ไม่พบข้อมูล');

  let t = '🔍 "' + q + '" — ' + results.length + ' รายการ\n\n';
  results.forEach((r, i) => {
    const mg = r.revenue > 0 ? (r.balance / r.revenue * 100).toFixed(1) : 0;
    t += (i + 1) + '. 📅 ' + r.date + '\n   ขาย:' + r.sold + ' | ' + fmtBot(r.balance) + ' (' + mg + '%)\n';
  });

  return txtMsg(t);
}

function botDayOfWeek(data, cmd) {
  const dayMap = {
    'อาทิตย์': 0, 'sunday': 0, 'จันทร์': 1, 'monday': 1, 'อังคาร': 2, 'tuesday': 2,
    'พุธ': 3, 'wednesday': 3, 'พฤหัส': 4, 'thursday': 4, 'พฤหัสบดี': 4,
    'ศุกร์': 5, 'friday': 5, 'เสาร์': 6, 'saturday': 6
  };

  const match = cmd.match(/\/day\s+(\S+)/i);
  if (!match) return txtMsg('❌ ระบุวัน เช่น /day จันทร์');

  const targetDay = dayMap[match[1].toLowerCase()];
  if (targetDay === undefined) return txtMsg('❌ ไม่พบวัน');

  const dd = data.filter(r => {
    if (!r.date) return false;
    return new Date(r.date).getDay() === targetDay;
  });

  const tn = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  if (!dd.length) return txtMsg('📅 ' + tn[targetDay] + '\n❌ ไม่มีข้อมูล');

  const rev = dd.reduce((s, r) => s + r.revenue, 0);
  const bal = dd.reduce((s, r) => s + r.balance, 0);
  const sold = dd.reduce((s, r) => s + r.sold, 0);

  return txtMsg('📅 วัน' + tn[targetDay] + '\n\n📊 ' + dd.length + ' วัน\n🛒 ' + sold + '\n💰 ' + fmtBot(rev) + '\n📈 ' + fmtBot(bal) + '\n📊 เฉลี่ย/วัน: ' + fmtBot(dd.length ? bal / dd.length : 0));
}

function botLatest(data, cmd) {
  const m = cmd.match(/\d+/);
  const limit = m ? parseInt(m[0]) : 5;
  const sorted = data.slice().sort((a, b) => b.date.localeCompare(a.date)).slice(0, limit);

  let t = '📋 ' + limit + ' รายการล่าสุด\n\n';
  sorted.forEach((r, i) => {
    const mg = r.revenue > 0 ? (r.balance / r.revenue * 100).toFixed(1) : 0;
    t += (i + 1) + '. 📅 ' + r.date + ' — ' + fmtBot(r.balance) + ' (' + mg + '%)\n';
  });

  return txtMsg(t);
}

function botTop(data, type, cmd) {
  const limitMatch = cmd.match(/\d+/);
  const limit = limitMatch ? parseInt(limitMatch[0]) : 5;
  const sorted = data.slice().sort((a, b) =>
    type === 'top' ? b.balance - a.balance : a.balance - b.balance);
  const topN = sorted.slice(0, limit);
  const emoji = type === 'top' ? '🏆' : '⚠️';

  let t = emoji + ' ' + type.toUpperCase() + ' ' + limit + '\n\n';
  topN.forEach((r, i) => {
    const mg = r.revenue > 0 ? (r.balance / r.revenue * 100).toFixed(1) : 0;
    t += (i + 1) + '. ' + r.date + ' — ' + fmtBot(r.balance) + ' (' + mg + '%)\n';
  });

  return txtMsg(t);
}

// ═══════════════════════════════════════════════════════════════
// 💬 LINE MESSAGE HANDLER
// ═══════════════════════════════════════════════════════════════

function handleLineMessageEnhanced(event) {
  const msg = event.message.text.trim();

  // 1. ลอง parse เป็น quick sale ก่อน
  if (recordQuickSale(msg, event)) return;

  // 2. ถ้าขึ้นต้นด้วย / หรือ !
  if (msg.startsWith('/') || msg.startsWith('!')) {
    try {
      const response = processBotCommand(msg);
      replyToLine(event.replyToken, response);
    } catch (e) {
      console.error('Command error:', e.message);
      replyToLine(event.replyToken, txtMsg('❌ เกิดข้อผิดพลาด: ' + e.message));
    }
    return;
  }

  // 3. Default reply
  replyToLine(event.replyToken, txtMsg(
    '🤖 สวัสดี! ผมคือ Sales Analytics Bot v' + SYSTEM_VERSION + '\n\n' +
    '📝 บันทึกข้อมูล: พิมพ์ "ขาย 10"\n' +
    '📊 ดูรายงาน: พิมพ์ /help\n\n' +
    '💡 ตัวอย่าง:\n' +
    '• ขาย 10\n' +
    '• ขาย 10 p50 s20\n' +
    '• /kpi\n' +
    '• /today'
  ));
}

function handlePostbackEvent(event) {
  const data = event.postback.data;
  const action = data.split('=')[1];
  const salesData = getAllSalesDataForBot();

  const actions = {
    'kpi': () => botKpiFlex(salesData),
    'today': () => botDaily(salesData, 'today'),
    'week': () => botWeekly(salesData),
    'month': () => botMonthly(salesData),
    'forecast': () => botForecast(salesData),
    'correlation': () => botCorrelationAnalysis(salesData),
    'whatif': () => botWhatIfAnalysis(salesData),
    'health': () => botHealthScore(salesData),
    'stock': () => botStockSummary(salesData),
    'pricing': () => botPricingOptimization(salesData)
  };

  const response = actions[action] ? actions[action]() : txtMsg('คำสั่งไม่ถูกต้อง');
  replyToLine(event.replyToken, response);
}

// ═══════════════════════════════════════════════════════════════
// ⏰ TRIGGER FUNCTIONS (รันอัตโนมัติตามเวลา)
// ═══════════════════════════════════════════════════════════════

/**
 * Daily Summary — รันทุกวัน 08:00
 */
function dailySummaryTrigger() {
  try {
    console.log('=== Daily Summary Trigger ===');
    const data = getAllSalesDataForBot();
    if (!data.length) {
      console.log('No data, skipping');
      return;
    }

    const result = sendKpiReportToLine();
    console.log('KPI sent:', result.success ? 'OK' : result.error);
  } catch (e) {
    console.error('dailySummaryTrigger error:', e.message);
  }
}

/**
 * Stock Alert — รันทุก 6 ชั่วโมง
 */
function checkStockAlertTrigger() {
  try {
    console.log('=== Stock Alert Trigger ===');
    if (!CONFIG.autoNotifyLowStock) return;

    const stock = getStockSummary();
    if (stock.success && stock.isLow) {
      sendLineNotification({
        type: 'stock_low',
        message: 'สต็อกต่ำ!\n\n📦 คงเหลือ: ' + stock.stock + ' ขวด\n🎯 เกณฑ์: ' + stock.threshold + '\n\n⚠️ ควรผลิตเพิ่ม ' + (stock.threshold - stock.stock) + ' ขวด'
      });
      console.log('Low stock alert sent');
    }
  } catch (e) {
    console.error('checkStockAlertTrigger error:', e.message);
  }
}

/**
 * Cleanup Old Reports — รันทุกวันอาทิตย์ 03:00
 */
function cleanupOldReportTrigger() {
  try {
    console.log('=== Cleanup Trigger ===');
    const ss = _ss();
    const old = ss.getSheetByName(REPORT_SHEET);
    if (old) {
      ss.deleteSheet(old);
      console.log('Old report deleted');
    }
  } catch (e) {
    console.error('cleanupOldReportTrigger error:', e.message);
  }
}

// ═══════════════════════════════════════════════════════════════
// 🧪 DIAGNOSTIC & TEST FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * ทดสอบการเข้าถึง Sheet
 * รันใน Apps Script Editor เพื่อ debug
 */
function testSheetAccess() {
  console.log('╔═══════════════════════════════════════╗');
  console.log('║  🧪 DIAGNOSTIC: Sheet Access Check  ║');
  console.log('╚═══════════════════════════════════════╝');

  console.log('\n📋 SPREADSHEET_ID:');
  console.log('   ' + SPREADSHEET_ID);

  console.log('\n📋 เปิด Spreadsheet:');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('   ✅ "' + ss.getName() + '"');
    console.log('   Sheets:');
    ss.getSheets().forEach(s => {
      console.log('      • "' + s.getName() + '" (' + s.getLastRow() + ' rows)');
    });
  } catch (e) {
    console.log('   ❌ ' + e.message);
    return;
  }

  console.log('\n📋 หา Sheet "' + SHEET_NAME + '":');
  try {
    const sheet = _sheet();
    console.log('   ✅ Last Row: ' + sheet.getLastRow());
    console.log('   ✅ Last Col: ' + sheet.getLastColumn());
  } catch (e) {
    console.log('   ❌ ' + e.message);
  }

  console.log('\n📋 ทดสอบ appendRow:');
  try {
    const sheet = _sheet();
    const testDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.appendRow([new Date(testDate), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, new Date()]);
    const lastRow = sheet.getLastRow();
    console.log('   ✅ Appended to row ' + lastRow);
    sheet.deleteRow(lastRow);
    console.log('   ✅ Cleaned up');
  } catch (e) {
    console.log('   ❌ ' + e.message);
  }

  console.log('\n📋 LINE Token:');
  if (LINE_CHANNEL_ACCESS_TOKEN && LINE_CHANNEL_ACCESS_TOKEN.length > 50) {
    console.log('   ✅ Length: ' + LINE_CHANNEL_ACCESS_TOKEN.length);
  } else {
    console.log('   ❌ Invalid');
  }

  console.log('\n📋 Triggers:');
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length) {
    triggers.forEach(t => console.log('   • ' + t.getHandlerFunction()));
  } else {
    console.log('   ⚠️ None — run setupSystem()');
  }

  console.log('\n✅ Diagnostic complete\n');
}

/**
 * ทดสอบ quick sale parser
 */
function testRecordQuickSale() {
  const samples = [
    'ขาย 10',
    'ขาย 10 p50',
    'ขาย 10 p50 s20 o10 sa5',
    'วันนี้ ขาย 10',
    'ขาย 10 pe5 cl5',
    '/sell 10'
  ];

  console.log('=== Test Quick Sale Parser ===');
  samples.forEach(s => {
    const c = s.replace(/^\//, '').trim();
    const isSale = c.match(/^(sell|ขาย|sale|add|บันทึก)/i) ||
                   c.match(/^\d{1,2}\/\d{1,2}/) ||
                   c.match(/^(วันนี้|พรุ่งนี้|เมื่อวาน)/i);

    const soldMatch = c.match(/(?:^|\s)(\d+)(?:\s|$)/);
    const pipeMatch = c.match(/\bp(\d+)/i);
    const shareMatch = c.match(/\bs(\d+)/i);

    console.log('"' + s + '" →');
    console.log('   isSale: ' + !!isSale);
    console.log('   sold: ' + (soldMatch ? soldMatch[1] : 0));
    console.log('   pipe: ' + (pipeMatch ? pipeMatch[1] : 0));
    console.log('   share: ' + (shareMatch ? shareMatch[1] : 0));
    console.log('');
  });
}

/**
 * ทดสอบส่ง LINE
 */
function testLineSend() {
  const result = sendCustomLineMessage(
    '🧪 Test Message\n\n' +
    'SSKYMT Analytics v' + SYSTEM_VERSION + '\n' +
    '📅 ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '\n\n' +
    'ถ้าได้รับข้อความนี้ = LINE ทำงานปกติ ✅'
  );
  console.log('LINE Test:', JSON.stringify(result));
}

// ═══════════════════════════════════════════════════════════════
// 📖 ONOPEN MENU (เพิ่มเมนูใน Google Sheets)
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('⚡ SSKYMT')
      .addItem('🚀 Setup System', 'setupSystem')
      .addItem('🔍 Verify System', 'verifySystem')
      .addSeparator()
      .addItem('📊 Generate Report', 'generateAnalyticsReport')
      .addItem('📈 Send KPI to LINE', 'sendKpiReportToLine')
      .addItem('🧠 Send Deep Analysis', 'sendDeepAnalysisToLine')
      .addSeparator()
      .addItem('🧪 Test Sheet Access', 'testSheetAccess')
      .addItem('🧪 Test LINE Send', 'testLineSend')
      .addSeparator()
      .addItem('📦 Seed Demo Data', 'seedDemoData')
      .addItem('🔄 Reset System', 'resetSystem')
      .addToUi();
  } catch (e) {
    console.error('onOpen error:', e.message);
  }
}

// ═══════════════════════════════════════════════════════════════
// ✅ END OF FILE — SSKYMT Analytics v5.0.0
// ═══════════════════════════════════════════════════════════════