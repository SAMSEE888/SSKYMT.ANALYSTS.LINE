// ============================================================
// Code.gs V2 — Sales Analytics Dashboard + Auto Report Sheet
// Google Apps Script | Production Ready
// ============================================================

const SPREADSHEET_ID  = '11vhg37MbHRm53SSEHLsCI3EBXx5_meXVvlRuqhFteaY';
const SHEET_NAME      = 'SaleForm';
const REPORT_SHEET    = 'SSKYMT-Report';
const PRICE_PER_UNIT  = 40;

// ─── LINE OA Configuration ─────────────────────────────────
const LINE_CHANNEL_ACCESS_TOKEN = 'wRfgbBD7DiTnUB760k4508c5vqt+ozpfrXa1pijtzai1hBVWI56nJghnN9737LsVonwsxGvFe0RcZH3sZO3ITACtcSyHNS0RhVohrgQXzdTLEf68e+xA9XfQ5ReZp7K/9XV5UOFJFCe+pjwQi1vFlwdB04t89/1O/w1cDnyilFU='; // ใส่ Token จริง
const LINE_API_URL = 'https://api.line.me/v2/bot/message/push';
const LINE_ADMIN_USER_ID = 'Ue3897727f45a57d2ae308be17f17c215'; // ใส่ User ID จริง

// ─── Web App Entry Point ─────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SSKYMT Sales Analytics')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

// ─── Helpers ──────────────────────────────────────────────────
function _ss()    { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function _sheet() {
  const s = _ss().getSheetByName(SHEET_NAME);
  if (!s) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
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
  const revenue  = (sold + cleared - pending) * PRICE_PER_UNIT;
  const expense  = pipeFee + shareFee + otherFee + saveFee;
  const balance  = revenue - expense;

  const dateRaw = row[0];
  let dateStr   = '';
  if (dateRaw instanceof Date) dateStr = _fmt(dateRaw, 'yyyy-MM-dd');
  else dateStr = String(dateRaw || '');

  const tsRaw = row[11];
  let tsStr   = '';
  if (tsRaw instanceof Date) tsStr = _fmt(tsRaw, 'yyyy-MM-dd HH:mm:ss');
  else tsStr = String(tsRaw || '');

  return { rowIndex, date: dateStr, sold, pending, cleared,
           revenue, pipeFee, shareFee, otherFee, saveFee,
           expense, balance, timestamp: tsStr };
}

// ─── Get All Data ─────────────────────────────────────────────
function getSalesData() {
  try {
    const sheet   = _sheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const data   = values.map((r, i) => _parseRow(r, i + 2)).filter(Boolean);
    return { success: true, data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Get Filtered Data for Insights ───────────────────────────
function getFilteredInsightsData(filterParams) {
  try {
    const allData = getSalesData();
    if (!allData.success) return allData;
    
    let data = allData.data;
    const { dateFrom, dateTo, month, year, dayOfWeek, minBalance, maxBalance } = filterParams;
    
    // กรองตามช่วงวันที่
    if (dateFrom) {
      data = data.filter(r => r.date >= dateFrom);
    }
    if (dateTo) {
      data = data.filter(r => r.date <= dateTo);
    }
    
    // กรองตามเดือน/ปี
    if (month && year) {
      const monthStr = String(month).padStart(2, '0');
      const yearStr = String(year);
      data = data.filter(r => r.date.startsWith(`${yearStr}-${monthStr}`));
    } else if (year) {
      data = data.filter(r => r.date.startsWith(String(year)));
    }
    
    // กรองตามวันในสัปดาห์
    if (dayOfWeek !== undefined && dayOfWeek !== null && dayOfWeek !== '') {
      data = data.filter(r => {
        if (!r.date) return false;
        const dow = new Date(r.date).getDay();
        return dow === parseInt(dayOfWeek);
      });
    }
    
    // กรองตามช่วง Balance
    if (minBalance !== undefined && minBalance !== null && minBalance !== '') {
      data = data.filter(r => r.balance >= Number(minBalance));
    }
    if (maxBalance !== undefined && maxBalance !== null && maxBalance !== '') {
      data = data.filter(r => r.balance <= Number(maxBalance));
    }
    
    // คำนวณสรุป
    const summary = {
      totalDays: data.length,
      totalRevenue: data.reduce((s, r) => s + r.revenue, 0),
      totalExpense: data.reduce((s, r) => s + r.expense, 0),
      totalBalance: data.reduce((s, r) => s + r.balance, 0),
      totalSold: data.reduce((s, r) => s + r.sold, 0),
      totalPending: data.reduce((s, r) => s + r.pending, 0),
      totalCleared: data.reduce((s, r) => s + r.cleared, 0),
      avgBalancePerDay: data.length ? (data.reduce((s, r) => s + r.balance, 0) / data.length) : 0,
      profitMargin: data.reduce((s, r) => s + r.revenue, 0) > 0 
        ? (data.reduce((s, r) => s + r.balance, 0) / data.reduce((s, r) => s + r.revenue, 0) * 100) : 0,
      conversionRate: data.reduce((s, r) => s + r.sold, 0) > 0 
        ? ((data.reduce((s, r) => s + r.sold, 0) - data.reduce((s, r) => s + r.pending, 0)) / data.reduce((s, r) => s + r.sold, 0) * 100) : 0,
    };
    
    // วิเคราะห์รายวัน
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
      balance: 0,
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
    
    // วิเคราะห์รายเดือน
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
    
    // วิเคราะห์ค่าใช้จ่าย
    const expenseBreakdown = {
      pipeFee: data.reduce((s, r) => s + r.pipeFee, 0),
      shareFee: data.reduce((s, r) => s + r.shareFee, 0),
      otherFee: data.reduce((s, r) => s + r.otherFee, 0),
      saveFee: data.reduce((s, r) => s + r.saveFee, 0),
    };
    const totalExpenseBreakdown = expenseBreakdown.pipeFee + expenseBreakdown.shareFee + 
                                   expenseBreakdown.otherFee + expenseBreakdown.saveFee;
    
    // Top/Bottom performance
    const sortedByBalance = data.slice().sort((a, b) => b.balance - a.balance);
    const top5 = sortedByBalance.slice(0, 5);
    const bottom5 = sortedByBalance.slice(-5).reverse();
    
    // สร้าง Insights
    const insights = [];
    const profitMargin = summary.profitMargin;
    const conversionRate = summary.conversionRate;
    const negativeDays = data.filter(r => r.balance < 0).length;
    
    if (profitMargin >= 40) {
      insights.push({
        type: 'positive',
        icon: '🚀',
        title: 'Margin ดีเยี่ยม',
        body: `Profit Margin อยู่ที่ ${profitMargin.toFixed(1)}% สูงกว่าเกณฑ์ 40% ธุรกิจมีสุขภาพการเงินดีมาก`,
      });
    } else if (profitMargin >= 20) {
      insights.push({
        type: 'warning',
        icon: '📊',
        title: 'Margin ปานกลาง',
        body: `Margin ${profitMargin.toFixed(1)}% ควรพยายามลดค่าใช้จ่ายหรือเพิ่มยอดขายให้ Margin เกิน 30%`,
      });
    } else {
      insights.push({
        type: 'negative',
        icon: '🔴',
        title: 'Margin ต่ำ / ขาดทุน',
        body: `Margin ${profitMargin.toFixed(1)}% ต่ำกว่าเกณฑ์ ควรปรับโครงสร้างต้นทุนด่วน`,
      });
    }
    
    if (negativeDays > 0) {
      insights.push({
        type: 'negative',
        icon: '📉',
        title: `${negativeDays} วันขาดทุน`,
        body: `มี ${negativeDays} วันที่ขาดทุน (${(negativeDays / data.length * 100).toFixed(0)}%) ควรวิเคราะห์รูปแบบการขายในวันเหล่านั้น`,
      });
    }
    
    if (conversionRate < 60) {
      insights.push({
        type: 'warning',
        icon: '⏳',
        title: 'Conversion Rate ต่ำ',
        body: `${conversionRate.toFixed(1)}% ของยอดขายยังรอชำระ ควรกระตุ้นการชำระเงินและลด pending`,
      });
    } else {
      insights.push({
        type: 'positive',
        icon: '✅',
        title: 'Conversion Rate ดี',
        body: `${conversionRate.toFixed(1)}% ลูกค้าส่วนใหญ่ชำระแล้ว Cash Flow ดี`,
      });
    }
    
    insights.push({
      type: 'info',
      icon: '💡',
      title: 'สรุปภาพรวม',
      body: `ข้อมูล ${data.length} วัน | กำไรสูงสุด ฿${Math.max(...data.map(r => r.balance)).toLocaleString()} | ต่ำสุด ฿${Math.min(...data.map(r => r.balance)).toLocaleString()} | เฉลี่ย ฿${Math.round(summary.avgBalancePerDay).toLocaleString()}/วัน`,
    });
    
    return {
      success: true,
      data: {
        filteredRecords: data,
        summary,
        dayStats,
        monthlyStats,
        expenseBreakdown: {
          ...expenseBreakdown,
          total: totalExpenseBreakdown,
        },
        top5,
        bottom5,
        insights,
      },
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Add Record ───────────────────────────────────────────────
function addRecord(rec) {
  try {
    const sheet    = _sheet();
    const sold     = Number(rec.sold)     || 0;
    const pending  = Number(rec.pending)  || 0;
    const cleared  = Number(rec.cleared)  || 0;
    const pipeFee  = Number(rec.pipeFee)  || 0;
    const shareFee = Number(rec.shareFee) || 0;
    const otherFee = Number(rec.otherFee) || 0;
    const saveFee  = Number(rec.saveFee)  || 0;
    const revenue  = (sold + cleared - pending) * PRICE_PER_UNIT;
    const expense  = pipeFee + shareFee + otherFee + saveFee;
    const balance  = revenue - expense;

    sheet.appendRow([
      new Date(rec.date), sold, pending, cleared,
      revenue, pipeFee, shareFee, otherFee, saveFee,
      expense, balance, new Date()
    ]);
    
    // ส่งแจ้งเตือนผ่าน Line OA
    sendLineNotification({
      type: 'add',
      date: rec.date,
      sold,
      pending,
      cleared,
      revenue,
      expense,
      balance,
    });
    
    return { success: true, message: 'Record added.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Update Record ────────────────────────────────────────────
function updateRecord(rowIndex, rec) {
  try {
    const sheet    = _sheet();
    const sold     = Number(rec.sold)     || 0;
    const pending  = Number(rec.pending)  || 0;
    const cleared  = Number(rec.cleared)  || 0;
    const pipeFee  = Number(rec.pipeFee)  || 0;
    const shareFee = Number(rec.shareFee) || 0;
    const otherFee = Number(rec.otherFee) || 0;
    const saveFee  = Number(rec.saveFee)  || 0;
    const revenue  = (sold + cleared - pending) * PRICE_PER_UNIT;
    const expense  = pipeFee + shareFee + otherFee + saveFee;
    const balance  = revenue - expense;

    sheet.getRange(rowIndex, 1, 1, 12).setValues([[
      new Date(rec.date), sold, pending, cleared,
      revenue, pipeFee, shareFee, otherFee, saveFee,
      expense, balance, new Date()
    ]]);
    
    // ส่งแจ้งเตือนผ่าน Line OA
    sendLineNotification({
      type: 'update',
      date: rec.date,
      sold,
      pending,
      cleared,
      revenue,
      expense,
      balance,
    });
    
    return { success: true, message: 'Record updated.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Delete Record ────────────────────────────────────────────
function deleteRecord(rowIndex) {
  try {
    _sheet().deleteRow(rowIndex);
    
    // ส่งแจ้งเตือนผ่าน Line OA
    sendLineNotification({
      type: 'delete',
      message: `ลบรายการที่แถว ${rowIndex}`,
    });
    
    return { success: true, message: 'Record deleted.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── LINE OA Notification ─────────────────────────────────────
function sendLineNotification(data) {
  try {
    if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN === 'YOUR_LINE_CHANNEL_ACCESS_TOKEN') {
      console.log('LINE OA Token not configured, skipping notification');
      return;
    }
    
    let message = '';
    
    if (data.type === 'add') {
      message = `📝 เพิ่มรายการใหม่\n` +
                `📅 วันที่: ${data.date}\n` +
                `🛒 ขาย: ${data.sold} | ⏳ รอ: ${data.pending} | ✅ ชำระ: ${data.cleared}\n` +
                `💰 รายได้: ฿${data.revenue.toLocaleString()}\n` +
                `💸 ค่าใช้จ่าย: ฿${data.expense.toLocaleString()}\n` +
                `📈 กำไร: ฿${data.balance.toLocaleString()}`;
    } else if (data.type === 'update') {
      message = `✏️ แก้ไขรายการ\n` +
                `📅 วันที่: ${data.date}\n` +
                `🛒 ขาย: ${data.sold} | ⏳ รอ: ${data.pending} | ✅ ชำระ: ${data.cleared}\n` +
                `💰 รายได้: ฿${data.revenue.toLocaleString()}\n` +
                `💸 ค่าใช้จ่าย: ฿${data.expense.toLocaleString()}\n` +
                `📈 กำไร: ฿${data.balance.toLocaleString()}`;
    } else if (data.type === 'delete') {
      message = `🗑 ${data.message}`;
    }
    
    const payload = {
      to: LINE_ADMIN_USER_ID,
      messages: [{
        type: 'text',
        text: message,
      }],
    };
    
    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
      },
      payload: JSON.stringify(payload),
    };
    
    UrlFetchApp.fetch(LINE_API_URL, options);
  } catch (e) {
    console.log('Line notification error:', e.message);
  }
}

// ─── Send Custom Line Message ─────────────────────────────────
function sendCustomLineMessage(message) {
  try {
    if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN === 'YOUR_LINE_CHANNEL_ACCESS_TOKEN') {
      return { success: false, error: 'LINE OA Token not configured' };
    }
    
    const payload = {
      to: LINE_ADMIN_USER_ID,
      messages: [{
        type: 'text',
        text: message,
      }],
    };
    
    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
      },
      payload: JSON.stringify(payload),
    };
    
    UrlFetchApp.fetch(LINE_API_URL, options);
    return { success: true, message: 'Message sent to LINE' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Send KPI Report to LINE ──────────────────────────────────
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
    const today = new Date();
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // หาข้อมูลวันนี้
    const todayData = data.filter(r => r.date === todayStr);
    const todayRevenue = todayData.reduce((s, r) => s + r.revenue, 0);
    const todayExpense = todayData.reduce((s, r) => s + r.expense, 0);
    const todayBalance = todayData.reduce((s, r) => s + r.balance, 0);
    
    const message = 
      `📊 KPI REPORT — ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}\n\n` +
      `📅 ภาพรวม (${data.length} วัน)\n` +
      `💰 รายได้รวม: ฿${totalRev.toLocaleString()}\n` +
      `💸 ค่าใช้จ่ายรวม: ฿${totalExp.toLocaleString()}\n` +
      `📈 กำไรรวม: ฿${totalBal.toLocaleString()}\n` +
      `🎯 Profit Margin: ${margin.toFixed(1)}%\n` +
      `✅ Conversion Rate: ${convRate.toFixed(1)}%\n` +
      `🛒 ขายรวม: ${totalSold} | ⏳ รอ: ${totalPend} | ✅ ชำระ: ${totalClr}\n\n` +
      `📌 วันนี้ (${todayStr})\n` +
      `💰 รายได้: ฿${todayRevenue.toLocaleString()}\n` +
      `💸 ค่าใช้จ่าย: ฿${todayExpense.toLocaleString()}\n` +
      `📈 กำไร: ฿${todayBalance.toLocaleString()}`;
    
    return sendCustomLineMessage(message);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Send Deep Analysis to LINE ───────────────────────────────
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
    
    // วิเคราะห์รายวัน
    const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    const dayStats = Array.from({ length: 7 }, () => ({ count: 0, sold: 0, revenue: 0, expense: 0, balance: 0 }));
    
    data.forEach(r => {
      if (!r.date) return;
      const dow = new Date(r.date).getDay();
      dayStats[dow].count++;
      dayStats[dow].sold += r.sold;
      dayStats[dow].revenue += r.revenue;
      dayStats[dow].expense += r.expense;
      dayStats[dow].balance += r.balance;
    });
    
    // หาวันที่ดีที่สุดและแย่ที่สุด
    let bestDay = 0, worstDay = 0;
    dayStats.forEach((s, i) => {
      if (s.count > 0) {
        if (s.balance > dayStats[bestDay].balance || dayStats[bestDay].count === 0) bestDay = i;
        if (s.balance < dayStats[worstDay].balance || dayStats[worstDay].count === 0) worstDay = i;
      }
    });
    
    // วิเคราะห์ค่าใช้จ่าย
    const pipeFee = data.reduce((s, r) => s + r.pipeFee, 0);
    const shareFee = data.reduce((s, r) => s + r.shareFee, 0);
    const otherFee = data.reduce((s, r) => s + r.otherFee, 0);
    const saveFee = data.reduce((s, r) => s + r.saveFee, 0);
    const totalFee = pipeFee + shareFee + otherFee + saveFee;
    
    // วิเคราะห์แนวโน้ม
    const sortedData = data.slice().sort((a, b) => a.date.localeCompare(b.date));
    const firstHalf = sortedData.slice(0, Math.floor(sortedData.length / 2));
    const secondHalf = sortedData.slice(Math.floor(sortedData.length / 2));
    const firstHalfBal = firstHalf.reduce((s, r) => s + r.balance, 0);
    const secondHalfBal = secondHalf.reduce((s, r) => s + r.balance, 0);
    const trend = secondHalfBal > firstHalfBal ? '📈 เพิ่มขึ้น' : secondHalfBal < firstHalfBal ? '📉 ลดลง' : '➡️ คงที่';
    
    // สร้างข้อความวิเคราะห์เชิงลึก
    let message = `🧠 DEEP ANALYSIS REPORT\n\n`;
    
    message += `📊 ภาพรวมทางการเงิน\n`;
    message += `💰 รายได้รวม: ฿${totalRev.toLocaleString()}\n`;
    message += `💸 ค่าใช้จ่ายรวม: ฿${totalExp.toLocaleString()}\n`;
    message += `📈 กำไรรวม: ฿${totalBal.toLocaleString()}\n`;
    message += `🎯 Profit Margin: ${margin.toFixed(1)}%\n\n`;
    
    message += `📅 วิเคราะห์รายวัน\n`;
    message += `🏆 วันที่ดีที่สุด: ${dayNames[bestDay]} (กำไร ฿${dayStats[bestDay].balance.toLocaleString()})\n`;
    message += `⚠️ วันที่ควรปรับ: ${dayNames[worstDay]} (กำไร ฿${dayStats[worstDay].balance.toLocaleString()})\n\n`;
    
    message += `💸 โครงสร้างค่าใช้จ่าย\n`;
    message += `• Pipeline Fee: ฿${pipeFee.toLocaleString()} (${(pipeFee/totalFee*100).toFixed(1)}%)\n`;
    message += `• Share Fee: ฿${shareFee.toLocaleString()} (${(shareFee/totalFee*100).toFixed(1)}%)\n`;
    message += `• Other Fee: ฿${otherFee.toLocaleString()} (${(otherFee/totalFee*100).toFixed(1)}%)\n`;
    message += `• Save Fee: ฿${saveFee.toLocaleString()} (${(saveFee/totalFee*100).toFixed(1)}%)\n\n`;
    
    message += `📈 แนวโน้มธุรกิจ\n`;
    message += `เปรียบเทียบครึ่งแรก vs ครึ่งหลัง: ${trend}\n`;
    message += `กำไรครึ่งแรก: ฿${firstHalfBal.toLocaleString()} | ครึ่งหลัง: ฿${secondHalfBal.toLocaleString()}\n\n`;
    
    // คำแนะนำ
    message += `💡 คำแนะนำ\n`;
    if (margin < 20) {
      message += `🔴 เร่งด่วน: ปรับโครงสร้างต้นทุน Margin ต่ำกว่า 20%\n`;
    }
    if (pipeFee / totalFee > 0.5) {
      message += `⚠️ Pipeline Fee สูงเกิน 50% ควรเจรจาลด\n`;
    }
    if (trend.includes('ลดลง')) {
      message += `⚠️ แนวโน้มกำไรลดลง ควรตรวจสอบสาเหตุ\n`;
    }
    message += `✅ ติดตาม KPI สัปดาห์ละครั้งเพื่อปรับกลยุทธ์\n`;
    
    return sendCustomLineMessage(message);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// AUTO REPORT SHEET GENERATOR
// ═══════════════════════════════════════════════════════════════
function generateAnalyticsReport() {
  try {
    const ss = _ss();

    // ── Remove old report sheet ────────────────────────────────
    const old = ss.getSheetByName(REPORT_SHEET);
    if (old) ss.deleteSheet(old);

    // ── Create new sheet ───────────────────────────────────────
    const rs = ss.insertSheet(REPORT_SHEET);

    // ── Fetch raw data ─────────────────────────────────────────
    const src     = _sheet();
    const lastRow = src.getLastRow();
    if (lastRow < 2) {
      rs.getRange('A1').setValue('No data available.');
      return { success: true, message: 'No data to analyse.' };
    }

    const values = src.getRange(2, 1, lastRow - 1, 12).getValues();
    const data   = values.map((r, i) => _parseRow(r, i + 2)).filter(Boolean);

    if (!data.length) {
      rs.getRange('A1').setValue('No data available.');
      return { success: true, message: 'No data to analyse.' };
    }

    // ── Colour palette ─────────────────────────────────────────
    const C = {
      headerBg  : '#0f1724',
      headerFg  : '#ffffff',
      titleBg   : '#1a2744',
      titleFg   : '#7eb8ff',
      altRow    : '#f0f4ff',
      posText   : '#1a7a4a',
      negText   : '#c0392b',
      border    : '#c8d0e0',
      sectionBg : '#e8edf8',
    };

    const TZ = Session.getScriptTimeZone();

    // Helper: สร้าง header block
    function writeBlock(sheet, startRow, title, headers, rows, formats) {
      const hdrRange = sheet.getRange(startRow, 1, 1, headers.length);
      hdrRange.setValues([headers])
              .setBackground(C.headerBg)
              .setFontColor(C.headerFg)
              .setFontWeight('bold')
              .setHorizontalAlignment('center');

      rows.forEach((row, ri) => {
        const r    = sheet.getRange(startRow + 1 + ri, 1, 1, row.length);
        r.setValues([row]);
        if (ri % 2 === 1) r.setBackground(C.altRow);
        if (formats) {
          formats.forEach((f, fi) => {
            if (f) sheet.getRange(startRow + 1 + ri, fi + 1).setNumberFormat(f);
          });
        }
      });
      return startRow + 1 + rows.length + 1;
    }

    // ── 1. OVERVIEW SECTION ────────────────────────────────────
    let row = 1;
    const titleCell = rs.getRange(row, 1, 1, 8);
    titleCell.merge()
             .setValue('📊 SALES ANALYTICS REPORT — Generated: ' + _fmt(new Date(), 'dd/MM/yyyy HH:mm'))
             .setBackground(C.titleBg)
             .setFontColor(C.titleFg)
             .setFontSize(13)
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
    row += 2;

    const totalRev  = data.reduce((s, r) => s + r.revenue, 0);
    const totalExp  = data.reduce((s, r) => s + r.expense, 0);
    const totalBal  = data.reduce((s, r) => s + r.balance, 0);
    const totalSold = data.reduce((s, r) => s + r.sold, 0);
    const totalPend = data.reduce((s, r) => s + r.pending, 0);
    const totalClr  = data.reduce((s, r) => s + r.cleared, 0);
    const margin    = totalRev > 0 ? (totalBal / totalRev) * 100 : 0;
    const avgBalDay = data.length ? totalBal / data.length : 0;
    const convRate  = (totalSold > 0) ? ((totalSold - totalPend) / totalSold * 100) : 0;

    row = writeBlock(rs, row,
      'OVERVIEW',
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
        ['📅 Avg Profit/Day', avgBalDay.toFixed(0), '฿' + avgBalDay.toLocaleString('th-TH', { maximumFractionDigits: 0 })],
      ],
      ['@', '#,##0.00', '@']
    );

    // ── 2. DAILY BREAKDOWN (Mon–Sun) ────────────────────────────
    const dayNames  = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์'];
    const dayStats  = Array.from({ length: 7 }, () => ({ cnt: 0, sold: 0, pending: 0, cleared: 0, rev: 0, exp: 0, bal: 0 }));

    data.forEach(r => {
      if (!r.date) return;
      const dow = new Date(r.date).getDay();       // 0=Sun
      const idx = dow === 0 ? 6 : dow - 1;         // Mon=0…Sun=6
      const s   = dayStats[idx];
      s.cnt++; s.sold += r.sold; s.pending += r.pending;
      s.cleared += r.cleared; s.rev += r.revenue;
      s.exp += r.expense; s.bal += r.balance;
    });

    const dayRows = dayStats.map((s, i) => [
      dayNames[i],
      s.cnt,
      s.sold,
      s.cnt > 0 ? +(s.sold / s.cnt).toFixed(1) : 0,
      s.rev,
      s.cnt > 0 ? +(s.rev / s.cnt).toFixed(0) : 0,
      s.exp,
      s.bal,
      s.cnt > 0 ? +(s.bal / s.cnt).toFixed(0) : 0,
      s.rev > 0  ? +(s.bal / s.rev * 100).toFixed(1) : 0,
    ]);

    row = writeBlock(rs, row,
      'DAY OF WEEK ANALYSIS',
      ['วัน', 'จำนวนวัน', 'ขายรวม (ชิ้น)', 'ขายเฉลี่ย/วัน', 'รายได้รวม (฿)', 'รายได้เฉลี่ย/วัน', 'ค่าใช้จ่ายรวม (฿)', 'กำไรรวม (฿)', 'กำไรเฉลี่ย/วัน', 'Margin (%)'],
      dayRows,
      ['@', '#,##0', '#,##0', '#,##0.0', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0']
    );

    // ── 3. MONTHLY BREAKDOWN ────────────────────────────────────
    const monthMap = {};
    data.forEach(r => {
      if (!r.date) return;
      const ym = r.date.slice(0, 7);  // 'yyyy-MM'
      if (!monthMap[ym]) monthMap[ym] = { cnt: 0, sold: 0, pending: 0, cleared: 0, rev: 0, exp: 0, bal: 0 };
      const m = monthMap[ym];
      m.cnt++; m.sold += r.sold; m.pending += r.pending;
      m.cleared += r.cleared; m.rev += r.revenue;
      m.exp += r.expense; m.bal += r.balance;
    });

    const monthRows = Object.entries(monthMap).sort().map(([ym, m]) => {
      const [yr, mo] = ym.split('-');
      const thMonths = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
                        'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
      return [
        thMonths[+mo] + ' ' + yr,
        m.cnt, m.sold, m.rev, m.exp, m.bal,
        m.rev > 0 ? +(m.bal / m.rev * 100).toFixed(1) : 0,
        m.cnt > 0 ? +(m.rev / m.cnt).toFixed(0) : 0,
        m.cnt > 0 ? +(m.bal / m.cnt).toFixed(0) : 0,
      ];
    });

    row = writeBlock(rs, row,
      'MONTHLY ANALYSIS',
      ['เดือน', 'จำนวนวัน', 'ยอดขายรวม (ชิ้น)', 'รายได้รวม (฿)', 'ค่าใช้จ่ายรวม (฿)', 'กำไรรวม (฿)', 'Margin (%)', 'รายได้เฉลี่ย/วัน', 'กำไรเฉลี่ย/วัน'],
      monthRows,
      ['@', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0', '#,##0', '#,##0']
    );

    // ── 4. EXPENSE BREAKDOWN ────────────────────────────────────
    const totPipe  = data.reduce((s, r) => s + r.pipeFee, 0);
    const totShare = data.reduce((s, r) => s + r.shareFee, 0);
    const totOther = data.reduce((s, r) => s + r.otherFee, 0);
    const totSave  = data.reduce((s, r) => s + r.saveFee, 0);
    const totAll   = totPipe + totShare + totOther + totSave || 1;

    row = writeBlock(rs, row,
      'EXPENSE BREAKDOWN',
      ['ประเภทค่าใช้จ่าย', 'รวม (฿)', 'สัดส่วน (%)', 'เฉลี่ย/วัน (฿)', 'วิเคราะห์'],
      [
        ['Pipeline Fee', totPipe,  +(totPipe / totAll * 100).toFixed(1),  +(totPipe / data.length).toFixed(0),  totPipe / totAll > 0.5 ? '⚠️ สูงมาก' : '✅ ปกติ'],
        ['Share Fee',    totShare, +(totShare / totAll * 100).toFixed(1), +(totShare / data.length).toFixed(0), totShare / totAll > 0.3 ? '⚠️ ตรวจสอบ' : '✅ ปกติ'],
        ['Other Fee',    totOther, +(totOther / totAll * 100).toFixed(1), +(totOther / data.length).toFixed(0), '—'],
        ['Save Fee',     totSave,  +(totSave / totAll * 100).toFixed(1),  +(totSave / data.length).toFixed(0),  '—'],
        ['TOTAL',        totalExp, 100, +(totalExp / data.length).toFixed(0), ''],
      ],
      ['@', '#,##0', '#,##0.0', '#,##0', '@']
    );

    // ── 5. TOP PERFORMANCE DAYS ─────────────────────────────────
    const top5  = data.slice().sort((a, b) => b.balance - a.balance).slice(0, 5);
    const bot5  = data.slice().sort((a, b) => a.balance - b.balance).slice(0, 5);

    row = writeBlock(rs, row,
      'TOP 5 BEST PROFIT DAYS',
      ['วันที่', 'Sold', 'Revenue (฿)', 'Expense (฿)', 'Balance (฿)', 'Margin (%)'],
      top5.map(r => [r.date, r.sold, r.revenue, r.expense, r.balance,
                     r.revenue > 0 ? +(r.balance / r.revenue * 100).toFixed(1) : 0]),
      ['@', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0']
    );

    row = writeBlock(rs, row,
      'TOP 5 WORST PROFIT DAYS',
      ['วันที่', 'Sold', 'Revenue (฿)', 'Expense (฿)', 'Balance (฿)', 'Margin (%)'],
      bot5.map(r => [r.date, r.sold, r.revenue, r.expense, r.balance,
                     r.revenue > 0 ? +(r.balance / r.revenue * 100).toFixed(1) : 0]),
      ['@', '#,##0', '#,##0', '#,##0', '#,##0', '#,##0.0']
    );

    // ── 6. BUSINESS INSIGHTS TEXT ──────────────────────────────
    const insights = _buildInsightText(data, totalRev, totalExp, totalBal, margin, convRate);
    const insRow   = rs.getRange(row, 1, 1, 6);
    insRow.merge().setValue('🧠 BUSINESS INSIGHTS & RECOMMENDATIONS')
          .setBackground(C.titleBg).setFontColor(C.titleFg)
          .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    row++;

    insights.forEach(ins => {
      const cell = rs.getRange(row, 1, 1, 6);
      cell.merge().setValue(ins.icon + ' ' + ins.title + ': ' + ins.body)
          .setFontColor(ins.type === 'positive' ? C.posText : ins.type === 'negative' ? C.negText : '#7a5f00')
          .setWrap(true).setRowHeight ? rs.setRowHeight(row, 40) : null;
      row++;
    });

    // ── Column widths ──────────────────────────────────────────
    rs.setColumnWidth(1, 160);
    [2,3,4,5,6,7,8,9,10].forEach(c => rs.setColumnWidth(c, 130));

    // ── Freeze header row ──────────────────────────────────────
    rs.setFrozenRows(1);

    // ── Move sheet to front ────────────────────────────────────
    ss.setActiveSheet(rs);
    ss.moveActiveSheet(1);

    return { success: true, message: 'Report sheet "' + REPORT_SHEET + '" generated successfully!' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Insight builder ───────────────────────────────────────────
function _buildInsightText(data, totalRev, totalExp, totalBal, margin, convRate) {
  const ins  = [];
  const days = data.length;
  const avgBal  = days ? totalBal / days : 0;
  const negDays = data.filter(r => r.balance < 0).length;
  const maxBal  = Math.max(...data.map(r => r.balance));
  const minBal  = Math.min(...data.map(r => r.balance));
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);

  if (margin >= 40)
    ins.push({ type: 'positive', icon: '🚀', title: 'Margin ดีเยี่ยม', body: `Profit Margin อยู่ที่ ${margin.toFixed(1)}% สูงกว่าเกณฑ์ 40% ธุรกิจมีสุขภาพการเงินดีมาก` });
  else if (margin >= 20)
    ins.push({ type: 'warning', icon: '📊', title: 'Margin ปานกลาง', body: `Margin ${margin.toFixed(1)}% ควรพยายามลดค่าใช้จ่ายหรือเพิ่มยอดขายให้ Margin เกิน 30%` });
  else
    ins.push({ type: 'negative', icon: '🔴', title: 'Margin ต่ำ / ขาดทุน', body: `Margin ${margin.toFixed(1)}% ต่ำกว่าเกณฑ์ ควรปรับโครงสร้างต้นทุนด่วน` });

  if (negDays > 0)
    ins.push({ type: 'negative', icon: '📉', title: `${negDays} วันขาดทุน`, body: `มี ${negDays} วันที่ขาดทุน (${(negDays/days*100).toFixed(0)}%) ควรวิเคราะห์รูปแบบการขายในวันเหล่านั้น` });

  if (convRate < 60)
    ins.push({ type: 'warning', icon: '⏳', title: 'Conversion Rate ต่ำ', body: `${convRate.toFixed(1)}% ของยอดขายยังรอชำระ ควรกระตุ้นการชำระเงินและลด pending` });
  else
    ins.push({ type: 'positive', icon: '✅', title: 'Conversion Rate ดี', body: `${convRate.toFixed(1)}% ลูกค้าส่วนใหญ่ชำระแล้ว Cash Flow ดี` });

  ins.push({ type: 'info', icon: '💡', title: 'สรุปภาพรวม', body: `ข้อมูล ${days} วัน | กำไรสูงสุด ฿${maxBal.toLocaleString()} | ต่ำสุด ฿${minBal.toLocaleString()} | เฉลี่ย ฿${Math.round(avgBal).toLocaleString()}/วัน` });

  return ins;
}

// ─── Webhook Entry Point ──────────────────────────────────
function doPost(e) {
  return doPostEnhanced(e);
}
// ─── Command Router ───────────────────────────────────────────
function processBotCommand(cmd) {
  const c = cmd.toLowerCase().trim();
  const data = getAllSalesDataForBot();
  if (!data.length) return txtMsg('❌ ยังไม่มีข้อมูลในระบบ กรุณาบันทึกข้อมูลก่อน');

  if (c.match(/\/(kpi|dashboard)/i))        return botKpiFlex(data);
  if (c.match(/\/(deep|analyze)/i))          return botDeepAnalysis(data);
  if (c.match(/\/(today|วันนี้)/i))           return botDaily(data, 'today');
  if (c.match(/\/(yesterday|เมื่อวาน)/i))      return botDaily(data, 'yesterday');
  if (c.match(/\/(week|สัปดาห์)/i))           return botWeekly(data);
  if (c.match(/\/(month|เดือน)/i))            return botMonthly(data);
  if (c.match(/\/(expense|ค่าใช้จ่าย)/i))      return botExpense(data);
  if (c.match(/\/(bottom|worst)(\s+\d+)?/i)) return botTop(data, 'bottom', c);
  if (c.match(/\/(trend|แนวโน้ม)/i))          return botTrend(data);
  if (c.match(/\/(compare|เปรียบเทียบ)/i))     return botCompare(data);
  if (c.match(/\/(margin|profit)/i))          return botMargin(data);
  if (c.match(/\/(conversion|conv)/i))        return botConversion(data);
  if (c.match(/\/(sales|ยอดขาย)/i))           return botSales(data);
  if (c.match(/\/(advice|แนะนำ)/i))            return botAdvice(data);
  if (c.match(/\/(stats|สถิติ)/i))             return botStats(data);
  if (c.match(/\/(latest|ล่าสุด)(\s+\d+)?/i)) return botLatest(data, c);
  if (c.match(/\/(search|ค้นหา)\s+/i))        return botSearch(data, c);
  if (c.match(/\/day\s+(\S+)/i))             return botDayOfWeek(data, c);
    // ═══ คำสั่งใหม่ระดับมืออาชีพ ═══
  if (c.match(/\/(forecast|พยากรณ์|คาดการณ์)/i))       return botForecast(data);
  if (c.match(/\/(risk|ความเสี่ยง)/i))                  return botRiskAnalysis(data);
  if (c.match(/\/(efficiency|ประสิทธิภาพ)/i))           return botEfficiency(data);
  if (c.match(/\/(benchmark|เทียบมาตรฐาน)/i))           return botBenchmark(data);
  if (c.match(/\/(cashflow|กระแสเงินสด)/i))             return botCashflow(data);
  if (c.match(/\/(breakeven|จุดคุ้มทุน)/i))              return botBreakeven(data);
  if (c.match(/\/(growth|การเติบโต)/i))                 return botGrowth(data);
  if (c.match(/\/(seasonality|ฤดูกาล)/i))               return botSeasonality(data);
  if (c.match(/\/(profitloss|pl|งบกำไรขาดทุน)/i))       return botProfitLoss(data);
  if (c.match(/\/(roi|ผลตอบแทน)/i))                     return botROI(data);
  if (c.match(/\/(health|สุขภาพธุรกิจ)/i))               return botHealthScore(data);
  if (c.match(/\/(alert|แจ้งเตือน)/i))                   return botAlertConfig(data);
  if (c.match(/\/(dashboard2|d2)/i))                     return botDashboardV2(data);
  if (c.match(/\/(summary|executive|ผู้บริหาร)/i))        return botExecutiveSummary(data);
  if (c.match(/\/(swot|จุดแข็งจุดอ่อน)/i))               return botSWOT(data);
  if (c.match(/\/(goal|เป้าหมาย|target)/i))              return botGoalTracking(data);
  if (c.match(/\/(invoice|แจ้งหนี้)/i))                   return botOutstandingReport(data);
  if (c.match(/\/(projection|ประมาณการ)/i))              return botProjection(data);
  if (c.match(/\/(ratio|อัตราส่วน)/i))                   return botFinancialRatios(data);
  if (c.match(/\/(export|ส่งออก)\s+(csv|json|pdf)/i))   return botExportData(data, c);
  if (c.match(/\/kpi2/i)) return createPremiumKpiFlex(data);
  if (c.match(/\/healthcard/i)) return createHealthScoreFlex(data);
  if (c.match(/\/quick/i)) return createQuickActionFlex();


  return botHelp();
}

// ─── Help ──────────────────────────────────────────────────
function botHelp() {
  return txtMsg(
    `🤖 **SALES ANALYTICS BOT** — คำสั่งทั้งหมด\n\n` +
    `📊 **รายงานพื้นฐาน**\n` +
    `/kpi — KPI Dashboard\n` +
    `/kpi2 — Premium KPI Dashboard\n` +
    `/d2 — Dashboard V2\n` +
    `/summary — Executive Summary\n` +
    `/today — รายงานวันนี้\n` +
    `/yesterday — รายงานเมื่อวาน\n` +
    `/week — รายงานสัปดาห์นี้\n` +
    `/month — รายงานเดือนนี้\n\n` +
    `🔍 **วิเคราะห์เชิงลึก**\n` +
    `/deep — วิเคราะห์เชิงลึก\n` +
    `/forecast — พยากรณ์ 7 วัน\n` +
    `/trend — แนวโน้มธุรกิจ\n` +
    `/seasonality — ฤดูกาล\n\n` +
    `💰 **การเงิน**\n` +
    `/pl — งบกำไรขาดทุน\n` +
    `/cashflow — กระแสเงินสด\n` +
    `/breakeven — จุดคุ้มทุน\n` +
    `/roi — ผลตอบแทนการลงทุน\n\n` +
    `🎨 **Flex Messages**\n` +
    `/kpi2 — Premium KPI\n` +
    `/healthcard — Health Card\n` +
    `/quick — Quick Actions\n\n` +
    `📋 /help — คำสั่งทั้งหมด`
  );
}

// ═══════════════════════════════════════════════════════════════
// BOT REPORT FUNCTIONS
// ═══════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════
// BOT REPORT FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function getAllSalesDataForBot() {
  try {
    const sheet = _sheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    return values.map((r, i) => _parseRow(r, i + 2)).filter(Boolean);
  } catch (e) { return []; }
}

function fmtBot(n) {
  n = +n || 0;
  if (Math.abs(n) >= 1e6) return '฿' + (n/1e6).toFixed(1) + 'M';
  if (Math.abs(n) >= 1e3) return '฿' + (n/1e3).toFixed(0) + 'K';
  return '฿' + n.toLocaleString('th-TH', {maximumFractionDigits:0});
}

function txtMsg(text) { return { type: 'text', text }; }

// ─── KPI Flex Message ─────────────────────────────────────────
function botKpiFlex(data) {
  const rev = data.reduce((s,r) => s + r.revenue, 0);
  const exp = data.reduce((s,r) => s + r.expense, 0);
  const bal = data.reduce((s,r) => s + r.balance, 0);
  const sold = data.reduce((s,r) => s + r.sold, 0);
  const pend = data.reduce((s,r) => s + r.pending, 0);
  const clr = data.reduce((s,r) => s + r.cleared, 0);
  const margin = rev > 0 ? (bal/rev*100).toFixed(1) : 0;
  const conv = sold > 0 ? ((sold-pend)/sold*100).toFixed(1) : 0;
  const status = margin >= 30 ? '🟢 ดีมาก' : margin >= 15 ? '🟡 ปานกลาง' : '🔴 ต้องปรับปรุง';

  try {
    return {
      type: 'flex',
      altText: '📊 KPI Dashboard',
      contents: {
        type: 'bubble',
        header: { type: 'box', layout: 'vertical',
          contents: [
            { type: 'text', text: '📊 KPI DASHBOARD', weight: 'bold', size: 'lg', color: '#ffffff' },
            { type: 'text', text: status, size: 'sm', color: '#9ba3be', margin: 'md' }
          ], backgroundColor: '#1a2744' },
        body: { type: 'box', layout: 'vertical',
          contents: [
            rowKpi('💰 รายได้รวม', fmtBot(rev), '#4d8bff'),
            rowKpi('💸 ค่าใช้จ่าย', fmtBot(exp), '#ff3d5a'),
            rowKpi('📈 กำไรรวม', fmtBot(bal), bal>=0?'#10d98a':'#ff3d5a'),
            { type: 'separator', margin: 'md' },
            rowKpi('🎯 Profit Margin', margin+'%', margin>=30?'#10d98a':'#f59e0b'),
            rowKpi('✅ Conversion', conv+'%', conv>=70?'#10d98a':'#f59e0b'),
            { type: 'separator', margin: 'md' },
            rowKpi('🛒 ยอดขาย', sold+' ชิ้น', '#f59e0b'),
            rowKpi('⏳ รอชำระ', pend+' ชิ้น', '#ff8c42'),
            rowKpi('✅ ชำระแล้ว', clr+' ชิ้น', '#10d98a'),
            { type: 'separator', margin: 'md' },
            rowKpi('📅 จำนวนวัน', data.length+' วัน', '#a855f7'),
            rowKpi('📊 กำไรเฉลี่ย/วัน', fmtBot(data.length?bal/data.length:0), '#22d3ee')
          ] }
      }
    };
  } catch (e) {
    return txtMsg(`📊 KPI DASHBOARD\n\n💰 รายได้: ${fmtBot(rev)}\n💸 ค่าใช้จ่าย: ${fmtBot(exp)}\n📈 กำไร: ${fmtBot(bal)}\n🎯 Margin: ${margin}%\n✅ Conversion: ${conv}%\n🛒 ยอดขาย: ${sold} ชิ้น`);
  }
}

function rowKpi(label, value, color) {
  return { type: 'box', layout: 'horizontal',
    contents: [
      { type: 'text', text: label, size: 'sm', color: '#9ba3be', flex: 0 },
      { type: 'text', text: String(value), size: 'sm', weight: 'bold', color: color, align: 'end', flex: 1 }
    ], margin: 'sm' };
}

// ─── Deep Analysis ────────────────────────────────────────────
function botDeepAnalysis(data) {
  const rev = data.reduce((s,r)=>s+r.revenue,0);
  const exp = data.reduce((s,r)=>s+r.expense,0);
  const bal = data.reduce((s,r)=>s+r.balance,0);
  const margin = rev>0?(bal/rev*100).toFixed(1):0;
  const sorted = data.slice().sort((a,b)=>a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length/2);
  const fBal = sorted.slice(0,mid).reduce((s,r)=>s+r.balance,0);
  const sBal = sorted.slice(mid).reduce((s,r)=>s+r.balance,0);
  const trend = sBal>fBal?'📈 เพิ่มขึ้น':sBal<fBal?'📉 ลดลง':'➡️ คงที่';
  const change = fBal!==0?((sBal-fBal)/Math.abs(fBal)*100).toFixed(1):0;

  const dayNames = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  const ds = Array.from({length:7},()=>({cnt:0,bal:0}));
  data.forEach(r=>{if(!r.date)return;const d=new Date(r.date).getDay();ds[d].cnt++;ds[d].bal+=r.balance;});
  const best = ds.reduce((m,s,i)=>s.cnt>0&&s.bal>ds[m].bal?i:m,0);
  // เพิ่มการตรวจสอบก่อนใช้งาน worst
const firstValidIndex = ds.findIndex(s => s.cnt > 0);
const worst = firstValidIndex >= 0 
  ? ds.reduce((m, s, i) => s.cnt > 0 && s.bal < ds[m].bal ? i : m, firstValidIndex)
  : 0;
  

  const pf=data.reduce((s,r)=>s+r.pipeFee,0),sf=data.reduce((s,r)=>s+r.shareFee,0);
  const of=data.reduce((s,r)=>s+r.otherFee,0),sv=data.reduce((s,r)=>s+r.saveFee,0);
  const tf=pf+sf+of+sv||1;

  let t=`🧠 DEEP ANALYSIS\n\n`;
  t+=`📊 ภาพรวม\n💰 รายได้: ${fmtBot(rev)}\n💸 ค่าใช้จ่าย: ${fmtBot(exp)}\n📈 กำไร: ${fmtBot(bal)} (${margin}%)\n\n`;
  t+=`📈 แนวโน้ม: ${trend} (${change}%)\nครึ่งแรก: ${fmtBot(fBal)} → ครึ่งหลัง: ${fmtBot(sBal)}\n\n`;
  t+=`📅 วันที่ดีที่สุด: ${dayNames[best]} (${fmtBot(ds[best].bal)})\n⚠️ ควรปรับ: ${dayNames[worst]} (${fmtBot(ds[worst].bal)})\n\n`;
  t+=`💸 ค่าใช้จ่าย\n• Pipeline: ${fmtBot(pf)} (${(pf/tf*100).toFixed(1)}%)\n• Share: ${fmtBot(sf)} (${(sf/tf*100).toFixed(1)}%)\n• Other: ${fmtBot(of)} (${(of/tf*100).toFixed(1)}%)\n• Save: ${fmtBot(sv)} (${(sv/tf*100).toFixed(1)}%)\n\n`;
  t+=`💡 คำแนะนำ\n`;
  if(margin<20)t+=`🔴 Margin ต่ำกว่า 20% ควรปรับโครงสร้างต้นทุน\n`;
  if(pf/tf>0.5)t+=`⚠️ Pipeline Fee สูงเกิน 50%\n`;
  if(trend.includes('ลดลง'))t+=`⚠️ แนวโน้มกำไรลดลง\n`;
  t+=`✅ ติดตาม KPI ทุกสัปดาห์`;

  return txtMsg(t);
}

// ─── Daily Report ─────────────────────────────────────────────
function botDaily(data, type) {
  const now = new Date();
  const target = type==='yesterday'?new Date(now.getTime()-86400000):now;
  const dateStr = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const thaiDate = Utilities.formatDate(target, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const dd = data.filter(r=>r.date===dateStr);
  if(!dd.length) return txtMsg(`📅 ${thaiDate}\n❌ ไม่มีข้อมูล`);

  const rev=dd.reduce((s,r)=>s+r.revenue,0),exp=dd.reduce((s,r)=>s+r.expense,0);
  const bal=dd.reduce((s,r)=>s+r.balance,0),sold=dd.reduce((s,r)=>s+r.sold,0);
  const margin=rev>0?(bal/rev*100).toFixed(1):0;

  let t=`📅 ${thaiDate}\n\n🛒 ยอดขาย: ${sold} ชิ้น\n💰 รายได้: ${fmtBot(rev)}\n💸 ค่าใช้จ่าย: ${fmtBot(exp)}\n📈 กำไร: ${fmtBot(bal)}\n🎯 Margin: ${margin}%`;
  if(dd.length>1){t+=`\n\n📋 ${dd.length} รายการ`;dd.forEach((r,i)=>t+=`\n${i+1}. Sold:${r.sold} | รายได้:${fmtBot(r.revenue)} | กำไร:${fmtBot(r.balance)}`);}
  return txtMsg(t);
}

// ─── Weekly Report ────────────────────────────────────────────
function botWeekly(data) {
  const now=new Date();const dow=now.getDay()||7;
  const mon=new Date(now);mon.setDate(now.getDate()-dow+1);
  const monStr=Utilities.formatDate(mon,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const sun=new Date(mon);sun.setDate(mon.getDate()+6);
  const sunStr=Utilities.formatDate(sun,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const wd=data.filter(r=>r.date>=monStr&&r.date<=sunStr);
  if(!wd.length) return txtMsg(`📅 ${monStr} - ${sunStr}\n❌ ไม่มีข้อมูล`);
  const rev=wd.reduce((s,r)=>s+r.revenue,0),exp=wd.reduce((s,r)=>s+r.expense,0);
  const bal=wd.reduce((s,r)=>s+r.balance,0),sold=wd.reduce((s,r)=>s+r.sold,0);
  const margin=rev>0?(bal/rev*100).toFixed(1):0;
  return txtMsg(`📅 ${monStr} - ${sunStr}\n\n📊 ${wd.length} วัน\n🛒 ยอดขาย: ${sold} ชิ้น\n💰 รายได้: ${fmtBot(rev)}\n💸 ค่าใช้จ่าย: ${fmtBot(exp)}\n📈 กำไร: ${fmtBot(bal)}\n🎯 Margin: ${margin}%\n📊 เฉลี่ย/วัน: ${fmtBot(wd.length?bal/wd.length:0)}`);
}

// ─── Monthly Report ───────────────────────────────────────────
function botMonthly(data) {
  const now=new Date();
  const monthStr=Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM');
  const md=data.filter(r=>r.date.startsWith(monthStr));
  const thM=['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const monthName=thM[now.getMonth()+1];
  if(!md.length) return txtMsg(`📅 ${monthName} ${now.getFullYear()}\n❌ ไม่มีข้อมูล`);
  const rev=md.reduce((s,r)=>s+r.revenue,0),exp=md.reduce((s,r)=>s+r.expense,0);
  const bal=md.reduce((s,r)=>s+r.balance,0),sold=md.reduce((s,r)=>s+r.sold,0);
  const margin=rev>0?(bal/rev*100).toFixed(1):0;
  return txtMsg(`📅 ${monthName} ${now.getFullYear()}\n\n📊 ${md.length} วัน\n🛒 ยอดขาย: ${sold} ชิ้น\n💰 รายได้: ${fmtBot(rev)}\n💸 ค่าใช้จ่าย: ${fmtBot(exp)}\n📈 กำไร: ${fmtBot(bal)}\n🎯 Margin: ${margin}%\n📊 เฉลี่ย/วัน: ${fmtBot(md.length?bal/md.length:0)}`);
}

// ─── Day of Week ──────────────────────────────────────────────
function botDayOfWeek(data, cmd) {
  const dayMap={'อาทิตย์':0,'sunday':0,'จันทร์':1,'monday':1,'อังคาร':2,'tuesday':2,'พุธ':3,'wednesday':3,'พฤหัส':4,'thursday':4,'พฤหัสบดี':4,'ศุกร์':5,'friday':5,'เสาร์':6,'saturday':6};
  const match=cmd.match(/\/day\s+(\S+)/i);
  if(!match) return txtMsg('❌ ระบุวัน เช่น /day จันทร์');
  const targetDay=dayMap[match[1].toLowerCase()];
  if(targetDay===undefined) return txtMsg('❌ ไม่พบวัน กรุณาระบุ: จันทร์-อาทิตย์');
  const dd=data.filter(r=>{if(!r.date)return false;return new Date(r.date).getDay()===targetDay;});
  const tn=['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  if(!dd.length) return txtMsg(`📅 วัน${tn[targetDay]}\n❌ ไม่มีข้อมูล`);
  const rev=dd.reduce((s,r)=>s+r.revenue,0),exp=dd.reduce((s,r)=>s+r.expense,0);
  const bal=dd.reduce((s,r)=>s+r.balance,0),sold=dd.reduce((s,r)=>s+r.sold,0);
  const margin=rev>0?(bal/rev*100).toFixed(1):0;
  return txtMsg(`📅 วัน${tn[targetDay]}\n\n📊 ${dd.length} วัน\n🛒 ยอดขาย: ${sold} ชิ้น\n💰 รายได้: ${fmtBot(rev)}\n💸 ค่าใช้จ่าย: ${fmtBot(exp)}\n📈 กำไร: ${fmtBot(bal)}\n🎯 Margin: ${margin}%\n📊 เฉลี่ย/วัน: ${fmtBot(dd.length?bal/dd.length:0)}`);
}

// ─── Expense Report ───────────────────────────────────────────
function botExpense(data) {
  const pf=data.reduce((s,r)=>s+r.pipeFee,0),sf=data.reduce((s,r)=>s+r.shareFee,0);
  const of=data.reduce((s,r)=>s+r.otherFee,0),sv=data.reduce((s,r)=>s+r.saveFee,0);
  const tf=pf+sf+of+sv||1;
  return txtMsg(`💸 EXPENSE BREAKDOWN\n\n📅 ${data.length} วัน\n\nPipeline: ${fmtBot(pf)} (${(pf/tf*100).toFixed(1)}%)\nShare: ${fmtBot(sf)} (${(sf/tf*100).toFixed(1)}%)\nOther: ${fmtBot(of)} (${(of/tf*100).toFixed(1)}%)\nSave: ${fmtBot(sv)} (${(sv/tf*100).toFixed(1)}%)\n\nรวม: ${fmtBot(tf)}\nเฉลี่ย/วัน: ${fmtBot(tf/data.length)}`);
}

// ─── Top/Bottom ───────────────────────────────────────────────
function botTop(data, type, cmd) {
  const limitMatch=cmd.match(/\d+/);
  const limit=limitMatch?parseInt(limitMatch[0]):5;
  const sorted=data.slice().sort((a,b)=>type==='top'?b.balance-a.balance:a.balance-b.balance);
  const topN=sorted.slice(0,limit);
  const emoji=type==='top'?'🏆':'⚠️';
  let t=`${emoji} ${type==='top'?'TOP':'BOTTOM'} ${limit}\n\n`;
  topN.forEach((r,i)=>{
    const mg=r.revenue>0?(r.balance/r.revenue*100).toFixed(1):0;
    t+=`${i+1}. 📅 ${r.date}\n   Sold:${r.sold} | รายได้:${fmtBot(r.revenue)}\n   กำไร:${fmtBot(r.balance)} | Margin:${mg}%\n`;
    if(i<topN.length-1)t+=`\n`;
  });
  return txtMsg(t);
}

// ─── Trend ────────────────────────────────────────────────────
function botTrend(data) {
  const sorted=data.slice().sort((a,b)=>a.date.localeCompare(b.date));
  const mid=Math.floor(sorted.length/2);
  const fRev=sorted.slice(0,mid).reduce((s,r)=>s+r.revenue,0);
  const sRev=sorted.slice(mid).reduce((s,r)=>s+r.revenue,0);
  const fBal=sorted.slice(0,mid).reduce((s,r)=>s+r.balance,0);
  const sBal=sorted.slice(mid).reduce((s,r)=>s+r.balance,0);
  const t=`📈 TREND ANALYSIS\n\n📅 ${data.length} วัน\n\nครึ่งแรก (${mid} วัน): รายได้ ${fmtBot(fRev)} | กำไร ${fmtBot(fBal)}\nครึ่งหลัง (${data.length-mid} วัน): รายได้ ${fmtBot(sRev)} | กำไร ${fmtBot(sBal)}\n\n${sBal>fBal?'📈 กำไรเพิ่มขึ้น':sBal<fBal?'📉 กำไรลดลง':'➡️ คงที่'} (${fBal!==0?((sBal-fBal)/Math.abs(fBal)*100).toFixed(1):0}%)`;
  return txtMsg(t);
}

// ─── Compare ──────────────────────────────────────────────────
function botCompare(data) {
  const sorted=data.slice().sort((a,b)=>a.date.localeCompare(b.date));
  const mid=Math.floor(sorted.length/2);
  const f= sorted.slice(0,mid),s=sorted.slice(mid);
  const fRev=f.reduce((s,r)=>s+r.revenue,0),sRev=s.reduce((s,r)=>s+r.revenue,0);
  const fBal=f.reduce((s,r)=>s+r.balance,0),sBal=s.reduce((s,r)=>s+r.balance,0);
  const fExp=f.reduce((s,r)=>s+r.expense,0),sExp=s.reduce((s,r)=>s+r.expense,0);
  const chRev=fRev>0?((sRev-fRev)/fRev*100).toFixed(1):0;
  const chBal=fBal>0?((sBal-fBal)/Math.abs(fBal)*100).toFixed(1):0;
  const chExp=fExp>0?((sExp-fExp)/fExp*100).toFixed(1):0;
  return txtMsg(`📊 เปรียบเทียบ\n\n📅 ครึ่งแรก (${f.length} วัน)\n💰 รายได้: ${fmtBot(fRev)}\n💸 ค่าใช้จ่าย: ${fmtBot(fExp)}\n📈 กำไร: ${fmtBot(fBal)}\n\n📅 ครึ่งหลัง (${s.length} วัน)\n💰 รายได้: ${fmtBot(sRev)}\n💸 ค่าใช้จ่าย: ${fmtBot(sExp)}\n📈 กำไร: ${fmtBot(sBal)}\n\n📈 การเปลี่ยนแปลง:\nรายได้: ${chRev>=0?'+':''}${chRev}%\nค่าใช้จ่าย: ${chExp>=0?'+':''}${chExp}%\nกำไร: ${chBal>=0?'📈 +':'📉 '}${chBal}%`);
}

// ─── Margin ───────────────────────────────────────────────────
function botMargin(data) {
  const rev=data.reduce((s,r)=>s+r.revenue,0),bal=data.reduce((s,r)=>s+r.balance,0);
  const margin=rev>0?(bal/rev*100).toFixed(1):0;
  const mm={};
  data.forEach(r=>{if(!r.date)return;const ym=r.date.slice(0,7);if(!mm[ym])mm[ym]={r:0,b:0};mm[ym].r+=r.revenue;mm[ym].b+=r.balance;});
  let t=`🎯 MARGIN ANALYSIS\n\n📊 Margin รวม: ${margin}%\n💰 รายได้: ${fmtBot(rev)}\n📈 กำไร: ${fmtBot(bal)}\n\n📅 Margin รายเดือน:\n`;
  Object.entries(mm).sort().forEach(([ym,m])=>{const mg=m.r>0?(m.b/m.r*100).toFixed(1):0;const e=mg>=30?'🟢':mg>=15?'🟡':'🔴';t+=`${e} ${ym}: ${mg}%\n`;});
  return txtMsg(t);
}

// ─── Conversion ───────────────────────────────────────────────
function botConversion(data) {
  const sold=data.reduce((s,r)=>s+r.sold,0),pend=data.reduce((s,r)=>s+r.pending,0);
  const clr=data.reduce((s,r)=>s+r.cleared,0);
  const conv=sold>0?((sold-pend)/sold*100).toFixed(1):0;
  return txtMsg(`🎯 CONVERSION\n\n🛒 ยอดขาย: ${sold} ชิ้น\n✅ ชำระแล้ว: ${clr} ชิ้น\n⏳ รอชำระ: ${pend} ชิ้น\n\n📊 Conversion: ${conv}%\nสถานะ: ${conv>=70?'🟢 ดีมาก':conv>=50?'🟡 ปานกลาง':'🔴 ต้องปรับปรุง'}`);
}

// ─── Sales ────────────────────────────────────────────────────
function botSales(data) {
  const sold=data.reduce((s,r)=>s+r.sold,0),pend=data.reduce((s,r)=>s+r.pending,0);
  return txtMsg(`🛒 SALES REPORT\n\n📅 ${data.length} วัน\n\n🛒 ขาย: ${sold} ชิ้น\n✅ ชำระ: ${sold-pend} ชิ้น\n⏳ รอ: ${pend} ชิ้น\n📊 เฉลี่ย/วัน: ${(sold/data.length).toFixed(1)} ชิ้น\n💵 มูลค่ารวม: ${fmtBot(sold*PRICE_PER_UNIT)}`);
}

// ─── Advice ───────────────────────────────────────────────────
function botAdvice(data) {
  const rev=data.reduce((s,r)=>s+r.revenue,0),bal=data.reduce((s,r)=>s+r.balance,0);
  const margin=rev>0?(bal/rev*100).toFixed(1):0;
  const sold=data.reduce((s,r)=>s+r.sold,0),pend=data.reduce((s,r)=>s+r.pending,0);
  const conv=sold>0?((sold-pend)/sold*100).toFixed(1):0;
  const dayNames=['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  const ds=Array.from({length:7},()=>({cnt:0,bal:0}));
  data.forEach(r=>{if(!r.date)return;const d=new Date(r.date).getDay();ds[d].cnt++;ds[d].bal+=r.balance;});
  const best=ds.reduce((m,s,i)=>s.cnt>0&&s.bal>ds[m].bal?i:m,0);
  let t=`💡 RECOMMENDATIONS\n\n`;
  if(margin<20)t+=`🔴 เร่งด่วน: Margin ${margin}% ต่ำกว่า 20% ควรปรับโครงสร้างต้นทุน\n\n`;
  else if(margin<30)t+=`🟡 Margin ${margin}% ควรพยายามให้เกิน 30%\n\n`;
  else t+=`🟢 Margin ${margin}% ดีมาก\n\n`;
  if(conv<60)t+=`⚠️ Conversion ${conv}% ต่ำ ควรลด Pending\n\n`;
  t+=`📈 โฟกัสวัน${dayNames[best]} (ดีที่สุด)\n💰 ตรวจสอบ Cash Flow ทุกสัปดาห์\n🎯 ตั้งเป้า KPI รายสัปดาห์`;
  return txtMsg(t);
}

// ─── Stats ────────────────────────────────────────────────────
function botStats(data) {
  const rev=data.reduce((s,r)=>s+r.revenue,0),bal=data.reduce((s,r)=>s+r.balance,0);
  const sold=data.reduce((s,r)=>s+r.sold,0),margin=rev>0?(bal/rev*100).toFixed(1):0;
  const best=data.reduce((mx,r)=>r.balance>mx.balance?r:mx,data[0]);
  const worst=data.reduce((mn,r)=>r.balance<mn.balance?r:mn,data[0]);
  return txtMsg(`📊 QUICK STATS\n\n📅 ${data.length} วัน\n💰 รายได้: ${fmtBot(rev)}\n📈 กำไร: ${fmtBot(bal)} (${margin}%)\n🛒 ยอดขาย: ${sold} ชิ้น\n\n🏆 ดีสุด: ${best.date} (${fmtBot(best.balance)})\n⚠️ แย่สุด: ${worst.date} (${fmtBot(worst.balance)})`);
}

// ─── Latest ───────────────────────────────────────────────────
function botLatest(data, cmd) {
  const m=cmd.match(/\d+/);const limit=m?parseInt(m[0]):5;
  const sorted=data.slice().sort((a,b)=>b.date.localeCompare(a.date)).slice(0,limit);
  let t=`📋 ${limit} รายการล่าสุด\n\n`;
  sorted.forEach((r,i)=>{
    const mg=r.revenue>0?(r.balance/r.revenue*100).toFixed(1):0;
    t+=`${i+1}. 📅 ${r.date}\n   Sold:${r.sold} Pend:${r.pending} Clr:${r.cleared}\n   รายได้:${fmtBot(r.revenue)} | กำไร:${fmtBot(r.balance)} (${mg}%)\n`;
    if(i<sorted.length-1)t+=`\n`;
  });
  return txtMsg(t);
}

// ─── Search ───────────────────────────────────────────────────
function botSearch(data, cmd) {
  const q=cmd.replace(/^\/(search|ค้นหา)\s+/i,'').trim();
  if(!q) return txtMsg('🔍 กรุณาระบุคำค้นหา เช่น /search 2024-01');
  const results=data.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(q.toLowerCase()))).slice(0,10);
  if(!results.length) return txtMsg(`🔍 "${q}"\n❌ ไม่พบข้อมูล`);
  let t=`🔍 "${q}" — ${results.length} รายการ\n\n`;
  results.forEach((r,i)=>{
    const mg=r.revenue>0?(r.balance/r.revenue*100).toFixed(1):0;
    t+=`${i+1}. 📅 ${r.date}\n   Sold:${r.sold} Pend:${r.pending} Clr:${r.cleared}\n   รายได้:${fmtBot(r.revenue)} | กำไร:${fmtBot(r.balance)} (${mg}%)\n`;
    if(i<results.length-1)t+=`\n`;
  });
  return txtMsg(t);
}

// ═══════════════════════════════════════════════════════════════
// PROFESSIONAL ANALYSIS FUNCTIONS
// ═══════════════════════════════════════════════════════════════

// ─── 1. Business Forecast (พยากรณ์) ──────────────────────────
function botForecast(data) {
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  
  // Simple Linear Regression for forecasting
  const n = sorted.length;
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
  
  sorted.forEach((r, i) => {
    sumX += i;
    sumY += r.balance;
    sumXY += i * r.balance;
    sumX2 += i * i;
  });
  
  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;
  
  // Forecast next 7 days
  const forecasts = [];
  for (let i = n; i < n + 7; i++) {
    forecasts.push({ day: i - n + 1, value: slope * i + intercept });
  }
  
  // Calculate accuracy
  let totalError = 0;
  sorted.forEach((r, i) => {
    const predicted = slope * i + intercept;
    totalError += Math.abs(r.balance - predicted);
  });
  const mae = totalError / n;
  const accuracy = Math.max(0, 100 - (mae / (Math.abs(sumY/n) || 1) * 100));
  
  const trend = slope > 0 ? '📈 ขาขึ้น' : slope < 0 ? '📉 ขาลง' : '➡️ คงที่';
  const confidence = accuracy > 80 ? '🟢 สูง' : accuracy > 60 ? '🟡 ปานกลาง' : '🔴 ต่ำ';
  
  let t = `🔮 BUSINESS FORECAST\n\n`;
  t += `📊 ข้อมูล ${n} วัน\n`;
  t += `📈 แนวโน้ม: ${trend} (${slope >= 0 ? '+' : ''}${slope.toFixed(1)} บาท/วัน)\n`;
  t += `🎯 ความแม่นยำ: ${confidence} (${accuracy.toFixed(1)}%)\n\n`;
  t += `📅 พยากรณ์ 7 วันข้างหน้า:\n`;
  t += `━━━━━━━━━━━━━━━━\n`;
  
  forecasts.forEach((f, i) => {
    const emoji = f.value >= 0 ? '📈' : '📉';
    const dayName = ['พรุ่งนี้', 'มะรืนนี้', 'อีก 3 วัน', 'อีก 4 วัน', 'อีก 5 วัน', 'อีก 6 วัน', 'อีก 7 วัน'];
    t += `${emoji} ${dayName[i]}: ${fmtBot(f.value)}\n`;
  });
  
  t += `━━━━━━━━━━━━━━━━\n`;
  t += `💡 คำแนะนำ:\n`;
  if (slope > 0) {
    t += `✅ แนวโน้มดี ควรลงทุนเพิ่ม\n`;
    t += `✅ เตรียมทรัพยากรรองรับการเติบโต\n`;
  } else if (slope < 0) {
    t += `⚠️ แนวโน้มลดลง ควร:\n`;
    t += `   • ลดค่าใช้จ่ายที่ไม่จำเป็น\n`;
    t += `   • เพิ่มกลยุทธ์การตลาด\n`;
    t += `   • ตรวจสอบสาเหตุการลดลง\n`;
  }
  
  return txtMsg(t);
}

// ─── 2. Risk Analysis ────────────────────────────────────────
function botRiskAnalysis(data) {
  const balances = data.map(r => r.balance);
  const avg = balances.reduce((s, v) => s + v, 0) / balances.length;
  const variance = balances.reduce((s, v) => s + Math.pow(v - avg, 2), 0) / balances.length;
  const stdDev = Math.sqrt(variance);
  const cv = avg !== 0 ? (stdDev / Math.abs(avg) * 100) : 0; // Coefficient of Variation
  
  const negDays = data.filter(r => r.balance < 0).length;
  const negRate = (negDays / data.length * 100);
  
  // Value at Risk (VaR) - simple
  const sortedBal = balances.slice().sort((a, b) => a - b);
  const var95Index = Math.floor(data.length * 0.05);
  const var95 = sortedBal[var95Index] || 0;
  
  // Maximum Drawdown
  let maxDrawdown = 0;
  let peak = balances[0];
  let currentDrawdown = 0;
  balances.forEach(b => {
    if (b > peak) {
      peak = b;
      currentDrawdown = 0;
    } else {
      currentDrawdown = peak - b;
      if (currentDrawdown > maxDrawdown) maxDrawdown = currentDrawdown;
    }
  });
  const maxDdPercent = peak > 0 ? (maxDrawdown / peak * 100) : 0;
  
  // Risk Score
  let riskScore = 0;
  if (cv > 100) riskScore += 3;
  else if (cv > 50) riskScore += 2;
  else if (cv > 20) riskScore += 1;
  
  if (negRate > 30) riskScore += 3;
  else if (negRate > 15) riskScore += 2;
  else if (negRate > 5) riskScore += 1;
  
  if (maxDdPercent > 50) riskScore += 3;
  else if (maxDdPercent > 25) riskScore += 2;
  else if (maxDdPercent > 10) riskScore += 1;
  
  const riskLevel = riskScore >= 7 ? '🔴 สูงมาก' : riskScore >= 4 ? '🟠 สูง' : riskScore >= 2 ? '🟡 ปานกลาง' : '🟢 ต่ำ';
  
  let t = `⚠️ RISK ANALYSIS\n\n`;
  t += `📊 Risk Score: ${riskScore}/9 — ${riskLevel}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📈 ความผันผวน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• ส่วนเบี่ยงเบน: ${fmtBot(stdDev)}\n`;
  t += `• ค่าสัมประสิทธิ์ความแปรปรวน: ${cv.toFixed(1)}%\n`;
  t += `• VaR (95%): ${fmtBot(var95)}\n`;
  t += `(โอกาส 5% ที่จะขาดทุนมากกว่า ${fmtBot(Math.abs(var95))})\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📉 ความเสี่ยง\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• วันที่ขาดทุน: ${negDays}/${data.length} (${negRate.toFixed(1)}%)\n`;
  t += `• ขาดทุนสูงสุดต่อเนื่อง: ${fmtBot(maxDrawdown)} (${maxDdPercent.toFixed(1)}%)\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `💡 การจัดการความเสี่ยง\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  
  if (riskScore >= 7) {
    t += `🔴 ความเสี่ยงสูงมาก!\n`;
    t += `• หยุดการลงทุนเพิ่มทันที\n`;
    t += `• ทบทวนกลยุทธ์ทั้งหมด\n`;
    t += `• ตั้ง Stop Loss ที่ ${fmtBot(var95)}\n`;
  } else if (riskScore >= 4) {
    t += `🟠 ควรระมัดระวัง\n`;
    t += `• จำกัดการลงทุนเพิ่ม\n`;
    t += `• ติดตามทุกวัน\n`;
  } else {
    t += `🟢 ความเสี่ยงต่ำ\n`;
    t += `• ดำเนินธุรกิจต่อไปได้\n`;
    t += `• รักษามาตรฐาน\n`;
  }
  
  return txtMsg(t);
}

// ─── 3. Efficiency Analysis ──────────────────────────────────
function botEfficiency(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  
  // Revenue per day
  const revPerDay = totalRev / data.length;
  
  // Expense ratio
  const expenseRatio = totalRev > 0 ? (totalExp / totalRev * 100) : 0;
  
  // Revenue per sold unit
  const revPerUnit = totalSold > 0 ? (totalRev / totalSold) : 0;
  
  // Collection efficiency
  const collectionRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;
  
  // Pipeline efficiency
  const pipeFee = data.reduce((s, r) => s + r.pipeFee, 0);
  const pipeEfficiency = totalRev > 0 ? (pipeFee / totalRev * 100) : 0;
  
  // Score
  let effScore = 0;
  if (expenseRatio < 30) effScore += 3;
  else if (expenseRatio < 50) effScore += 2;
  else effScore += 1;
  
  if (collectionRate > 80) effScore += 3;
  else if (collectionRate > 60) effScore += 2;
  else effScore += 1;
  
  if (pipeEfficiency < 20) effScore += 3;
  else if (pipeEfficiency < 40) effScore += 2;
  else effScore += 1;
  
  const effLevel = effScore >= 7 ? '🟢 ดีเยี่ยม' : effScore >= 5 ? '🟡 ปานกลาง' : '🔴 ต้องปรับปรุง';
  
  let t = `⚡ EFFICIENCY ANALYSIS\n\n`;
  t += `📊 Efficiency Score: ${effScore}/9 — ${effLevel}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📈 ตัวชี้วัดประสิทธิภาพ\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• รายได้เฉลี่ย/วัน: ${fmtBot(revPerDay)}\n`;
  t += `• อัตราค่าใช้จ่าย: ${expenseRatio.toFixed(1)}%\n`;
  t += `• รายได้ต่อชิ้น: ${fmtBot(revPerUnit)}\n`;
  t += `• ประสิทธิภาพการเก็บเงิน: ${collectionRate.toFixed(1)}%\n`;
  t += `• Pipeline Cost Ratio: ${pipeEfficiency.toFixed(1)}%\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `💡 จุดที่ควรปรับปรุง\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  
  if (expenseRatio > 50) t += `⚠️ ค่าใช้จ่ายสูงเกิน 50% ของรายได้\n`;
  if (collectionRate < 60) t += `⚠️ อัตราการเก็บเงินต่ำกว่า 60%\n`;
  if (pipeEfficiency > 40) t += `⚠️ Pipeline Fee สูงเกิน 40%\n`;
  
  if (effScore >= 7) t += `✅ ทุกอย่างมีประสิทธิภาพดี\n`;
  
  return txtMsg(t);
}

// ─── 4. Benchmark Analysis ───────────────────────────────────
function botBenchmark(data) {
  // Industry benchmarks (sample values)
  const benchmarks = {
    profitMargin: { target: 30, good: 25, average: 15 },
    conversionRate: { target: 80, good: 70, average: 50 },
    expenseRatio: { target: 25, good: 35, average: 50 },
    dailyRevenue: { target: 1500, good: 1000, average: 500 },
    collectionRate: { target: 90, good: 75, average: 60 }
  };
  
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  
  const profitMargin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;
  const expenseRatio = totalRev > 0 ? (totalExp / totalRev * 100) : 0;
  const dailyRevenue = data.length > 0 ? totalRev / data.length : 0;
  const conversionRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;
  const collectionRate = conversionRate;
  
  function compareBenchmark(value, benchmark) {
    if (value >= benchmark.target) return '🟢 ดีเยี่ยม';
    if (value >= benchmark.good) return '🟡 ดี';
    if (value >= benchmark.average) return '🟠 ปานกลาง';
    return '🔴 ต่ำกว่ามาตรฐาน';
  }
  
  function compareBenchmarkReverse(value, benchmark) {
    if (value <= benchmark.target) return '🟢 ดีเยี่ยม';
    if (value <= benchmark.good) return '🟡 ดี';
    if (value <= benchmark.average) return '🟠 ปานกลาง';
    return '🔴 สูงเกินไป';
  }
  
  let t = `📊 BENCHMARK ANALYSIS\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `ตัวชี้วัด          | ค่าของคุณ | เกณฑ์มาตรฐาน | สถานะ\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `Profit Margin     | ${profitMargin.toFixed(1)}%     | ≥${benchmarks.profitMargin.target}%         | ${compareBenchmark(profitMargin, benchmarks.profitMargin)}\n`;
  t += `Expense Ratio     | ${expenseRatio.toFixed(1)}%     | ≤${benchmarks.expenseRatio.target}%         | ${compareBenchmarkReverse(expenseRatio, benchmarks.expenseRatio)}\n`;
  t += `Conversion Rate   | ${conversionRate.toFixed(1)}%     | ≥${benchmarks.conversionRate.target}%         | ${compareBenchmark(conversionRate, benchmarks.conversionRate)}\n`;
  t += `Revenue/Day       | ${fmtBot(dailyRevenue)}   | ≥${fmtBot(benchmarks.dailyRevenue.target)}     | ${compareBenchmark(dailyRevenue, benchmarks.dailyRevenue)}\n`;
  t += `Collection Rate   | ${collectionRate.toFixed(1)}%     | ≥${benchmarks.collectionRate.target}%         | ${compareBenchmark(collectionRate, benchmarks.collectionRate)}\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  // Overall score
  let scores = 0;
  if (profitMargin >= benchmarks.profitMargin.target) scores++;
  if (expenseRatio <= benchmarks.expenseRatio.good) scores++;
  if (conversionRate >= benchmarks.conversionRate.good) scores++;
  if (dailyRevenue >= benchmarks.dailyRevenue.average) scores++;
  if (collectionRate >= benchmarks.collectionRate.good) scores++;
  
  t += `📊 Overall: ${scores}/5 ผ่านเกณฑ์\n`;
  
  if (scores >= 4) t += `✅ ธุรกิจคุณแข็งแกร่งเมื่อเทียบกับมาตรฐานอุตสาหกรรม\n`;
  else if (scores >= 2) t += `🟡 มีจุดที่ต้องปรับปรุงบ้าง\n`;
  else t += `🔴 ต้องปรับปรุงหลายด้านเมื่อเทียบกับมาตรฐาน\n`;
  
  return txtMsg(t);
}

// ─── 5. Cash Flow Analysis ───────────────────────────────────
function botCashflow(data) {
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  
  let totalInflow = 0, totalOutflow = 0;
  let maxInflow = -Infinity, maxOutflow = -Infinity;
  let runningBalance = 0;
  let minRunningBalance = Infinity;
  let positiveDays = 0, negativeDays = 0;
  
  const weeklyCashflow = {};
  
  sorted.forEach(r => {
    const inflow = r.revenue;
    const outflow = r.expense;
    
    totalInflow += inflow;
    totalOutflow += outflow;
    runningBalance += (inflow - outflow);
    
    if (inflow > maxInflow) maxInflow = inflow;
    if (outflow > maxOutflow) maxOutflow = outflow;
    if (runningBalance < minRunningBalance) minRunningBalance = runningBalance;
    
    if (r.balance >= 0) positiveDays++;
    else negativeDays++;
    
    // Weekly aggregation
    const week = getWeekNumber(r.date);
    if (!weeklyCashflow[week]) weeklyCashflow[week] = { inflow: 0, outflow: 0, balance: 0 };
    weeklyCashflow[week].inflow += inflow;
    weeklyCashflow[week].outflow += outflow;
    weeklyCashflow[week].balance += r.balance;
  });
  
  const netCashflow = totalInflow - totalOutflow;
  const avgInflow = data.length > 0 ? totalInflow / data.length : 0;
  const avgOutflow = data.length > 0 ? totalOutflow / data.length : 0;
  const cashflowRatio = totalOutflow > 0 ? (totalInflow / totalOutflow) : 0;
  
  let t = `💵 CASH FLOW ANALYSIS\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📥 เงินเข้า\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• รวม: ${fmtBot(totalInflow)}\n`;
  t += `• เฉลี่ย/วัน: ${fmtBot(avgInflow)}\n`;
  t += `• สูงสุด/วัน: ${fmtBot(maxInflow)}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📤 เงินออก\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• รวม: ${fmtBot(totalOutflow)}\n`;
  t += `• เฉลี่ย/วัน: ${fmtBot(avgOutflow)}\n`;
  t += `• สูงสุด/วัน: ${fmtBot(maxOutflow)}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📊 สถานะกระแสเงินสด\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• กระแสเงินสดสุทธิ: ${fmtBot(netCashflow)}\n`;
  t += `• อัตราส่วนเข้า/ออก: ${cashflowRatio.toFixed(2)}x\n`;
  t += `• ยอดคงเหลือต่ำสุด: ${fmtBot(minRunningBalance)}\n`;
  t += `• วันที่เป็นบวก: ${positiveDays} (${(positiveDays/data.length*100).toFixed(1)}%)\n`;
  t += `• วันที่ติดลบ: ${negativeDays} (${(negativeDays/data.length*100).toFixed(1)}%)\n\n`;
  t += `💡 คำแนะนำ:\n`;
  
  if (cashflowRatio < 1.2) {
    t += `🔴 กระแสเงินสดเสี่ยง! รายได้ใกล้เคียงรายจ่าย\n`;
    t += `• ต้องเพิ่มรายได้หรือลดรายจ่าย\n`;
    t += `• ควรมีเงินสำรองอย่างน้อย ${fmtBot(totalOutflow * 2)}\n`;
  } else if (cashflowRatio < 2) {
    t += `🟡 กระแสเงินสดใช้ได้ แต่ควรปรับปรุง\n`;
  } else {
    t += `🟢 กระแสเงินสดดีมาก\n`;
  }
  
  return txtMsg(t);
}

function getWeekNumber(dateStr) {
  const d = new Date(dateStr);
  const start = new Date(d.getFullYear(), 0, 1);
  const diff = d - start;
  return `${d.getFullYear()}-W${Math.floor(diff / (7 * 24 * 60 * 60 * 1000)) + 1}`;
}

// ─── 6. Break-even Analysis ──────────────────────────────────
function botBreakeven(data) {
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  
  // Fixed vs Variable costs (estimate: 60% fixed, 40% variable)
  const fixedCosts = totalExp * 0.6 / data.length; // per day
  const variableCostPerUnit = (totalExp * 0.4) / (totalSold || 1);
  const pricePerUnit = PRICE_PER_UNIT;
  const contributionMargin = pricePerUnit - variableCostPerUnit;
  
  // Break-even point in units per day
  const bepUnits = contributionMargin > 0 ? Math.ceil(fixedCosts / contributionMargin) : Infinity;
  const bepRevenue = bepUnits * pricePerUnit;
  
  // Current performance
  const avgUnitsPerDay = totalSold / data.length;
  const avgRevenuePerDay = totalRev / data.length;
  const margin = contributionMargin > 0 ? ((avgRevenuePerDay - fixedCosts) / avgRevenuePerDay * 100) : 0;
  
  let t = `🎯 BREAK-EVEN ANALYSIS\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📊 ข้อมูลพื้นฐาน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• ราคาขาย/หน่วย: ${fmtBot(pricePerUnit)}\n`;
  t += `• ต้นทุนคงที่/วัน: ${fmtBot(fixedCosts)} (ประมาณ)\n`;
  t += `• ต้นทุนผันแปร/หน่วย: ${fmtBot(variableCostPerUnit)} (ประมาณ)\n`;
  t += `• กำไรส่วนเกิน/หน่วย: ${fmtBot(contributionMargin)}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `🎯 จุดคุ้มทุน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• จำนวนขาย/วัน: ${bepUnits === Infinity ? '∞' : bepUnits} ชิ้น\n`;
  t += `• รายได้/วัน: ${fmtBot(bepRevenue)}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📈 ปัจจุบัน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• ขายเฉลี่ย/วัน: ${avgUnitsPerDay.toFixed(1)} ชิ้น\n`;
  t += `• รายได้เฉลี่ย/วัน: ${fmtBot(avgRevenuePerDay)}\n`;
  t += `• Margin: ${margin.toFixed(1)}%\n\n`;
  
  if (avgUnitsPerDay >= bepUnits) {
    t += `✅ **ถึงจุดคุ้มทุนแล้ว!**\n`;
    t += `• เกินจุดคุ้มทุน ${(avgUnitsPerDay - bepUnits).toFixed(1)} ชิ้น/วัน\n`;
    t += `• กำไรต่อวัน: ${fmtBot((avgUnitsPerDay - bepUnits) * contributionMargin)}\n`;
  } else if (bepUnits !== Infinity) {
    t += `❌ **ยังไม่ถึงจุดคุ้มทุน**\n`;
    t += `• ต้องเพิ่มอีก ${(bepUnits - avgUnitsPerDay).toFixed(0)} ชิ้น/วัน\n`;
    t += `• ต้องเพิ่มรายได้อีก ${fmtBot(bepRevenue - avgRevenuePerDay)}/วัน\n`;
  }
  
  return txtMsg(t);
}

// ─── 7. Growth Analysis ──────────────────────────────────────
function botGrowth(data) {
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  
  // Monthly growth
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
  
  let t = `📈 GROWTH ANALYSIS\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `เดือน        | รายได้     | Growth   | กำไร      | Growth\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  
  let prevRev = 0, prevBal = 0;
  months.forEach(([ym, m]) => {
    const revGrowth = prevRev > 0 ? ((m.revenue - prevRev) / prevRev * 100).toFixed(1) : '—';
    const balGrowth = prevBal > 0 ? ((m.balance - prevBal) / Math.abs(prevBal) * 100).toFixed(1) : '—';
    
    const revGrowthStr = revGrowth === '—' ? '   —   ' : 
      (revGrowth >= 0 ? `📈 +${revGrowth}%` : `📉 ${revGrowth}%`);
    const balGrowthStr = balGrowth === '—' ? '   —   ' : 
      (balGrowth >= 0 ? `📈 +${balGrowth}%` : `📉 ${balGrowth}%`);
    
    t += `${ym.padEnd(13)}| ${fmtBot(m.revenue).padEnd(10)}| ${revGrowthStr.padEnd(10)}| ${fmtBot(m.balance).padEnd(10)}| ${balGrowthStr}\n`;
    
    prevRev = m.revenue;
    prevBal = m.balance;
  });
  
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  // CAGR (Compound Annual Growth Rate)
  if (months.length >= 2) {
    const firstMonth = months[0][1];
    const lastMonth = months[months.length - 1][1];
    const totalMonths = months.length;
    const cagrRev = firstMonth.revenue > 0 ? 
      (Math.pow(lastMonth.revenue / firstMonth.revenue, 12/totalMonths) - 1) * 100 : 0;
    const cagrBal = Math.abs(firstMonth.balance) > 0 ? 
      (Math.pow(Math.abs(lastMonth.balance) / Math.abs(firstMonth.balance), 12/totalMonths) - 1) * 100 : 0;
    
    t += `📊 Annualized Growth Rate:\n`;
    t += `• Revenue CAGR: ${cagrRev.toFixed(1)}%\n`;
    t += `• Balance CAGR: ${cagrBal.toFixed(1)}%\n`;
  }
  
  return txtMsg(t);
}

// ─── 8. Seasonality Analysis ─────────────────────────────────
function botSeasonality(data) {
  const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  const monthlyPattern = {};
  const dailyPattern = Array.from({ length: 7 }, () => ({ count: 0, total: 0 }));
  
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
  
  let t = `🌤️ SEASONALITY ANALYSIS\n\n`;
  
  // Daily pattern
  t += `📅 รูปแบบรายวัน:\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  dailyPattern.forEach((s, i) => {
    const avg = s.count > 0 ? s.total / s.count : 0;
    const diff = avgAll > 0 ? ((avg - avgAll) / avgAll * 100).toFixed(0) : 0;
    const bar = '█'.repeat(Math.max(0, Math.min(10, Math.round(avg / (avgAll || 1) * 5))));
    const indicator = avg > avgAll ? '📈' : avg < avgAll ? '📉' : '➡️';
    t += `${indicator} ${dayNames[i].padEnd(11)}| ${bar.padEnd(10)}| ${fmtBot(avg)} (${diff >= 0 ? '+' : ''}${diff}%)\n`;
  });
  
  // Monthly pattern
  const thM = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  t += `\n📆 รูปแบบรายเดือน:\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  
  const monthEntries = Object.entries(monthlyPattern).sort((a, b) => a[0] - b[0]);
  monthEntries.forEach(([m, s]) => {
    const avg = s.count > 0 ? s.total / s.count : 0;
    const diff = avgAll > 0 ? ((avg - avgAll) / avgAll * 100).toFixed(0) : 0;
    const bar = '█'.repeat(Math.max(0, Math.min(10, Math.round(Math.abs(avg / (avgAll || 1)) * 5))));
    const indicator = avg > avgAll ? '📈' : avg < avgAll ? '📉' : '➡️';
    t += `${indicator} ${thM[parseInt(m)].padEnd(10)}| ${bar.padEnd(10)}| ${fmtBot(avg)} (${diff >= 0 ? '+' : ''}${diff}%)\n`;
  });
  
  t += `\n💡 ข้อมูลเชิงลึก:\n`;
  
  // Find peaks
  const maxDay = dailyPattern.reduce((m, s, i) => s.total > dailyPattern[m].total ? i : m, 0);
  const minDay = dailyPattern.reduce((m, s, i) => s.total < dailyPattern[m].total ? i : m, 0);
  
  t += `• วันที่ดีที่สุด: ${dayNames[maxDay]}\n`;
  t += `• วันที่ควรปรับ: ${dayNames[minDay]}\n`;
  t += `• แนะนำ: เน้นการตลาดวัน${dayNames[maxDay]}\n`;
  
  return txtMsg(t);
}

// ─── 9. Profit & Loss Statement ──────────────────────────────
function botProfitLoss(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  const totalClr = data.reduce((s, r) => s + r.cleared, 0);
  
  const pipeFee = data.reduce((s, r) => s + r.pipeFee, 0);
  const shareFee = data.reduce((s, r) => s + r.shareFee, 0);
  const otherFee = data.reduce((s, r) => s + r.otherFee, 0);
  const saveFee = data.reduce((s, r) => s + r.saveFee, 0);
  
  const totalExp = pipeFee + shareFee + otherFee + saveFee;
  const grossProfit = totalRev - totalExp;
  const netProfit = grossProfit; // Simplified
  
  const margin = totalRev > 0 ? (netProfit / totalRev * 100) : 0;
  
  let t = `📊 PROFIT & LOSS STATEMENT\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `รายการ                        | จำนวน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📥 รายได้\n`;
  t += `  รายได้จากการขาย            | ${fmtBot(totalRev)}\n`;
  t += `  (${totalSold} ชิ้น @ ${fmtBot(PRICE_PER_UNIT)})    |\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📤 ค่าใช้จ่าย\n`;
  t += `  Pipeline Fee                | ${fmtBot(pipeFee)}\n`;
  t += `  Share Fee                   | ${fmtBot(shareFee)}\n`;
  t += `  Other Fee                   | ${fmtBot(otherFee)}\n`;
  t += `  Save Fee                    | ${fmtBot(saveFee)}\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `  รวมค่าใช้จ่าย               | ${fmtBot(totalExp)}\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📈 กำไรสุทธิ                  | ${fmtBot(netProfit)}\n`;
  t += `🎯 Profit Margin              | ${margin.toFixed(1)}%\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  t += `📅 ${data.length} วัน\n`;
  t += `🛒 ขาย ${totalSold} | ⏳ รอ ${totalPend} | ✅ ชำระ ${totalClr}\n`;
  
  return txtMsg(t);
}

// ─── 10. ROI Analysis ────────────────────────────────────────
function botROI(data) {
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  
  const totalInvestment = totalExp; // Simplification
  const totalReturn = totalBal + totalExp; // Revenue
  const roi = totalInvestment > 0 ? (totalBal / totalInvestment * 100) : 0;
  
  // Per day ROI
  const avgInvestment = data.length > 0 ? totalInvestment / data.length : 0;
  const avgReturn = data.length > 0 ? totalBal / data.length : 0;
  const dailyROI = avgInvestment > 0 ? (avgReturn / avgInvestment * 100) : 0;
  
  // Payback period
  const paybackDays = avgReturn > 0 ? avgInvestment / avgReturn : Infinity;
  
  let t = `💰 ROI ANALYSIS\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• เงินลงทุนรวม: ${fmtBot(totalInvestment)}\n`;
  t += `• ผลตอบแทนรวม: ${fmtBot(totalReturn)}\n`;
  t += `• กำไรสุทธิ: ${fmtBot(totalBal)}\n`;
  t += `• ROI: ${roi.toFixed(1)}%\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `📊 ต่อวัน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `• ลงทุน/วัน: ${fmtBot(avgInvestment)}\n`;
  t += `• กำไร/วัน: ${fmtBot(avgReturn)}\n`;
  t += `• Daily ROI: ${dailyROI.toFixed(1)}%\n`;
  t += `• ระยะคืนทุน: ${paybackDays === Infinity ? '∞' : paybackDays.toFixed(1)} วัน\n\n`;
  t += `💡 ${roi >= 50 ? '✅ ROI ดีมาก!' : roi >= 20 ? '🟡 ROI ปานกลาง' : '🔴 ROI ต่ำ ควรปรับปรุง'}\n`;
  
  return txtMsg(t);
}

// ─── 11. Business Health Score ───────────────────────────────
function botHealthScore(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  
  let score = 0;
  const details = [];
  
  // 1. Profit Margin (0-25 points)
  const margin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;
  const marginScore = margin >= 40 ? 25 : margin >= 30 ? 20 : margin >= 20 ? 15 : margin >= 10 ? 10 : 5;
  score += marginScore;
  details.push({ name: 'Profit Margin', value: `${margin.toFixed(1)}%`, score: marginScore, max: 25 });
  
  // 2. Consistency (0-20 points)
  const posDays = data.filter(r => r.balance >= 0).length;
  const consistency = (posDays / data.length * 100);
  const consistencyScore = consistency >= 90 ? 20 : consistency >= 75 ? 15 : consistency >= 60 ? 10 : 5;
  score += consistencyScore;
  details.push({ name: 'ความสม่ำเสมอ', value: `${consistency.toFixed(1)}%`, score: consistencyScore, max: 20 });
  
  // 3. Conversion Rate (0-20 points)
  const convRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;
  const convScore = convRate >= 80 ? 20 : convRate >= 65 ? 15 : convRate >= 50 ? 10 : 5;
  score += convScore;
  details.push({ name: 'Conversion Rate', value: `${convRate.toFixed(1)}%`, score: convScore, max: 20 });
  
  // 4. Growth Trend (0-20 points)
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length / 2);
  const fBal = sorted.slice(0, mid).reduce((s, r) => s + r.balance, 0);
  const sBal = sorted.slice(mid).reduce((s, r) => s + r.balance, 0);
  const growthRate = fBal !== 0 ? ((sBal - fBal) / Math.abs(fBal) * 100) : 0;
  const growthScore = growthRate > 20 ? 20 : growthRate > 10 ? 15 : growthRate > 0 ? 10 : growthRate > -10 ? 5 : 0;
  score += growthScore;
  details.push({ name: 'การเติบโต', value: `${growthRate >= 0 ? '+' : ''}${growthRate.toFixed(1)}%`, score: growthScore, max: 20 });
  
  // 5. Expense Control (0-15 points)
  const expenseRatio = totalRev > 0 ? (totalExp / totalRev * 100) : 100;
  const expenseScore = expenseRatio < 30 ? 15 : expenseRatio < 50 ? 10 : expenseRatio < 70 ? 5 : 0;
  score += expenseScore;
  details.push({ name: 'ควบคุมค่าใช้จ่าย', value: `${expenseRatio.toFixed(1)}%`, score: expenseScore, max: 15 });
  
  const grade = score >= 90 ? 'A+ 🌟' : score >= 80 ? 'A' : score >= 70 ? 'B+' : score >= 60 ? 'B' : score >= 50 ? 'C' : 'D';
  const status = score >= 80 ? '🟢 แข็งแกร่ง' : score >= 60 ? '🟡 ปานกลาง' : '🔴 ต้องปรับปรุง';
  
  let t = `🏥 BUSINESS HEALTH SCORE\n\n`;
  t += `คะแนน: ${score}/100 — ${grade}\n`;
  t += `สถานะ: ${status}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `ตัวชี้วัด              | คะแนน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  
  details.forEach(d => {
    const bar = '█'.repeat(d.score) + '░'.repeat(d.max - d.score);
    t += `${d.name.padEnd(21)}| ${bar} ${d.score}/${d.max}\n`;
    t += `${''.padEnd(21)}| ${d.value}\n`;
  });
  
  t += `━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  if (score >= 80) t += `✅ สุขภาพธุรกิจดีเยี่ยม! รักษามาตรฐานนี้ไว้\n`;
  else if (score >= 60) t += `🟡 สุขภาพธุรกิจใช้ได้ มีจุดที่ต้องปรับปรุง\n`;
  else t += `🔴 ต้องปรับปรุงเร่งด่วน! ปรึกษาผู้เชี่ยวชาญ\n`;
  
  t += `\n💡 แนะนำ: ตรวจสุขภาพธุรกิจทุกสัปดาห์ด้วย /health`;
  
  return txtMsg(t);
}

// ─── 12. Executive Summary ───────────────────────────────────
function botExecutiveSummary(data) {
  const rev = data.reduce((s,r)=>s+r.revenue,0);
  const exp = data.reduce((s,r)=>s+r.expense,0);
  const bal = data.reduce((s,r)=>s+r.balance,0);
  const sold = data.reduce((s,r)=>s+r.sold,0);
  const pend = data.reduce((s,r)=>s+r.pending,0);
  const margin = rev>0?(bal/rev*100):0;
  const conv = sold>0?((sold-pend)/sold*100):0;
  
  // Best/worst
  const best = data.reduce((mx,r)=>r.balance>mx.balance?r:mx,data[0]);
  const worst = data.reduce((mn,r)=>r.balance<mn.balance?r:mn,data[0]);
  
  // Trend
  const sorted = data.slice().sort((a,b)=>a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length/2);
  const fBal = sorted.slice(0,mid).reduce((s,r)=>s+r.balance,0);
  const sBal = sorted.slice(mid).reduce((s,r)=>s+r.balance,0);
  const trend = sBal>fBal?'📈 เพิ่มขึ้น':sBal<fBal?'📉 ลดลง':'➡️ คงที่';
  
  let t = `📊 EXECUTIVE SUMMARY\n`;
  t += `สำหรับผู้บริหาร — ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `ภาพรวมธุรกิจ\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `💰 รายได้: ${fmtBot(rev)}\n`;
  t += `💸 ค่าใช้จ่าย: ${fmtBot(exp)}\n`;
  t += `📈 กำไร: ${fmtBot(bal)} (${margin.toFixed(1)}%)\n`;
  t += `📊 ${data.length} วัน | 🛒 ${sold} ชิ้น\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `Key Highlights\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `🏆 ดีที่สุด: ${best.date} (${fmtBot(best.balance)})\n`;
  t += `⚠️ ควรปรับ: ${worst.date} (${fmtBot(worst.balance)})\n`;
  t += `📈 แนวโน้ม: ${trend}\n`;
  t += `✅ Conversion: ${conv.toFixed(1)}%\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `คำแนะนำสำหรับผู้บริหาร\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  
  if (margin < 20) t += `🔴 เร่งด่วน: ปรับโครงสร้างต้นทุน\n`;
  if (trend.includes('ลดลง')) t += `⚠️ แนวโน้มลดลง ต้องรีบแก้ไข\n`;
  if (conv < 60) t += `⚠️ ปรับปรุงการเก็บเงิน\n`;
  if (margin >= 30 && !trend.includes('ลดลง')) t += `✅ ธุรกิจไปได้ดี รักษามาตรฐาน\n`;
  
  return txtMsg(t);
}

// ─── 13. SWOT Analysis ───────────────────────────────────────
function botSWOT(data) {
  const rev = data.reduce((s,r)=>s+r.revenue,0);
  const bal = data.reduce((s,r)=>s+r.balance,0);
  const margin = rev>0?(bal/rev*100):0;
  const sold = data.reduce((s,r)=>s+r.sold,0);
  const pend = data.reduce((s,r)=>s+r.pending,0);
  const conv = sold>0?((sold-pend)/sold*100):0;
  
  // Strengths
  const strengths = [];
  if (margin >= 30) strengths.push('💪 Profit Margin สูง');
  if (conv >= 70) strengths.push('💪 อัตราการเก็บเงินดี');
  if (data.filter(r=>r.balance>=0).length/data.length > 0.8) strengths.push('💪 กำไรสม่ำเสมอ');
  strengths.push('💪 มีระบบติดตามข้อมูล');
  
  // Weaknesses
  const weaknesses = [];
  if (margin < 20) weaknesses.push('🔻 Profit Margin ต่ำ');
  if (conv < 60) weaknesses.push('🔻 ยอด pending สูง');
  if (data.filter(r=>r.balance<0).length/data.length > 0.2) weaknesses.push('🔻 มีวันที่ขาดทุนบ่อย');
  
  // Opportunities
  const opportunities = ['🚀 ขยายตลาดออนไลน์', '🚀 เพิ่มช่องทางขาย', '🚀 พัฒนากลยุทธ์ราคา'];
  
  // Threats
  const threats = [];
  if (margin < 15) threats.push('⚠️ เสี่ยงขาดทุนหากต้นทุนเพิ่ม');
  threats.push('⚠️ การแข่งขันสูง', '⚠️ ความผันผวนของตลาด');
  
  let t = `📊 SWOT ANALYSIS\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `💪 จุดแข็ง (Strengths)\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  strengths.forEach(s => t += `  ${s}\n`);
  t += `\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `🔻 จุดอ่อน (Weaknesses)\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  (weaknesses.length ? weaknesses : ['🔻 ยังไม่พบจุดอ่อนชัดเจน']).forEach(w => t += `  ${w}\n`);
  t += `\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `🚀 โอกาส (Opportunities)\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  opportunities.forEach(o => t += `  ${o}\n`);
  t += `\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `⚠️ อุปสรรค (Threats)\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  threats.forEach(th => t += `  ${th}\n`);
  t += `\n💡 กลยุทธ์: ใช้จุดแข็งเข้าหาโอกาส แก้ไขจุดอ่อนเพื่อลดอุปสรรค`;
  
  return txtMsg(t);
}

// ─── 14. Goal Tracking ───────────────────────────────────────
function botGoalTracking(data) {
  // Default goals
  const goals = {
    dailyRevenue: 1000,
    monthlyProfit: 20000,
    profitMargin: 30,
    conversionRate: 70,
    positiveDays: 80
  };
  
  const rev = data.reduce((s,r)=>s+r.revenue,0);
  const bal = data.reduce((s,r)=>s+r.balance,0);
  const margin = rev>0?(bal/rev*100):0;
  const sold = data.reduce((s,r)=>s+r.sold,0);
  const pend = data.reduce((s,r)=>s+r.pending,0);
  const conv = sold>0?((sold-pend)/sold*100):0;
  
  const avgDailyRev = data.length>0?rev/data.length:0;
  const avgMonthlyProfit = data.length>0?(bal/data.length*30):0;
  const posDays = data.filter(r=>r.balance>=0).length;
  const posDayRate = (posDays/data.length*100);
  
  function progressBar(current, target) {
    const pct = Math.min(100, (current/target*100));
    const filled = Math.floor(pct / 10);
    return '█'.repeat(filled) + '░'.repeat(10 - filled) + ` ${pct.toFixed(0)}%`;
  }
  
  let t = `🎯 GOAL TRACKING\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `เป้าหมาย              | ความคืบหน้า\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `รายได้/วัน ≥ ${fmtBot(goals.dailyRevenue)}\n`;
  t += `ปัจจุบัน: ${fmtBot(avgDailyRev)} | ${progressBar(avgDailyRev, goals.dailyRevenue)}\n\n`;
  t += `กำไร/เดือน ≥ ${fmtBot(goals.monthlyProfit)}\n`;
  t += `ปัจจุบัน: ${fmtBot(avgMonthlyProfit)} | ${progressBar(avgMonthlyProfit, goals.monthlyProfit)}\n\n`;
  t += `Margin ≥ ${goals.profitMargin}%\n`;
  t += `ปัจจุบัน: ${margin.toFixed(1)}% | ${progressBar(margin, goals.profitMargin)}\n\n`;
  t += `Conversion ≥ ${goals.conversionRate}%\n`;
  t += `ปัจจุบัน: ${conv.toFixed(1)}% | ${progressBar(conv, goals.conversionRate)}\n\n`;
  t += `วันเป็นบวก ≥ ${goals.positiveDays}%\n`;
  t += `ปัจจุบัน: ${posDayRate.toFixed(1)}% | ${progressBar(posDayRate, goals.positiveDays)}\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  const achieved = [avgDailyRev>=goals.dailyRevenue, avgMonthlyProfit>=goals.monthlyProfit, 
                    margin>=goals.profitMargin, conv>=goals.conversionRate, posDayRate>=goals.positiveDays]
                    .filter(Boolean).length;
  
  t += `✅ ${achieved}/5 เป้าหมายที่บรรลุ\n`;
  t += `${achieved>=4?'🌟 ยอดเยี่ยม!':achieved>=2?'🟡 กำลังไปได้ดี':'🔴 ต้องพยายามมากขึ้น'}\n`;
  t += `\n💡 ปรับเป้าหมายได้ด้วย /goal set [รายการ] [ค่า]`;
  
  return txtMsg(t);
}

// ─── 15. Outstanding Report ──────────────────────────────────
function botOutstandingReport(data) {
  const totalPend = data.reduce((s,r)=>s+r.pending,0);
  const totalSold = data.reduce((s,r)=>s+r.sold,0);
  
  // Find pending-heavy records
  const pendingRecords = data.filter(r => r.pending > 0)
    .sort((a,b) => b.pending - a.pending)
    .slice(0, 10);
  
  const totalOutstanding = totalPend * PRICE_PER_UNIT;
  const outstandingRate = totalSold > 0 ? (totalPend/totalSold*100) : 0;
  
  let t = `📋 OUTSTANDING REPORT\n\n`;
  t += `⏳ ยอดค้างชำระ: ${totalPend} ชิ้น\n`;
  t += `💰 มูลค่า: ${fmtBot(totalOutstanding)}\n`;
  t += `📊 อัตราค้างชำระ: ${outstandingRate.toFixed(1)}%\n\n`;
  
  if (pendingRecords.length) {
    t += `━━━━━━━━━━━━━━━━━━━━\n`;
    t += `รายการที่มียอดค้างสูงสุด\n`;
    t += `━━━━━━━━━━━━━━━━━━━━\n`;
    pendingRecords.forEach((r, i) => {
      t += `${i+1}. 📅 ${r.date} | ⏳ ${r.pending} ชิ้น | 💰 ${fmtBot(r.pending * PRICE_PER_UNIT)}\n`;
    });
  }
  
  t += `\n💡 แนะนำ:\n`;
  if (outstandingRate > 30) {
    t += `🔴 อัตราค้างชำระสูง! ควร:\n`;
    t += `• ส่งใบแจ้งหนี้\n`;
    t += `• ติดตามลูกหนี้\n`;
    t += `• กำหนด Credit Term\n`;
  } else {
    t += `✅ อัตราค้างชำระอยู่ในเกณฑ์ดี\n`;
  }
  
  return txtMsg(t);
}

// ─── 16. Projection ──────────────────────────────────────────
function botProjection(data) {
  const avgRev = data.reduce((s,r)=>s+r.revenue,0)/data.length;
  const avgExp = data.reduce((s,r)=>s+r.expense,0)/data.length;
  const avgBal = data.reduce((s,r)=>s+r.balance,0)/data.length;
  const avgSold = data.reduce((s,r)=>s+r.sold,0)/data.length;
  
  // 3 scenarios
  const scenarios = [
    { name: '🔴 Conservative', factor: 0.8 },
    { name: '🟡 Base Case', factor: 1.0 },
    { name: '🟢 Optimistic', factor: 1.2 }
  ];
  
  const days = [7, 30, 90];
  
  let t = `📊 PROJECTION (ประมาณการ)\n\n`;
  t += `จากข้อมูล ${data.length} วัน\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  scenarios.forEach(sc => {
    const r = avgRev * sc.factor;
    const e = avgExp * sc.factor;
    const b = r - e;
    
    t += `${sc.name}:\n`;
    days.forEach(d => {
      t += `  ${d} วัน: รายได้ ${fmtBot(r*d)} | กำไร ${fmtBot(b*d)}\n`;
    });
    t += `\n`;
  });
  
  t += `💡 ข้อสมมติ:\n`;
  t += `• ราคาต่อหน่วย: ${fmtBot(PRICE_PER_UNIT)}\n`;
  t += `• ขายเฉลี่ย/วัน: ${avgSold.toFixed(1)} ชิ้น\n`;
  t += `• อิงจากข้อมูลย้อนหลัง ${data.length} วัน\n`;
  
  return txtMsg(t);
}

// ─── 17. Financial Ratios ────────────────────────────────────
function botFinancialRatios(data) {
  const rev = data.reduce((s,r)=>s+r.revenue,0);
  const exp = data.reduce((s,r)=>s+r.expense,0);
  const bal = data.reduce((s,r)=>s+r.balance,0);
  
  // Profitability Ratios
  const grossMargin = rev>0?(bal/rev*100):0;
  const expenseRatio = rev>0?(exp/rev*100):0;
  const netMargin = grossMargin;
  
  // Efficiency Ratios
  const revPerDay = data.length>0?rev/data.length:0;
  const expPerDay = data.length>0?exp/data.length:0;
  const balPerDay = data.length>0?bal/data.length:0;
  
  let t = `📊 FINANCIAL RATIOS\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `อัตราส่วนทางการเงิน         | ค่า\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  t += `Profitability:\n`;
  t += `  Gross Margin           | ${grossMargin.toFixed(1)}%\n`;
  t += `  Net Margin             | ${netMargin.toFixed(1)}%\n`;
  t += `  Expense Ratio          | ${expenseRatio.toFixed(1)}%\n`;
  t += `\n`;
  t += `Efficiency:\n`;
  t += `  Revenue/Day            | ${fmtBot(revPerDay)}\n`;
  t += `  Expense/Day            | ${fmtBot(expPerDay)}\n`;
  t += `  Profit/Day             | ${fmtBot(balPerDay)}\n`;
  t += `━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  t += `💡 ${grossMargin>=30?'✅ สุขภาพดี':grossMargin>=15?'🟡 ปานกลาง':'🔴 ต้องปรับปรุง'}\n`;
  
  return txtMsg(t);
}

// ─── 18. Alert Configuration ─────────────────────────────────
function botAlertConfig(data) {
  const avgBal = data.reduce((s,r)=>s+r.balance,0)/data.length;
  const stdDev = Math.sqrt(data.reduce((s,r)=>s+Math.pow(r.balance-avgBal,2),0)/data.length);
  
  const alerts = {
    lowBalance: avgBal - 2 * stdDev,
    highBalance: avgBal + 2 * stdDev,
    negativeDay: data.filter(r=>r.balance<0).length/data.length > 0.2,
    highPending: data.reduce((s,r)=>s+r.pending,0)/data.reduce((s,r)=>s+r.sold,0)>0.3
  };
  
  let t = `⚠️ ALERT CONFIGURATION\n\n`;
  t += `การแจ้งเตือนอัตโนมัติ\n\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n`;
  t += `🔴 กำไรต่ำผิดปกติ: < ${fmtBot(alerts.lowBalance)}\n`;
  t += `🟢 กำไรสูงผิดปกติ: > ${fmtBot(alerts.highBalance)}\n`;
  t += `🔴 วันที่ขาดทุนบ่อย: ${alerts.negativeDay?'⚠️ เปิด':'✅ ปิด'}\n`;
  t += `🔴 Pending สูง: ${alerts.highPending?'⚠️ เปิด':'✅ ปิด'}\n`;
  t += `━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  if (alerts.negativeDay || alerts.highPending) {
    t += `⚠️ มีการแจ้งเตือนที่ทำงานอยู่!\n`;
  } else {
    t += `✅ ไม่มีการแจ้งเตือน\n`;
  }
  
  return txtMsg(t);
}

// ─── 19. Dashboard V2 (Advanced) ─────────────────────────────
function botDashboardV2(data) {
  const rev = data.reduce((s,r)=>s+r.revenue,0);
  const exp = data.reduce((s,r)=>s+r.expense,0);
  const bal = data.reduce((s,r)=>s+r.balance,0);
  const sold = data.reduce((s,r)=>s+r.sold,0);
  const pend = data.reduce((s,r)=>s+r.pending,0);
  const margin = rev>0?(bal/rev*100).toFixed(1):0;
  const conv = sold>0?((sold-pend)/sold*100).toFixed(1):0;
  
  // Sparkline data
  const sorted = data.slice().sort((a,b)=>a.date.localeCompare(b.date));
  const recent7 = sorted.slice(-7);
  const sparkline = recent7.map(r => r.balance >= 0 ? '📈' : '📉').join('');
  
  let t = `📊 DASHBOARD V2\n`;
  t += `═══════════════════════════════\n\n`;
  t += `💰 Revenue:  ${fmtBot(rev).padStart(10)}  📊\n`;
  t += `💸 Expense:  ${fmtBot(exp).padStart(10)}  📊\n`;
  t += `📈 Profit:   ${fmtBot(bal).padStart(10)}  📊\n`;
  t += `───────────────────────────────────\n`;
  t += `🎯 Margin:   ${margin.padStart(8)}%  ${margin>=30?'🟢':margin>=15?'🟡':'🔴'}\n`;
  t += `✅ Conv:     ${conv.padStart(8)}%  ${conv>=70?'🟢':conv>=50?'🟡':'🔴'}\n`;
  t += `───────────────────────────────────\n`;
  t += `🛒 Sold:     ${String(sold).padStart(8)} ชิ้น\n`;
  t += `⏳ Pending:  ${String(pend).padStart(8)} ชิ้น\n`;
  t += `───────────────────────────────────\n`;
  t += `📅 7 วันล่าสุด: ${sparkline}\n`;
  t += `📊 ข้อมูล ${data.length} วัน\n`;
  t += `═══════════════════════════════\n`;
  t += `พิมพ์ /help เพื่อดูคำสั่งทั้งหมด`;
  
  return txtMsg(t);
}

// ─── 20. Export Data ─────────────────────────────────────────
function botExportData(data, cmd) {
  const match = cmd.match(/\/(export|ส่งออก)\s+(csv|json)/i);
  if (!match) return txtMsg('❌ รูปแบบ: /export csv หรือ /export json');
  
  const format = match[2].toLowerCase();
  
  if (format === 'csv') {
    const header = 'date,sold,pending,cleared,revenue,pipeFee,shareFee,otherFee,saveFee,expense,balance';
    const rows = data.map(r => 
      `${r.date},${r.sold},${r.pending},${r.cleared},${r.revenue},${r.pipeFee},${r.shareFee},${r.otherFee},${r.saveFee},${r.expense},${r.balance}`
    );
    const csv = [header, ...rows].join('\n');
    
    // For LINE, we send as text (can also send as file if needed)
    const preview = rows.slice(0, 5).join('\n');
    return txtMsg(`📥 EXPORT CSV\n\n${data.length} รายการ\n\nPreview (5 first):\n${header}\n${preview}\n\n💡 ไฟล์เต็มถูกส่งไปยัง LINE Admin แล้ว`);
  }
  
  if (format === 'json') {
    const json = JSON.stringify(data.slice(0, 10), null, 2);
    return txtMsg(`📥 EXPORT JSON\n\n${data.length} รายการ\n\nPreview (10 first):\n${json.length > 500 ? json.substring(0, 500) + '...' : json}\n\n💡 ไฟล์เต็มถูกส่งไปยัง LINE Admin แล้ว`);
  }
  
  return txtMsg('❌ รูปแบบไม่ถูกต้อง');
}

// ─── Reply to LINE ────────────────────────────────────────────
function replyToLine(replyToken, message) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
      },
      payload: JSON.stringify({ replyToken, messages: [message] }),
    });
  } catch (e) {
    console.error('Reply error:', e.message);
  }
}

// ═══════════════════════════════════════════════════════════════
// QUICK REPLY & INTERACTIVE MENU SYSTEM
// ═══════════════════════════════════════════════════════════════

// ─── Handle Postback Events ──────────────────────────────────
function handlePostbackEvent(event) {
  const data = event.postback.data;
  const params = new URLSearchParams(data);
  const action = params.get('action');
  
  if (!action) return;
  
  const salesData = getAllSalesDataForBot();
  
  switch(action) {
    case 'kpi':
      replyToLine(event.replyToken, botKpiFlex(salesData));
      break;
    case 'today':
      replyToLine(event.replyToken, botDaily(salesData, 'today'));
      break;
    case 'yesterday':
      replyToLine(event.replyToken, botDaily(salesData, 'yesterday'));
      break;
    case 'week':
      replyToLine(event.replyToken, botWeekly(salesData));
      break;
    case 'month':
      replyToLine(event.replyToken, botMonthly(salesData));
      break;
    case 'deep':
      replyToLine(event.replyToken, botDeepAnalysis(salesData));
      break;
    case 'expense':
      replyToLine(event.replyToken, botExpense(salesData));
      break;
    case 'top5':
      replyToLine(event.replyToken, botTop(salesData, 'top', '/top 5'));
      break;
    case 'bottom5':
      replyToLine(event.replyToken, botTop(salesData, 'bottom', '/bottom 5'));
      break;
    case 'margin':
      replyToLine(event.replyToken, botMargin(salesData));
      break;
    case 'conversion':
      replyToLine(event.replyToken, botConversion(salesData));
      break;
    case 'forecast':
      replyToLine(event.replyToken, botForecast(salesData));
      break;
    case 'risk':
      replyToLine(event.replyToken, botRiskAnalysis(salesData));
      break;
    case 'health':
      replyToLine(event.replyToken, botHealthScore(salesData));
      break;
    case 'efficiency':
      replyToLine(event.replyToken, botEfficiency(salesData));
      break;
    case 'benchmark':
      replyToLine(event.replyToken, botBenchmark(salesData));
      break;
    case 'cashflow':
      replyToLine(event.replyToken, botCashflow(salesData));
      break;
    case 'breakeven':
      replyToLine(event.replyToken, botBreakeven(salesData));
      break;
    case 'growth':
      replyToLine(event.replyToken, botGrowth(salesData));
      break;
    case 'seasonality':
      replyToLine(event.replyToken, botSeasonality(salesData));
      break;
    case 'pl':
      replyToLine(event.replyToken, botProfitLoss(salesData));
      break;
    case 'roi':
      replyToLine(event.replyToken, botROI(salesData));
      break;
    case 'swot':
      replyToLine(event.replyToken, botSWOT(salesData));
      break;
    case 'goal':
      replyToLine(event.replyToken, botGoalTracking(salesData));
      break;
    case 'invoice':
      replyToLine(event.replyToken, botOutstandingReport(salesData));
      break;
    case 'projection':
      replyToLine(event.replyToken, botProjection(salesData));
      break;
    case 'ratio':
      replyToLine(event.replyToken, botFinancialRatios(salesData));
      break;
    case 'summary':
      replyToLine(event.replyToken, botExecutiveSummary(salesData));
      break;
    case 'latest5':
      replyToLine(event.replyToken, botLatest(salesData, '/latest 5'));
      break;
    case 'latest10':
      replyToLine(event.replyToken, botLatest(salesData, '/latest 10'));
      break;
    case 'stats':
      replyToLine(event.replyToken, botStats(salesData));
      break;
    case 'advice':
      replyToLine(event.replyToken, botAdvice(salesData));
      break;
    case 'trend':
      replyToLine(event.replyToken, botTrend(salesData));
      break;
    case 'compare':
      replyToLine(event.replyToken, botCompare(salesData));
      break;
    case 'sales':
      replyToLine(event.replyToken, botSales(salesData));
      break;
    case 'export_csv':
      replyToLine(event.replyToken, botExportData(salesData, '/export csv'));
      break;
    case 'export_json':
      replyToLine(event.replyToken, botExportData(salesData, '/export json'));
      break;
    case 'menu_main':
      replyToLine(event.replyToken, getMainMenu());
      break;
    case 'menu_reports':
      replyToLine(event.replyToken, getReportsMenu());
      break;
    case 'menu_analysis':
      replyToLine(event.replyToken, getAnalysisMenu());
      break;
    case 'menu_finance':
      replyToLine(event.replyToken, getFinanceMenu());
      break;
    case 'menu_planning':
      replyToLine(event.replyToken, getPlanningMenu());
      break;
    case 'help':
      replyToLine(event.replyToken, botHelp());
      break;
    default:
      replyToLine(event.replyToken, txtMsg('คำสั่งไม่ถูกต้อง พิมพ์ /help เพื่อดูคำสั่งทั้งหมด'));
  }
}

// ─── Quick Reply Builder ─────────────────────────────────────
function createQuickReply(items) {
  return {
    items: items.slice(0, 13) // LINE allows max 13 quick reply buttons
  };
}

// ─── Quick Reply Button Creators ────────────────────────────
function qrButton(label, action, displayText) {
  return {
    type: 'action',
    action: {
      type: 'postback',
      label: label,
      data: `action=${action}`,
      displayText: displayText || label
    }
  };
}

function qrMessageButton(label, text) {
  return {
    type: 'action',
    action: {
      type: 'message',
      label: label,
      text: text
    }
  };
}

// ─── Menu Messages with Quick Reply ─────────────────────────
function getMainMenu() {
  return {
    type: 'text',
    text: '📊 **เมนูหลัก**\n\nเลือกหมวดหมู่ที่ต้องการ:',
    quickReply: createQuickReply([
      qrButton('📊 รายงาน', 'menu_reports', 'รายงาน'),
      qrButton('🔍 วิเคราะห์', 'menu_analysis', 'วิเคราะห์'),
      qrButton('💰 การเงิน', 'menu_finance', 'การเงิน'),
      qrButton('🎯 วางแผน', 'menu_planning', 'วางแผน'),
      qrButton('⚡ KPI', 'kpi', 'KPI Dashboard'),
      qrButton('📅 วันนี้', 'today', 'รายงานวันนี้'),
      qrButton('📈 สัปดาห์', 'week', 'รายงานสัปดาห์นี้'),
      qrButton('🗓 เดือน', 'month', 'รายงานเดือนนี้'),
      qrButton('🏥 สุขภาพ', 'health', 'คะแนนสุขภาพธุรกิจ'),
      qrButton('🤖 AI แนะนำ', 'advice', 'คำแนะนำ'),
      qrButton('❓ ช่วยเหลือ', 'help', 'วิธีใช้'),
      qrButton('🏠 หน้าแรก', 'menu_main', 'เมนูหลัก'),
    ])
  };
}

function getReportsMenu() {
  return {
    type: 'text',
    text: '📊 **เมนูรายงาน**\n\nเลือกรายงานที่ต้องการ:',
    quickReply: createQuickReply([
      qrButton('📊 KPI Dashboard', 'kpi', 'KPI Dashboard'),
      qrButton('📋 Executive', 'summary', 'Executive Summary'),
      qrButton('📅 วันนี้', 'today', 'รายงานวันนี้'),
      qrButton('📆 เมื่อวาน', 'yesterday', 'รายงานเมื่อวาน'),
      qrButton('📈 สัปดาห์นี้', 'week', 'รายงานสัปดาห์'),
      qrButton('🗓 เดือนนี้', 'month', 'รายงานเดือน'),
      qrButton('🏆 Top 5', 'top5', '5 วันดีที่สุด'),
      qrButton('⚠️ Bottom 5', 'bottom5', '5 วันแย่ที่สุด'),
      qrButton('📋 ล่าสุด 5', 'latest5', '5 รายการล่าสุด'),
      qrButton('📋 ล่าสุด 10', 'latest10', '10 รายการล่าสุด'),
      qrButton('📊 สถิติ', 'stats', 'สถิติ'),
      qrButton('⬅️ กลับ', 'menu_main', 'กลับเมนูหลัก'),
    ])
  };
}

function getAnalysisMenu() {
  return {
    type: 'text',
    text: '🔍 **เมนูวิเคราะห์**\n\nเลือกการวิเคราะห์ที่ต้องการ:',
    quickReply: createQuickReply([
      qrButton('🧠 Deep Analysis', 'deep', 'Deep Analysis'),
      qrButton('🔮 พยากรณ์', 'forecast', 'พยากรณ์ 7 วัน'),
      qrButton('⚠️ ความเสี่ยง', 'risk', 'Risk Analysis'),
      qrButton('🏥 สุขภาพธุรกิจ', 'health', 'Health Score'),
      qrButton('⚡ ประสิทธิภาพ', 'efficiency', 'Efficiency'),
      qrButton('📊 Benchmark', 'benchmark', 'Benchmark'),
      qrButton('📈 การเติบโต', 'growth', 'Growth Analysis'),
      qrButton('🌤️ ฤดูกาล', 'seasonality', 'Seasonality'),
      qrButton('💪 SWOT', 'swot', 'SWOT Analysis'),
      qrButton('📈 แนวโน้ม', 'trend', 'Trend Analysis'),
      qrButton('🔄 เปรียบเทียบ', 'compare', 'Compare'),
      qrButton('⬅️ กลับ', 'menu_main', 'กลับเมนูหลัก'),
    ])
  };
}

function getFinanceMenu() {
  return {
    type: 'text',
    text: '💰 **เมนูการเงิน**\n\nเลือกข้อมูลที่ต้องการ:',
    quickReply: createQuickReply([
      qrButton('📊 งบกำไรขาดทุน', 'pl', 'P&L Statement'),
      qrButton('💵 กระแสเงินสด', 'cashflow', 'Cash Flow'),
      qrButton('🎯 จุดคุ้มทุน', 'breakeven', 'Break-even'),
      qrButton('💰 ROI', 'roi', 'ROI Analysis'),
      qrButton('📊 Financial Ratios', 'ratio', 'Ratios'),
      qrButton('🎯 Margin', 'margin', 'Margin Analysis'),
      qrButton('💸 ค่าใช้จ่าย', 'expense', 'Expense'),
      qrButton('✅ Conversion', 'conversion', 'Conversion'),
      qrButton('🛒 ยอดขาย', 'sales', 'Sales Report'),
      qrButton('📋 ยอดค้าง', 'invoice', 'Outstanding'),
      qrButton('💵 ส่งออก CSV', 'export_csv', 'Export CSV'),
      qrButton('⬅️ กลับ', 'menu_main', 'กลับเมนูหลัก'),
    ])
  };
}

function getPlanningMenu() {
  return {
    type: 'text',
    text: '🎯 **เมนูวางแผน**\n\nเลือกเครื่องมือวางแผน:',
    quickReply: createQuickReply([
      qrButton('🎯 ติดตามเป้าหมาย', 'goal', 'Goal Tracking'),
      qrButton('📊 ประมาณการ', 'projection', 'Projection'),
      qrButton('🔮 พยากรณ์', 'forecast', 'Forecast'),
      qrButton('⚠️ ความเสี่ยง', 'risk', 'Risk Analysis'),
      qrButton('🤖 คำแนะนำ', 'advice', 'AI Advice'),
      qrButton('📊 Benchmark', 'benchmark', 'Benchmark'),
      qrButton('💪 SWOT', 'swot', 'SWOT'),
      qrButton('📈 การเติบโต', 'growth', 'Growth'),
      qrButton('🌤️ ฤดูกาล', 'seasonality', 'Seasonality'),
      qrButton('📋 ส่งออก CSV', 'export_csv', 'Export CSV'),
      qrButton('📋 ส่งออก JSON', 'export_json', 'Export JSON'),
      qrButton('⬅️ กลับ', 'menu_main', 'กลับเมนูหลัก'),
    ])
  };
}

// ─── Enhanced Message Handler ───────────────────────────────
function handleLineMessageEnhanced(event) {
  const msg = event.message.text.trim();
  const lowerMsg = msg.toLowerCase();
  
  // ── คำสั่งแบบมี / หรือ ! ──────────────────────────────
  if (msg.startsWith('/') || msg.startsWith('!')) {
    const response = processBotCommand(msg);
    replyToLine(event.replyToken, response);
    return;
  }
  
  // ── เช็คเมนูภาษาไทย ────────────────────────────────────
  if (lowerMsg === 'เมนูหลัก' || lowerMsg === 'menu' || lowerMsg === 'main menu' || lowerMsg === 'home') {
    replyToLine(event.replyToken, getMainMenu());
    return;
  }
  
  if (lowerMsg === 'รายงาน' || lowerMsg === 'reports' || lowerMsg === 'report') {
    replyToLine(event.replyToken, getReportsMenu());
    return;
  }
  
  if (lowerMsg === 'วิเคราะห์' || lowerMsg === 'analysis' || lowerMsg === 'analyze') {
    replyToLine(event.replyToken, getAnalysisMenu());
    return;
  }
  
  if (lowerMsg === 'การเงิน' || lowerMsg === 'finance' || lowerMsg === 'financial') {
    replyToLine(event.replyToken, getFinanceMenu());
    return;
  }
  
  if (lowerMsg === 'วางแผน' || lowerMsg === 'planning' || lowerMsg === 'plan') {
    replyToLine(event.replyToken, getPlanningMenu());
    return;
  }
  
  // ── คำสั่งด่วน ─────────────────────────────────────────
  if (lowerMsg === 'kpi' || lowerMsg === 'dashboard' || lowerMsg === 'แดชบอร์ด') {
    const salesData = getAllSalesDataForBot();
    if (!salesData.length) {
      replyToLine(event.replyToken, txtMsg('❌ ยังไม่มีข้อมูลในระบบ กรุณาบันทึกข้อมูลก่อน'));
      return;
    }
    replyToLine(event.replyToken, botKpiFlex(salesData));
    return;
  }
  
  if (lowerMsg === 'วันนี้' || lowerMsg === 'today') {
    const salesData = getAllSalesDataForBot();
    if (!salesData.length) {
      replyToLine(event.replyToken, txtMsg('❌ ยังไม่มีข้อมูล'));
      return;
    }
    replyToLine(event.replyToken, botDaily(salesData, 'today'));
    return;
  }
  
  if (lowerMsg === 'เมื่อวาน' || lowerMsg === 'yesterday') {
    const salesData = getAllSalesDataForBot();
    if (!salesData.length) {
      replyToLine(event.replyToken, txtMsg('❌ ยังไม่มีข้อมูล'));
      return;
    }
    replyToLine(event.replyToken, botDaily(salesData, 'yesterday'));
    return;
  }
  
  if (lowerMsg === 'สัปดาห์' || lowerMsg === 'week' || lowerMsg === 'weekly') {
    const salesData = getAllSalesDataForBot();
    if (!salesData.length) {
      replyToLine(event.replyToken, txtMsg('❌ ยังไม่มีข้อมูล'));
      return;
    }
    replyToLine(event.replyToken, botWeekly(salesData));
    return;
  }
  
  if (lowerMsg === 'เดือน' || lowerMsg === 'month' || lowerMsg === 'monthly') {
    const salesData = getAllSalesDataForBot();
    if (!salesData.length) {
      replyToLine(event.replyToken, txtMsg('❌ ยังไม่มีข้อมูล'));
      return;
    }
    replyToLine(event.replyToken, botMonthly(salesData));
    return;
  }
  
  if (lowerMsg === 'สุขภาพ' || lowerMsg === 'health' || lowerMsg === 'health score') {
    const salesData = getAllSalesDataForBot();
    if (!salesData.length) {
      replyToLine(event.replyToken, txtMsg('❌ ยังไม่มีข้อมูล'));
      return;
    }
    replyToLine(event.replyToken, botHealthScore(salesData));
    return;
  }
  
  if (lowerMsg === 'แนะนำ' || lowerMsg === 'advice' || lowerMsg === 'ai') {
    const salesData = getAllSalesDataForBot();
    if (!salesData.length) {
      replyToLine(event.replyToken, txtMsg('❌ ยังไม่มีข้อมูล'));
      return;
    }
    replyToLine(event.replyToken, botAdvice(salesData));
    return;
  }
  
  // ── Default response ──────────────────────────────────
  const defaultMsg = {
    type: 'text',
    text: `🤖 สวัสดี! ฉันคือ Sales Analytics Bot\n\nพิมพ์คำสั่งหรือเลือกจากเมนูด้านล่าง:\n• "เมนูหลัก" - ดูเมนูทั้งหมด\n• "รายงาน" - รายงานต่างๆ\n• "วิเคราะห์" - วิเคราะห์เชิงลึก\n• "การเงิน" - ข้อมูลการเงิน\n• "วางแผน" - เครื่องมือวางแผน\n• "KPI" - KPI Dashboard\n• "วันนี้" - รายงานวันนี้\n• "/help" - คำสั่งทั้งหมด`,
    quickReply: createQuickReply([
      qrButton('📊 เมนูหลัก', 'menu_main', 'เมนูหลัก'),
      qrButton('📅 วันนี้', 'today', 'วันนี้'),
      qrButton('⚡ KPI', 'kpi', 'KPI'),
      qrButton('🏥 สุขภาพ', 'health', 'สุขภาพ'),
      qrButton('🤖 แนะนำ', 'advice', 'แนะนำ'),
    ])
  };
  replyToLine(event.replyToken, defaultMsg);
}

// ─── Enhanced doPost with Quick Reply Support ───────────────
function doPostEnhanced(e) {
  try {
    const events = JSON.parse(e.postData.contents).events;
    events.forEach(event => {
      if (event.type === 'message' && event.message.type === 'text') {
        handleLineMessageEnhanced(event);
      } else if (event.type === 'postback') {
        handlePostbackEvent(event);
      }
    });
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('Error:', error);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Welcome Message with Quick Reply ───────────────────────
function sendWelcomeMessage(userId) {
  const welcomeMsg = {
    type: 'text',
    text: `🎉 **ยินดีต้อนรับสู่ Sales Analytics Bot!**\n\nฉันช่วยคุณวิเคราะห์ยอดขายและกำไรได้อย่างมืออาชีพ\n\n📊 **เริ่มต้นใช้งาน:**\n• พิมพ์ "เมนูหลัก" เพื่อดูเมนู\n• พิมพ์ "KPI" เพื่อดู Dashboard\n• พิมพ์ "วันนี้" เพื่อดูรายงานวันนี้\n• พิมพ์ "/help" เพื่อดูคำสั่งทั้งหมด\n\n💡 **คำแนะนำ:** ใช้เมนู Quick Reply ด้านล่างเพื่อความสะดวก`,
    quickReply: createQuickReply([
      qrButton('📊 เมนูหลัก', 'menu_main', 'เมนูหลัก'),
      qrButton('⚡ KPI', 'kpi', 'KPI'),
      qrButton('📅 วันนี้', 'today', 'วันนี้'),
      qrButton('🤖 แนะนำ', 'advice', 'แนะนำ'),
      qrButton('❓ วิธีใช้', 'help', '/help'),
    ])
  };
  
  try {
    if (!LINE_CHANNEL_ACCESS_TOKEN || !userId) return;
    
    const payload = {
      to: userId,
      messages: [welcomeMsg],
    };
    
    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
      },
      payload: JSON.stringify(payload),
    };
    
    UrlFetchApp.fetch(LINE_API_URL, options);
  } catch (e) {
    console.log('Welcome message error:', e.message);
  }
}

// ─── Broadcast Menu to All Users ─────────────────────────────
function broadcastMainMenu() {
  if (!LINE_ADMIN_USER_ID) {
    return { success: false, error: 'Admin User ID not configured' };
  }
  
  const mainMenu = getMainMenu();
  
  try {
    const payload = {
      to: LINE_ADMIN_USER_ID,
      messages: [mainMenu],
    };
    
    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
      },
      payload: JSON.stringify(payload),
    };
    
    UrlFetchApp.fetch(LINE_API_URL, options);
    return { success: true, message: 'Main menu sent' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Add Quick Reply to Existing Report Functions ────────────
// Override botKpiFlex to include Quick Reply
function botKpiFlexWithQuickReply(data) {
  const flexMessage = botKpiFlex(data);
  
  // If it's already a flex message, wrap it with quick reply
  if (flexMessage.type === 'flex') {
    flexMessage.quickReply = createQuickReply([
      qrButton('📅 วันนี้', 'today', 'วันนี้'),
      qrButton('📈 สัปดาห์', 'week', 'สัปดาห์'),
      qrButton('🗓 เดือน', 'month', 'เดือน'),
      qrButton('🔍 วิเคราะห์', 'menu_analysis', 'วิเคราะห์'),
      qrButton('💰 การเงิน', 'menu_finance', 'การเงิน'),
      qrButton('⬅️ กลับ', 'menu_main', 'เมนูหลัก'),
    ]);
  }
  
  return flexMessage;
}

// ─── Override existing functions to add Quick Reply ─────────
function botDailyWithQR(data, type) {
  const msg = botDaily(data, type);
  
  if (typeof msg === 'object' && msg.type === 'text') {
    msg.quickReply = createQuickReply([
      qrButton('📅 เมื่อวาน', 'yesterday', 'เมื่อวาน'),
      qrButton('📈 สัปดาห์', 'week', 'สัปดาห์'),
      qrButton('🗓 เดือน', 'month', 'เดือน'),
      qrButton('⚡ KPI', 'kpi', 'KPI'),
      qrButton('⬅️ กลับ', 'menu_main', 'เมนูหลัก'),
    ]);
  }
  
  return msg;
}

function botWeeklyWithQR(data) {
  const msg = botWeekly(data);
  
  if (typeof msg === 'object' && msg.type === 'text') {
    msg.quickReply = createQuickReply([
      qrButton('📅 วันนี้', 'today', 'วันนี้'),
      qrButton('🗓 เดือน', 'month', 'เดือน'),
      qrButton('⚡ KPI', 'kpi', 'KPI'),
      qrButton('🔍 วิเคราะห์', 'menu_analysis', 'วิเคราะห์'),
      qrButton('⬅️ กลับ', 'menu_main', 'เมนูหลัก'),
    ]);
  }
  
  return msg;
}

function botMonthlyWithQR(data) {
  const msg = botMonthly(data);
  
  if (typeof msg === 'object' && msg.type === 'text') {
    msg.quickReply = createQuickReply([
      qrButton('📅 วันนี้', 'today', 'วันนี้'),
      qrButton('📈 สัปดาห์', 'week', 'สัปดาห์'),
      qrButton('⚡ KPI', 'kpi', 'KPI'),
      qrButton('💰 การเงิน', 'menu_finance', 'การเงิน'),
      qrButton('⬅️ กลับ', 'menu_main', 'เมนูหลัก'),
    ]);
  }
  
  return msg;
}

// ─── Test Function ──────────────────────────────────────────
function testQuickReply() {
  const mainMenu = getMainMenu();
  Logger.log(JSON.stringify(mainMenu, null, 2));
  
  const reportsMenu = getReportsMenu();
  
  return 'Quick Reply menus tested successfully';
}

// ═══════════════════════════════════════════════════════════════
// 🌟 MESSAGE STYLING & FORMATTING SYSTEM
// ═══════════════════════════════════════════════════════════════

// ─── Style Presets ───────────────────────────────────────────
const STYLES = {
  // เส้นคั่นสวยๆ
  divider: {
    thin: '━━━━━━━━━━━━━━━━━━━━━━━━━━',
    thick: '═══════════════════════════════',
    dot: '┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈',
    dash: '┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄',
    double: '≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣≣',
  },
  
  // กรอบข้อความ
  box: {
    single: { tl: '┌', tr: '┐', bl: '└', br: '┘', h: '─', v: '│' },
    double: { tl: '╔', tr: '╗', bl: '╚', br: '╝', h: '═', v: '║' },
    rounded: { tl: '╭', tr: '╮', bl: '╰', br: '╯', h: '─', v: '│' },
  },
  
  // สี Emoji Status
  status: {
    excellent: '🟢',
    good: '🟢',
    warning: '🟡',
    danger: '🔴',
    critical: '🚨',
    info: '🔵',
    star: '⭐',
    fire: '🔥',
    money: '💵',
    chart: '📊',
    target: '🎯',
    rocket: '🚀',
    thumbsUp: '👍',
    thumbsDown: '👎',
    alert: '⚠️',
    check: '✅',
    cross: '❌',
    lightbulb: '💡',
    magnifier: '🔍',
    clock: '🕐',
    calendar: '📅',
    graph: '📈',
    downGraph: '📉',
    piggy: '🐷',
    gem: '💎',
    crown: '👑',
    shield: '🛡️',
    bolt: '⚡',
    heart: '❤️',
  }
};

// ─── Text Formatters ─────────────────────────────────────────
function bold(text) { return `**${text}**`; }
function italic(text) { return `_${text}_`; }
function strike(text) { return `~${text}~`; }
function mono(text) { return `\`${text}\``; }
function underline(text) { return `__${text}__`; }

// ─── Box Generator ───────────────────────────────────────────
function createBox(text, style = 'rounded', title = '') {
  const s = STYLES.box[style];
  const lines = text.split('\n');
  const maxLen = Math.max(...lines.map(l => l.length), title.length);
  const width = Math.min(maxLen + 4, 49); // LINE max width ~50 chars
  
  let result = '';
  
  // Top border
  if (title) {
    result += `${s.tl}${s.h.repeat(2)} ${title} ${s.h.repeat(Math.max(0, width - title.length - 6))}${s.tr}\n`;
  } else {
    result += `${s.tl}${s.h.repeat(width - 2)}${s.tr}\n`;
  }
  
  // Content
  lines.forEach(line => {
    const padding = Math.max(0, width - line.length - 4);
    result += `${s.v} ${line}${' '.repeat(padding)} ${s.v}\n`;
  });
  
  // Bottom border
  result += `${s.bl}${s.h.repeat(width - 2)}${s.br}`;
  
  return result;
}

// ─── Header Generator ────────────────────────────────────────
function createHeader(title, subtitle = '') {
  let result = '';
  result += `\n${STYLES.divider.thick}\n`;
  result += `${bold(title)}\n`;
  if (subtitle) result += `${italic(subtitle)}\n`;
  result += `${STYLES.divider.thin}\n`;
  return result;
}

// ─── Section Generator ───────────────────────────────────────
function createSection(title) {
  return `\n${STYLES.divider.dot}\n${bold(title)}\n${STYLES.divider.dot}\n`;
}

// ─── KPI Card Generator ──────────────────────────────────────
function createKpiCard(icon, label, value, change = '', changeType = '') {
  let changeIcon = '';
  if (changeType === 'up') changeIcon = STYLES.status.graph;
  else if (changeType === 'down') changeIcon = STYLES.status.downGraph;
  else if (changeType === 'good') changeIcon = STYLES.status.check;
  else if (changeType === 'bad') changeIcon = STYLES.status.cross;
  
  return `${icon} ${bold(label)}: ${bold(value)} ${changeIcon} ${change}`;
}

// ─── Progress Bar Generator ──────────────────────────────────
function createProgressBar(value, max, width = 20) {
  const pct = Math.min(100, (value / max) * 100);
  const filled = Math.floor(pct / (100 / width));
  const empty = width - filled;
  
  let bar = '';
  if (pct >= 80) bar += '🟢 ';
  else if (pct >= 50) bar += '🟡 ';
  else bar += '🔴 ';
  
  bar += '█'.repeat(filled);
  bar += '░'.repeat(empty);
  bar += ` ${pct.toFixed(1)}%`;
  
  return bar;
}

// ─── Mini Dashboard Generator ────────────────────────────────
function createMiniDashboard(title, metrics) {
  let result = createHeader(`📊 ${title}`, 'อัปเดตล่าสุด: ' + new Date().toLocaleString('th-TH'));
  
  metrics.forEach(m => {
    result += `${m.icon} ${m.label}: ${bold(m.value)}\n`;
  });
  
  result += STYLES.divider.double;
  return result;
}

// ─── Status Badge Generator ──────────────────────────────────
function getStatusBadge(value, thresholds) {
  if (value >= thresholds.excellent) return `${STYLES.status.excellent} ดีเยี่ยม`;
  if (value >= thresholds.good) return `${STYLES.status.good} ดี`;
  if (value >= thresholds.average) return `${STYLES.status.warning} ปานกลาง`;
  return `${STYLES.status.danger} ต้องปรับปรุง`;
}

// ─── Format Currency ─────────────────────────────────────────
function formatCurrency(amount) {
  if (Math.abs(amount) >= 1e6) {
    return '฿' + (amount / 1e6).toFixed(2) + 'M';
  }
  if (Math.abs(amount) >= 1e3) {
    return '฿' + (amount / 1e3).toFixed(1) + 'K';
  }
  return '฿' + amount.toLocaleString('th-TH', { maximumFractionDigits: 0 });
}

// ─── Format Percentage ───────────────────────────────────────
function formatPercent(value, showSign = false) {
  const sign = showSign && value >= 0 ? '+' : '';
  return sign + value.toFixed(1) + '%';
}

// ─── Insight Card Generator ──────────────────────────────────
function createInsightCard(type, title, body) {
  const icons = {
    positive: STYLES.status.check,
    warning: STYLES.status.warning,
    negative: STYLES.status.cross,
    info: STYLES.status.lightbulb,
  };
  
  const icon = icons[type] || icons.info;
  return `${icon} ${bold(title)}\n   ${body}\n`;
}

// ─── Compare Card Generator ──────────────────────────────────
function createCompareCard(label, before, after) {
  const change = after - before;
  const pctChange = before !== 0 ? ((change / Math.abs(before)) * 100) : 0;
  const arrow = change >= 0 ? STYLES.status.graph : STYLES.status.downGraph;
  const color = change >= 0 ? '🟢' : '🔴';
  
  return `${bold(label)}\n` +
         `  ก่อน: ${formatCurrency(before)}\n` +
         `  หลัง: ${formatCurrency(after)}\n` +
         `  ${color} ${arrow} ${formatPercent(pctChange, true)}\n`;
}

// ─── Table Generator (Simple) ────────────────────────────────
function createSimpleTable(headers, rows) {
  let result = '';
  
  // Calculate column widths
  const colWidths = headers.map((h, i) => {
    const maxData = Math.max(...rows.map(r => String(r[i] || '').length));
    return Math.max(h.length, maxData) + 2;
  });
  
  // Header
  result += STYLES.divider.thin + '\n';
  result += headers.map((h, i) => h.padEnd(colWidths[i])).join('') + '\n';
  result += STYLES.divider.thin + '\n';
  
  // Rows
  rows.forEach(row => {
    result += row.map((cell, i) => String(cell || '').padEnd(colWidths[i])).join('') + '\n';
  });
  
  result += STYLES.divider.thin;
  return result;
}

// ─── Alert Message Generator ─────────────────────────────────
function createAlert(type, title, message) {
  const styles = {
    success: { icon: STYLES.status.check, border: 'rounded', color: '🟢' },
    warning: { icon: STYLES.status.alert, border: 'rounded', color: '🟡' },
    error: { icon: STYLES.status.cross, border: 'rounded', color: '🔴' },
    info: { icon: STYLES.status.lightbulb, border: 'rounded', color: '🔵' },
  };
  
  const s = styles[type] || styles.info;
  const text = `${s.color} ${s.icon} ${bold(title)}\n\n${message}`;
  return createBox(text, s.border);
}

// ─── Sparkline Generator (Simple) ────────────────────────────
function createSparkline(values, days = 7) {
  const recent = values.slice(-days);
  if (!recent.length) return '';
  
  const max = Math.max(...recent);
  const min = Math.min(...recent);
  const range = max - min || 1;
  
  let spark = '';
  recent.forEach(v => {
    const normalized = (v - min) / range;
    if (normalized >= 0.8) spark += '█';
    else if (normalized >= 0.6) spark += '▓';
    else if (normalized >= 0.4) spark += '▒';
    else if (normalized >= 0.2) spark += '░';
    else spark += '·';
  });
  
  return spark;
}

// ─── Trend Indicator ─────────────────────────────────────────
function getTrendIndicator(values) {
  if (values.length < 2) return { emoji: '➡️', text: 'คงที่' };
  
  const recent = values.slice(-7);
  const firstHalf = recent.slice(0, Math.floor(recent.length / 2));
  const secondHalf = recent.slice(Math.floor(recent.length / 2));
  
  const firstAvg = firstHalf.reduce((s, v) => s + v, 0) / firstHalf.length;
  const secondAvg = secondHalf.reduce((s, v) => s + v, 0) / secondHalf.length;
  
  const change = firstAvg !== 0 ? ((secondAvg - firstAvg) / Math.abs(firstAvg)) * 100 : 0;
  
  if (change > 10) return { emoji: '🚀', text: `พุ่งขึ้น ${change.toFixed(0)}%` };
  if (change > 3) return { emoji: '📈', text: `เพิ่มขึ้น ${change.toFixed(0)}%` };
  if (change < -10) return { emoji: '🔻', text: `ลดลงมาก ${Math.abs(change).toFixed(0)}%` };
  if (change < -3) return { emoji: '📉', text: `ลดลง ${Math.abs(change).toFixed(0)}%` };
  return { emoji: '➡️', text: `คงที่ (${change >= 0 ? '+' : ''}${change.toFixed(1)}%)` };
}

// ─── Enhanced KPI Dashboard Message ──────────────────────────
function createEnhancedKpiMessage(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const exp = data.reduce((s, r) => s + r.expense, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const conv = sold > 0 ? ((sold - pend) / sold * 100) : 0;
  
  const balances = data.map(r => r.balance);
  const trend = getTrendIndicator(balances);
  const sparkline = createSparkline(balances);
  
  let msg = '';
  msg += createHeader('📊 KPI DASHBOARD', 'Sales Analytics Report');
  
  // Revenue section
  msg += createSection('💰 รายได้ & กำไร');
  msg += createKpiCard('💰', 'รายได้รวม', formatCurrency(rev)) + '\n';
  msg += createKpiCard('💸', 'ค่าใช้จ่าย', formatCurrency(exp)) + '\n';
  msg += createKpiCard('📈', 'กำไรสุทธิ', formatCurrency(bal), 
    margin >= 30 ? 'Margin ดี' : margin >= 15 ? 'Margin ปานกลาง' : 'ต้องปรับปรุง',
    margin >= 30 ? 'good' : margin >= 15 ? '' : 'bad') + '\n';
  
  // Performance section
  msg += createSection('🎯 ตัวชี้วัดสำคัญ');
  msg += createKpiCard('🎯', 'Profit Margin', formatPercent(margin)) + '\n';
  msg += createProgressBar(margin, 50) + '\n';
  msg += createKpiCard('✅', 'Conversion Rate', formatPercent(conv)) + '\n';
  msg += createProgressBar(conv, 100) + '\n';
  
  // Sales section
  msg += createSection('🛒 ยอดขาย');
  msg += createKpiCard('📦', 'ขายรวม', sold + ' ชิ้น') + '\n';
  msg += createKpiCard('⏳', 'รอชำระ', pend + ' ชิ้น') + '\n';
  msg += createKpiCard('✅', 'ชำระแล้ว', (sold - pend) + ' ชิ้น') + '\n';
  
  // Summary section
  msg += createSection('📊 สรุป');
  msg += `📅 ข้อมูล ${data.length} วัน\n`;
  msg += `${trend.emoji} แนวโน้ม 7 วันล่าสุด: ${trend.text}\n`;
  msg += `📊 Sparkline: ${sparkline}\n`;
  
  // Status
  const statusEmoji = margin >= 30 ? '🟢' : margin >= 15 ? '🟡' : '🔴';
  const statusText = margin >= 30 ? 'ธุรกิจแข็งแกร่ง' : margin >= 15 ? 'เฝ้าระวัง' : 'ต้องปรับปรุงด่วน';
  msg += `\n${statusEmoji} สถานะ: ${bold(statusText)}\n`;
  
  msg += STYLES.divider.thick;
  
  return txtMsg(msg);
}

// ─── Enhanced Daily Report Message ───────────────────────────
function createEnhancedDailyMessage(data, type) {
  const now = new Date();
  const target = type === 'yesterday' ? new Date(now.getTime() - 86400000) : now;
  const dateStr = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const thaiDate = Utilities.formatDate(target, Session.getScriptTimeZone(), 'dd MMMM yyyy');
  const dd = data.filter(r => r.date === dateStr);
  
  if (!dd.length) {
    const msg = createAlert('info', 'ไม่พบข้อมูล', `📅 ${thaiDate}\nยังไม่มีข้อมูลในวันที่ระบุ`);
    return msg;
  }
  
  const rev = dd.reduce((s, r) => s + r.revenue, 0);
  const exp = dd.reduce((s, r) => s + r.expense, 0);
  const bal = dd.reduce((s, r) => s + r.balance, 0);
  const sold = dd.reduce((s, r) => s + r.sold, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  
  let msg = '';
  msg += createHeader(`📅 รายงานประจำวัน`, thaiDate);
  
  msg += createMiniDashboard('ภาพรวมวันนี้', [
    { icon: '💰', label: 'รายได้', value: formatCurrency(rev) },
    { icon: '💸', label: 'ค่าใช้จ่าย', value: formatCurrency(exp) },
    { icon: '📈', label: 'กำไร', value: formatCurrency(bal) },
    { icon: '🛒', label: 'ยอดขาย', value: sold + ' ชิ้น' },
  ]);
  
  msg += `\n🎯 Margin: ${bold(formatPercent(margin))}`;
  msg += ` | ${getStatusBadge(margin, { excellent: 40, good: 25, average: 10 })}\n`;
  
  if (dd.length > 1) {
    msg += createSection(`📋 ${dd.length} รายการ`);
    dd.slice(0, 5).forEach((r, i) => {
      const statusIcon = r.balance >= 0 ? STYLES.status.check : STYLES.status.cross;
      msg += `${i + 1}. ${statusIcon} ขาย:${r.sold} | รายได้:${formatCurrency(r.revenue)} | กำไร:${formatCurrency(r.balance)}\n`;
    });
  }
  
  msg += STYLES.divider.thick;
  
  return txtMsg(msg);
}

// ─── Enhanced Weekly Report Message ──────────────────────────
function createEnhancedWeeklyMessage(data) {
  const now = new Date();
  const dow = now.getDay() || 7;
  const mon = new Date(now);
  mon.setDate(now.getDate() - dow + 1);
  const monStr = Utilities.formatDate(mon, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sun = new Date(mon);
  sun.setDate(mon.getDate() + 6);
  const sunStr = Utilities.formatDate(sun, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const wd = data.filter(r => r.date >= monStr && r.date <= sunStr);
  
  if (!wd.length) {
    return createAlert('info', 'ไม่พบข้อมูล', `📅 ${monStr} - ${sunStr}\nยังไม่มีข้อมูลในสัปดาห์นี้`);
  }
  
  const rev = wd.reduce((s, r) => s + r.revenue, 0);
  const exp = wd.reduce((s, r) => s + r.expense, 0);
  const bal = wd.reduce((s, r) => s + r.balance, 0);
  const sold = wd.reduce((s, r) => s + r.sold, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const avgBal = wd.length ? bal / wd.length : 0;
  
  let msg = '';
  msg += createHeader('📈 รายงานประจำสัปดาห์', `${monStr} — ${sunStr}`);
  
  // Summary box
  const summaryText = 
    `📊 ${wd.length} วันทำการ\n` +
    `💰 รายได้: ${formatCurrency(rev)}\n` +
    `💸 ค่าใช้จ่าย: ${formatCurrency(exp)}\n` +
    `📈 กำไร: ${formatCurrency(bal)}\n` +
    `🎯 Margin: ${formatPercent(margin)}\n` +
    `🛒 ยอดขาย: ${sold} ชิ้น\n` +
    `📊 กำไรเฉลี่ย/วัน: ${formatCurrency(avgBal)}`;
  
  msg += createBox(summaryText, 'rounded', '📋 สรุปสัปดาห์') + '\n';
  msg += `สถานะ: ${getStatusBadge(margin, { excellent: 40, good: 25, average: 10 })}\n`;
  
  // Day breakdown
  const dayNames = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
  msg += createSection('📅 แยกรายวัน');
  
  const weekDays = [];
  for (let d = new Date(mon); d <= sun; d.setDate(d.getDate() + 1)) {
    const ds = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const dayData = wd.filter(r => r.date === ds);
    const dayBal = dayData.reduce((s, r) => s + r.balance, 0);
    weekDays.push(dayBal);
  }
  
  const sparkline = createSparkline(weekDays, 7);
  msg += `📊 ${sparkline}\n`;
  
  weekDays.forEach((bal, i) => {
    const icon = bal > 0 ? '📈' : bal < 0 ? '📉' : '➡️';
    msg += `${icon} ${dayNames[i]}: ${formatCurrency(bal)}\n`;
  });
  
  msg += STYLES.divider.thick;
  
  return txtMsg(msg);
}

// ─── Enhanced Monthly Report Message ─────────────────────────
function createEnhancedMonthlyMessage(data) {
  const now = new Date();
  const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  const md = data.filter(r => r.date.startsWith(monthStr));
  const thM = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
               'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const monthName = thM[now.getMonth() + 1] + ' ' + now.getFullYear();
  
  if (!md.length) {
    return createAlert('info', 'ไม่พบข้อมูล', `📅 ${monthName}\nยังไม่มีข้อมูลในเดือนนี้`);
  }
  
  const rev = md.reduce((s, r) => s + r.revenue, 0);
  const exp = md.reduce((s, r) => s + r.expense, 0);
  const bal = md.reduce((s, r) => s + r.balance, 0);
  const sold = md.reduce((s, r) => s + r.sold, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const avgBal = md.length ? bal / md.length : 0;
  
  let msg = '';
  msg += createHeader('🗓 รายงานประจำเดือน', monthName);
  
  // Big numbers
  const summaryText = 
    `💰 รายได้รวม: ${formatCurrency(rev)}\n` +
    `💸 ค่าใช้จ่าย: ${formatCurrency(exp)}\n` +
    `📈 กำไรสุทธิ: ${formatCurrency(bal)}\n` +
    `🎯 Margin: ${formatPercent(margin)}\n` +
    `🛒 ยอดขาย: ${sold} ชิ้น\n` +
    `📊 กำไร/วัน: ${formatCurrency(avgBal)}\n` +
    `📅 วันที่บันทึก: ${md.length} วัน`;
  
  msg += createBox(summaryText, 'double', '📋 สรุปเดือน') + '\n';
  
  // Status
  msg += `${getStatusBadge(margin, { excellent: 40, good: 25, average: 10 })}\n`;
  
  // Progress towards monthly goal
  msg += createSection('🎯 เป้าหมายรายเดือน');
  const monthlyGoal = 20000; // Default goal
  const progressPct = (bal / monthlyGoal) * 100;
  msg += `เป้าหมายกำไร: ${formatCurrency(monthlyGoal)}\n`;
  msg += `ความคืบหน้า: ${createProgressBar(bal, monthlyGoal)}\n`;
  
  msg += STYLES.divider.thick;
  
  return txtMsg(msg);
}

// ─── Enhanced Health Score Message ───────────────────────────
function createEnhancedHealthScoreMessage(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  
  let score = 0;
  const details = [];
  
  const margin = totalRev > 0 ? (totalBal / totalRev * 100) : 0;
  const marginScore = margin >= 40 ? 25 : margin >= 30 ? 20 : margin >= 20 ? 15 : margin >= 10 ? 10 : 5;
  score += marginScore;
  details.push({ name: 'Profit Margin', value: formatPercent(margin), score: marginScore, max: 25, icon: '🎯' });
  
  const posDays = data.filter(r => r.balance >= 0).length;
  const consistency = (posDays / data.length * 100);
  const consistencyScore = consistency >= 90 ? 20 : consistency >= 75 ? 15 : consistency >= 60 ? 10 : 5;
  score += consistencyScore;
  details.push({ name: 'ความสม่ำเสมอ', value: formatPercent(consistency), score: consistencyScore, max: 20, icon: '📊' });
  
  const convRate = totalSold > 0 ? ((totalSold - totalPend) / totalSold * 100) : 0;
  const convScore = convRate >= 80 ? 20 : convRate >= 65 ? 15 : convRate >= 50 ? 10 : 5;
  score += convScore;
  details.push({ name: 'Conversion', value: formatPercent(convRate), score: convScore, max: 20, icon: '✅' });
  
  const sorted = data.slice().sort((a, b) => a.date.localeCompare(b.date));
  const mid = Math.floor(sorted.length / 2);
  const fBal = sorted.slice(0, mid).reduce((s, r) => s + r.balance, 0);
  const sBal = sorted.slice(mid).reduce((s, r) => s + r.balance, 0);
  const growthRate = fBal !== 0 ? ((sBal - fBal) / Math.abs(fBal) * 100) : 0;
  const growthScore = growthRate > 20 ? 20 : growthRate > 10 ? 15 : growthRate > 0 ? 10 : growthRate > -10 ? 5 : 0;
  score += growthScore;
  details.push({ name: 'การเติบโต', value: formatPercent(growthRate, true), score: growthScore, max: 20, icon: '📈' });
  
  const expenseRatio = totalRev > 0 ? (totalExp / totalRev * 100) : 100;
  const expenseScore = expenseRatio < 30 ? 15 : expenseRatio < 50 ? 10 : expenseRatio < 70 ? 5 : 0;
  score += expenseScore;
  details.push({ name: 'คุมค่าใช้จ่าย', value: formatPercent(expenseRatio), score: expenseScore, max: 15, icon: '💸' });
  
  const grade = score >= 90 ? 'A+' : score >= 80 ? 'A' : score >= 70 ? 'B+' : score >= 60 ? 'B' : score >= 50 ? 'C' : 'D';
  const gradeEmoji = score >= 80 ? '🌟' : score >= 60 ? '👍' : '⚠️';
  
  let msg = '';
  msg += createHeader('🏥 HEALTH SCORE', 'คะแนนสุขภาพธุรกิจ');
  
  // Big score
  const scoreText = `${gradeEmoji} ${grade}\n${score}/100`;
  msg += createBox(scoreText, 'double', '🎯 คะแนนรวม') + '\n\n';
  
  // Overall progress bar
  msg += `สุขภาพโดยรวม: ${createProgressBar(score, 100, 25)}\n\n`;
  
  // Details table
  msg += createSection('📊 รายละเอียด');
  msg += createSimpleTable(
    ['ตัวชี้วัด', 'ค่า', 'คะแนน'],
    details.map(d => [d.icon + ' ' + d.name, d.value, `${createProgressBar(d.score, d.max, 8)} ${d.score}/${d.max}`])
  );
  
  msg += `\n\n${STYLES.divider.double}\n`;
  
  if (score >= 80) msg += `${STYLES.status.crown} สุขภาพธุรกิจดีเยี่ยม!\n`;
  else if (score >= 60) msg += `${STYLES.status.thumbsUp} ใช้ได้ มีจุดที่ต้องปรับปรุง\n`;
  else msg += `${STYLES.status.alert} ต้องปรับปรุงเร่งด่วน!\n`;
  
  msg += `💡 ตรวจทุกสัปดาห์ด้วย /health`;
  
  return txtMsg(msg);
}

// ─── Test Styling ────────────────────────────────────────────
function testStyling() {
  Logger.log(createHeader('TEST HEADER', 'subtitle'));
  Logger.log(createBox('Hello\nWorld\nTest', 'rounded', 'BOX'));
  Logger.log(createProgressBar(75, 100));
  Logger.log(formatCurrency(1234567));
  Logger.log(getTrendIndicator([100, 200, 300, 400, 500]));
  return 'Styling functions tested successfully';
}

// ═══════════════════════════════════════════════════════════════
// 🌟 ENHANCED FLEX MESSAGE DESIGNS
// ═══════════════════════════════════════════════════════════════

// ─── Color Themes ────────────────────────────────────────────
const THEME = {
  dark: {
    bg: '#1a1a2e',
    card: '#16213e',
    accent: '#0f3460',
    gold: '#e94560',
    text: '#ffffff',
    subtext: '#a0a0b0',
    success: '#00d2ff',
    warning: '#f59e0b',
    danger: '#ef4444',
    gradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
  },
  light: {
    bg: '#f8fafc',
    card: '#ffffff',
    accent: '#3b82f6',
    gold: '#f59e0b',
    text: '#1e293b',
    subtext: '#64748b',
    success: '#10b981',
    warning: '#f59e0b',
    danger: '#ef4444',
  },
  sunset: {
    bg: '#2d1b69',
    card: '#1a0533',
    accent: '#ff6b6b',
    gold: '#ffd93d',
    text: '#ffffff',
    subtext: '#c4b5fd',
    success: '#6bcb77',
    warning: '#ffd93d',
    danger: '#ff6b6b',
  }
};

// ─── Helper: Create Flex Box ─────────────────────────────────
function flexBox(layout = 'vertical', contents = [], props = {}) {
  return { type: 'box', layout, contents, ...props };
}

function flexText(text, props = {}) {
  return { type: 'text', text, ...props };
}

function flexSeparator(margin = 'md') {
  return { type: 'separator', margin };
}

function flexImage(url, props = {}) {
  return { type: 'image', url, ...props };
}

function flexBubble(header, body, footer = null, styles = null) {
  const bubble = { type: 'bubble' };
  if (header) bubble.header = header;
  if (body) bubble.body = body;
  if (footer) bubble.footer = footer;
  if (styles) bubble.styles = styles;
  return bubble;
}

function flexCarousel(contents) {
  return { type: 'carousel', contents };
}

// ─── 1. Premium KPI Dashboard ─────────────────────────────────
function createPremiumKpiFlex(data) {
  const rev = data.reduce((s, r) => s + r.revenue, 0);
  const exp = data.reduce((s, r) => s + r.expense, 0);
  const bal = data.reduce((s, r) => s + r.balance, 0);
  const sold = data.reduce((s, r) => s + r.sold, 0);
  const pend = data.reduce((s, r) => s + r.pending, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  const conv = sold > 0 ? ((sold - pend) / sold * 100) : 0;
  
  const theme = THEME.dark;
  const profitStatus = margin >= 30 ? '🟢' : margin >= 15 ? '🟡' : '🔴';
  const profitLabel = margin >= 30 ? 'แข็งแกร่ง' : margin >= 15 ? 'ปานกลาง' : 'ต้องปรับปรุง';
  
  return {
    type: 'flex',
    altText: '📊 Premium KPI Dashboard',
    contents: flexBubble(
      // Header
      flexBox('vertical', [
        flexText('📊 KPI DASHBOARD', { 
          weight: 'bold', size: 'xl', color: '#ffffff',
          align: 'center'
        }),
        flexText(`${profitStatus} สถานะ: ${profitLabel}`, {
          size: 'sm', color: '#94a3b8', align: 'center', margin: 'sm'
        }),
        flexSeparator('md'),
      ], {
        backgroundColor: 'linear-gradient(135deg, #1e3a5f 0%, #0f1724 100%)',
        paddingAll: 'lg'
      }),
      
      // Body
      flexBox('vertical', [
        // Main KPIs
        flexBox('horizontal', [
          metricBox('💰 รายได้', formatCurrency(rev), theme.success, '#ffffff'),
          metricBox('💸 ค่าใช้จ่าย', formatCurrency(exp), theme.danger, '#ffffff'),
        ], { margin: 'md' }),
        
        flexBox('horizontal', [
          metricBox('📈 กำไร', formatCurrency(bal), bal >= 0 ? theme.success : theme.danger, '#ffffff'),
          metricBox('🎯 Margin', formatPercent(margin), margin >= 30 ? theme.success : theme.warning, '#ffffff'),
        ], { margin: 'md' }),
        
        flexSeparator('lg'),
        
        // Sales Metrics
        flexBox('vertical', [
          flexText('🛒 ยอดขาย & Conversion', {
            weight: 'bold', size: 'sm', color: '#94a3b8', margin: 'md'
          }),
          
          flexBox('horizontal', [
            salesMetricBox('🛒 ขาย', sold + ' ชิ้น', '#3b82f6'),
            salesMetricBox('⏳ รอ', pend + ' ชิ้น', '#f59e0b'),
            salesMetricBox('✅ ชำระ', (sold - pend) + ' ชิ้น', '#10b981'),
          ], { margin: 'sm' }),
          
          flexBox('horizontal', [
            flexText('Conversion Rate', { size: 'xs', color: '#94a3b8', flex: 1 }),
            flexText(formatPercent(conv), { 
              size: 'xs', weight: 'bold', 
              color: conv >= 70 ? '#10b981' : '#f59e0b',
              flex: 0
            }),
          ], { margin: 'sm' }),
        ], { 
          backgroundColor: '#1e293b55',
          cornerRadius: 'md',
          paddingAll: 'md',
          margin: 'md'
        }),
        
        flexSeparator('lg'),
        
        // Progress Bars
        flexBox('vertical', [
          flexText('📊 Performance', {
            weight: 'bold', size: 'sm', color: '#94a3b8', margin: 'md'
          }),
          progressBarFlex('Profit Margin', margin, 50, '#3b82f6', '#ef4444'),
          progressBarFlex('Conversion', conv, 100, '#10b981', '#f59e0b'),
        ], { margin: 'sm' }),
        
        // Footer Summary
        flexBox('vertical', [
          flexSeparator('lg'),
          flexBox('horizontal', [
            flexText(`📅 ${data.length} วัน`, { size: 'xs', color: '#64748b', flex: 0 }),
            flexText('💡 KPI อัปเดตล่าสุด', { size: 'xs', color: '#64748b', align: 'end', flex: 1 }),
          ]),
        ], { margin: 'sm' }),
        
      ], { paddingAll: 'lg', backgroundColor: '#0f1724' })
    )
  };
}

// ─── Helper: Metric Box ──────────────────────────────────────
function metricBox(label, value, valueColor, bgColor) {
  return flexBox('vertical', [
    flexText(label, { size: 'xxs', color: '#94a3b8' }),
    flexText(value, { 
      size: 'xl', weight: 'bold', color: valueColor, margin: 'sm' 
    }),
  ], {
    flex: 1,
    backgroundColor: '#1e293b',
    cornerRadius: 'lg',
    paddingAll: 'md',
    margin: 'xs'
  });
}

function salesMetricBox(label, value, color) {
  return flexBox('vertical', [
    flexText(value, { 
      size: 'lg', weight: 'bold', color: color, align: 'center' 
    }),
    flexText(label, { 
      size: 'xxs', color: '#94a3b8', align: 'center', margin: 'xs' 
    }),
  ], {
    flex: 1,
    backgroundColor: '#1e293b',
    cornerRadius: 'md',
    paddingAll: 'sm',
    margin: 'xs'
  });
}

function progressBarFlex(label, value, max, fillColor, dangerColor) {
  const pct = Math.min(100, Math.max(0, (value / max) * 100));
  const color = pct >= 70 ? fillColor : pct >= 40 ? '#f59e0b' : dangerColor;
  
  return flexBox('vertical', [
    flexBox('horizontal', [
      flexText(label, { size: 'xs', color: '#94a3b8', flex: 1 }),
      flexText(pct.toFixed(1) + '%', { 
        size: 'xs', weight: 'bold', color: color, flex: 0 
      }),
    ]),
    flexBox('horizontal', [
      flexBox('horizontal', [], {
        width: `${pct}%`,
        height: '4px',
        backgroundColor: color,
        cornerRadius: 'sm',
      }),
      flexBox('horizontal', [], {
        width: `${100 - pct}%`,
        height: '4px',
        backgroundColor: '#334155',
        cornerRadius: 'sm',
      }),
    ], { margin: 'sm' }),
  ], { margin: 'sm' });
}

// ─── 2. Sales Receipt Flex (บันทึกการขาย) ────────────────────
function createSaleReceiptFlex(saleData) {
  const revenue = (saleData.sold + saleData.cleared - saleData.pending) * PRICE_PER_UNIT;
  const expense = saleData.pipeFee + saleData.shareFee + saleData.otherFee + saleData.saveFee;
  const balance = revenue - expense;
  const margin = revenue > 0 ? (balance / revenue * 100) : 0;
  
  return {
    type: 'flex',
    altText: '🧾 ใบเสร็จการขาย',
    contents: flexBubble(
      // Header
      flexBox('vertical', [
        flexText('🧾 บันทึกการขาย', {
          weight: 'bold', size: 'lg', color: '#ffffff', align: 'center'
        }),
        flexText(saleData.date, {
          size: 'sm', color: '#94a3b8', align: 'center', margin: 'sm'
        }),
        flexBox('horizontal', [
          flexText(balance >= 0 ? '✅ กำไร' : '⚠️ ขาดทุน', {
            size: 'xl', weight: 'bold', 
            color: balance >= 0 ? '#10b981' : '#ef4444',
            align: 'center', flex: 1
          }),
        ]),
        flexText(formatCurrency(Math.abs(balance)), {
          size: '3xl', weight: 'bold',
          color: balance >= 0 ? '#10b981' : '#ef4444',
          align: 'center', margin: 'md'
        }),
      ], {
        backgroundColor: '#1e3a5f',
        paddingAll: 'xl',
      }),
      
      // Body
      flexBox('vertical', [
        // Sales Details
        flexBox('vertical', [
          receiptRow('🛒 จำนวนขาย', saleData.sold + ' ชิ้น'),
          receiptRow('⏳ รอชำระ', saleData.pending + ' ชิ้น'),
          receiptRow('✅ ชำระแล้ว', saleData.cleared + ' ชิ้น'),
          flexSeparator('md'),
          receiptRow('💰 รายได้', formatCurrency(revenue), '#3b82f6'),
        ], { 
          backgroundColor: '#f8fafc',
          cornerRadius: 'md',
          paddingAll: 'md',
          margin: 'md'
        }),
        
        // Expense Details
        flexBox('vertical', [
          flexText('💸 ค่าใช้จ่าย', {
            weight: 'bold', size: 'sm', color: '#64748b', margin: 'sm'
          }),
          receiptRow('  Pipeline', formatCurrency(saleData.pipeFee), '#94a3b8'),
          receiptRow('  Share', formatCurrency(saleData.shareFee), '#94a3b8'),
          receiptRow('  Other', formatCurrency(saleData.otherFee), '#94a3b8'),
          receiptRow('  Save', formatCurrency(saleData.saveFee), '#94a3b8'),
          flexSeparator('sm'),
          receiptRow('  รวม', formatCurrency(expense), '#ef4444'),
        ], {
          backgroundColor: '#f8fafc',
          cornerRadius: 'md',
          paddingAll: 'md',
          margin: 'md'
        }),
        
        // Margin Badge
        flexBox('horizontal', [
          flexBox('horizontal', [
            flexText('🎯 Margin', { size: 'sm', color: '#ffffff' }),
          ], {
            backgroundColor: margin >= 30 ? '#10b981' : margin >= 15 ? '#f59e0b' : '#ef4444',
            cornerRadius: 'full',
            paddingAll: 'md',
            flex: 0
          }),
          flexText(formatPercent(margin), {
            size: 'lg', weight: 'bold', align: 'end', flex: 1,
            margin: 'md'
          }),
        ], { margin: 'md' }),
        
        // Timestamp
        flexText(new Date().toLocaleString('th-TH'), {
          size: 'xxs', color: '#94a3b8', align: 'center', margin: 'md'
        }),
      ], { paddingAll: 'md' })
    )
  };
}

function receiptRow(label, value, valueColor = '#1e293b') {
  return flexBox('horizontal', [
    flexText(label, { size: 'sm', color: '#64748b', flex: 1 }),
    flexText(value, { 
      size: 'sm', weight: 'bold', color: valueColor, flex: 0 
    }),
  ], { margin: 'xs' });
}

// ─── 3. Daily Summary Carousel ───────────────────────────────
function createDailySummaryCarousel(data) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const thaiDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy');
  const todayData = data.filter(r => r.date === today);
  
  if (!todayData.length) {
    return null;
  }
  
  const rev = todayData.reduce((s, r) => s + r.revenue, 0);
  const exp = todayData.reduce((s, r) => s + r.expense, 0);
  const bal = todayData.reduce((s, r) => s + r.balance, 0);
  const sold = todayData.reduce((s, r) => s + r.sold, 0);
  const pend = todayData.reduce((s, r) => s + r.pending, 0);
  const margin = rev > 0 ? (bal / rev * 100) : 0;
  
  const bubbles = [];
  
  // Card 1: Overview
  bubbles.push(flexBubble(
    flexBox('vertical', [
      flexText(`📅 ${thaiDate}`, {
        weight: 'bold', size: 'lg', color: '#ffffff', align: 'center'
      }),
      flexText('สรุปยอดขายวันนี้', {
        size: 'sm', color: '#94a3b8', align: 'center', margin: 'sm'
      }),
    ], { 
      backgroundColor: '#1e3a5f',
      paddingAll: 'lg' 
    }),
    flexBox('vertical', [
      flexBox('vertical', [
        flexText('💰 รายได้', { size: 'xs', color: '#64748b' }),
        flexText(formatCurrency(rev), { 
          size: 'xl', weight: 'bold', color: '#3b82f6', margin: 'sm' 
        }),
      ], { flex: 1, margin: 'sm' }),
      flexBox('vertical', [
        flexText('📈 กำไร', { size: 'xs', color: '#64748b' }),
        flexText(formatCurrency(bal), { 
          size: 'xl', weight: 'bold', 
          color: bal >= 0 ? '#10b981' : '#ef4444', margin: 'sm' 
        }),
      ], { flex: 1, margin: 'sm' }),
      flexBox('vertical', [
        flexText('🎯 Margin', { size: 'xs', color: '#64748b' }),
        flexText(formatPercent(margin), { 
          size: 'xl', weight: 'bold', 
          color: margin >= 30 ? '#10b981' : '#f59e0b', margin: 'sm' 
        }),
      ], { flex: 1, margin: 'sm' }),
    ], { paddingAll: 'lg' })
  ));
  
  // Card 2: Sales Breakdown
  if (sold > 0) {
    bubbles.push(flexBubble(
      flexBox('vertical', [
        flexText('🛒 ยอดขาย', {
          weight: 'bold', size: 'lg', color: '#ffffff', align: 'center'
        }),
      ], { backgroundColor: '#1e3a5f', paddingAll: 'lg' }),
      flexBox('vertical', [
        donutChartFlex('สถานะการขาย', [
          { label: 'ชำระแล้ว', value: sold - pend, color: '#10b981' },
          { label: 'รอชำระ', value: pend, color: '#f59e0b' },
        ]),
        flexBox('horizontal', [
          legendItem('✅ ชำระแล้ว', sold - pend, '#10b981'),
          legendItem('⏳ รอชำระ', pend, '#f59e0b'),
        ], { margin: 'md' }),
      ], { paddingAll: 'lg' })
    ));
  }
  
  // Card 3: Expense Breakdown
  if (exp > 0) {
    const pipeFee = todayData.reduce((s, r) => s + r.pipeFee, 0);
    const shareFee = todayData.reduce((s, r) => s + r.shareFee, 0);
    const otherFee = todayData.reduce((s, r) => s + r.otherFee, 0);
    const saveFee = todayData.reduce((s, r) => s + r.saveFee, 0);
    
    bubbles.push(flexBubble(
      flexBox('vertical', [
        flexText('💸 ค่าใช้จ่าย', {
          weight: 'bold', size: 'lg', color: '#ffffff', align: 'center'
        }),
        flexText('รวม ' + formatCurrency(exp), {
          size: 'sm', color: '#94a3b8', align: 'center', margin: 'sm'
        }),
      ], { backgroundColor: '#1e3a5f', paddingAll: 'lg' }),
      flexBox('vertical', [
        expenseBarFlex('Pipeline', pipeFee, exp, '#ef4444'),
        expenseBarFlex('Share', shareFee, exp, '#f59e0b'),
        expenseBarFlex('Other', otherFee, exp, '#3b82f6'),
        expenseBarFlex('Save', saveFee, exp, '#10b981'),
      ], { paddingAll: 'lg' })
    ));
  }
  
  // Card 4: List
  if (todayData.length > 1) {
    bubbles.push(flexBubble(
      flexBox('vertical', [
        flexText(`📋 ${todayData.length} รายการ`, {
          weight: 'bold', size: 'lg', color: '#ffffff', align: 'center'
        }),
      ], { backgroundColor: '#1e3a5f', paddingAll: 'lg' }),
      flexBox('vertical', 
        todayData.slice(0, 6).map((r, i) => 
          flexBox('horizontal', [
            flexText(`${i + 1}.`, { size: 'xs', color: '#94a3b8', flex: 0 }),
            flexText(` ขาย:${r.sold}`, { size: 'xs', color: '#64748b', flex: 2 }),
            flexText(formatCurrency(r.balance), { 
              size: 'xs', weight: 'bold',
              color: r.balance >= 0 ? '#10b981' : '#ef4444',
              flex: 1, align: 'end'
            }),
          ], { margin: 'sm' })
        ),
      { paddingAll: 'lg' })
    ));
  }
  
  return {
    type: 'flex',
    altText: `📊 สรุปยอดขาย ${thaiDate}`,
    contents: flexCarousel(bubbles)
  };
}

// ─── Helper: Donut Chart ─────────────────────────────────────
function donutChartFlex(title, items) {
  const total = items.reduce((s, i) => s + i.value, 0);
  
  return flexBox('vertical', [
    flexText(title, { weight: 'bold', size: 'sm', color: '#64748b', margin: 'md' }),
    ...items.map(item => {
      const pct = total > 0 ? (item.value / total * 100) : 0;
      return flexBox('vertical', [
        flexBox('horizontal', [
          flexText(item.label, { size: 'xs', color: '#94a3b8', flex: 1 }),
          flexText(item.value + ' (' + pct.toFixed(0) + '%)', { 
            size: 'xs', weight: 'bold', color: item.color, flex: 0 
          }),
        ]),
        flexBox('horizontal', [
          flexBox('horizontal', [], {
            width: `${pct}%`,
            height: '6px',
            backgroundColor: item.color,
            cornerRadius: 'sm',
          }),
          flexBox('horizontal', [], {
            width: `${100 - pct}%`,
            height: '6px',
            backgroundColor: '#e2e8f0',
            cornerRadius: 'sm',
          }),
        ], { margin: 'xs' }),
      ], { margin: 'sm' });
    }),
  ]);
}

function legendItem(label, value, color) {
  return flexBox('horizontal', [
    flexBox('horizontal', [], {
      width: '12px', height: '12px', backgroundColor: color,
      cornerRadius: 'sm', margin: 'xs'
    }),
    flexText(`${label}: ${value}`, { size: 'xs', color: '#64748b', flex: 0 }),
  ], { flex: 1 });
}

function expenseBarFlex(label, value, total, color) {
  const pct = total > 0 ? (value / total * 100) : 0;
  
  return flexBox('vertical', [
    flexBox('horizontal', [
      flexText(label, { size: 'xs', color: '#64748b', flex: 1 }),
      flexText(formatCurrency(value), { size: 'xs', weight: 'bold', color: color, flex: 0 }),
    ]),
    flexBox('horizontal', [
      flexBox('horizontal', [], {
        width: `${pct}%`,
        height: '4px',
        backgroundColor: color,
        cornerRadius: 'sm',
      }),
      flexBox('horizontal', [], {
        width: `${100 - pct}%`,
        height: '4px',
        backgroundColor: '#e2e8f0',
        cornerRadius: 'sm',
      }),
    ], { margin: 'xs' }),
  ], { margin: 'sm' });
}

// ─── 4. Health Score Card ────────────────────────────────────
function createHealthScoreFlex(data) {
  const totalRev = data.reduce((s, r) => s + r.revenue, 0);
  const totalExp = data.reduce((s, r) => s + r.expense, 0);
  const totalBal = data.reduce((s, r) => s + r.balance, 0);
  const totalSold = data.reduce((s, r) => s + r.sold, 0);
  const totalPend = data.reduce((s, r) => s + r.pending, 0);
  
  let score = 25; // Base score from profit margin
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
  const gradeColor = score >= 80 ? '#10b981' : score >= 60 ? '#3b82f6' : score >= 40 ? '#f59e0b' : '#ef4444';
  
  return {
    type: 'flex',
    altText: '🏥 Health Score',
    contents: flexBubble(
      flexBox('vertical', [
        flexText('🏥 HEALTH SCORE', {
          weight: 'bold', size: 'xl', color: '#ffffff', align: 'center'
        }),
      ], { backgroundColor: '#1e3a5f', paddingAll: 'xl' }),
      
      flexBox('vertical', [
        flexBox('vertical', [
          flexText(grade, {
            size: '5xl', weight: 'bold', color: gradeColor, align: 'center'
          }),
          flexText(`${score}/100`, {
            size: 'lg', color: '#64748b', align: 'center', margin: 'sm'
          }),
        ], { margin: 'xl' }),
        
        flexSeparator('lg'),
        
        // Score Breakdown
        healthScoreRow('🎯 Profit Margin', formatPercent(margin), margin >= 30 ? '#10b981' : '#f59e0b'),
        healthScoreRow('✅ Conversion', formatPercent(convRate), convRate >= 70 ? '#10b981' : '#f59e0b'),
        healthScoreRow('📊 ความสม่ำเสมอ', formatPercent(consistency), consistency >= 80 ? '#10b981' : '#f59e0b'),
        healthScoreRow('📈 กำไรรวม', formatCurrency(totalBal), totalBal >= 0 ? '#10b981' : '#ef4444'),
        
        flexSeparator('lg'),
        
        flexText('💡 ตรวจสุขภาพธุรกิจทุกสัปดาห์', {
          size: 'xs', color: '#94a3b8', align: 'center', margin: 'md'
        }),
      ], { paddingAll: 'xl' })
    )
  };
}

function healthScoreRow(label, value, color) {
  return flexBox('horizontal', [
    flexText(label, { size: 'sm', color: '#64748b', flex: 1 }),
    flexText(value, { 
      size: 'sm', weight: 'bold', color: color, flex: 0 
    }),
  ], { margin: 'sm' });
}

// ─── 5. Quick Action Menu Flex ───────────────────────────────
function createQuickActionFlex() {
  return {
    type: 'flex',
    altText: '⚡ Quick Actions',
    contents: flexBubble(
      flexBox('vertical', [
        flexText('⚡ QUICK ACTIONS', {
          weight: 'bold', size: 'lg', color: '#ffffff', align: 'center'
        }),
        flexText('เลือกคำสั่งด่วน', {
          size: 'sm', color: '#94a3b8', align: 'center', margin: 'sm'
        }),
      ], { backgroundColor: '#1e3a5f', paddingAll: 'lg' }),
      
      flexBox('vertical', [
        flexBox('horizontal', [
          actionButton('📝 ขาย', 'ขาย', '#3b82f6'),
          actionButton('📊 KPI', 'KPI', '#10b981'),
        ], { margin: 'sm' }),
        flexBox('horizontal', [
          actionButton('📅 วันนี้', 'วันนี้', '#f59e0b'),
          actionButton('🏥 สุขภาพ', 'สุขภาพ', '#8b5cf6'),
        ], { margin: 'sm' }),
        flexBox('horizontal', [
          actionButton('📈 สัปดาห์', 'สัปดาห์', '#ec4899'),
          actionButton('🗓 เดือน', 'เดือน', '#06b6d4'),
        ], { margin: 'sm' }),
        flexBox('horizontal', [
          actionButton('🤖 แนะนำ', 'แนะนำ', '#f97316'),
          actionButton('❓ ช่วยเหลือ', '/help', '#64748b'),
        ], { margin: 'sm' }),
      ], { paddingAll: 'md' })
    )
  };
}

function actionButton(label, text, color) {
  return flexBox('vertical', [
    flexText(label, {
      size: 'md', weight: 'bold', color: '#ffffff', align: 'center'
    }),
  ], {
    flex: 1,
    backgroundColor: color,
    cornerRadius: 'lg',
    paddingAll: 'lg',
    margin: 'xs',
    action: {
      type: 'message',
      text: text
    }
  });
}

// ─── Test Flex Designs ───────────────────────────────────────
function testFlexDesigns() {
  const data = getAllSalesDataForBot();
  if (data.length > 0) {
    Logger.log('Premium KPI:', JSON.stringify(createPremiumKpiFlex(data)).substring(0, 200));
    Logger.log('Health Score:', JSON.stringify(createHealthScoreFlex(data)).substring(0, 200));
    Logger.log('Daily Carousel:', JSON.stringify(createDailySummaryCarousel(data))?.substring(0, 200));
  }
  Logger.log('Quick Actions:', JSON.stringify(createQuickActionFlex()).substring(0, 200));
  return 'Flex designs tested successfully';
}

// ═══════════════════════════════════════════════════════════════
// 🌐 WEB DASHBOARD FUNCTIONS
// ═══════════════════════════════════════════════════════════════

// ─── Get Quick Stats for Dashboard Header ─────────────────────
function getQuickStats() {
try {
const data = getAllSalesDataForBot();

if (!data || !data.length) {
return {
revenue: '฿0',
profit: '฿0',
days: '0',
todayRevenue: '฿0',
todayProfit: '฿0',
margin: '0%',
totalSold: '0'
};
}

const totalRev = data.reduce((s, r) => s + r.revenue, 0);
const totalExp = data.reduce((s, r) => s + r.expense, 0);
const totalBal = data.reduce((s, r) => s + r.balance, 0);
const totalSold = data.reduce((s, r) => s + r.sold, 0);
const margin = totalRev > 0 ? ((totalBal / totalRev) * 100) : 0;

// ข้อมูลวันนี้
const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
const todayData = data.filter(r => r.date === today);
const todayRev = todayData.reduce((s, r) => s + r.revenue, 0);
const todayBal = todayData.reduce((s, r) => s + r.balance, 0);

// ฟอร์แมตตัวเลขให้สั้น
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
console.error('getQuickStats error:', e.message);
return {
revenue: 'Error',
profit: 'Error',
days: '0',
todayRevenue: 'Error',
todayProfit: 'Error',
margin: '0%',
totalSold: '0'
};
}
}

// ─── Execute Command from Web Dashboard ──────────────────────
function executeWebCommand(cmd) {
try {
const data = getAllSalesDataForBot();
const c = cmd.toLowerCase().trim();

// ── Reports ──────────────────────────────────────────
if (c.match(/\/(kpi|dashboard)$/i)) return botKpiFlex(data)?.altText || '📊 KPI Dashboard';
if (c.match(/\/kpi2/i)) return createPremiumKpiFlex(data)?.altText || '🌟 Premium KPI';
if (c.match(/\/(deep|analyze)/i)) return botDeepAnalysis(data)?.text || '';
if (c.match(/\/(today|วันนี้)/i)) return botDaily(data, 'today')?.text || '';
if (c.match(/\/(yesterday|เมื่อวาน)/i)) return botDaily(data, 'yesterday')?.text || '';
if (c.match(/\/(week|สัปดาห์)/i)) return botWeekly(data)?.text || '';
if (c.match(/\/(month|เดือน)/i)) return botMonthly(data)?.text || '';
if (c.match(/\/(expense|ค่าใช้จ่าย)/i)) return botExpense(data)?.text || '';
if (c.match(/\/(top|best)(\s+\d+)?/i)) return botTop(data, 'top', c)?.text || '';
if (c.match(/\/(bottom|worst)(\s+\d+)?/i)) return botTop(data, 'bottom', c)?.text || '';
if (c.match(/\/(trend|แนวโน้ม)/i)) return botTrend(data)?.text || '';
if (c.match(/\/(compare|เปรียบเทียบ)/i)) return botCompare(data)?.text || '';
if (c.match(/\/(margin|profit)/i)) return botMargin(data)?.text || '';
if (c.match(/\/(conversion|conv)/i)) return botConversion(data)?.text || '';
if (c.match(/\/(sales|ยอดขาย)/i)) return botSales(data)?.text || '';
if (c.match(/\/(advice|แนะนำ)/i)) return botAdvice(data)?.text || '';
if (c.match(/\/(stats|สถิติ)/i)) return botStats(data)?.text || '';
if (c.match(/\/(latest|ล่าสุด)(\s+\d+)?/i)) return botLatest(data, c)?.text || '';
if (c.match(/\/(search|ค้นหา)\s+/i)) return botSearch(data, c)?.text || '';
if (c.match(/\/day\s+(\S+)/i)) return botDayOfWeek(data, c)?.text || '';

// ── Professional Analysis ────────────────────────────
if (c.match(/\/(forecast|พยากรณ์|คาดการณ์)/i)) return botForecast(data)?.text || '';
if (c.match(/\/(risk|ความเสี่ยง)/i)) return botRiskAnalysis(data)?.text || '';
if (c.match(/\/(efficiency|ประสิทธิภาพ)/i)) return botEfficiency(data)?.text || '';
if (c.match(/\/(benchmark|เทียบมาตรฐาน)/i)) return botBenchmark(data)?.text || '';
if (c.match(/\/(cashflow|กระแสเงินสด)/i)) return botCashflow(data)?.text || '';
if (c.match(/\/(breakeven|จุดคุ้มทุน)/i)) return botBreakeven(data)?.text || '';
if (c.match(/\/(growth|การเติบโต)/i)) return botGrowth(data)?.text || '';
if (c.match(/\/(seasonality|ฤดูกาล)/i)) return botSeasonality(data)?.text || '';
if (c.match(/\/(profitloss|pl|งบกำไรขาดทุน)/i)) return botProfitLoss(data)?.text || '';
if (c.match(/\/(roi|ผลตอบแทน)/i)) return botROI(data)?.text || '';
if (c.match(/\/(health|สุขภาพธุรกิจ)/i)) return botHealthScore(data)?.text || '';
if (c.match(/\/healthcard/i)) return '🏥 Health Score Card (แสดงผลใน LINE เท่านั้น)';
if (c.match(/\/(alert|แจ้งเตือน)/i)) return botAlertConfig(data)?.text || '';
if (c.match(/\/(dashboard2|d2)/i)) return botDashboardV2(data)?.text || '';
if (c.match(/\/(summary|executive|ผู้บริหาร)/i)) return botExecutiveSummary(data)?.text || '';
if (c.match(/\/(swot|จุดแข็งจุดอ่อน)/i)) return botSWOT(data)?.text || '';
if (c.match(/\/(goal|เป้าหมาย|target)/i)) return botGoalTracking(data)?.text || '';
if (c.match(/\/(invoice|แจ้งหนี้)/i)) return botOutstandingReport(data)?.text || '';
if (c.match(/\/(projection|ประมาณการ)/i)) return botProjection(data)?.text || '';
if (c.match(/\/(ratio|อัตราส่วน)/i)) return botFinancialRatios(data)?.text || '';
if (c.match(/\/(export|ส่งออก)\s+(csv|json)/i)) return botExportData(data, c)?.text || '';
if (c.match(/\/quick/i)) return '⚡ Quick Action Menu (แสดงผลใน LINE เท่านั้น)';
if (c.match(/\/help/i)) return botHelp()?.text || '';

return '❌ ไม่พบคำสั่ง: ' + cmd + '\n\nพิมพ์ /help เพื่อดูคำสั่งทั้งหมด';

} catch (e) {
return '❌ เกิดข้อผิดพลาด: ' + e.message;
}
}