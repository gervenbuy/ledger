/**
 * 大軍工作室財務健康 APP - Apps Script 後端 v3
 * ============================================
 * 架構跟前兩版不同：不再用「每個帳本一個分頁」，
 * 改成「單一交易紀錄總表」+「業務類型」欄位，
 * 方便做應收應付、損益兩平、現金續航、固定成本等跨帳本的彙整分析。
 *
 * 分頁結構：
 *   - 交易紀錄：所有收支明細（取代原本 個人／迪特軍EV／電商神助手／課程教學 四個分頁）
 *   - 設定：現金餘額、老闆月薪目標
 *   - 固定成本：每月固定支出項目
 *   （原本四個舊分頁不會被刪除，保留當備份，執行 migrateOldLedgersToTransactions() 可以把舊資料搬過來）
 *
 * 部署方式：
 *   Apps Script 編輯器整段覆蓋貼上 → 存檔 → 部署 → 管理部署作業 → 編輯現有部署 → 版本選「新版本」→ 部署
 * ============================================
 */

// ===== 設定區 =====
const TRANSACTION_SHEET = '交易紀錄';
const TRANSACTION_HEADER = ['ID', '日期', '收支類型', '業務類型', '分類', '客戶', '說明', '金額', '帳戶', '付款方式', '付款狀態', '專案ID', '備註', '建立時間', '更新時間'];

const SETTINGS_SHEET = '設定';
const FIXED_COST_SHEET = '固定成本';

const OLD_LEDGER_NAMES = ['個人', '迪特軍EV', '電商神助手', '課程教學'];

const BUSINESS_TYPES = ['個人', '迪特軍EV', '電商神助手', '課程教學', '工作室共用'];

const CATEGORIES = {
  '個人': {
    '收入': ['薪資', '其他收入'],
    '支出': ['餐飲', '交通', '居住', '娛樂', '醫療', '其他支出']
  },
  '迪特軍EV': {
    '收入': ['車輛銷售', '維修收入', '零件銷售', '其他收入'],
    '支出': ['零件採購', '店租', '水電', '人事', '工具設備', '其他支出']
  },
  '電商神助手': {
    '收入': ['顧問服務費', 'AI行銷服務', '網站建置費', '其他收入'],
    '支出': ['廣告投放', '軟體訂閱', '外包費用', '其他支出']
  },
  '課程教學': {
    '收入': ['線上課程', '實體課程', '企業培訓', '其他收入'],
    '支出': ['場地費', '教材製作', '行銷推廣', '其他支出']
  },
  '工作室共用': {
    '收入': ['其他收入'],
    '支出': ['軟體訂閱', '辦公用品', '共同行銷', '雜項支出', '其他支出']
  }
};

const ACCOUNTS = [{ name: '現金' }, { name: '銀行帳戶' }, { name: '電子支付' }];

const ACCESS_PASSWORD = 'Gerven0530';

// ===== 主要進入點 =====
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    let action, password, month, data;
    if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      action = body.action; password = body.password; month = body.month; data = body.data;
    } else {
      const p = e.parameter || {};
      action = p.action; password = p.password; month = p.month;
      data = p.data ? JSON.parse(p.data) : null;
    }

    if (action !== 'verifyPassword') {
      if (password !== ACCESS_PASSWORD) {
        return jsonResponse({ success: false, error: '密碼錯誤或未驗證' });
      }
    }

    switch (action) {
      case 'verifyPassword':
        return jsonResponse({ success: password === ACCESS_PASSWORD });
      case 'getBootstrapData':
        return getBootstrapData(month);
      case 'addTransaction':
        return addTransaction(data);
      case 'updateTransaction':
        return updateTransaction(data);
      case 'deleteTransaction':
        return deleteTransaction(data && data.id, month);
      case 'saveSettings':
        return saveSettings(data, month);
      case 'addFixedCost':
        return addFixedCost(data, month);
      default:
        return jsonResponse({ success: false, error: '未知的操作: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ===== 工具函式 =====
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 把 Sheet 讀出來的日期值統一轉成 'yyyy-MM-dd'，不管是 Date 物件還是字串
function toYMD(val) {
  if (val === null || val === undefined || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'GMT+8', 'yyyy-MM-dd');
  }
  return String(val).substring(0, 10);
}

function shiftMonth(yyyyMM, deltaMonths) {
  const parts = yyyyMM.split('-');
  let y = parseInt(parts[0]);
  let m = parseInt(parts[1]) + deltaMonths;
  while (m < 1) { m += 12; y -= 1; }
  while (m > 12) { m -= 12; y += 1; }
  return y + '-' + (m < 10 ? '0' + m : m);
}

function genId() {
  return 'T' + new Date().getTime() + Math.floor(Math.random() * 1000);
}

function getSheetOrCreate(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(header);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getTransactionSheet() { return getSheetOrCreate(TRANSACTION_SHEET, TRANSACTION_HEADER); }
function getSettingsSheet() { return getSheetOrCreate(SETTINGS_SHEET, ['項目', '數值']); }
function getFixedCostSheet() { return getSheetOrCreate(FIXED_COST_SHEET, ['項目', '金額', '分類']); }

// ===== 讀取所有交易（內部用） =====
function readTransactions() {
  const sheet = getTransactionSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, TRANSACTION_HEADER.length).getValues();
  const list = [];
  data.forEach(function (row, i) {
    const type = row[2];
    if (type !== '收入' && type !== '支出') return; // 跳過異常列
    list.push({
      rowIndex: i + 2,
      id: row[0],
      date: toYMD(row[1]),
      type: type,
      businessType: row[3],
      category: row[4],
      customer: row[5],
      description: row[6],
      amount: parseFloat(row[7]) || 0,
      account: row[8],
      paymentMethod: row[9],
      paymentStatus: row[10] || '已收/付',
      projectId: row[11],
      note: row[12],
      createdAt: row[13],
      updatedAt: row[14]
    });
  });
  return list;
}

// ===== 新增一筆交易 =====
function addTransaction(data) {
  if (!data || !data.date || !data.amount) {
    return jsonResponse({ success: false, error: '缺少日期或金額' });
  }
  const sheet = getTransactionSheet();
  const id = genId();
  const now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([
    id, data.date, data.type, data.businessType, data.category,
    data.customer || '', data.description || '', parseFloat(data.amount) || 0,
    data.account || '', data.paymentMethod || '', data.paymentStatus || '已收/付',
    data.projectId || '', data.note || '', now, now
  ]);
  return getBootstrapData(data.date.substring(0, 7));
}

// ===== 修改一筆交易 =====
function updateTransaction(data) {
  if (!data || !data.id) return jsonResponse({ success: false, error: '缺少交易 ID' });
  const sheet = getTransactionSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonResponse({ success: false, error: '找不到資料' });

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowIndex = -1;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.id) { rowIndex = i + 2; break; }
  }
  if (rowIndex === -1) return jsonResponse({ success: false, error: '找不到這筆交易' });

  const now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(rowIndex, 2, 1, 12).setValues([[
    data.date, data.type, data.businessType, data.category,
    data.customer || '', data.description || '', parseFloat(data.amount) || 0,
    data.account || '', data.paymentMethod || '', data.paymentStatus || '已收/付',
    data.projectId || '', data.note || ''
  ]]);
  sheet.getRange(rowIndex, 15).setValue(now);
  return getBootstrapData(data.date.substring(0, 7));
}

// ===== 刪除一筆交易 =====
function deleteTransaction(id, month) {
  if (!id) return jsonResponse({ success: false, error: '缺少交易 ID' });
  const sheet = getTransactionSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) { sheet.deleteRow(i + 2); break; }
    }
  }
  return getBootstrapData(month);
}

// ===== 設定：現金餘額／老闆月薪目標 =====
function getSettingsValues() {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  const result = { CASH_BALANCE: 0, OWNER_SALARY_TARGET: 0 };
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(function (r) {
      if (r[0] === 'CASH_BALANCE') result.CASH_BALANCE = parseFloat(r[1]) || 0;
      if (r[0] === 'OWNER_SALARY_TARGET') result.OWNER_SALARY_TARGET = parseFloat(r[1]) || 0;
    });
  }
  return result;
}

function saveSettings(data, month) {
  const sheet = getSettingsSheet();
  function upsert(key, val) {
    const lastRow = sheet.getLastRow();
    let found = -1;
    if (lastRow > 1) {
      const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) { found = i + 2; break; }
      }
    }
    if (found > 0) sheet.getRange(found, 2).setValue(val);
    else sheet.appendRow([key, val]);
  }
  upsert('CASH_BALANCE', parseFloat(data.CASH_BALANCE) || 0);
  upsert('OWNER_SALARY_TARGET', parseFloat(data.OWNER_SALARY_TARGET) || 0);
  return getBootstrapData(month);
}

// ===== 固定成本 =====
function getFixedCostList() {
  const sheet = getFixedCostSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data
    .map(function (r) { return { item: r[0], amount: parseFloat(r[1]) || 0, category: r[2] || '其他' }; })
    .filter(function (x) { return x.item; });
}

function addFixedCost(data, month) {
  if (!data || !data.item || !data.amount) {
    return jsonResponse({ success: false, error: '請填項目與金額' });
  }
  const sheet = getFixedCostSheet();
  sheet.appendRow([data.item, parseFloat(data.amount) || 0, data.category || '其他']);
  return getBootstrapData(month);
}

// ===== 主要彙總：儀表板 + 帳本 + 趨勢 + 應收應付 =====
function getBootstrapData(month) {
  const targetMonth = month || Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM');
  const all = readTransactions();
  const monthTx = all.filter(function (t) { return t.date && t.date.indexOf(targetMonth) === 0; });

  let revenue = 0, expenses = 0;
  const byBusiness = {};
  BUSINESS_TYPES.forEach(function (b) { byBusiness[b] = { revenue: 0, expense: 0, profit: 0, margin: 0 }; });

  monthTx.forEach(function (t) {
    if (t.type === '收入') {
      revenue += t.amount;
      if (byBusiness[t.businessType]) byBusiness[t.businessType].revenue += t.amount;
    } else if (t.type === '支出') {
      expenses += t.amount;
      if (byBusiness[t.businessType]) byBusiness[t.businessType].expense += t.amount;
    }
  });
  Object.keys(byBusiness).forEach(function (b) {
    const x = byBusiness[b];
    x.profit = x.revenue - x.expense;
    x.margin = x.revenue > 0 ? (x.profit / x.revenue * 100) : 0;
  });

  const fixedCostList = getFixedCostList();
  const fixedCosts = fixedCostList.reduce(function (s, x) { return s + x.amount; }, 0);

  const settings = getSettingsValues();
  const cashBalance = settings.CASH_BALANCE;
  const ownerSalary = settings.OWNER_SALARY_TARGET;

  const netProfit = revenue - expenses - fixedCosts;
  const breakevenRevenue = expenses + fixedCosts;
  const monthlyBurn = (expenses + fixedCosts) - revenue;
  let runwayMonths;
  if (monthlyBurn <= 0) {
    runwayMonths = 99;
  } else {
    runwayMonths = cashBalance > 0 ? Math.min(99, cashBalance / monthlyBurn) : 0;
  }

  let health;
  if (netProfit >= 0 && runwayMonths >= 3) health = '綠燈';
  else if (runwayMonths >= 1) health = '黃燈';
  else health = '紅燈';

  // 應收應付（跨所有月份，只要還沒結清）
  const unsettled = all.filter(function (t) { return t.paymentStatus === '未收/付'; });
  const receivable = unsettled.filter(function (t) { return t.type === '收入'; }).reduce(function (s, x) { return s + x.amount; }, 0);
  const payable = unsettled.filter(function (t) { return t.type === '支出'; }).reduce(function (s, x) { return s + x.amount; }, 0);
  const outstandingItems = unsettled
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; })
    .slice(0, 30)
    .map(function (t) {
      return { id: t.id, date: t.date, businessType: t.businessType, category: t.category, customer: t.customer, description: t.description, type: t.type, amount: t.amount };
    });

  // 近 6 個月趨勢
  const trend = [];
  for (let i = 5; i >= 0; i--) {
    const m = shiftMonth(targetMonth, -i);
    const tx = all.filter(function (t) { return t.date && t.date.indexOf(m) === 0; });
    const rev = tx.filter(function (t) { return t.type === '收入'; }).reduce(function (s, x) { return s + x.amount; }, 0);
    const exp = tx.filter(function (t) { return t.type === '支出'; }).reduce(function (s, x) { return s + x.amount; }, 0);
    trend.push({ month: m, revenue: rev, expense: exp });
  }

  // 分類清單
  const categories = [];
  Object.keys(CATEGORIES).forEach(function (biz) {
    ['收入', '支出'].forEach(function (type) {
      (CATEGORIES[biz][type] || []).forEach(function (cat) {
        categories.push({ name: cat, type: type, businessType: biz });
      });
    });
  });

  return jsonResponse({
    success: true,
    month: targetMonth,
    dashboard: {
      month: targetMonth,
      revenue: revenue,
      expenses: expenses,
      netProfit: netProfit,
      fixedCosts: fixedCosts,
      cashBalance: cashBalance,
      breakevenRevenue: breakevenRevenue,
      runwayMonths: Math.round(runwayMonths * 10) / 10,
      health: health,
      ownerSalary: ownerSalary,
      byBusiness: byBusiness
    },
    transactions: monthTx.map(function (t) {
      return { id: t.id, date: t.date, type: t.type, businessType: t.businessType, category: t.category, customer: t.customer, description: t.description, amount: t.amount, account: t.account, paymentMethod: t.paymentMethod, paymentStatus: t.paymentStatus, projectId: t.projectId, note: t.note };
    }),
    trend: trend,
    outstanding: { receivable: receivable, payable: payable, count: outstandingItems.length, items: outstandingItems },
    fixedCosts: fixedCostList,
    businessTypes: BUSINESS_TYPES,
    categories: categories,
    accounts: ACCOUNTS
  });
}

// ===== 一次性搬移工具：把舊的四個分頁資料搬進新的「交易紀錄」總表 =====
// 使用方式：在 Apps Script 編輯器選這個函式，按「執行」（不用部署，也不透過網頁）
// 執行完可以到執行紀錄（左側時鐘圖示）看 Logger 訊息，或直接打開「交易紀錄」分頁確認
// 原本四個分頁不會被刪除，繼續保留當備份
function migrateOldLedgersToTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newSheet = getTransactionSheet();
  let migratedCount = 0;

  OLD_LEDGER_NAMES.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    data.forEach(function (row) {
      const type = row[1];
      if (type !== '收入' && type !== '支出') return; // 跳過殘留的重複標題列等異常資料

      const dateVal = toYMD(row[0]);
      const category = row[2];
      const amount = parseFloat(row[3]) || 0;
      const item = row[4] || '';
      const paymentMethod = row[5] || '';
      const note = row[6] || '';
      let createdAt = row[7];
      if (Object.prototype.toString.call(createdAt) === '[object Date]') {
        createdAt = Utilities.formatDate(createdAt, 'GMT+8', 'yyyy-MM-dd HH:mm:ss');
      } else if (!createdAt) {
        createdAt = dateVal;
      }

      const id = 'M' + Utilities.getUuid().substring(0, 8);
      newSheet.appendRow([
        id, dateVal, type, name, category,
        '', item, amount,
        paymentMethod || '現金', paymentMethod || '',
        '已收/付', '', note, createdAt, createdAt
      ]);
      migratedCount++;
    });
  });

  Logger.log('搬移完成，共搬移 ' + migratedCount + ' 筆資料到「' + TRANSACTION_SHEET + '」分頁。');
}
