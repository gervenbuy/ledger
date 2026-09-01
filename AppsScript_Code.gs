/**
 * 大軍記帳系統 - Apps Script 後端 v2
 * ============================================
 * 更新內容（相對 v1）：
 *   - 新增「目標設定」分頁：各帳本可設定月營收目標、月支出上限
 *   - 新增財務健康分數 API：綜合儲蓄率／帳本結構／成長趨勢／目標達成度
 * 原本的記帳功能（addRecord / getRecords / deleteRecord / getSummary / getCategories）完全保留，
 * Google Sheets 裡既有的記帳資料不受影響。
 *
 * 部署方式跟 v1 一樣：
 *   Apps Script 編輯器整段覆蓋貼上 → 存檔 → 部署 → 管理部署作業 → 編輯現有部署 → 版本選「新版本」→ 部署
 *   （用「管理部署作業」更新，網址不會變，前端 API_URL 不用改）
 * ============================================
 */

// ===== 設定區 =====
const SHEET_NAMES = ['個人', '迪特軍EV', '電商神助手', '課程教學'];
const HEADER = ['日期', '收支類型', '分類', '金額', '項目說明', '付款方式', '備註', '建立時間'];

const GOAL_SHEET_NAME = '目標設定';
const GOAL_HEADER = ['帳本', '年月', '營收目標', '支出上限', '建立時間'];

// 簡易登入密碼
const ACCESS_PASSWORD = 'Gerven0530';

// 健康分數權重
const HEALTH_WEIGHTS = {
  savingsRate: 0.35,   // 儲蓄率
  structure: 0.25,     // 帳本結構健康度
  growth: 0.25,        // 成長趨勢
  goalProgress: 0.15   // 目標達成度
};

// ===== 主要進入點 =====
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const params = e.parameter;
    const action = params.action;

    if (action !== 'verifyPassword') {
      if (params.password !== ACCESS_PASSWORD) {
        return jsonResponse({ success: false, error: '密碼錯誤或未驗證' });
      }
    }

    switch (action) {
      case 'verifyPassword':
        return jsonResponse({ success: params.password === ACCESS_PASSWORD });
      case 'addRecord':
        return addRecord(params);
      case 'getRecords':
        return getRecords(params);
      case 'deleteRecord':
        return deleteRecord(params);
      case 'getSummary':
        return getSummary(params);
      case 'getCategories':
        return getCategories(params);
      case 'setGoal':
        return setGoal(params);
      case 'getGoals':
        return getGoals(params);
      case 'getHealthScore':
        return getHealthScore(params);
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

function getSheet(ledgerName) {
  if (SHEET_NAMES.indexOf(ledgerName) === -1) {
    throw new Error('無效的帳本名稱: ' + ledgerName);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ledgerName);
  if (!sheet) {
    sheet = ss.insertSheet(ledgerName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
    sheet.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getGoalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GOAL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(GOAL_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(GOAL_HEADER);
    sheet.getRange(1, 1, 1, GOAL_HEADER.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatMonth(dateStr) {
  // 'yyyy-MM-dd' -> 'yyyy-MM'
  return dateStr.toString().substring(0, 7);
}

function shiftMonth(yyyyMM, deltaMonths) {
  const parts = yyyyMM.split('-');
  let y = parseInt(parts[0]);
  let m = parseInt(parts[1]) + deltaMonths;
  while (m < 1) { m += 12; y -= 1; }
  while (m > 12) { m -= 12; y += 1; }
  return y + '-' + (m < 10 ? '0' + m : m);
}

// ===== 新增一筆記帳 =====
function addRecord(params) {
  const sheet = getSheet(params.ledger);
  const now = new Date();
  const row = [
    params.date || Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd'),
    params.type,
    params.category,
    parseFloat(params.amount),
    params.item || '',
    params.paymentMethod || '',
    params.note || '',
    Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss')
  ];
  sheet.appendRow(row);
  return jsonResponse({ success: true, row: sheet.getLastRow() });
}

// ===== 刪除一筆記帳 =====
function deleteRecord(params) {
  const sheet = getSheet(params.ledger);
  const rowIndex = parseInt(params.rowIndex);
  if (rowIndex <= 1 || rowIndex > sheet.getLastRow()) {
    return jsonResponse({ success: false, error: '無效的列號' });
  }
  sheet.deleteRow(rowIndex);
  return jsonResponse({ success: true });
}

// ===== 取得記錄列表 =====
function getRecords(params) {
  const sheet = getSheet(params.ledger);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonResponse({ success: true, records: [] });

  const data = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  const monthFilter = params.month;

  const records = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    let dateStr = row[0];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, 'GMT+8', 'yyyy-MM-dd');
    }
    if (monthFilter && !dateStr.toString().startsWith(monthFilter)) continue;

    records.push({
      rowIndex: i + 2,
      date: dateStr,
      type: row[1],
      category: row[2],
      amount: row[3],
      item: row[4],
      paymentMethod: row[5],
      note: row[6],
      createdAt: row[7]
    });
  }
  records.sort((a, b) => (a.date < b.date ? 1 : -1));
  return jsonResponse({ success: true, records: records });
}

// ===== 取得指定帳本＋月份的收支合計（內部工具，不含 http response） =====
function computeMonthTotals(ledgerName, monthFilter) {
  const sheet = getSheet(ledgerName);
  const lastRow = sheet.getLastRow();
  let income = 0, expense = 0;
  const byCategory = {};

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
    data.forEach(function (row) {
      let dateStr = row[0];
      if (dateStr instanceof Date) {
        dateStr = Utilities.formatDate(dateStr, 'GMT+8', 'yyyy-MM-dd');
      }
      if (monthFilter && !dateStr.toString().startsWith(monthFilter)) return;

      const type = row[1];
      const category = row[2];
      const amount = parseFloat(row[3]) || 0;

      if (type === '收入') income += amount;
      else if (type === '支出') expense += amount;

      const key = type + '|' + category;
      byCategory[key] = (byCategory[key] || 0) + amount;
    });
  }
  return { income: income, expense: expense, balance: income - expense, byCategory: byCategory };
}

// ===== 取得摘要統計 =====
function getSummary(params) {
  const monthFilter = params.month;
  const targetLedgers = params.ledger === 'all' ? SHEET_NAMES : [params.ledger];

  const result = {};
  let grandIncome = 0;
  let grandExpense = 0;

  targetLedgers.forEach(function (ledgerName) {
    const totals = computeMonthTotals(ledgerName, monthFilter);
    result[ledgerName] = totals;
    grandIncome += totals.income;
    grandExpense += totals.expense;
  });

  return jsonResponse({
    success: true,
    summary: result,
    grandTotal: { income: grandIncome, expense: grandExpense, balance: grandIncome - grandExpense }
  });
}

// ===== 各帳本分類清單 =====
const DEFAULT_CATEGORIES = {
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
  }
};

function getCategories(params) {
  const ledger = params.ledger;
  if (!DEFAULT_CATEGORIES[ledger]) {
    return jsonResponse({ success: false, error: '無效的帳本名稱' });
  }
  return jsonResponse({ success: true, categories: DEFAULT_CATEGORIES[ledger] });
}

// ===== 目標設定：新增/更新（同帳本+同年月會覆蓋舊的） =====
function setGoal(params) {
  if (SHEET_NAMES.indexOf(params.ledger) === -1) {
    return jsonResponse({ success: false, error: '無效的帳本名稱' });
  }
  const sheet = getGoalSheet();
  const lastRow = sheet.getLastRow();
  const revenueGoal = parseFloat(params.revenueGoal) || 0;
  const expenseCap = parseFloat(params.expenseCap) || 0;
  const now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss');

  let foundRow = -1;
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, GOAL_HEADER.length).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === params.ledger && data[i][1] === params.month) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 3, 1, 2).setValues([[revenueGoal, expenseCap]]);
  } else {
    sheet.appendRow([params.ledger, params.month, revenueGoal, expenseCap, now]);
  }
  return jsonResponse({ success: true });
}

// ===== 目標設定：取得（單帳本+單月，或該月全部帳本） =====
function getGoals(params) {
  const sheet = getGoalSheet();
  const lastRow = sheet.getLastRow();
  const goals = {};
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, GOAL_HEADER.length).getValues();
    data.forEach(function (row) {
      const ledgerName = row[0];
      const month = row[1];
      if (params.month && month !== params.month) return;
      if (params.ledger && params.ledger !== 'all' && ledgerName !== params.ledger) return;
      goals[ledgerName] = { revenueGoal: row[2], expenseCap: row[3] };
    });
  }
  return jsonResponse({ success: true, goals: goals });
}

function getGoalFor(ledgerName, month) {
  const sheet = getGoalSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, GOAL_HEADER.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === ledgerName && data[i][1] === month) {
      return { revenueGoal: data[i][2], expenseCap: data[i][3] };
    }
  }
  return null;
}

// ===== 財務健康分數 =====
// params.month: 'yyyy-MM'，未提供則用當月
function getHealthScore(params) {
  const now = new Date();
  const thisMonth = params.month || Utilities.formatDate(now, 'GMT+8', 'yyyy-MM');
  const lastMonth = shiftMonth(thisMonth, -1);
  const twoMonthsAgo = shiftMonth(thisMonth, -2);
  const lastYearSameMonth = shiftMonth(thisMonth, -12);

  // 全帳本合計（本月 / 上月 / 去年同期）
  let thisIncome = 0, thisExpense = 0;
  let lastIncome = 0, lastExpense = 0;
  let yoyIncome = 0, yoyHasData = false;

  const perLedger = {};
  const structureWarnings = [];

  SHEET_NAMES.forEach(function (ledgerName) {
    const t0 = computeMonthTotals(ledgerName, thisMonth);
    const t1 = computeMonthTotals(ledgerName, lastMonth);
    const t2 = computeMonthTotals(ledgerName, twoMonthsAgo);
    const tYoY = computeMonthTotals(ledgerName, lastYearSameMonth);

    thisIncome += t0.income; thisExpense += t0.expense;
    lastIncome += t1.income; lastExpense += t1.expense;

    // 去年同期只要任一帳本有資料就算有
    if (tYoY.income > 0 || tYoY.expense > 0) {
      yoyHasData = true;
      yoyIncome += tYoY.income;
    }

    // 連續2個月支出>收入 → 結構警示
    const badThisMonth = t0.expense > t0.income && (t0.income > 0 || t0.expense > 0);
    const badLastMonth = t1.expense > t1.income && (t1.income > 0 || t1.expense > 0);
    const isWarning = badThisMonth && badLastMonth;
    if (isWarning) {
      structureWarnings.push(ledgerName);
    }

    perLedger[ledgerName] = {
      thisMonth: t0,
      lastMonth: t1,
      warning: isWarning
    };
  });

  // 1. 儲蓄率分數 (35%)
  let savingsRate = 0;
  if (thisIncome > 0) {
    savingsRate = (thisIncome - thisExpense) / thisIncome;
  } else if (thisExpense > 0) {
    savingsRate = -1; // 沒收入卻有支出，視為最差
  }
  // -30% => 0分, 0% => 50分, +30% => 100分，線性內插並夾在 0-100
  let savingsScore = 50 + (savingsRate / 0.3) * 50;
  savingsScore = Math.max(0, Math.min(100, savingsScore));

  // 2. 帳本結構健康度 (25%)
  const structureScore = 100 - (structureWarnings.length / SHEET_NAMES.length) * 100;

  // 3. 成長趨勢 (25%) — 以月對月為主
  let growthScore = 70; // 無上月資料時給中性分數
  let momGrowthRate = null;
  if (lastIncome > 0) {
    momGrowthRate = (thisIncome - lastIncome) / lastIncome;
    // -20% => 0分, 0% => 50分, +10% => 100分
    if (momGrowthRate >= 0) {
      growthScore = 50 + Math.min(momGrowthRate / 0.10, 1) * 50;
    } else {
      growthScore = 50 - Math.min(-momGrowthRate / 0.20, 1) * 50;
    }
    growthScore = Math.max(0, Math.min(100, growthScore));
  }
  let yoyGrowthRate = null;
  if (yoyHasData && yoyIncome > 0) {
    yoyGrowthRate = (thisIncome - yoyIncome) / yoyIncome;
  }

  // 4. 目標達成度 (15%)
  let goalScore = 70; // 無目標時給中性分數
  let goalDetails = [];
  let hasAnyGoal = false;
  SHEET_NAMES.forEach(function (ledgerName) {
    const goal = getGoalFor(ledgerName, thisMonth);
    if (!goal) return;
    hasAnyGoal = true;
    const actual = perLedger[ledgerName].thisMonth;
    const revenueProgress = goal.revenueGoal > 0 ? actual.income / goal.revenueGoal : null;
    const expenseRatio = goal.expenseCap > 0 ? actual.expense / goal.expenseCap : null;
    goalDetails.push({
      ledger: ledgerName,
      revenueGoal: goal.revenueGoal,
      actualRevenue: actual.income,
      revenueProgress: revenueProgress,
      expenseCap: goal.expenseCap,
      actualExpense: actual.expense,
      expenseRatio: expenseRatio
    });
  });
  if (hasAnyGoal) {
    // 每筆目標算一個小分數再平均：營收達成率(上限100分) + 支出未超標(超標扣分)
    let sum = 0, count = 0;
    goalDetails.forEach(function (g) {
      if (g.revenueProgress !== null) {
        sum += Math.min(g.revenueProgress, 1.2) / 1.2 * 100; // 達成或超標都給高分，封頂1.2倍
        count++;
      }
      if (g.expenseRatio !== null) {
        // 支出比例 <=1 表示沒超標；越接近0分數越高但不用太苛刻
        const ratioScore = g.expenseRatio <= 1 ? (100 - g.expenseRatio * 30) : Math.max(0, 70 - (g.expenseRatio - 1) * 100);
        sum += ratioScore;
        count++;
      }
    });
    if (count > 0) goalScore = Math.max(0, Math.min(100, sum / count));
  }

  const totalScore = Math.round(
    savingsScore * HEALTH_WEIGHTS.savingsRate +
    structureScore * HEALTH_WEIGHTS.structure +
    growthScore * HEALTH_WEIGHTS.growth +
    goalScore * HEALTH_WEIGHTS.goalProgress
  );

  let light = 'red';
  if (totalScore >= 75) light = 'green';
  else if (totalScore >= 50) light = 'yellow';

  // 白話建議
  const suggestions = [];
  if (structureWarnings.length > 0) {
    suggestions.push('「' + structureWarnings.join('、') + '」連續兩個月支出大於收入，建議檢視成本結構。');
  }
  if (savingsRate < 0) {
    suggestions.push('本月整體支出已超過收入，建議優先檢視非必要支出。');
  } else if (savingsRate < 0.1 && thisIncome > 0) {
    suggestions.push('本月儲蓄率偏低（' + Math.round(savingsRate * 100) + '%），建議留意支出佔比。');
  }
  if (momGrowthRate !== null && momGrowthRate < -0.1) {
    suggestions.push('本月營收較上月下降 ' + Math.round(-momGrowthRate * 100) + '%，建議留意業務量變化。');
  }
  if (hasAnyGoal) {
    goalDetails.forEach(function (g) {
      if (g.revenueProgress !== null && g.revenueProgress < 0.5) {
        suggestions.push(g.ledger + ' 本月營收目標達成率僅 ' + Math.round(g.revenueProgress * 100) + '%，離目標還有一段距離。');
      }
      if (g.expenseRatio !== null && g.expenseRatio > 1) {
        suggestions.push(g.ledger + ' 本月支出已超過設定上限 ' + Math.round((g.expenseRatio - 1) * 100) + '%。');
      }
    });
  }
  if (suggestions.length === 0) {
    suggestions.push('本月整體財務狀況穩定，維持目前節奏即可。');
  }

  return jsonResponse({
    success: true,
    month: thisMonth,
    totalScore: totalScore,
    light: light,
    subScores: {
      savingsRate: { score: Math.round(savingsScore), value: Math.round(savingsRate * 1000) / 10 },
      structure: { score: Math.round(structureScore), warnings: structureWarnings },
      growth: { score: Math.round(growthScore), mom: momGrowthRate !== null ? Math.round(momGrowthRate * 1000) / 10 : null, yoy: yoyGrowthRate !== null ? Math.round(yoyGrowthRate * 1000) / 10 : null },
      goalProgress: { score: Math.round(goalScore), hasGoal: hasAnyGoal, details: goalDetails }
    },
    totals: { income: thisIncome, expense: thisExpense, balance: thisIncome - thisExpense },
    suggestions: suggestions
  });
}
