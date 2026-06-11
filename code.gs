/* ============================================================
   mamibuy 行政交接系統 — Google Apps Script 後端
   
   部署方式：
   1. 開啟 Google Sheets → 延伸功能 → Apps Script
   2. 把這整段貼進 code.gs（取代預設內容）
   3. 點「部署」→「新增部署」
   4. 類型選「網頁應用程式」
   5. 「執行身分」選「我」
   6. 「誰可以存取」選「所有人」（或限定組織）
   7. 點「部署」→ 複製 Web App URL
   8. 把 URL 貼進 index.html 的設定頁
   ============================================================ */

// ── 主路由 ───────────────────────────────────────────────────

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const action = e.parameter.action;
    let result;

    switch (action) {
      case 'sync_all':
        result = syncAll();
        break;
      case 'list':
        result = listSheet(e.parameter.sheet);
        break;
      case 'get':
        result = getRow(e.parameter.sheet, e.parameter.id);
        break;
      case 'create':
        result = createRow(e.parameter.sheet, JSON.parse(e.postData.contents));
        break;
      case 'update':
        result = updateRow(e.parameter.sheet, JSON.parse(e.postData.contents));
        break;
      case 'delete':
        result = deleteRow(e.parameter.sheet, e.parameter.id);
        break;
      case 'upload_image':
        result = uploadImage(JSON.parse(e.postData.contents));
        break;
      case 'init_seed':
        result = initSeed(JSON.parse(e.postData.contents));
        break;
      default:
        result = { status: 'error', message: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 取得所有工作表（一次拿完）──────────────────────────────────

function syncAll() {
  const sheets = ['tasks', 'sops', 'accounts', 'contacts', 'callguide'];
  const data = {};
  sheets.forEach(name => {
    data[name] = readSheetAsJson(name);
  });
  return { status: 'ok', data: data };
}

// ── CRUD ─────────────────────────────────────────────────────

function listSheet(sheetName) {
  return { status: 'ok', data: readSheetAsJson(sheetName) };
}

function getRow(sheetName, id) {
  const rows = readSheetAsJson(sheetName);
  const row = rows.find(r => r.id === id);
  if (!row) return { status: 'error', message: 'Not found: ' + id };
  return { status: 'ok', data: row };
}

function createRow(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return { status: 'error', message: 'Sheet not found: ' + sheetName };

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  
  // 自動產生 id（如果沒帶）
  if (!data.id) {
    const prefix = { tasks:'c_', sops:'sop_', accounts:'a_', contacts:'ct_', callguide:'cg_' };
    data.id = (prefix[sheetName] || 'x_') + Date.now();
  }

  const row = headers.map(h => {
    const val = data[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'object') return JSON.stringify(val); // 陣列轉 JSON 字串
    if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
    return val;
  });

  ws.appendRow(row);
  return { status: 'ok', data: data };
}

function updateRow(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return { status: 'error', message: 'Sheet not found: ' + sheetName };

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('id');
  if (idCol === -1) return { status: 'error', message: 'No id column' };

  const allData = ws.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][idCol] === data.id) {
      targetRow = i + 1; // Sheet 是 1-based
      break;
    }
  }
  if (targetRow === -1) return { status: 'error', message: 'Not found: ' + data.id };

  const row = headers.map(h => {
    const val = data[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
    return val;
  });

  ws.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { status: 'ok', data: data };
}

function deleteRow(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return { status: 'error', message: 'Sheet not found: ' + sheetName };

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('id');
  const allData = ws.getDataRange().getValues();

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][idCol] === id) {
      ws.deleteRow(i + 1);
      return { status: 'ok', data: { id: id } };
    }
  }
  return { status: 'error', message: 'Not found: ' + id };
}

// ── 圖片上傳到 Google Drive ─────────────────────────────────

function uploadImage(data) {
  // data = { base64: "data:image/png;base64,...", filename: "photo.png", folderId: "xxx" }
  const folderId = data.folderId;
  if (!folderId) return { status: 'error', message: 'No folderId provided' };

  try {
    const folder = DriveApp.getFolderById(folderId);
    
    // 解析 base64
    const base64str = data.base64.replace(/^data:image\/\w+;base64,/, '');
    const mimeMatch = data.base64.match(/^data:(image\/\w+);base64,/);
    const mimeType = mimeMatch ? mimeMatch[1] : 'image/png';
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64str), mimeType, data.filename || 'image.png');
    const file = folder.createFile(blob);

    // 嘗試設定共用權限（若帳號有限制會靜默略過，資料夾繼承權限即可）
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {}

    // 回傳可直接當 <img src> 用的 URL
    const fileId = file.getId();
    const url = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1200';

    return { status: 'ok', data: { url: url, fileId: fileId } };
  } catch (err) {
    return { status: 'error', message: 'Upload failed: ' + err.message };
  }
}

// ── 初始化寫入 SEED_DATA ────────────────────────────────────

function initSeed(seedData) {
  // seedData = { tasks:[...], sops:[...], accounts:[...], contacts:[...], callguide:[...] }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = {};

  Object.keys(seedData).forEach(sheetName => {
    let ws = ss.getSheetByName(sheetName);
    
    // 如果工作表不存在，建立它
    if (!ws) {
      ws = ss.insertSheet(sheetName);
    } else {
      // 清空現有資料（保留標題或全清）
      ws.clear();
    }

    const items = seedData[sheetName];
    if (!items || items.length === 0) {
      results[sheetName] = 0;
      return;
    }

    // 取得所有欄位名（從第一筆推斷）
    const headers = Object.keys(items[0]);
    
    // 寫入標題列
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 寫入資料列
    const rows = items.map(item => {
      return headers.map(h => {
        const val = item[h];
        if (val === undefined || val === null) return '';
        if (typeof val === 'object') return JSON.stringify(val);
        if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
        return val;
      });
    });

    if (rows.length > 0) {
      ws.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    results[sheetName] = rows.length;
  });

  return { status: 'ok', data: results };
}

// ── 工具函式 ─────────────────────────────────────────────────

function readSheetAsJson(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(sheetName);
  if (!ws || ws.getLastRow() <= 1) return [];

  const data = ws.getDataRange().getValues();
  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, ci) => {
      let val = data[i][ci];
      
      // 嘗試解析 JSON 字串（weekdays, monthDays 等陣列欄位）
      if (typeof val === 'string' && val.startsWith('[')) {
        try { val = JSON.parse(val); } catch (e) {}
      }
      
      // 解析布林值
      if (val === 'TRUE') val = true;
      if (val === 'FALSE') val = false;
      
      obj[h] = val;
    });
    rows.push(obj);
  }

  return rows;
}
