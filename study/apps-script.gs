/**
 * 楷宸讀書系統：雲端同步用的 Google Apps Script
 * 用法：在 Google 試算表裡按「擴充功能 → Apps Script」，把這整段貼上，
 * 再「部署 → 新增部署作業 → 網頁應用程式」（執行身分：我；誰可以存取：所有人）。
 * 資料會存在試算表的「data」分頁：每一列是一筆紀錄（key、內容、更新時間）。
 */
const SHEET_NAME = 'data';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'json', 'updated']);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 讀取全部紀錄
function doGet() {
  const values = sheet_().getDataRange().getValues();
  const data = {};
  for (let i = 1; i < values.length; i++) {
    const key = values[i][0];
    if (!key) continue;
    try { data[key] = JSON.parse(values[i][1]); } catch (e) {}
  }
  return json_({ ok: true, data: data });
}

// 新增、修改或刪除紀錄
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const ops = body.ops || [];
    const sh = sheet_();
    const keys = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(function (r) { return r[0]; });
    ops.forEach(function (op) {
      const idx = keys.indexOf(op.key);
      if (op.action === 'delete') {
        if (idx > 0) { sh.deleteRow(idx + 1); keys.splice(idx, 1); }
        return;
      }
      const row = [op.key, JSON.stringify(op.value), new Date()];
      if (idx > 0) {
        sh.getRange(idx + 1, 1, 1, 3).setValues([row]);
      } else {
        sh.appendRow(row);
        keys.push(op.key);
      }
    });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
