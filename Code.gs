// =================================================
// 學生評語助手 - Google Apps Script 獨立專案
// 部署方式：執行身分「我」、存取對象「所有人」
// 此腳本與任何試算表無綁定，透過傳入的 sheetUrl 存取
// =================================================

const SHEET_NAMES = {
  Chinese:  '國語',
  Math:     '數學',
  Society:  '社會',
  Science:  '自然',
  Comments: '評語'
};

// ===== 欄位索引（0 = A 欄，依您的試算表調整）=====

// 國語 —— 期中 (C~M = index 2~12)
const CH_MID = { attitude:2, workbook:3, note:4, hw:5, dictation:6, sheet:7, quiz:8, daily:9, exam:10, avg:11, rank:12 };
// 國語 —— 期末 (N~X = index 13~23)
const CH_FIN = { attitude:13, workbook:14, note:15, hw:16, dictation:17, sheet:18, quiz:19, daily:20, exam:21, avg:22, rank:23 };

// 數學 —— 期中
const MA_MID = { attitude:2, textbookAvg:3, workbookAvg:4, hw:5, heavyAvg:6, sheetAvg:7, paperQuiz:8, daily:9, exam:10, avg:11, rank:12 };
// 數學 —— 期末
const MA_FIN = { attitude:13, textbookAvg:14, workbookAvg:15, hw:16, heavyAvg:17, sheetAvg:18, paperQuiz:19, daily:20, exam:21, avg:22, rank:23 };

// 社會 / 自然 —— 期中 (C~F = index 2~5)
const SI_MID = { daily:2, exam:3, avg:4, rank:5 };
// 社會 / 自然 —— 期末 (G~H = index 6~7，無平均/排名欄)
const SI_FIN = { daily:6, exam:7 };

// 評語工作表：A=座號 B=姓名 C=生成評語1 D=生成評語2 E=生成評語3 F=時間 G=正式評語
const CMT = { id:0, name:1, gen1:2, gen2:3, gen3:4, timestamp:5, final:6 };

// =================================================


function doGet(e) {
  const action   = e.parameter.action   || '';
  const sheetUrl = e.parameter.sheetUrl || '';

  let result;
  try {
    // createTemplate 不需要現有試算表，單獨處理
    if (action === 'createTemplate') {
      result = { ok:true, data: { url: createTemplate() } };
    } else {
      const ss = resolveSpreadsheet(sheetUrl);
      if      (action === 'getStudents') result = { ok:true, data: getStudents(ss) };
      else if (action === 'getScores')   result = { ok:true, data: getScores(ss) };
      else if (action === 'getComments') result = { ok:true, data: getComments(ss) };
      else if (action === 'saveComment') {
        saveComment(ss,
          e.parameter.id,
          e.parameter.name,
          b64Dec(e.parameter.gen1 || ''),
          b64Dec(e.parameter.gen2 || ''),
          b64Dec(e.parameter.gen3 || '')
        );
        result = { ok:true };
      }
      else if (action === 'saveFinalComment') {
        saveFinalComment(ss,
          e.parameter.id,
          b64Dec(e.parameter.comment || '')
        );
        result = { ok:true };
      }
      else if (action === 'getPromptTemplate') result = { ok:true, data: getPromptTemplate(ss) };
      else if (action === 'savePromptTemplate') {
        savePromptTemplate(ss, decodeURIComponent(e.parameter.template || ''));
        result = { ok:true };
      }
      else if (action === 'savePromptCell') {
        savePromptCell(ss,
          parseInt(e.parameter.cell) || 1,
          b64Dec(e.parameter.value || '')
        );
        result = { ok:true };
      }
      else if (action === 'getTraitData') result = { ok:true, data: getTraitData(ss) };
      else if (action === 'saveTraitDim') {
        saveTraitDimByIndex(ss,
          parseInt(e.parameter.dimIndex || '0'),
          JSON.parse(b64Dec(e.parameter.data || '[]'))
        );
        result = { ok:true };
      }
      else result = { ok:false, error:'未知的 action' };
    }
  } catch(err) {
    result = { ok:false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── base64 解碼（支援 URL-safe base64，相容舊版 decodeURIComponent）──
function b64Dec(str) {
  if (!str) return '';
  try {
    var b64 = str.replace(/-/g, '+').replace(/_/g, '/');
    var pad = b64.length % 4;
    if (pad) b64 += '==='.slice(pad);
    return Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString('UTF-8');
  } catch(e) {
    try { return decodeURIComponent(str); } catch(e2) { return str; }
  }
}

// ── 從網址解析試算表 ───────────────────────────────────────
function resolveSpreadsheet(sheetUrl) {
  if (!sheetUrl) throw new Error('請提供 Google Sheet 網址');
  const match = sheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) throw new Error('無效的 Google Sheet 網址，請確認網址格式正確');
  return SpreadsheetApp.openById(match[1]);
}

// ── 讀取學生名單（從「國語」工作表前兩欄）──────────────────
function getStudents(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.Chinese);
  if (!sheet) throw new Error('找不到「' + SHEET_NAMES.Chinese + '」工作表');
  const rows = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < rows.length; i++) {
    const id = rows[i][0], name = rows[i][1];
    if (!id && !name) continue;
    students.push({ id: id.toString(), name: name.toString() });
  }
  return students;
}

// ── 讀取所有科目成績 ──────────────────────────────────────
function getScores(ss) {
  const all = {};
  readSubject(ss, SHEET_NAMES.Chinese, all, '國語', CH_MID, CH_FIN,
    ['attitude','workbook','note','hw','dictation','sheet','quiz','daily','exam','avg','rank']);
  readSubject(ss, SHEET_NAMES.Math, all, '數學', MA_MID, MA_FIN,
    ['attitude','textbookAvg','workbookAvg','hw','heavyAvg','sheetAvg','paperQuiz','daily','exam','avg','rank']);
  readSubject(ss, SHEET_NAMES.Society, all, '社會', SI_MID, SI_FIN, ['daily','exam','avg','rank']);
  readSubject(ss, SHEET_NAMES.Science, all, '自然', SI_MID, SI_FIN, ['daily','exam','avg','rank']);
  return all;
}

function readSubject(ss, sheetName, all, subject, midCols, finCols, fields) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const id = rows[i][0];
    if (!id) continue;
    const sid = id.toString();
    if (!all[sid]) all[sid] = { midterm:{}, final:{} };
    all[sid].midterm[subject] = pick(rows[i], midCols, fields);
    all[sid].final[subject]   = pick(rows[i], finCols, fields);
  }
}

function pick(row, colMap, fields) {
  const obj = {};
  fields.forEach(f => {
    if (colMap[f] !== undefined) {
      const v = row[colMap[f]];
      obj[f] = (v !== '' && v !== null && v !== undefined) ? v.toString() : '';
    }
  });
  return obj;
}

// ── 讀取評語（回傳完整三則 + 正式評語）──────────────────────
function getComments(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.Comments);
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const comments = {};
  for (let i = 1; i < rows.length; i++) {
    const id = rows[i][CMT.id];
    if (!id) continue;
    comments[id.toString()] = {
      gen1:      (rows[i][CMT.gen1]      || '').toString(),
      gen2:      (rows[i][CMT.gen2]      || '').toString(),
      gen3:      (rows[i][CMT.gen3]      || '').toString(),
      timestamp: (rows[i][CMT.timestamp] || '').toString(),
      final:     (rows[i][CMT.final]     || '').toString()
    };
  }
  return comments;
}

// ── 儲存三則生成評語 + 時間戳記 ─────────────────────────────
function saveComment(ss, studentId, studentName, gen1, gen2, gen3) {
  let sheet = ss.getSheetByName(SHEET_NAMES.Comments);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.Comments);
    sheet.appendRow(['座號', '姓名', '生成評語1', '生成評語2', '生成評語3', '時間', '正式評語']);
  }
  const ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][CMT.id].toString() === studentId.toString()) {
      sheet.getRange(i+1, CMT.gen1+1, 1, 4).setValues([[gen1, gen2, gen3, ts]]);
      return;
    }
  }
  // 新增一列（保留正式評語欄為空）
  sheet.appendRow([studentId, studentName, gen1, gen2, gen3, ts, '']);
}

// ── 建立範本試算表（虛擬學生，乾淨結構）────────────────────
function createTemplate() {
  const ss = SpreadsheetApp.create('學生評語助手－範本（請重新命名）');

  const dummies = [[1,'范小明'],[2,'林小美'],[3,'王大華'],[4,'陳美玲'],[5,'李志偉']];

  // 國語
  const ch = ss.insertSheet('國語');
  ch.appendRow(['座號','姓名','期中態度','期中國習','期中筆記','期中作業','期中聽寫','期中考卷','期中平時測驗','期中平時成績','期中月考成績','期中平均','期中排名','期末態度','期末國習','期末筆記','期末作業','期末聽寫','期末考卷','期末平時測驗','期末平時成績','期末月考成績','期末平均','期末排名']);
  dummies.forEach(([id, name]) => ch.appendRow([id, name, ...Array(22).fill('')]));

  // 數學
  const ma = ss.insertSheet('數學');
  ma.appendRow(['座號','姓名','期中態度','期中數課平均','期中數習平均','期中作業','期中數重平均','期中考卷平均','期中紙筆測驗','期中平時成績','期中月考成績','期中平均','期中排名','期末態度','期末數課平均','期末數習平均','期末作業','期末數重平均','期末考卷平均','期末紙筆測驗','期末平時成績','期末月考成績','期末平均','期末排名']);
  dummies.forEach(([id, name]) => ma.appendRow([id, name, ...Array(22).fill('')]));

  // 社會
  const so = ss.insertSheet('社會');
  so.appendRow(['座號','姓名','期中平時成績','期中月考成績','期中平均','期中排名','期末平時成績','期末月考成績']);
  dummies.forEach(([id, name]) => so.appendRow([id, name, '', '', '', '', '', '']));

  // 自然
  const sc = ss.insertSheet('自然');
  sc.appendRow(['座號','姓名','期中平時成績','期中月考成績','期中平均','期中排名','期末平時成績','期末月考成績']);
  dummies.forEach(([id, name]) => sc.appendRow([id, name, '', '', '', '', '', '']));

  // 刪除預設工作表
  ['Sheet1','工作表1'].forEach(n => { try { ss.deleteSheet(ss.getSheetByName(n)); } catch(e) {} });

  return ss.getUrl();
}

// ── 提示詞範本（儲存在「提示詞」工作表 A1）──────────────────
function getPromptTemplate(ss) {
  const sheet = ss.getSheetByName('提示詞');
  if (!sheet) return '';
  // 讀取 B1:B4，合併為完整提示詞
  const vals = sheet.getRange('B1:B4').getValues();
  const parts = vals.map(row => (row[0] || '').toString().trim()).filter(v => v);
  return parts.join('\n\n');
}

function savePromptTemplate(ss, template) {
  let sheet = ss.getSheetByName('提示詞');
  if (!sheet) {
    sheet = ss.insertSheet('提示詞');
    sheet.setColumnWidth(2, 800);
  }
  sheet.getRange('B1').setValue(template);
}

// ── 儲存單格提示詞（B1~B4）──────────────────────────────────
function savePromptCell(ss, cellIndex, value) {
  let sheet = ss.getSheetByName('提示詞');
  if (!sheet) {
    sheet = ss.insertSheet('提示詞');
    sheet.setColumnWidth(2, 800);
  }
  if (cellIndex >= 1 && cellIndex <= 4) {
    sheet.getRange(cellIndex, 2).setValue(value);
  }
}

// ── 儲存正式評語（只更新 G 欄）────────────────────────────
function saveFinalComment(ss, studentId, finalComment) {
  let sheet = ss.getSheetByName(SHEET_NAMES.Comments);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.Comments);
    sheet.appendRow(['座號', '姓名', '生成評語1', '生成評語2', '生成評語3', '時間', '正式評語']);
  }
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][CMT.id].toString() === studentId.toString()) {
      sheet.getRange(i+1, CMT.final+1).setValue(finalComment);
      return;
    }
  }
  // 若尚無此學生列，先建立
  sheet.appendRow([studentId, '', '', '', '', '', finalComment]);
}

// ── 讀取個人特質庫（個人特質工作表，25 列）──────────────────
function getTraitData(ss) {
  const sheet = ss.getSheetByName('個人特質');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return rows.map(r => ({
    dimId:   r[0].toString(),
    dimName: r[1].toString(),
    subId:   r[2].toString(),
    subName: r[3].toString(),
    green:   r[4].toString().split(/[,、]/).map(s => s.trim()).filter(Boolean),
    yellow:  r[5].toString().split(/[,、]/).map(s => s.trim()).filter(Boolean),
    red:     r[6].toString().split(/[,、]/).map(s => s.trim()).filter(Boolean),
  }));
}

// ── 用列號直接更新一個向度的 5 列（不需比對 subId，避免 GAS 型別轉型問題）──
function saveTraitDimByIndex(ss, dimIndex, rows) {
  var sheet = ss.getSheetByName('個人特質');
  if (!sheet) {
    sheet = ss.insertSheet('個人特質');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(5, 3, 300);
  }

  var startRow = dimIndex * 5 + 2; // 1-indexed，第 0 向度從第 2 列開始
  var endRow   = startRow + 4;

  // 若目前列數不足，先補空列
  while (sheet.getLastRow() < endRow) {
    sheet.appendRow(['', '', '', '', '', '', '']);
  }

  // 確保第一列是欄標題
  if (sheet.getRange(1, 1).getValue() === '') {
    sheet.getRange(1, 1, 1, 7).setValues(
      [['向度id', '向度name', '子項目id', '子項目name', 'green', 'yellow', 'red']]
    );
  }

  var values = rows.map(function(r) {
    return [r.dimId, r.dimName, r.subId, r.subName, r.green, r.yellow, r.red];
  });
  sheet.getRange(startRow, 1, 5, 7).setValues(values);
}
