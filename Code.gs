// ========================================
// 設定
// ========================================
// ★現行スプレッドシート (顧客リスト・マスタ・案件管理 等が入っている正本) のID
// スプレッドシートURL: https://docs.google.com/spreadsheets/d/【ここがID】/edit
const MAIN_SS_ID = '1_guv8ou75mJ1SceBymgOYS9bUGztirx_7PARcXsk8mM';

const TEMPLATE_SS_ID = '1_guv8ou75mJ1SceBymgOYS9bUGztirx_7PARcXsk8mM';
const DESTINATION_FOLDER_ID = '11oyvYyiWDIz5ogII2i-abbw1gJtKDHTJ';

// バインド済みでも独立スクリプトでもどちらでも動くようにする
function getMainSS_() {
  if (MAIN_SS_ID && MAIN_SS_ID.indexOf('ここに') === -1) {
    return SpreadsheetApp.openById(MAIN_SS_ID);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('スプレッドシートに紐付いていません。Code.gs の MAIN_SS_ID に現行スプレッドシートのIDを設定してください。');
  }
  return active;
}

// ★診断用: GASエディタの関数一覧から手動実行してください
// 各シートが存在するかチェックします
function runDiagnostic() {
  const ss = getMainSS_();
  console.log('SS名:', ss.getName());
  console.log('SS URL:', ss.getUrl());
  console.log('---');
  const required = Object.values(SHEET_NAMES);
  required.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      console.log('✅ ' + name + ' (' + sh.getLastRow() + '行)');
    } else {
      console.log('❌ ' + name + ' が見つかりません');
    }
  });
  console.log('---');
  console.log('テンプレSS:');
  try {
    const tpl = SpreadsheetApp.openById(TEMPLATE_SS_ID);
    console.log('✅ テンプレSS開けました: ' + tpl.getName());
    console.log('  シート一覧: ' + tpl.getSheets().map(s => s.getName()).join(', '));
  } catch (e) {
    console.log('❌ テンプレSS開けません: ' + e);
  }
}

const SHEET_NAMES = {
  CLIENT: '顧客リスト',
  UNIT_PRICE: 'マスタ_単価',
  PHARMA: 'マスタ_工数_製薬WS',
  WS_NORMAL: 'マスタ_工数_WS製薬以外',
  PANEL: 'マスタ_工数_パネル',
  PROJECT: '案件管理',
  DETAIL: '見積明細',
  TRAVEL: '旅費明細'
};

const PROJECT_COL = { ID:1, TITLE:2, CLIENT:3, DATE:4, SOURCE:5, STATUS:6, MEMO:7, TOTAL:8, TYPE:9 };
const DETAIL_COL  = { ID:1, PROJECT_ID:2, EVENT_TYPE:3, HOURS:4, IS_MEDICAL:5, UNIT_PRICE:6, RANK_NAME:7, SUBTOTAL:8, NOTE:9, PREP_HOURS:10 };
const TRAVEL_COL  = { ID:1, PROJECT_ID:2, TYPE:3, FROM:4, TO:5, PLACE:6, PRICE_IN:7, COUNT:8, IS_ROUND_TRIP:9 };

// ========================================
// エントリーポイント
// ========================================
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('たがやす見積シミュレーターv4')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ========================================
// 初期データ取得 (1 RPC で全部返す → 高速化)
// ========================================
function getInitialData() {
  return {
    clients: getClientList_(),
    unitPrices: getUnitPriceMaster_(),
    pharma: getMasterByName_(SHEET_NAMES.PHARMA),
    wsNormal: getMasterByName_(SHEET_NAMES.WS_NORMAL),
    panel: getMasterByName_(SHEET_NAMES.PANEL, 3),
    templateStructure: getTemplateStructure_(),
    illustStructure: getIllustStructure_()
  };
}

// イラスト専用テンプレ (内訳明細イラストのみ) から品目構造を取得
function getIllustStructure_() {
  const sh = SpreadsheetApp.openById(TEMPLATE_SS_ID).getSheetByName('内訳明細イラストのみ');
  if (!sh || sh.getLastRow() < 1) return [];
  const data = sh.getRange(1, 1, sh.getLastRow(), 4).getValues();  // A:D 列
  const result = [];
  const seenStageNums = new Set();
  let stage = '';
  for (const row of data) {
    const cell = String(row[0]).trim();
    if (!cell) continue;
    const stageMatch = cell.match(/^[\(（](\d+)/);
    if (stageMatch) {
      const num = stageMatch[1];
      if (seenStageNums.has(num)) break;
      seenStageNums.add(num);
      // 管理費・特急料金・小計 等のステージはアプリ画面では扱わない
      if (/管理費|諸経費|特急|小計|合計/.test(cell)) {
        stage = '';
        continue;
      }
      stage = cell;
      continue;
    }
    if (/管理費|諸経費|小計|合計|出精|消費税|内訳|品名/.test(cell)) continue;
    if (!stage) continue;  // ステージ番号前のカテゴリヘッダーはスキップ
    const price = Number(row[3]) || 0;
    const unit = String(row[2] || '').trim();
    // 単位なし行 = カテゴリヘッダー / 区切り行 とみなしてスキップ
    if (!unit) continue;
    result.push({ stage, role: cell, price, unit });
  }
  return result;
}

function getClientList_() {
  const sh = getMainSS_().getSheetByName(SHEET_NAMES.CLIENT);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().filter(String);
}

function getUnitPriceMaster_() {
  const sh = getMainSS_().getSheetByName(SHEET_NAMES.UNIT_PRICE);
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const prices = {};
  data.forEach(r => { if (r[0]) prices[r[0]] = Number(r[2]); });
  return prices;
}

function getMasterByName_(sheetName, fixedCols) {
  const sh = getMainSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const cols = fixedCols || sh.getLastColumn();
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
  return fillDownStage_(data);
}

function fillDownStage_(data) {
  let last = '';
  return data.map(row => {
    if (row[0] && String(row[0]).trim() !== '') last = row[0];
    else row[0] = last;
    return row;
  });
}

function getTemplateStructure_() {
  const sh = SpreadsheetApp.openById(TEMPLATE_SS_ID).getSheetByName('内訳明細');
  if (!sh || sh.getLastRow() < 1) return [];
  const data = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  const result = [];
  const seenStageNums = new Set();
  let stage = '';
  for (const [raw] of data) {
    const cell = String(raw).trim();
    if (!cell) continue;
    const stageMatch = cell.match(/^[\(（](\d+)/);
    if (stageMatch) {
      const num = stageMatch[1];
      // 同じステージ番号が再登場したら以降は重複なので打ち切り
      if (seenStageNums.has(num)) break;
      seenStageNums.add(num);
      stage = cell;
      continue;
    }
    if (/諸経費|管理費|小計|合計|旅費|交通費/.test(cell)) continue;
    if (stage) result.push({ stage, role: cell });
  }
  return result;
}

// ========================================
// 履歴 / 読み込み / 削除
// ========================================
function getRecentProjects() {
  const sh = getMainSS_().getSheetByName(SHEET_NAMES.PROJECT);
  if (!sh || sh.getLastRow() < 2) return [];
  const colNum = Math.max(sh.getLastColumn(), PROJECT_COL.TYPE);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, colNum).getValues();
  const out = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    let dateStr = '';
    try { dateStr = r[PROJECT_COL.DATE - 1] ? new Date(r[PROJECT_COL.DATE - 1]).toLocaleDateString() : ''; } catch (e) {}
    out.push({
      id: r[PROJECT_COL.ID - 1],
      title: r[PROJECT_COL.TITLE - 1],
      client: r[PROJECT_COL.CLIENT - 1],
      date: dateStr,
      total: r[PROJECT_COL.TOTAL - 1],
      type: r[PROJECT_COL.TYPE - 1] || 'other'
    });
    if (out.length >= 100) break;
  }
  return out;
}

function getProjectDetail(projectId) {
  const ss = getMainSS_();
  const pSh = ss.getSheetByName(SHEET_NAMES.PROJECT);
  const pData = pSh.getDataRange().getValues();
  const projectRow = pData.find(r => r[PROJECT_COL.ID - 1] === projectId);
  if (!projectRow) throw new Error('案件が見つかりません');

  let discount = 0, overheadRate = 30;
  const memo = projectRow[PROJECT_COL.MEMO - 1] ? String(projectRow[PROJECT_COL.MEMO - 1]) : '';
  memo.split(',').forEach(part => {
    part = part.trim();
    if (part.startsWith('値引:')) discount = part.split(':')[1];
    if (part.startsWith('諸経費:')) overheadRate = part.split(':')[1].replace('%', '');
  });
  const savedType = projectRow[PROJECT_COL.TYPE - 1] || 'ws_pharma';

  const dSh = ss.getSheetByName(SHEET_NAMES.DETAIL);
  const dData = dSh.getDataRange().getValues();
  const items = dData
    .filter(r => r[DETAIL_COL.PROJECT_ID - 1] === projectId)
    .map(r => ({
      eventType: r[DETAIL_COL.EVENT_TYPE - 1],
      hours: r[DETAIL_COL.HOURS - 1],
      prepHours: Number(r[DETAIL_COL.PREP_HOURS - 1]) || 0,
      isMedical: r[DETAIL_COL.IS_MEDICAL - 1],
      unitPrice: r[DETAIL_COL.UNIT_PRICE - 1],
      rankName: r[DETAIL_COL.RANK_NAME - 1],
      subtotal: r[DETAIL_COL.SUBTOTAL - 1],
      extraDesc: r[DETAIL_COL.NOTE - 1] || '',
      isHeader: (r[DETAIL_COL.NOTE - 1] === 'IS_HEADER')
    }));

  const tSh = ensureTravelSheet_();
  const travelDetails = [];
  if (tSh.getLastRow() > 1) {
    const tData = tSh.getRange(2, 1, tSh.getLastRow() - 1, 9).getValues();
    tData.forEach(r => {
      if (r[TRAVEL_COL.PROJECT_ID - 1] !== projectId) return;
      const rt = r[TRAVEL_COL.IS_ROUND_TRIP - 1];
      travelDetails.push({
        type: r[TRAVEL_COL.TYPE - 1],
        from: r[TRAVEL_COL.FROM - 1] || '',
        to: r[TRAVEL_COL.TO - 1] || '',
        place: r[TRAVEL_COL.PLACE - 1] || '',
        priceIn: Number(r[TRAVEL_COL.PRICE_IN - 1]) || 0,
        count: Number(r[TRAVEL_COL.COUNT - 1]) || 1,
        isRoundTrip: rt === true || rt === 'TRUE' || rt === 'true'
      });
    });
  }

  return {
    clientName: projectRow[PROJECT_COL.CLIENT - 1],
    projectTitle: projectRow[PROJECT_COL.TITLE - 1],
    discountAmount: discount,
    overheadRate: overheadRate,
    projectType: savedType,
    items: items,
    travelDetails: travelDetails
  };
}

function deleteProject(projectId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'ロック取得失敗' };
  try {
    const ss = getMainSS_();
    deleteRowsByMatch_(ss.getSheetByName(SHEET_NAMES.PROJECT), PROJECT_COL.ID, projectId);
    deleteRowsByMatch_(ss.getSheetByName(SHEET_NAMES.DETAIL), DETAIL_COL.PROJECT_ID, projectId);
    deleteRowsByMatch_(ensureTravelSheet_(), TRAVEL_COL.PROJECT_ID, projectId);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// ヘッダー以外を読み込み→不要行除外→書き戻し (高速)
function deleteRowsByMatch_(sheet, colIndex, value) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const keep = data.filter(r => r[colIndex - 1] !== value);
  sheet.getRange(2, 1, data.length, lastCol).clearContent();
  if (keep.length > 0) {
    sheet.getRange(2, 1, keep.length, lastCol).setValues(keep);
  }
}

function ensureTravelSheet_() {
  const ss = getMainSS_();
  let sh = ss.getSheetByName(SHEET_NAMES.TRAVEL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.TRAVEL);
    sh.getRange(1, 1, 1, 9).setValues([['travelId','projectId','type','from','to','place','priceIn','count','isRoundTrip']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ========================================
// 保存
// ========================================
function saveDataToSheet(formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: '他の処理が実行中です。少し時間をおいて再度お試しください。' };
  }
  try {
    const ss = getMainSS_();
    const pSh = ss.getSheetByName(SHEET_NAMES.PROJECT);
    const dSh = ss.getSheetByName(SHEET_NAMES.DETAIL);
    const tSh = ensureTravelSheet_();
    const dateStr = new Date().toLocaleDateString();
    const travelTotal = parseInt(formData.travelTotal) || 0;
    const memoStr = '値引:' + formData.discountAmount + ', 諸経費:' + formData.overheadRate + '%, 旅費:' + travelTotal;

    let projectId = formData.projectId;
    let isUpdate = false;

    if (projectId) {
      const pData = pSh.getDataRange().getValues();
      const idx = pData.findIndex(r => r[PROJECT_COL.ID - 1] === projectId);
      if (idx > 0) {
        isUpdate = true;
        const row = [projectId, formData.projectTitle, formData.clientName, dateStr,
                     'Web入力', '作成中', memoStr, formData.grandTotal, formData.selectedType];
        pSh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
        deleteRowsByMatch_(dSh, DETAIL_COL.PROJECT_ID, projectId);
        deleteRowsByMatch_(tSh, TRAVEL_COL.PROJECT_ID, projectId);
      }
    }
    if (!projectId || !isUpdate) {
      projectId = 'PROJ_' + new Date().getTime();
      pSh.appendRow([projectId, formData.projectTitle, formData.clientName, dateStr,
                     'Web入力', '作成中', memoStr, formData.grandTotal, formData.selectedType]);
    }

    // 明細を一括書き込み (高速)
    if (formData.items && formData.items.length > 0) {
      const rows = formData.items.map(it => [
        'DET_' + Math.random().toString(36).slice(-8),
        projectId, it.eventType, it.hours, it.isMedical,
        it.unitPrice, it.rankName, it.subtotal,
        it.isHeader ? 'IS_HEADER' : (it.extraDesc || ''),
        it.prepHours || 0
      ]);
      const startRow = dSh.getLastRow() + 1;
      dSh.getRange(startRow, 1, rows.length, 10).setValues(rows);
    }

    // 旅費明細を一括書き込み
    if (formData.travelDetails && formData.travelDetails.length > 0) {
      const rows = formData.travelDetails.map(t => [
        'TRV_' + Math.random().toString(36).slice(-8),
        projectId, t.type, t.from || '', t.to || '',
        t.place || '', t.priceIn || 0, t.count || 1, !!t.isRoundTrip
      ]);
      const startRow = tSh.getLastRow() + 1;
      tSh.getRange(startRow, 1, rows.length, 9).setValues(rows);
    }

    const result = updateSpreadsheetTemplate_(formData, dateStr);

    return {
      success: true,
      url: result.url,
      pdfUrl: result.pdfUrl,
      projectId: projectId,
      sheetTotal: result.sheetTotal,
      message: isUpdate ? '上書き保存しました！' : '新規保存しました！'
    };
  } catch (e) {
    console.error(e.stack || e);
    return { success: false, message: 'エラーが発生しました:\n' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function normalizeKey_(str) {
  if (!str) return '';
  return String(str)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[（）()]/g, '')
    .replace(/[ 　]/g, '')
    .trim();
}

// ========================================
// 見積書テンプレ生成 (高速版: 列ごとに getValues→加工→setValues)
// ========================================
function updateSpreadsheetTemplate_(data, dateStr) {
  const d = new Date();
  const safeDate = d.getFullYear() + ('0'+(d.getMonth()+1)).slice(-2) + ('0'+d.getDate()).slice(-2);
  const fileName = '【見積内訳】' + data.clientName + '様_' + data.projectTitle + '_' + safeDate;

  const copied = Drive.Files.copy({ title: fileName, parents: [{id: DESTINATION_FOLDER_ID}] }, TEMPLATE_SS_ID);
  const newSs = SpreadsheetApp.openById(copied.id);
  let targetName = '内訳明細';
  if (data.selectedType === 'graphic') targetName = '内訳明細グラレコのみ';
  else if (data.selectedType === 'project') targetName = '内訳明細イラストのみ';

  newSs.getSheets().forEach(s => {
    if (s.getName() !== targetName) { try { newSs.deleteSheet(s); } catch (e) {} }
  });

  let sheet = newSs.getSheetByName(targetName);
  if (!sheet) throw new Error('テンプレートに「' + targetName + '」シートが見つかりません');

  // タイプ別の印刷範囲 (リネーム前の元シート名で判定)
  const printRanges = {
    '内訳明細': 'A1:F47',
    '内訳明細グラレコのみ': 'A1:F21',
    '内訳明細イラストのみ': 'A1:F45'
  };
  const printRange = printRanges[targetName] || 'A1:F50';

  if (sheet.getName() !== '内訳明細') sheet.setName('内訳明細');

  // PDF出力用URL (A4縦・印刷範囲指定済み・現在のシート状態で出力)
  const sheetGid = sheet.getSheetId();
  const pdfUrl = 'https://docs.google.com/spreadsheets/d/' + copied.id +
    '/export?format=pdf&size=a4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false' +
    '&gid=' + sheetGid + '&range=' + printRange;

  if (targetName === '内訳明細') {
    sheet.getRange('A3').setValue(data.clientName + ' 御中');
    sheet.getRange('A5').setValue('件名: ' + data.projectTitle);
  }
  sheet.getRange('F1').setValue(dateStr);
  // 印刷範囲外 (G1) にPDFリンクを設置 (印刷時は出ない)
  sheet.getRange('G1').setFormula('=HYPERLINK("' + pdfUrl + '","📄 印刷用PDFを開く")')
    .setFontWeight('bold').setFontColor('#1a73e8').setFontSize(11);

  // メインテーブル: A:E を一括読み込み (5回 → 2回のRPCに圧縮)
  const lastRow = sheet.getLastRow();
  const mainRange = sheet.getRange(1, 1, lastRow, 5);
  const allValues = mainRange.getValues();
  const allFormulas = mainRange.getFormulas();
  const colA = allValues.map(r => [r[0]]);
  const colB = allValues.map(r => [r[1]]);
  const colC = allValues.map(r => [r[2]]);
  const colD = allValues.map(r => [r[3]]);
  const colE = allValues.map(r => [r[4]]);
  const modB = new Array(lastRow).fill(false);
  const modD = new Array(lastRow).fill(false);
  const modE = new Array(lastRow).fill(false);

  const masterPrices = getUnitPriceMaster_();
  const itemMap = {};
  data.items.forEach(it => {
    if (it.isHeader) return;
    // 本番時間 + (準備等時間 ÷ 2) = テンプレに書き込む実効時間 (準備等は半額のため)
    const effectiveHours = (Number(it.hours) || 0) + (Number(it.prepHours) || 0) / 2;
    const obj = { hours: effectiveHours, price: it.unitPrice };
    itemMap[normalizeKey_(it.eventType + '_' + it.rankName)] = obj;
    itemMap[normalizeKey_(it.rankName)] = obj;
  });

  let stage = '';
  let totalRowIndex = -1;
  const travelTotal = parseInt(data.travelTotal) || 0;
  const people = parseInt(data.grPeople) || 1;

  for (let i = 0; i < colA.length; i++) {
    const cellA = String(colA[i][0]).trim();
    if (cellA === '合計') totalRowIndex = i + 1;
    if (!cellA) continue;

    // グラレコ専用 (B7=人数, D7=1人分小計式, E7=B7*D7。B9〜B11 は1人分の値)
    if (data.selectedType === 'graphic') {
      const findHrs = (name, fallback) => {
        const it = data.items.find(x => !x.isHeader && x.rankName === name);
        return it ? (Number(it.hours) || 0) : fallback;
      };
      if (cellA === 'グラフィックレコーディング' && (!colD[i][0] || colD[i][0] === '')) { colB[i][0] = people; modB[i]=true; continue; }
      if (cellA === '基本料金') { colB[i][0] = findHrs('基本料金', 1); colD[i][0] = 80000; modB[i]=true; modD[i]=true; continue; }
      if (cellA === 'グラフィックレコーディング' && colC[i][0] === '時間') { colB[i][0] = findHrs('グラフィックレコーディング', 2); colD[i][0] = 20000; modB[i]=true; modD[i]=true; continue; }
      if (cellA === '待機') { colB[i][0] = findHrs('待機', 1); colD[i][0] = 10000; modB[i]=true; modD[i]=true; continue; }
    }

    if (cellA.indexOf('諸経費') > -1) {
      colD[i][0] = (parseFloat(data.overheadRate) || 30) / 100;
      modD[i] = true;
      continue;
    }
    if (cellA.indexOf('旅費') > -1 || cellA.indexOf('交通費') > -1) {
      if (travelTotal > 0) { colB[i][0] = 1; colE[i][0] = travelTotal; }
      else { colB[i][0] = ''; colE[i][0] = ''; }
      modB[i] = true; modE[i] = true;
      continue;
    }
    if (cellA.indexOf('出精値引き') > -1 && data.discountAmount > 0) {
      colE[i][0] = -data.discountAmount;
      modE[i] = true;
      continue;
    }
    if (cellA.match(/^[\(（]\d+/)) { stage = cellA; continue; }

    if (data.selectedType !== 'graphic') {
      const mapKey = normalizeKey_(stage) + '_' + normalizeKey_(cellA);
      const onlyRoleKey = normalizeKey_(cellA);
      if (itemMap[mapKey]) {
        colB[i][0] = itemMap[mapKey].hours;
        colD[i][0] = itemMap[mapKey].price;
        modB[i]=true; modD[i]=true;
      } else if (itemMap[onlyRoleKey] && !stage) {
        colB[i][0] = itemMap[onlyRoleKey].hours;
        colD[i][0] = itemMap[onlyRoleKey].price;
        modB[i]=true; modD[i]=true;
      } else if (masterPrices[cellA]) {
        colD[i][0] = masterPrices[cellA];
        colB[i][0] = '';
        modB[i]=true; modD[i]=true;
      }
    }
  }

  // 一括書き戻し: 修正していないセルは元の数式を保持 (B,C,D,E を1回のsetValuesで)
  const writeData = new Array(lastRow);
  for (let i = 0; i < lastRow; i++) {
    const bF = allFormulas[i][1];
    const cF = allFormulas[i][2];
    const dF = allFormulas[i][3];
    const eF = allFormulas[i][4];
    writeData[i] = [
      modB[i] ? colB[i][0] : (bF || colB[i][0]),
      cF || colC[i][0],
      modD[i] ? colD[i][0] : (dF || colD[i][0]),
      modE[i] ? colE[i][0] : (eF || colE[i][0])
    ];
  }
  sheet.getRange(1, 2, lastRow, 4).setValues(writeData);

  // 旅費明細 (H6〜M)
  sheet.getRange('H6:M').clearContent();
  sheet.getRange('H6:M6').setValues([['項目','詳細 (区間・宿泊地など)','片道単価(税込)','片道/往復','人数','小計(税抜)']])
    .setFontWeight('bold').setBackground('#e3f2fd');
  if (data.travelDetails && data.travelDetails.length > 0) {
    const startRow = 7;
    const rows = data.travelDetails.map(d => {
      const detailStr = (d.type === '交通費') ? (d.from + '～' + d.to) : (d.place || '');
      const wayStr = (d.type === '交通費') ? (d.isRoundTrip ? '往復' : '片道') : '－';
      return [d.type, detailStr, d.priceIn, wayStr, d.count];
    });
    sheet.getRange(startRow, 8, rows.length, 5).setValues(rows);
    const formulas = rows.map((_, idx) => {
      const r = startRow + idx;
      return ['=ROUND(J'+r+'/1.1) * L'+r+' * IF(K'+r+'="往復", 2, 1)'];
    });
    sheet.getRange(startRow, 13, rows.length, 1).setFormulas(formulas);
    const totalRow = startRow + rows.length;
    sheet.getRange(totalRow, 12).setValue('合計(税抜):').setFontWeight('bold');
    sheet.getRange(totalRow, 13).setFormula('=SUM(M' + startRow + ':M' + (totalRow - 1) + ')').setFontWeight('bold');
    // 列幅はテンプレートで事前設定推奨 (毎回設定すると ~600ms 遅くなる)
  }

  // 共有設定: Advanced Drive API で setSharing より高速
  try {
    Drive.Permissions.insert(
      { role: 'writer', type: 'anyone', withLink: true },
      copied.id,
      { sendNotificationEmails: false }
    );
  } catch (e) {
    console.warn('setSharing skipped (domain policy): ' + e);
  }
  // SpreadsheetApp.flush()と sheetTotal読み取りは省略 (体感速度優先)
  return { url: 'https://docs.google.com/spreadsheets/d/' + copied.id + '/edit', pdfUrl: pdfUrl, sheetTotal: 0 };
}

function authorizeDrive() {
  Drive.Files.list({maxResults: 1});
  DriveApp.getRootFolder().getName();
  console.log('認証完了！');
}
