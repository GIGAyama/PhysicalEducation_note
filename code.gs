/**
 * スポ☆活ダッシュボード - Server Logic
 * GIGA Standard v2 Compliant
 * [Auto-Recovery & Dynamic Header Mapping Enabled]
 */

const APP_NAME = 'スポ☆活ダッシュボード';
const PROPERTIES = PropertiesService.getScriptProperties();

// ==== 定義（必須シートと必須列） ====
const DEF_MEMBERS = ['メールアドレス', '出席番号', '氏名', '権限'];
const DEF_LOGS = ['タイムスタンプ', 'メールアドレス', '単元名', '入力タイプ', '観点', 'コメント', 'deletedAt', '宛先', '単元ID', '授業回'];
const DEF_CONFIG = ['項目', '値'];
const DEF_UNITS = ['単元ID', '単元名', '総時間数', '単元目標', '作成日時', 'ステータス']; // ステータス: active or archived

// ==========================================
// 自己修復 & ヘッダーマッピング機能 (Auto-Recovery)
// ==========================================

/**
 * 常に健全なスプレッドシートとマッピングデータを取得する
 */
function getHealthySpreadsheet() {
  const ssId = PROPERTIES.getProperty('SPREADSHEET_ID');
  if (!ssId) throw new Error('システムが初期化されていません。');
  
  let ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  } catch (e) {
    throw new Error('スプレッドシートにアクセスできません。削除されたか権限がありません。');
  }

  // シートの存在確認と復旧
  const membersSheet = ensureSheet(ss, '名簿', DEF_MEMBERS);
  const logSheet = ensureSheet(ss, 'ログ', DEF_LOGS);
  const configSheet = ensureSheet(ss, '設定', DEF_CONFIG);
  const unitSheet = ensureSheet(ss, '単元マスター', DEF_UNITS);

  // ヘッダー列の存在確認と動的マッピング
  const membersMap = ensureHeadersAndGetMap(membersSheet, DEF_MEMBERS);
  const logMap = ensureHeadersAndGetMap(logSheet, DEF_LOGS);
  const configMap = ensureHeadersAndGetMap(configSheet, DEF_CONFIG);
  const unitMap = ensureHeadersAndGetMap(unitSheet, DEF_UNITS);

  // 設定シートの必須項目デフォルト値確保
  ensureConfigDefault(configSheet, configMap, '現在の単元名', '未設定'); // 下位互換用
  ensureConfigDefault(configSheet, configMap, '現在の単元ID', '');
  ensureConfigDefault(configSheet, configMap, '現在の授業回', '1');

  return { ss, membersSheet, membersMap, logSheet, logMap, configSheet, configMap, unitSheet, unitMap };
}

/**
 * シートが存在しなければ作成する
 */
function ensureSheet(ss, sheetName, defaultHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(defaultHeaders);
    sheet.setFrozenRows(1);
    formatHeader(sheet);
    
    // 名簿シート作成時は実行者を教員として登録
    if (sheetName === '名簿') {
      sheet.appendRow([Session.getActiveUser().getEmail(), '0', '先生', 'teacher']);
    }
  }
  return sheet;
}

/**
 * 必須列が存在するか確認し、なければ追加。列名とインデックス(0始まり)の連想配列を返す
 */
function ensureHeadersAndGetMap(sheet, requiredHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
  }
  
  let lastCol = sheet.getLastColumn() || 1;
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let map = {};
  let headersModified = false;

  // 既存ヘッダーのインデックスをマッピング
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] !== "") map[headers[i]] = i;
  }

  // 必須ヘッダーが欠損している場合は末尾に追加して復旧
  for (let i = 0; i < requiredHeaders.length; i++) {
    let req = requiredHeaders[i];
    if (!(req in map)) {
      let nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(req);
      map[req] = nextCol - 1;
      headersModified = true;
    }
  }

  if (headersModified) {
    formatHeader(sheet);
  }
  return map;
}

/**
 * 設定シートに特定のキーが存在しなければデフォルト値を追加する
 */
function ensureConfigDefault(sheet, map, key, defaultValue) {
  const data = sheet.getDataRange().getValues();
  const keyIdx = map['項目'];
  const valIdx = map['値'];
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyIdx] === key) return;
  }
  let newRow = [];
  newRow[keyIdx] = key;
  newRow[valIdx] = defaultValue;
  sheet.appendRow(newRow);
}

function formatHeader(sheet) {
  const lastCol = sheet.getLastColumn() || 1;
  const range = sheet.getRange(1, 1, 1, lastCol);
  range.setBackground("#f8f9fa");
  range.setFontWeight("bold");
}


// ==========================================
// メイン処理
// ==========================================

function doGet(e) {
  const ssId = PROPERTIES.getProperty('SPREADSHEET_ID');
  const template = HtmlService.createTemplateFromFile('index');
  template.isSetup = !ssId;
  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initializeSetup() {
  try {
    const ss = SpreadsheetApp.create(APP_NAME + ' データ');
    PROPERTIES.setProperty('SPREADSHEET_ID', ss.getId());
    
    // getHealthySpreadsheetを呼び出すことで自動的に必要な全シートとヘッダーが生成される
    getHealthySpreadsheet();
    
    // デフォルトのシート1が残っていれば削除
    const sheet1 = ss.getSheetByName('シート1');
    if (sheet1) ss.deleteSheet(sheet1);
    
    return { success: true, url: ss.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getUserInfo() {
  const email = Session.getActiveUser().getEmail();
  const { membersSheet, membersMap } = getHealthySpreadsheet();
  
  const data = membersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][membersMap['メールアドレス']];
    if (!rowEmail) continue; // 空行スキップ
    
    if (rowEmail === email && rowEmail !== "") {
      return {
        email: email,
        number: data[i][membersMap['出席番号']],
        name: data[i][membersMap['氏名']],
        role: data[i][membersMap['権限']]
      };
    }
  }
  return { email: email, number: '?', name: 'ゲスト（名簿未登録）', role: 'student' };
}

function getStudentData(email) {
  const { configSheet, configMap, membersSheet, membersMap, logSheet, logMap } = getHealthySpreadsheet();
  
  // 設定取得
  const configData = configSheet.getDataRange().getValues();
  let currentUnit = '未設定';
  let currentUnitId = '';
  let currentSession = '1';
  let refMediaId = '';
  for(let i = 1; i < configData.length; i++){
    const k = configData[i][configMap['項目']];
    const v = configData[i][configMap['値']];
    if(k === '現在の単元名') currentUnit = v;
    if(k === '現在の単元ID') currentUnitId = v;
    if(k === '現在の授業回') currentSession = v;
    if(k === 'お手本メディアID') refMediaId = v;
  }

  // 単元マスターから詳細情報取得
  let totalSessions = '';
  let unitGoal = '';
  if (currentUnitId) {
    const unitSheet = ss.getSheetByName('単元マスター');
    if (unitSheet) {
      const uData = unitSheet.getDataRange().getValues();
      const uMap = ensureHeadersAndGetMap(unitSheet, DEF_UNITS);
      for (let i = 1; i < uData.length; i++) {
        if (uData[i][uMap['単元ID']] === currentUnitId) {
          totalSessions = uData[i][uMap['総時間数']];
          unitGoal = uData[i][uMap['単元目標']];
          break;
        }
      }
    }
  }

  // クラスメイト取得
  const mData = membersSheet.getDataRange().getValues();
  let classmates = [];
  let emailToName = {};
  for (let i = 1; i < mData.length; i++) {
    let rowEmail = mData[i][membersMap['メールアドレス']];
    if (!rowEmail) continue; // 空行スキップ

    let rowName = mData[i][membersMap['氏名']];
    let rowRole = mData[i][membersMap['権限']];
    
    emailToName[rowEmail] = rowName;
    if (rowRole === 'student' && rowEmail !== email) {
      classmates.push({ email: rowEmail, name: rowName });
    }
  }

  // ログ集計
  const data = logSheet.getDataRange().getValues();
  let radarCounts = { '知る': 0, '見る': 0, 'する': 0, '支える': 0 };
  let logs = [];
  let receivedThanks = [];
  
  const todayStr = new Date().toDateString();
  let hasTodayGoal = false;
  let hasTodayReflect = false;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const logEmail = row[logMap['メールアドレス']];
    const rawDate = row[logMap['タイムスタンプ']];

    // 空行や必須データがない行はスキップしてエラーを防ぐ
    if (!logEmail || !rawDate) continue; 
    if (row[logMap['deletedAt']]) continue; // 削除済み（空文字/null/undefined全対応）

    const type = row[logMap['入力タイプ']];
    const aspect = row[logMap['観点']];
    
    const dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) continue; // 不正な日付をスキップ
    
    const dateStr = dateObj.toDateString();
    const targetEmail = logMap['宛先'] !== undefined ? (row[logMap['宛先']] || "") : "";

    // 自分の記録
    if (logEmail === email && type !== 'サンクスカード') {
      if (aspect in radarCounts) radarCounts[aspect]++;
      
      logs.push({
        date: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd"),
        type: type,
        aspect: aspect,
        comment: row[logMap['コメント']]
      });

      if (dateStr === todayStr) {
        if (type === '事前目標') hasTodayGoal = true;
        if (type === '事後振り返り') hasTodayReflect = true;
      }
    }

    // サンクスカード受信
    if (type === 'サンクスカード' && targetEmail === email) {
      receivedThanks.push({
        date: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd"),
        senderName: emailToName[logEmail] || '友達',
        comment: row[logMap['コメント']]
      });
    }
  }

  return {
    radarData: [radarCounts['知る'], radarCounts['見る'], radarCounts['する'], radarCounts['支える']],
    logs: logs.reverse(),
    receivedThanks: receivedThanks.reverse(),
    hasTodayGoal: hasTodayGoal,
    hasTodayReflect: hasTodayReflect,
    currentUnit: currentUnit,
    currentUnitId: currentUnitId,
    currentSession: currentSession,
    totalSessions: totalSessions,
    unitGoal: unitGoal,
    classmates: classmates,
    refMediaId: refMediaId
  };
}

function saveLog(email, type, aspect, comment, targetEmail = "") {
  const { logSheet, logMap, configSheet, configMap } = getHealthySpreadsheet();
  
  const configData = configSheet.getDataRange().getValues();
  let currentUnit = '未設定';
  let currentUnitId = '';
  let currentSession = '1';
  for(let i = 1; i < configData.length; i++){
    const k = configData[i][configMap['項目']];
    const v = configData[i][configMap['値']];
    if(k === '現在の単元名') currentUnit = v;
    if(k === '現在の単元ID') currentUnitId = v;
    if(k === '現在の授業回') currentSession = v;
  }

  // マッピングを使って正しく配列を組み立てる（列が入れ替わっていても対応）
  const maxCol = Math.max(...Object.values(logMap)) + 1;
  let newRow = new Array(maxCol).fill("");
  
  newRow[logMap['タイムスタンプ']] = new Date();
  newRow[logMap['メールアドレス']] = email;
  newRow[logMap['単元名']] = currentUnit;
  newRow[logMap['入力タイプ']] = type;
  newRow[logMap['観点']] = aspect;
  newRow[logMap['コメント']] = comment;
  newRow[logMap['deletedAt']] = "";
  newRow[logMap['単元ID']] = currentUnitId;
  newRow[logMap['授業回']] = currentSession;
  if (logMap['宛先'] !== undefined) {
    newRow[logMap['宛先']] = targetEmail;
  }

  logSheet.appendRow(newRow);
  return { success: true };
}

function getTeacherData() {
  const { configSheet, configMap, membersSheet, membersMap, logSheet, logMap } = getHealthySpreadsheet();
  
  const configData = configSheet.getDataRange().getValues();
  let currentUnit = '未設定';
  let currentUnitId = '';
  let currentSession = '1';
  let refMediaId = '';
  for(let i = 1; i < configData.length; i++){
    const k = configData[i][configMap['項目']];
    const v = configData[i][configMap['値']];
    if(k === '現在の単元名') currentUnit = v;
    if(k === '現在の単元ID') currentUnitId = v;
    if(k === '現在の授業回') currentSession = v;
    if(k === 'お手本メディアID') refMediaId = v;
  }
  
  // 単元マスターから詳細情報取得
  let totalSessions = '';
  if (currentUnitId) {
    const unitSheet = ss.getSheetByName('単元マスター');
    if (unitSheet) {
      const uData = unitSheet.getDataRange().getValues();
      const uMap = ensureHeadersAndGetMap(unitSheet, DEF_UNITS);
      for (let i = 1; i < uData.length; i++) {
        if (uData[i][uMap['単元ID']] === currentUnitId) {
          totalSessions = uData[i][uMap['総時間数']];
          break;
        }
      }
    }
  }

  const mData = membersSheet.getDataRange().getValues();
  const lData = logSheet.getDataRange().getValues();
  
  const todayStr = new Date().toDateString();
  let students = [];
  
  for (let i = 1; i < mData.length; i++) {
    let rowEmail = mData[i][membersMap['メールアドレス']];
    if (!rowEmail) continue; // 空行スキップ

    if (mData[i][membersMap['権限']] === 'student') {
      students.push({
        email: mData[i][membersMap['メールアドレス']],
        number: mData[i][membersMap['出席番号']],
        name: mData[i][membersMap['氏名']],
        todayGoalAspect: '未',
        todayReflectAspect: '未',
        todayReflectComment: '',
        allAspects: [],
        alertBias: false
      });
    }
  }

  for (let i = 1; i < lData.length; i++) {
    const row = lData[i];
    const email = row[logMap['メールアドレス']];
    const rawDate = row[logMap['タイムスタンプ']];

    // 空行スキップ
    if (!email || !rawDate) continue;
    if (row[logMap['deletedAt']]) continue; // 削除済み
    
    const type = row[logMap['入力タイプ']];
    const aspect = row[logMap['観点']];
    
    const dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) continue; // 不正な日付をスキップ

    const dateStr = dateObj.toDateString();
    
    let student = students.find(s => s.email === email);
    if (student) {
      if (type === '事後振り返り') {
        student.allAspects.push(aspect);
      }
      if (dateStr === todayStr) {
        if (type === '事前目標') student.todayGoalAspect = aspect;
        if (type === '事後振り返り') {
          student.todayReflectAspect = aspect;
          student.todayReflectComment = row[logMap['コメント']];
        }
      }
    }
  }
  
  students.forEach(s => {
    if (s.allAspects.length >= 3) {
      const last3 = s.allAspects.slice(-3);
      if (last3[0] === last3[1] && last3[1] === last3[2]) {
        s.alertBias = last3[0];
      }
    }
  });
  
  students.sort((a, b) => (parseInt(a.number, 10) || 0) - (parseInt(b.number, 10) || 0));
  return { 
    students: students, 
    currentUnit: currentUnit, 
    currentUnitId: currentUnitId,
    currentSession: currentSession,
    totalSessions: totalSessions,
    refMediaId: refMediaId 
  };
}

function updateUnitName(newName) {
  const { configSheet, configMap } = getHealthySpreadsheet();
  const data = configSheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++){
    if(data[i][configMap['項目']] === '現在の単元名') {
      configSheet.getRange(i + 1, configMap['値'] + 1).setValue(newName);
      return {success: true, newUnit: newName};
    }
  }
  
  let newRow = [];
  newRow[configMap['項目']] = '現在の単元名';
  newRow[configMap['値']] = newName;
  configSheet.appendRow(newRow);
  return {success: true, newUnit: newName};
}

function updateRefMediaId(newId) {
  const { configSheet, configMap } = getHealthySpreadsheet();
  const data = configSheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++){
    if(data[i][configMap['項目']] === 'お手本メディアID') {
      configSheet.getRange(i + 1, configMap['値'] + 1).setValue(newId);
      return {success: true, newId: newId};
    }
  }
  
  let newRow = [];
  newRow[configMap['項目']] = 'お手本メディアID';
  newRow[configMap['値']] = newId;
  configSheet.appendRow(newRow);
  return {success: true, newId: newId};
}

function getReferenceImageBase64() {
  const { configSheet, configMap } = getHealthySpreadsheet();
  const data = configSheet.getDataRange().getValues();
  let fileId = '';
  for(let i = 1; i < data.length; i++){
    if(data[i][configMap['項目']] === 'お手本メディアID') fileId = data[i][configMap['値']];
  }
  if (!fileId) throw new Error('お手本画像が設定されていません。');
  
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  // 拡張子によらずBase64テキスト化することでCORS等のブロックを回避
  return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
}

function getStudentDetailForTeacher(studentEmail) {
  const { membersSheet, membersMap, logSheet, logMap } = getHealthySpreadsheet();
  
  const logData = logSheet.getDataRange().getValues();
  const mData = membersSheet.getDataRange().getValues();
  
  let studentName = '不明な児童';
  let emailToName = {};
  for (let i = 1; i < mData.length; i++) {
    let email = mData[i][membersMap['メールアドレス']];
    if (!email) continue; // 空行スキップ

    let name = mData[i][membersMap['氏名']];
    emailToName[email] = name;
    if (email === studentEmail) studentName = name;
  }

  let radarCounts = { '知る': 0, '見る': 0, 'する': 0, '支える': 0 };
  let logs = [];
  let receivedThanks = [];

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const logEmail = row[logMap['メールアドレス']];
    const rawDate = row[logMap['タイムスタンプ']];

    // 空行スキップ
    if (!logEmail || !rawDate) continue;
    if (row[logMap['deletedAt']]) continue; // 削除済み

    const unit = row[logMap['単元名']];
    const type = row[logMap['入力タイプ']];
    const aspect = row[logMap['観点']];
    const comment = row[logMap['コメント']];
    const targetEmail = logMap['宛先'] !== undefined ? (row[logMap['宛先']] || "") : "";
    
    const dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) continue; // 不正な日付をスキップ

    if (logEmail === studentEmail && type !== 'サンクスカード') {
      if (aspect in radarCounts) radarCounts[aspect]++;
      logs.push({
        date: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd"),
        unit: unit,
        session: row[logMap['授業回']] || '',
        type: type,
        aspect: aspect,
        comment: comment
      });
    }

    if (type === 'サンクスカード' && targetEmail === studentEmail) {
      receivedThanks.push({
        date: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd"),
        senderName: emailToName[logEmail] || '友達',
        comment: comment
      });
    }
  }

  return {
    name: studentName,
    radarData: [radarCounts['知る'], radarCounts['見る'], radarCounts['する'], radarCounts['支える']],
    logs: logs.reverse(),
    receivedThanks: receivedThanks.reverse()
  };
}

function getAllLogsForCsv() {
  const { membersSheet, membersMap, logSheet, logMap } = getHealthySpreadsheet();
  
  const mData = membersSheet.getDataRange().getValues();
  let emailToName = {};
  for (let i = 1; i < mData.length; i++) {
    let email = mData[i][membersMap['メールアドレス']];
    if (!email) continue; // 空行スキップ
    emailToName[email] = mData[i][membersMap['氏名']];
  }

  const lData = logSheet.getDataRange().getValues();
  let csvData = [['タイムスタンプ', '氏名', '単元名', '授業回', '入力タイプ', '観点', 'コメント', '宛先(サンクスカード)']];
  
  for (let i = 1; i < lData.length; i++) {
    const row = lData[i];
    const logEmail = row[logMap['メールアドレス']];
    const rawDate = row[logMap['タイムスタンプ']];

    // 空行スキップ
    if (!logEmail || !rawDate) continue;
    if (row[logMap['deletedAt']]) continue; // 削除済み
    
    const dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) continue; // 不正な日付をスキップ
    
    const timestamp = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    const name = emailToName[logEmail] || '不明';
    const unit = row[logMap['単元名']];
    const session = row[logMap['授業回']] || '';
    const type = row[logMap['入力タイプ']];
    const aspect = row[logMap['観点']];
    const comment = row[logMap['コメント']];
    const targetEmail = logMap['宛先'] !== undefined ? row[logMap['宛先']] : "";
    const targetName = targetEmail ? (emailToName[targetEmail] || '不明') : '';

    csvData.push([timestamp, name, unit, session, type, aspect, comment, targetName]);
  }
  
  return csvData;
}

// ==========================================
// 単元管理 API (Phase 1追加機能)
// ==========================================

function createUnit(unitName, totalSessions, unitGoal) {
  const { ss, configSheet, configMap, unitSheet, unitMap } = getHealthySpreadsheet();
  
  // 重複チェックなどはせず単純に新規作成する
  const newUnitId = 'U-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  
  // 全既存単元をアーカイブ
  const uData = unitSheet.getDataRange().getValues();
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][unitMap['ステータス']] === 'active') {
      unitSheet.getRange(i + 1, unitMap['ステータス'] + 1).setValue('archived');
    }
  }

  // 新規単元の追加
  let newRow = new Array(DEF_UNITS.length).fill('');
  newRow[unitMap['単元ID']] = newUnitId;
  newRow[unitMap['単元名']] = unitName;
  newRow[unitMap['総時間数']] = totalSessions;
  newRow[unitMap['単元目標']] = unitGoal;
  newRow[unitMap['作成日時']] = new Date();
  newRow[unitMap['ステータス']] = 'active';
  unitSheet.appendRow(newRow);

  // 設定シートの更新（古い「単元名」キーも一応塞いでおく）
  updateConfigValue(configSheet, configMap, '現在の単元名', unitName);
  updateConfigValue(configSheet, configMap, '現在の単元ID', newUnitId);
  updateConfigValue(configSheet, configMap, '現在の授業回', '1');

  return { success: true, newUnitId: newUnitId };
}

function updateSession(sessionNumber) {
  const { configSheet, configMap } = getHealthySpreadsheet();
  updateConfigValue(configSheet, configMap, '現在の授業回', sessionNumber.toString());
  return { success: true, session: sessionNumber };
}

// 共通ヘルパー: 設定シートの特定キーを更新する
function updateConfigValue(sheet, map, key, newValue) {
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(data[i][map['項目']] === key) {
      sheet.getRange(i + 1, map['値'] + 1).setValue(newValue);
      return;
    }
  }
  // なければ追記
  let newRow = [];
  newRow[map['項目']] = key;
  newRow[map['値']] = newValue;
  sheet.appendRow(newRow);
}
