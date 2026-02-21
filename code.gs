/**
 * 体育ノート - Server Logic
 * GIGA Standard v2 Compliant
 * [Auto-Recovery & Dynamic Header Mapping Enabled]
 */

const APP_NAME = '体育ノート';
const PROPERTIES = PropertiesService.getScriptProperties();

// ==== 定義（必須シートと必須列） ====
const DEF_MEMBERS = ['メールアドレス', '出席番号', '氏名', '権限'];
const DEF_LOGS = ['タイムスタンプ', 'メールアドレス', '単元名', '入力タイプ', '観点', 'コメント', 'deletedAt', '宛先', '単元ID', '授業回'];
const DEF_CONFIG = ['項目', '値'];
const DEF_UNITS = ['単元ID', '単元名', '総時間数', '単元目標', '作成日時', 'ステータス', '授業詳細JSON'];

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
    // 削除されている・権限がない場合の自己修復：自動で新規作成してリカバリする
    ss = SpreadsheetApp.create(APP_NAME + ' データ (自動復旧)');
    PROPERTIES.setProperty('SPREADSHEET_ID', ss.getId());
    // 古いシート1が残るかもしれないので確保処理後に消すなどはお好みですが、ここでは最低限新規SSに置き換えます
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

  // スプレッドシート側の仕様変更により「現在の単元」などは設定シートを使わない
  // （下位互換性のため残すことは可能だが、今後は単元マスターのみ参照）

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
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://drive.google.com/uc?id=1g7PlclDf5CynFyHrBCBI9yZD_xqSChYn&.png');
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
    
    // お手本フォルダも同時に作成・プロパティ登録
    getHealthyMediaFolder();
    
    // デフォルトのシート1が残っていれば削除
    const sheet1 = ss.getSheetByName('シート1');
    if (sheet1) ss.deleteSheet(sheet1);
    
    // ScriptApp.getService().getUrl() で現在のウェブアプリ（実行可能API）URLを取得してフロントへ渡す
    return { success: true, url: ss.getUrl(), appUrl: ScriptApp.getService().getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// お手本メディアフォルダ管理 (Auto-Recovery)
// ==========================================

/**
 * 常に健全なお手本フォルダを返す。
 * MEDIA_FOLDER_ID プロパティが未設定 / フォルダが削除された場合は自動作成して復旧する。
 */
function getHealthyMediaFolder() {
  const folderId = PROPERTIES.getProperty('MEDIA_FOLDER_ID');
  
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      // ゴミ箱に入っていたら復旧できないので別扱い
      if (!folder.isTrashed()) return folder;
    } catch (e) {
      // フォルダが見つからない → 下で再作成
    }
  }
  
  // 自己修復：新しいフォルダを作成してプロパティに登録
  const newFolder = DriveApp.createFolder(APP_NAME + ' お手本フォルダ');
  PROPERTIES.setProperty('MEDIA_FOLDER_ID', newFolder.getId());
  Logger.log('お手本フォルダを新規作成しました: ' + newFolder.getId());
  return newFolder;
}

/**
 * Googleドライブ全体からメディアファイル（画像・スライド・動画）を検索して一覧を返す
 */
function listMediaFiles(searchQuery = '') {
  try {
    const IMAGE_TYPES = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    const SLIDE_TYPE  = 'application/vnd.google-apps.presentation';
    const VIDEO_TYPES = ['video/mp4', 'video/quicktime', 'video/x-msvideo'];

    let mimeCondition = '(' + [
      ...IMAGE_TYPES.map(t => "mimeType='" + t + "'"),
      "mimeType='" + SLIDE_TYPE + "'",
      ...VIDEO_TYPES.map(t => "mimeType='" + t + "'")
    ].join(' or ') + ')';

    let q = mimeCondition + " and trashed = false";
    
    // キーワードが指定されている場合はファイル名で絞り込み
    if (searchQuery && typeof searchQuery === 'string' && searchQuery.trim() !== '') {
      let escapedQuery = searchQuery.replace(/'/g, "\\'"); // 'をエスケープ
      q += " and title contains '" + escapedQuery + "'";
    }

    const files = DriveApp.searchFiles(q);
    const result = [];
    
    // 最大50件取得
    while (files.hasNext() && result.length < 50) {
      const file = files.next();
      const mime = file.getMimeType();
      let type = 'other';
      
      if (IMAGE_TYPES.indexOf(mime) !== -1) type = 'image';
      else if (mime === SLIDE_TYPE) type = 'slide';
      else if (VIDEO_TYPES.indexOf(mime) !== -1) type = 'video';
      
      result.push({
        id: file.getId(),
        name: file.getName(),
        type: type,
        url: type === 'slide'
          ? 'https://docs.google.com/presentation/d/' + file.getId() + '/preview'
          : 'https://drive.google.com/file/d/' + file.getId() + '/view'
      });
    }
    
    // 名前順ソート
    result.sort(function(a, b) { return a.name.localeCompare(b.name, 'ja'); });
    return { success: true, files: result };
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
  const { ss, configSheet, configMap, membersSheet, membersMap, logSheet, logMap } = getHealthySpreadsheet();
  
  // アクティブな複数単元をマスターから取得
  let activeUnits = [];
  const unitSheet = ss.getSheetByName('単元マスター');
  if (unitSheet) {
    const uData = unitSheet.getDataRange().getValues();
    const uMap = ensureHeadersAndGetMap(unitSheet, DEF_UNITS);
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][uMap['ステータス']] === 'active') {
        let details = [];
        try {
          details = JSON.parse(uData[i][uMap['授業詳細JSON']] || '[]');
        } catch(e) { /* parse error */ }
        
        activeUnits.push({
          id: uData[i][uMap['単元ID']],
          name: uData[i][uMap['単元名']],
          totalSessions: uData[i][uMap['総時間数']],
          goal: uData[i][uMap['単元目標']],
          sessions: details
        });
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
        comment: row[logMap['コメント']],
        unit: row[logMap['単元名']] || '',
        unitId: row[logMap['単元ID']] || '',
        session: row[logMap['授業回']] || ''
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
    activeUnits: activeUnits,
    classmates: classmates
  };
}

function saveLog(email, type, aspect, comment, targetEmail = "", selectedUnitName = "", selectedUnitId = "", selectedSession = "") {
  const { logSheet, logMap } = getHealthySpreadsheet();
  
  // マッピングを使って正しく配列を組み立てる（列が入れ替わっていても対応）
  const maxCol = Math.max(...Object.values(logMap)) + 1;
  let newRow = new Array(maxCol).fill("");
  
  newRow[logMap['タイムスタンプ']] = new Date();
  newRow[logMap['メールアドレス']] = String(email).trim();
  newRow[logMap['単元名']] = selectedUnitName || '未設定';
  newRow[logMap['入力タイプ']] = type;
  newRow[logMap['観点']] = aspect;
  newRow[logMap['コメント']] = comment;
  newRow[logMap['deletedAt']] = "";
  newRow[logMap['単元ID']] = selectedUnitId || '';
  newRow[logMap['授業回']] = selectedSession || '';
  if (logMap['宛先'] !== undefined) {
    newRow[logMap['宛先']] = targetEmail;
  }

  logSheet.appendRow(newRow);
  return { success: true };
}

function getTeacherData() {
  const { ss, configSheet, configMap, membersSheet, membersMap, logSheet, logMap } = getHealthySpreadsheet();
  
  // アクティブな複数単元をマスターから取得
  let activeUnits = [];
  const unitSheet = ss.getSheetByName('単元マスター');
  if (unitSheet) {
    const uData = unitSheet.getDataRange().getValues();
    const uMap = ensureHeadersAndGetMap(unitSheet, DEF_UNITS);
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][uMap['ステータス']] === 'active') {
        let details = [];
        try {
          details = JSON.parse(uData[i][uMap['授業詳細JSON']] || '[]');
        } catch(e) {}
        
        activeUnits.push({
          id: uData[i][uMap['単元ID']],
          name: uData[i][uMap['単元名']],
          totalSessions: uData[i][uMap['総時間数']],
          goal: uData[i][uMap['単元目標']],
          sessions: details
        });
      }
    }
  }

  const mData = membersSheet.getDataRange().getValues();
  const lData = logSheet.getDataRange().getValues();
  
  // ==== クラスポートフォリオ用集計 ====
  let classStats = {
    radarData: { '知る': 0, '見る': 0, 'する': 0, '支える': 0 },
    dailyTrend: {} // { "MM/dd": { count: 0 } }
  };
  
  // 月日フォーマット作成ヘルパー
  const getMMDD = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd");
  
  // 過去2週間の日付の器を作っておく
  const today = new Date();
  for (let d = 14; d >= 0; d--) {
    let tempDate = new Date();
    tempDate.setDate(today.getDate() - d);
    classStats.dailyTrend[getMMDD(tempDate)] = 0;
  }

  
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
        
        // クラス全体のレーダーチャート集計
        if (classStats.radarData[aspect] !== undefined) {
          classStats.radarData[aspect]++;
        }
        
        // クラス全体の日別提出推移（過去2週間以内ならカウント）
        const mmdd = getMMDD(dateObj);
        if (classStats.dailyTrend[mmdd] !== undefined) {
          classStats.dailyTrend[mmdd]++;
        }
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
    activeUnits: activeUnits,
    classStats: classStats
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

function getReferenceImageBase64(fileId) {
  if (!fileId) throw new Error('ファイルIDが指定されていません。');
  
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    // Base64テキスト化することでCORS等のブロックを回避
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    throw new Error('ファイルを取得できませんでした。共有権限を確認してください: ' + e.message);
  }
}

function getStudentDetailForTeacher(studentEmail) {
  const { membersSheet, membersMap, logSheet, logMap } = getHealthySpreadsheet();
  
  const logData = logSheet.getDataRange().getValues();
  const mData = membersSheet.getDataRange().getValues();
  
  let studentName = '不明な児童';
  let emailToName = {};
  const targetEmail = String(studentEmail).trim();

  for (let i = 1; i < mData.length; i++) {
    let email = mData[i][membersMap['メールアドレス']];
    if (!email) continue; // 空行スキップ

    let cleanEmail = String(email).trim();
    let name = mData[i][membersMap['氏名']];
    emailToName[cleanEmail] = name;
    if (cleanEmail === targetEmail) studentName = name;
  }

  let radarCounts = { '知る': 0, '見る': 0, 'する': 0, '支える': 0 };
  let logs = [];
  let receivedThanks = [];

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const logEmail = String(row[logMap['メールアドレス']] || "").trim();
    const rawDate = row[logMap['タイムスタンプ']];

    // 空行スキップ
    if (!logEmail || !rawDate) continue;
    if (row[logMap['deletedAt']]) continue; // 削除済み

    const unit = row[logMap['単元名']];
    const type = row[logMap['入力タイプ']];
    const aspect = row[logMap['観点']];
    const comment = row[logMap['コメント']];
    const targetThanksEmail = logMap['宛先'] !== undefined ? String(row[logMap['宛先']] || "").trim() : "";
    
    const dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) continue; // 不正な日付をスキップ

    if (logEmail === targetEmail && type !== 'サンクスカード') {
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

    if (type === 'サンクスカード' && targetThanksEmail === targetEmail) {
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
// 単元管理 API (Phase 1 複数単元並行稼働対応)
// ==========================================

function createUnit(unitName, totalSessions, unitGoal, sessionsJsonData) {
  const { ss, unitSheet, unitMap } = getHealthySpreadsheet();
  
  const newUnitId = 'U-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  
  // 新規単元の追加（過去の単元はアーカイブせず、複数activeを許容する）
  let newRow = new Array(DEF_UNITS.length).fill('');
  newRow[unitMap['単元ID']] = newUnitId;
  newRow[unitMap['単元名']] = unitName;
  newRow[unitMap['総時間数']] = totalSessions;
  newRow[unitMap['単元目標']] = unitGoal;
  newRow[unitMap['作成日時']] = new Date();
  newRow[unitMap['ステータス']] = 'active';
  newRow[unitMap['授業詳細JSON']] = JSON.stringify(sessionsJsonData || []);
  unitSheet.appendRow(newRow);

  return { success: true, newUnitId: newUnitId };
}

function updateUnit(unitId, unitName, totalSessions, unitGoal, sessionsJsonData) {
  const { unitSheet, unitMap } = getHealthySpreadsheet();
  const uData = unitSheet.getDataRange().getValues();
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][unitMap['単元ID']] === unitId) {
      unitSheet.getRange(i + 1, unitMap['単元名'] + 1).setValue(unitName);
      unitSheet.getRange(i + 1, unitMap['総時間数'] + 1).setValue(totalSessions);
      unitSheet.getRange(i + 1, unitMap['単元目標'] + 1).setValue(unitGoal);
      unitSheet.getRange(i + 1, unitMap['授業詳細JSON'] + 1).setValue(JSON.stringify(sessionsJsonData || []));
      return { success: true };
    }
  }
  return { success: false, message: '更新対象の単元が見つかりません' };
}

function archiveUnit(unitId) {
  const { unitSheet, unitMap } = getHealthySpreadsheet();
  const uData = unitSheet.getDataRange().getValues();
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][unitMap['単元ID']] === unitId) {
      unitSheet.getRange(i + 1, unitMap['ステータス'] + 1).setValue('archived');
      return { success: true };
    }
  }
  return { success: false, message: '単元が見つかりません' };
}

// (updateSession, updateConfigValueなどの古い設定APIは廃止)
