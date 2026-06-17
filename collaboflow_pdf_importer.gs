// ==========================================
// 設定（ここだけ変更してください）
// ==========================================
const CONFIG = {
  // PDFを置くGoogleドライブのフォルダID
  PDF_FOLDER_ID: 'ここにフォルダIDを入力',

  // 書き込み先スプレッドシートID
  SPREADSHEET_ID: 'ここにスプレッドシートIDを入力',

  // シート名
  SHEET_APPLICATIONS: '申請一覧',
  SHEET_VISITORS:     '入館者一覧',
  SHEET_LOG:          '処理ログ',

  // 処理済みPDFの移動先フォルダID（空文字なら移動しない）
  DONE_FOLDER_ID: '',
};

// ==========================================
// メイン処理（手動実行 or トリガーで呼び出す）
// ==========================================
function processNewPDFs() {
  const folder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const files  = folder.getFilesByType(MimeType.PDF);

  // 未処理ファイルを全件収集
  const allFiles = [];
  while (files.hasNext()) {
    const f = files.next();
    if (!isAlreadyProcessed(f.getId())) {
      allFiles.push(f);
    }
  }

  if (allFiles.length === 0) {
    Logger.log('処理対象のPDFはありません。');
    return;
  }

  Logger.log(`処理対象: ${allFiles.length} 件`);

  let successCount = 0;
  let errorCount   = 0;

  allFiles.forEach((file, index) => {
    Logger.log(`[${index + 1}/${allFiles.length}] 処理中: ${file.getName()}`);
    try {
      const data = extractDataFromPDF(file);
      if (data) {
        writeToSpreadsheet(data);
        markAsProcessed(file.getId(), file.getName());

        // 処理済みフォルダへ移動
        if (CONFIG.DONE_FOLDER_ID) {
          const doneFolder = DriveApp.getFolderById(CONFIG.DONE_FOLDER_ID);
          doneFolder.addFile(file);
          folder.removeFile(file);
        }

        successCount++;
        writeLog(file.getName(), '成功', `入館者 ${data.visitors.length} 名`);
        Logger.log(`  → 完了 (入館者 ${data.visitors.length} 名)`);
      }
    } catch (e) {
      errorCount++;
      writeLog(file.getName(), 'エラー', e.message);
      Logger.log(`  → エラー: ${e.message}`);
    }

    // Drive API の過負荷を避けるため少し待機
    if (index < allFiles.length - 1) Utilities.sleep(1500);
  });

  Logger.log(`完了 — 成功: ${successCount} 件 / エラー: ${errorCount} 件`);
}

// ==========================================
// PDFからデータ抽出
// ==========================================
function extractDataFromPDF(pdfFile) {
  // PDFをGoogleドキュメントに変換してテキスト取得
  const tempDoc = Drive.Files.copy(
    { title: '__temp__' + pdfFile.getId(), mimeType: MimeType.GOOGLE_DOCS },
    pdfFile.getId(),
    { convert: true }
  );

  Utilities.sleep(2000); // 変換完了待ち

  let text = '';
  try {
    const doc = DocumentApp.openById(tempDoc.id);
    text = doc.getBody().getText();
  } finally {
    DriveApp.getFileById(tempDoc.id).setTrashed(true); // 必ず一時ファイルを削除
  }

  Logger.log('--- 抽出テキスト ---\n' + text + '\n---');

  return {
    fileName:       pdfFile.getName(),
    fileId:         pdfFile.getId(),
    company:        extractField(text, /会社名[^\n]*\n([^\n]+)/),
    name:           extractField(text, /氏名[^\n]*\n([^\n]+)/),
    phone:          extractField(text, /電話番号[^\n]*\n([^\n]+)/),
    startDate:      extractField(text, /入館開始日[^\n]*\n(\d{4}\/\d{2}\/\d{2})/),
    endDate:        extractField(text, /入館終了日[^\n]*\n(\d{4}\/\d{2}\/\d{2})/),
    entryTime:      extractField(text, /入館予定時刻[^\n]*\n[^\n]*\n(\d{2}:\d{2})/),
    exitTime:       extractField(text, /退館予定時刻[^\n]*\n[^\n]*\n(\d{2}:\d{2})/),
    entryType:      extractField(text, /(連続入館|断続入館|両方)/),
    authorizedCard: extractField(text, /オーソライズドカード[^\n]*\n([^\n]+)/),
    purpose:        extractField(text, /入館目的[^\n]*\n([^\n]+)/),
    accessArea:     extractAccessArea(text),
    loadingWork:    extractField(text, /搬出入作業[^\n]*\n(有|無)/),
    hazardous:      extractField(text, /火気[^\n]*\n(有|無)/),
    status:         extractField(text, /申請状況：\s*([^\n]+)/),
    processedAt:    new Date().toLocaleString('ja-JP'),
    visitors:       extractVisitors(text),
  };
}

// ==========================================
// 正規表現でフィールド抽出
// ==========================================
function extractField(text, regex) {
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

// ==========================================
// 入室箇所の抽出
// ==========================================
function extractAccessArea(text) {
  const definitions = [
    { pattern: /yes\s+1[Ff]\s*Common/i,   label: '1F共用部' },
    { pattern: /yes\s+2[Ff]\s*Common/i,   label: '2F共用部' },
    { pattern: /yes\s+4[Ff]\s*Common/i,   label: '4F共用部' },
    { pattern: /yes\s+2[Ff]\s*Office/i,   label: '2Fオフィス201' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?101/i, label: '1F会議室101' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?102/i, label: '1F会議室102' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?103/i, label: '1F会議室103' },
  ];

  const found = definitions.filter(d => d.pattern.test(text)).map(d => d.label);
  if (found.length > 0) return found.join(', ');

  // フォールバック: セクション全体から最初の意味ある行を取得
  const m = text.match(/入室箇所[\s\S]*?(?=搬出入作業)/);
  if (m) {
    return m[0]
      .replace(/入室箇所[^\n]*/,'')
      .split('\n')
      .map(l => l.trim())
      .filter(l => l && !/^yes$/i.test(l) && !/^no$/i.test(l))
      .join(', ')
      .trim();
  }
  return '';
}

// ==========================================
// 入館者リストの抽出
// ==========================================
function extractVisitors(text) {
  const visitors = [];

  // 入館者情報セクションを切り出す
  const sectionMatch = text.match(/入館者情報[\s\S]*?(?=申請状況|$)/);
  if (!sectionMatch) return visitors;

  const lines = sectionMatch[0]
    .split('\n')
    .map(l => l.trim())
    .filter(l => l && !l.startsWith('入館者') && !l.startsWith('Visitor') && !l.startsWith('取り込み'));

  const phoneRegex = /^[\d\+][\d\-\+\s]{7,}$/;

  for (let i = 0; i < lines.length; i++) {
    const cleaned = lines[i].replace(/\s/g, '');
    if (phoneRegex.test(cleaned)) {
      const phone       = lines[i].trim();
      const visitorName = i >= 1 ? lines[i - 1].trim() : '';
      const company     = i >= 2 ? lines[i - 2].trim() : '';

      if (visitorName && !visitorName.includes('Download') && !visitorName.match(/^\d+$/)) {
        visitors.push({ company, name: visitorName, phone });
      }
    }
  }

  return visitors;
}

// ==========================================
// スプレッドシートへの書き込み
// ==========================================
function writeToSpreadsheet(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // --- 申請一覧シート ---
  const appSheet = getOrCreateSheet(ss, CONFIG.SHEET_APPLICATIONS, [
    'ファイル名', '申請者会社名', '申請者氏名', '申請者電話番号',
    '入館開始日', '入館終了日', '入館予定時刻', '退館予定時刻',
    '連続断続区分', 'オーソライズドカード', '入館目的', '入室箇所',
    '搬出入作業', '火気危険物', '申請状況', '処理日時',
  ]);

  appSheet.appendRow([
    data.fileName,   data.company,   data.name,     data.phone,
    data.startDate,  data.endDate,   data.entryTime, data.exitTime,
    data.entryType,  data.authorizedCard, data.purpose, data.accessArea,
    data.loadingWork, data.hazardous, data.status,   data.processedAt,
  ]);

  // --- 入館者一覧シート ---
  const visitorSheet = getOrCreateSheet(ss, CONFIG.SHEET_VISITORS, [
    'ファイル名', '申請者会社名', '入館開始日', '入館者会社名', '入館者氏名', '入館者電話番号',
  ]);

  data.visitors.forEach(v => {
    visitorSheet.appendRow([
      data.fileName, data.company, data.startDate,
      v.company, v.name, v.phone,
    ]);
  });
}

// ==========================================
// ログシートへの書き込み
// ==========================================
function writeLog(fileName, status, message) {
  const ss        = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const logSheet  = getOrCreateSheet(ss, CONFIG.SHEET_LOG, ['処理日時', 'ファイル名', '状況', 'メッセージ']);
  logSheet.appendRow([new Date().toLocaleString('ja-JP'), fileName, status, message]);
}

// ==========================================
// シート取得（なければ作成してヘッダー設定）
// ==========================================
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4472C4')
      .setFontColor('white')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 150);
  }
  return sheet;
}

// ==========================================
// 処理済み管理（PropertiesServiceで記録）
// ==========================================
function isAlreadyProcessed(fileId) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('processed_' + fileId) !== null;
}

function markAsProcessed(fileId, fileName) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('processed_' + fileId, new Date().toISOString() + ' | ' + fileName);
}

// 処理済み記録をリセット（再処理したい場合に手動実行）
function resetProcessedHistory() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('処理済み履歴をリセットしました。');
}

// ==========================================
// 5分ごとの自動実行トリガー設定（初回1回だけ実行）
// ==========================================
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processNewPDFs')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('トリガー設定完了（5分ごとに自動実行）');
}
