// ==========================================
// 設定（ここだけ変更してください）
// ==========================================
const CONFIG = {
  // PDFを一時保存するGoogleドライブのフォルダID
  // フォルダURLの /folders/XXXXXXXX の部分
  PDF_FOLDER_ID: 'ここにフォルダIDを入力',

  // 書き込み先スプレッドシートID
  // スプレッドシートURLの /d/XXXXXXXX/ の部分
  SPREADSHEET_ID: 'ここにスプレッドシートIDを入力',

  // シート名
  SHEET_APPLICATIONS: '申請一覧',
  SHEET_VISITORS:     '入館者一覧',
  SHEET_LOG:          '処理ログ',

  // 処理済みPDFの移動先フォルダID（空文字なら移動しない）
  DONE_FOLDER_ID: '',
};

// ==========================================
// スプレッドシート起動時にメニューを追加
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📄 PDF取込')
    .addItem('PDFファイルを選択して取込', 'showUploadDialog')
    .addSeparator()
    .addItem('フォルダ内の未処理PDFを一括取込', 'processNewPDFs')
    .addSeparator()
    .addItem('処理済み履歴をリセット', 'resetProcessedHistory')
    .addToUi();
}

// ==========================================
// ダイアログ表示（ローカルPDFを選択して取込）
// ==========================================
function showUploadDialog() {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Noto Sans JP', sans-serif; padding: 16px; color: #333; }
    h3 { margin-top: 0; color: #4472C4; }
    .drop-zone {
      border: 2px dashed #4472C4; border-radius: 8px;
      padding: 30px; text-align: center; cursor: pointer;
      background: #f8f9ff; margin-bottom: 12px;
    }
    .drop-zone:hover { background: #eef0ff; }
    #fileInput { display: none; }
    #fileList { font-size: 13px; color: #555; margin: 8px 0; min-height: 20px; }
    .btn-primary {
      background: #4472C4; color: white; border: none;
      padding: 10px 24px; border-radius: 4px; cursor: pointer; font-size: 14px;
    }
    .btn-secondary {
      background: #888; color: white; border: none;
      padding: 10px 24px; border-radius: 4px; cursor: pointer; font-size: 14px;
      margin-left: 8px;
    }
    button:disabled { background: #aaa !important; cursor: not-allowed; }
    #status { margin-top: 12px; font-size: 13px; line-height: 1.8; }
    .ok   { color: #2e7d32; }
    .err  { color: #c62828; }
    .info { color: #1565c0; }
    progress { width: 100%; margin-top: 8px; }
  </style>
</head>
<body>
  <h3>📄 PDFファイルを取込</h3>
  <div class="drop-zone" onclick="document.getElementById('fileInput').click()">
    <div>クリックしてPDFを選択</div>
    <div style="font-size:12px; color:#888; margin-top:6px;">複数選択可（Ctrl/Cmd+クリック）</div>
  </div>
  <input type="file" id="fileInput" accept=".pdf" multiple onchange="onFilesSelected(this.files)">
  <div id="fileList">ファイルが選択されていません</div>
  <button class="btn-primary" id="startBtn" onclick="startUpload()" disabled>取込開始</button>
  <button class="btn-secondary" onclick="google.script.host.close()">閉じる</button>
  <progress id="progress" max="100" value="0" style="display:none; margin-top:12px;"></progress>
  <div id="status"></div>

<script>
  let selectedFiles = [];

  function onFilesSelected(files) {
    selectedFiles = Array.from(files).filter(f => f.type === 'application/pdf');
    if (selectedFiles.length === 0) {
      document.getElementById('fileList').textContent = 'PDFファイルを選択してください';
      document.getElementById('startBtn').disabled = true;
      return;
    }
    document.getElementById('fileList').innerHTML =
      selectedFiles.map((f, i) =>
        i + 1 + '. ' + f.name + ' (' + (f.size / 1024).toFixed(0) + 'KB)'
      ).join('<br>');
    document.getElementById('startBtn').disabled = false;
  }

  async function startUpload() {
    document.getElementById('startBtn').disabled = true;
    const progress = document.getElementById('progress');
    const status   = document.getElementById('status');
    progress.style.display = 'block';
    progress.value = 0;
    status.innerHTML = '';

    for (let i = 0; i < selectedFiles.length; i++) {
      const file = selectedFiles[i];
      status.innerHTML += '<span class="info">[' + (i+1) + '/' + selectedFiles.length + '] 処理中: ' + file.name + '</span><br>';
      progress.value = Math.round((i / selectedFiles.length) * 100);

      try {
        const base64 = await toBase64(file);
        const result = await new Promise((resolve, reject) => {
          google.script.run
            .withSuccessHandler(resolve)
            .withFailureHandler(reject)
            .uploadAndProcessPDF(base64, file.name);
        });
        status.innerHTML += '<span class="ok">✓ ' + file.name + ' — ' + result + '</span><br>';
      } catch (e) {
        status.innerHTML += '<span class="err">✗ ' + file.name + ' — ' + (e.message || e) + '</span><br>';
      }
    }

    progress.value = 100;
    status.innerHTML += '<br><b>すべての処理が完了しました。</b>';
    document.getElementById('startBtn').disabled = false;
  }

  function toBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => resolve(e.target.result.split(',')[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }
</script>
</body>
</html>
  `)
  .setWidth(480)
  .setHeight(460)
  .setTitle('PDFファイル取込');

  SpreadsheetApp.getUi().showModalDialog(html, 'PDFファイル取込');
}

// ==========================================
// ダイアログから呼ばれる処理
// Base64のPDFを受け取り → Drive保存 → 抽出 → シート記入
// ==========================================
function uploadAndProcessPDF(base64Data, fileName) {
  const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, fileName);
  const folder  = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const pdfFile = folder.createFile(blob);

  try {
    if (isAlreadyProcessed(pdfFile.getId())) {
      pdfFile.setTrashed(true);
      return 'スキップ（処理済み）';
    }
    const data = extractDataFromPDF(pdfFile);
    writeToSpreadsheet(data);
    markAsProcessed(pdfFile.getId(), fileName);
    writeLog(fileName, '成功', '入館者 ' + data.visitors.length + ' 名');

    if (CONFIG.DONE_FOLDER_ID) {
      const doneFolder = DriveApp.getFolderById(CONFIG.DONE_FOLDER_ID);
      doneFolder.addFile(pdfFile);
      folder.removeFile(pdfFile);
    }
    return '完了 — 入館者 ' + data.visitors.length + ' 名を記入しました';

  } catch (e) {
    pdfFile.setTrashed(true);
    writeLog(fileName, 'エラー', e.message);
    throw new Error(e.message);
  }
}

// ==========================================
// フォルダ内の未処理PDFを一括取込（メニューまたはトリガーから実行）
// ==========================================
function processNewPDFs() {
  const folder   = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const files    = folder.getFilesByType(MimeType.PDF);
  const allFiles = [];

  while (files.hasNext()) {
    const f = files.next();
    if (!isAlreadyProcessed(f.getId())) allFiles.push(f);
  }

  if (allFiles.length === 0) {
    Logger.log('処理対象のPDFはありません。');
    return;
  }

  Logger.log('処理対象: ' + allFiles.length + ' 件');

  let successCount = 0;
  let errorCount   = 0;

  allFiles.forEach((file, index) => {
    Logger.log('[' + (index + 1) + '/' + allFiles.length + '] 処理中: ' + file.getName());
    try {
      const data = extractDataFromPDF(file);
      writeToSpreadsheet(data);
      markAsProcessed(file.getId(), file.getName());

      if (CONFIG.DONE_FOLDER_ID) {
        const doneFolder = DriveApp.getFolderById(CONFIG.DONE_FOLDER_ID);
        doneFolder.addFile(file);
        folder.removeFile(file);
      }
      successCount++;
      writeLog(file.getName(), '成功', '入館者 ' + data.visitors.length + ' 名');
      Logger.log('  → 完了（入館者 ' + data.visitors.length + ' 名）');
    } catch (e) {
      errorCount++;
      writeLog(file.getName(), 'エラー', e.message);
      Logger.log('  → エラー: ' + e.message);
    }

    if (index < allFiles.length - 1) Utilities.sleep(1500);
  });

  Logger.log('完了 — 成功: ' + successCount + ' 件 / エラー: ' + errorCount + ' 件');
}

// ==========================================
// PDFからデータ抽出
// ==========================================
function extractDataFromPDF(pdfFile) {
  const tempDoc = Drive.Files.copy(
    { title: '__temp__' + pdfFile.getId(), mimeType: MimeType.GOOGLE_DOCS },
    pdfFile.getId(),
    { convert: true }
  );

  Utilities.sleep(2000);

  let text = '';
  try {
    text = DocumentApp.openById(tempDoc.id).getBody().getText();
  } finally {
    DriveApp.getFileById(tempDoc.id).setTrashed(true);
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
    { pattern: /yes\s+1[Ff]\s*Common/i,        label: '1F共用部' },
    { pattern: /yes\s+2[Ff]\s*Common/i,        label: '2F共用部' },
    { pattern: /yes\s+4[Ff]\s*Common/i,        label: '4F共用部' },
    { pattern: /yes\s+2[Ff]\s*Office/i,        label: '2Fオフィス201' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?101/i, label: '1F会議室101' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?102/i, label: '1F会議室102' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?103/i, label: '1F会議室103' },
  ];

  const found = definitions.filter(d => d.pattern.test(text)).map(d => d.label);
  if (found.length > 0) return found.join(', ');

  const m = text.match(/入室箇所[\s\S]*?(?=搬出入作業)/);
  if (m) {
    return m[0]
      .replace(/入室箇所[^\n]*/, '')
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
  const sectionMatch = text.match(/入館者情報[\s\S]*?(?=申請状況|$)/);
  if (!sectionMatch) return visitors;

  const lines = sectionMatch[0]
    .split('\n')
    .map(l => l.trim())
    .filter(l => l && !l.startsWith('入館者') && !l.startsWith('Visitor') && !l.startsWith('取り込み'));

  const phoneRegex = /^[\d\+][\d\-\+\s]{7,}$/;

  for (let i = 0; i < lines.length; i++) {
    if (phoneRegex.test(lines[i].replace(/\s/g, ''))) {
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

  const appSheet = getOrCreateSheet(ss, CONFIG.SHEET_APPLICATIONS, [
    'ファイル名', '申請者会社名', '申請者氏名', '申請者電話番号',
    '入館開始日', '入館終了日', '入館予定時刻', '退館予定時刻',
    '連続断続区分', 'オーソライズドカード', '入館目的', '入室箇所',
    '搬出入作業', '火気危険物', '申請状況', '処理日時',
  ]);
  appSheet.appendRow([
    data.fileName,      data.company,       data.name,          data.phone,
    data.startDate,     data.endDate,       data.entryTime,     data.exitTime,
    data.entryType,     data.authorizedCard, data.purpose,      data.accessArea,
    data.loadingWork,   data.hazardous,     data.status,        data.processedAt,
  ]);

  const visitorSheet = getOrCreateSheet(ss, CONFIG.SHEET_VISITORS, [
    'ファイル名', '申請者会社名', '入館開始日', '入館者会社名', '入館者氏名', '入館者電話番号',
  ]);
  data.visitors.forEach(v => {
    visitorSheet.appendRow([data.fileName, data.company, data.startDate, v.company, v.name, v.phone]);
  });
}

// ==========================================
// ログシートへの書き込み
// ==========================================
function writeLog(fileName, status, message) {
  const ss       = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const logSheet = getOrCreateSheet(ss, CONFIG.SHEET_LOG, ['処理日時', 'ファイル名', '状況', 'メッセージ']);
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
// 処理済み管理（同じPDFの二重登録を防ぐ）
// ==========================================
function isAlreadyProcessed(fileId) {
  return PropertiesService.getScriptProperties().getProperty('processed_' + fileId) !== null;
}

function markAsProcessed(fileId, fileName) {
  PropertiesService.getScriptProperties()
    .setProperty('processed_' + fileId, new Date().toISOString() + ' | ' + fileName);
}

// 処理済み記録をリセット（再処理したい場合にメニューから実行）
function resetProcessedHistory() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('処理済み履歴をリセットしました。');
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
