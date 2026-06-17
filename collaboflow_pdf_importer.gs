// ==========================================
// 設定（ここだけ変更してください）
// ==========================================
const CONFIG = {
  // PDFを一時保存するGoogleドライブのフォルダID
  PDF_FOLDER_ID:  'ここにフォルダIDを入力',

  // 書き込み先スプレッドシートID
  SPREADSHEET_ID: 'ここにスプレッドシートIDを入力',

  // 処理ログシート名
  SHEET_LOG: '処理ログ',
};

// ==========================================
// 会社名キーワード → シート振り分けルール
// 新しい会社を追加する場合は keywords に追記してください
// ==========================================
const SHEET_ROUTING = [
  {
    sheet: 'EXEO・新菱冷熱・東急建設',
    keywords: ['エクシオ', 'exeo', '新菱', '東急建設', 'デルタ電子', 'delta', 'SEC齋藤'],
  },
  {
    sheet: 'PDG',
    keywords: ['PDG', 'ターナー', 'Townsend', 'NGK', '鴻池', 'セキュリティーアウトカム'],
  },
  {
    sheet: 'ABC工事関係者（4F権限者）',
    keywords: ['H3C', 'Wesco', 'Anixter', 'eXtrreak', 'Alibaba'],
  },
  {
    sheet: '4F ABC',
    keywords: ['Zenlayer'],
  },
  // 上記に一致しない場合は「その他」へ
];

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
// ファイル選択ダイアログ
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
        (i + 1) + '. ' + f.name + ' (' + (f.size / 1024).toFixed(0) + 'KB)'
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
      status.innerHTML += '<span class="info">[' + (i+1) + '/' + selectedFiles.length + '] 処理中: ' + file.name + '...</span><br>';
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
    const data    = extractDataFromPDF(pdfFile);
    const results = writeToSheets(data);
    markAsProcessed(pdfFile.getId(), fileName);
    writeLog(fileName, '成功', results);
    return results;
  } catch (e) {
    pdfFile.setTrashed(true);
    writeLog(fileName, 'エラー', e.message);
    throw new Error(e.message);
  }
}

// ==========================================
// フォルダ内の未処理PDFを一括取込
// ==========================================
function processNewPDFs() {
  const folder   = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const files    = folder.getFilesByType(MimeType.PDF);
  const allFiles = [];

  while (files.hasNext()) {
    const f = files.next();
    if (!isAlreadyProcessed(f.getId())) allFiles.push(f);
  }

  if (allFiles.length === 0) { Logger.log('処理対象なし'); return; }
  Logger.log('処理対象: ' + allFiles.length + ' 件');

  allFiles.forEach((file, index) => {
    Logger.log('[' + (index + 1) + '/' + allFiles.length + '] ' + file.getName());
    try {
      const data    = extractDataFromPDF(file);
      const results = writeToSheets(data);
      markAsProcessed(file.getId(), file.getName());
      writeLog(file.getName(), '成功', results);
      Logger.log('  → ' + results);
    } catch (e) {
      writeLog(file.getName(), 'エラー', e.message);
      Logger.log('  → エラー: ' + e.message);
    }
    if (index < allFiles.length - 1) Utilities.sleep(1500);
  });
}

// ==========================================
// PDF → Googleドキュメント変換でテキスト抽出（無料・Drive API使用）
// ==========================================
function extractDataFromPDF(pdfFile) {
  // PDFをGoogleドキュメントに変換
  const tempDoc = Drive.Files.copy(
    { title: '__temp__' + pdfFile.getId(), mimeType: MimeType.GOOGLE_DOCS },
    pdfFile.getId(),
    { convert: true }
  );

  Utilities.sleep(2000); // 変換完了待ち

  let text = '';
  try {
    text = DocumentApp.openById(tempDoc.id).getBody().getText();
  } finally {
    DriveApp.getFileById(tempDoc.id).setTrashed(true); // 一時ファイルを必ず削除
  }

  Logger.log('=== 抽出テキスト ===\n' + text + '\n===================');

  // ==========================================
  // 書式の違いに対応した柔軟な抽出
  // ==========================================
  const applicantCompany = extractAny(text, [
    /会社名[\/\*\s]*Company Name\s*[\n\r]+([^\n\r]+)/,
    /会社名[^\n]*\n([^\n]+)/,
    /Company Name[^\n]*\n([^\n]+)/,
  ]);

  const applicantName = extractAny(text, [
    /氏名[\/\*\s]*Name\s*[\n\r]+([^\n\r]+)/,
    /氏名[^\n]*\n([^\n]+)/,
    /\*Name[^\n]*\n([^\n]+)/,
  ]);

  const applicantPhone = extractAny(text, [
    /電話番号[\/\*\s]*Phone Number\s*[\n\r]+([^\n\r]+)/,
    /電話番号[^\n]*\n([^\n]+)/,
    /Phone Number[^\n]*\n([^\n]+)/,
  ]);

  const startDate = extractAny(text, [
    /入館開始日[^\n]*\n(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/,
    /Start Date[^\n]*\n(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/,
    /(\d{4}\/\d{2}\/\d{2})(?=.*入館開始|.*Start)/,
  ]);

  const endDate = extractAny(text, [
    /入館終了日[^\n]*\n(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/,
    /End Date[^\n]*\n(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/,
  ]);

  const entryTime = extractAny(text, [
    /入館予定時刻[^\n]*\n[^\n]*\n(\d{1,2}:\d{2})/,
    /Scheduled Entry Time[^\n]*\n[^\n]*\n(\d{1,2}:\d{2})/,
    /入館予定時刻[^\n]*\n(\d{1,2}:\d{2})/,
  ]);

  const exitTime = extractAny(text, [
    /退館予定時刻[^\n]*\n[^\n]*\n(\d{1,2}:\d{2})/,
    /Scheduled Exit Time[^\n]*\n[^\n]*\n(\d{1,2}:\d{2})/,
    /退館予定時刻[^\n]*\n(\d{1,2}:\d{2})/,
  ]);

  const entryType = extractAny(text, [
    /(連続入館|断続入館|両方)/,
    /(Continuous Entry|Intermittent Entry|Both)/,
  ]);

  const purpose = extractAny(text, [
    /入館目的[^\n]*\n([^\n]+)/,
    /Purpose of Entry[^\n]*\n([^\n]+)/,
  ]);

  const accessArea = extractAccessArea(text);

  const visitors = extractVisitors(text, applicantCompany);

  const data = {
    fileName:        pdfFile.getName(),
    processedAt:     new Date().toLocaleString('ja-JP'),
    applicantCompany,
    applicantName,
    applicantPhone,
    startDate,
    endDate,
    entryTime,
    exitTime,
    entryType,
    purpose,
    accessArea,
    visitors,
  };

  Logger.log('=== 解析結果 ===');
  Logger.log('申請者: ' + applicantCompany + ' / ' + applicantName);
  Logger.log('期間: ' + startDate + '～' + endDate + ' ' + entryTime + '～' + exitTime);
  Logger.log('入室箇所: ' + accessArea);
  Logger.log('入館者数: ' + visitors.length);
  visitors.forEach((v, i) => Logger.log('  ' + (i+1) + '. ' + v.company + ' / ' + v.name + ' / ' + v.phone));

  return data;
}

// ==========================================
// 複数パターンで順に試して最初にマッチした値を返す
// ==========================================
function extractAny(text, patterns) {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1] && match[1].trim()) {
      return match[1].trim();
    }
  }
  return '';
}

// ==========================================
// 入室箇所の抽出（yes/Yes チェックボックス形式 + テキスト形式に対応）
// ==========================================
function extractAccessArea(text) {
  const checkPatterns = [
    { pattern: /yes\s+1[Ff]\s*Common/i,        label: '1F共用部' },
    { pattern: /yes\s+2[Ff]\s*Common/i,        label: '2F共用部' },
    { pattern: /yes\s+4[Ff]\s*Common/i,        label: '4F共用部' },
    { pattern: /yes\s+2[Ff]\s*Office/i,        label: '2Fオフィス201' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?101/i, label: '1F会議室101' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?102/i, label: '1F会議室102' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?103/i, label: '1F会議室103' },
    { pattern: /yes\s+1[Ff]\s*UPS/i,           label: '1F UPS室' },
    { pattern: /yes\s+5[Ff]/i,                 label: '5F' },
  ];

  const found = checkPatterns.filter(d => d.pattern.test(text)).map(d => d.label);
  if (found.length > 0) return found.join(', ');

  // チェックボックス形式でない場合はテキストから抽出
  const m = text.match(/入室箇所[^\n]*\n([\s\S]*?)(?=搬出入作業|Loading and Unloading|$)/);
  if (m) {
    return m[1]
      .split('\n')
      .map(l => l.trim())
      .filter(l => l && !/^yes$/i.test(l) && !/^no$/i.test(l) && !/^\*?Access Area/i.test(l))
      .join(', ')
      .trim();
  }
  return '';
}

// ==========================================
// 入館者リストの抽出（複数行テーブル形式に対応）
// ==========================================
function extractVisitors(text, fallbackCompany) {
  const visitors = [];

  // 入館者情報セクションを切り出す
  const sectionMatch = text.match(/入館者情報[\s\S]*?(?=申請状況|$)/);
  if (!sectionMatch) {
    // セクションが見つからない場合は申請者本人を入館者とする
    return [];
  }

  const lines = sectionMatch[0]
    .split('\n')
    .map(l => l.trim())
    .filter(l =>
      l &&
      !l.startsWith('入館者情報') &&
      !l.startsWith('Visitor Information') &&
      !l.startsWith('入館者会社名') &&
      !l.startsWith('Visitor Company') &&
      !l.startsWith('取り込み') &&
      !l.startsWith('Download') &&
      !/^\d+$/.test(l)          // 行番号のみの行を除外
    );

  const phoneRegex = /^[\d\+][\d\-\+\(\)\s]{7,}$/;

  for (let i = 0; i < lines.length; i++) {
    const cleaned = lines[i].replace(/[\s\-\(\)]/g, '');
    if (phoneRegex.test(lines[i].replace(/\s/g, ''))) {
      const phone       = lines[i].trim();
      const visitorName = i >= 1 ? lines[i - 1].trim() : '';
      const company     = i >= 2 ? lines[i - 2].trim() : (fallbackCompany || '');

      if (visitorName && !visitorName.includes('Download') && !visitorName.match(/^\d+$/)) {
        visitors.push({
          company: company || fallbackCompany || '',
          name:    visitorName,
          phone:   phone,
        });
      }
    }
  }

  return visitors;
}

// ==========================================
// シートへの振り分けと書き込み
// ==========================================
function writeToSheets(data) {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const period  = formatPeriod(data);
  const summary = [];

  // 入館者リストが空なら申請者本人を使う
  const visitors = (data.visitors && data.visitors.length > 0)
    ? data.visitors
    : [{ company: data.applicantCompany, name: data.applicantName, phone: data.applicantPhone }];

  visitors.forEach(visitor => {
    if (!visitor.company) visitor.company = data.applicantCompany;
    if (!visitor.phone)   visitor.phone   = '';

    // 申請者会社名も加味してシートを決定
    const sheetName = resolveSheet(visitor.company + ' ' + data.applicantCompany);
    const sheet     = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log('シートが見つかりません: ' + sheetName);
      summary.push(visitor.name + ' → シート未発見(' + sheetName + ')');
      return;
    }

    const row = buildRow(sheetName, visitor, period, data.accessArea);
    sheet.appendRow(row);
    summary.push(visitor.name + ' → ' + sheetName);
  });

  return summary.join(', ');
}

// ==========================================
// 期間文字列の整形
// ==========================================
function formatPeriod(data) {
  let str = '';
  if (data.startDate) str += data.startDate;
  if (data.endDate)   str += '～' + data.endDate;
  if (data.entryTime && data.exitTime) {
    str += '\n' + data.entryTime + '～' + data.exitTime;
  }
  return str;
}

// ==========================================
// 会社名からシートを特定
// ==========================================
function resolveSheet(companyText) {
  const text = (companyText || '').toLowerCase();
  for (const rule of SHEET_ROUTING) {
    if (rule.keywords.some(kw => text.includes(kw.toLowerCase()))) {
      return rule.sheet;
    }
  }
  return 'その他';
}

// ==========================================
// シートごとの列順で行データを作成
// ==========================================
function buildRow(sheetName, visitor, period, accessArea) {
  // 電話番号列があるシート
  const withPhone = ['EXEO・新菱冷熱・東急建設', 'ABC工事関係者（4F権限者）', 'その他'];

  if (withPhone.includes(sheetName)) {
    // 会社名 / 氏名 / 電話番号 / 期間 / ID番号 / 入館時間 / 退館時間 / カード番号 / 入室箇所
    return [visitor.company, visitor.name, visitor.phone, period, '', '', '', '', accessArea];
  } else {
    // 会社名 / 氏名 / 期間 / ID番号 / 入館時間 / 退館時間 / カード番号 / 入室箇所
    return [visitor.company, visitor.name, period, '', '', '', '', accessArea];
  }
}

// ==========================================
// ログシートへの書き込み
// ==========================================
function writeLog(fileName, status, message) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let logSheet = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEET_LOG);
    logSheet.appendRow(['処理日時', 'ファイル名', '状況', 'メッセージ']);
    logSheet.getRange(1, 1, 1, 4).setBackground('#4472C4').setFontColor('white').setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }
  logSheet.appendRow([new Date().toLocaleString('ja-JP'), fileName, status, message]);
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

function resetProcessedHistory() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('処理済み履歴をリセットしました。');
}

// ==========================================
// 自動実行トリガー設定（必要な場合のみ実行）
// ==========================================
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processNewPDFs').timeBased().everyMinutes(5).create();
  Logger.log('トリガー設定完了（5分ごとに自動実行）');
}
