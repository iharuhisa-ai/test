// ==========================================
// 設定（ここだけ変更してください）
// ==========================================
const CONFIG = {
  // PDFを一時保存するGoogleドライブのフォルダID
  PDF_FOLDER_ID: 'ここにフォルダIDを入力',

  // 書き込み先スプレッドシートID（送付のExcelをGoogleスプレッドシートに変換したもの）
  SPREADSHEET_ID: 'ここにスプレッドシートIDを入力',

  // Gemini API キー（Google AI Studio で取得: https://aistudio.google.com/apikey）
  GEMINI_API_KEY: 'ここにAPIキーを入力',

  // 処理ログシート名
  SHEET_LOG: '処理ログ',
};

// ==========================================
// 会社名キーワード → シート振り分けルール
// ==========================================
const SHEET_ROUTING = [
  {
    sheet: 'EXEO・新菱冷熱・東急建設',
    keywords: ['エクシオ', 'exeo', 'EXEO', '新菱', '東急建設', 'デルタ電子', 'delta', 'SEC齋藤'],
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
    keywords: ['Zenlayer', 'ABC'],
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
    const data    = extractDataWithAI(pdfFile);
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
      const data    = extractDataWithAI(file);
      const results = writeToSheets(data);
      markAsProcessed(file.getId(), file.getName());
      writeLog(file.getName(), '成功', results);
      Logger.log('  → ' + results);
    } catch (e) {
      writeLog(file.getName(), 'エラー', e.message);
      Logger.log('  → エラー: ' + e.message);
    }
    if (index < allFiles.length - 1) Utilities.sleep(2000);
  });
}

// ==========================================
// Gemini API で PDF からデータ抽出
// ==========================================
function extractDataWithAI(pdfFile) {
  const base64Pdf = Utilities.base64Encode(pdfFile.getBlob().getBytes());

  const prompt = `
以下のTY1入館申請書PDFを解析して、JSON形式で情報を抽出してください。

抽出するJSON形式:
{
  "applicant": {
    "company": "申請者の会社名",
    "name": "申請者の氏名",
    "phone": "申請者の電話番号"
  },
  "period": {
    "startDate": "入館開始日 (YYYY/MM/DD形式)",
    "endDate": "入館終了日 (YYYY/MM/DD形式)",
    "entryTime": "入館予定時刻 (HH:MM形式)",
    "exitTime": "退館予定時刻 (HH:MM形式)"
  },
  "accessArea": "入室箇所（すべて記載）",
  "purpose": "入館目的",
  "entryType": "連続入館 or 断続入館 or 両方",
  "visitors": [
    {
      "company": "入館者の会社名",
      "name": "入館者の氏名",
      "phone": "入館者の電話番号"
    }
  ]
}

注意:
- 入館者情報テーブルに複数人いる場合はすべてvisitorsに含める
- 入館者会社名が空欄の場合は申請者会社名を使う
- JSONのみを返し、説明文は不要
`;

  const url = 'https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash-latest:generateContent?key=' + CONFIG.GEMINI_API_KEY;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: 'application/pdf', data: base64Pdf } },
      ],
    }],
    generationConfig: { temperature: 0 },
  };

  // リトライ付きAPIコール（最大3回、429エラー時は30秒待機）
  let response;
  for (let attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() === 200) break;

    const errBody = JSON.parse(response.getContentText());
    if (errBody.error && errBody.error.code === 429) {
      Logger.log('レート制限のため ' + (attempt * 30) + '秒待機... (' + attempt + '/3)');
      Utilities.sleep(attempt * 30 * 1000);
    } else {
      break;
    }
  }

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini API エラー: ' + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  let jsonText = result.candidates[0].content.parts[0].text;
  Logger.log('AI抽出結果:\n' + jsonText);

  // マークダウンのコードブロックを除去
  jsonText = jsonText.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();

  const data       = JSON.parse(jsonText);
  data.fileName    = pdfFile.getName();
  data.processedAt = new Date().toLocaleString('ja-JP');
  return data;
}

// ==========================================
// シートへの振り分けと書き込み
// ==========================================
function writeToSheets(data) {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const period  = formatPeriod(data.period);
  const summary = [];

  // 入館者リストが空なら申請者本人を使う
  const visitors = (data.visitors && data.visitors.length > 0)
    ? data.visitors
    : [{ company: data.applicant.company, name: data.applicant.name, phone: data.applicant.phone }];

  visitors.forEach(visitor => {
    if (!visitor.company) visitor.company = data.applicant.company;
    if (!visitor.phone)   visitor.phone   = '';

    const sheetName = resolveSheet(visitor.company + ' ' + data.applicant.company);
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
function formatPeriod(period) {
  if (!period) return '';
  let str = (period.startDate || '') + '～' + (period.endDate || '');
  if (period.entryTime && period.exitTime) {
    str += '\n' + period.entryTime + '～' + period.exitTime;
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
// 処理済み管理
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
// 使用可能なGeminiモデルを確認（初回セットアップ時に実行してモデル名を確認）
// ==========================================
function listModels() {
  const url = 'https://generativelanguage.googleapis.com/v1/models?key=' + CONFIG.GEMINI_API_KEY;
  const res = UrlFetchApp.fetch(url);
  const models = JSON.parse(res.getContentText()).models;
  models
    .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes('generateContent'))
    .forEach(m => Logger.log(m.name));
}

// ==========================================
// 自動実行トリガー設定（必要な場合のみ実行）
// ==========================================
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processNewPDFs').timeBased().everyMinutes(5).create();
  Logger.log('トリガー設定完了（5分ごとに自動実行）');
}
