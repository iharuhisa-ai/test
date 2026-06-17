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
// シート振り分けルール
// filenameKeywords: ファイル名で先に判定（会社名より優先）
// keywords        : 会社名で判定
// ==========================================
const SHEET_ROUTING = [
  // ── ファイル名ベース（Tenant申請書）──
  {
    sheet:            'サンライズ',
    filenameKeywords: ['tenant_001', 'tenant001'],
  },
  {
    sheet:            'ABC工事関係者（4F権限者）',
    filenameKeywords: ['tenant_002', 'tenant002'],
  },
  // ── 会社名ベース（通常申請書）──
  {
    sheet:    'EXEO・新菱冷熱・東急建設',
    keywords: ['エクシオ', 'exeo', '新菱', '東急建設', 'デルタ電子', 'delta', 'sec齋藤'],
  },
  {
    sheet:    'PDG',
    keywords: ['pdg', 'ターナー', 'townsend', 'ngk', '鴻池', 'セキュリティーアウトカム'],
  },
  {
    sheet:    'ABC工事関係者（4F権限者）',
    keywords: ['h3c', 'wesco', 'anixter', 'extrreak', 'alibaba'],
  },
  {
    sheet:    '4F ABC',
    keywords: ['zenlayer'],
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
  const fileName = pdfFile.getName();

  // ファイル名からPDF種別を判定
  const isTenant001 = /tenant.?001/i.test(fileName);
  const isTenant002 = /tenant.?002/i.test(fileName);

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

  const accessArea = extractAccessArea(text);

  // Tenant_001: 番号付きリスト形式（会社名・電話番号なし）
  // Tenant_002/通常: 電話番号ベースの解析
  const visitors = isTenant001
    ? extractTenant001Visitors(text)
    : extractVisitors(text, applicantCompany);

  const data = {
    fileName,
    isTenant001,
    isTenant002,
    processedAt:     new Date().toLocaleString('ja-JP'),
    applicantCompany,
    applicantName,
    applicantPhone,
    startDate,
    endDate,
    entryTime,
    exitTime,
    accessArea,
    visitors,
  };

  Logger.log('=== 解析結果 ===');
  Logger.log('種別: ' + (isTenant001 ? 'Tenant001(サンライズ)' : isTenant002 ? 'Tenant002(ABC)' : '通常'));
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
    // Tenant_001 形式（CollaborFlow UI）
    { pattern: /yes\s+DH307/i,               label: 'DH307' },
    { pattern: /yes\s+DH308/i,               label: 'DH308' },
    // 通常申請書形式
    { pattern: /yes\s+1[Ff]\s*Common/i,      label: '1F共用部' },
    { pattern: /yes\s+2[Ff]\s*Common/i,      label: '2F共用部' },
    { pattern: /yes\s+4[Ff]\s*Common/i,      label: '4F共用部' },
    { pattern: /yes\s+2[Ff]\s*Office/i,      label: '2Fオフィス201' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?101/i, label: '1F会議室101' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?102/i, label: '1F会議室102' },
    { pattern: /yes\s+1[Ff]\s*Meeting.*?103/i, label: '1F会議室103' },
    { pattern: /yes\s+1[Ff]\s*UPS/i,         label: '1F UPS室' },
    { pattern: /yes\s+5[Ff]/i,               label: '5F' },
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
// Tenant_001 入館者リスト抽出
// 形式: 番号付きリスト（会社名・電話番号なし）
// ==========================================
function extractTenant001Visitors(text) {
  const visitors = [];

  // 入館者情報セクションを切り出す（なければ全体を対象）
  const sectionMatch = text.match(/(?:入館者情報|Visitor Information)([\s\S]*?)(?=申請状況|承認|$)/i);
  const section = sectionMatch ? sectionMatch[1] : text;

  Logger.log('=== Tenant001 入館者セクション ===\n' + section);

  const lines = section.split('\n');
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    // 「1: NAME」「1. NAME」「1 NAME」「No.1 NAME」など柔軟に対応
    const m = trimmed.match(/^(?:No\.?\s*)?(\d+)\s*[:\.\s]\s*([A-Za-z぀-鿿][^\n]{1,60})$/);
    if (m) {
      const name = m[2].trim();
      if (
        name.length > 1 &&
        !/^(name|company|visitor|phone|no\.|number)/i.test(name) &&
        !/^[\d\s]+$/.test(name)
      ) {
        visitors.push({ company: 'サンライズ', name, phone: '' });
      }
    }
  }

  // セクション形式で取れなかった場合: テキスト全体から番号+名前を探す
  if (visitors.length === 0) {
    const allMatches = [...text.matchAll(/^(\d+)\s*[:\.\s]\s*([A-Z][A-Z\s\-]{2,50})$/gm)];
    for (const m of allMatches) {
      const name = m[2].trim();
      if (!/^(NAME|COMPANY|VISITOR|PHONE)/i.test(name)) {
        visitors.push({ company: 'サンライズ', name, phone: '' });
      }
    }
  }

  return visitors;
}

// ==========================================
// 通常申請書の入館者リスト抽出（電話番号ベース）
// ==========================================
function extractVisitors(text, fallbackCompany) {
  const visitors = [];

  // 入館者情報セクションを切り出す
  const sectionMatch = text.match(/入館者情報[\s\S]*?(?=申請状況|$)/);
  if (!sectionMatch) return [];

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
      !/^\d+$/.test(l)
    );

  const phoneRegex = /^[\d\+][\d\-\+\(\)\s]{7,}$/;

  for (let i = 0; i < lines.length; i++) {
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

    // ファイル名を最優先にシートを決定（Tenant_001/002 判定）
    const sheetName = resolveSheet(
      visitor.company + ' ' + data.applicantCompany,
      data.fileName
    );
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log('シートが見つかりません: ' + sheetName);
      summary.push(visitor.name + ' → シート未発見(' + sheetName + ')');
      return;
    }

    const row = buildRow(sheetName, visitor, period, data.accessArea);
    // 氏名列の最終入力行の次に書き込む
    const nameCol = (sheetName === 'サンライズ') ? 6 : 2;
    const nextRow = getNextRowByColumn(sheet, nameCol);
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    summary.push(visitor.name + ' → ' + sheetName);
  });

  return summary.join(', ');
}

// ==========================================
// 指定列の最終入力行の次の行番号を返す
// ==========================================
function getNextRowByColumn(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 2; // ヘッダーのみの場合は2行目から
  const values = sheet.getRange(1, col, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '') return i + 2; // 次の行
  }
  return 2; // データなしの場合は2行目から
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
// 会社名 or ファイル名からシートを特定
// ==========================================
function resolveSheet(companyText, fileName) {
  // ファイル名ベースを先に評価
  const fn = (fileName || '').toLowerCase();
  for (const rule of SHEET_ROUTING) {
    if (rule.filenameKeywords && rule.filenameKeywords.some(kw => fn.includes(kw))) {
      return rule.sheet;
    }
  }
  // 会社名ベース
  const text = (companyText || '').toLowerCase();
  for (const rule of SHEET_ROUTING) {
    if (rule.keywords && rule.keywords.some(kw => text.includes(kw))) {
      return rule.sheet;
    }
  }
  return 'その他';
}

// ==========================================
// シートごとの列順で行データを作成
// ==========================================
function buildRow(sheetName, visitor, period, accessArea) {
  switch (sheetName) {
    case 'サンライズ':
      // ステータス / MGRチェック / 一人目 / 二人目 / 申請書内番号 / 氏名 / 期間 / ID番号 / 入館時間 / 退館時間
      return ['', '', '', '', '', visitor.name, period, '', '', ''];

    case 'EXEO・新菱冷熱・東急建設':
      // A:会社名 B:氏名 C:電話番号 D:期間 E:ID番号 F:入館時間 G:退館時間 H:カード番号 I:(備考) J:入室箇所
      return [visitor.company, visitor.name, visitor.phone, period, '', '', '', '', '', accessArea];

    case 'PDG':
      // A:会社名 B:氏名 C:電話番号 D:期間 E:ID番号 F:入館時間 G:退館時間 H:カード番号 I:(備考) J:入室箇所
      return [visitor.company, visitor.name, visitor.phone, period, '', '', '', '', '', accessArea];

    case 'ABC工事関係者（4F権限者）':
      // A:会社名 B:氏名 C:電話番号 D:期間 E:ID番号 F:入館時間 G:退館時間 H:カード番号 I:(備考) J:入室箇所
      return [visitor.company, visitor.name, visitor.phone, period, '', '', '', '', '', accessArea];

    case '4F ABC':
      // A:会社名 B:氏名 C:電話番号 D:期間 E:ID番号 F:入館時間 G:退館時間 H:カード番号 I:(備考) J:入室箇所
      return [visitor.company, visitor.name, visitor.phone, period, '', '', '', '', '', accessArea];

    default:
      // その他: A:会社名 B:氏名 C:電話番号 D:期間 E:ID番号 F:入館時間 G:退館時間 H:カード番号 I:(備考) J:入室箇所
      return [visitor.company, visitor.name, visitor.phone, period, '', '', '', '', '', accessArea];
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
