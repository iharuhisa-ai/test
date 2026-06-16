/**
 * 問い合わせメール自動集計スクリプト（AI解析付き）
 *
 * 【セットアップ手順】
 * 1. Google Apps Script (script.google.com) で新規プロジェクトを作成
 * 2. このファイルの内容を貼り付ける
 * 3. CONFIG の各値を自分の環境に合わせて変更
 * 4. スクリプトプロパティに GEMINI_API_KEY を設定
 *    （エディタ上部: プロジェクトの設定 → スクリプトプロパティ → 追加）
 * 5. 初回は setup() を手動実行してシートとトリガーを作成する
 */

// ===== 設定項目 =====
const CONFIG = {
  // 転記先スプレッドシートのID（URLの /d/XXXX/ の部分）
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',

  // 監視する Gmail ラベル名（例: "問い合わせ"）
  // 空文字 '' にすると全受信トレイが対象
  LABEL_NAME: '問い合わせ',

  // 転記先シート名
  SHEET_NAME: '問い合わせ一覧',

  // 追加の検索クエリ（例: 'subject:お問い合わせ'）
  EXTRA_QUERY: '',

  // Gemini API モデル
  GEMINI_MODEL: 'gemini-2.0-flash',
};

// ===== ヘッダー定義 =====
const HEADERS = [
  '受信日時',
  '件名',
  '差出人名',
  '差出人メールアドレス',
  '会社名',
  '氏名',
  '電話番号',
  '問い合わせ種別',
  '問い合わせ内容（AI要約）',
  '緊急度',
  '本文（原文冒頭）',
  'スレッドID',
  'ステータス',
  '担当者',
  '対応メモ',
];

/**
 * 初回セットアップ: シート作成 + トリガー登録
 * Apps Script エディタから手動で一度だけ実行してください
 */
function setup() {
  _ensureSheet();
  _registerTrigger();
  Logger.log('セットアップ完了。トリガーを登録しました。');
}

/**
 * トリガーから呼び出されるメイン関数
 */
function syncInquiryEmails() {
  const sheet = _getSheet();
  const processedIds = _getProcessedIds();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  const query = _buildQuery();
  const threads = GmailApp.search(query, 0, 20); // AI解析があるので1回あたり20件

  if (threads.length === 0) return;

  const newRows = [];

  threads.forEach((thread) => {
    const threadId = thread.getId();
    if (processedIds.has(threadId)) return;

    const message = thread.getMessages()[0];
    const from = message.getFrom();
    const emailMatch = from.match(/<(.+?)>/);
    const emailAddress = emailMatch ? emailMatch[1] : from;
    const displayName = emailMatch ? from.replace(/<.+?>/, '').trim() : from;

    const body = message.getPlainBody().trim();
    const bodyPreview = body.replace(/\s+/g, ' ').slice(0, 300);

    // AI解析
    let parsed = {
      companyName: '',
      personName: '',
      phone: '',
      inquiryType: '',
      summary: '',
      urgency: '通常',
    };

    if (apiKey) {
      try {
        parsed = _analyzeWithGemini(apiKey, message.getSubject(), body);
      } catch (e) {
        Logger.log(`AI解析エラー (threadId: ${threadId}): ${e.message}`);
      }
    }

    newRows.push([
      message.getDate(),
      message.getSubject(),
      displayName,
      emailAddress,
      parsed.companyName,
      parsed.personName,
      parsed.phone,
      parsed.inquiryType,
      parsed.summary,
      parsed.urgency,
      bodyPreview,
      threadId,
      '未対応',
      '',
      '',
    ]);

    processedIds.add(threadId);
  });

  if (newRows.length > 0) {
    const insertRow = 2;
    sheet.insertRowsBefore(insertRow, newRows.length);
    const range = sheet.getRange(insertRow, 1, newRows.length, HEADERS.length);
    range.setValues(newRows);

    sheet.getRange(insertRow, 1, newRows.length, 1)
      .setNumberFormat('yyyy/MM/dd HH:mm');

    // ステータスのプルダウン
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['未対応', '対応中', '完了', '保留'], true)
      .build();
    sheet.getRange(insertRow, 13, newRows.length, 1).setDataValidation(statusRule);

    // 緊急度による色付け
    for (let i = 0; i < newRows.length; i++) {
      const urgency = newRows[i][9];
      if (urgency === '高') {
        sheet.getRange(insertRow + i, 1, 1, HEADERS.length).setBackground('#fce8e6');
      } else if (urgency === '中') {
        sheet.getRange(insertRow + i, 1, 1, HEADERS.length).setBackground('#fef9e7');
      }
    }

    _saveProcessedIds(processedIds);
    Logger.log(`${newRows.length} 件のメールを転記しました。`);
  }
}

/**
 * Gemini API でメール本文を解析して構造化情報を返す
 */
function _analyzeWithGemini(apiKey, subject, body) {
  const prompt = `以下のメール件名と本文を解析して、JSON形式で情報を抽出してください。

件名: ${subject}

本文:
${body.slice(0, 2000)}

以下のJSON形式で返してください（余分なテキストなし、JSONのみ）:
{
  "companyName": "会社名（不明の場合は空文字）",
  "personName": "氏名（不明の場合は空文字）",
  "phone": "電話番号（不明の場合は空文字）",
  "inquiryType": "問い合わせ種別（例: 製品問い合わせ、見積もり依頼、クレーム、採用、その他）",
  "summary": "問い合わせ内容の要約（100文字以内）",
  "urgency": "緊急度（高・中・通常のいずれか）"
}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      maxOutputTokens: 512,
    },
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error(result.error.message);
  }

  const text = result.candidates[0].content.parts[0].text.trim();
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  return jsonMatch ? JSON.parse(jsonMatch[0]) : {};
}

// ===== 内部ヘルパー関数 =====

function _buildQuery() {
  const parts = [];
  if (CONFIG.LABEL_NAME) parts.push(`label:${CONFIG.LABEL_NAME}`);
  if (CONFIG.EXTRA_QUERY) parts.push(CONFIG.EXTRA_QUERY);
  parts.push('newer_than:7d');
  return parts.join(' ');
}

function _getSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = _ensureSheet();
  return sheet;
}

function _ensureSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    const widths = [150, 250, 130, 200, 150, 120, 130, 130, 350, 80, 300, 120, 100, 120, 250];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

function _getProcessedIds() {
  const raw = PropertiesService.getScriptProperties().getProperty('processedThreadIds');
  return raw ? new Set(JSON.parse(raw)) : new Set();
}

function _saveProcessedIds(idSet) {
  const arr = [...idSet].slice(-5000);
  PropertiesService.getScriptProperties().setProperty('processedThreadIds', JSON.stringify(arr));
}

function _registerTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'syncInquiryEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 5分おきのポーリング（Gmail のメール受信トリガーのバックアップ）
  ScriptApp.newTrigger('syncInquiryEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
}
