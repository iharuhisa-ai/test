/**
 * 問い合わせメール自動集計スクリプト（AI解析付き）
 *
 * 【セットアップ手順】
 * 1. Google Apps Script (script.google.com) で新規プロジェクトを作成
 * 2. このファイルの内容を貼り付ける
 * 3. CONFIG の各値を自分の環境に合わせて変更
 * 4. スクリプトプロパティに GEMINI_API_KEY を設定（sk-... で始まるキー）
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
  GEMINI_MODEL: 'gemini-1.5-flash',
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
  '最終返信日時',
  '返信件数',
  '最新返信内容',
];

// 列インデックス（0始まり）
const COL = {
  RECEIVED:      0,
  SUBJECT:       1,
  FROM_NAME:     2,
  FROM_EMAIL:    3,
  COMPANY:       4,
  PERSON:        5,
  PHONE:         6,
  TYPE:          7,
  SUMMARY:       8,
  URGENCY:       9,
  BODY:          10,
  THREAD_ID:     11,
  STATUS:        12,
  ASSIGNEE:      13,
  MEMO:          14,
  LAST_REPLY_AT: 15,
  REPLY_COUNT:   16,
  LAST_REPLY:    17,
};

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
 * 新規スレッド → 新しい行を追加
 * 既存スレッドに返信あり → 最終返信日時・返信件数・最新返信内容を更新
 */
function syncInquiryEmails() {
  const sheet = _getSheet();
  const processedIds = _getProcessedIds();
  const apiKey = ScriptApp.getOAuthToken();
  const threadRowMap = _buildThreadRowMap(sheet);


  const query = _buildQuery();
  const threads = GmailApp.search(query, 0, 20);

  if (threads.length === 0) return;

  const newRows = [];
  let updatedCount = 0;

  threads.forEach((thread) => {
    const threadId = thread.getId();
    const messages = thread.getMessages();
    const firstMsg = messages[0];
    const lastMsg = messages[messages.length - 1];
    const replyCount = messages.length;

    // ── 既存スレッドへの返信 ──
    if (threadRowMap[threadId]) {
      const rowNum = threadRowMap[threadId];
      const lastReplyPreview = lastMsg.getPlainBody().replace(/\s+/g, ' ').trim().slice(0, 300);
      sheet.getRange(rowNum, COL.LAST_REPLY_AT + 1).setValue(lastMsg.getDate())
        .setNumberFormat('yyyy/MM/dd HH:mm');
      sheet.getRange(rowNum, COL.REPLY_COUNT + 1).setValue(replyCount);
      sheet.getRange(rowNum, COL.LAST_REPLY + 1).setValue(lastReplyPreview);
      updatedCount++;
      return;
    }

    // ── 新規スレッド ──
    if (processedIds.has(threadId)) return;

    const from = firstMsg.getFrom();
    const emailMatch = from.match(/<(.+?)>/);
    const emailAddress = emailMatch ? emailMatch[1] : from;
    const displayName = emailMatch ? from.replace(/<.+?>/, '').trim() : from;

    const body = firstMsg.getPlainBody().trim();
    const bodyPreview = body.replace(/\s+/g, ' ').slice(0, 300);
    const lastReplyPreview = replyCount > 1
      ? lastMsg.getPlainBody().replace(/\s+/g, ' ').trim().slice(0, 300)
      : '';

    let parsed = { companyName: '', personName: '', phone: '', inquiryType: '', summary: '', urgency: '通常' };
    if (apiKey) {
      try {
        parsed = _analyzeWithOpenAI(apiKey, firstMsg.getSubject(), body);
      } catch (e) {
        Logger.log(`AI解析エラー (threadId: ${threadId}): ${e.message}`);
      }
    }

    newRows.push([
      firstMsg.getDate(),
      firstMsg.getSubject(),
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
      replyCount > 1 ? lastMsg.getDate() : '',
      replyCount,
      lastReplyPreview,
    ]);

    processedIds.add(threadId);
  });

  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, HEADERS.length).setValues(newRows);
    sheet.getRange(startRow, COL.RECEIVED + 1, newRows.length, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    sheet.getRange(startRow, COL.LAST_REPLY_AT + 1, newRows.length, 1).setNumberFormat('yyyy/MM/dd HH:mm');

    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['未対応', '対応中', '完了', '保留'], true)
      .build();
    sheet.getRange(startRow, COL.STATUS + 1, newRows.length, 1).setDataValidation(statusRule);

    for (let i = 0; i < newRows.length; i++) {
      const urgency = newRows[i][COL.URGENCY];
      if (urgency === '高') sheet.getRange(startRow + i, 1, 1, HEADERS.length).setBackground('#fce8e6');
      else if (urgency === '中') sheet.getRange(startRow + i, 1, 1, HEADERS.length).setBackground('#fef9e7');
    }

    _saveProcessedIds(processedIds);
    Logger.log(`新規 ${newRows.length} 件追加、更新 ${updatedCount} 件。`);
  } else {
    Logger.log(`更新 ${updatedCount} 件。`);
  }
}

/**
 * Gemini API でメール本文を解析して構造化情報を返す
 */
function _analyzeWithOpenAI(apiKey, subject, body) {
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

  const url = `https://generativelanguage.googleapis.com/v1/models/${CONFIG.GEMINI_MODEL}:generateContent`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      maxOutputTokens: 512,
    },
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
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

/**
 * 過去メールを一括取り込み
 * 初回セットアップ後に一度だけ実行してください
 * ラベルが付いているすべてのメールをスプレッドシートに転記します
 */
function importAllExistingEmails() {
  const sheet = _getSheet();
  const processedIds = _getProcessedIds();
  const apiKey = ScriptApp.getOAuthToken();

  // 日付制限なしでラベルのメールをすべて取得（500件まで）
  const parts = [];
  if (CONFIG.LABEL_NAME) parts.push(`label:${CONFIG.LABEL_NAME}`);
  if (CONFIG.EXTRA_QUERY) parts.push(CONFIG.EXTRA_QUERY);
  const query = parts.join(' ') || 'in:inbox';

  let start = 0;
  const batchSize = 20; // AI解析があるため1バッチ20件
  let totalImported = 0;

  while (true) {
    const threads = GmailApp.search(query, start, batchSize);
    if (threads.length === 0) break;

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

      let parsed = { companyName: '', personName: '', phone: '', inquiryType: '', summary: '', urgency: '通常' };
      if (apiKey) {
        try {
          parsed = _analyzeWithOpenAI(apiKey, message.getSubject(), body);
        } catch (e) {
          Logger.log(`AI解析エラー (threadId: ${threadId}): ${e.message}`);
        }
      }

      const msgs = thread.getMessages();
      const lastMsg = msgs[msgs.length - 1];
      const replyCount = msgs.length;
      const lastReplyPreview = replyCount > 1
        ? lastMsg.getPlainBody().replace(/\s+/g, ' ').trim().slice(0, 300) : '';

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
        replyCount > 1 ? lastMsg.getDate() : '',
        replyCount,
        lastReplyPreview,
      ]);

      processedIds.add(threadId);
    });

    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, HEADERS.length).setValues(newRows);
      sheet.getRange(startRow, COL.RECEIVED + 1, newRows.length, 1).setNumberFormat('yyyy/MM/dd HH:mm');
      sheet.getRange(startRow, COL.LAST_REPLY_AT + 1, newRows.length, 1).setNumberFormat('yyyy/MM/dd HH:mm');

      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['未対応', '対応中', '完了', '保留'], true)
        .build();
      sheet.getRange(startRow, COL.STATUS + 1, newRows.length, 1).setDataValidation(statusRule);

      for (let i = 0; i < newRows.length; i++) {
        const urgency = newRows[i][COL.URGENCY];
        if (urgency === '高') sheet.getRange(startRow + i, 1, 1, HEADERS.length).setBackground('#fce8e6');
        else if (urgency === '中') sheet.getRange(startRow + i, 1, 1, HEADERS.length).setBackground('#fef9e7');
      }

      totalImported += newRows.length;
      _saveProcessedIds(processedIds);
    }

    // 取得件数がbatchSizeより少なければ最終ページ
    if (threads.length < batchSize) break;
    start += batchSize;

    // GAS の実行時間制限（6分）を考慮して途中で止める
    if (totalImported >= 200) {
      Logger.log(`200件に達したため一時停止。続きは再度 importAllExistingEmails を実行してください。`);
      break;
    }
  }

  Logger.log(`合計 ${totalImported} 件のメールを取り込みました。`);
}



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

    const widths = [150, 250, 130, 200, 150, 120, 130, 130, 350, 80, 300, 120, 100, 120, 250, 150, 80, 300];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

/**
 * シートのスレッドID列を読んで { threadId: rowNumber } のマップを返す
 */
function _buildThreadRowMap(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const threadIds = sheet.getRange(2, COL.THREAD_ID + 1, lastRow - 1, 1).getValues();
  const map = {};
  threadIds.forEach((row, i) => {
    if (row[0]) map[row[0]] = i + 2; // 行番号は1始まり、ヘッダー分+1
  });
  return map;
}

function _getProcessedIds() {
  const raw = PropertiesService.getScriptProperties().getProperty('processedThreadIds');
  return raw ? new Set(JSON.parse(raw)) : new Set();
}

function _saveProcessedIds(idSet) {
  const arr = [...idSet].slice(-5000);
  PropertiesService.getScriptProperties().setProperty('processedThreadIds', JSON.stringify(arr));
}

/**
 * テスト用: サンプルメールをスプレッドシートに書き込む
 * Apps Script エディタから手動で実行してください
 */
function testWithSampleEmail() {
  const sheet = _getSheet();
  const apiKey = ScriptApp.getOAuthToken();

  const subject = '【お伺いしたい件】2026年6月28日 宿泊予約 アテンドサービス / Family Haus 0628';
  const body = `株式会社G.S.P.Corporation
セック様

いつも大変お世話になっております。
Bespoke Japan Travelの宮島と申します。
下記のサービスにつき、ご相談させていただきたくご連絡させていただいております。
現時点で日程を押さえていただく必要はございませんが、空き状況をご教示いただけないでしょうか。
お忙しいところ恐縮ですが、早めにご返信いただけると大変助かります。
何卒よろしくお願い申し上げます。

**********************
●日時：2026年6月28日（日）
●人数: アメリカより大人４名
※英語対応を希望いたします。
●フライト：羽田空港  19:25着 AS831便
●ご依頼内容：空港アテンドサービス
ブリッジでお迎え後、入国審査にご同行いただき、弊社手配のハイヤーまでご案内。
●ご料金：60,500円（税込み、管理手数料10％込）/ お客様1名あたりでお間違いなかったでしょうか
●キャンセルポリシー：業務日より14日前より50％、7日前より100％
●支払い方法：事前銀行振り込み
※ご請求書をメールでお送りいただきたいと存じます。お支払い期限がございましたらご教示ください
**********************

何卒よろしくお願い申し上げます。

*Bespoke Japan Travel 株式会社*
宮島菜実 Nami MIYAJIMA
Tel：050-6883-1811
FAX：050-4561-7326
email：operation@bespokejapantravel.com
Web：https://www.bespokejapantravel.com`;

  let parsed = {
    companyName: 'Bespoke Japan Travel 株式会社',
    personName: '宮島菜実',
    phone: '050-6883-1811',
    inquiryType: '空港アテンドサービス問い合わせ',
    summary: '2026/6/28 羽田空港 19:25着 AS831便、米国人大人4名の空港アテンドサービス依頼。料金・キャンセルポリシーの確認。',
    urgency: '通常',
  };

  if (apiKey) {
    try {
      parsed = _analyzeWithOpenAI(apiKey, subject, body);
      Logger.log('AI解析結果: ' + JSON.stringify(parsed));
    } catch (e) {
      Logger.log('AI解析エラー（フォールバック値を使用）: ' + e.message);
    }
  }

  const insertRow = sheet.getLastRow() + 1;
  const range = sheet.getRange(insertRow, 1, 1, HEADERS.length);
  range.setValues([[
    new Date('2026-05-29T12:13:18+09:00'),
    subject,
    'Bespoke Japan Travel',
    'operation@bespokejapantravel.com',
    parsed.companyName,
    parsed.personName,
    parsed.phone,
    parsed.inquiryType,
    parsed.summary,
    parsed.urgency,
    body.replace(/\s+/g, ' ').slice(0, 300),
    'test-thread-family-haus-0628',
    '未対応',
    '',
    '',
    '',
    1,
    '',
  ]]);

  sheet.getRange(insertRow, 1, 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未対応', '対応中', '完了', '保留'], true)
    .build();
  sheet.getRange(insertRow, 13, 1, 1).setDataValidation(statusRule);

  Logger.log('テストメールをスプレッドシートに書き込みました。');
}

/**
 * Gemini API の動作確認用デバッグ関数
 */
function debugTestGemini() {
  const apiKey = ScriptApp.getOAuthToken();
  Logger.log('✅ OAuthトークン確認: ' + apiKey.slice(0, 8) + '...');

  const url = `https://generativelanguage.googleapis.com/v1/models/${CONFIG.GEMINI_MODEL}:generateContent`;
  const payload = {
    contents: [{ parts: [{ text: '株式会社テスト 山田太郎です。電話: 03-1234-5678\n{"companyName":"","personName":"","phone":"","inquiryType":"","summary":"","urgency":"通常"} の形式でJSONのみ返してください。' }] }],
    generationConfig: { responseMimeType: 'application/json', maxOutputTokens: 256 },
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    const text = response.getContentText();
    Logger.log('HTTPステータス: ' + statusCode);

    if (statusCode !== 200) {
      Logger.log('❌ APIエラー: ' + text.slice(0, 300));
      return;
    }

    const result = JSON.parse(text);
    const parsed = result.candidates[0].content.parts[0].text;
    Logger.log('✅ 解析成功: ' + parsed);

  } catch (e) {
    Logger.log('❌ 例外発生: ' + e.message);
  }
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
