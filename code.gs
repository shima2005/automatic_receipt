// 列のインデックス設定 (A=1, B=2, C=3, D=4...)
const URL_COL_INDEX = 4;      // D列: レシート画像URL
const COL_COUNT_FORM_DATA = 4; // A〜D列までがフォームの基本データ

/**
 * 【自動実行】フォーム送信時に起動するトリガー関数
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  
  // フォームから回答された基本データ (A:タイムスタンプ, B:名前, C:区分, D:URL)
  const baseData = [
    e.values[0], 
    e.values[1], 
    e.values[2], 
    e.values[3]  
  ];
  
  coreReceiptProcessor(sheet, row, baseData);
}

/**
 * 【手動実行】スプレッドシートのメニューから選択した行を再処理する
 */
function retrySelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  if (row <= 1) {
    ui.alert('データが入っている行（2行目以降）を選択してください。');
    return;
  }
  
  const baseData = sheet.getRange(row, 1, 1, COL_COUNT_FORM_DATA).getValues()[0];
  
  if (!baseData[3] || baseData[3].toString().trim() === '') {
    ui.alert('選択した行のD列にURLが見つかりません。');
    return;
  }

  // 既にデータがある場合の警告
  const existingCheck = sheet.getRange(row, 5).getValue();
  if (existingCheck) {
    const res = ui.alert('確認', '既にデータが存在します。上書きして再処理しますか？', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
  }

  ss.toast('Gemini APIを呼び出し中です。完了までお待ちください...', '処理開始', 10);
  
  try {
    coreReceiptProcessor(sheet, row, baseData);
    ss.toast('解析が完了し、シートを更新しました。', '完了');
  } catch (error) {
    ui.alert('エラーが発生しました', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * スプレッドシート起動時にメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧾 レシート処理')
    .addItem('選択した行を再解析する', 'retrySelectedRow')
    .addToUi();
}

/**
 * Google Driveのファイル取得でエラーが出た場合にリトライする関数
 */
function getFileWithRetry(fileId) {
  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    try {
      return DriveApp.getFileById(fileId);
    } catch (e) {
      if (i === maxRetries - 1) throw e; // 3回ダメならエラーを投げる
      Logger.log(`ドライブエラーのためリトライ中... (${i + 1}回目)`);
      Utilities.sleep((i + 1) * 1000); // 1秒、2秒と待機時間を増やす
    }
  }
}

// ==========================================
// コアロジック（解析・行操作の共通処理）
// ==========================================

function coreReceiptProcessor(sheet, startRow, baseData) {
  // LockServiceで同時実行による行の重複書き込みを防止
  const lock = LockService.getScriptLock();
  try {
    // 最大30秒間順番待ち
    lock.waitLock(30000); 

    const urlString = baseData[3].toString();
    const isAdvance = (baseData[2] === "立て替え");
    
    // URLを分割してファイルIDを取得
    const urls = urlString.split(',').map(url => url.trim()).filter(url => url.length > 0);
    const fileIds = urls.map(url => {
      const match = url.match(/id=([a-zA-Z0-9_-]+)/);
      return match ? match[1] : null;
    }).filter(id => id !== null);

    if (fileIds.length === 0) throw new Error("有効なGoogleドライブのURLが見つかりませんでした。");

    // API呼び出し（503対策のリトライ機能付き）
    const resultsArray = fetchGeminiWithRetry(fileIds);

    let currentRow = startRow;
    let isFirstWrite = true;

    // 解析結果を1件ずつ行に書き込む
    for (const res of resultsArray) {
      // インデックスを元に対応するURLを1つだけ特定
      const specificUrl = (res.image_index !== undefined && urls[res.image_index]) 
                          ? urls[res.image_index] 
                          : urls[0];

      const bookkeepingData = [
        res.date,     // E列: 日付
        "",           // F列: 空
        "",           // G列: 空
        res.category, // H列: 品目名
        baseData[1],  // I列: 名前
        isAdvance,    // J列: 立替
        false,           // K列: 空
        res.amount    // L列: 金額
      ];

      if (isFirstWrite) {
        // 1件目は元の行を更新 (D列を単体URLにし、E-Lを埋める)
        sheet.getRange(currentRow, URL_COL_INDEX, 1, 9).setValues([[specificUrl, ...bookkeepingData]]);
        sheet.getRange(currentRow, 11).insertCheckboxes();
        isFirstWrite = false;
      } else {
        // 2件目以降は行を挿入してコピー
        sheet.insertRowAfter(currentRow);
        currentRow++;
        const newRow = [
          baseData[0], // A: タイムスタンプ
          baseData[1], // B: 名前
          baseData[2], // C: 区分
          specificUrl, // D: 単体URL
          ...bookkeepingData // E-L: 解析結果
        ];
        sheet.getRange(currentRow, 1, 1, newRow.length).setValues([newRow]);
        sheet.getRange(currentRow, 11).insertCheckboxes();
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// API呼び出し部
// ==========================================

/**
 * 503エラー等の際、最大3回までリトライを行うAPI実行関数
 */
function fetchGeminiWithRetry(fileIds) {
  let lastError;
  const maxRetries = 3;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      return callGeminiAPI(fileIds);
    } catch (e) {
      lastError = e;
      // 503エラーやサーバー負荷の場合、数秒待機してリトライ
      if (e.message.includes('503') || e.message.includes('overloaded')) {
        Logger.log(`Gemini APIが混雑しています。リトライ中... (${i + 1}/${maxRetries})`);
        Utilities.sleep((i + 1) * 3000);
        continue;
      }
      throw e; // 503以外の致命的なエラーは即座に投げる
    }
  }
  throw new Error(`Gemini APIが混雑しており、${maxRetries}回試行しましたが失敗しました。時間をおいて再試行してください。内容: ${lastError}`);
}

/**
 * 実際のAPIリクエスト処理
 */
function callGeminiAPI(fileIds) {
  const parts = [];

  const prompt = `あなたはレシート専門AIです。
各画像データの直前に「image_index: 数字」というテキストを配置しました。
指示に従い、各画像から情報を抽出してJSONで回答してください。

# 抽出ルール
1. 'amount': 合計金額（数値）
2. 'date': 日付（YYYY-MM-DD）
3. 'category': 品目名（ガソリン代、駐車場代、レンタカー代、フェリー代、食品代（原材料名など）、飲み物代など）
4. 'image_index': 各画像に付随する「image_index」の数値をそのまま使用してください。

# 出力形式
JSON配列形式のみを出力してください。説明やMarkdown記法は一切不要です。
例：[{"image_index": 0, "category": "ガソリン代", "amount": 3500, "date": "2026-04-28"}]`;

  parts.push({ text: prompt });

  for (let i = 0; i < fileIds.length; i++) {
    parts.push({ text: "image_index: " + i });
    const file = getFileWithRetry(fileIds[i]);
    const blob = file.getBlob();
    parts.push({
      inlineData: {
        mimeType: blob.getContentType(),
        data: Utilities.base64Encode(blob.getBytes())
      }
    });
  }

  const payload = {
    contents: [{ parts: parts }],
    generationConfig: { 
      responseMimeType: "application/json",
      temperature: 0.1 
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`Gemini API呼び出しでエラーが発生しました。コード: ${responseCode} / メッセージ: ${responseText}`);
  }

  const data = JSON.parse(responseText);
  
  if (!data.candidates || !data.candidates[0].content) {
    throw new Error("APIから有効な回答が得られませんでした。画像が読み取れない可能性があります。");
  }

  let jsonString = data.candidates[0].content.parts[0].text;
  
  jsonString = jsonString.trim().replace(/^```json/g, '').replace(/```$/g, '');

  const results = JSON.parse(jsonString);
  
  // 型の補正
  results.forEach(r => {
    if (r.amount) r.amount = parseFloat(r.amount);
    if (r.image_index !== undefined) r.image_index = parseInt(r.image_index);
  });
  
  return results;
}
