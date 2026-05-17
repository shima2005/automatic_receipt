/**
 * 三菱ＵＦＪ‐ＶＩＳＡビジネスデビット利用通知メールから情報を抽出し
 * スプレッドシートに記録するスクリプト
 */

function extractMufgDebitInfo() {
  const SPREADSHEET_ID = null; 
  const SHEET_NAME = '利用明細';
  const SUMMARY_SHEET_NAME = '月別集計';
  const LABEL_NAME = '処理済み';
  
  const SEARCH_QUERY = '三菱ＵＦＪ ビジネスデビット after:2025/09/01';
  
  try {
    const spreadsheet = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      sheet.appendRow(['受信日時', '利用先', '利用金額（円）', '利用者', 'チェック済', 'MessageID']);
      sheet.setFrozenRows(1);
      // MessageID列（F列）は管理用なので非表示にしても良い
      sheet.hideColumns(6);
    }

    let label = GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);

    const threads = GmailApp.search(SEARCH_QUERY + ' -label:' + LABEL_NAME, 0, 100);
    console.log('検索クエリ: ' + SEARCH_QUERY);
    console.log('未処理の検索結果: ' + threads.length + ' 件見つかりました');

    if (threads.length > 0) {
      threads.forEach(thread => {
        console.log('処理中スレッド: ' + thread.getFirstMessageSubject());
        const messages = thread.getMessages();
        messages.forEach(message => {
          processMessage(message, sheet);
        });
        thread.addLabel(label);
        thread.markRead();
      });
      updateMonthlySummary(spreadsheet, SHEET_NAME, SUMMARY_SHEET_NAME);
    }

  } catch (e) {
    console.error('実行エラー: ' + e.toString());
  }
}

function processMessage(message, sheet) {
  const subject = message.getSubject();
  if (subject.indexOf('ビジネスデビット') === -1) return;
  
  const body = message.getPlainBody();
  const date = message.getDate();
  const messageId = message.getId(); // Gmailと紐付けるためのID
  
  try {
    const nameMatch = body.match(/   様/);
    let userName = nameMatch ? nameMatch[1].trim() : '不明';

    const amountMatch = body.match(/ご利用金額（円）\s*:\s*([\d,-]+)/);
    let amountStr = amountMatch ? amountMatch[1] : null;
    
    const merchantMatch = body.match(/ご利用先\s*　*\s*:\s*(.+)/);
    let merchant = merchantMatch ? merchantMatch[1].trim() : null;

    if (amountStr && merchant) {
      const amount = parseInt(amountStr.replace(/,/g, ''), 10);
      
      // データの追加（チェックボックス用のfalseとMessageIDを追加）
      const lastRow = sheet.getLastRow() + 1;
      sheet.appendRow([date, merchant, amount, userName, false, messageId]);
      
      // E列（5列目）にチェックボックスを設定
      sheet.getRange(lastRow, 5).insertCheckboxes();
    }
  } catch (e) {
    console.error('メッセージ表示エラー: ' + e.toString());
  }
}

/**
 * 【追加機能】レシート申請フォームのチェック状態を同期する
 */
function syncFromReceiptForm() {
  // --- 設定項目 ---
  const SOURCE_SS_ID = ''; 
  const SOURCE_SHEET_NAME = 'フォームの回答 1';
  const TARGET_SHEET_NAME = '利用明細';
  // ----------------

  try {
    const targetSs = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = targetSs.getSheetByName(TARGET_SHEET_NAME);
    const sourceSs = SpreadsheetApp.openById(SOURCE_SS_ID);
    const sourceSheet = sourceSs.getSheetByName(SOURCE_SHEET_NAME);

    if (!targetSheet || !sourceSheet) {
      throw new Error('シート名が正しくありません。');
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    const targetData = targetSheet.getDataRange().getValues();

    // 申請フォーム側（ソース）のデータを整理：キーは「日付_金額」
    const checkedMap = {};
    for (let i = 1; i < sourceData.length; i++) {
      const dateVal = sourceData[i][4];   // E列: 日付
      const isChecked = sourceData[i][10]; // K列: カード確認済み
      const amount = sourceData[i][11];    // L列: 金額

      if (isChecked === true && dateVal instanceof Date) {
        // 日付を yyyy-MM-dd 形式に固定して金額と結合
        const dateStr = Utilities.formatDate(dateVal, 'JST', 'yyyy-MM-dd');
        const key = dateStr + '_' + amount;
        checkedMap[key] = true;
      }
    }

    // 利用明細側（ターゲット）をスキャンしてチェックを入れる
    let updateCount = 0;
    for (let j = 1; j < targetData.length; j++) {
      const dateVal = targetData[j][0];   // A列: 受信日時
      const amount = targetData[j][2];    // C列: 金額
      const currentCheck = targetData[j][4]; // E列: チェック済（現状）

      if (dateVal instanceof Date && !currentCheck) {
        const dateStr = Utilities.formatDate(dateVal, 'JST', 'yyyy-MM-dd');
        const key = dateStr + '_' + amount;

        // 申請フォーム側に一致するデータがあればチェック
        if (checkedMap[key]) {
          targetSheet.getRange(j + 1, 5).setValue(true);
          updateCount++;
        }
      }
    }

    console.log('同期完了: ' + updateCount + ' 件更新');
    SpreadsheetApp.getUi().alert('同期完了', updateCount + ' 件の明細をチェック済みにしました。', SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    console.error('同期エラー: ' + e.toString());
    SpreadsheetApp.getUi().alert('同期エラー: ' + e.toString());
  }
}

/**
 * スプレッドシートのチェックをGmailのラベルに反映させる機能
 */
function syncCheckboxesToGmail() {
  const SHEET_NAME = '利用明細';
  const TARGET_LABEL = '確認済み';
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    let label = GmailApp.getUserLabelByName(TARGET_LABEL) || GmailApp.createLabel(TARGET_LABEL);
    let count = 0;

    for (let i = 1; i < data.length; i++) {
      const isChecked = data[i][4]; // E列: チェック済
      const messageId = data[i][5]; // F列: MessageID
      
      if (isChecked === true && messageId) {
        try {
          const msg = GmailApp.getMessageById(messageId);
          const thread = msg.getThread();
          let processed = false; // 処理を行ったかどうかのフラグ
          
          // 1. 「確認済み」ラベルをつける
          const hasLabel = thread.getLabels().some(l => l.getName() === TARGET_LABEL);
          if (!hasLabel) {
            thread.addLabel(label);
            processed = true;
          }

          // 2. 受信トレイにある場合はアーカイブする
          if (thread.isInInbox()) {
            thread.moveToArchive();
            processed = true;
          }
          
          if (processed) count++;

        } catch (err) { console.warn('ID失効またはメール削除済み: ' + messageId); }
      }
    }
    SpreadsheetApp.getUi().alert('ラベル同期＆アーカイブ完了', count + ' 件のメールを処理（ラベル付与/アーカイブ）しました。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { 
    SpreadsheetApp.getUi().alert('エラー: ' + e.toString()); 
  }
}
// --- 以下、既存の集計・メニュー処理（一部更新） ---

function updateMonthlySummary(spreadsheet, sourceSheetName, summarySheetName) {
  const sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  let summarySheet = spreadsheet.getSheetByName(summarySheetName);
  if (!summarySheet) summarySheet = spreadsheet.insertSheet(summarySheetName);
  summarySheet.clear();
  summarySheet.appendRow(['年月', '合計利用金額（円）']);
  summarySheet.setFrozenRows(1);

  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const monthlyTotals = {};
  for (let i = 1; i < data.length; i++) {
    const date = new Date(data[i][0]);
    const amount = Number(data[i][2]);
    if (isNaN(date.getTime()) || isNaN(amount)) continue;
    const monthKey = Utilities.formatDate(date, 'JST', 'yyyy-MM');
    monthlyTotals[monthKey] = (monthlyTotals[monthKey] || 0) + amount;
  }

  const sortedMonths = Object.keys(monthlyTotals).sort();
  sortedMonths.forEach(month => summarySheet.appendRow([month, monthlyTotals[month]]));
  updateChart(summarySheet);
}

function updateChart(sheet) {
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getDataRange())
    .setPosition(2, 4, 0, 0)
    .setOption('title', '月別利用金額推移')
    .build();
  sheet.insertChart(chart);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('デビット管理')
    .addItem('メールからデータ抽出', 'extractMufgDebitInfo')
    .addItem('集計を手動更新', 'manualUpdateSummary')
    .addSeparator()
    .addItem('レシート申請と同期', 'syncFromReceiptForm')
    .addItem('チェックをGmailに反映', 'syncCheckboxesToGmail')
    .addSeparator()
    .addItem('【注意】過去のメールを全再取得', 'reprocessAllHistoricalEmails')
    .addToUi();
}

function manualUpdateSummary() {
  updateMonthlySummary(SpreadsheetApp.getActiveSpreadsheet(), '利用明細', '月別集計');
  SpreadsheetApp.getUi().alert('更新完了');
}

function reprocessAllHistoricalEmails() {
  const SPREADSHEET_ID = null;
  const SHEET_NAME = '利用明細';
  const SEARCH_QUERY = '三菱ＵＦＪ ビジネスデビット after:2025/09/01';
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('全期間再取得', '2025年9月以降のすべてのメールを再スキャンします。既存のデータがある場合、その下に追記されます。実行しますか？', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    const threads = GmailApp.search(SEARCH_QUERY, 0, 200);
    console.log('再スキャン対象: ' + threads.length + ' 件のスレッド');
    threads.forEach(t => {
      console.log('再処理中: ' + t.getFirstMessageSubject());
      t.getMessages().forEach(m => processMessage(m, sheet));
    });
    updateMonthlySummary(ss, SHEET_NAME, '月別集計');
    ui.alert('完了', threads.length + ' 件処理しました。');
  } catch (e) { ui.alert('エラー: ' + e.toString()); }
}
