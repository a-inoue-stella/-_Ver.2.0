/**
 * 【v2.0】クオーレ様向けタスク管理ツール Backend Logic
 * Feature: プロセス管理、工数管理、AIプラン取り込み、Chat通知
 */

// --- 1. 設定 (CONFIG) ---
// ★修正：シート名を日本語版に合わせて変更
const CONFIG = {
  SHEET_TASK: 'タスク管理',
  SHEET_PROCESS: 'プロセスマスタ',
  SHEET_DASHBOARD: 'ダッシュボード',
  
  // 列番号 (A列=1) ※変更なし
  COL_PROCESS_ID: 1,
  COL_TASK_ID: 2,
  COL_PROCESS_NAME: 3,
  COL_TASK_NAME: 4,
  COL_ASSIGNEE: 5,
  COL_STATUS: 6,
  COL_EST_HOURS: 7, 
  COL_START: 8,
  COL_DUE: 9,
  COL_NOTIFY: 10,   
  
  CELL_WEBHOOK: 'E2'
};

/**
 * メニューバー追加
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚡️ タスク管理v2.0')
    .addItem('🤖 AIプラン取り込み (JSON)', 'openImportModal')
    .addSeparator()
    .addItem('🔔 リマインド送信 (手動)', 'sendReminders')
    .addToUi();
}

/* ==========================================================================
   機能1: AIプラン取り込み (JSON解析 & DB展開)
   ========================================================================== */

/**
 * 1-1. 入力用モーダルの表示 (修正版：完了通知機能付き)
 */
function openImportModal() {
  const html = `
    <div style="font-family:sans-serif; padding:10px;">
      <h3>🤖 AIプラン取り込み</h3>
      <p>Geminiが生成したJSONを貼り付けてください。</p>
      <textarea id="json" style="width:100%; height:300px; font-family:monospace;"></textarea>
      <br><br>
      <button id="btn" onclick="runImport()" style="padding:10px 20px; font-weight:bold; cursor:pointer;">取り込み実行</button>
      <div id="status" style="margin-top:10px; font-weight:bold;"></div>
      <script>
        function runImport() {
          const json = document.getElementById('json').value;
          if (!json) {
            alert("JSONが入力されていません");
            return;
          }
          
          // ボタンを無効化し、処理中表示にする
          const btn = document.getElementById('btn');
          const status = document.getElementById('status');
          btn.disabled = true;
          btn.innerText = "処理中...";
          status.innerText = '🔄 スプレッドシートに書き込んでいます...少々お待ちください。';

          google.script.run
            .withSuccessHandler(msg => {
              // ★完了時の挙動：アラートを出して閉じる
              status.innerText = '✅ 完了しました！';
              window.alert(msg); // ポップアップ通知
              google.script.host.close(); // モーダルを閉じる
            })
            .withFailureHandler(err => {
              // エラー時はボタンを戻す
              btn.disabled = false;
              btn.innerText = "取り込み実行";
              status.innerText = '❌ エラー: ' + err.message;
              window.alert('エラーが発生しました:\\n' + err.message);
            })
            .processAiPlan(json);
        }
      </script>
    </div>
  `;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(600).setHeight(550), 'AIプランナー連携');
}

/**
 * 1-2. JSON解析とDBへの書き込み (サーバー側処理)
 * ★修正版：日付から時間情報を削除 (00:00:00化) して書き込む
 */
function processAiPlan(jsonString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTask = ss.getSheetByName(CONFIG.SHEET_TASK);
  const sheetProcess = ss.getSheetByName(CONFIG.SHEET_PROCESS);

  try {
    const planData = JSON.parse(jsonString);
    if (!Array.isArray(planData)) throw new Error("JSONは配列形式である必要があります");

    // --- A. Process_DB の更新 (Insert Only) ---
    const lastRowP = sheetProcess.getLastRow();
    const existingIds = new Set();
    
    if (lastRowP > 1) {
      const ids = sheetProcess.getRange(2, 1, lastRowP - 1, 1).getValues().flat();
      ids.forEach(id => { if(id) existingIds.add(id); });
    }

    const newProcesses = [];
    const seenProcIdsInJson = new Set(); 

    planData.forEach(item => {
      const pId = item.process_id;
      const pName = item.process_name || "";

      if (!pId) return;
      if (seenProcIdsInJson.has(pId)) return; 
      seenProcIdsInJson.add(pId);

      if (!existingIds.has(pId)) {
        newProcesses.push([pId, pName, "AI自動生成(新規)"]);
        existingIds.add(pId);
      }
    });

    if (newProcesses.length > 0) {
      const insertRow = sheetProcess.getLastRow() + 1;
      sheetProcess.getRange(insertRow, 1, newProcesses.length, 3).setValues(newProcesses);
    }

    // --- B. Task_DB の更新 ---
    const existTaskIds = sheetTask.getRange("B2:B").getValues().flat();
    let maxId = 0;
    existTaskIds.forEach(id => {
      if (typeof id === 'string' && id.startsWith('TASK-')) {
        const num = parseInt(id.replace('TASK-', ''), 10);
        if (!isNaN(num) && num > maxId) maxId = num;
      }
    });

    const newTasksPart1 = []; 
    const newTasksPart2 = []; 

    planData.forEach((item, i) => {
      const nextId = maxId + i + 1;
      const taskId = 'TASK-' + ('000' + nextId).slice(-3);
      
      // ★修正ポイント：日付オブジェクトの時間をリセットする
      const today = new Date();
      today.setHours(0, 0, 0, 0); // 時・分・秒・ミリ秒を0にする

      const start = new Date(today);
      const due = new Date(today);
      
      if (item.start_offset !== undefined) start.setDate(today.getDate() + item.start_offset);
      if (item.due_offset !== undefined) due.setDate(today.getDate() + item.due_offset);

      newTasksPart1.push([item.process_id || "", taskId]);
      newTasksPart2.push([
        item.task_name || "",       
        item.assignee_name || "",   
        "⚪️ 未着手",                
        item.est_hours || 1,        
        start, // 時間が0:00になったDateオブジェクト
        due,   // 時間が0:00になったDateオブジェクト
        false                       
      ]);
    });

    // 書き込み
    const valsA = sheetTask.getRange("A1:A").getValues().flat();
    let realLastRow = valsA.length;
    while (realLastRow > 0 && valsA[realLastRow - 1] === "") {
      realLastRow--;
    }
    const startRow = realLastRow + 1;

    if (newTasksPart1.length > 0) {
      sheetTask.getRange(startRow, 1, newTasksPart1.length, 2).setValues(newTasksPart1);
      sheetTask.getRange(startRow, 4, newTasksPart2.length, 7).setValues(newTasksPart2);
    }

    ss.toast(`タスク${newTasksPart1.length}件を取り込みました。`, "🤖 取り込み完了", 5);
    return `✅ 成功！\nタスク ${newTasksPart1.length}件を追加しました。\n(新規プロセス: ${newProcesses.length}件)`;

  } catch (e) {
    throw e;
  }
}

/* ==========================================================================
   機能2: 通知トリガー (チェックボックスONで通知)
   ========================================================================== */

function onCheck(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Task_DBシートの「Notify(J列)」がTRUEになった時のみ発動
  if (sheet.getName() !== CONFIG.SHEET_TASK) return;
  if (range.getColumn() !== CONFIG.COL_NOTIFY) return;
  if (e.value !== "TRUE") return;

  sendNotificationCard(sheet, range.getRow());
}

/**
 * 行データを取得してChatに送信し、チェックを外す (修正版)
 */
function sendNotificationCard(sheet, row) {
  const data = sheet.getRange(row, 1, 1, 10).getValues()[0];
  
  // データのマッピング
  const taskInfo = {
    processName: data[CONFIG.COL_PROCESS_NAME - 1], // 工程名
    taskName:    data[CONFIG.COL_TASK_NAME - 1],    // タスク名
    assignee:    data[CONFIG.COL_ASSIGNEE - 1],     // 担当者
    status:      data[CONFIG.COL_STATUS - 1],       // ステータス
    estHours:    data[CONFIG.COL_EST_HOURS - 1],    // 工数
    due:         data[CONFIG.COL_DUE - 1]           // 期限日(Date)
  };

  const webhookUrl = getWebhookUrl();
  if (!webhookUrl) {
    Browser.msgBox("Webhook URLが設定されていません (ダッシュボード!D2)");
    sheet.getRange(row, CONFIG.COL_NOTIFY).setValue(false);
    return;
  }

  // カード作成（通常通知モード）
  const payload = createCardPayload(taskInfo, "NORMAL");
  
  // 送信
  sendToWebhook(webhookUrl, payload);

  // チェックを戻す
  sheet.getRange(row, CONFIG.COL_NOTIFY).setValue(false);
}

/**
 * ★追加：リッチなカード通知を作成する共通関数
 * type: "NORMAL" | "REMIND_DELAY" | "REMIND_TODAY" | "REMIND_TOMORROW"
 */
function createCardPayload(d, type) {
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const dateStr = d.due ? Utilities.formatDate(d.due, Session.getScriptTimeZone(), 'yyyy/MM/dd') : '未設定';

  // --- 1. ヘッダーのデザイン定義 ---
  let headerTitle = "【通知】タスク更新";
  let headerSubtitle = "タスク管理Botより";
  let headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/notifications_black_48dp.png";
  
  // ステータスやタイプによる分岐
  if (type === "REMIND_DELAY") {
    headerTitle = "🔥 【遅延】期限が過ぎています！";
    headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/warning_amber_black_48dp.png"; // ビックリマーク
  } else if (type === "REMIND_TODAY") {
    headerTitle = "⏰ 【今日】本日が対応期限です";
    headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/alarm_black_48dp.png"; // 時計
  } else if (type === "REMIND_TOMORROW") {
    headerTitle = "⚠️ 【明日】明日が期限です";
    headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/event_black_48dp.png"; // カレンダー
  } else if (d.status === "🟡 確認待ち") {
    headerTitle = "🟡 【確認依頼】承認をお願いします";
    headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/account_circle_black_48dp.png"; // 人型
  } else if (d.status === "🟢 完了") {
    headerTitle = "🟢 【完了】タスクが完了しました";
    headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/check_circle_black_48dp.png"; // チェック
  }

  // --- 2. カードの構築 ---
  return {
    "cardsV2": [
      {
        "cardId": "task-card-" + new Date().getTime(),
        "card": {
          "header": {
            "title": headerTitle,
            "subtitle": headerSubtitle,
            "imageUrl": headerIcon,
            "imageType": "SQUARE" // アイコンを大きく表示 [4117]
          },
          "sections": [
            {
              "widgets": [
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "DESCRIPTION" },
                    "topLabel": "タスク / 工程",
                    "text": `<b>${d.taskName}</b><br><font color="#666666">${d.processName}</font>`,
                    "wrapText": true
                  }
                },
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "PERSON" },
                    "topLabel": "担当者",
                    "text": `<b>${d.assignee}</b>`
                  }
                },
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "BOOKMARK" },
                    "topLabel": "ステータス",
                    "text": `<b>${d.status}</b>`
                  }
                },
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "CLOCK" },
                    "topLabel": "期限日",
                    "text": `<b>${dateStr}</b>`
                  }
                }
              ]
            },
            {
              "widgets": [
                {
                  "buttonList": {
                    "buttons": [
                      {
                        "text": "シートを開く",
                        "onClick": {
                          "openLink": { "url": sheetUrl }
                        }
                      }
                    ]
                  }
                }
              ]
            }
          ]
        }
      }
    ]
  };
}

/**
 * 4. リマインド実行 (修正版：期限切れ・今日・明日を区別して通知)
 * メニュー「🔔 リマインド送信」から実行
 */
function sendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_TASK);
  const webhookUrl = getWebhookUrl();

  if (!webhookUrl) {
    Browser.msgBox("Webhook URLが設定されていません");
    return;
  }

  // データ取得 (ヘッダー除く)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Browser.msgBox("データがありません");
    return;
  }
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  // 日付の基準を作成 (時刻は0:00にリセット)
  const today = new Date();
  today.setHours(0,0,0,0);
  
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  
  let alertCount = 0;

  data.forEach(row => {
    const taskInfo = {
      processName: row[CONFIG.COL_PROCESS_NAME - 1],
      taskName:    row[CONFIG.COL_TASK_NAME - 1],
      assignee:    row[CONFIG.COL_ASSIGNEE - 1],
      status:      row[CONFIG.COL_STATUS - 1],
      estHours:    row[CONFIG.COL_EST_HOURS - 1],
      due:         row[CONFIG.COL_DUE - 1]
    };

    // 完了済み、または期限設定なし、タスク名なしはスキップ
    if (taskInfo.status === "🟢 完了" || !taskInfo.taskName || !taskInfo.due) return;

    // 期限日(Date型)の時刻リセット
    const deadline = new Date(taskInfo.due);
    deadline.setHours(0,0,0,0);

    let type = "";

    // 判定ロジック
    if (deadline.getTime() < today.getTime()) {
      type = "REMIND_DELAY";    // 期限切れ
    } else if (deadline.getTime() === today.getTime()) {
      type = "REMIND_TODAY";    // 今日
    } else if (deadline.getTime() === tomorrow.getTime()) {
      type = "REMIND_TOMORROW"; // 明日
    }

    // 対象なら通知
    if (type !== "") {
      const payload = createCardPayload(taskInfo, type);
      sendToWebhook(webhookUrl, payload);
      alertCount++;
      Utilities.sleep(300); // 連続送信によるエラー防止のウェイト
    }
  });

  if(alertCount > 0) {
    Browser.msgBox(`送信完了：${alertCount}件のリマインドを送りました。`);
  } else {
    Browser.msgBox("リマインド対象（遅延・今日・明日）のタスクはありませんでした。");
  }
}

/* ==========================================================================
   ユーティリティ
   ========================================================================== */

function getWebhookUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  return dashboard.getRange(CONFIG.CELL_WEBHOOK).getValue();
}

function sendToWebhook(url, payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}
