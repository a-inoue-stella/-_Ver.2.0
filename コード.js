/**
 * 【v2.0】クオーレ様向けタスク管理ツール Backend Logic
 * Feature: プロセス管理、工数管理、AIプラン取り込み、Chat通知
 */

// --- 1. 設定 (CONFIG) ---
// シートの列番号や設定値をここで一元管理します
const CONFIG = {
  SHEET_TASK: 'Task_DB',
  SHEET_PROCESS: 'Process_DB',
  SHEET_DASHBOARD: 'Dashboard',
  
  // Task_DBの列番号 (A列=1)
  COL_PROCESS_ID: 1,
  COL_TASK_ID: 2,
  COL_PROCESS_NAME: 3,
  COL_TASK_NAME: 4,
  COL_ASSIGNEE: 5,
  COL_STATUS: 6,
  COL_EST_HOURS: 7, // ★新規: 工数
  COL_START: 8,
  COL_DUE: 9,
  COL_NOTIFY: 10,   // J列: 通知チェックボックス
  
  // DashboardのWebhook URL入力セル
  CELL_WEBHOOK: 'D2'
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
 * 1-1. 入力用モーダルの表示
 */
function openImportModal() {
  const html = `
    <div style="font-family:sans-serif; padding:10px;">
      <h3>🤖 AIプラン取り込み</h3>
      <p>Geminiが生成したJSONを貼り付けてください。</p>
      <textarea id="json" style="width:100%; height:300px; font-family:monospace;"></textarea>
      <br><br>
      <button onclick="runImport()" style="padding:10px 20px; font-weight:bold; cursor:pointer;">取り込み実行</button>
      <div id="status" style="margin-top:10px; font-weight:bold;"></div>
      <script>
        function runImport() {
          const json = document.getElementById('json').value;
          document.getElementById('status').innerText = '処理中...';
          google.script.run
            .withSuccessHandler(msg => document.getElementById('status').innerText = msg)
            .withFailureHandler(err => document.getElementById('status').innerText = 'エラー: ' + err.message)
            .processAiPlan(json);
        }
      </script>
    </div>
  `;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(600).setHeight(500), 'AIプランナー連携');
}

/**
 * 1-2. JSON解析とDBへの書き込み (サーバー側処理)
 */
function processAiPlan(jsonString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTask = ss.getSheetByName(CONFIG.SHEET_TASK);
  const sheetProcess = ss.getSheetByName(CONFIG.SHEET_PROCESS);

  try {
    // JSONパース（配列であることを期待）
    const planData = JSON.parse(jsonString);
    if (!Array.isArray(planData)) throw new Error("JSONは配列形式である必要があります");

    // --- A. Process_DB の更新 ---
    // 既存のプロセスIDを取得して重複を防ぐ
    const existProcs = sheetProcess.getRange("A2:A").getValues().flat().filter(String);
    const newProcesses = [];
    const seenProcIds = new Set(existProcs);

    planData.forEach(item => {
      if (item.process_id && !seenProcIds.has(item.process_id)) {
        newProcesses.push([item.process_id, item.process_name || "", "AI生成"]);
        seenProcIds.add(item.process_id);
      }
    });

    if (newProcesses.length > 0) {
      const lastRowP = sheetProcess.getLastRow();
      sheetProcess.getRange(lastRowP + 1, 1, newProcesses.length, 3).setValues(newProcesses);
    }

    // --- B. Task_DB の更新 ---
    // Task_IDの最大値を取得して連番生成
    const existTaskIds = sheetTask.getRange("B2:B").getValues().flat();
    let maxId = 0;
    existTaskIds.forEach(id => {
      if (typeof id === 'string' && id.startsWith('TASK-')) {
        const num = parseInt(id.replace('TASK-', ''), 10);
        if (!isNaN(num) && num > maxId) maxId = num;
      }
    });

    const newTasks = planData.map((item, i) => {
      const nextId = maxId + i + 1;
      const taskId = 'TASK-' + ('000' + nextId).slice(-3);
      
      // 日付計算 (今日 + offset)
      const today = new Date();
      const start = new Date(today);
      const due = new Date(today);
      if (item.due_offset) due.setDate(today.getDate() + item.due_offset);

      return [
        item.process_id || "",      // A: Process_ID
        taskId,                     // B: Task_ID
        "",                         // C: Process_Name (数式で自動表示)
        item.task_name || "",       // D: Task_Name
        item.assignee_name || "",   // E: Assignee
        "⚪️ 未着手",                // F: Status
        item.est_hours || 1,        // G: Est_Hours (工数)
        start,                      // H: Start
        due,                        // I: Due
        false,                      // J: Notify
        ""                          // K: Gantt (数式)
      ];
    });

    // 書き込み（C列, K列は数式が入っている前提なので上書き注意だが、
    // 今回のシート構築スクリプトではARRAYFORMULAを使っているため、
    // 空欄を書き込んでも数式が生きる、もしくは値として書き込む）
    // ※今回は値として書き込みます。C列はARRAYFORMULAが入っているので空文字でOK。
    
    // A列(Process_ID)の最終行を探して追記
    const lastRowT = sheetTask.getLastRow(); 
    // ※getLastRowはデータがある最終行。数式だけの行はカウントされない場合があるが、
    // 配列渡しで書き込むため、正確な位置特定が必要。
    // 安全のため、A列の値を見て最終行を判定
    const valsA = sheetTask.getRange("A1:A").getValues().flat();
    let realLastRow = valsA.length;
    while (realLastRow > 0 && valsA[realLastRow - 1] === "") {
      realLastRow--;
    }
    
    sheetTask.getRange(realLastRow + 1, 1, newTasks.length, newTasks[0].length).setValues(newTasks);

    return `✅ 成功！ ${newTasks.length}件のタスクと${newProcesses.length}件のプロセスを追加しました。`;

  } catch (e) {
    return "❌ エラー: " + e.message;
  }
}

/* ==========================================================================
   機能2: 通知トリガー (チェックボックスONで通知)
   ========================================================================== */

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Task_DBシートの「Notify(J列)」がTRUEになった時のみ発動
  if (sheet.getName() !== CONFIG.SHEET_TASK) return;
  if (range.getColumn() !== CONFIG.COL_NOTIFY) return;
  if (e.value !== "TRUE") return;

  sendNotificationCard(sheet, range.getRow());
}

/**
 * 行データを取得してChatに送信し、チェックを外す
 */
function sendNotificationCard(sheet, row) {
  // データの取得
  const rowData = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const data = {
    processName: rowData[CONFIG.COL_PROCESS_ID - 1], // Process_IDから名前引くのは複雑なのでIDか、VLOOKUP済のC列を取るか
    // C列の値を取りたいが、getRowDataだと生の値。
    // ここでは簡便のため、C列（Process_Name）を直接取得しにいく
    processNameReal: sheet.getRange(row, CONFIG.COL_PROCESS_NAME).getValue(),
    taskName: rowData[CONFIG.COL_TASK_NAME - 1],
    assignee: rowData[CONFIG.COL_ASSIGNEE - 1],
    status: rowData[CONFIG.COL_STATUS - 1],
    estHours: rowData[CONFIG.COL_EST_HOURS - 1],
    due: rowData[CONFIG.COL_DUE - 1]
  };

  const webhookUrl = getWebhookUrl();
  if (!webhookUrl) {
    Browser.msgBox("Webhook URLが設定されていません (Dashboard!D2)");
    sheet.getRange(row, CONFIG.COL_NOTIFY).setValue(false);
    return;
  }

  // カードペイロード作成
  const payload = createCardV2(data);
  
  // 送信
  sendToWebhook(webhookUrl, payload);

  // チェックを戻す
  sheet.getRange(row, CONFIG.COL_NOTIFY).setValue(false);
}

/**
 * v2.0用 リッチなカード通知を作成
 */
function createCardV2(d) {
  const dateStr = d.due ? Utilities.formatDate(d.due, Session.getScriptTimeZone(), 'MM/dd') : '未定';
  
  return {
    "cardsV2": [{
      "cardId": "task-card",
      "card": {
        "header": {
          "title": "【タスク通知】" + d.taskName,
          "subtitle": `工程: ${d.processNameReal} | 工数: ${d.estHours}h`,
          "imageUrl": "https://www.gstatic.com/images/icons/material/system/2x/assignment_ind_black_48dp.png",
          "imageType": "CIRCLE"
        },
        "sections": [
          {
            "widgets": [
              {
                "decoratedText": {
                  "startIcon": { "knownIcon": "PERSON" },
                  "topLabel": "担当者",
                  "text": `<b>${d.assignee}</b>`
                }
              },
              {
                "decoratedText": {
                  "startIcon": { "knownIcon": "CLOCK" },
                  "topLabel": "期限 / 状況",
                  "text": `${dateStr}  <font color="${d.status=='🟢 完了'?'#00AA00':'#FF0000'}">${d.status}</font>`
                }
              }
            ]
          },
          {
            "widgets": [
              {
                "buttonList": {
                  "buttons": [{
                    "text": "シートを開く",
                    "onClick": {
                      "openLink": { "url": SpreadsheetApp.getActiveSpreadsheet().getUrl() }
                    }
                  }]
                }
              }
            ]
          }
        ]
      }
    }]
  };
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

// リマインド機能（手動実行用）
// 今回はデモ用なので、単純に「未完了タスク」をいくつかピックアップして通知する簡易版
function sendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_TASK);
  const data = sheet.getDataRange().getValues();
  const webhookUrl = getWebhookUrl();

  let count = 0;
  // ヘッダー飛ばして走査
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[CONFIG.COL_STATUS - 1];
    const due = row[CONFIG.COL_DUE - 1];
    
    // 「進行中」かつ「今日以前」のものがあれば通知
    // デモ演出用: 条件を緩くして、1つ見つけたら通知して終わる（スパム防止）
    if (status === "🔵 進行中" && count < 1) {
      // 無理やり通知関数を呼ぶ（行番号は i+1）
      sendNotificationCard(sheet, i + 1);
      count++;
    }
  }
  
  if (count === 0) Browser.msgBox("リマインド対象（進行中）が見つかりませんでした。ステータスを変更して試してください。");
  else Browser.msgBox("リマインドを1件送信しました（デモ用制限）");
}