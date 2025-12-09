/**
 * クオーレ様向けタスク管理ツール v2.0 構築スクリプト（バインド版）
 * 現在開いているスプレッドシートに対して、シート構成と設定を一括適用します。
 * ※注意: 同名のシート（Dashboard等）が既にある場合、削除して作り直します。
 */
function setupV2DemoSheet_Bound() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 作成するシート名の定義
  const targetSheets = {
    dashboard: "Dashboard",
    taskDB: "Task_DB",
    processDB: "Process_DB",
    dropdowns: "Dropdowns"
  };

  // 1. 同名の既存シートがあれば削除 (リセット処理)
  Object.values(targetSheets).forEach(name => {
    const existing = ss.getSheetByName(name);
    if (existing) {
      ss.deleteSheet(existing);
    }
  });

  // 2. シートの新規作成
  const sheetDashboard = ss.insertSheet(targetSheets.dashboard);
  const sheetTaskDB = ss.insertSheet(targetSheets.taskDB);
  const sheetProcessDB = ss.insertSheet(targetSheets.processDB);
  const sheetDropdowns = ss.insertSheet(targetSheets.dropdowns);

  // ==========================================
  // 3. Dropdowns シート設定（担当者リスト置き場）
  // ==========================================
  sheetDropdowns.getRange("A1").setValue("【担当者リスト】(マスタから転記)").setFontWeight("bold").setBackground("#d9ead3");
  // デモ用仮データ
  const initialAssignees = [["本田 啓夫"], ["佐藤 料理長"], ["鈴木 買出"], ["AI アシスタント"]];
  sheetDropdowns.getRange(2, 1, initialAssignees.length, 1).setValues(initialAssignees);

  // ==========================================
  // 4. Process_DB シート設定（工程マスタ）
  // ==========================================
  const processHeaders = ["Process_ID", "Process_Name", "Description"];
  const processData = [
    ["P-01", "買出し", "食材や備品の調達フェーズ"],
    ["P-02", "下準備", "食材のカット、下味付け"],
    ["P-03", "調理", "加熱調理プロセス"],
    ["P-04", "盛り付け", "提供前の最終仕上げ"]
  ];

  sheetProcessDB.getRange(1, 1, 1, processHeaders.length).setValues([processHeaders])
    .setFontWeight("bold").setBackground("#cfe2f3");
  sheetProcessDB.getRange(2, 1, processData.length, processData[0].length).setValues(processData);

  // ==========================================
  // 5. Task_DB シート設定（メイン入力画面）
  // ==========================================
  const taskHeaders = [
    "Process_ID", "Task_ID", "Process_Name", "Task_Name", 
    "Assignee", "Status", "Est_Hours", "Start_Date", "Due_Date", "Notify", "Gantt"
  ];
  
  sheetTaskDB.getRange(1, 1, 1, taskHeaders.length).setValues([taskHeaders])
    .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
  
  sheetTaskDB.setColumnWidth(4, 250); 
  sheetTaskDB.setColumnWidth(11, 200);

  // --- 入力規則 (プルダウン) ---
  
  // E列: Assignee (DropdownsシートのA列を参照)
  const ruleAssignee = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheetDropdowns.getRange("A2:A"))
    .setAllowInvalid(true).build();
  sheetTaskDB.getRange("E2:E100").setDataValidation(ruleAssignee);

  // F列: Status (固定リスト)
  const ruleStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(["⚪️ 未着手", "🔵 進行中", "🟢 完了", "🟡 確認待ち"])
    .setAllowInvalid(true).build();
  sheetTaskDB.getRange("F2:F100").setDataValidation(ruleStatus);

  // J列: Notify (チェックボックス)
  const ruleCheck = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  sheetTaskDB.getRange("J2:J100").setDataValidation(ruleCheck);

  // --- 数式 (VLOOKUP, SPARKLINE) ---
  // C列: Process_Name
  sheetTaskDB.getRange("C2").setFormula('=ARRAYFORMULA(IFERROR(VLOOKUP(A2:A, Process_DB!A:B, 2, FALSE), ""))');

  // K列: 簡易ガントチャート
  sheetTaskDB.getRange("K2").setFormula('=ARRAYFORMULA(IF((I2:I="")+(I2:I<TODAY()), "", SPARKLINE(I2:I-TODAY(), {"charttype","bar";"max",30;"min",0;"color1","#6aa84f"})))');

  // --- 条件付き書式 ---
  const rangeAll = sheetTaskDB.getRange("A2:K100");
  
  // 1. 完了行グレーアウト
  const ruleGray = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="🟢 完了"')
    .setBackground("#EFEFEF")
    .setFontColor("#999999")
    .setRanges([rangeAll])
    .build();
  
  // 2. プロセスID区切り (背景色変更)
  const ruleProcessGroup = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2<>$A1') 
    .setBackground("#fff2cc") 
    .setRanges([sheetTaskDB.getRange("A2:K100")])
    .build();

  const rules = sheetTaskDB.getConditionalFormatRules();
  rules.push(ruleGray);
  rules.push(ruleProcessGroup);
  sheetTaskDB.setConditionalFormatRules(rules);

  sheetTaskDB.setFrozenRows(1);
  sheetTaskDB.setFrozenColumns(4);

  // ==========================================
  // 6. Dashboard シート設定
  // ==========================================
  sheetDashboard.getRange("A1").setValue("【リソース負荷状況】(未完了タスクの工数合計)");
  sheetDashboard.getRange("A2").setFormula('=QUERY(Task_DB!E:G, "select E, sum(G) where F != \'🟢 完了\' and E is not null group by E label sum(G) \'残工数(h)\'", 1)');

  sheetDashboard.getRange("D1").setValue("【設定】Google Chat Webhook URL");
  sheetDashboard.getRange("D2").setBackground("#fff2cc").setValue("");

  sheetDashboard.getRange("D4").setValue("【KPI】期限切れタスク数");
  sheetDashboard.getRange("D5").setFormula('=COUNTIFS(Task_DB!I:I, "<"&TODAY(), Task_DB!F:F, "<>🟢 完了")');
  sheetDashboard.getRange("D5").setFontColor("red").setFontWeight("bold").setFontSize(14);

  // ==========================================
  // 7. 不要シートの掃除
  // ==========================================
  // 作成した4シート以外（元々あった「シート1」など）を削除
  const createdSheetNames = Object.values(targetSheets);
  const allSheets = ss.getSheets();
  
  if (allSheets.length > createdSheetNames.length) {
    allSheets.forEach(sheet => {
      if (!createdSheetNames.includes(sheet.getName())) {
        try {
          ss.deleteSheet(sheet);
        } catch (e) {
          // 削除エラー（最後の1枚など）は無視
          console.log("シート削除スキップ: " + sheet.getName());
        }
      }
    });
  }

  // Dashboardをアクティブにする
  ss.setActiveSheet(sheetDashboard);
  Browser.msgBox("✅ シート構築が完了しました！");
}