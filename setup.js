/**
 * クオーレ様向けタスク管理ツール v2.0 構築スクリプト
 * ・Masterシートなし
 * ・「Dropdowns」シートを作成し、担当者リストをそこから参照する形式に変更
 */
function createV2DemoSheet_WithDropdown() {
  // 1. 新規スプレッドシート作成
  const ss = SpreadsheetApp.create("【デモv2.0】クオーレ様タスク管理ツール_プルダウン連携版");
  const defaultSheet = ss.getSheets()[0];

  // --- シートの作成 ---
  const sheetDashboard = ss.insertSheet("Dashboard");
  const sheetTaskDB = ss.insertSheet("Task_DB");
  const sheetProcessDB = ss.insertSheet("Process_DB");
  const sheetDropdowns = ss.insertSheet("Dropdowns"); // ★新規追加：プルダウン用シート
  
  // デフォルトの「シート1」を削除
  ss.deleteSheet(defaultSheet);

  // ==========================================
  // 2. Dropdowns シート設定（担当者リスト置き場）
  // ==========================================
  // 後でマスタから転記しやすいよう、A列を担当者リスト枠として空けておきます
  sheetDropdowns.getRange("A1").setValue("【担当者リスト】(マスタから転記)").setFontWeight("bold").setBackground("#d9ead3");
  // デモ用に仮のデータを入れておきます（後で上書きしてください）
  const initialAssignees = [["本田 啓夫"], ["佐藤 料理長"], ["鈴木 買出"], ["AI アシスタント"]];
  sheetDropdowns.getRange(2, 1, initialAssignees.length, 1).setValues(initialAssignees);

  // ==========================================
  // 3. Process_DB シート設定（工程マスタ）
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
  // 4. Task_DB シート設定（メイン入力画面）
  // ==========================================
  const taskHeaders = [
    "Process_ID", "Task_ID", "Process_Name", "Task_Name", 
    "Assignee", "Status", "Est_Hours", "Start_Date", "Due_Date", "Notify", "Gantt"
  ];
  
  sheetTaskDB.getRange(1, 1, 1, taskHeaders.length).setValues([taskHeaders])
    .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
  
  sheetTaskDB.setColumnWidth(4, 250); 
  sheetTaskDB.setColumnWidth(11, 200);

  // --- 入力規則 (プルダウン) の設定 ---
  
  // E列: Assignee (★変更点：DropdownsシートのA列を参照するように設定)
  const ruleAssignee = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheetDropdowns.getRange("A2:A")) // A列全体を範囲指定
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

  // --- 数式の設定 ---
  // C列: Process_Name
  sheetTaskDB.getRange("C2").setFormula('=ARRAYFORMULA(IFERROR(VLOOKUP(A2:A, Process_DB!A:B, 2, FALSE), ""))');

  // K列: 簡易ガントチャート
  sheetTaskDB.getRange("K2").setFormula('=ARRAYFORMULA(IF((I2:I="")+(I2:I<TODAY()), "", SPARKLINE(I2:I-TODAY(), {"charttype","bar";"max",30;"min",0;"color1","#6aa84f"})))');

  // --- 条件付き書式の設定 ---
  const rangeAll = sheetTaskDB.getRange("A2:K100");
  const ruleGray = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="🟢 完了"')
    .setBackground("#EFEFEF")
    .setFontColor("#999999")
    .setRanges([rangeAll])
    .build();
  
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
  // 5. Dashboard シート設定
  // ==========================================
  sheetDashboard.getRange("A1").setValue("【リソース負荷状況】(未完了タスクの工数合計)");
  sheetDashboard.getRange("A2").setFormula('=QUERY(Task_DB!E:G, "select E, sum(G) where F != \'🟢 完了\' and E is not null group by E label sum(G) \'残工数(h)\'", 1)');

  sheetDashboard.getRange("D1").setValue("【設定】Google Chat Webhook URL");
  sheetDashboard.getRange("D2").setBackground("#fff2cc").setValue("");

  sheetDashboard.getRange("D4").setValue("【KPI】期限切れタスク数");
  sheetDashboard.getRange("D5").setFormula('=COUNTIFS(Task_DB!I:I, "<"&TODAY(), Task_DB!F:F, "<>🟢 完了")');
  sheetDashboard.getRange("D5").setFontColor("red").setFontWeight("bold").setFontSize(14);

  Logger.log("作成完了URL: " + ss.getUrl());
}