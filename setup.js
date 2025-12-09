/**
 * 【v2.3】クオーレ様向けタスク管理ツール (列構成修正版)
 * A-J: タスク情報
 * K  : Group_ID (計算用・非表示)
 * L~ : ガントチャート
 */
function createV2DemoSheet_Corrected() {
  const ss = SpreadsheetApp.create("【デモv2.3】クオーレ様タスク管理ツール_列修正版");
  const defaultSheet = ss.getSheets()[0];

  const sheetDashboard = ss.insertSheet("Dashboard");
  const sheetTaskDB = ss.insertSheet("Task_DB");
  const sheetProcessDB = ss.insertSheet("Process_DB");
  const sheetDropdowns = ss.insertSheet("Dropdowns");
  
  ss.deleteSheet(defaultSheet);

  // --- Dropdowns ---
  sheetDropdowns.getRange("A1").setValue("【担当者リスト】").setFontWeight("bold").setBackground("#d9ead3");
  const initialAssignees = [["本田 啓夫"], ["佐藤 料理長"], ["鈴木 買出"], ["AI アシスタント"]];
  sheetDropdowns.getRange(2, 1, initialAssignees.length, 1).setValues(initialAssignees);

  // --- Process_DB ---
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

  // --- Task_DB 設定 ---
  const fixedHeaders = [
    "Process_ID", "Task_ID", "Process_Name", "Task_Name", 
    "Assignee", "Status", "Est_Hours", "Start_Date", "Due_Date", "Notify"
  ];
  
  sheetTaskDB.getRange(1, 1, 1, fixedHeaders.length).setValues([fixedHeaders])
    .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
  
  // ★変更点1：K列を計算用列に設定
  sheetTaskDB.getRange("K1").setValue("Group_ID");
  // 数式：A列(Process_ID)がユニークリストの何番目か
  sheetTaskDB.getRange("K2").setFormula('=ARRAYFORMULA(IF(A2:A="", "", MATCH(A2:A, UNIQUE(A2:A), 0)))');
  sheetTaskDB.hideColumns(11); // K列を隠す

  // ★変更点2：L列(12列目)以降をガントチャートに設定
  const today = new Date();
  const dateHeaders = [];
  for (let i = 0; i < 60; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() + i);
    dateHeaders.push(d);
  }
  sheetTaskDB.getRange(1, 12, 1, dateHeaders.length) // 12列目から書き込み
    .setValues([dateHeaders])
    .setNumberFormat("M/d")
    .setBackground("#f3f3f3")
    .setFontColor("black")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 列幅
  sheetTaskDB.setColumnWidth(4, 250);
  sheetTaskDB.setColumnWidths(12, 60, 25); // L列以降を狭く

  // 入力規則
  const ruleAssignee = SpreadsheetApp.newDataValidation().requireValueInRange(sheetDropdowns.getRange("A2:A")).setAllowInvalid(true).build();
  sheetTaskDB.getRange("E2:E100").setDataValidation(ruleAssignee);
  const ruleStatus = SpreadsheetApp.newDataValidation().requireValueInList(["⚪️ 未着手", "🔵 進行中", "🟢 完了", "🟡 確認待ち"]).setAllowInvalid(true).build();
  sheetTaskDB.getRange("F2:F100").setDataValidation(ruleStatus);
  const ruleCheck = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheetTaskDB.getRange("J2:J100").setDataValidation(ruleCheck);

  // 数式 (C列)
  sheetTaskDB.getRange("C2").setFormula('=ARRAYFORMULA(IFERROR(VLOOKUP(A2:A, Process_DB!A:B, 2, FALSE), ""))');

  // --- 条件付き書式 ---
  const rules = sheetTaskDB.getConditionalFormatRules();

  // 1. 完了行グレーアウト
  const ruleGray = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="🟢 完了"')
    .setBackground("#EFEFEF")
    .setFontColor("#999999")
    .setRanges([sheetTaskDB.getRange("A2:BM100")])
    .build();
  rules.push(ruleGray);

  // 2. プロセスごとの色分け (A~D列)
  // ★修正：K列($K2)を参照して奇数判定
  const rangeProcessCols = sheetTaskDB.getRange("A2:D100");
  const rulePink = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISODD($K2)') 
    .setBackground("#EAD1DC") // ピンク
    .setRanges([rangeProcessCols])
    .build();
  rules.push(rulePink);

  // 3. ガントチャートバー
  // ★修正：日付はL$1から、範囲はL2から
  const ganttRange = sheetTaskDB.getRange(2, 12, 100, 60); // L列から
  const ruleGantt = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(L$1>=$H2, L$1<=$I2)')
    .setBackground("#6aa84f")
    .setRanges([ganttRange])
    .build();
  rules.push(ruleGantt);

  // 4. 今日線
  const ruleToday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=L$1=TODAY()')
    .setBackground("#fff2cc")
    .setRanges([ganttRange])
    .build();
  rules.push(ruleToday);

  sheetTaskDB.setConditionalFormatRules(rules);

  sheetTaskDB.setFrozenRows(1);
  sheetTaskDB.setFrozenColumns(4);

  // --- Dashboard ---
  sheetDashboard.getRange("A1").setValue("【リソース負荷状況】");
  sheetDashboard.getRange("A2").setFormula('=QUERY(Task_DB!E:G, "select E, sum(G) where F != \'🟢 完了\' and E is not null group by E label sum(G) \'残工数(h)\'", 1)');
  sheetDashboard.getRange("D1").setValue("【設定】Google Chat Webhook URL");
  sheetDashboard.getRange("D2").setBackground("#fff2cc");

  Logger.log("URL: " + ss.getUrl());
}