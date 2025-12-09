/**
 * 【v2.1】クオーレ様向けタスク管理ツール構築 (本格ガントチャート版)
 * K列以降に日付を展開し、条件付き書式で期間を塗りつぶします。
 */
function createV2DemoSheet_Gantt() {
  const ss = SpreadsheetApp.create("【デモv2.1】クオーレ様タスク管理ツール_ガントチャート版");
  const defaultSheet = ss.getSheets()[0];

  const sheetDashboard = ss.insertSheet("Dashboard");
  const sheetTaskDB = ss.insertSheet("Task_DB");
  const sheetProcessDB = ss.insertSheet("Process_DB");
  
  ss.deleteSheet(defaultSheet);

  // --- Process_DB 設定 ---
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
  
  // 固定列のヘッダーセット
  sheetTaskDB.getRange(1, 1, 1, fixedHeaders.length).setValues([fixedHeaders])
    .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
  
  // ★変更点：K列以降に「日付ヘッダー」を展開 (今日から60日分)
  const today = new Date();
  const dateHeaders = [];
  for (let i = 0; i < 60; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() + i);
    dateHeaders.push(d);
  }
  // K1セルから日付を書き込み
  sheetTaskDB.getRange(1, 11, 1, dateHeaders.length) // 11列目(K列)から
    .setValues([dateHeaders])
    .setNumberFormat("M/d") // 日付フォーマット
    .setBackground("#f3f3f3")
    .setFontColor("black")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 列幅調整
  sheetTaskDB.setColumnWidth(4, 250); // Task_Name
  sheetTaskDB.setColumnWidths(11, 60, 25); // ガントチャートエリアを狭く(25px)して見やすく

  // --- 入力規則 ---
  const demoAssignees = ["本田 啓夫", "佐藤 料理長", "鈴木 買出", "AI アシスタント"];
  const ruleAssignee = SpreadsheetApp.newDataValidation().requireValueInList(demoAssignees).setAllowInvalid(true).build();
  sheetTaskDB.getRange("E2:E100").setDataValidation(ruleAssignee);

  const ruleStatus = SpreadsheetApp.newDataValidation().requireValueInList(["⚪️ 未着手", "🔵 進行中", "🟢 完了", "🟡 確認待ち"]).setAllowInvalid(true).build();
  sheetTaskDB.getRange("F2:F100").setDataValidation(ruleStatus);

  const ruleCheck = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheetTaskDB.getRange("J2:J100").setDataValidation(ruleCheck);

  // --- 数式 ---
  sheetTaskDB.getRange("C2").setFormula('=ARRAYFORMULA(IFERROR(VLOOKUP(A2:A, Process_DB!A:B, 2, FALSE), ""))');

  // --- 条件付き書式 (ガントチャートの描画) ---
  const rules = sheetTaskDB.getConditionalFormatRules();

  // 1. ガントチャートバー (期間塗りつぶし)
  // 範囲: K2:BM100 (日付エリア)
  // 条件: カレンダーの日付(K$1)が、開始日($H2)以上 かつ 期限($I2)以下 の場合
  const ganttRange = sheetTaskDB.getRange(2, 11, 100, 60);
  const ruleGantt = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(K$1>=$H2, K$1<=$I2)')
    .setBackground("#6aa84f") // 緑色
    .setRanges([ganttRange])
    .build();
  rules.push(ruleGantt);

  // 2. 今日線 (縦ライン)
  const ruleToday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=K$1=TODAY()')
    .setBackground("#fff2cc") // 薄い黄色
    .setRanges([ganttRange])
    .build();
  rules.push(ruleToday);

  // 3. 完了行グレーアウト (全体)
  const ruleGray = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="🟢 完了"')
    .setBackground("#EFEFEF")
    .setFontColor("#999999")
    .setRanges([sheetTaskDB.getRange("A2:BM100")])
    .build();
  rules.push(ruleGray);

  // 4. プロセス区切り
  const ruleProcessGroup = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2<>$A1')
    .setBackground("#e6b8af") // 少し濃い色で区切り
    .setRanges([sheetTaskDB.getRange("A2:A100")]) // A列のみ色付け
    .build();
  rules.push(ruleProcessGroup);

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