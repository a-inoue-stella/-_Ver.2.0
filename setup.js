/**
 * 【v2.5】クオーレ様向けタスク管理ツール (日本語版 & 4色プロセス)
 * シート名・項目名を日本語化し、4色プロセス色分けを適用します。
 */
function createV2DemoSheet_Japanese() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ★変更：シート名を日本語に
  const targetSheets = {
    dashboard: "ダッシュボード",
    taskDB: "タスク管理",
    processDB: "プロセスマスタ",
    dropdowns: "担当者リスト"
  };

  // 1. リセット処理
  Object.values(targetSheets).forEach(name => {
    const existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
  });

  // 2. 新規作成
  const sheetDashboard = ss.insertSheet(targetSheets.dashboard);
  const sheetTaskDB = ss.insertSheet(targetSheets.taskDB);
  const sheetProcessDB = ss.insertSheet(targetSheets.processDB);
  const sheetDropdowns = ss.insertSheet(targetSheets.dropdowns);

  // --- 担当者リスト (Dropdowns) ---
  sheetDropdowns.getRange("A1").setValue("【担当者リスト】").setFontWeight("bold").setBackground("#d9ead3");
  const initialAssignees = [["本田 啓夫"], ["佐藤 料理長"], ["鈴木 買出"], ["AI アシスタント"]];
  sheetDropdowns.getRange(2, 1, initialAssignees.length, 1).setValues(initialAssignees);

  // --- 工程マスタ (Process_DB) ---
  // ★変更：項目名を日本語に
  const processHeaders = ["工程ID", "工程名", "説明"];
  const processData = [
    ["P-01", "買出し", "食材や備品の調達フェーズ"],
    ["P-02", "下準備", "食材のカット、下味付け"],
    ["P-03", "調理", "加熱調理プロセス"],
    ["P-04", "盛り付け", "提供前の最終仕上げ"]
  ];
  sheetProcessDB.getRange(1, 1, 1, processHeaders.length).setValues([processHeaders])
    .setFontWeight("bold").setBackground("#cfe2f3");
  sheetProcessDB.getRange(2, 1, processData.length, processData[0].length).setValues(processData);

  // --- タスク管理 (Task_DB) ---
  // ★変更：項目名を日本語に
  const fixedHeaders = [
    "工程ID", "タスクID", "工程名", "タスク名", 
    "担当者", "ステータス", "想定工数(h)", "開始日", "期限日", "通知"
  ];
  
  sheetTaskDB.getRange(1, 1, 1, fixedHeaders.length).setValues([fixedHeaders])
    .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
  
  // 計算用列 (K列) ※ヘッダー名変更
  sheetTaskDB.getRange("K1").setValue("グループID");
  sheetTaskDB.getRange("K2").setFormula('=ARRAYFORMULA(IF(A2:A="", "", MATCH(A2:A, UNIQUE(A2:A), 0)))');
  sheetTaskDB.hideColumns(11);

  // ガントチャート (L列以降)
  const today = new Date();
  const dateHeaders = [];
  for (let i = 0; i < 60; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() + i);
    dateHeaders.push(d);
  }
  sheetTaskDB.getRange(1, 12, 1, dateHeaders.length)
    .setValues([dateHeaders])
    .setNumberFormat("M/d")
    .setBackground("#f3f3f3")
    .setFontColor("black")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 列幅・固定
  sheetTaskDB.setColumnWidth(4, 250); // タスク名
  sheetTaskDB.setColumnWidths(12, 60, 25); // ガントチャート
  sheetTaskDB.setFrozenRows(1);
  sheetTaskDB.setFrozenColumns(4);

  // 入力規則
  const ruleAssignee = SpreadsheetApp.newDataValidation().requireValueInRange(sheetDropdowns.getRange("A2:A")).setAllowInvalid(true).build();
  sheetTaskDB.getRange("E2:E100").setDataValidation(ruleAssignee);
  const ruleStatus = SpreadsheetApp.newDataValidation().requireValueInList(["⚪️ 未着手", "🔵 進行中", "🟢 完了", "🟡 確認待ち"]).setAllowInvalid(true).build();
  sheetTaskDB.getRange("F2:F100").setDataValidation(ruleStatus);
  const ruleCheck = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheetTaskDB.getRange("J2:J100").setDataValidation(ruleCheck);

  // 数式 (C列: 工程名) ★シート名参照を日本語に変更
  // VLOOKUP(A2:A, '工程マスタ'!A:B, 2, FALSE)
  sheetTaskDB.getRange("C2").setFormula("=ARRAYFORMULA(IFERROR(VLOOKUP(A2:A, '工程マスタ'!A:B, 2, FALSE), \"\"))");

  // --- 条件付き書式 (4色分け) ---
  const rules = sheetTaskDB.getConditionalFormatRules();
  const rangeProcessCols = sheetTaskDB.getRange("A2:D100"); 

  // 1. 完了行グレーアウト
  const ruleGray = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="🟢 完了"')
    .setBackground("#EFEFEF").setFontColor("#999999")
    .setRanges([sheetTaskDB.getRange("A2:BM100")])
    .build();
  rules.push(ruleGray);

  // 2. プロセス4色分け (K列参照)
  const colors = ["#F4CCCC", "#D9EAD3", "#CFE2F3", "#FFF2CC"];
  colors.forEach((color, index) => {
    const remainder = (index + 1) % 4;
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=MOD($K2, 4) = ${remainder}`)
      .setBackground(color)
      .setRanges([rangeProcessCols])
      .build();
    rules.push(rule);
  });

  // 3. ガントチャートバー (L列以降)
  const ganttRange = sheetTaskDB.getRange(2, 12, 100, 60);
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

  // --- ダッシュボード (Dashboard) ---
  sheetDashboard.getRange("A1").setValue("【リソース負荷状況】(未完了タスクの工数合計)");
  // ★シート名参照を日本語に変更
  sheetDashboard.getRange("A2").setFormula("=QUERY('タスク管理'!E:G, \"select E, sum(G) where F != '🟢 完了' and E is not null group by E label sum(G) '残工数(h)'\", 1)");
  
  sheetDashboard.getRange("D1").setValue("【設定】Google Chat Webhook URL");
  sheetDashboard.getRange("D2").setBackground("#fff2cc");
  
  sheetDashboard.getRange("D7").setValue("【KPI】完了タスク数");
  sheetDashboard.getRange("D8").setFormula("=COUNTIF('タスク管理'!F:F, \"🟢 完了\")");
  sheetDashboard.getRange("D8").setFontColor("green").setFontWeight("bold").setFontSize(14);

  // 不要シート削除
  const allSheets = ss.getSheets();
  if (allSheets.length > 4) {
    allSheets.forEach(sheet => {
      if (!Object.values(targetSheets).includes(sheet.getName())) {
        try { ss.deleteSheet(sheet); } catch(e){}
      }
    });
  }
  
  ss.setActiveSheet(sheetTaskDB);
  Browser.msgBox("✅ 日本語版シートを作成しました");
}