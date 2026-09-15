/**
 * Main.gs
 * エントリーポイントとメニュー機能
 */

const SETUP_FUNCTION = 'dailySystemRun';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✨ 課題自動取得システム')
    .addItem('1. WebClass 認証情報を設定 (ID/PW)', 'showAuthDialog')
    .addItem('2. Tasks・自動実行設定を完了 (初回のみ)', 'showTasksSetupDialog')
    .addSeparator()
    .addItem('3. 今すぐ実行（テスト）', SETUP_FUNCTION)
    .addSeparator()
    .addItem('課題の取りこぼしをチェック', 'checkMissingTasks')
    .addSeparator()
    .addItem('4. 設定をすべてリセット', 'resetAllSettings')
    .addToUi();
}

/**
 * 1. 認証情報設定ダイアログ表示
 */
function showAuthDialog() {
  // Settings.htmlを読み込みます
  const html = HtmlService.createHtmlOutputFromFile('Setting')
    .setWidth(450).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'WebClass 認証情報の設定');
}

/**
 * 2. Tasks・自動実行設定ダイアログ表示
 */
function showTasksSetupDialog() {
  // Setting_Tasks.htmlを読み込みます
  const html = HtmlService.createHtmlOutputFromFile('Setting_Tasks')
    .setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tasks・自動実行設定');
}

/**
 * Tasks設定保存後に呼ばれる処理（トリガー設定など）
 * Setting_Tasks.html から呼び出されます
 */
function runPostTasksSetup(settings) {
  const ui = SpreadsheetApp.getUi();
  try {
    // 1. Tasksリストの連携・作成はUtils.gsのsaveTasksDataFromHtml内で完了しています

    // 2. トリガー設定
    setupDailyTrigger(settings.triggerHour);

    const times = parseTriggerHours(settings.triggerHour).map(h => `${h}時台`).join('・');
    ui.alert(`✅ 設定完了\nTasksリスト「${settings.taskListName}」と連携し、毎日 ${times} の自動実行を設定しました。`);
  } catch (e) {
    log(`🚨 設定エラー: ${e.message}`);
    throw e; // HTML側にエラーを返す
  }
}

/**
 * 定期実行トリガーの設定。
 * 複数時刻に対応（例: "6,18" で1日2回）。
 * 毎回すべて作り直すので、何度実行しても重複しない。
 */
function setupDailyTrigger(hourSpec) {
  const hours = parseTriggerHours(hourSpec);

  // 既存の同ハンドラのトリガーを一旦すべて削除
  let removed = 0;
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === SETUP_FUNCTION) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }

  for (const h of hours) {
    ScriptApp.newTrigger(SETUP_FUNCTION).timeBased().everyDays(1).atHour(h).create();
  }

  log(`✅ 自動実行トリガーを再設定しました: 毎日 ${hours.join('時, ')}時台（旧トリガー${removed}件を削除）`);
}

/**
 * 4. 設定をすべてリセット (Utils.gsのresetAllを呼び出し)
 */
function resetAllSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🚨 設定リセットの確認',
    'WebClass認証情報、TasksリストID、自動実行トリガーをすべて削除します。よろしいですか？\n\n（この操作は元に戻せません）',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    try {
      Settings.resetAll(); // Utils.gsの関数を呼び出し
      ui.alert('✅ すべての設定とトリガーを削除しました。システムを再利用するには、再度メニュー1, 2を実行してください。');
    } catch (e) {
      ui.alert(`🚨 リセットエラー: ${e.message}`);
    }
  }
}

/**
 * 日次実行メイン関数
 */
function dailySystemRun() {
  const startedAt = new Date();
  log(`--- システム実行開始 (${formatStamp(startedAt)}) ---`);
  Health.reset();

  // 各段は独立して動かす。WebClassがコケてもClassroomは試す。
  try {
    processWebClass();
  } catch (e) {
    Health.add(`WebClassの取得に失敗: ${e.message}`);
  }

  try {
    processClassroom();
  } catch (e) {
    Health.add(`Classroomの取得に失敗: ${e.message}`);
  }

  try {
    processTasksSync();
  } catch (e) {
    Health.add(`Tasksへの同期に失敗: ${e.message}`);
  }

  const finishedAt = new Date();
  const elapsedSec = Math.round((finishedAt - startedAt) / 1000);
  log(`--- システム実行完了 (${formatStamp(finishedAt)} / 所要 ${elapsedSec}秒) ---`);

  Health.syncToTasks();
  trimLogSheet();
}