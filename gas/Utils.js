/**
 * Utils.gs
 * システムの共通ヘルパー、設定管理、ログ、シート操作を担う。
 * * 依存: 
 * - Config.gs (定数)
 * - Main.gs (関数: runPostTasksSetup, 定数: SETUP_FUNCTION)
 * - Tasks API サービス
 */

// SETUP_FUNCTION は Main.gs で定義されています。（ここでは再宣言しない）


/**
 * HTMLから呼び出される認証情報保存関数
 */
function saveAuthDataFromHtml(userid, password) {
  if (!userid || !password) throw new Error('IDとパスワードが必要です。');
  Settings.saveAuth({ userid: String(userid), password: String(password) });
  return true; 
}

/**
 * HTMLから呼び出されるTasks設定保存関数
 */
function saveTasksDataFromHtml(settings) {
  const oldListId = Settings.getTaskListId(); 
  
  // 1. Tasksリストの検索・作成（IDの取得と保存）をまず実行
  const listId = setupTasksList(settings.taskListName); 
  Settings.setTaskListId(listId); 
  
  // 2. Tasks IDが変わった場合、シートの全課題データをクリア
  if (oldListId !== listId) {
      log('🚨 TasksリストIDが変更されました。新しい環境で再スタートするため、シートの全課題データをクリアし、次回の実行で全て再取得します。');
      // AppLogic.gs で定義された関数を呼び出す
      clearAssignmentSheets(); 
  }
  
  // 3. その他の設定を保存
  Settings.saveTasks(settings);

  // 4. トリガー設定（Main.gsの関数呼び出し）
  runPostTasksSetup(settings);
  return true;
}

/**
 * HTML表示用の設定値取得
 */
function getTasksSettingsForHtml() {
  return {
    taskListName: Settings.getSetting('taskListName') || '大学課題',
    triggerHour: Settings.getSetting('triggerHour') || '6,18',
    cleanupDays: Settings.getSetting('cleanupDays') || '30'
  };
}


// --- 共通ヘルパー関数 ---

/**
 * ログ記録
 */
function log(message) {
  Logger.log(message);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_LOG);
      sheet.appendRow(['タイムスタンプ', 'メッセージ']);
    }
    sheet.appendRow([new Date(), message]);
  } catch (e) {
    console.error('ログ記録エラー: ' + e.message);
  }
}

/**
 * 実行の健全性を集め、異常があればTasksに「不具合タスク」を1件だけ置く。
 *
 * 「壊れているのに黙って動き続ける」のを防ぐのが目的。
 * 普段からウィジェットで見ているTasksに出すので、わざわざ別の場所を見に行かなくていい。
 * タスクは常に1件だけ（毎日増えない）。復旧したら自動で消える。
 */
const Health = {
  problems: [],

  /** 異常を記録する。ログにも残る。 */
  add: function(message) {
    this.problems.push(message);
    log(`⚠️ [要確認] ${message}`);
  },

  reset: function() {
    this.problems = [];
  },

  /** 今日の日付（JST基準）をTasksのdue用RFC3339文字列にする */
  _todayDue: function() {
    const jst = new Date(Date.now() + 9 * 3600 * 1000);
    return new Date(Date.UTC(jst.getUTCFullYear(), jst.getUTCMonth(), jst.getUTCDate())).toISOString();
  },

  /**
   * 異常の有無をTasks上の1件のタスクに反映する。
   * 異常あり → 作成 or 更新 / 異常なし → 削除
   */
  syncToTasks: function() {
    const listId = Settings.getTaskListId();
    if (!listId) {
      if (this.problems.length > 0) log('🚨 TasksリストIDが未設定のため、不具合を通知できませんでした。');
      return;
    }

    const savedId = Settings.getSetting('healthTaskId');

    // --- 復旧した: 不具合タスクを消す ---
    if (this.problems.length === 0) {
      if (savedId) {
        try {
          Tasks.Tasks.remove(listId, savedId);
          log('✅ 復旧を確認したので、不具合タスクを削除しました。');
        } catch (e) {
          // 既にユーザーが消している場合。何もしなくてよい。
        }
        Settings.deleteSetting('healthTaskId');
      }
      log('✅ 健全性チェック: 異常なし');
      return;
    }

    // --- 異常あり: 1件だけ置く ---
    const task = {
      title: `⚠️ 課題の自動取得に問題あり (${this.problems.length}件)`,
      notes: [
        '課題が正しく取得できていません。',
        '',
        ...this.problems.map((p, i) => `${i + 1}. ${p}`),
        '',
        `最終確認: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm')}`,
        '※ 直れば、このタスクは次回の実行で自動的に消えます。'
      ].join('\n'),
      due: this._todayDue(),
      status: 'needsAction' // 完了にされていても、直るまでは再び表に出す
    };

    if (savedId) {
      try {
        Tasks.Tasks.patch(task, listId, savedId);
        log(`⚠️ 不具合タスクを更新しました (${this.problems.length}件)`);
        return;
      } catch (e) {
        // ユーザーが消していた → 作り直す
        Settings.deleteSetting('healthTaskId');
      }
    }

    try {
      const created = Tasks.Tasks.insert(task, listId);
      Settings.setSetting('healthTaskId', created.id);
      log(`⚠️ 不具合タスクをTasksに追加しました (${this.problems.length}件)`);
    } catch (e) {
      log(`🚨 不具合タスクの作成に失敗: ${e.message}`);
    }
  }
};


/**
 * 実行時刻の指定文字列を時のリストに変換する。
 * "6,18" / "6, 18" / "6" / 6 のいずれも受け付ける。
 * 不正な値は捨て、重複を除いて昇順にする。
 */
function parseTriggerHours(value) {
  const hours = String(value == null ? '' : value)
    .split(',')
    .map(v => String(v).trim())
    .filter(v => v !== '')            // Number('') が 0 になるので先に落とす
    .map(v => Number(v))
    .filter(n => Number.isInteger(n) && n >= 0 && n <= 23);

  const unique = Array.from(new Set(hours)).sort((a, b) => a - b);
  if (unique.length === 0) throw new Error('実行時間は0〜23の数字で指定してください（例: 6,18）。');
  return unique;
}


/**
 * 日付文字列をパースしてDateオブジェクトを返すヘルパー関数。
 */
function parseAssignmentDate(dateStr) {
  if (!dateStr) return null;
  const cleanStr = String(dateStr).trim().replace(/(\d{4})[\/年](\d{1,2})[\/月](\d{1,2})[\日]?/g, '$1/$2/$3');
  const date = new Date(cleanStr);
  return isNaN(date.getTime()) ? null : date;
}


// --- 設定管理オブジェクト ---

/**
 * 設定管理オブジェクト
 */
const Settings = {
  getSetting: function(key) {
    return PropertiesService.getUserProperties().getProperty(key);
  },

  setSetting: function(key, value) {
    PropertiesService.getUserProperties().setProperty(key, String(value));
  },

  deleteSetting: function(key) {
    PropertiesService.getUserProperties().deleteProperty(key);
  },
  
  saveAuth: function(data) {
    PropertiesService.getUserProperties().setProperties({
      'userid': data.userid,
      'password': data.password
    });
  },

  saveTasks: function(data) {
    PropertiesService.getUserProperties().setProperties({
      'taskListName': String(data.taskListName),
      'triggerHour': String(data.triggerHour),
      'cleanupDays': String(data.cleanupDays)
    });
  },

  getTaskListId: function() {
    return PropertiesService.getUserProperties().getProperty('taskListId');
  },

  setTaskListId: function(id) {
    PropertiesService.getUserProperties().setProperty('taskListId', id);
  },
  
  /**
   * TasksリストIDをPropertiesServiceから明示的に削除する
   */
  deleteTaskListId: function() {
    PropertiesService.getUserProperties().deleteProperty('taskListId');
  },
  
  /**
   * すべてのユーザープロパティと日次実行トリガーを削除する
   */
  resetAll: function() {
    PropertiesService.getUserProperties().deleteAllProperties();
    
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction() === SETUP_FUNCTION) {
          ScriptApp.deleteTrigger(t);
      }
    }
    log('すべての設定と自動実行トリガーを削除しました。');
  }
};


// --- Tasks連携ヘルパー ---

/**
 * Tasksリストの検索・作成
 */
function setupTasksList(listName) {
  const lists = Tasks.Tasklists.list().items;
  let targetId = null;
  
  // 1. 既存のリストを名前で検索
  if (lists) {
    for (const list of lists) {
      if (list.title === listName) {
        targetId = list.id;
        log(`既存のTasksリスト「${listName}」を再発見しました。`);
        break;
      }
    }
  }
  
  // 2. 見つからなければ新規作成
  if (!targetId) {
    const newList = Tasks.Tasklists.insert({ title: listName });
    targetId = newList.id;
    log(`Tasksリスト「${listName}」を新規作成しました。`);
  }

  return targetId; 
}


// --- シート操作ヘルパー ---

/**
 * シート書き込み共通処理
 * ★修正: 既存データのTasks ID/Flagを保持したまま更新するように変更
 */
const SheetUtils = {
  writeToSheet: function(sheetName, newAssignments) { // newAssignmentsはWebClass/Classroomから取得したデータ
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        log(`シート「${sheetName}」を新規作成しました。`);
    }

    // --- 既存のTasks IDとFlagを退避する処理を追加 ---
    const idMap = new Map();
    const lastRow = sheet.getLastRow();
    
    // データが既にある場合、Tasks ID等を退避
    if (lastRow > 1) {
      // 全データを取得 (Linkは5列目, TasksIDは6列目, Flagは7列目 ※0始まり)
      const currentData = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
      
      currentData.forEach(row => {
        const link = row[5]; // 一意のキーとして課題リンクを使用
        const taskId = row[6];
        const flag = row[7];
        
        // Tasks IDまたはフラグがある場合のみ記録
        if (link && (taskId || flag)) {
          idMap.set(link, { id: taskId, flag: flag });
        }
      });
    }

    // --- 新しいデータに既存情報をマージ ---
    newAssignments.forEach(row => {
      const link = row[5];
      if (idMap.has(link)) {
        const saved = idMap.get(link);
        row[6] = saved.id;   // Tasks IDを復元
        row[7] = saved.flag; // 登録済みフラグを復元
      }
    });

    // 1. シートをクリア (情報はマージ済みなので安全)
    if (lastRow > 1) {
      const lastColumn = sheet.getLastColumn();
      if (lastColumn > 0) {
          sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
      }
    }

    // 2. ヘッダーとデータを書き込み
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    
    if (newAssignments.length > 0) {
      sheet.getRange(2, 1, newAssignments.length, newAssignments[0].length).setValues(newAssignments);
    }
    SpreadsheetApp.flush();
    log(`✅ ${newAssignments.length}件を「${sheetName}」へ更新完了 (重複防止処理済み)`);
  }
};
