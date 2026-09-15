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
 * ログやメール本文に埋め込む日時文字列。
 * ログシートのタイムスタンプ列とは別に、メッセージ本文にも日時を残したい場面で使う
 * （メッセージだけをコピーしたときに日時が失われないようにするため）。
 */
function formatStamp(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

/**
 * ログシートが無限に伸びないように、古い行を落として直近だけ残す。
 * 1回のdeleteRowsでまとめて消すので、行数が多くても呼び出しは1回。
 */
function trimLogSheet(maxRows) {
  const limit = Number(maxRows || MAX_LOG_ROWS);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_LOG);
    if (!sheet) return;

    const dataRows = sheet.getLastRow() - 1; // ヘッダーを除く
    if (dataRows <= limit) return;

    const removeCount = dataRows - limit;
    sheet.deleteRows(2, removeCount); // 2行目（最も古い行）から削除
    log(`🧹 古いログ${removeCount}行を削除 (直近${limit}行を保持)`);
  } catch (e) {
    console.error('ログ整理エラー: ' + e.message);
  }
}


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
    // 先頭が = + @ の文字列はスプレッドシートに数式として解釈され #ERROR! になる。
    // 先頭に ' を付けると強制的に文字列として扱われ、この ' はセルに表示されない。
    const safe = /^[=+@]/.test(String(message)) ? `'${message}` : message;
    sheet.appendRow([new Date(), safe]);
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
          log('✅ 復旧を確認。不具合タスクを削除しました。');
        } catch (e) {
          // 既にユーザーが消している場合。何もしなくてよい。
        }
        Settings.deleteSetting('healthTaskId');
      }
      log('✅ 異常なし');
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
        log(`⚠️ 不具合タスクを更新 (${this.problems.length}件)`);
        return;
      } catch (e) {
        // ユーザーが消していた → 作り直す
        Settings.deleteSetting('healthTaskId');
      }
    }

    try {
      const created = Tasks.Tasks.insert(task, listId);
      Settings.setSetting('healthTaskId', created.id);
      log(`⚠️ 不具合タスクをTasksに追加 (${this.problems.length}件)`);
    } catch (e) {
      log(`🚨 不具合タスクの作成に失敗: ${e.message}`);
    }
  }
};


/**
 * Google API の一時的な障害（503 "The service is currently unavailable" など）に備えて
 * 指数バックオフで再試行する。
 *
 * 注意: これを使うのは Google のAPIに対してのみ。
 * WebClass（大学のサーバ）へのリクエストは、無駄に負荷をかけないため再試行しない。
 */
function retryOnTransient(label, fn, attempts = 3) {
  for (let i = 0; i < attempts; i++) {
    try {
      return fn();
    } catch (e) {
      const msg = String((e && e.message) || e);
      const isTransient = /currently unavailable|internal error|backend error|try again|timed? ?out|rate limit|quota|\b50[0-9]\b/i.test(msg);

      if (!isTransient || i === attempts - 1) throw e;

      const waitMs = Math.pow(2, i) * 1000; // 1秒 → 2秒
      log(`  ・${label}: 一時的なエラー。${waitMs / 1000}秒後に再試行 (${i + 1}/${attempts - 1})`);
      Utilities.sleep(waitMs);
    }
  }
}


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
 * Classroom APIの dueDate / dueTime を Date に変換する。
 * どちらも仕様上UTCなので、UTCとして組み立てる。
 * （ローカル時刻として組み立てると、日本時間では9時間ずれる）
 */
function fromClassroomDue(dueDate, dueTime) {
  if (!dueDate) return null;
  const t = dueTime || {};
  const d = new Date(Date.UTC(dueDate.year, dueDate.month - 1, dueDate.day, t.hours || 0, t.minutes || 0));
  return isNaN(d.getTime()) ? null : d;
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
    log('[設定] すべての設定と自動実行トリガーを削除しました。');
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
        log(`[設定] 既存のTasksリスト「${listName}」を使用します。`);
        break;
      }
    }
  }
  
  // 2. 見つからなければ新規作成
  if (!targetId) {
    const newList = Tasks.Tasklists.insert({ title: listName });
    targetId = newList.id;
    log(`[設定] Tasksリスト「${listName}」を作成しました。`);
  }

  return targetId; 
}


// --- シート操作ヘルパー ---

/**
 * シート書き込み共通処理
 * ★修正: 既存データのTasks ID/Flagを保持したまま更新するように変更
 */
const SheetUtils = {
  writeToSheet: function(sheetName, newAssignments) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      log(`[設定] シート「${sheetName}」を作成しました。`);
    }

    // --- 既存行を課題リンクで引けるようにする ---
    const existingByLink = new Map();
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const currentData = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
      currentData.forEach(row => {
        const link = row[COL.LINK]; // 課題リンクを一意キーとして使う
        if (link) existingByLink.set(link, row);
      });
    }

    // --- 新しいデータに既存のTasks ID / フラグを引き継ぐ ---
    const newLinks = new Set();
    newAssignments.forEach(row => {
      const link = row[COL.LINK];
      newLinks.add(link);
      const prev = existingByLink.get(link);
      if (prev) {
        row[COL.TASK_ID] = prev[COL.TASK_ID]; // Tasks ID
        row[COL.FLAG] = prev[COL.FLAG]; // 登録済みフラグ
      }
    });

    // --- 今回取得できなかったが、既にTasksへ登録済み/処理済みの行は残す ---
    // 取得が一時的に失敗した際に行ごと消えると、Tasks IDを失って
    // 次回の実行で同じ課題が二重登録されるため。
    // 不要になった行は _cleanup() が期限経過後に削除する。
    const preserved = [];
    existingByLink.forEach((row, link) => {
      if (newLinks.has(link)) return;
      if (row[COL.TASK_ID] || row[COL.FLAG]) preserved.push(row);
    });

    const rows = newAssignments.concat(preserved);

    // --- シートをクリアして書き戻す ---
    if (lastRow > 1) {
      const lastColumn = sheet.getLastColumn();
      if (lastColumn > 0) {
        sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
      }
    }

    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, HEADER.length).setValues(rows);
    }
    SpreadsheetApp.flush();

    const suffix = preserved.length > 0 ? ` / 登録済みのため保持 ${preserved.length}件` : '';
    log(`  → ${newAssignments.length}件を「${sheetName}」へ書き込み${suffix}`);
  }
};
