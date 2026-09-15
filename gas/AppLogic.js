/**
 * AppLogic.gs
 * 課題データの取得、変換、Tasks連携のロジックを管理するコアファイル。
 * 依存: 
 * - Config.gs (定数)
 * - Utils.gs (Settings, SheetUtils, log, parseAssignmentDate)
 * - WebClassClient.gs (WebClassClientクラス)
 * - Parser.gs (WebClassParser)
 * - Tasks API サービス, Classroom API サービス
 */

/**
 * Tasksリスト再設定時、または強制リセット時に課題シートのデータをクリアする
 * (ヘッダー行とTasks ID/フラグだけでなく、課題全体をクリアし、次回全て再取得させる)
 */
function clearAssignmentSheets() {
  log('--- 課題シートの全データクリア開始 (Tasksリスト再設定のため) ---');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    // ヘッダー行 (1行目) を残して、2行目以降の全データをクリア
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // データが存在する場合のみクリア実行
    if (lastRow > 1 && lastCol > 0) {
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        log(`✅ シート「${name}」の全課題データをクリアしました。`);
    }
  });
  log('--- 課題シートの全データクリア完了 ---');
}


/**
 * WebClassから課題を取得し、シートに書き込む
 */
function processWebClass() {
  log('--- WebClass課題取得開始 ---');
  const u = Settings.getSetting('userid');
  const p = Settings.getSetting('password');
  
  if (!u || !p) {
    throw new Error('WebClass認証情報が未設定です。メニューから設定してください。');
  }

  const client = new WebClassClient();
  let dashUrl;
  try {
    dashUrl = client.login(u, p);
  } catch(e) {
    log(`🚨 ログイン失敗: ${e.message}`);
    throw new Error('WebClassへのログインに失敗しました。認証情報を確認してください。');
  }

  const dashHtml = client.fetchWithSession(dashUrl);
  const courses = WebClassParser.parseDashboard(dashHtml);
  if (courses.length === 0) {
    Health.add('WebClassのコースを1件も検出できませんでした。WebClass側のHTML構造が変わった可能性があります。');
  }

  const rows = [];
  courses.forEach(c => {
    let cName = c.name.replace(/^\s*\d+\s*/, '').replace(/\s*\(.*\)\s*$/, '').trim();
    try {
      const html = client.fetchWithSession(c.url);
      const asses = WebClassParser.parseCourseContents(html);
      
      asses.forEach(a => {
        // Tasks ID(6) と フラグ(7) は空でセット
        rows.push(['WebClass', cName, a.title, a.start, a.end, a.shareLink, '', '']);
      });
    } catch (e) {
      Health.add(`WebClass「${cName}」の課題取得に失敗: ${e.message}`);
    }
    Utilities.sleep(500); 
  });
  
  SheetUtils.writeToSheet(SHEET_NAME_WEBCLASS, rows);
  log('--- WebClass課題取得完了 ---');
}

/**
 * Classroomのコース一覧をページネーション込みで全件取得
 */
function _listAllClassroomCourses() {
  const out = [];
  let pageToken = null;
  do {
    const res = Classroom.Courses.list({
      courseStates: ['ACTIVE'],
      pageSize: 100,
      pageToken: pageToken || undefined
    });
    if (res.courses) out.push(...res.courses);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return out;
}

/**
 * 1コース分の課題をページネーション込みで全件取得
 */
function _listAllCourseWork(courseId) {
  const out = [];
  let pageToken = null;
  do {
    const res = Classroom.Courses.CourseWork.list(courseId, {
      courseWorkStates: ['PUBLISHED'],
      pageSize: 100,
      pageToken: pageToken || undefined
    });
    if (res.courseWork) out.push(...res.courseWork);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return out;
}

/**
 * Google Classroomから課題を取得し、シートに書き込む
 */
function processClassroom() {
  log('--- Classroom課題取得開始 ---');

  let courses;
  try {
    courses = retryOnTransient('Classroomコース一覧', () => _listAllClassroomCourses());
  } catch (e) {
    log(`🚨 Classroomコース一覧の取得に失敗: ${e.message}`);
    return;
  }
  log(`Classroomコースを${courses.length}件検出`);

  const rows = [];
  let failed = 0;

  // コース単位でtryを切る。1コースの失敗で全滅させない。
  courses.forEach(c => {
    try {
      const works = retryOnTransient(c.name, () => _listAllCourseWork(c.id));
      let dated = 0;

      works.forEach(w => {
        if (!w.dueDate) return;

        const d = w.dueDate;
        const t = w.dueTime || {};
        // dueDate/dueTime はAPI仕様上UTC。UTCとして組み立ててからスクリプトのTZで整形する。
        const dt = new Date(Date.UTC(d.year, d.month - 1, d.day, t.hours || 0, t.minutes || 0));
        const dueStr = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

        rows.push(['Classroom', c.name, w.title, '', dueStr, w.alternateLink, '', '']);
        dated++;
      });

      log(`  ✅ ${c.name}: 課題${works.length}件 / 期限付き${dated}件`);
    } catch (e) {
      failed++;
      Health.add(`Classroom「${c.name}」の課題取得に失敗: ${e.message}`);
    }
  });

  if (courses.length === 0) {
    Health.add('ClassroomのACTIVEなコースが0件でした。学期の切り替わりでアーカイブされた可能性があります。');
  } else if (failed === courses.length) {
    Health.add('Classroomの全コースで課題取得に失敗しました。OAuthスコープ不足が濃厚です (classroom.coursework.me.readonly)。');
  }

  // 全滅時に既存シートを空で上書きして消し飛ばさない
  if (rows.length === 0 && (failed > 0 || courses.length === 0)) {
    log('⚠️ 取得0件かつ失敗ありのため、シート上書きをスキップしました（既存データを保持）。');
    return;
  }

  SheetUtils.writeToSheet(SHEET_NAME_CLASSROOM, rows);
  log('--- Classroom課題取得完了 ---');
}

/**
 * スプレッドシートとTasksの同期処理
 */
function processTasksSync() {
  const listId = Settings.getTaskListId();
  if (!listId) {
    log('⚠️ TasksリストIDが未設定のため、同期・登録処理をスキップしました。');
    return;
  }
  
  log('--- Tasks同期処理開始 ---');

  // TasksリストIDの有効性チェック
  try {
    Tasks.Tasklists.get(listId); 
  } catch (e) {
    if (e.message.includes('Not Found') || e.message.includes('not found')) {
      log(`🚨 TasksリストID「${listId}」が見つかりませんでした。リストが削除された可能性があります。`);
      Settings.deleteTaskListId(); 
      log('✅ 無効なTasksリストIDを削除しました。Tasks連携を再開するには、メニューの「2. Tasks・自動実行設定を完了」から再設定してください。');
      return; 
    }
    log(`🚨 Tasks APIエラーにより同期中断: ${e.message}`);
    throw new Error(`Tasks APIエラーにより同期中断: ${e.message}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetDataMap = new Map();
  const allRows = [];

  // 各シートを読み込み、2シートを1つの配列に統合する。
  // 書き戻す先が分かるように、シート名と元の行番号を後ろに付けておく。
  // （Tasks ID / フラグの復元は SheetUtils.writeToSheet が取得時に済ませているので、ここでは不要）
  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length).getValues();
    sheetDataMap.set(name, { rows: rows, sheet: sheet, updated: false });

    rows.forEach((row, originalIndex) => allRows.push([...row, name, originalIndex]));
  });

  if (allRows.length === 0) {
    log('同期対象の課題が見つかりませんでした。');
    _cleanup(ss); 
    return;
  }
  
  // 4. 統合した全課題を、締切の遅い順にソートする (Tasksへの登録順を決定)
  allRows.sort((a, b) => {
    const dateA = parseAssignmentDate(a[4]); 
    const dateB = parseAssignmentDate(b[4]);

    const timeA = dateA ? dateA.getTime() : Infinity;
    const timeB = dateB ? dateB.getTime() : Infinity;

    return timeB - timeA; 
  });

  // 5. 締切の遅い順にTasksへの同期・登録処理を実行
  allRows.forEach(fullRow => {
    const [src, course, title, start, due, link, taskId, flag, sheetName, originalIndex] = fullRow;
    
    const sheetContext = sheetDataMap.get(sheetName);
    const originalRow = sheetContext.rows[originalIndex]; // これはマージ後のデータ

    
    // --- 課題の完了状態をTasksからシートへ同期（originalRowを操作） ---
    if (originalRow[COL.TASK_ID] && ![FLAG.COMPLETED, FLAG.DELETED].includes(originalRow[COL.FLAG])) {
      try {
        const taskStatus = Tasks.Tasks.get(listId, originalRow[COL.TASK_ID]).status;
        if (taskStatus === 'completed') {
          originalRow[COL.FLAG] = FLAG.COMPLETED; sheetContext.updated = true;
        }
      } catch(e) { 
        if(e.message.includes('NotFound')) { 
          originalRow[COL.FLAG] = FLAG.DELETED; sheetContext.updated = true; 
          log(`Tasksから削除された課題を検出: ${title}`);
        }
      }
    }

    // --- 新規課題をTasksに登録（originalRowを操作） ---
    // Tasks IDが空（まだ登録されていない）場合にのみ登録を試みる
    if (!originalRow[COL.TASK_ID] && !TERMINAL_FLAGS.includes(originalRow[COL.FLAG])) {
      
      let dueObj = parseAssignmentDate(due); 
      
      if (!dueObj) {
         originalRow[COL.FLAG] = FLAG.SKIPPED_NODATE; sheetContext.updated = true;
         return;
      }

      // 既に期限が過ぎているかチェック (1日余裕)
      if (dueObj.getTime() < new Date().getTime() - (24 * 3600 * 1000)) { 
        originalRow[COL.FLAG] = FLAG.EXPIRED; sheetContext.updated = true; 
        log(`期限切れの課題を検出: ${title}`);
        return;
      }

      try {
        const diff = (dueObj.getTime() - new Date().getTime()) / 86400000;
        const urgent = diff <= 3; 
        const dueDisp = Utilities.formatDate(dueObj, Session.getScriptTimeZone(), 'MM/dd(E) HH:mm');
        
        let taskDue = new Date(dueObj);
        
        const task = {
          title: `${urgent ? '🔥 ' : ''}[${course}] ${title} (${dueDisp}まで)`,
          due: taskDue.toISOString(),
          notes: `リンク:\n${link}\n\n期限: ${dueDisp}\nソース: ${src}`
        };
        
        const t = Tasks.Tasks.insert(task, listId);
        
        originalRow[COL.TASK_ID] = t.id; 
        originalRow[COL.FLAG] = FLAG.REGISTERED; 
        sheetContext.updated = true;
        log(`Tasks登録: ${task.title}`);
      } catch(e) {
        Health.add(`Tasksへの登録に失敗: ${title} - ${e.message}`);
      }
    }
  });

  // 6. 更新されたデータをソートし、元のシートに書き戻す
  sheetDataMap.forEach((context, name) => {
    if (context.updated) {
      // 課題を期限の早い順にソート（シート表示用）
      context.rows.sort((a, b) => {
        const dateA = parseAssignmentDate(a[4]);
        const dateB = parseAssignmentDate(b[4]);

        const timeA = dateA ? dateA.getTime() : Infinity;
        const timeB = dateB ? dateB.getTime() : Infinity;

        return timeA - timeB;
      });
      
      // シートに書き戻す
      context.sheet.getRange(2, 1, context.rows.length, context.rows[0].length).setValues(context.rows);
      SpreadsheetApp.flush();
    }
  });
  
  _cleanup(ss); 
  log('--- Tasks同期処理完了 ---');
}

/**
 * 期限切れ、完了済み、削除済みタスクをシートから削除（整理）
 */
function _cleanup(ss) {
  const days = Number(Settings.getSetting('cleanupDays') || 30);
  const thresh = days * 86400000; 
  const now = new Date().getTime();
  
  log(`--- シートクリーンアップ開始 (猶予期間: ${days}日) ---`);
  let removed = 0;

  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const rows = sheet.getDataRange().getValues();
    
    for (let i = rows.length - 1; i >= 1; i--) {
      const row = rows[i];
      const [,,,, due,,, flag] = row;

      const dueObj = parseAssignmentDate(due);
      const pastGrace = dueObj && (now - dueObj.getTime()) > thresh;

      let shouldDelete = false;

      // 期限から猶予期間が過ぎた行は、状態によらず削除する。
      // REGISTERED のまま残り続ける行（大学側から消えた課題など）もここで掃除される。
      // 期限が未来の行は、まだ扱う必要があるので残す。
      if (pastGrace) shouldDelete = true;

      // 終了済みなのに期限が読み取れない行は、いつ消してよいか判断できないので即削除する。
      if (TERMINAL_FLAGS.includes(flag) && (flag === FLAG.SKIPPED_NODATE || !dueObj)) {
        shouldDelete = true;
      }

      if (shouldDelete) {
        sheet.deleteRow(i + 1);
        removed++;
      }
    }
  });
  log(`--- シートクリーンアップ完了 (${removed}行を削除) ---`);
}
