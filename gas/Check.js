/**
 * Check.gs
 * 課題の取りこぼしチェック。
 *
 * WebClass と Classroom から見えている「全項目」と、実際のGoogle Tasksの中身を
 * 突き合わせて、ToDoに入っていない項目とその理由を一覧にする。
 * 「ToDoに入るべきなのに入っていない」ものだけを ★要確認 として拾うのが目的。
 *
 * メニューから手動で実行する。自動実行からは呼ばれない。
 */
function checkMissingTasks() {
  log('===== 取りこぼしチェック 開始 =====');

  const listId = resolveTaskList();
  if (!listId) {
    log('🚨 Tasksリストを特定できないため、チェックできません。');
    return;
  }

  // --- 1. 現在ToDoにある課題のリンクを集める（完了済み・非表示も含める） ---
  const registeredLinks = new Set();
  let pageToken = null;
  do {
    const res = Tasks.Tasks.list(listId, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100,
      pageToken: pageToken || undefined
    });
    (res.items || []).forEach(t => {
      const m = String(t.notes || '').match(/https?:\/\/\S+/);
      if (m) registeredLinks.add(m[0]);
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  log(`ToDoに登録されている課題: ${registeredLinks.size}件`);

  // --- 2. 両方のソースから全項目を集める（期限の有無を問わず） ---
  const items = _collectAllItems();

  // --- 3. 突き合わせ ---
  const now = new Date().getTime();
  let inTodo = 0;
  let excluded = 0;
  const missing = [];

  let currentCourse = '';
  items.forEach(it => {
    if (it.course !== currentCourse) {
      currentCourse = it.course;
      log(`■ ${it.course} (${it.source})`);
    }

    if (registeredLinks.has(it.link)) {
      inTodo++;
      log(`    ✅ ${it.title}`);
      return;
    }

    // ToDoに無い。入らなくて正しいのか、漏れなのかを判定する。
    const due = it.due;
    let reason;

    if (!due) {
      reason = String(it.rawDue || '').trim() === ''
        ? '対象外: 期限が設定されていない'
        : `★要確認: 期限「${it.rawDue}」を解釈できない`;
    } else if (due.getTime() < now - 24 * 3600 * 1000) {
      reason = `対象外: 期限切れ (${Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')})`;
    } else {
      reason = `★要確認: 期限が有効なのにToDoに無い (${Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')})`;
    }

    if (reason.indexOf('★') === 0) missing.push(`${it.course} / ${it.title} — ${reason}`);
    else excluded++;

    log(`    ${reason.indexOf('★') === 0 ? '🚨' : '⏭'} ${it.title} — ${reason}`);
  });

  // --- 4. まとめ ---
  log('----------');
  log(`全項目 ${items.length}件 / ToDoにあり ${inTodo}件 / 対象外 ${excluded}件 / ★要確認 ${missing.length}件`);

  if (missing.length > 0) {
    log('🚨 ToDoに入るべきなのに入っていない項目:');
    missing.forEach(m => log(`    ${m}`));
  } else {
    log('✅ 取りこぼしはありません。');
  }
  log('===== 取りこぼしチェック 完了 =====');

  return { total: items.length, inTodo: inTodo, excluded: excluded, missing: missing };
}

/**
 * WebClassとClassroomから、期限の有無を問わず全項目を集める。
 * チェック専用。シートにもTasksにも一切書き込まない。
 */
function _collectAllItems() {
  const items = [];

  // --- WebClass ---
  const u = Settings.getSetting('userid');
  const p = Settings.getSetting('password');
  if (u && p) {
    try {
      const client = new WebClassClient();
      const dashHtml = client.fetchWithSession(client.login(u, p));
      WebClassParser.parseDashboard(dashHtml).forEach(c => {
        const cName = c.name.replace(/^\s*\d+\s*/, '').replace(/\s*\(.*\)\s*$/, '').trim();
        try {
          WebClassParser.parseCourseContents(client.fetchWithSession(c.url)).forEach(a => {
            items.push({
              source: 'WebClass', course: cName, title: a.title,
              link: a.shareLink, rawDue: a.end, due: parseAssignmentDate(a.end)
            });
          });
        } catch (e) {
          log(`⚠️ WebClass「${cName}」の取得に失敗: ${e.message}`);
        }
        Utilities.sleep(500);
      });
    } catch (e) {
      log(`🚨 WebClassの取得に失敗: ${e.message}`);
    }
  } else {
    log('⚠️ WebClass認証情報が未設定のため、WebClassはチェックできません。');
  }

  // --- Classroom（期限なしの課題も含めて集める） ---
  try {
    _listAllClassroomCourses().forEach(c => {
      try {
        _listAllCourseWork(c.id).forEach(w => {
          const due = fromClassroomDue(w.dueDate, w.dueTime);
          items.push({
            source: 'Classroom', course: c.name, title: w.title,
            link: w.alternateLink,
            rawDue: w.dueDate ? JSON.stringify(w.dueDate) : '',
            due: due
          });
        });
      } catch (e) {
        log(`⚠️ Classroom「${c.name}」の取得に失敗: ${e.message}`);
      }
    });
  } catch (e) {
    log(`🚨 Classroomの取得に失敗: ${e.message}`);
  }

  return items;
}
