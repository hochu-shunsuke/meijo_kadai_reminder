/**
 * Config.gs
 * システム全体で使用する共通定数。
 */

// --- WebClass関連 URL ---
const WEBCLASS_BASE_URL = 'https://rpwebcls.meijo-u.ac.jp';
const SSO_URL = 'https://slbsso.meijo-u.ac.jp/opensso/json/authenticate';
const ACS_URL = WEBCLASS_BASE_URL + '/simplesaml/module.php/saml/sp/saml2-acs.php/default-sp';

// --- スプレッドシート設定 ---
const SHEET_NAME_WEBCLASS = 'WebClass課題';
const SHEET_NAME_CLASSROOM = 'Classroom課題';
const SHEET_NAME_LOG = 'ログ';
const HEADER = ['ソース', '授業名', '課題タイトル', '開始日時', '終了日時', '課題リンク (URL)', 'Tasks ID', '登録済みフラグ'];

// HEADER の並びに対応する列番号（0始まり）。
// 列を増減するときは HEADER と COL の両方を必ず一緒に直すこと。
const COL = {
  SOURCE: 0,
  COURSE: 1,
  TITLE: 2,
  START: 3,
  END: 4,
  LINK: 5,    // 課題を一意に識別するキーとして使う
  TASK_ID: 6,
  FLAG: 7
};

// 登録済みフラグ（H列）が取りうる値
const FLAG = {
  REGISTERED: 'REGISTERED',        // Tasksへ登録済み
  COMPLETED: 'COMPLETED',          // Tasks側で完了された
  DELETED: 'DELETED',              // Tasks側で削除された
  EXPIRED: 'EXPIRED',              // 登録する前に期限が過ぎていた
  SKIPPED_NODATE: 'SKIPPED_NODATE' // 期限が読み取れなかった
};

// これらのフラグが付いた課題は、二度とTasksへ登録しない
const TERMINAL_FLAGS = [FLAG.COMPLETED, FLAG.DELETED, FLAG.EXPIRED, FLAG.SKIPPED_NODATE];

// --- システム定数 ---
const MAX_REDIRECTS = 15;

// ログシートに残す最大行数（ヘッダーを除く）。
// 1実行あたり30〜35行程度なので、5000行で約150実行分（1日2回なら約2か月半）。
// 2列しか使わないので1万セル程度。スプレッドシートの上限(1000万セル)から見れば誤差。
const MAX_LOG_ROWS = 5000;

// --- 正規表現 ---
const REGEX = {
  ID: /id=([a-f0-9]+)/,
  REDIRECT: /(?:window\.location\.href\s*=\s*|content\s*=\s*[\"']0;\s*URL=)['"]([^\"']+)[\"']/,
  SAML_RESPONSE: /<input type="hidden" name="SAMLResponse" value="([^"]+)"/,
  RELAY_STATE: /<input type="hidden" name="RelayState" value="([^"]+)"/,
  FORM_ACTION: /<form method="post" action="([^"]+)"/
};

// --- User-Agent リスト ---
const USER_AGENTS = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
  'Mozilla/5.0 (iPhone; CPU iPhone OS 17_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.6 Mobile/15E148 Safari/604.1'
];