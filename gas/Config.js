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

// --- システム定数 ---
const MAX_REDIRECTS = 15;

// --- API呼び出しのリトライ設定 ---
// Classroom/Tasks の Advanced Service は自動リトライしない。
// 503が1回返っただけでそのコースの課題が丸ごと落ちるので、こちらで粘る。
const RETRY_MAX_ATTEMPTS = 4;              // 初回 + 再試行3回
const RETRY_BASE_WAIT_MS = 1000;           // 1s -> 2s -> 4s
const CLASSROOM_COURSE_INTERVAL_MS = 300;  // コース間のウェイト（レート制限の緩和）

// 時間をおけば直る類のエラー。APIが返す英文メッセージで判定する。
// 権限不足やNot Foundは何度やっても同じなので、ここには入れない。
const TRANSIENT_ERROR_PATTERNS = [
  'currently unavailable',
  'temporarily unavailable',
  'Service unavailable',
  'Internal error',
  'Backend Error',
  'backendError',
  'Rate Limit Exceeded',
  'rateLimitExceeded',
  'userRateLimitExceeded',
  'Too many requests',
  'Deadline exceeded',
  'timed out'
];

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