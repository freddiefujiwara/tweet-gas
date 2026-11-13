/**
 * tweet-gas: X (Twitter) posting backend on Google Apps Script.
 *
 * - Posts to POST https://api.twitter.com/2/tweets with OAuth 1.0a user context.
 * - The four API keys live in Script Properties (never in code):
 *     X_CONSUMER_KEY, X_CONSUMER_SECRET, X_ACCESS_TOKEN, X_ACCESS_TOKEN_SECRET
 * - The web app (doPost) only accepts requests carrying a Google ID token
 *   whose verified email is ALLOWED_EMAIL. Optional Script Property
 *   GOOGLE_CLIENT_ID additionally pins the token audience.
 * - A comment inbox: "list_comments" collects comments on the
 *   owner's posts into an "inbox" sheet tab, and "reply_comment"
 *   answers one of them. Both use the normal user auth path below.
 * - A DM inbox: "list_dms" collects direct-message events into a
 *   "dm" sheet tab (with a separate "dm_state" tab for the
 *   read/replied markers), "mark_dm_read" advances the read
 *   marker, and "reply_dm" answers a conversation. Same user
 *   auth path as the comment inbox.
 * - One machine-to-machine exception: social-gas reads the X numbers
 *   (followers + recent post metrics) through the "sns_snapshot"
 *   action, authorized by the SYNC_TOKEN script property — the same
 *   sync-token pattern alarm-gas uses for the time-table sync. The
 *   X keys stay in this project; social-gas never holds them.
 */

export const API_BASE = 'https://api.twitter.com/2';
export const ALLOWED_EMAIL = 'zegenvs@gmail.com';
export const SCREEN_NAME = 'freddiefujiwara';
export const TOKENINFO_URL = 'https://oauth2.googleapis.com/tokeninfo';
export const LOG_SHEET_NAME = 'log';
export const LOG_DETAIL_MAX = 40000;

// --- Temporary share sessions (temporary access links) ---
// A logged-in owner can issue a URL token that grants the same access
// as the owner for SHARE_TTL_SECONDS without a Google account (for
// tweet, that means posting). Sessions are stored in a dedicated
// sheet (token hashes only, never raw tokens), which doubles as an
// audit log. The sheet is never exposed through the API.

export const SHARE_SHEET_NAME = 'share_sessions';
// One month, so the owner can hand the link over well in advance.
export const SHARE_TTL_SECONDS = 30 * 24 * 60 * 60;
export const SHARE_HEADERS = [
  'token_hash',
  'issued_at',
  'expires_at',
  'revoked_at',
  'access_count',
  'last_accessed_at',
];

// --- Environment separation (staging) ---
// Optional Script Properties; when unset, behavior is identical to the
// original production setup:
// - SPREADSHEET_ID: open this spreadsheet by ID instead of the container
//   spreadsheet, so a standalone staging project can use its own file.
// - DRY_RUN ('1' / 'true'): never call the X API. Posting returns a
//   synthetic result so the whole flow can be exercised safely.

function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function isDryRun_() {
  const value = PropertiesService.getScriptProperties().getProperty('DRY_RUN');
  return value === '1' || value === 'true';
}

function getSessionSheet(sheetName) {
  const ss = spreadsheet_();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(SHARE_HEADERS);
  }
  return sheet;
}

function hashShareToken(token) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token))
    .map((byte) => ((byte % 256) + 256) % 256)
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('');
}

function readSessionRows(sheetName) {
  const sheet = getSessionSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  return { sheet, values };
}

function shareCellIso(value) {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
}

function revokeActiveSessions(sheetName, nowIso) {
  const { sheet, values } = readSessionRows(sheetName);
  values.forEach((row, index) => {
    if (index === 0) {
      return;
    }
    const revokedAt = shareCellIso(row[3]);
    const expiresAt = shareCellIso(row[2]);
    if (!revokedAt && expiresAt && expiresAt > nowIso) {
      sheet.getRange(index + 1, 4).setValue(nowIso);
    }
  });
}

function revokeActiveShareSessions(nowIso) {
  revokeActiveSessions(SHARE_SHEET_NAME, nowIso);
}

function issueSession(sheetName, tokenKey) {
  const now = new Date();
  const nowIso = now.toISOString();
  const expiresIso = new Date(now.getTime() + SHARE_TTL_SECONDS * 1000).toISOString();
  revokeActiveSessions(sheetName, nowIso);
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  getSessionSheet(sheetName).appendRow([hashShareToken(token), nowIso, expiresIso, '', 0, '']);
  return {
    [tokenKey]: token,
    issued_at: nowIso,
    expires_at: expiresIso,
    ttl_seconds: SHARE_TTL_SECONDS,
  };
}

function issueShareSession() {
  return issueSession(SHARE_SHEET_NAME, 'share_token');
}

function revokeShareSession() {
  revokeActiveShareSessions(new Date().toISOString());
  return { revoked: true };
}

function validateSessionToken(sheetName, token) {
  if (!token) {
    return null;
  }
  const hash = hashShareToken(token);
  const { sheet, values } = readSessionRows(sheetName);
  const nowIso = new Date().toISOString();
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    if (row[0] !== hash) {
      continue;
    }
    if (shareCellIso(row[3])) {
      return null;
    }
    const expiresAt = shareCellIso(row[2]);
    if (!expiresAt || expiresAt <= nowIso) {
      return null;
    }
    sheet.getRange(index + 1, 5).setValue(Number(row[4] || 0) + 1);
    sheet.getRange(index + 1, 6).setValue(nowIso);
    return { expires_at: expiresAt };
  }
  return null;
}

function validateShareToken(token) {
  return validateSessionToken(SHARE_SHEET_NAME, token);
}

// --- Owner sessions (long-lived sign-in) ---
// Google ID tokens expire after about an hour, which used to drop the
// owner back at the sign-in gate. After a Google sign-in the frontend
// exchanges the ID token for one of these 30-day owner sessions (same
// storage scheme as share sessions, separate sheet) and uses it for
// all later calls. Issuing a new one revokes the previous ones, and
// the frontend revokes its session on logout.

export const OWNER_SHEET_NAME = 'owner_sessions';

function issueOwnerSession() {
  return issueSession(OWNER_SHEET_NAME, 'session_token');
}

function validateOwnerToken(token) {
  return validateSessionToken(OWNER_SHEET_NAME, token);
}

function revokeOwnerSession(token) {
  if (token) {
    const hash = hashShareToken(token);
    const { sheet, values } = readSessionRows(OWNER_SHEET_NAME);
    const nowIso = new Date().toISOString();
    values.forEach((row, index) => {
      if (index > 0 && row[0] === hash && !shareCellIso(row[3])) {
        sheet.getRange(index + 1, 4).setValue(nowIso);
      }
    });
  }
  return { revoked: true };
}

/**
 * Appends one row to the "log" sheet of the bound spreadsheet.
 * Logging must never break the main flow, so errors are swallowed.
 * detail is stored as a JSON string in a single cell.
 */
export function logToSheet(action, detail) {
  try {
    const ss = spreadsheet_();
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.appendRow(['time', 'action', 'detail']);
    }
    sheet.appendRow([new Date().toISOString(), action, JSON.stringify(detail)]);
  } catch (e) {
    // ignore logging failures
  }
}

/**
 * RFC 3986 percent-encoding as required by OAuth 1.0a.
 */
export function percentEncode(value) {
  return encodeURIComponent(String(value)).replace(/[!'()*]/g, (c) =>
    '%' + c.charCodeAt(0).toString(16).toUpperCase()
  );
}

/**
 * Builds the OAuth 1.0a signature base string.
 * params: array of [key, value] pairs (oauth_* plus query/form parameters).
 */
export function signatureBaseString(method, baseUrl, params) {
  const normalized = params
    .map(([k, v]) => [percentEncode(k), percentEncode(v)])
    .sort((a, b) => (a[0] === b[0] ? (a[1] < b[1] ? -1 : 1) : a[0] < b[0] ? -1 : 1))
    .map(([k, v]) => `${k}=${v}`)
    .join('&');
  return [
    method.toUpperCase(),
    percentEncode(baseUrl),
    percentEncode(normalized),
  ].join('&');
}

/**
 * HMAC-SHA1 signature via Apps Script Utilities.
 */
export function hmacSha1Base64(baseString, consumerSecret, tokenSecret) {
  const key = `${percentEncode(consumerSecret)}&${percentEncode(tokenSecret || '')}`;
  // There is no computeHmacSha1Signature in Apps Script; use the generic API.
  const bytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_1,
    baseString,
    key
  );
  return Utilities.base64Encode(bytes);
}

/**
 * Signs a request and returns the OAuth Authorization header value.
 * request: { method, baseUrl, query: {}, form: {} }
 * credentials: { consumerKey, consumerSecret, accessToken, accessTokenSecret }
 * overrides: { nonce, timestamp } for tests.
 */
export function buildAuthorizationHeader(request, credentials, overrides = {}) {
  const nonce =
    overrides.nonce || Utilities.getUuid().replace(/-/g, '');
  const timestamp =
    overrides.timestamp || Math.floor(Date.now() / 1000).toString();

  const oauthParams = [
    ['oauth_consumer_key', credentials.consumerKey],
    ['oauth_nonce', nonce],
    ['oauth_signature_method', 'HMAC-SHA1'],
    ['oauth_timestamp', timestamp],
    ['oauth_token', credentials.accessToken],
    ['oauth_version', '1.0'],
  ];

  const collected = [...oauthParams];
  Object.entries(request.query || {}).forEach(([k, v]) => collected.push([k, String(v)]));
  Object.entries(request.form || {}).forEach(([k, v]) => collected.push([k, String(v)]));

  const baseString = signatureBaseString(request.method, request.baseUrl, collected);
  const signature = hmacSha1Base64(
    baseString,
    credentials.consumerSecret,
    credentials.accessTokenSecret
  );

  const headerParams = [...oauthParams, ['oauth_signature', signature]];
  return (
    'OAuth ' +
    headerParams
      .map(([k, v]) => `${percentEncode(k)}="${percentEncode(v)}"`)
      .join(', ')
  );
}

/**
 * Reads the four X API keys from Script Properties.
 */
export function getCredentials() {
  const props = PropertiesService.getScriptProperties();
  const credentials = {
    consumerKey: props.getProperty('X_CONSUMER_KEY'),
    consumerSecret: props.getProperty('X_CONSUMER_SECRET'),
    accessToken: props.getProperty('X_ACCESS_TOKEN'),
    accessTokenSecret: props.getProperty('X_ACCESS_TOKEN_SECRET'),
  };
  const missing = Object.entries(credentials)
    .filter(([, v]) => !v)
    .map(([k]) => k);
  if (missing.length > 0) {
    throw new Error('Missing Script Properties: ' + missing.join(', '));
  }
  return credentials;
}

function fetchJson(url, options) {
  const response = UrlFetchApp.fetch(url, { ...options, muteHttpExceptions: true });
  const code = response.getResponseCode();
  const text = response.getContentText();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {
    // keep json as null; handled by caller via status code
  }
  // Log everything the API returned. The query string is stripped because
  // the tokeninfo URL carries the ID token itself (claims in the response
  // body are fine to log; the credential is not). DM endpoints are the one
  // exception: a DM response body must never be stored anywhere (the same
  // rule probeDmAccess follows), so only the status is logged for them.
  const bareUrl = url.split('?')[0];
  const isDmEndpoint = /\/(dm_events|dm_conversations)(\/|$)/.test(bareUrl);
  logToSheet('fetch', {
    method: (options.method || 'get').toUpperCase(),
    url: bareUrl,
    status: code,
    response: isDmEndpoint
      ? '(not logged: DM endpoint)'
      : text.length > LOG_DETAIL_MAX ? text.slice(0, LOG_DETAIL_MAX) + '…(truncated)' : text,
  });
  return { code, json, text };
}

// --- Weighted text length (the X 280-character rule) -------------------
// NOTE: up to afa8b79 this project only checked for empty text; the
// weighted-length validation the comment inbox needs did not exist
// here yet. It is implemented once in validateTweetText() below and
// shared by postTweet() and postReply(), so both posting paths apply
// exactly the same rule. Weights follow the X counting rules: most
// Latin / kana code points weigh 1, CJK and emoji weigh 2, and a URL
// weighs 23 regardless of its raw length.

export const MAX_WEIGHTED_LENGTH = 280;
export const URL_WEIGHTED_LENGTH = 23;

function codePointWeight_(codePoint) {
  if (
    codePoint <= 0x10ff ||
    (codePoint >= 0x2000 && codePoint <= 0x200d) ||
    (codePoint >= 0x2010 && codePoint <= 0x201f) ||
    (codePoint >= 0x2032 && codePoint <= 0x2037)
  ) {
    return 1;
  }
  return 2;
}

/**
 * The X weighted length of a text: URLs count as
 * URL_WEIGHTED_LENGTH each, every other code point as 1 or 2.
 */
export function weightedTweetLength(text) {
  const raw = String(text == null ? '' : text);
  let length = 0;
  let rest = raw;
  const urlPattern = /https?:\/\/[^\s]+/g;
  let match = urlPattern.exec(raw);
  while (match) {
    length += URL_WEIGHTED_LENGTH;
    rest = rest.replace(match[0], '');
    match = urlPattern.exec(raw);
  }
  for (const ch of rest) {
    length += codePointWeight_(ch.codePointAt(0));
  }
  return length;
}

/**
 * Shared text validation for posts and replies: non-empty and
 * within the weighted 280 limit. Throws on invalid text.
 */
export function validateTweetText(text) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    throw new Error('text is empty');
  }
  const length = weightedTweetLength(text);
  if (length > MAX_WEIGHTED_LENGTH) {
    throw new Error(
      `text is too long: weighted length ${length} exceeds ${MAX_WEIGHTED_LENGTH}`
    );
  }
  return text;
}

/**
 * Posts a tweet. Returns { id, text, url }.
 */
export function postTweet(text) {
  validateTweetText(text);
  if (isDryRun_()) {
    return { id: 'dry-run', text, url: null, dry_run: true };
  }
  const credentials = getCredentials();
  const url = `${API_BASE}/tweets`;
  const authorization = buildAuthorizationHeader(
    { method: 'POST', baseUrl: url },
    credentials
  );
  const { code, json, text: body } = fetchJson(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: authorization },
    payload: JSON.stringify({ text }),
  });
  if (code < 200 || code >= 300 || !json || !json.data || !json.data.id) {
    throw new Error(`X API post failed (${code}): ${body}`);
  }
  return {
    id: json.data.id,
    text: json.data.text,
    url: `https://x.com/${SCREEN_NAME}/status/${json.data.id}`,
  };
}

/**
 * Fetches a tweet by id for verification. Returns { id, text } or null.
 */
export function getTweet(id) {
  const credentials = getCredentials();
  const url = `${API_BASE}/tweets/${encodeURIComponent(id)}`;
  const authorization = buildAuthorizationHeader(
    { method: 'GET', baseUrl: url },
    credentials
  );
  const { code, json } = fetchJson(url, {
    method: 'get',
    headers: { Authorization: authorization },
  });
  if (code < 200 || code >= 300 || !json || !json.data) {
    return null;
  }
  return { id: json.data.id, text: json.data.text };
}

// --- Machine sync token (GAS-to-GAS) ---
// The same pattern alarm-gas uses for the time-table sync: a single
// long-lived token for machine-to-machine calls, kept only in the
// SYNC_TOKEN script property, never in code. Here the token is
// scoped to exactly one read-only action, "sns_snapshot": the
// caller (social-gas) holds the same value as its
// TWEET_GAS_SYNC_TOKEN property and POSTs it with the action, and
// gets the X numbers below back. The token cannot post, issue
// share links, or do anything else.

export const SYNC_ACTIONS = ['sns_snapshot'];

export function validateSyncToken(token) {
  if (!token) {
    return false;
  }
  const expected = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  return !!expected && String(token) === expected;
}

// How many of the owner's newest posts the SNS snapshot carries.
export const SNS_SNAPSHOT_POSTS_LIMIT = 50;

/**
 * Compact, secret-free summary of an X API error response. Only
 * the status and the response body's own message fields are used,
 * so no credential can leak into a log or a caller's snapshot.
 */
export function summarizeXApiError(code, text) {
  let summary = `HTTP ${code}`;
  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    parsed = null;
  }
  const errors = parsed && Array.isArray(parsed.errors) ? parsed.errors : null;
  if (errors && errors.length) {
    const first = errors[0] || {};
    if (first.code !== undefined && first.code !== null) {
      summary += ` code=${first.code}`;
    }
    if (first.message) {
      summary += `: ${String(first.message).slice(0, 200)}`;
    }
  } else if (parsed && parsed.detail) {
    summary += `: ${String(parsed.detail).slice(0, 200)}`;
  } else if (parsed && parsed.title) {
    summary += `: ${String(parsed.title).slice(0, 200)}`;
  }
  return summary;
}

/**
 * One signed GET against the X API with query parameters. The
 * query values join the OAuth signature (buildAuthorizationHeader
 * takes them as request.query) and the request URL.
 * Returns { code, json, text } like fetchJson.
 */
export function signedGet(path, query, credentials) {
  const baseUrl = `${API_BASE}${path}`;
  const queryString = Object.entries(query)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`)
    .join('&');
  const url = queryString ? `${baseUrl}?${queryString}` : baseUrl;
  const authorization = buildAuthorizationHeader(
    { method: 'GET', baseUrl, query },
    credentials
  );
  return fetchJson(url, { method: 'get', headers: { Authorization: authorization } });
}

// A metric value as a number, or null when the API did not report
// one (absent and non-numeric both read as null, never as 0).
const xMetricOrNull = (value) => {
  if (value === null || value === undefined) {
    return null;
  }
  const num = Number(value);
  return Number.isFinite(num) ? num : null;
};

/**
 * One post entry of the SNS snapshot, from one tweet object of
 * GET /2/users/:id/tweets. Counts the API did not report stay null.
 */
export function buildSnsSnapshotPost(tweet) {
  const metrics = (tweet && tweet.public_metrics) || {};
  return {
    id: tweet && tweet.id != null ? String(tweet.id) : null,
    posted_at: (tweet && tweet.created_at) || null,
    text: tweet && tweet.text != null ? String(tweet.text) : '',
    views: xMetricOrNull(metrics.impression_count),
    likes: xMetricOrNull(metrics.like_count),
    reposts: xMetricOrNull(metrics.retweet_count),
    replies: xMetricOrNull(metrics.reply_count),
  };
}

/**
 * The X numbers social-gas collects hourly: the owner's follower
 * count and newest posts with their public metrics, read with the
 * same OAuth 1.0a credentials this project posts with — the X keys
 * stay in this project only. The two halves fail independently:
 * a failed half reports nulls plus a summary in x_error /
 * posts_error, and the snapshot as a whole still returns ok.
 * Under DRY_RUN nothing is fetched and every value is null.
 */
export function buildSnsSnapshot() {
  const snapshot = {
    ok: true,
    action: 'sns_snapshot',
    collected_at: new Date().toISOString(),
    x: { followers: null, username: null, recent_posts: null },
    x_error: null,
    posts_error: null,
  };
  if (isDryRun_()) {
    snapshot.dry_run = true;
    return snapshot;
  }
  let credentials;
  try {
    credentials = getCredentials();
  } catch (err) {
    snapshot.x_error = String(err && err.message ? err.message : err);
    snapshot.posts_error = snapshot.x_error;
    return snapshot;
  }
  const me = signedGet('/users/me', { 'user.fields': 'public_metrics' }, credentials);
  const user = me.json && me.json.data ? me.json.data : null;
  if (me.code < 200 || me.code >= 300 || !user || user.id == null) {
    snapshot.x_error = `x profile: ${summarizeXApiError(me.code, me.text)}`;
    snapshot.posts_error = snapshot.x_error;
    return snapshot;
  }
  snapshot.x.username = user.username != null ? String(user.username) : null;
  snapshot.x.followers = xMetricOrNull(
    user.public_metrics ? user.public_metrics.followers_count : null
  );
  const list = signedGet(
    `/users/${encodeURIComponent(String(user.id))}/tweets`,
    {
      max_results: String(SNS_SNAPSHOT_POSTS_LIMIT),
      'tweet.fields': 'created_at,public_metrics',
      exclude: 'retweets,replies',
    },
    credentials
  );
  const items = list.json && Array.isArray(list.json.data) ? list.json.data : null;
  if (list.code < 200 || list.code >= 300 || !items) {
    snapshot.posts_error = `x posts: ${summarizeXApiError(list.code, list.text)}`;
    return snapshot;
  }
  snapshot.x.recent_posts = items.map(buildSnsSnapshotPost);
  return snapshot;
}

// --- Comment inbox ------------------------------------------------------
// Two user-authenticated actions (the same Google ID token / share
// session path as posting; the sync token can NOT reach them — it is
// scoped to sns_snapshot only):
// - "list_comments": collect comments on the owner's posts from the
//   mentions timeline, a per-conversation recent search, and the
//   quote tweets of the owner's quoted posts, store them in
//   the "inbox" sheet, and return them newest first. Every comment
//   carries a "kind" (mention / reply / quote; quote wins whenever a
//   tweet quotes one of the owner's posts).
// - "reply_comment": post a reply to one stored comment and mark
//   the inbox row as replied. With as_quote it quotes the comment
//   instead of replying to it.
// Authorization caveat: mentions and especially the recent-search
// endpoint are not available on every X API plan/tier, and an app
// without the right access gets 401/403 for them. list_comments
// therefore treats search and quote failures as degraded results
// (the other sources, plus a note) instead of failing the action.

export const INBOX_SHEET_NAME = 'inbox';
export const INBOX_HEADERS = [
  'comment_id',
  'author_id',
  'author_username',
  'author_name',
  'text',
  'created_at',
  'in_reply_to_tweet_id',
  'conversation_id',
  'status',
  'our_reply_id',
  'first_seen_at',
  'last_seen_at',
  'kind',
];
// 1-based column numbers in the inbox sheet, matching INBOX_HEADERS.
const INBOX_COL_STATUS = 9;
const INBOX_COL_OUR_REPLY_ID = 10;
const INBOX_COL_LAST_SEEN_AT = 12;
const INBOX_COL_KIND = 13;

// The three comment kinds stored in the inbox "kind" column.
export const INBOX_KINDS = ['mention', 'reply', 'quote'];
// Priority when the same tweet is seen through several sources (or
// a stored row already carries a kind): quote beats reply beats
// mention, so a kind once recorded is never downgraded or erased
// by a later collection.
const INBOX_KIND_PRIORITY_ = { mention: 1, reply: 2, quote: 3 };

export const INBOX_MENTIONS_MAX_RESULTS = 100;
// The conversation search costs one search request per own post,
// and recent search is the most quota-sensitive endpoint here, so
// the targets are capped at the newest posts of the last 7 days:
// at most INBOX_SEARCH_POSTS_LIMIT search calls per list_comments
// run, which keeps a polling frontend from draining the quota while
// still covering the window where new comments realistically land.
export const INBOX_SEARCH_WINDOW_DAYS = 7;
export const INBOX_SEARCH_POSTS_LIMIT = 10;
export const INBOX_SEARCH_MAX_RESULTS = 100;
// Third source: quote tweets. The owner's posts come from the one
// shared fetch (at most 100 of them, newest first); among them,
// only those whose public_metrics report quote_count > 0 are
// queried, and the quote_tweets calls themselves are capped at
// INBOX_QUOTE_CALLS_LIMIT per list_comments run, so the quote
// source stays bounded like the conversation search. There is no
// age window: a quoted post keeps surfacing its quotes no matter
// how old it is (quotes are rare, so the call cap is enough).
export const INBOX_QUOTE_CALLS_LIMIT = 10;
export const INBOX_QUOTE_MAX_RESULTS = 100;

const INBOX_TWEET_FIELDS = 'created_at,conversation_id,in_reply_to_user_id,referenced_tweets';

function getInboxSheet() {
  const ss = spreadsheet_();
  let sheet = ss.getSheetByName(INBOX_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INBOX_SHEET_NAME);
    sheet.appendRow(INBOX_HEADERS);
    return sheet;
  }
  // Migrate a pre-kind sheet in place: rows written before the
  // "kind" column existed simply have no value in it, which the
  // rest of the inbox treats as the legacy (mention/reply) kind.
  try {
    const values = sheet.getDataRange().getValues();
    const header = values && values.length > 0 ? values[0] : null;
    if (header && header.length < INBOX_HEADERS.length) {
      for (let col = header.length + 1; col <= INBOX_HEADERS.length; col += 1) {
        sheet.getRange(1, col).setValue(INBOX_HEADERS[col - 1]);
      }
    }
  } catch (e) {
    // header migration must never break collection
  }
  return sheet;
}

/**
 * The kind of one inbox comment. A referenced_tweets entry of
 * type "quoted" that points at one of the owner's own posts makes
 * it a quote, and that wins over every other signal (including a
 * "replied_to" reference on the same tweet and the source the
 * tweet was collected through). Without a quote reference the
 * source kind wins when one is given, then a "replied_to"
 * reference makes it a reply, and anything else is a mention.
 * ownPostIds may be an array/Set of the owner's post ids, or null
 * when they are unknown (then any "quoted" reference counts).
 */
export function determineInboxKind(tweet, ownPostIds, sourceKind) {
  const referenced =
    tweet && Array.isArray(tweet.referenced_tweets) ? tweet.referenced_tweets : [];
  const ownSet =
    ownPostIds == null
      ? null
      : new Set((Array.isArray(ownPostIds) ? ownPostIds : [...ownPostIds]).map(String));
  const quotesOwnPost = referenced.some(
    (ref) =>
      ref &&
      ref.type === 'quoted' &&
      ref.id != null &&
      (ownSet === null || ownSet.has(String(ref.id)))
  );
  if (quotesOwnPost) {
    return 'quote';
  }
  if (sourceKind && INBOX_KINDS.includes(sourceKind)) {
    return sourceKind;
  }
  if (referenced.some((ref) => ref && ref.type === 'replied_to')) {
    return 'reply';
  }
  return 'mention';
}

// The higher-priority kind of two (possibly empty) kinds, or ''
// when neither is a known kind. Used so a stored kind is never
// erased or downgraded by a re-collection (see INBOX_KIND_PRIORITY_).
function mergeInboxKind_(existingKind, newKind) {
  const existing = INBOX_KINDS.includes(existingKind) ? existingKind : '';
  const incoming = INBOX_KINDS.includes(newKind) ? newKind : '';
  if (!existing) {
    return incoming;
  }
  if (!incoming) {
    return existing;
  }
  return INBOX_KIND_PRIORITY_[incoming] > INBOX_KIND_PRIORITY_[existing]
    ? incoming
    : existing;
}

// The kind a comment is reported with: its recorded kind, or —
// for a legacy row whose kind cell is still empty — the kind it
// would have carried before the column existed (a reply when it
// replies to a tweet, a mention otherwise).
function publicInboxKind_(comment) {
  if (comment && INBOX_KINDS.includes(comment.kind)) {
    return comment.kind;
  }
  return comment && comment.in_reply_to_tweet_id ? 'reply' : 'mention';
}

function usersByIdFromIncludes_(json) {
  const users = {};
  const list =
    json && json.includes && Array.isArray(json.includes.users) ? json.includes.users : [];
  list.forEach((user) => {
    if (user && user.id != null) {
      users[String(user.id)] = {
        username: user.username != null ? String(user.username) : '',
        name: user.name != null ? String(user.name) : '',
      };
    }
  });
  return users;
}

/**
 * Normalizes one raw X tweet object into an inbox comment, or
 * returns null for items the inbox must not carry: the owner's own
 * posts and reposts (referenced_tweets with type "retweeted", plus
 * the legacy "RT @" text prefix when references are absent).
 * A tweet with no text at all still counts when it reaches here
 * through the quote source: the quote itself is the comment.
 * options: { ownPostIds, sourceKind } — see determineInboxKind.
 */
export function normalizeInboxComment(tweet, usersById, ownId, options = {}) {
  if (!tweet || tweet.id == null) {
    return null;
  }
  const authorId = tweet.author_id != null ? String(tweet.author_id) : '';
  if (authorId && String(ownId) === authorId) {
    return null;
  }
  const referenced = Array.isArray(tweet.referenced_tweets) ? tweet.referenced_tweets : [];
  if (referenced.some((ref) => ref && ref.type === 'retweeted')) {
    return null;
  }
  const text = tweet.text != null ? String(tweet.text) : '';
  if (referenced.length === 0 && text.startsWith('RT @')) {
    return null;
  }
  const repliedTo = referenced.find((ref) => ref && ref.type === 'replied_to');
  const author = (usersById && usersById[authorId]) || {};
  return {
    comment_id: String(tweet.id),
    author_id: authorId,
    author_username: author.username || '',
    author_name: author.name || '',
    text,
    created_at: tweet.created_at || '',
    in_reply_to_tweet_id: repliedTo && repliedTo.id != null ? String(repliedTo.id) : '',
    conversation_id: tweet.conversation_id != null ? String(tweet.conversation_id) : '',
    kind: determineInboxKind(tweet, options.ownPostIds, options.sourceKind),
  };
}

function normalizeInboxComments_(json, ownId, options = {}) {
  const usersById = usersByIdFromIncludes_(json);
  const items = json && Array.isArray(json.data) ? json.data : [];
  return items
    .map((tweet) => normalizeInboxComment(tweet, usersById, ownId, options))
    .filter((comment) => comment !== null);
}

/**
 * Chooses the owner's posts whose conversations get a reply
 * search: newest first, only those created within the last
 * INBOX_SEARCH_WINDOW_DAYS days, capped at
 * INBOX_SEARCH_POSTS_LIMIT (see the constant's comment for why).
 */
export function selectConversationSearchTargets(posts, nowMs) {
  const cutoff = nowMs - INBOX_SEARCH_WINDOW_DAYS * 24 * 60 * 60 * 1000;
  return (Array.isArray(posts) ? posts : [])
    .filter((post) => post && post.id != null && post.created_at)
    .map((post) => ({ id: String(post.id), createdMs: Date.parse(post.created_at) }))
    .filter((post) => Number.isFinite(post.createdMs) && post.createdMs >= cutoff)
    .sort((a, b) => b.createdMs - a.createdMs)
    .slice(0, INBOX_SEARCH_POSTS_LIMIT)
    .map((post) => post.id);
}

/**
 * Chooses the owner's posts whose quote tweets get fetched:
 * among the fetched posts (the one shared fetch, at most 100),
 * newest first, only posts whose public_metrics report
 * quote_count > 0 are kept — however old they are — and the
 * result is capped at INBOX_QUOTE_CALLS_LIMIT so a list_comments
 * run never makes more quote_tweets calls than that.
 */
export function selectQuoteTargets(posts) {
  return (Array.isArray(posts) ? posts : [])
    .filter((post) => post && post.id != null && post.created_at)
    .map((post) => ({
      id: String(post.id),
      createdMs: Date.parse(post.created_at),
      quoteCount: xMetricOrNull(post.public_metrics ? post.public_metrics.quote_count : null),
    }))
    .filter((post) => Number.isFinite(post.createdMs))
    .sort((a, b) => b.createdMs - a.createdMs)
    .filter((post) => post.quoteCount != null && post.quoteCount > 0)
    .slice(0, INBOX_QUOTE_CALLS_LIMIT)
    .map((post) => post.id);
}

/**
 * Upserts comments into the inbox sheet and returns them with
 * their stored status / our_reply_id merged in. Existing rows
 * keep status, our_reply_id and first_seen_at untouched; only the
 * descriptive fields and last_seen_at are refreshed, so collecting
 * again can never un-reply a comment that was already answered.
 * Storage failures are swallowed like logToSheet failures (the
 * caller still gets the freshly collected comments, all "new").
 */
export function upsertInboxComments(comments) {
  const nowIso = new Date().toISOString();
  let existingById = new Map();
  let sheet = null;
  try {
    sheet = getInboxSheet();
    const values = sheet.getDataRange().getValues();
    values.forEach((row, index) => {
      if (index > 0 && row[0] != null && row[0] !== '') {
        existingById.set(String(row[0]), { row, rowNumber: index + 1 });
      }
    });
  } catch (e) {
    sheet = null;
    existingById = new Map();
  }
  return (Array.isArray(comments) ? comments : []).map((comment) => {
    const existing = existingById.get(comment.comment_id);
    // The kind is merged by priority and never taken from a single
    // side alone: a stored "quote" survives a later mention/reply
    // re-collection, a legacy row (kind still empty) adopts the
    // freshly determined kind, and a kind-less incoming comment
    // (direct callers of this function) falls back to the kind its
    // reply reference implies.
    const incomingKind = INBOX_KINDS.includes(comment.kind)
      ? comment.kind
      : comment.in_reply_to_tweet_id
        ? 'reply'
        : 'mention';
    const merged = {
      ...comment,
      status: existing ? String(existing.row[INBOX_COL_STATUS - 1] || 'new') : 'new',
      our_reply_id: existing ? String(existing.row[INBOX_COL_OUR_REPLY_ID - 1] || '') : '',
      kind: mergeInboxKind_(
        existing ? String(existing.row[INBOX_COL_KIND - 1] || '') : '',
        incomingKind
      ),
    };
    if (!sheet) {
      return merged;
    }
    try {
      if (existing) {
        const refresh = [
          [2, comment.author_id],
          [3, comment.author_username],
          [4, comment.author_name],
          [5, comment.text],
          [6, comment.created_at],
          [7, comment.in_reply_to_tweet_id],
          [8, comment.conversation_id],
          [INBOX_COL_LAST_SEEN_AT, nowIso],
          [INBOX_COL_KIND, merged.kind],
        ];
        refresh.forEach(([col, value]) => {
          sheet.getRange(existing.rowNumber, col).setValue(value);
        });
      } else {
        sheet.appendRow([
          comment.comment_id,
          comment.author_id,
          comment.author_username,
          comment.author_name,
          comment.text,
          comment.created_at,
          comment.in_reply_to_tweet_id,
          comment.conversation_id,
          'new',
          '',
          nowIso,
          nowIso,
          merged.kind,
        ]);
      }
    } catch (e) {
      // one bad row must not lose the rest of the batch
    }
    return merged;
  });
}

function publicInboxComment_(comment) {
  return {
    comment_id: comment.comment_id,
    author_id: comment.author_id,
    author_username: comment.author_username,
    author_name: comment.author_name,
    text: comment.text,
    created_at: comment.created_at,
    in_reply_to_tweet_id: comment.in_reply_to_tweet_id,
    conversation_id: comment.conversation_id,
    status: comment.status,
    our_reply_id: comment.our_reply_id,
    kind: publicInboxKind_(comment),
  };
}

/**
 * Backs the "list_comments" action. Returns
 * { ok: true, comments: [...newest first], note? }.
 * Mentions failure throws (the action then reports ok:false via
 * doPost); search and quote failures only degrade with a note.
 * Under DRY_RUN the X API is never called.
 */
export function listComments() {
  if (isDryRun_()) {
    return { ok: true, dry_run: true, comments: [] };
  }
  const credentials = getCredentials();
  const me = signedGet('/users/me', {}, credentials);
  const meData = me.json && me.json.data ? me.json.data : null;
  if (me.code < 200 || me.code >= 300 || !meData || meData.id == null) {
    throw new Error(`X API profile failed: ${summarizeXApiError(me.code, me.text)}`);
  }
  const ownId = String(meData.id);

  // The owner's own posts are fetched first (one call shared by
  // the conversation search and the quote source): their ids are
  // what lets a "quoted" reference be recognized as quoting the
  // owner, and their public_metrics carry the quote counts. A
  // failure here only degrades the two per-post sources below —
  // the mentions timeline further down stands on its own.
  const ownPosts = signedGet(
    `/users/${encodeURIComponent(ownId)}/tweets`,
    {
      max_results: '100',
      exclude: 'retweets,replies',
      'tweet.fields': 'created_at,public_metrics',
    },
    credentials
  );
  const ownItems =
    ownPosts.code >= 200 && ownPosts.code < 300 && ownPosts.json && Array.isArray(ownPosts.json.data)
      ? ownPosts.json.data
      : null;
  const ownPostIds = ownItems ? ownItems.map((post) => String(post.id)) : null;

  const mentions = signedGet(
    `/users/${encodeURIComponent(ownId)}/mentions`,
    {
      max_results: String(INBOX_MENTIONS_MAX_RESULTS),
      expansions: 'author_id',
      'tweet.fields': INBOX_TWEET_FIELDS,
    },
    credentials
  );
  if (mentions.code < 200 || mentions.code >= 300 || !mentions.json) {
    throw new Error(`X API mentions failed: ${summarizeXApiError(mentions.code, mentions.text)}`);
  }
  const byId = new Map();
  // Merges one collected comment into the map: a tweet seen
  // through several sources is kept once, and its kind is the
  // highest-priority kind any source determined for it.
  const addComment = (comment) => {
    const existing = byId.get(comment.comment_id);
    if (!existing) {
      byId.set(comment.comment_id, comment);
      return;
    }
    const kind = mergeInboxKind_(existing.kind, comment.kind);
    if (kind !== existing.kind) {
      byId.set(comment.comment_id, { ...existing, kind });
    }
  };
  normalizeInboxComments_(mentions.json, ownId, { ownPostIds }).forEach(addComment);

  // Second source: direct replies to the owner's recent posts,
  // which the mentions timeline can miss (e.g. when the reply does
  // not @-mention the owner). Any failure here — the endpoint is
  // plan-gated, so 401/403 is expected on tiers without search —
  // degrades to the mentions result instead of failing the action.
  // Third source: quote tweets of the owner's quoted posts.
  // It degrades the same way, independently of the search source.
  const notes = [];
  if (!ownItems) {
    notes.push('conversation reply search unavailable; showing mentions only');
    notes.push('quote search unavailable; quotes may be missing');
  } else {
    const targets = selectConversationSearchTargets(ownItems, Date.now());
    let searchFailed = false;
    targets.forEach((postId) => {
      const found = signedGet(
        '/tweets/search/recent',
        {
          query: `conversation_id:${postId} is:reply`,
          max_results: String(INBOX_SEARCH_MAX_RESULTS),
          expansions: 'author_id',
          'tweet.fields': INBOX_TWEET_FIELDS,
        },
        credentials
      );
      if (found.code < 200 || found.code >= 300 || !found.json) {
        searchFailed = true;
        return;
      }
      normalizeInboxComments_(found.json, ownId, { ownPostIds }).forEach(addComment);
    });
    if (searchFailed) {
      notes.push('conversation reply search failed for some posts; results may be incomplete');
    }

    const quoteTargets = selectQuoteTargets(ownItems);
    let quoteFailed = false;
    quoteTargets.forEach((postId) => {
      const found = signedGet(
        `/tweets/${encodeURIComponent(postId)}/quote_tweets`,
        {
          max_results: String(INBOX_QUOTE_MAX_RESULTS),
          expansions: 'author_id',
          'tweet.fields': INBOX_TWEET_FIELDS,
        },
        credentials
      );
      if (found.code < 200 || found.code >= 300 || !found.json) {
        quoteFailed = true;
        return;
      }
      normalizeInboxComments_(found.json, ownId, {
        ownPostIds,
        sourceKind: 'quote',
      }).forEach(addComment);
    });
    if (quoteFailed) {
      notes.push('quote search failed for some posts; results may be incomplete');
    }
  }

  const stored = upsertInboxComments([...byId.values()]);
  stored.sort((a, b) => (a.created_at < b.created_at ? 1 : a.created_at > b.created_at ? -1 : 0));
  const result = { ok: true, comments: stored.map(publicInboxComment_) };
  if (notes.length > 0) {
    result.note = notes.join('; ');
  }
  return result;
}

/**
 * Posts a reply to the tweet commentId. Same shape and DRY_RUN
 * behavior as postTweet. Returns { id, text, url }.
 * With options.asQuote the post quotes commentId instead:
 * the payload carries quote_tweet_id and no reply object at all.
 */
export function postReply(text, commentId, options = {}) {
  validateTweetText(text);
  if (!commentId || String(commentId).trim() === '') {
    throw new Error('comment_id is empty');
  }
  if (isDryRun_()) {
    return { id: 'dry-run', text, url: null, dry_run: true };
  }
  const credentials = getCredentials();
  const url = `${API_BASE}/tweets`;
  const authorization = buildAuthorizationHeader(
    { method: 'POST', baseUrl: url },
    credentials
  );
  const { code, json, text: body } = fetchJson(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: authorization },
    payload: options.asQuote
      ? JSON.stringify({ text, quote_tweet_id: String(commentId) })
      : JSON.stringify({
          text,
          reply: { in_reply_to_tweet_id: String(commentId) },
        }),
  });
  if (code < 200 || code >= 300 || !json || !json.data || !json.data.id) {
    throw new Error(`X API post failed (${code}): ${body}`);
  }
  return {
    id: json.data.id,
    text: json.data.text,
    url: `https://x.com/${SCREEN_NAME}/status/${json.data.id}`,
  };
}

/**
 * Marks one inbox row as replied and records our reply's id.
 * Returns true when a matching row was found. Storage failures
 * are swallowed (false) so a posted reply is never reported as
 * failed just because the bookkeeping write failed.
 */
export function markInboxCommentReplied(commentId, replyId) {
  try {
    const sheet = getInboxSheet();
    const values = sheet.getDataRange().getValues();
    for (let index = 1; index < values.length; index += 1) {
      if (String(values[index][0]) === String(commentId)) {
        sheet.getRange(index + 1, INBOX_COL_STATUS).setValue('replied');
        sheet.getRange(index + 1, INBOX_COL_OUR_REPLY_ID).setValue(String(replyId));
        return true;
      }
    }
  } catch (e) {
    // ignore bookkeeping failures
  }
  return false;
}

/**
 * Backs the "reply_comment" action: validates the text with the
 * shared weighted-length rule, posts the reply, verifies it with
 * GET /2/tweets/:id exactly like the plain post flow, then marks
 * the inbox row replied. With options.asQuote the post quotes the
 * comment instead of replying to it; verification and the inbox
 * update are exactly the same, and the result carries
 * quote_tweet_id (and as_quote) instead of in_reply_to_tweet_id.
 * Under DRY_RUN nothing is posted and the inbox row is left
 * untouched.
 */
export function replyToComment(commentId, text, options = {}) {
  const asQuote = !!(options && options.asQuote);
  const posted = postReply(text, commentId, { asQuote });
  const verified = posted.dry_run ? { id: posted.id, text: posted.text } : getTweet(posted.id);
  const result = {
    ok: true,
    id: posted.id,
    url: posted.url,
    postedText: posted.text,
    verifiedText: verified ? verified.text : null,
    comment_id: String(commentId),
  };
  if (asQuote) {
    result.as_quote = true;
    result.quote_tweet_id = String(commentId);
  } else {
    result.in_reply_to_tweet_id = String(commentId);
  }
  if (posted.dry_run) {
    result.dry_run = true;
    return result;
  }
  result.inbox_updated = markInboxCommentReplied(commentId, posted.id);
  return result;
}

// --- DM inbox -----------------------------------------------------------
// Three more user-authenticated actions (the same Google ID token /
// share session path as the comment inbox; the sync token can NOT
// reach them — it is scoped to sns_snapshot only):
// - "list_dms": fetch the newest DM events (GET /2/dm_events,
//   MessageCreate only, at most DM_EVENTS_MAX_PAGES pages), upsert
//   them into the "dm" sheet, and return them grouped into
//   conversations, newest conversation first.
// - "mark_dm_read": advance a conversation's read marker.
// - "reply_dm": send one DM into a conversation and advance both
//   the read and replied markers to the sent event.
// Storage mirrors the comment inbox: per the owner's explicit
// decision the "dm" tab keeps the full message text, exactly like
// the "inbox" tab keeps comment text. Read/replied markers live in
// the separate "dm_state" tab, so re-collecting events can never
// clobber them (the inbox folds the same state into its status
// column instead, because a comment's reply is itself a tweet).

export const DM_SHEET_NAME = 'dm';
export const DM_HEADERS = [
  'event_id',
  'conversation_id',
  'sender_id',
  'sender_username',
  'sender_name',
  'text',
  'created_at',
  'is_mine',
  'first_seen_at',
  'last_seen_at',
];
// 1-based column numbers in the dm sheet, matching DM_HEADERS.
const DM_COL_LAST_SEEN_AT = 10;

export const DM_STATE_SHEET_NAME = 'dm_state';
export const DM_STATE_HEADERS = [
  'conversation_id',
  'last_read_event_id',
  'last_replied_event_id',
  'updated_at',
];
// 1-based column numbers in the dm_state sheet.
const DM_STATE_COL_LAST_READ = 2;
const DM_STATE_COL_LAST_REPLIED = 3;
const DM_STATE_COL_UPDATED_AT = 4;

export const DM_EVENTS_MAX_RESULTS = 100;
// The events endpoint is polled by the frontend, so fetching is
// capped at two pages (200 events) per list_dms run no matter how
// many next_tokens the API offers.
export const DM_EVENTS_MAX_PAGES = 2;
// Each conversation carries at most this many newest messages.
export const DM_MESSAGES_PER_CONVERSATION = 20;
// DM texts allow far more than a tweet: the limit here is a plain
// code-point count (NOT the weighted 280 rule), per the API
// contract agreed with the frontend.
export const DM_REPLY_MAX_CODE_POINTS = 10000;

function getDmSheet() {
  const ss = spreadsheet_();
  let sheet = ss.getSheetByName(DM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DM_SHEET_NAME);
    sheet.appendRow(DM_HEADERS);
  }
  return sheet;
}

function getDmStateSheet() {
  const ss = spreadsheet_();
  let sheet = ss.getSheetByName(DM_STATE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DM_STATE_SHEET_NAME);
    sheet.appendRow(DM_STATE_HEADERS);
  }
  return sheet;
}

/**
 * Compares two X event ids as numeric strings. X ids exceed
 * Number.MAX_SAFE_INTEGER, so the comparison goes through BigInt
 * (with a length-then-lexicographic fallback for non-numeric ids)
 * instead of ever falling back to Number, where distinct ids can
 * collapse into the same value. Returns -1, 0 or 1.
 */
export function compareDmEventIds(a, b) {
  const sa = a == null ? '' : String(a);
  const sb = b == null ? '' : String(b);
  if (sa === sb) {
    return 0;
  }
  try {
    const ba = BigInt(sa);
    const bb = BigInt(sb);
    return ba < bb ? -1 : 1;
  } catch (e) {
    if (sa.length !== sb.length) {
      return sa.length < sb.length ? -1 : 1;
    }
    return sa < sb ? -1 : 1;
  }
}

// Chronological order for DM messages: created_at first, the
// numeric event id breaks ties (and a missing created_at).
function compareDmMessages_(a, b) {
  const ca = a.created_at || '';
  const cb = b.created_at || '';
  if (ca !== cb) {
    return ca < cb ? -1 : 1;
  }
  return compareDmEventIds(a.event_id, b.event_id);
}

// User lookup for DM payloads. Unlike the inbox variant, absent
// names/usernames stay null: the DM API contract types them as
// string|null, not as empty strings.
function dmUsersById_(json) {
  const users = {};
  const list =
    json && json.includes && Array.isArray(json.includes.users) ? json.includes.users : [];
  list.forEach((user) => {
    if (user && user.id != null) {
      users[String(user.id)] = {
        username: user.username != null ? String(user.username) : null,
        name: user.name != null ? String(user.name) : null,
      };
    }
  });
  return users;
}

/**
 * Normalizes one raw dm_events item into a DM message, or returns
 * null for items that cannot be a message: no id, no conversation
 * id, or an event_type other than MessageCreate (the request asks
 * for MessageCreate only, but the filter is re-applied here so a
 * stray event type can never enter the sheet). The returned
 * object also carries participant_ids for conversation grouping;
 * publicDmMessage_ strips it from API responses.
 */
export function normalizeDmEvent(raw, usersById, ownId) {
  if (!raw || raw.id == null || raw.dm_conversation_id == null) {
    return null;
  }
  if (raw.event_type && raw.event_type !== 'MessageCreate') {
    return null;
  }
  const senderId = raw.sender_id != null ? String(raw.sender_id) : '';
  const sender = (usersById && usersById[senderId]) || {};
  return {
    event_id: String(raw.id),
    conversation_id: String(raw.dm_conversation_id),
    sender_id: senderId,
    sender_username: sender.username != null ? sender.username : null,
    sender_name: sender.name != null ? sender.name : null,
    text: raw.text != null ? String(raw.text) : '',
    created_at: raw.created_at || '',
    is_mine: senderId !== '' && String(ownId) === senderId,
    participant_ids: Array.isArray(raw.participant_ids)
      ? raw.participant_ids.map((id) => String(id))
      : [],
  };
}

function publicDmMessage_(message) {
  return {
    event_id: message.event_id,
    conversation_id: message.conversation_id,
    sender_id: message.sender_id,
    sender_username: message.sender_username != null ? message.sender_username : null,
    sender_name: message.sender_name != null ? message.sender_name : null,
    text: message.text,
    created_at: message.created_at,
    is_mine: !!message.is_mine,
  };
}

/**
 * Upserts DM events into the "dm" sheet, keyed by event_id.
 * Existing rows keep first_seen_at; the descriptive fields and
 * last_seen_at are refreshed, so re-collecting never duplicates a
 * row. Storage failures are swallowed like logToSheet failures.
 */
export function upsertDmEvents(events) {
  const nowIso = new Date().toISOString();
  let sheet = null;
  const existingById = new Map();
  try {
    sheet = getDmSheet();
    const values = sheet.getDataRange().getValues();
    values.forEach((row, index) => {
      if (index > 0 && row[0] != null && row[0] !== '') {
        existingById.set(String(row[0]), { row, rowNumber: index + 1 });
      }
    });
  } catch (e) {
    sheet = null;
  }
  if (!sheet) {
    return;
  }
  (Array.isArray(events) ? events : []).forEach((event) => {
    try {
      const existing = existingById.get(event.event_id);
      if (existing) {
        const refresh = [
          [2, event.conversation_id],
          [3, event.sender_id],
          [4, event.sender_username],
          [5, event.sender_name],
          [6, event.text],
          [7, event.created_at],
          [8, event.is_mine],
          [DM_COL_LAST_SEEN_AT, nowIso],
        ];
        refresh.forEach(([col, value]) => {
          sheet.getRange(existing.rowNumber, col).setValue(value == null ? '' : value);
        });
      } else {
        sheet.appendRow([
          event.event_id,
          event.conversation_id,
          event.sender_id,
          event.sender_username == null ? '' : event.sender_username,
          event.sender_name == null ? '' : event.sender_name,
          event.text,
          event.created_at,
          event.is_mine,
          nowIso,
          nowIso,
        ]);
      }
    } catch (e) {
      // one bad row must not lose the rest of the batch
    }
  });
}

/**
 * Reads the whole "dm_state" tab into a Map keyed by
 * conversation_id. Missing/unreadable state reads as empty, which
 * only makes conversations look unread — never the reverse.
 */
export function readDmState() {
  const byConversation = new Map();
  try {
    const sheet = getDmStateSheet();
    const values = sheet.getDataRange().getValues();
    values.forEach((row, index) => {
      if (index > 0 && row[0] != null && row[0] !== '') {
        byConversation.set(String(row[0]), {
          last_read_event_id:
            row[DM_STATE_COL_LAST_READ - 1] != null && row[DM_STATE_COL_LAST_READ - 1] !== ''
              ? String(row[DM_STATE_COL_LAST_READ - 1])
              : null,
          last_replied_event_id:
            row[DM_STATE_COL_LAST_REPLIED - 1] != null &&
            row[DM_STATE_COL_LAST_REPLIED - 1] !== ''
              ? String(row[DM_STATE_COL_LAST_REPLIED - 1])
              : null,
        });
      }
    });
  } catch (e) {
    // treated as no state at all
  }
  return byConversation;
}

/**
 * Advances the dm_state markers of one conversation. A marker is
 * only ever moved forward (numeric event-id comparison), never
 * backwards, and a missing conversation row is created. Returns
 * the resulting markers, or null when storage failed (bookkeeping
 * failures are swallowed, like the inbox's).
 */
export function updateDmState(conversationId, updates) {
  try {
    const sheet = getDmStateSheet();
    const values = sheet.getDataRange().getValues();
    const nowIso = new Date().toISOString();
    const convId = String(conversationId);
    let rowNumber = -1;
    let current = { last_read_event_id: null, last_replied_event_id: null };
    for (let index = 1; index < values.length; index += 1) {
      if (String(values[index][0]) === convId) {
        rowNumber = index + 1;
        current = {
          last_read_event_id:
            values[index][DM_STATE_COL_LAST_READ - 1] != null &&
            values[index][DM_STATE_COL_LAST_READ - 1] !== ''
              ? String(values[index][DM_STATE_COL_LAST_READ - 1])
              : null,
          last_replied_event_id:
            values[index][DM_STATE_COL_LAST_REPLIED - 1] != null &&
            values[index][DM_STATE_COL_LAST_REPLIED - 1] !== ''
              ? String(values[index][DM_STATE_COL_LAST_REPLIED - 1])
              : null,
        };
        break;
      }
    }
    const advance = (cur, next) =>
      next != null && next !== '' && (!cur || compareDmEventIds(next, cur) > 0)
        ? String(next)
        : cur;
    const nextState = {
      last_read_event_id:
        updates && updates.last_read_event_id !== undefined
          ? advance(current.last_read_event_id, updates.last_read_event_id)
          : current.last_read_event_id,
      last_replied_event_id:
        updates && updates.last_replied_event_id !== undefined
          ? advance(current.last_replied_event_id, updates.last_replied_event_id)
          : current.last_replied_event_id,
    };
    if (rowNumber === -1) {
      sheet.appendRow([
        convId,
        nextState.last_read_event_id || '',
        nextState.last_replied_event_id || '',
        nowIso,
      ]);
    } else {
      sheet.getRange(rowNumber, DM_STATE_COL_LAST_READ).setValue(nextState.last_read_event_id || '');
      sheet
        .getRange(rowNumber, DM_STATE_COL_LAST_REPLIED)
        .setValue(nextState.last_replied_event_id || '');
      sheet.getRange(rowNumber, DM_STATE_COL_UPDATED_AT).setValue(nowIso);
    }
    return nextState;
  } catch (e) {
    return null;
  }
}

/**
 * Groups normalized DM events into the API's Conversation shape:
 * conversations ordered by their latest event (newest first),
 * messages chronological (oldest first) and capped at the newest
 * DM_MESSAGES_PER_CONVERSATION. unread / needs_reply derive from
 * the dm_state markers: both need the latest event to be someone
 * else's and newer than the respective marker (or no marker).
 */
export function buildDmConversations(events, stateByConversation, usersById, ownId) {
  const groups = new Map();
  (Array.isArray(events) ? events : []).forEach((event) => {
    if (!groups.has(event.conversation_id)) {
      groups.set(event.conversation_id, []);
    }
    groups.get(event.conversation_id).push(event);
  });
  const conversations = [];
  groups.forEach((list, conversationId) => {
    const sorted = [...list].sort(compareDmMessages_);
    const messages = sorted.slice(-DM_MESSAGES_PER_CONVERSATION).map(publicDmMessage_);
    const latestEvent = messages[messages.length - 1];
    // Participants: everyone the events name, minus the owner,
    // in first-appearance order. When participant_ids are absent,
    // fall back to the distinct other senders.
    const seen = new Set();
    const participants = [];
    const addParticipant = (id, fallback) => {
      const pid = String(id);
      if (pid === String(ownId) || seen.has(pid)) {
        return;
      }
      seen.add(pid);
      const user = (usersById && usersById[pid]) || fallback || {};
      participants.push({
        id: pid,
        name: user.name != null ? user.name : null,
        username: user.username != null ? user.username : null,
      });
    };
    sorted.forEach((event) => {
      (event.participant_ids || []).forEach((pid) => addParticipant(pid, null));
    });
    if (participants.length === 0) {
      sorted.forEach((event) => {
        if (!event.is_mine && event.sender_id) {
          addParticipant(event.sender_id, {
            name: event.sender_name,
            username: event.sender_username,
          });
        }
      });
    }
    const state = (stateByConversation && stateByConversation.get(conversationId)) || {};
    const unread =
      !latestEvent.is_mine &&
      (!state.last_read_event_id ||
        compareDmEventIds(latestEvent.event_id, state.last_read_event_id) > 0);
    const needsReply =
      !latestEvent.is_mine &&
      (!state.last_replied_event_id ||
        compareDmEventIds(latestEvent.event_id, state.last_replied_event_id) > 0);
    conversations.push({
      conversation_id: conversationId,
      participants,
      latest_event: latestEvent,
      messages,
      unread,
      needs_reply: needsReply,
    });
  });
  conversations.sort((a, b) => compareDmMessages_(b.latest_event, a.latest_event));
  return conversations;
}

/**
 * Backs the "list_dms" action. Returns
 * { ok: true, conversations: [...] }. A dm_events failure throws
 * (the action then reports ok:false via doPost), like a mentions
 * failure in list_comments. Under DRY_RUN the X API is never
 * called and no sheet is touched.
 */
export function listDms() {
  if (isDryRun_()) {
    return { ok: true, dry_run: true, conversations: [] };
  }
  const credentials = getCredentials();
  const me = signedGet('/users/me', {}, credentials);
  const meData = me.json && me.json.data ? me.json.data : null;
  if (me.code < 200 || me.code >= 300 || !meData || meData.id == null) {
    throw new Error(`X API profile failed: ${summarizeXApiError(me.code, me.text)}`);
  }
  const ownId = String(meData.id);
  const usersById = {
    [ownId]: {
      username: meData.username != null ? String(meData.username) : null,
      name: meData.name != null ? String(meData.name) : null,
    },
  };
  const byId = new Map();
  let paginationToken = null;
  for (let page = 0; page < DM_EVENTS_MAX_PAGES; page += 1) {
    const query = {
      event_types: 'MessageCreate',
      max_results: String(DM_EVENTS_MAX_RESULTS),
      expansions: 'sender_id,participant_ids',
      'dm_event.fields':
        'id,text,event_type,dm_conversation_id,created_at,sender_id,participant_ids',
      'user.fields': 'name,username',
    };
    if (paginationToken) {
      query.pagination_token = paginationToken;
    }
    const res = signedGet('/dm_events', query, credentials);
    if (res.code < 200 || res.code >= 300 || !res.json) {
      throw new Error(`X API dm_events failed: ${summarizeXApiError(res.code, res.text)}`);
    }
    Object.assign(usersById, dmUsersById_(res.json));
    const items = Array.isArray(res.json.data) ? res.json.data : [];
    items.forEach((raw) => {
      const event = normalizeDmEvent(raw, usersById, ownId);
      if (event && !byId.has(event.event_id)) {
        byId.set(event.event_id, event);
      }
    });
    paginationToken =
      res.json.meta && res.json.meta.next_token ? String(res.json.meta.next_token) : null;
    if (!paginationToken) {
      break;
    }
  }
  const events = [...byId.values()];
  upsertDmEvents(events);
  const conversations = buildDmConversations(events, readDmState(), usersById, ownId);
  return { ok: true, conversations };
}

/**
 * Backs the "mark_dm_read" action: advances the conversation's
 * last_read_event_id (never backwards) and stamps updated_at.
 * Under DRY_RUN nothing is written.
 */
export function markDmRead(conversationId, eventId) {
  if (!conversationId || String(conversationId).trim() === '') {
    throw new Error('conversation_id is empty');
  }
  if (!eventId || String(eventId).trim() === '') {
    throw new Error('event_id is empty');
  }
  const result = {
    ok: true,
    conversation_id: String(conversationId),
    event_id: String(eventId),
  };
  if (isDryRun_()) {
    result.dry_run = true;
    return result;
  }
  updateDmState(result.conversation_id, { last_read_event_id: result.event_id });
  return result;
}

/**
 * DM text validation: a non-empty string of at most
 * DM_REPLY_MAX_CODE_POINTS Unicode code points. Deliberately NOT
 * the weighted 280 rule — that limit belongs to tweets, and
 * spreading code points (not UTF-16 units) is what counts an
 * emoji as one character here. Throws on invalid text.
 */
export function validateDmText(text) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    throw new Error('text is empty');
  }
  const count = [...text].length;
  if (count > DM_REPLY_MAX_CODE_POINTS) {
    throw new Error(
      `text is too long: ${count} code points exceeds ${DM_REPLY_MAX_CODE_POINTS}`
    );
  }
  return text;
}

/**
 * Sends one DM into a conversation via
 * POST /2/dm_conversations/:id/messages. The JSON body is not
 * part of the OAuth signature, exactly like postTweet. Returns
 * { event_id } with the X response's dm_event_id (null when the
 * API did not return one). Under DRY_RUN nothing is sent.
 */
export function postDmMessage(conversationId, text) {
  validateDmText(text);
  if (!conversationId || String(conversationId).trim() === '') {
    throw new Error('conversation_id is empty');
  }
  if (isDryRun_()) {
    return { event_id: null, dry_run: true };
  }
  const credentials = getCredentials();
  const url = `${API_BASE}/dm_conversations/${encodeURIComponent(String(conversationId))}/messages`;
  const authorization = buildAuthorizationHeader(
    { method: 'POST', baseUrl: url },
    credentials
  );
  const { code, json, text: body } = fetchJson(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: authorization },
    payload: JSON.stringify({ text }),
  });
  if (code < 200 || code >= 300) {
    throw new Error(`X API DM post failed (${code}): ${body}`);
  }
  return {
    event_id: json && json.data && json.data.dm_event_id != null
      ? String(json.data.dm_event_id)
      : null,
  };
}

/**
 * Backs the "reply_dm" action: sends the DM, then advances both
 * dm_state markers to the sent event and stores the sent event in
 * the "dm" tab as one of our own messages (both only when the API
 * returned a dm_event_id). The bookkeeping never fails a reply
 * that was already sent. Under DRY_RUN nothing is sent and no
 * sheet changes.
 */
export function replyDm(conversationId, text) {
  const posted = postDmMessage(conversationId, text);
  const convId = String(conversationId);
  if (posted.dry_run) {
    return { ok: true, conversation_id: convId, event_id: null, dry_run: true };
  }
  if (posted.event_id) {
    updateDmState(convId, {
      last_read_event_id: posted.event_id,
      last_replied_event_id: posted.event_id,
    });
    try {
      const credentials = getCredentials();
      const me = signedGet('/users/me', {}, credentials);
      const meData = me.json && me.json.data ? me.json.data : null;
      if (meData && meData.id != null) {
        upsertDmEvents([
          {
            event_id: posted.event_id,
            conversation_id: convId,
            sender_id: String(meData.id),
            sender_username: meData.username != null ? String(meData.username) : null,
            sender_name: meData.name != null ? String(meData.name) : null,
            text: String(text),
            created_at: new Date().toISOString(),
            is_mine: true,
            participant_ids: [],
          },
        ]);
      }
    } catch (e) {
      // the reply is sent; only the local copy is missing
    }
  }
  return { ok: true, conversation_id: convId, event_id: posted.event_id };
}

/**
 * Verifies a Google ID token. Returns the token payload when the token is
 * valid, email-verified, and the email is ALLOWED_EMAIL; otherwise null.
 */
export function verifyIdToken(idToken) {
  if (!idToken) return null;
  const { code, json } = fetchJson(
    `${TOKENINFO_URL}?id_token=${encodeURIComponent(idToken)}`,
    { method: 'get' }
  );
  if (code !== 200 || !json) return null;
  const emailVerified = json.email_verified === true || json.email_verified === 'true';
  if (!emailVerified) return null;
  if (json.email !== ALLOWED_EMAIL) return null;
  const clientId = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
  if (clientId && json.aud !== clientId) return null;
  return json;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

/**
 * Health check.
 */
export function doGet() {
  return jsonOutput({ ok: true, service: 'tweet-gas', time: new Date().toISOString() });
}

/**
 * POST JSON API. Bodies:
 * - { idToken, text }                       -> posts and verifies.
 * - { share_token, text }                   -> same, via a share link.
 * - { session_token, ... }                  -> same, via the 30-day
 *   owner session exchanged at sign-in (see "Owner sessions").
 * - { idToken, action: "exchange_session" } -> issues an owner session.
 * - { session_token, action: "revoke_session" } -> revokes it (logout).
 * - { idToken|session_token, action: "issue_share" } -> issues a share link.
 * - { idToken|session_token, action: "revoke_share" } -> revokes links.
 * - { share_token, action: "validate_share" } -> checks a link and
 *   returns its expiry as { ok, _share: { expires_at } }.
 * - { sync_token, action: "sns_snapshot" }  -> the X numbers for
 *   social-gas (machine-to-machine; see the sync token section).
 * - { idToken|share_token, action: "list_comments" } -> comments on
 *   the owner's posts, collected into the "inbox" sheet.
 * - { idToken|share_token, action: "reply_comment", comment_id,
 *   text, as_quote? } -> posts a reply to that comment and marks
 *   it replied; with as_quote: true it quotes the comment instead.
 * - { idToken|share_token, action: "list_dms" } -> DM events
 *   collected into the "dm" sheet, grouped into conversations.
 * - { idToken|share_token, action: "mark_dm_read", conversation_id,
 *   event_id } -> advances the conversation's read marker.
 * - { idToken|share_token, action: "reply_dm", conversation_id,
 *   text } -> sends a DM into that conversation.
 * Returns { ok, id, url, postedText, verifiedText } (post) or
 * { ok: false, error }.
 */
export function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    // Machine-to-machine sync (social-gas): the sync token is
    // checked before every other credential, and it unlocks the
    // read-only sns_snapshot action only — the same scoping
    // alarm-gas applies to its sync token.
    if (body.sync_token) {
      if (!validateSyncToken(body.sync_token)) {
        logToSheet('doPost', { ok: false, error: 'sync token invalid' });
        return jsonOutput({ ok: false, status: 401, error: 'sync token invalid' });
      }
      if (!SYNC_ACTIONS.includes(body.action)) {
        logToSheet('doPost', { ok: false, error: 'sync token may only read the SNS snapshot' });
        return jsonOutput({
          ok: false,
          status: 403,
          error: 'sync token may only read the SNS snapshot',
        });
      }
      const snapshot = buildSnsSnapshot();
      logToSheet('sns_snapshot', {
        ok: true,
        followers: snapshot.x.followers,
        posts: Array.isArray(snapshot.x.recent_posts) ? snapshot.x.recent_posts.length : null,
        x_error: snapshot.x_error,
        posts_error: snapshot.posts_error,
      });
      return jsonOutput({ ...snapshot, _sync: true });
    }
    if (body.action === 'exchange_session') {
      const identity = verifyIdToken(body.idToken);
      if (!identity) {
        logToSheet('doPost', { ok: false, error: 'unauthorized' });
        return jsonOutput({ ok: false, error: 'unauthorized' });
      }
      const result = issueOwnerSession();
      logToSheet('exchange_session', { ok: true, email: identity.email });
      return jsonOutput({ ok: true, ...result });
    }
    if (body.action === 'revoke_session') {
      revokeOwnerSession(body.session_token);
      return jsonOutput({ ok: true, revoked: true });
    }
    if (body.action === 'issue_share' || body.action === 'revoke_share') {
      const identity = verifyIdToken(body.idToken);
      const sessionAuth = !identity && Boolean(validateOwnerToken(body.session_token));
      if (!identity && !sessionAuth) {
        logToSheet('doPost', { ok: false, error: 'unauthorized' });
        return jsonOutput({ ok: false, error: 'unauthorized' });
      }
      const result =
        body.action === 'issue_share' ? issueShareSession() : revokeShareSession();
      logToSheet(body.action, { ok: true, email: identity ? identity.email : 'owner-session' });
      return jsonOutput({ ok: true, ...result });
    }
    let identity = null;
    let shareSession = null;
    let ownerSession = null;
    if (body.session_token) {
      ownerSession = validateOwnerToken(body.session_token);
      if (!ownerSession) {
        logToSheet('doPost', { ok: false, error: 'unauthorized' });
        return jsonOutput({ ok: false, error: 'unauthorized' });
      }
    } else if (body.share_token) {
      shareSession = validateShareToken(body.share_token);
      if (!shareSession) {
        logToSheet('doPost', { ok: false, error: 'unauthorized' });
        return jsonOutput({ ok: false, error: 'unauthorized' });
      }
    } else {
      identity = verifyIdToken(body.idToken);
      if (!identity) {
        logToSheet('doPost', { ok: false, error: 'unauthorized' });
        return jsonOutput({ ok: false, error: 'unauthorized' });
      }
    }
    if (body.action === 'validate_share') {
      // Reached only with a valid share session (see above); an ID
      // token alone cannot validate a share link.
      if (!shareSession) {
        logToSheet('doPost', { ok: false, error: 'unauthorized' });
        return jsonOutput({ ok: false, error: 'unauthorized' });
      }
      return jsonOutput({
        ok: true,
        _share: { expires_at: shareSession.expires_at },
      });
    }
    if (
      body.action === 'list_comments' ||
      body.action === 'reply_comment' ||
      body.action === 'list_dms' ||
      body.action === 'mark_dm_read' ||
      body.action === 'reply_dm'
    ) {
      // Comment and DM inboxes: reached only through the user auth
      // path above (Google ID token or share session) — the sync
      // token was already rejected for anything but sns_snapshot.
      const actor = identity ? identity.email : ownerSession ? '(owner session)' : '(share link)';
      if (body.action === 'list_dms') {
        const inbox = listDms();
        const result = { ...inbox };
        if (shareSession) {
          result._share = { expires_at: shareSession.expires_at };
        }
        logToSheet('list_dms', {
          email: actor,
          ok: true,
          count: inbox.conversations.length,
          dry_run: !!inbox.dry_run,
        });
        return jsonOutput(result);
      }
      if (body.action === 'mark_dm_read') {
        const marked = markDmRead(body.conversation_id, body.event_id);
        const result = { ...marked };
        if (shareSession) {
          result._share = { expires_at: shareSession.expires_at };
        }
        logToSheet('mark_dm_read', { email: actor, ...marked });
        return jsonOutput(result);
      }
      if (body.action === 'reply_dm') {
        const replied = replyDm(body.conversation_id, body.text);
        const result = { ...replied };
        if (shareSession) {
          result._share = { expires_at: shareSession.expires_at };
        }
        // Metadata only: the reply text itself is not copied into
        // the action log (the sent event lives in the "dm" tab).
        logToSheet('reply_dm', { email: actor, ...replied });
        return jsonOutput(result);
      }
      if (body.action === 'list_comments') {
        const inbox = listComments();
        const result = { ...inbox };
        if (shareSession) {
          result._share = { expires_at: shareSession.expires_at };
        }
        logToSheet('list_comments', {
          email: actor,
          ok: true,
          count: inbox.comments.length,
          note: inbox.note || null,
          dry_run: !!inbox.dry_run,
        });
        return jsonOutput(result);
      }
      const replied = replyToComment(body.comment_id, body.text, {
        asQuote: body.as_quote === true,
      });
      const result = { ...replied };
      if (shareSession) {
        result._share = { expires_at: shareSession.expires_at };
      }
      logToSheet('reply_comment', { email: actor, ...result });
      return jsonOutput(result);
    }
    const posted = postTweet(body.text);
    const verified = posted.dry_run ? { id: posted.id, text: posted.text } : getTweet(posted.id);
    const result = {
      ok: true,
      id: posted.id,
      url: posted.url,
      postedText: posted.text,
      verifiedText: verified ? verified.text : null,
    };
    if (posted.dry_run) {
      result.dry_run = true;
    }
    if (shareSession) {
      result._share = { expires_at: shareSession.expires_at };
    }
    logToSheet('doPost', {
      email: identity ? identity.email : ownerSession ? '(owner session)' : '(share link)',
      ...result,
    });
    return jsonOutput(result);
  } catch (err) {
    const message = String(err && err.message ? err.message : err);
    logToSheet('doPost', { ok: false, error: message });
    return jsonOutput({ ok: false, error: message });
  }
}

/**
 * Manual smoke test from the Apps Script editor: posts a marked test tweet.
 */
export function postTestTweet() {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const result = postTweet(`テスト投稿です（tweet-gas 動作確認 ${stamp}）`);
  Logger.log(JSON.stringify(result));
  return result;
}

// --- DM access probe (phase 2 prerequisite) ---
// Editor-only check: can the current X credentials read the DM
// endpoints at all? The answer decides whether a DM inbox
// (phase 2) can reuse the existing OAuth 1.0a keys, or whether the
// app permission has to include Direct Messages and the access
// token be re-issued. This is deliberately NOT a doPost action, so
// no new web-reachable surface is added. No DM content is ever
// read out, logged or returned: only the HTTP status and, on
// failure, the error title/detail the API itself returns.

export const DM_EVENTS_URL = 'https://api.x.com/2/dm_events';

/**
 * Reduces one dm_events response to { status, ok } plus, on
 * success, only how many events came back (the events themselves
 * are discarded unread) and, on failure, only the error
 * title/detail.
 */
export function summarizeDmProbeResponse(code, text) {
  const result = { status: code, ok: code >= 200 && code < 300 };
  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    parsed = null;
  }
  if (!parsed) {
    return result;
  }
  if (result.ok) {
    if (Array.isArray(parsed.data)) {
      result.count = parsed.data.length;
    }
    return result;
  }
  if (parsed.title) {
    result.title = String(parsed.title).slice(0, 200);
  }
  if (parsed.detail) {
    result.detail = String(parsed.detail).slice(0, 200);
  }
  if (
    !result.title &&
    !result.detail &&
    Array.isArray(parsed.errors) &&
    parsed.errors.length
  ) {
    const first = parsed.errors[0] || {};
    if (first.message) {
      result.detail = String(first.message).slice(0, 200);
    }
  }
  return result;
}

/**
 * Manual probe from the Apps Script editor: one signed GET of
 * GET /2/dm_events?max_results=1 with the existing credentials.
 * Under DRY_RUN the X API is never called. The fetch bypasses
 * fetchJson() on purpose: that helper logs the full response body
 * to the log sheet, and a DM response body must never be stored
 * anywhere. Only the status summary reaches the Logger, the
 * return value and the log sheet.
 */
export function probeDmAccess() {
  if (isDryRun_()) {
    const result = { dry_run: true };
    Logger.log(JSON.stringify(result));
    return result;
  }
  const credentials = getCredentials();
  const query = { max_results: 1 };
  const authorization = buildAuthorizationHeader(
    { method: 'GET', baseUrl: DM_EVENTS_URL, query },
    credentials
  );
  const response = UrlFetchApp.fetch(`${DM_EVENTS_URL}?max_results=1`, {
    method: 'get',
    headers: { Authorization: authorization },
    muteHttpExceptions: true,
  });
  const result = summarizeDmProbeResponse(
    response.getResponseCode(),
    response.getContentText()
  );
  logToSheet('probe_dm_access', { status: result.status, ok: result.ok });
  Logger.log(JSON.stringify(result));
  return result;
}
