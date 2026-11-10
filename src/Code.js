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

function getShareSheet() {
  const ss = spreadsheet_();
  let sheet = ss.getSheetByName(SHARE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHARE_SHEET_NAME);
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

function readShareRows() {
  const sheet = getShareSheet();
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

function revokeActiveShareSessions(nowIso) {
  const { sheet, values } = readShareRows();
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

function issueShareSession() {
  const now = new Date();
  const nowIso = now.toISOString();
  const expiresIso = new Date(now.getTime() + SHARE_TTL_SECONDS * 1000).toISOString();
  revokeActiveShareSessions(nowIso);
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  getShareSheet().appendRow([hashShareToken(token), nowIso, expiresIso, '', 0, '']);
  return {
    share_token: token,
    issued_at: nowIso,
    expires_at: expiresIso,
    ttl_seconds: SHARE_TTL_SECONDS,
  };
}

function revokeShareSession() {
  revokeActiveShareSessions(new Date().toISOString());
  return { revoked: true };
}

function validateShareToken(token) {
  if (!token) {
    return null;
  }
  const hash = hashShareToken(token);
  const { sheet, values } = readShareRows();
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
  // body are fine to log; the credential is not).
  logToSheet('fetch', {
    method: (options.method || 'get').toUpperCase(),
    url: url.split('?')[0],
    status: code,
    response: text.length > LOG_DETAIL_MAX ? text.slice(0, LOG_DETAIL_MAX) + '…(truncated)' : text,
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
//   mentions timeline plus a per-conversation recent search, store
//   them in the "inbox" sheet, and return them newest first.
// - "reply_comment": post a reply to one stored comment and mark
//   the inbox row as replied.
// Authorization caveat: mentions and especially the recent-search
// endpoint are not available on every X API plan/tier, and an app
// without the right access gets 401/403 for them. list_comments
// therefore treats search failure as a degraded result (mentions
// only, plus a note) instead of failing the whole action.

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
];
// 1-based column numbers in the inbox sheet, matching INBOX_HEADERS.
const INBOX_COL_STATUS = 9;
const INBOX_COL_OUR_REPLY_ID = 10;
const INBOX_COL_LAST_SEEN_AT = 12;

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

const INBOX_TWEET_FIELDS = 'created_at,conversation_id,in_reply_to_user_id,referenced_tweets';

function getInboxSheet() {
  const ss = spreadsheet_();
  let sheet = ss.getSheetByName(INBOX_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INBOX_SHEET_NAME);
    sheet.appendRow(INBOX_HEADERS);
  }
  return sheet;
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
 */
export function normalizeInboxComment(tweet, usersById, ownId) {
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
  };
}

function normalizeInboxComments_(json, ownId) {
  const usersById = usersByIdFromIncludes_(json);
  const items = json && Array.isArray(json.data) ? json.data : [];
  return items
    .map((tweet) => normalizeInboxComment(tweet, usersById, ownId))
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
    const merged = {
      ...comment,
      status: existing ? String(existing.row[INBOX_COL_STATUS - 1] || 'new') : 'new',
      our_reply_id: existing ? String(existing.row[INBOX_COL_OUR_REPLY_ID - 1] || '') : '',
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
  };
}

/**
 * Backs the "list_comments" action. Returns
 * { ok: true, comments: [...newest first], note? }.
 * Mentions failure throws (the action then reports ok:false via
 * doPost); search failure only degrades to mentions with a note.
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
  normalizeInboxComments_(mentions.json, ownId).forEach((comment) => {
    byId.set(comment.comment_id, comment);
  });

  // Second source: direct replies to the owner's recent posts,
  // which the mentions timeline can miss (e.g. when the reply does
  // not @-mention the owner). Any failure here — the endpoint is
  // plan-gated, so 401/403 is expected on tiers without search —
  // degrades to the mentions result instead of failing the action.
  let note = null;
  const ownPosts = signedGet(
    `/users/${encodeURIComponent(ownId)}/tweets`,
    {
      max_results: '100',
      exclude: 'retweets,replies',
      'tweet.fields': 'created_at',
    },
    credentials
  );
  const ownItems =
    ownPosts.code >= 200 && ownPosts.code < 300 && ownPosts.json && Array.isArray(ownPosts.json.data)
      ? ownPosts.json.data
      : null;
  if (!ownItems) {
    note = 'conversation reply search unavailable; showing mentions only';
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
      normalizeInboxComments_(found.json, ownId).forEach((comment) => {
        if (!byId.has(comment.comment_id)) {
          byId.set(comment.comment_id, comment);
        }
      });
    });
    if (searchFailed) {
      note = 'conversation reply search failed for some posts; results may be incomplete';
    }
  }

  const stored = upsertInboxComments([...byId.values()]);
  stored.sort((a, b) => (a.created_at < b.created_at ? 1 : a.created_at > b.created_at ? -1 : 0));
  const result = { ok: true, comments: stored.map(publicInboxComment_) };
  if (note) {
    result.note = note;
  }
  return result;
}

/**
 * Posts a reply to the tweet commentId. Same shape and DRY_RUN
 * behavior as postTweet. Returns { id, text, url }.
 */
export function postReply(text, commentId) {
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
    payload: JSON.stringify({
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
 * the inbox row replied. Under DRY_RUN nothing is posted and the
 * inbox row is left untouched.
 */
export function replyToComment(commentId, text) {
  const posted = postReply(text, commentId);
  const verified = posted.dry_run ? { id: posted.id, text: posted.text } : getTweet(posted.id);
  const result = {
    ok: true,
    id: posted.id,
    url: posted.url,
    postedText: posted.text,
    verifiedText: verified ? verified.text : null,
    comment_id: String(commentId),
    in_reply_to_tweet_id: String(commentId),
  };
  if (posted.dry_run) {
    result.dry_run = true;
    return result;
  }
  result.inbox_updated = markInboxCommentReplied(commentId, posted.id);
  return result;
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
 * - { idToken, action: "issue_share" }      -> issues a share link.
 * - { idToken, action: "revoke_share" }     -> revokes active links.
 * - { share_token, action: "validate_share" } -> checks a link and
 *   returns its expiry as { ok, _share: { expires_at } }.
 * - { sync_token, action: "sns_snapshot" }  -> the X numbers for
 *   social-gas (machine-to-machine; see the sync token section).
 * - { idToken|share_token, action: "list_comments" } -> comments on
 *   the owner's posts, collected into the "inbox" sheet.
 * - { idToken|share_token, action: "reply_comment", comment_id,
 *   text } -> posts a reply to that comment and marks it replied.
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
    if (body.action === 'issue_share' || body.action === 'revoke_share') {
      const identity = verifyIdToken(body.idToken);
      if (!identity) {
        logToSheet('doPost', { ok: false, error: 'unauthorized' });
        return jsonOutput({ ok: false, error: 'unauthorized' });
      }
      const result =
        body.action === 'issue_share' ? issueShareSession() : revokeShareSession();
      logToSheet(body.action, { ok: true, email: identity.email });
      return jsonOutput({ ok: true, ...result });
    }
    let identity = null;
    let shareSession = null;
    if (body.share_token) {
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
    if (body.action === 'list_comments' || body.action === 'reply_comment') {
      // Comment inbox: reached only through the user auth path
      // above (Google ID token or share session) — the sync token
      // was already rejected for anything but sns_snapshot.
      const actor = identity ? identity.email : '(share link)';
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
      const replied = replyToComment(body.comment_id, body.text);
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
      email: identity ? identity.email : '(share link)',
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
