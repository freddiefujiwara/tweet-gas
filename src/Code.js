/**
 * tweet-gas: X (Twitter) posting backend on Google Apps Script.
 *
 * - Posts to POST https://api.twitter.com/2/tweets with OAuth 1.0a user context.
 * - The four API keys live in Script Properties (never in code):
 *     X_CONSUMER_KEY, X_CONSUMER_SECRET, X_ACCESS_TOKEN, X_ACCESS_TOKEN_SECRET
 * - The web app (doPost) only accepts requests carrying a Google ID token
 *   whose verified email is ALLOWED_EMAIL. Optional Script Property
 *   GOOGLE_CLIENT_ID additionally pins the token audience.
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

/**
 * Posts a tweet. Returns { id, text, url }.
 */
export function postTweet(text) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    throw new Error('text is empty');
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
