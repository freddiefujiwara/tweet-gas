/**
 * tweet-gas: X (Twitter) posting backend on Google Apps Script.
 *
 * - Posts to POST https://api.twitter.com/2/tweets with OAuth 1.0a user context.
 * - The four API keys live in Script Properties (never in code):
 *     X_CONSUMER_KEY, X_CONSUMER_SECRET, X_ACCESS_TOKEN, X_ACCESS_TOKEN_SECRET
 * - The web app (doPost) only accepts requests carrying a Google ID token
 *   whose verified email is ALLOWED_EMAIL. Optional Script Property
 *   GOOGLE_CLIENT_ID additionally pins the token audience.
 */

export const API_BASE = 'https://api.twitter.com/2';
export const ALLOWED_EMAIL = 'zegenvs@gmail.com';
export const SCREEN_NAME = 'freddiefujiwara';
export const TOKENINFO_URL = 'https://oauth2.googleapis.com/tokeninfo';
export const LOG_SHEET_NAME = 'log';
export const LOG_DETAIL_MAX = 40000;

/**
 * Appends one row to the "log" sheet of the bound spreadsheet.
 * Logging must never break the main flow, so errors are swallowed.
 * detail is stored as a JSON string in a single cell.
 */
export function logToSheet(action, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
 * POST { idToken, text } -> posts and verifies.
 * Returns { ok, id, url, postedText, verifiedText } or { ok: false, error }.
 */
export function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const identity = verifyIdToken(body.idToken);
    if (!identity) {
      logToSheet('doPost', { ok: false, error: 'unauthorized' });
      return jsonOutput({ ok: false, error: 'unauthorized' });
    }
    const posted = postTweet(body.text);
    const verified = getTweet(posted.id);
    const result = {
      ok: true,
      id: posted.id,
      url: posted.url,
      postedText: posted.text,
      verifiedText: verified ? verified.text : null,
    };
    logToSheet('doPost', { email: identity.email, ...result });
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
