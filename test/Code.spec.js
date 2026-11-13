import { describe, it, expect, vi, beforeEach } from 'vitest';
import crypto from 'node:crypto';
import * as Code from '../src/Code.js';

// --- GAS service mocks -------------------------------------------------

const mockGetProperty = vi.fn();
const mockFetch = vi.fn();

let uuidCounter = 0;
global.Utilities = {
  MacAlgorithm: { HMAC_SHA_1: 'HMAC_SHA_1' },
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  computeHmacSignature: (algorithm, value, key) => {
    if (algorithm !== 'HMAC_SHA_1') throw new Error('unexpected algorithm: ' + algorithm);
    return crypto.createHmac('sha1', key).update(value, 'utf8').digest();
  },
  computeDigest: (algorithm, value) => {
    if (algorithm !== 'SHA_256') throw new Error('unexpected algorithm: ' + algorithm);
    return [...crypto.createHash('sha256').update(String(value), 'utf8').digest()];
  },
  base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
  getUuid: () =>
    `00000000-0000-4000-8000-${String(uuidCounter++).padStart(12, '0')}`,
  formatDate: () => '2026-11-01 05:00:00',
};
global.PropertiesService = {
  getScriptProperties: () => ({ getProperty: mockGetProperty }),
};
global.UrlFetchApp = { fetch: mockFetch };

const appendedRows = [];
let sheetExists = true;
const mockSheet = { appendRow: (row) => appendedRows.push(row) };

const SHARE_HEADERS = [
  'token_hash',
  'issued_at',
  'expires_at',
  'revoked_at',
  'access_count',
  'last_accessed_at',
];
let shareValues = [];
let shareSheetExists = true;
const mockShareSheet = {
  appendRow: (row) => shareValues.push([...row]),
  getDataRange: () => ({ getValues: () => shareValues }),
  getRange: (row, col) => ({
    setValue: (value) => {
      shareValues[row - 1][col - 1] = value;
    },
  }),
};

let ownerValues = [];
let ownerSheetExists = true;
const mockOwnerSheet = {
  appendRow: (row) => ownerValues.push([...row]),
  getDataRange: () => ({ getValues: () => ownerValues }),
  getRange: (row, col) => ({
    setValue: (value) => {
      ownerValues[row - 1][col - 1] = value;
    },
  }),
};

function installSpreadsheetApp() {
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => {
        if (name === 'share_sessions') {
          return shareSheetExists ? mockShareSheet : null;
        }
        if (name === 'owner_sessions') {
          return ownerSheetExists ? mockOwnerSheet : null;
        }
        return sheetExists ? mockSheet : null;
      },
      insertSheet: (name) => {
        if (name === 'share_sessions') {
          shareSheetExists = true;
          return mockShareSheet;
        }
        if (name === 'owner_sessions') {
          ownerSheetExists = true;
          return mockOwnerSheet;
        }
        sheetExists = true;
        return mockSheet;
      },
    }),
  };
}
installSpreadsheetApp();
global.ContentService = {
  MimeType: { JSON: 'JSON' },
  createTextOutput: (text) => ({
    text,
    setMimeType() {
      return this;
    },
  }),
};
global.Logger = { log: vi.fn() };

const KEYS = {
  X_CONSUMER_KEY: 'ck',
  X_CONSUMER_SECRET: 'cs',
  X_ACCESS_TOKEN: 'at',
  X_ACCESS_TOKEN_SECRET: 'ats',
};

function fetchResponse(code, obj) {
  return {
    getResponseCode: () => code,
    getContentText: () => (typeof obj === 'string' ? obj : JSON.stringify(obj)),
  };
}

beforeEach(() => {
  vi.clearAllMocks();
  appendedRows.length = 0;
  sheetExists = true;
  shareValues = [[...SHARE_HEADERS]];
  shareSheetExists = true;
  ownerValues = [[...SHARE_HEADERS]];
  ownerSheetExists = true;
  uuidCounter = 0;
  installSpreadsheetApp();
  mockGetProperty.mockImplementation((name) => KEYS[name] ?? null);
});

// --- Tests ---------------------------------------------------------------

describe('percentEncode', () => {
  it('encodes per RFC 3986', () => {
    expect(Code.percentEncode('Ladies + Gentlemen!')).toBe('Ladies%20%2B%20Gentlemen%21');
    expect(Code.percentEncode('abc-._~')).toBe('abc-._~');
  });
});

describe('buildAuthorizationHeader', () => {
  it('matches the RFC 5849 signature test vector', () => {
    const header = Code.buildAuthorizationHeader(
      {
        method: 'GET',
        baseUrl: 'http://photos.example.net/photos',
        query: { file: 'vacation.jpg', size: 'original' },
      },
      {
        consumerKey: 'dpf43f3p2l4k3l03',
        consumerSecret: 'kd94hf93k423kf44',
        accessToken: 'nnch734d00sl2jdk',
        accessTokenSecret: 'pfkkdhi9sl3r4s00',
      },
      { nonce: 'kllo9940pd9333jh', timestamp: '1191242096' }
    );
    expect(header).toContain('oauth_signature="tR3%2BTy81lMeYAr%2FFid0kMTYa%2FWM%3D"');
  });
});

describe('postTweet', () => {
  it('posts and returns id and url', () => {
    mockFetch.mockReturnValue(fetchResponse(201, { data: { id: '123', text: 'hello' } }));
    const result = Code.postTweet('hello');
    expect(result).toEqual({
      id: '123',
      text: 'hello',
      url: 'https://x.com/freddiefujiwara/status/123',
    });
    const [url, options] = mockFetch.mock.calls[0];
    expect(url).toBe('https://api.twitter.com/2/tweets');
    expect(options.headers.Authorization).toMatch(/^OAuth /);
    expect(JSON.parse(options.payload)).toEqual({ text: 'hello' });
  });

  it('throws on API error', () => {
    mockFetch.mockReturnValue(fetchResponse(403, { detail: 'forbidden' }));
    expect(() => Code.postTweet('hello')).toThrow(/403/);
  });

  it('throws on empty text', () => {
    expect(() => Code.postTweet('  ')).toThrow(/empty/);
  });

  it('throws when script properties are missing', () => {
    mockGetProperty.mockReturnValue(null);
    expect(() => Code.postTweet('hello')).toThrow(/Missing Script Properties/);
  });
});

describe('getTweet', () => {
  it('returns the tweet on success', () => {
    mockFetch.mockReturnValue(fetchResponse(200, { data: { id: '123', text: 'hello' } }));
    expect(Code.getTweet('123')).toEqual({ id: '123', text: 'hello' });
  });

  it('returns null on failure', () => {
    mockFetch.mockReturnValue(fetchResponse(404, {}));
    expect(Code.getTweet('999')).toBeNull();
  });
});

describe('verifyIdToken', () => {
  it('accepts the allowed verified email', () => {
    mockFetch.mockReturnValue(
      fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true', aud: 'cid' })
    );
    mockGetProperty.mockImplementation((name) => (name === 'GOOGLE_CLIENT_ID' ? 'cid' : KEYS[name] ?? null));
    const payload = Code.verifyIdToken('token');
    expect(payload.email).toBe('zegenvs@gmail.com');
  });

  it('rejects other emails', () => {
    mockFetch.mockReturnValue(
      fetchResponse(200, { email: 'someone@example.com', email_verified: 'true' })
    );
    expect(Code.verifyIdToken('token')).toBeNull();
  });

  it('rejects unverified email', () => {
    mockFetch.mockReturnValue(
      fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'false' })
    );
    expect(Code.verifyIdToken('token')).toBeNull();
  });

  it('rejects wrong audience when GOOGLE_CLIENT_ID is set', () => {
    mockFetch.mockReturnValue(
      fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true', aud: 'other' })
    );
    mockGetProperty.mockImplementation((name) => (name === 'GOOGLE_CLIENT_ID' ? 'cid' : KEYS[name] ?? null));
    expect(Code.verifyIdToken('token')).toBeNull();
  });

  it('rejects invalid token responses', () => {
    mockFetch.mockReturnValue(fetchResponse(400, { error: 'invalid' }));
    expect(Code.verifyIdToken('bad')).toBeNull();
    expect(Code.verifyIdToken('')).toBeNull();
  });
});

describe('doPost', () => {
  const event = (obj) => ({ postData: { contents: JSON.stringify(obj) } });

  it('returns unauthorized for a bad token', () => {
    mockFetch.mockReturnValue(fetchResponse(400, {}));
    const out = Code.doPost(event({ idToken: 'bad', text: 'hi' }));
    expect(JSON.parse(out.text)).toEqual({ ok: false, error: 'unauthorized' });
  });

  it('posts and verifies for the allowed user', () => {
    mockFetch
      .mockReturnValueOnce(fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' }))
      .mockReturnValueOnce(fetchResponse(201, { data: { id: '555', text: 'hi' } }))
      .mockReturnValueOnce(fetchResponse(200, { data: { id: '555', text: 'hi' } }));
    const out = Code.doPost(event({ idToken: 'good', text: 'hi' }));
    const body = JSON.parse(out.text);
    expect(body.ok).toBe(true);
    expect(body.id).toBe('555');
    expect(body.url).toBe('https://x.com/freddiefujiwara/status/555');
    expect(body.verifiedText).toBe('hi');
  });
});

describe('spreadsheet logging', () => {
  it('logs the full API response of a post to the log sheet', () => {
    mockFetch.mockReturnValue(fetchResponse(201, { data: { id: '123', text: 'hello' } }));
    Code.postTweet('hello');
    expect(appendedRows).toHaveLength(1);
    const [time, action, detail] = appendedRows[0];
    expect(action).toBe('fetch');
    const parsed = JSON.parse(detail);
    expect(parsed.method).toBe('POST');
    expect(parsed.url).toBe('https://api.twitter.com/2/tweets');
    expect(parsed.status).toBe(201);
    expect(parsed.response).toContain('"id":"123"');
    expect(typeof time).toBe('string');
  });

  it('does not log DM response bodies to the log sheet', () => {
    mockFetch.mockReturnValue(
      fetchResponse(200, { data: [{ id: '1', text: 'secret dm body' }] })
    );
    Code.signedGet('/dm_events', { max_results: '5' }, {
      apiKey: 'k',
      apiKeySecret: 's',
      accessToken: 't',
      accessTokenSecret: 'ts',
    });
    expect(appendedRows).toHaveLength(1);
    const [, action, detail] = appendedRows[0];
    expect(action).toBe('fetch');
    const parsed = JSON.parse(detail);
    expect(parsed.status).toBe(200);
    expect(parsed.response).toBe('(not logged: DM endpoint)');
    expect(detail).not.toContain('secret dm body');
  });

  it('strips the query string (ID token) from logged URLs', () => {
    mockFetch.mockReturnValue(
      fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' })
    );
    Code.verifyIdToken('SECRET_TOKEN_VALUE');
    const detail = JSON.parse(appendedRows[0][2]);
    expect(detail.url).toBe('https://oauth2.googleapis.com/tokeninfo');
    expect(detail.response).not.toContain('SECRET_TOKEN_VALUE');
  });

  it('creates the log sheet with a header when missing', () => {
    sheetExists = false;
    mockFetch.mockReturnValue(fetchResponse(201, { data: { id: '1', text: 'x' } }));
    Code.postTweet('x');
    expect(appendedRows[0]).toEqual(['time', 'action', 'detail']);
    expect(appendedRows).toHaveLength(2);
  });

  it('never breaks posting when logging fails', () => {
    global.SpreadsheetApp = {
      getActiveSpreadsheet: () => {
        throw new Error('no spreadsheet');
      },
    };
    mockFetch.mockReturnValue(fetchResponse(201, { data: { id: '7', text: 'ok' } }));
    expect(Code.postTweet('ok').id).toBe('7');
  });
});

describe('doGet', () => {
  it('returns a health payload', () => {
    const out = Code.doGet();
    const body = JSON.parse(out.text);
    expect(body.ok).toBe(true);
    expect(body.service).toBe('tweet-gas');
  });
});

describe('doPost share sessions', () => {
  const event = (obj) => ({ postData: { contents: JSON.stringify(obj) } });
  const post = (obj) => JSON.parse(Code.doPost(event(obj)).text);
  const shareHashOf = (token) =>
    crypto.createHash('sha256').update(String(token), 'utf8').digest('hex');

  // Routes fetches by URL: tokeninfo for auth, X API for posting.
  function mockAllFetches() {
    mockFetch.mockImplementation((url, options) => {
      if (String(url).includes('tokeninfo')) {
        return fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' });
      }
      if (options && options.method === 'post') {
        return fetchResponse(201, { data: { id: '555', text: 'hi' } });
      }
      return fetchResponse(200, { data: { id: '555', text: 'hi' } });
    });
  }

  const issueToken = () => {
    mockAllFetches();
    const result = post({ idToken: 'good', action: 'issue_share' });
    expect(result.ok).toBe(true);
    return result.share_token;
  };

  it('requires an ID token to issue a share session', () => {
    mockFetch.mockReturnValue(fetchResponse(400, {}));
    expect(post({ idToken: 'bad', action: 'issue_share' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('requires an ID token to revoke share sessions', () => {
    mockFetch.mockReturnValue(fetchResponse(400, {}));
    expect(post({ idToken: 'bad', action: 'revoke_share' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('issues a 30-day share session and stores only its hash', () => {
    const before = Date.now();
    mockAllFetches();
    const result = post({ idToken: 'good', action: 'issue_share' });
    expect(result.ok).toBe(true);
    expect(result.ttl_seconds).toBe(30 * 24 * 60 * 60);
    expect(result.share_token).toMatch(/^[0-9a-f]{64}$/);
    expect(Date.parse(result.expires_at) - before).toBeGreaterThan(29 * 24 * 60 * 60 * 1000);
    expect(Date.parse(result.expires_at) - before).toBeLessThanOrEqual(
      30 * 24 * 60 * 60 * 1000 + 5000
    );
    expect(shareValues).toHaveLength(2);
    expect(shareValues[1][0]).toBe(shareHashOf(result.share_token));
    expect(shareValues[1][0]).not.toBe(result.share_token);
    expect(shareValues[1][4]).toBe(0);
  });

  it('creates the share sheet with headers when it does not exist yet', () => {
    shareValues = [];
    shareSheetExists = false;
    mockAllFetches();
    const result = post({ idToken: 'good', action: 'issue_share' });
    expect(result.ok).toBe(true);
    expect(shareValues[0]).toEqual(SHARE_HEADERS);
    expect(shareValues).toHaveLength(2);
  });

  it('revokes the previous session when a new one is issued', () => {
    const first = issueToken();
    const second = issueToken();
    expect(first).not.toBe(second);
    expect(shareValues[1][3]).not.toBe('');
    expect(shareValues[2][3]).toBe('');
    expect(post({ share_token: first, text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('posts with a valid share token and counts the access', () => {
    const token = issueToken();
    mockAllFetches();
    const result = post({ share_token: token, text: 'hi' });
    expect(result.ok).toBe(true);
    expect(result.id).toBe('555');
    expect(result.url).toBe('https://x.com/freddiefujiwara/status/555');
    expect(result._share.expires_at).toBe(shareValues[1][2]);
    expect(shareValues[1][4]).toBe(1);
    expect(shareValues[1][5]).not.toBe('');
  });

  it('validates a share token without posting', () => {
    const token = issueToken();
    mockAllFetches();
    const result = post({ share_token: token, action: 'validate_share' });
    expect(result).toEqual({
      ok: true,
      _share: { expires_at: shareValues[1][2] },
    });
    // Only the issue-time tokeninfo fetch plus none for posting.
    const postCalls = mockFetch.mock.calls.filter(
      ([url]) => String(url) === 'https://api.twitter.com/2/tweets'
    );
    expect(postCalls).toHaveLength(0);
  });

  it('rejects an expired share token', () => {
    const token = 'expired-token';
    shareValues.push([
      shareHashOf(token),
      new Date(Date.now() - 3600_000).toISOString(),
      new Date(Date.now() - 1800_000).toISOString(),
      '',
      0,
      '',
    ]);
    expect(post({ share_token: token, text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('rejects a revoked share token', () => {
    const token = 'revoked-token';
    shareValues.push([
      shareHashOf(token),
      new Date().toISOString(),
      new Date(Date.now() + 1800_000).toISOString(),
      new Date().toISOString(),
      0,
      '',
    ]);
    expect(post({ share_token: token, text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('rejects an unknown share token', () => {
    expect(post({ share_token: 'no-such-token', text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('revokes the active session via the revoke action', () => {
    const token = issueToken();
    mockAllFetches();
    expect(post({ idToken: 'good', action: 'revoke_share' })).toEqual({
      ok: true,
      revoked: true,
    });
    expect(post({ share_token: token, text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });
});

describe('doPost owner sessions', () => {
  const event = (obj) => ({ postData: { contents: JSON.stringify(obj) } });
  const post = (obj) => JSON.parse(Code.doPost(event(obj)).text);
  const hashOf = (token) =>
    crypto.createHash('sha256').update(String(token), 'utf8').digest('hex');

  function mockAllFetches() {
    mockFetch.mockImplementation((url, options) => {
      if (String(url).includes('tokeninfo')) {
        return fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' });
      }
      if (options && options.method === 'post') {
        return fetchResponse(201, { data: { id: '555', text: 'hi' } });
      }
      return fetchResponse(200, { data: { id: '555', text: 'hi' } });
    });
  }

  const exchange = () => {
    mockAllFetches();
    const result = post({ idToken: 'good', action: 'exchange_session' });
    expect(result.ok).toBe(true);
    return result.session_token;
  };

  it('requires an ID token to exchange for an owner session', () => {
    mockFetch.mockReturnValue(fetchResponse(400, {}));
    expect(post({ idToken: 'bad', action: 'exchange_session' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('issues a 30-day owner session and stores only its hash', () => {
    const token = exchange();
    const result = post({ idToken: 'good', action: 'exchange_session' });
    expect(result.ttl_seconds).toBe(30 * 24 * 60 * 60);
    expect(token).toMatch(/^[0-9a-f]{64}$/);
    expect(ownerValues).toHaveLength(3);
    expect(ownerValues[2][0]).toBe(hashOf(result.session_token));
    expect(ownerValues[2][0]).not.toBe(result.session_token);
  });

  it('posts with a valid owner session token', () => {
    const token = exchange();
    mockAllFetches();
    const result = post({ session_token: token, text: 'hi' });
    expect(result.ok).toBe(true);
    expect(result.id).toBe('555');
    expect(ownerValues[1][4]).toBe(1);
  });

  it('rejects an unknown owner session token', () => {
    expect(post({ session_token: 'no-such-token', text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('revokes the previous owner session when a new one is exchanged', () => {
    const first = exchange();
    const second = exchange();
    expect(first).not.toBe(second);
    expect(post({ session_token: first, text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
    mockAllFetches();
    expect(post({ session_token: second, text: 'hi' }).ok).toBe(true);
  });

  it('revokes the owner session via the revoke_session action', () => {
    const token = exchange();
    expect(post({ session_token: token, action: 'revoke_session' })).toEqual({
      ok: true,
      revoked: true,
    });
    expect(post({ session_token: token, text: 'hi' })).toEqual({
      ok: false,
      error: 'unauthorized',
    });
  });

  it('accepts an owner session as proof for issuing a share link', () => {
    const token = exchange();
    const result = post({ session_token: token, action: 'issue_share' });
    expect(result.ok).toBe(true);
    expect(result.share_token).toMatch(/^[0-9a-f]{64}$/);
  });
});

describe('staging environment separation', () => {
  it('postTweet returns a dry-run result without credentials or an X API call when DRY_RUN=1', () => {
    mockGetProperty.mockImplementation((name) => (name === 'DRY_RUN' ? '1' : null));
    const result = Code.postTweet('hello staging');
    expect(result).toEqual({ id: 'dry-run', text: 'hello staging', url: null, dry_run: true });
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('postTweet still posts normally when DRY_RUN is unset', () => {
    mockFetch.mockReturnValue(fetchResponse(201, { data: { id: '123', text: 'hello' } }));
    const result = Code.postTweet('hello');
    expect(result.id).toBe('123');
    expect(result.dry_run).toBeUndefined();
  });

  it('logToSheet opens the spreadsheet by SPREADSHEET_ID when set', () => {
    const byIdSheet = { appendRow: vi.fn() };
    const openById = vi.fn(() => ({ getSheetByName: () => byIdSheet, insertSheet: () => byIdSheet }));
    global.SpreadsheetApp = {
      getActiveSpreadsheet: () => {
        throw new Error('must not use the container spreadsheet');
      },
      openById,
    };
    mockGetProperty.mockImplementation((name) => (name === 'SPREADSHEET_ID' ? 'stg-file-id' : null));
    Code.logToSheet('test', { ok: true });
    expect(openById).toHaveBeenCalledWith('stg-file-id');
    expect(byIdSheet.appendRow).toHaveBeenCalled();
  });

  it('doPost in dry-run mode posts nothing and reports dry_run', () => {
    const event = (obj) => ({ postData: { contents: JSON.stringify(obj) } });
    mockGetProperty.mockImplementation((name) => (name === 'DRY_RUN' ? 'true' : null));
    mockFetch.mockImplementation((url) => {
      if (String(url).startsWith('https://oauth2.googleapis.com/tokeninfo')) {
        return fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' });
      }
      throw new Error('unexpected fetch: ' + url);
    });
    const out = Code.doPost(event({ idToken: 'tok', text: 'hello dry' }));
    expect(JSON.parse(out.text)).toMatchObject({
      ok: true,
      id: 'dry-run',
      dry_run: true,
      postedText: 'hello dry',
    });
    const xCalls = mockFetch.mock.calls.filter(([url]) => String(url).includes('api.twitter.com'));
    expect(xCalls).toHaveLength(0);
  });
});

// --- Machine sync token / SNS snapshot --------------------------------

describe('sns_snapshot sync action', () => {
  const event = (obj) => ({ postData: { contents: JSON.stringify(obj) } });
  const post = (obj) => JSON.parse(Code.doPost(event(obj)).text);

  const PROPS = {
    ...KEYS,
    SYNC_TOKEN: 'sync-secret',
  };

  const ME = {
    data: {
      id: '1234567890',
      username: 'freddiefujiwara',
      public_metrics: { followers_count: 321, following_count: 45 },
    },
  };
  const TWEETS = {
    data: [
      {
        id: '2001',
        created_at: '2026-11-09T10:00:00.000Z',
        text: '川口のブログを更新しました',
        public_metrics: {
          impression_count: 500,
          like_count: 12,
          retweet_count: 3,
          reply_count: 2,
        },
      },
      {
        id: '2002',
        created_at: '2026-11-08T10:00:00.000Z',
        text: '表示回数だけ無い投稿',
        public_metrics: { like_count: 1, retweet_count: 0, reply_count: 0 },
      },
    ],
  };

  beforeEach(() => {
    mockGetProperty.mockImplementation((name) => PROPS[name] ?? null);
  });

  it('rejects a wrong sync token with 401 and never calls the X API', () => {
    const out = post({ sync_token: 'wrong', action: 'sns_snapshot' });
    expect(out).toMatchObject({ ok: false, status: 401, error: 'sync token invalid' });
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('rejects every other action for the sync token with 403', () => {
    const out = post({ sync_token: 'sync-secret', text: '投稿されてはいけない' });
    expect(out).toMatchObject({ ok: false, status: 403 });
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('rejects the sync action while SYNC_TOKEN is unset', () => {
    mockGetProperty.mockImplementation((name) => KEYS[name] ?? null);
    const out = post({ sync_token: 'sync-secret', action: 'sns_snapshot' });
    expect(out).toMatchObject({ ok: false, status: 401 });
  });

  it('returns followers and post metrics for a valid sync call', () => {
    mockFetch.mockImplementation((url) => {
      const u = String(url);
      if (u.startsWith('https://api.twitter.com/2/users/me')) {
        return fetchResponse(200, ME);
      }
      if (u.startsWith('https://api.twitter.com/2/users/1234567890/tweets')) {
        return fetchResponse(200, TWEETS);
      }
      throw new Error('unexpected fetch: ' + u);
    });
    const out = post({ sync_token: 'sync-secret', action: 'sns_snapshot' });
    expect(out.ok).toBe(true);
    expect(out._sync).toBe(true);
    expect(out.x.followers).toBe(321);
    expect(out.x.username).toBe('freddiefujiwara');
    expect(out.x.recent_posts).toHaveLength(2);
    expect(out.x.recent_posts[0]).toEqual({
      id: '2001',
      posted_at: '2026-11-09T10:00:00.000Z',
      text: '川口のブログを更新しました',
      views: 500,
      likes: 12,
      reposts: 3,
      replies: 2,
    });
    // A metric the API did not report stays null, never 0.
    expect(out.x.recent_posts[1].views).toBeNull();
    expect(out.x_error).toBeNull();
    expect(out.posts_error).toBeNull();
    // Both requests are signed (OAuth Authorization header present).
    for (const [, options] of mockFetch.mock.calls) {
      expect(options.headers.Authorization).toMatch(/^OAuth /);
    }
  });

  it('keeps the followers when the posts fetch fails', () => {
    mockFetch.mockImplementation((url) => {
      const u = String(url);
      if (u.startsWith('https://api.twitter.com/2/users/me')) {
        return fetchResponse(200, ME);
      }
      return fetchResponse(429, {
        errors: [{ code: 88, message: 'Rate limit exceeded' }],
      });
    });
    const out = post({ sync_token: 'sync-secret', action: 'sns_snapshot' });
    expect(out.ok).toBe(true);
    expect(out.x.followers).toBe(321);
    expect(out.x.recent_posts).toBeNull();
    expect(out.x_error).toBeNull();
    expect(out.posts_error).toContain('HTTP 429');
    expect(out.posts_error).toContain('Rate limit exceeded');
  });

  it('reports a profile failure on both error fields', () => {
    mockFetch.mockImplementation(() =>
      fetchResponse(401, { title: 'Unauthorized', detail: 'bad token' })
    );
    const out = post({ sync_token: 'sync-secret', action: 'sns_snapshot' });
    expect(out.ok).toBe(true);
    expect(out.x.followers).toBeNull();
    expect(out.x.recent_posts).toBeNull();
    expect(out.x_error).toContain('HTTP 401');
    expect(out.posts_error).toBe(out.x_error);
  });

  it('returns nulls without fetching under DRY_RUN', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'DRY_RUN' ? '1' : (PROPS[name] ?? null)
    );
    const out = post({ sync_token: 'sync-secret', action: 'sns_snapshot' });
    expect(out.ok).toBe(true);
    expect(out.dry_run).toBe(true);
    expect(out.x).toEqual({ followers: null, username: null, recent_posts: null });
    expect(mockFetch).not.toHaveBeenCalled();
  });
});

// --- DM access probe -------------------------------------------------

describe('probeDmAccess', () => {
  const DM_BODY = {
    data: [
      {
        id: 'dm-event-1',
        event_type: 'JoinConversation',
        text: 'SECRET DM TEXT',
        sender_id: '4242424242',
        dm_conversation_id: 'conv-987654321',
      },
    ],
    meta: { result_count: 1 },
  };

  it('calls dm_events once, signed, and returns only status/ok/count on 200', () => {
    mockFetch.mockReturnValue(fetchResponse(200, DM_BODY));
    const result = Code.probeDmAccess();
    expect(result).toEqual({ status: 200, ok: true, count: 1 });
    expect(mockFetch).toHaveBeenCalledTimes(1);
    const [url, options] = mockFetch.mock.calls[0];
    expect(url).toBe('https://api.x.com/2/dm_events?max_results=1');
    expect(options.method).toBe('get');
    expect(options.headers.Authorization).toMatch(/^OAuth /);
  });

  it('never leaks DM content into the result, the Logger or the log sheet', () => {
    mockFetch.mockReturnValue(fetchResponse(200, DM_BODY));
    const result = Code.probeDmAccess();
    const surfaces =
      JSON.stringify(result) +
      JSON.stringify(appendedRows) +
      JSON.stringify(Logger.log.mock.calls);
    expect(surfaces).not.toContain('SECRET DM TEXT');
    expect(surfaces).not.toContain('4242424242');
    expect(surfaces).not.toContain('conv-987654321');
    expect(surfaces).not.toContain('dm-event-1');
  });

  it('extracts only the error title/detail on 403', () => {
    mockFetch.mockReturnValue(
      fetchResponse(403, {
        title: 'Forbidden',
        detail: 'This endpoint is not accessible with the current app permission',
        resource_id: 'must-not-be-copied',
      })
    );
    const result = Code.probeDmAccess();
    expect(result).toEqual({
      status: 403,
      ok: false,
      title: 'Forbidden',
      detail: 'This endpoint is not accessible with the current app permission',
    });
    expect(JSON.stringify(result)).not.toContain('must-not-be-copied');
  });

  it('falls back to errors[0].message when title/detail are absent', () => {
    mockFetch.mockReturnValue(
      fetchResponse(403, { errors: [{ code: 200, message: 'Forbidden.' }] })
    );
    const result = Code.probeDmAccess();
    expect(result).toEqual({ status: 403, ok: false, detail: 'Forbidden.' });
  });

  it('never calls the X API under DRY_RUN', () => {
    mockGetProperty.mockImplementation((name) => (name === 'DRY_RUN' ? '1' : null));
    const result = Code.probeDmAccess();
    expect(result).toEqual({ dry_run: true });
    expect(mockFetch).not.toHaveBeenCalled();
  });
});

describe('summarizeDmProbeResponse', () => {
  it('omits count when a 200 body has no data array', () => {
    expect(Code.summarizeDmProbeResponse(200, '{"meta":{"result_count":0}}')).toEqual({
      status: 200,
      ok: true,
    });
  });

  it('tolerates a non-JSON error body', () => {
    expect(Code.summarizeDmProbeResponse(500, 'oops')).toEqual({
      status: 500,
      ok: false,
    });
  });
});
