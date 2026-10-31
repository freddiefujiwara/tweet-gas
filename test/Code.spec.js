import { describe, it, expect, vi, beforeEach } from 'vitest';
import crypto from 'node:crypto';
import * as Code from '../src/Code.js';

// --- GAS service mocks -------------------------------------------------

const mockGetProperty = vi.fn();
const mockFetch = vi.fn();

global.Utilities = {
  MacAlgorithm: { HMAC_SHA_1: 'HMAC_SHA_1' },
  computeHmacSignature: (algorithm, value, key) => {
    if (algorithm !== 'HMAC_SHA_1') throw new Error('unexpected algorithm: ' + algorithm);
    return crypto.createHmac('sha1', key).update(value, 'utf8').digest();
  },
  base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
  getUuid: () => '12345678-1234-4234-8234-123456789012',
  formatDate: () => '2026-11-01 05:00:00',
};
global.PropertiesService = {
  getScriptProperties: () => ({ getProperty: mockGetProperty }),
};
global.UrlFetchApp = { fetch: mockFetch };

const appendedRows = [];
let sheetExists = true;
const mockSheet = { appendRow: (row) => appendedRows.push(row) };
global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: () => (sheetExists ? mockSheet : null),
    insertSheet: () => {
      sheetExists = true;
      return mockSheet;
    },
  }),
};
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
