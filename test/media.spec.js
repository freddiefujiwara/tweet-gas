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
  base64Decode: (text) => [...Buffer.from(text, 'base64')],
  newBlob: (bytes, contentType, name) => ({ bytes, contentType, name }),
  getUuid: () =>
    `00000000-0000-4000-8000-${String(uuidCounter++).padStart(12, '0')}`,
  formatDate: () => '2026-11-01 05:00:00',
};
global.PropertiesService = {
  getScriptProperties: () => ({ getProperty: mockGetProperty }),
};
global.UrlFetchApp = { fetch: mockFetch };

const sheets = new Map();
function makeSheet(name) {
  const sheet = {
    name,
    values: [],
    appendRow(row) {
      this.values.push([...row]);
    },
    getDataRange() {
      return { getValues: () => this.values };
    },
    getRange(row, col) {
      return {
        setValue: (value) => {
          while (sheet.values.length < row) sheet.values.push([]);
          sheet.values[row - 1][col - 1] = value;
        },
      };
    },
  };
  return sheet;
}
const spreadsheet = {
  getSheetByName: (name) => sheets.get(name) || null,
  insertSheet: (name) => {
    const sheet = makeSheet(name);
    sheets.set(name, sheet);
    return sheet;
  },
};
global.SpreadsheetApp = {
  getActiveSpreadsheet: () => spreadsheet,
  openById: () => spreadsheet,
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
  ALLOWED_EMAIL: 'owner@example.com',
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
  sheets.clear();
  uuidCounter = 0;
  mockGetProperty.mockImplementation((name) => KEYS[name] ?? null);
});

// A tiny valid-looking PNG payload (content does not matter here).
const PNG = { data: Buffer.from([0x89, 0x50, 0x4e, 0x47]).toString('base64'), type: 'image/png' };

// --- Tests ---------------------------------------------------------------

describe('validateMediaItem', () => {
  it('accepts a supported image item', () => {
    expect(Code.validateMediaItem(PNG)).toBe(PNG);
    expect(Code.validateMediaItem({ data: 'AAAA', type: 'image/jpeg' })).toBeTruthy();
  });

  it('rejects an unsupported type', () => {
    expect(() => Code.validateMediaItem({ data: 'AAAA', type: 'video/mp4' })).toThrow(
      /unsupported media type/
    );
  });

  it('rejects empty data and non-objects', () => {
    expect(() => Code.validateMediaItem({ data: '', type: 'image/png' })).toThrow(/empty/);
    expect(() => Code.validateMediaItem(null)).toThrow(/object/);
  });

  it('rejects an item over the size limit', () => {
    const big = { data: 'A'.repeat(7_000_001), type: 'image/png' };
    expect(() => Code.validateMediaItem(big)).toThrow(/too large/);
  });
});

describe('uploadMedia', () => {
  it('uploads as multipart with an OAuth header and returns the media id', () => {
    mockFetch.mockReturnValue(fetchResponse(200, { data: { id: 'media-1', media_key: '3_x' } }));
    expect(Code.uploadMedia(PNG)).toBe('media-1');
    const [url, options] = mockFetch.mock.calls[0];
    expect(url).toBe('https://api.x.com/2/media/upload');
    expect(options.headers.Authorization).toMatch(/^OAuth /);
    expect(options.payload.media_category).toBe('tweet_image');
    expect(options.payload.media.contentType).toBe('image/png');
  });

  it('accepts the legacy media_id_string response shape', () => {
    mockFetch.mockReturnValue(fetchResponse(200, { media_id_string: 'legacy-9' }));
    expect(Code.uploadMedia(PNG)).toBe('legacy-9');
  });

  it('throws on an upload error', () => {
    mockFetch.mockReturnValue(fetchResponse(403, { detail: 'forbidden' }));
    expect(() => Code.uploadMedia(PNG)).toThrow(/media upload failed \(403\)/);
  });
});

describe('postTweet with media', () => {
  it('uploads each image first, then posts with media_ids', () => {
    mockFetch
      .mockReturnValueOnce(fetchResponse(200, { data: { id: 'm1' } }))
      .mockReturnValueOnce(fetchResponse(200, { data: { id: 'm2' } }))
      .mockReturnValueOnce(fetchResponse(201, { data: { id: '123', text: 'hello' } }));
    const result = Code.postTweet('hello', [PNG, PNG]);
    expect(result.id).toBe('123');
    expect(mockFetch.mock.calls).toHaveLength(3);
    const [url, options] = mockFetch.mock.calls[2];
    expect(url).toBe('https://api.twitter.com/2/tweets');
    expect(JSON.parse(options.payload)).toEqual({
      text: 'hello',
      media: { media_ids: ['m1', 'm2'] },
    });
  });

  it('posts without media exactly as before', () => {
    mockFetch.mockReturnValue(fetchResponse(201, { data: { id: '123', text: 'hello' } }));
    Code.postTweet('hello');
    const [, options] = mockFetch.mock.calls[0];
    expect(JSON.parse(options.payload)).toEqual({ text: 'hello' });
  });

  it('rejects more than 4 images before any API call', () => {
    expect(() => Code.postTweet('hello', [PNG, PNG, PNG, PNG, PNG])).toThrow(/too many media/);
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('returns a dry-run result without uploading when DRY_RUN=1', () => {
    mockGetProperty.mockImplementation((name) => (name === 'DRY_RUN' ? '1' : null));
    const result = Code.postTweet('hello staging', [PNG]);
    expect(result).toEqual({ id: 'dry-run', text: 'hello staging', url: null, dry_run: true });
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('still validates media in dry-run mode', () => {
    mockGetProperty.mockImplementation((name) => (name === 'DRY_RUN' ? '1' : null));
    expect(() => Code.postTweet('hello', [{ data: 'AAAA', type: 'video/mp4' }])).toThrow(
      /unsupported media type/
    );
  });
});
