import { describe, it, expect, vi, beforeEach } from 'vitest';
import crypto from 'node:crypto';
import * as Code from '../src/Code.js';

// Abnormal-case tests: deformed values arriving from outside this
// project — X API response JSON (tweet/user/metric objects, error
// bodies, usage payloads) and request bodies from the frontend
// (post text, reply/mark targets). A value whose toString is not a
// function (UNSTR below) is producible from JSON
// (JSON.parse('{"toString":[]}')) and makes String()/Number() throw
// "Cannot convert object to primitive value"; arrays and booleans
// silently coerce (Number(true) === 1, String(['123']) === '123').
// One deformed item must never kill a whole list/snapshot, invent a
// metric, or become the target of a post/DM/mark.

// --- GAS service mocks (same generic-sheet harness as inbox.spec.js) ---

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
  X_CONSUMER_KEY: 'ck',
  X_CONSUMER_SECRET: 'cs',
  X_ACCESS_TOKEN: 'at',
  X_ACCESS_TOKEN_SECRET: 'as',
};

function fetchResponse(code, obj) {
  return {
    getResponseCode: () => code,
    getContentText: () => (typeof obj === 'string' ? obj : JSON.stringify(obj)),
  };
}

const dryRun = () =>
  mockGetProperty.mockImplementation((name) =>
    name === 'DRY_RUN' ? '1' : (KEYS[name] ?? null)
  );

beforeEach(() => {
  vi.clearAllMocks();
  sheets.clear();
  uuidCounter = 0;
  mockGetProperty.mockImplementation((name) => KEYS[name] ?? null);
});

// A JSON-producible value whose toString is not a function.
const UNSTR = { toString: [] };

describe('metrics from the X API (snapshot posts, usage, quote targets)', () => {
  it('reads boolean / array / blank metrics as null, never as invented numbers', () => {
    for (const bad of [true, false, [5], [], '', ' ']) {
      const post = Code.buildSnsSnapshotPost({
        id: '1',
        public_metrics: { impression_count: bad },
      });
      expect(post.views).toBeNull();
    }
  });

  it('does not throw on an unstringifiable metric; it reads as null', () => {
    const post = Code.buildSnsSnapshotPost({
      id: '1',
      public_metrics: { impression_count: UNSTR },
    });
    expect(post.views).toBeNull();
    expect(Code.buildUsageCredits({ total_balance: UNSTR }).credit_balance_usd).toBeNull();
    expect(
      Code.buildUsageCredits({ free_grants: [{ amount: UNSTR }] }).free_grants
    ).toEqual([{ amount_usd: null, expires_at: null }]);
    expect(
      Code.buildUsageDaily({
        daily_project_usage: [{ date: '2026-01-01', usage: UNSTR }],
      })
    ).toEqual([{ date: '2026-01-01', tweets_consumed: null }]);
    expect(
      Code.buildUsageDaily({
        daily_project_usage: [
          { date: '2026-01-01', usage: [{ tweets_consumed: UNSTR }, { tweets_consumed: 2 }] },
        ],
      })
    ).toEqual([{ date: '2026-01-01', tweets_consumed: 2 }]);
  });

  it('reads boolean / array usage amounts as null, never as invented money', () => {
    expect(Code.buildUsageCredits({ total_balance: true }).credit_balance_usd).toBeNull();
    expect(
      Code.buildUsageCredits({ free_grants: [{ amount: [500] }] }).free_grants
    ).toEqual([{ amount_usd: null, expires_at: null }]);
  });

  it('keeps genuine metrics: numbers and numeric strings pass through', () => {
    const post = Code.buildSnsSnapshotPost({
      id: '1',
      public_metrics: { impression_count: 42, like_count: '5' },
    });
    expect(post.views).toBe(42);
    expect(post.likes).toBe(5);
  });

  it('selectQuoteTargets survives a deformed quote_count and keeps valid targets', () => {
    const now = new Date().toISOString();
    const posts = [
      { id: 'bad', created_at: now, public_metrics: { quote_count: UNSTR } },
      { id: 'good', created_at: now, public_metrics: { quote_count: 3 } },
    ];
    expect(Code.selectQuoteTargets(posts)).toEqual(['good']);
  });
});

describe('SNS snapshot post shaping', () => {
  it('does not throw on unstringifiable id / text; they read as null / empty', () => {
    expect(Code.buildSnsSnapshotPost({ id: UNSTR }).id).toBeNull();
    expect(Code.buildSnsSnapshotPost({ id: '1', text: UNSTR }).text).toBe('');
  });

  it('does not leak an array text as "a,b"', () => {
    expect(Code.buildSnsSnapshotPost({ id: '1', text: ['a', 'b'] }).text).toBe('');
  });

  it('carries posted_at only from strings', () => {
    expect(Code.buildSnsSnapshotPost({ id: '1', created_at: 12345 }).posted_at).toBeNull();
    expect(Code.buildSnsSnapshotPost({ id: '1', created_at: UNSTR }).posted_at).toBeNull();
    expect(
      Code.buildSnsSnapshotPost({ id: '1', created_at: '2026-01-01T00:00:00.000Z' }).posted_at
    ).toBe('2026-01-01T00:00:00.000Z');
  });
});

describe('comment inbox shaping', () => {
  it('does not throw on unstringifiable text; it reads as empty', () => {
    const comment = Code.normalizeInboxComment(
      { id: '9', author_id: '2', text: UNSTR },
      {},
      '111'
    );
    expect(comment.text).toBe('');
  });

  it('keeps every field a string: deformed created_at / ids read as empty', () => {
    const comment = Code.normalizeInboxComment(
      {
        id: '9',
        author_id: UNSTR,
        text: 'x',
        created_at: UNSTR,
        conversation_id: ['42'],
        referenced_tweets: [{ type: 'replied_to', id: UNSTR }],
      },
      {},
      '111'
    );
    for (const key of [
      'comment_id',
      'author_id',
      'author_username',
      'author_name',
      'text',
      'created_at',
      'in_reply_to_tweet_id',
      'conversation_id',
    ]) {
      expect(typeof comment[key]).toBe('string');
    }
    expect(comment.created_at).toBe('');
    expect(comment.conversation_id).toBe('');
  });

  it('determineInboxKind does not throw on an unstringifiable quoted id', () => {
    const kind = Code.determineInboxKind(
      { referenced_tweets: [{ type: 'quoted', id: UNSTR }] },
      ['999']
    );
    expect(kind).toBe('mention');
  });

  it('determineInboxKind does not throw on a non-iterable ownPostIds', () => {
    const kind = Code.determineInboxKind(
      { referenced_tweets: [{ type: 'quoted', id: '999' }] },
      42
    );
    expect(kind).toBe('mention');
  });

  it('one deformed mention does not kill list_comments', async () => {
    mockFetch.mockImplementation((url) => {
      if (url.includes('/users/me')) {
        return fetchResponse(200, { data: { id: '111', username: 'freddiefujiwara' } });
      }
      if (url.includes('/mentions')) {
        return fetchResponse(200, {
          data: [
            {
              id: '2001',
              author_id: '222',
              text: 'normal comment',
              created_at: '2026-11-19T00:00:00.000Z',
              conversation_id: '2001',
            },
            {
              id: '2002',
              author_id: '333',
              text: UNSTR,
              created_at: '2026-11-19T01:00:00.000Z',
              conversation_id: '2002',
            },
          ],
          includes: {
            users: [
              { id: '222', username: 'alice', name: 'Alice' },
              { id: '333', username: UNSTR, name: 'Bob' },
            ],
          },
        });
      }
      // the owner's own posts fetch: empty is fine
      return fetchResponse(200, { data: [] });
    });
    const result = await Code.listComments();
    expect(result.ok).toBe(true);
    expect(result.comments).toHaveLength(2);
    const deformed = result.comments.find((c) => c.comment_id === '2002');
    expect(deformed.text).toBe('');
    expect(deformed.author_username).toBe('');
  });
});

describe('DM inbox shaping', () => {
  it('does not throw on unstringifiable text / participants; created_at stays a string', () => {
    const event = Code.normalizeDmEvent(
      {
        id: '1',
        dm_conversation_id: 'c1',
        sender_id: '2',
        text: UNSTR,
        created_at: UNSTR,
        participant_ids: [UNSTR, '2'],
      },
      {},
      '111'
    );
    expect(event.text).toBe('');
    expect(event.created_at).toBe('');
    expect(event.participant_ids).toEqual(['2']);
  });

  it('a normalized batch with deformed fields still builds conversations', () => {
    const events = [
      Code.normalizeDmEvent(
        { id: '1', dm_conversation_id: 'c1', sender_id: '2', text: 'a', created_at: UNSTR },
        {},
        '111'
      ),
      Code.normalizeDmEvent(
        {
          id: '2',
          dm_conversation_id: 'c1',
          sender_id: '2',
          text: 'b',
          created_at: '2026-01-01T00:00:00.000Z',
        },
        {},
        '111'
      ),
    ];
    const conversations = Code.buildDmConversations(events, new Map(), {}, '111');
    expect(conversations).toHaveLength(1);
    expect(conversations[0].messages).toHaveLength(2);
  });

  it('buildDmConversations skips non-object entries instead of throwing', () => {
    const valid = {
      event_id: '1',
      conversation_id: 'c1',
      sender_id: '2',
      sender_username: null,
      sender_name: null,
      text: 'hi',
      created_at: '2026-01-01T00:00:00.000Z',
      is_mine: false,
      participant_ids: [],
    };
    const conversations = Code.buildDmConversations([null, 42, valid], new Map(), {}, '111');
    expect(conversations).toHaveLength(1);
  });
});

describe('X API error summarizers', () => {
  it('does not throw on unstringifiable message / detail / code', () => {
    expect(
      Code.summarizeXApiError(403, '{"errors":[{"code":88,"message":{"toString":[]}}]}')
    ).toBe('HTTP 403 code=88');
    expect(Code.summarizeXApiError(403, '{"detail":{"toString":[]}}')).toBe('HTTP 403');
    expect(
      Code.summarizeXApiError(403, '{"errors":[{"code":{"toString":[]}}]}')
    ).toBe('HTTP 403');
  });

  it('summarizeDmProbeResponse does not throw on an unstringifiable title', () => {
    expect(Code.summarizeDmProbeResponse(403, '{"title":{"toString":[]}}')).toEqual({
      status: 403,
      ok: false,
    });
  });
});

describe('usage daily dates', () => {
  it('skips days that do not exist instead of carrying them verbatim', () => {
    const daily = Code.buildUsageDaily({
      daily_project_usage: [
        { date: '2026-02-30T00:00:00.000Z', usage: [{ tweets_consumed: 3 }] },
        { date: '2026-13-01T00:00:00.000Z', usage: [{ tweets_consumed: 3 }] },
        { date: '2023-02-29T00:00:00.000Z', usage: [{ tweets_consumed: 3 }] },
        { date: '2024-02-29T00:00:00.000Z', usage: [{ tweets_consumed: 3 }] },
        { date: '2026-11-10T00:00:00.000Z', usage: [{ tweets_consumed: 3 }] },
      ],
    });
    expect(daily).toEqual([
      { date: '2024-02-29', tweets_consumed: 3 },
      { date: '2026-11-10', tweets_consumed: 3 },
    ]);
  });
});

describe('action target ids from request bodies', () => {
  it('postReply rejects array / unstringifiable comment ids instead of coercing them', () => {
    dryRun();
    expect(() => Code.postReply('hi', ['123'])).toThrow('comment_id is empty');
    expect(() => Code.postReply('hi', UNSTR)).toThrow('comment_id is empty');
    expect(() => Code.postReply('hi', { id: '123' })).toThrow('comment_id is empty');
  });

  it('postReply still accepts string and number comment ids', () => {
    dryRun();
    expect(Code.postReply('hi', '123').dry_run).toBe(true);
    expect(Code.postReply('hi', 123).dry_run).toBe(true);
  });

  it('markDmRead rejects array / unstringifiable ids instead of coercing them', () => {
    dryRun();
    expect(() => Code.markDmRead(['c1'], 'e1')).toThrow('conversation_id is empty');
    expect(() => Code.markDmRead('c1', UNSTR)).toThrow('event_id is empty');
    expect(Code.markDmRead('c1', 'e1').dry_run).toBe(true);
  });

  it('postDmMessage rejects an array conversation id instead of sending into it', () => {
    dryRun();
    expect(() => Code.postDmMessage(['c1'], 'hi')).toThrow('conversation_id is empty');
    expect(() => Code.postDmMessage(UNSTR, 'hi')).toThrow('conversation_id is empty');
    expect(Code.postDmMessage('c1', 'hi').dry_run).toBe(true);
  });
});

describe('sheet upserts at the payload boundary', () => {
  it('upsertInboxComments skips non-object entries instead of throwing', () => {
    const valid = {
      comment_id: '100',
      author_id: '2',
      author_username: 'u',
      author_name: 'n',
      text: 'hi',
      created_at: '2026-01-01T00:00:00.000Z',
      in_reply_to_tweet_id: '',
      conversation_id: '100',
      kind: 'mention',
    };
    const merged = Code.upsertInboxComments([null, 42, 'x', valid]);
    expect(merged).toHaveLength(1);
    expect(merged[0].comment_id).toBe('100');
    expect(merged[0].status).toBe('new');
  });

  it('upsertDmEvents already tolerates a null entry (pinned)', () => {
    expect(() => Code.upsertDmEvents([null])).not.toThrow();
  });
});

describe('conversation search target selection', () => {
  it('skips posts whose created_at is not a string instead of throwing', () => {
    const recent = new Date().toISOString();
    const posts = [
      { id: 'bad', created_at: UNSTR },
      { id: 'num', created_at: 12345 },
      { id: 'good', created_at: recent },
    ];
    expect(Code.selectConversationSearchTargets(posts, Date.now())).toEqual(['good']);
  });
});

describe('text validation boundaries (pinned behaviour)', () => {
  it('accepts exactly 280 weighted and rejects 281', () => {
    expect(Code.validateTweetText('a'.repeat(280))).toBe('a'.repeat(280));
    expect(() => Code.validateTweetText('a'.repeat(281))).toThrow('too long');
    expect(Code.validateTweetText('あ'.repeat(140))).toBe('あ'.repeat(140));
    expect(() => Code.validateTweetText('あ'.repeat(141))).toThrow('too long');
    // 256 letters + space + one URL (23) = exactly 280
    expect(Code.validateTweetText('a'.repeat(256) + ' https://x.co')).toContain('https://');
    expect(() => Code.validateTweetText('a'.repeat(257) + ' https://x.co')).toThrow('too long');
    // astral characters weigh 2 each: 140 emoji = exactly 280
    expect(Code.validateTweetText('🎉'.repeat(140))).toBe('🎉'.repeat(140));
    expect(() => Code.validateTweetText('🎉'.repeat(141))).toThrow('too long');
  });

  it('rejects non-string texts as empty (the caller-side guard for the counter)', () => {
    expect(() => Code.validateTweetText(UNSTR)).toThrow('text is empty');
    expect(() => Code.validateTweetText(42)).toThrow('text is empty');
    expect(() => Code.validateTweetText(['a'])).toThrow('text is empty');
  });

  it('counts a lone surrogate without throwing', () => {
    expect(Code.weightedTweetLength('abc\uD800')).toBe(5);
    expect(Code.validateTweetText('abc\uD800')).toBe('abc\uD800');
  });

  it('validateDmText counts code points: 10000 pass, 10001 throw', () => {
    expect(Code.validateDmText('🎉'.repeat(10000))).toBe('🎉'.repeat(10000));
    expect(() => Code.validateDmText('🎉'.repeat(10001))).toThrow('too long');
    expect(() => Code.validateDmText(UNSTR)).toThrow('text is empty');
  });
});
