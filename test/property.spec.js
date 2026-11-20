import { describe, it, expect, vi, beforeEach } from 'vitest';
import fc from 'fast-check';
import crypto from 'node:crypto';
import * as Code from '../src/Code.js';

// Property-based tests for the pure transforms that shape data
// from outside this project (X API JSON, request bodies): the
// shaping functions must never throw on JSON-producible values,
// their output fields must keep their documented types, the
// weighted counter must agree with the documented counting rules,
// and well-formed payloads must shape exactly as before.

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

beforeEach(() => {
  vi.clearAllMocks();
  sheets.clear();
  uuidCounter = 0;
  mockGetProperty.mockImplementation((name) => KEYS[name] ?? null);
});

// A JSON-producible value whose toString is not a function:
// String()/Number() throw on it. It is a legitimate member of the
// input domain of every function that reads X API JSON or a
// request body, because JSON.parse can produce it.
const UNSTR = { toString: [] };
const anyField = () =>
  fc.oneof(
    fc.string(),
    fc.double(),
    fc.boolean(),
    fc.constant(null),
    fc.constant(undefined),
    fc.array(fc.jsonValue()),
    fc.object(),
    fc.constant(UNSTR)
  );

const isFiniteOrNull = (v) => v === null || (typeof v === 'number' && Number.isFinite(v));

describe('weighted tweet length', () => {
  it('never throws and is monotone under appending one code point', () => {
    // The alphabet can form URLs, so appends exercise the
    // plain-text -> URL transition as well as plain growth.
    const alphabet = ['h', 't', 'p', 's', ':', '/', 'a', ' ', 'あ', '🎉'];
    fc.assert(
      fc.property(
        fc.array(fc.constantFrom(...alphabet), { maxLength: 60 }).map((a) => a.join('')),
        fc.constantFrom(...alphabet),
        (text, ch) => {
          const before = Code.weightedTweetLength(text);
          const after = Code.weightedTweetLength(text + ch);
          return (
            Number.isInteger(before) && before >= 0 && after >= before
          );
        }
      ),
      { numRuns: 100 }
    );
  });

  it('agrees with the documented per-code-point weights on URL-free text', () => {
    const weight = (cp) =>
      cp <= 0x10ff ||
      (cp >= 0x2000 && cp <= 0x200d) ||
      (cp >= 0x2010 && cp <= 0x201f) ||
      (cp >= 0x2032 && cp <= 0x2037)
        ? 1
        : 2;
    fc.assert(
      fc.property(
        fc
          .string({ unit: 'grapheme', maxLength: 80 })
          .filter((s) => !/https?:\/\//.test(s)),
        (text) => {
          const expected = [...text].reduce((sum, ch) => sum + weight(ch.codePointAt(0)), 0);
          return Code.weightedTweetLength(text) === expected;
        }
      ),
      { numRuns: 100 }
    );
  });

  it('validateTweetText throws "too long" exactly when the weight exceeds 280', () => {
    fc.assert(
      fc.property(fc.string({ unit: 'grapheme', maxLength: 400 }), (text) => {
        if (text.trim() === '') {
          return true; // the empty case is pinned in abnormal.spec.js
        }
        const length = Code.weightedTweetLength(text);
        let error = null;
        try {
          Code.validateTweetText(text);
        } catch (e) {
          error = e;
        }
        return length > 280
          ? error !== null && /too long/.test(error.message)
          : error === null;
      }),
      { numRuns: 100 }
    );
  });
});

describe('SNS snapshot post shaping', () => {
  it('never throws on JSON-domain tweets; output fields keep their types', () => {
    fc.assert(
      fc.property(
        fc.record({
          id: anyField(),
          text: anyField(),
          created_at: anyField(),
          public_metrics: fc.record({
            impression_count: anyField(),
            like_count: anyField(),
            retweet_count: anyField(),
            reply_count: anyField(),
          }),
        }),
        (tweet) => {
          const post = Code.buildSnsSnapshotPost(tweet);
          return (
            (post.id === null || typeof post.id === 'string') &&
            (post.posted_at === null || typeof post.posted_at === 'string') &&
            typeof post.text === 'string' &&
            isFiniteOrNull(post.views) &&
            isFiniteOrNull(post.likes) &&
            isFiniteOrNull(post.reposts) &&
            isFiniteOrNull(post.replies)
          );
        }
      ),
      { numRuns: 100 }
    );
  });

  it('shapes well-formed posts identically (metrics pass through)', () => {
    fc.assert(
      fc.property(
        fc.record({
          id: fc.string({ minLength: 1 }),
          text: fc.string(),
          created_at: fc.constant('2026-11-10T00:00:00.000Z'),
          views: fc.integer({ min: 0, max: 1_000_000 }),
          likes: fc.integer({ min: 0, max: 1_000_000 }),
        }),
        ({ id, text, created_at, views, likes }) => {
          const post = Code.buildSnsSnapshotPost({
            id,
            text,
            created_at,
            public_metrics: { impression_count: views, like_count: likes },
          });
          return (
            post.id === id &&
            post.text === text &&
            post.posted_at === created_at &&
            post.views === views &&
            post.likes === likes
          );
        }
      ),
      { numRuns: 100 }
    );
  });
});

describe('comment inbox shaping', () => {
  it('never throws on JSON-domain tweets; a kept comment has only string fields', () => {
    fc.assert(
      fc.property(
        fc.record({
          id: anyField(),
          author_id: anyField(),
          text: anyField(),
          created_at: anyField(),
          conversation_id: anyField(),
        }),
        fc.option(fc.dictionary(fc.string(), fc.record({ username: anyField(), name: anyField() })), {
          nil: undefined,
        }),
        (tweet, usersById) => {
          const comment = Code.normalizeInboxComment(tweet, usersById, '111');
          if (comment === null) {
            return true;
          }
          return [
            'comment_id',
            'author_id',
            'author_username',
            'author_name',
            'text',
            'created_at',
            'in_reply_to_tweet_id',
            'conversation_id',
            'kind',
          ].every((key) => typeof comment[key] === 'string');
        }
      ),
      { numRuns: 100 }
    );
  });
});

describe('DM event shaping and conversation building', () => {
  it('normalizeDmEvent never throws; kept events have string fields and string participants', () => {
    fc.assert(
      fc.property(
        fc.record({
          id: anyField(),
          dm_conversation_id: anyField(),
          sender_id: anyField(),
          text: anyField(),
          created_at: anyField(),
          participant_ids: fc.option(fc.array(anyField()), { nil: undefined }),
        }),
        (raw) => {
          const event = Code.normalizeDmEvent(raw, {}, '111');
          if (event === null) {
            return true;
          }
          return (
            [
              'event_id',
              'conversation_id',
              'sender_id',
              'text',
              'created_at',
            ].every((key) => typeof event[key] === 'string') &&
            event.participant_ids.every((id) => typeof id === 'string') &&
            typeof event.is_mine === 'boolean'
          );
        }
      ),
      { numRuns: 100 }
    );
  });

  it('conversations cap messages at 20, stay newest-first, and keep messages in their conversation', () => {
    const eventArb = fc.record({
      n: fc.integer({ min: 0, max: 10_000 }),
      conversation_id: fc.constantFrom('c1', 'c2', 'c3'),
      mine: fc.boolean(),
    });
    fc.assert(
      fc.property(fc.array(eventArb, { maxLength: 60 }), (specs) => {
        const events = specs.map(({ n, conversation_id, mine }) => ({
          event_id: String(1_000_000 + n),
          conversation_id,
          sender_id: mine ? '111' : '222',
          sender_username: 'someone',
          sender_name: 'Someone',
          text: `m${n}`,
          // one distinct minute per n keeps chronological order total
          created_at: new Date(Date.UTC(2026, 0, 1) + n * 60_000).toISOString(),
          is_mine: mine,
          participant_ids: ['111', '222'],
        }));
        const conversations = Code.buildDmConversations(events, new Map(), {}, '111');
        if (!conversations.every((c) => c.messages.length <= 20)) return false;
        if (
          !conversations.every((c) =>
            c.messages.every((m) => m.conversation_id === c.conversation_id)
          )
        ) {
          return false;
        }
        const totalMessages = conversations.reduce((sum, c) => sum + c.messages.length, 0);
        if (totalMessages > events.length) return false;
        for (let i = 1; i < conversations.length; i += 1) {
          const prev = conversations[i - 1].latest_event;
          const cur = conversations[i].latest_event;
          const order =
            prev.created_at === cur.created_at
              ? Code.compareDmEventIds(prev.event_id, cur.event_id)
              : prev.created_at < cur.created_at
                ? -1
                : 1;
          if (order < 0) return false;
        }
        return true;
      }),
      { numRuns: 100 }
    );
  });
});

describe('DM event id comparison', () => {
  it('agrees with BigInt ordering on decimal id strings and is antisymmetric', () => {
    const idArb = fc
      .array(fc.constantFrom('0', '1', '2', '3', '4', '5', '6', '7', '8', '9'), {
        minLength: 1,
        maxLength: 25,
      })
      .map((digits) => digits.join(''));
    fc.assert(
      fc.property(idArb, idArb, (a, b) => {
        const ab = Code.compareDmEventIds(a, b);
        const ba = Code.compareDmEventIds(b, a);
        const expected = BigInt(a) < BigInt(b) ? -1 : BigInt(a) > BigInt(b) ? 1 : 0;
        return ab === expected && ab + ba === 0;
      }),
      { numRuns: 100 }
    );
  });
});

describe('usage daily shaping', () => {
  const realDate = (y, m, d) => {
    const dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  };

  it('emits only real calendar dates, sorted ascending, and never throws', () => {
    fc.assert(
      fc.property(
        fc.array(
          fc.record({
            y: fc.integer({ min: 2023, max: 2027 }),
            m: fc.integer({ min: 1, max: 13 }),
            d: fc.integer({ min: 1, max: 31 }),
            usage: anyField(),
          }),
          { maxLength: 12 }
        ),
        (entries) => {
          const daily = Code.buildUsageDaily({
            daily_project_usage: entries.map(({ y, m, d, usage }) => ({
              date: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}T00:00:00.000Z`,
              usage: Array.isArray(usage) ? usage : [{ tweets_consumed: usage }],
            })),
          });
          if (!Array.isArray(daily)) return false;
          for (const entry of daily) {
            const [y, m, d] = entry.date.split('-').map(Number);
            if (!realDate(y, m, d)) return false;
            if (!isFiniteOrNull(entry.tweets_consumed)) return false;
          }
          for (let i = 1; i < daily.length; i += 1) {
            if (daily[i - 1].date > daily[i].date) return false;
          }
          return true;
        }
      ),
      { numRuns: 100 }
    );
  });

  it('sums the per-app counts of a well-formed day exactly', () => {
    fc.assert(
      fc.property(
        fc.array(fc.integer({ min: 0, max: 10_000 }), { maxLength: 8 }),
        (counts) => {
          const daily = Code.buildUsageDaily({
            daily_project_usage: [
              {
                date: '2026-11-10T00:00:00.000Z',
                usage: counts.map((tweets_consumed) => ({ tweets_consumed })),
              },
            ],
          });
          const expected = counts.length === 0 ? null : counts.reduce((a, b) => a + b, 0);
          return (
            Array.isArray(daily) &&
            daily.length === 1 &&
            daily[0].date === '2026-11-10' &&
            daily[0].tweets_consumed === expected
          );
        }
      ),
      { numRuns: 100 }
    );
  });
});
