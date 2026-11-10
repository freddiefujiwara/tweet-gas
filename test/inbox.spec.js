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

// Generic in-memory spreadsheet: any number of named sheets, each
// backed by a plain values matrix (unlike Code.spec.js, whose mocks
// only model the log and share_sessions sheets).
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

// --- Fixtures ------------------------------------------------------------

const hoursAgo = (h) => new Date(Date.now() - h * 3600_000).toISOString();
const daysAgo = (d) => new Date(Date.now() - d * 24 * 3600_000).toISOString();

const MENTIONS = {
  data: [
    {
      id: '1001',
      author_id: '222',
      text: '@freddiefujiwara 参考になりました！',
      created_at: hoursAgo(2),
      conversation_id: '900',
      in_reply_to_user_id: '111',
      referenced_tweets: [{ type: 'replied_to', id: '900' }],
    },
    {
      // The owner's own post showing up in mentions: excluded.
      id: '1002',
      author_id: '111',
      text: '自分の投稿',
      created_at: hoursAgo(1),
      conversation_id: '1002',
    },
    {
      // A repost of someone else's tweet: excluded.
      id: '1003',
      author_id: '333',
      text: 'RT @someone すごい',
      created_at: hoursAgo(1),
      conversation_id: '1003',
      referenced_tweets: [{ type: 'retweeted', id: '700' }],
    },
    {
      // Legacy RT without referenced_tweets: excluded too.
      id: '1005',
      author_id: '444',
      text: 'RT @someone すごい',
      created_at: hoursAgo(1),
      conversation_id: '1005',
    },
  ],
  includes: {
    users: [
      { id: '222', username: 'alice', name: 'Alice A' },
      { id: '333', username: 'carol', name: 'Carol C' },
      { id: '444', username: 'dave', name: 'Dave D' },
    ],
  },
};

const OWN_POSTS = {
  data: [
    { id: '900', created_at: daysAgo(1), text: '一昨日の投稿' },
    // Older than the 7-day search window: never a search target.
    { id: '800', created_at: daysAgo(30), text: '古い投稿' },
  ],
};

const SEARCH_RESULTS = {
  data: [
    {
      // Same tweet id as a mention: must be deduplicated.
      id: '1001',
      author_id: '222',
      text: '@freddiefujiwara 参考になりました！',
      created_at: hoursAgo(2),
      conversation_id: '900',
      referenced_tweets: [{ type: 'replied_to', id: '900' }],
    },
    {
      // A direct reply that never @-mentions the owner: only the
      // conversation search finds this one.
      id: '1004',
      author_id: '444',
      text: 'メンションなしの直返信です',
      created_at: hoursAgo(1),
      conversation_id: '900',
      referenced_tweets: [{ type: 'replied_to', id: '900' }],
    },
  ],
  includes: {
    users: [
      { id: '222', username: 'alice', name: 'Alice A' },
      { id: '444', username: 'dave', name: 'Dave D' },
    ],
  },
};

function routeFetch(overrides = {}) {
  mockFetch.mockImplementation((url, options) => {
    const u = String(url);
    if (u.includes('tokeninfo')) {
      return fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' });
    }
    if (u.startsWith('https://api.twitter.com/2/users/me')) {
      return fetchResponse(200, { data: { id: '111', username: 'freddiefujiwara' } });
    }
    if (u.includes('/mentions')) {
      return overrides.mentions || fetchResponse(200, MENTIONS);
    }
    if (u.includes('/tweets/search/recent')) {
      return overrides.search
        ? overrides.search(u)
        : fetchResponse(200, SEARCH_RESULTS);
    }
    if (u.startsWith('https://api.twitter.com/2/users/111/tweets')) {
      return overrides.ownPosts || fetchResponse(200, OWN_POSTS);
    }
    if (u === 'https://api.twitter.com/2/tweets' && options && options.method === 'post') {
      const payload = JSON.parse(options.payload);
      return fetchResponse(201, { data: { id: '2001', text: payload.text } });
    }
    if (u.startsWith('https://api.twitter.com/2/tweets/')) {
      return fetchResponse(200, { data: { id: u.split('/').pop(), text: 'ありがとうございます！' } });
    }
    throw new Error('unexpected fetch: ' + u);
  });
}

const event = (obj) => ({ postData: { contents: JSON.stringify(obj) } });
const post = (obj) => JSON.parse(Code.doPost(event(obj)).text);
const xCalls = () =>
  mockFetch.mock.calls.filter(([url]) => String(url).includes('api.twitter.com'));

function seedInbox(rows) {
  const sheet = makeSheet('inbox');
  sheet.values = [[...Code.INBOX_HEADERS], ...rows.map((row) => [...row])];
  sheets.set('inbox', sheet);
  return sheet;
}

const inboxRowFor = (commentId) =>
  (sheets.get('inbox')?.values || []).find((row) => String(row[0]) === commentId);

// --- normalize / weighted length / target selection ----------------------

describe('normalizeInboxComment', () => {
  const users = { 222: { username: 'alice', name: 'Alice A' } };

  it('maps a mention tweet into an inbox comment with author info', () => {
    const comment = Code.normalizeInboxComment(MENTIONS.data[0], users, '111');
    expect(comment).toEqual({
      comment_id: '1001',
      author_id: '222',
      author_username: 'alice',
      author_name: 'Alice A',
      text: '@freddiefujiwara 参考になりました！',
      created_at: MENTIONS.data[0].created_at,
      in_reply_to_tweet_id: '900',
      conversation_id: '900',
    });
  });

  it('excludes the owner own posts and reposts', () => {
    expect(Code.normalizeInboxComment(MENTIONS.data[1], users, '111')).toBeNull();
    expect(Code.normalizeInboxComment(MENTIONS.data[2], users, '111')).toBeNull();
    expect(Code.normalizeInboxComment(MENTIONS.data[3], users, '111')).toBeNull();
  });
});

describe('weightedTweetLength / validateTweetText', () => {
  it('weighs latin as 1, CJK as 2 and a URL as 23', () => {
    expect(Code.weightedTweetLength('hello')).toBe(5);
    expect(Code.weightedTweetLength('あいう')).toBe(6);
    expect(Code.weightedTweetLength('see https://example.com/a/very/long/path')).toBe(27);
  });

  it('rejects empty and over-280 weighted text', () => {
    expect(() => Code.validateTweetText('  ')).toThrow(/empty/);
    expect(() => Code.validateTweetText('a'.repeat(281))).toThrow(/too long/);
    expect(() => Code.validateTweetText('あ'.repeat(141))).toThrow(/too long/);
    expect(Code.validateTweetText('あ'.repeat(140))).toBe('あ'.repeat(140));
  });
});

describe('selectConversationSearchTargets', () => {
  it('keeps only posts inside the 7-day window, newest first, capped', () => {
    const posts = [];
    for (let i = 0; i < 15; i += 1) {
      posts.push({ id: String(1000 + i), created_at: hoursAgo(i + 1) });
    }
    posts.push({ id: 'old', created_at: daysAgo(8) });
    const targets = Code.selectConversationSearchTargets(posts, Date.now());
    expect(targets).toHaveLength(Code.INBOX_SEARCH_POSTS_LIMIT);
    expect(targets[0]).toBe('1000');
    expect(targets).not.toContain('old');
  });
});

// --- list_comments --------------------------------------------------------

describe('list_comments action', () => {
  it('merges mentions and conversation replies, deduplicates and stores them', () => {
    routeFetch();
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(true);
    expect(out.note).toBeUndefined();
    // Newest first; the search-only reply (1004) comes first, the
    // duplicated mention/search hit (1001) appears exactly once,
    // and own posts / reposts never show up.
    expect(out.comments.map((c) => c.comment_id)).toEqual(['1004', '1001']);
    expect(out.comments[1]).toMatchObject({
      author_username: 'alice',
      author_name: 'Alice A',
      in_reply_to_tweet_id: '900',
      conversation_id: '900',
      status: 'new',
      our_reply_id: '',
    });
    // The mentions request carries the required expansions/fields.
    const mentionsCall = mockFetch.mock.calls.find(([url]) =>
      String(url).includes('/mentions')
    );
    expect(String(mentionsCall[0])).toContain('expansions=author_id');
    expect(String(mentionsCall[0])).toContain('referenced_tweets');
    // The search targets only the recent post (900), never the old one (800).
    const searchCalls = mockFetch.mock.calls.filter(([url]) =>
      String(url).includes('/tweets/search/recent')
    );
    expect(searchCalls).toHaveLength(1);
    expect(decodeURIComponent(String(searchCalls[0][0]))).toContain(
      'query=conversation_id:900 is:reply'
    );
    // Stored in the inbox sheet with the full column layout.
    const values = sheets.get('inbox').values;
    expect(values[0]).toEqual(Code.INBOX_HEADERS);
    expect(values).toHaveLength(3);
    expect(inboxRowFor('1001')[8]).toBe('new');
    expect(inboxRowFor('1001')[10]).not.toBe(''); // first_seen_at
    expect(inboxRowFor('1001')[11]).not.toBe(''); // last_seen_at
  });

  it('keeps the stored status and our_reply_id when collecting again', () => {
    routeFetch();
    post({ idToken: 'good', action: 'list_comments' });
    const row = inboxRowFor('1001');
    row[8] = 'replied';
    row[9] = '2000';
    const out = post({ idToken: 'good', action: 'list_comments' });
    const again = out.comments.find((c) => c.comment_id === '1001');
    expect(again.status).toBe('replied');
    expect(again.our_reply_id).toBe('2000');
    // No duplicate row was appended for the same comment.
    expect(sheets.get('inbox').values).toHaveLength(3);
    expect(inboxRowFor('1001')[8]).toBe('replied');
    expect(inboxRowFor('1001')[9]).toBe('2000');
  });

  it('degrades to mentions with a note when the search endpoint fails', () => {
    routeFetch({
      search: () =>
        fetchResponse(403, {
          errors: [{ code: 453, message: 'endpoint not available on this plan' }],
        }),
    });
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(true);
    expect(out.comments.map((c) => c.comment_id)).toEqual(['1001']);
    expect(out.note).toMatch(/search/);
  });

  it('degrades with a note when even the own-posts fetch fails', () => {
    routeFetch({ ownPosts: fetchResponse(429, { errors: [{ message: 'rate limited' }] }) });
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(true);
    expect(out.comments.map((c) => c.comment_id)).toEqual(['1001']);
    expect(out.note).toMatch(/mentions only/);
  });

  it('fails the action when mentions itself fails', () => {
    routeFetch({ mentions: fetchResponse(401, { title: 'Unauthorized' }) });
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/mentions failed/);
  });

  it('returns an empty dry-run result without any X API call', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'DRY_RUN' ? '1' : (KEYS[name] ?? null)
    );
    routeFetch();
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out).toEqual({ ok: true, dry_run: true, comments: [] });
    expect(xCalls()).toHaveLength(0);
    expect(sheets.get('inbox')).toBeUndefined();
  });

  it('cannot be reached through the sync token', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'SYNC_TOKEN' ? 'sync-secret' : (KEYS[name] ?? null)
    );
    const out = post({ sync_token: 'sync-secret', action: 'list_comments' });
    expect(out).toMatchObject({ ok: false, status: 403 });
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('requires user authentication', () => {
    mockFetch.mockReturnValue(fetchResponse(400, {}));
    const out = post({ idToken: 'bad', action: 'list_comments' });
    expect(out).toEqual({ ok: false, error: 'unauthorized' });
  });
});

describe('upsertInboxComments', () => {
  it('preserves status / our_reply_id of an existing row', () => {
    seedInbox([
      ['1001', '222', 'alice', 'Alice A', 'old text', hoursAgo(5), '900', '900',
        'replied', '2000', hoursAgo(5), hoursAgo(5)],
    ]);
    const [merged] = Code.upsertInboxComments([
      {
        comment_id: '1001',
        author_id: '222',
        author_username: 'alice',
        author_name: 'Alice A',
        text: 'new text',
        created_at: hoursAgo(5),
        in_reply_to_tweet_id: '900',
        conversation_id: '900',
      },
    ]);
    expect(merged.status).toBe('replied');
    expect(merged.our_reply_id).toBe('2000');
    const row = inboxRowFor('1001');
    expect(row[8]).toBe('replied');
    expect(row[9]).toBe('2000');
    expect(row[4]).toBe('new text'); // descriptive fields do refresh
    expect(sheets.get('inbox').values).toHaveLength(2);
  });
});

// --- reply_comment --------------------------------------------------------

describe('reply_comment action', () => {
  const seedRow = () =>
    seedInbox([
      ['1001', '222', 'alice', 'Alice A', 'コメント', hoursAgo(3), '900', '900',
        'new', '', hoursAgo(3), hoursAgo(3)],
    ]);

  it('posts the reply payload, verifies it, and marks the inbox row replied', () => {
    seedRow();
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_comment',
      comment_id: '1001',
      text: 'ありがとうございます！',
    });
    expect(out.ok).toBe(true);
    expect(out.id).toBe('2001');
    expect(out.url).toBe('https://x.com/freddiefujiwara/status/2001');
    expect(out.verifiedText).toBe('ありがとうございます！');
    const postCall = mockFetch.mock.calls.find(
      ([url, options]) =>
        String(url) === 'https://api.twitter.com/2/tweets' && options.method === 'post'
    );
    expect(JSON.parse(postCall[1].payload)).toEqual({
      text: 'ありがとうございます！',
      reply: { in_reply_to_tweet_id: '1001' },
    });
    const row = inboxRowFor('1001');
    expect(row[8]).toBe('replied');
    expect(row[9]).toBe('2001');
    // The reply is recorded on the log sheet like a normal post.
    const logRows = sheets.get('log').values;
    const replyLog = logRows.find((r) => r[1] === 'reply_comment');
    expect(replyLog).toBeTruthy();
    expect(JSON.parse(replyLog[2]).id).toBe('2001');
  });

  it('rejects text over the weighted 280 limit without calling X', () => {
    seedRow();
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_comment',
      comment_id: '1001',
      text: 'a'.repeat(281),
    });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/too long/);
    expect(xCalls()).toHaveLength(0);
    expect(inboxRowFor('1001')[8]).toBe('new');
  });

  it('rejects a missing comment_id without calling X', () => {
    routeFetch();
    const out = post({ idToken: 'good', action: 'reply_comment', text: 'hi' });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/comment_id/);
    expect(xCalls()).toHaveLength(0);
  });

  it('in dry-run mode posts nothing and leaves the inbox row untouched', () => {
    seedRow();
    mockGetProperty.mockImplementation((name) =>
      name === 'DRY_RUN' ? 'true' : (KEYS[name] ?? null)
    );
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_comment',
      comment_id: '1001',
      text: 'ありがとうございます！',
    });
    expect(out).toMatchObject({ ok: true, dry_run: true, id: 'dry-run' });
    expect(xCalls()).toHaveLength(0);
    const row = inboxRowFor('1001');
    expect(row[8]).toBe('new');
    expect(row[9]).toBe('');
  });
});
