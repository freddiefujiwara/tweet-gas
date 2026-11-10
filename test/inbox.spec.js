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
    if (u.includes('/quote_tweets')) {
      return overrides.quote
        ? overrides.quote(u)
        : fetchResponse(200, { data: [] });
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
      // The mention fixture replies to post 900, so its kind is
      // "reply" (the shape gained the kind field with the quote
      // source; see the kind tests below).
      kind: 'reply',
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

// --- Quote source (third collection path) --------------------------------

const OWN_POSTS_QUOTED = {
  data: [
    {
      id: '900',
      created_at: daysAgo(1),
      text: '一昨日の投稿',
      public_metrics: { quote_count: 3, retweet_count: 0, reply_count: 1 },
    },
  ],
};

const QUOTE_RESULTS = {
  data: [
    {
      id: '3001',
      author_id: '555',
      text: 'これは参考になる',
      created_at: hoursAgo(0.5),
      conversation_id: '3001',
      referenced_tweets: [{ type: 'quoted', id: '900' }],
    },
    {
      // A quote with no text of its own: the quote itself counts.
      id: '3002',
      author_id: '666',
      text: '',
      created_at: hoursAgo(4),
      conversation_id: '3002',
      referenced_tweets: [{ type: 'quoted', id: '900' }],
    },
    {
      // The same tweet id as the mention/search hit 1001, this
      // time also quoting post 900: deduplicated, kind upgraded.
      id: '1001',
      author_id: '222',
      text: '@freddiefujiwara 参考になりました！',
      created_at: hoursAgo(2),
      conversation_id: '900',
      referenced_tweets: [
        { type: 'replied_to', id: '900' },
        { type: 'quoted', id: '900' },
      ],
    },
  ],
  includes: {
    users: [
      { id: '555', username: 'eve', name: 'Eve E' },
      { id: '666', username: 'frank', name: 'Frank F' },
      { id: '222', username: 'alice', name: 'Alice A' },
    ],
  },
};

const quoteCalls = () =>
  mockFetch.mock.calls.filter(([url]) => String(url).includes('/quote_tweets'));

describe('kind determination', () => {
  const base = {
    id: '4001',
    author_id: '222',
    text: 'x',
    created_at: hoursAgo(1),
    conversation_id: '4001',
  };

  it('prefers quote when a tweet both replies to and quotes an own post', () => {
    const tweet = {
      ...base,
      referenced_tweets: [
        { type: 'replied_to', id: '900' },
        { type: 'quoted', id: '900' },
      ],
    };
    const comment = Code.normalizeInboxComment(tweet, {}, '111', {
      ownPostIds: ['900', '901'],
    });
    expect(comment.kind).toBe('quote');
    expect(Code.determineInboxKind(tweet, ['900'], 'reply')).toBe('quote');
  });

  it('does not call a quote of someone else a quote kind', () => {
    const tweet = {
      ...base,
      referenced_tweets: [
        { type: 'replied_to', id: '900' },
        { type: 'quoted', id: '999' },
      ],
    };
    expect(Code.determineInboxKind(tweet, ['900'])).toBe('reply');
  });

  it('classifies a plain mention as mention and a plain reply as reply', () => {
    expect(Code.normalizeInboxComment({ ...base }, {}, '111').kind).toBe('mention');
    const reply = {
      ...base,
      referenced_tweets: [{ type: 'replied_to', id: '900' }],
    };
    expect(Code.normalizeInboxComment(reply, {}, '111').kind).toBe('reply');
  });
});

describe('selectQuoteTargets', () => {
  it('keeps quoted posts of any age, newest first, and skips quote_count 0', () => {
    const posts = [
      { id: 'a', created_at: daysAgo(10), public_metrics: { quote_count: 0 } },
      { id: 'b', created_at: daysAgo(20), public_metrics: { quote_count: 2 } },
      { id: 'c', created_at: daysAgo(5), public_metrics: { quote_count: 1 } },
      { id: 'd', created_at: daysAgo(40), public_metrics: { quote_count: 9 } },
      { id: 'e', created_at: daysAgo(3) },
      { id: 'f', created_at: daysAgo(200), public_metrics: { quote_count: 3 } },
    ];
    expect(Code.selectQuoteTargets(posts)).toEqual(['c', 'b', 'd', 'f']);
  });

  it('caps the calls at 10, newest quoted posts first', () => {
    const posts = [];
    for (let i = 0; i < 20; i += 1) {
      posts.push({
        id: String(2000 + i),
        created_at: hoursAgo(i + 1),
        public_metrics: { quote_count: 1 },
      });
    }
    const targets = Code.selectQuoteTargets(posts);
    expect(targets).toHaveLength(Code.INBOX_QUOTE_CALLS_LIMIT);
    expect(targets[0]).toBe('2000');
  });

  it('considers quoted posts beyond the newest 15 posts', () => {
    const posts = [];
    for (let i = 0; i < 17; i += 1) {
      posts.push({
        id: String(3000 + i),
        created_at: daysAgo(i + 1),
        public_metrics: { quote_count: 0 },
      });
    }
    posts.push({
      id: '3999',
      created_at: daysAgo(60),
      public_metrics: { quote_count: 5 },
    });
    expect(Code.selectQuoteTargets(posts)).toEqual(['3999']);
  });
});

describe('list_comments quote source', () => {
  it('merges quotes, deduplicates by tweet id, and records kind', () => {
    routeFetch({
      ownPosts: fetchResponse(200, OWN_POSTS_QUOTED),
      quote: () => fetchResponse(200, QUOTE_RESULTS),
    });
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(true);
    expect(out.note).toBeUndefined();
    // Newest first; the text-less quote (3002) is included, and
    // 1001 appears exactly once although three sources returned it.
    expect(out.comments.map((c) => c.comment_id)).toEqual(['3001', '1004', '1001', '3002']);
    const byCommentId = Object.fromEntries(out.comments.map((c) => [c.comment_id, c]));
    expect(byCommentId['3001'].kind).toBe('quote');
    expect(byCommentId['3002'].kind).toBe('quote');
    expect(byCommentId['3002'].text).toBe('');
    expect(byCommentId['1004'].kind).toBe('reply');
    // Quote wins for the tweet that also arrived as a reply.
    expect(byCommentId['1001'].kind).toBe('quote');
    // Exactly one quote_tweets call, for post 900, with the
    // required expansions and fields.
    expect(quoteCalls()).toHaveLength(1);
    const quoteUrl = String(quoteCalls()[0][0]);
    expect(quoteUrl).toContain('/tweets/900/quote_tweets');
    expect(quoteUrl).toContain('expansions=author_id');
    expect(quoteUrl).toContain('referenced_tweets');
    // The kind is stored as the last inbox column.
    expect(sheets.get('inbox').values[0]).toEqual(Code.INBOX_HEADERS);
    expect(Code.INBOX_HEADERS[12]).toBe('kind');
    expect(inboxRowFor('3001')[12]).toBe('quote');
    expect(inboxRowFor('1001')[12]).toBe('quote');
  });

  it('does not call quote_tweets for posts with quote_count 0', () => {
    routeFetch({
      ownPosts: fetchResponse(200, {
        data: [
          { id: '900', created_at: daysAgo(1), public_metrics: { quote_count: 0 } },
          { id: '901', created_at: daysAgo(2), public_metrics: {} },
        ],
      }),
    });
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(true);
    expect(quoteCalls()).toHaveLength(0);
  });

  it('caps quote_tweets calls at 10 even when many old posts are quoted', () => {
    const posts = [];
    for (let i = 0; i < 15; i += 1) {
      posts.push({
        // 35..49 days old: outside the old 30-day quote window
        // and outside the 7-day conversation search window, yet
        // still quote targets now that the age window is gone.
        id: String(5000 + i),
        created_at: daysAgo(35 + i),
        public_metrics: { quote_count: 1 },
      });
    }
    routeFetch({ ownPosts: fetchResponse(200, { data: posts }) });
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(true);
    expect(quoteCalls()).toHaveLength(Code.INBOX_QUOTE_CALLS_LIMIT);
  });

  it('degrades with a note when the quote endpoint fails', () => {
    routeFetch({
      ownPosts: fetchResponse(200, OWN_POSTS_QUOTED),
      quote: () => fetchResponse(403, { errors: [{ message: 'forbidden' }] }),
    });
    const out = post({ idToken: 'good', action: 'list_comments' });
    expect(out.ok).toBe(true);
    // Mentions plus conversation replies are still returned.
    expect(out.comments.map((c) => c.comment_id)).toEqual(['1004', '1001']);
    expect(out.note).toMatch(/quote/);
  });
});

describe('upsertInboxComments kind handling', () => {
  const commentWithKind = (kind) => ({
    comment_id: '1001',
    author_id: '222',
    author_username: 'alice',
    author_name: 'Alice A',
    text: 'text',
    created_at: hoursAgo(5),
    in_reply_to_tweet_id: '900',
    conversation_id: '900',
    kind,
  });

  it('never erases or downgrades a stored kind on re-collection', () => {
    seedInbox([
      ['1001', '222', 'alice', 'Alice A', 'old text', hoursAgo(5), '900', '900',
        'replied', '2000', hoursAgo(5), hoursAgo(5), 'quote'],
    ]);
    const [merged] = Code.upsertInboxComments([commentWithKind('mention')]);
    expect(merged.kind).toBe('quote');
    expect(merged.status).toBe('replied');
    expect(inboxRowFor('1001')[12]).toBe('quote');
  });

  it('fills the kind of a legacy row whose kind cell is still empty', () => {
    // A pre-kind row: only the original 12 columns are present.
    seedInbox([
      ['1001', '222', 'alice', 'Alice A', 'old text', hoursAgo(5), '900', '900',
        'new', '', hoursAgo(5), hoursAgo(5)],
    ]);
    const [merged] = Code.upsertInboxComments([commentWithKind('quote')]);
    expect(merged.kind).toBe('quote');
    expect(inboxRowFor('1001')[12]).toBe('quote');
  });

  it('adds the kind header to a sheet created before the column existed', () => {
    const sheet = makeSheet('inbox');
    sheet.values = [
      [...Code.INBOX_HEADERS.slice(0, 12)],
      ['1001', '222', 'alice', 'Alice A', 'old text', hoursAgo(5), '900', '900',
        'new', '', hoursAgo(5), hoursAgo(5)],
    ];
    sheets.set('inbox', sheet);
    Code.upsertInboxComments([commentWithKind('quote')]);
    expect(sheet.values[0]).toEqual(Code.INBOX_HEADERS);
    expect(inboxRowFor('1001')[12]).toBe('quote');
  });

  it('writes the kind when appending a new row', () => {
    const [merged] = Code.upsertInboxComments([commentWithKind('quote')]);
    expect(merged.kind).toBe('quote');
    const row = inboxRowFor('1001');
    expect(row).toHaveLength(13);
    expect(row[12]).toBe('quote');
  });
});

// --- reply_comment with as_quote ------------------------------------------

describe('reply_comment action with as_quote', () => {
  const seedRow = () =>
    seedInbox([
      ['1001', '222', 'alice', 'Alice A', 'コメント', hoursAgo(3), '900', '900',
        'new', '', hoursAgo(3), hoursAgo(3)],
    ]);

  it('posts a quote payload (quote_tweet_id, no reply), verifies and marks replied', () => {
    seedRow();
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_comment',
      comment_id: '1001',
      text: 'ありがとうございます！',
      as_quote: true,
    });
    expect(out.ok).toBe(true);
    expect(out.id).toBe('2001');
    expect(out.as_quote).toBe(true);
    expect(out.quote_tweet_id).toBe('1001');
    expect(out.in_reply_to_tweet_id).toBeUndefined();
    expect(out.verifiedText).toBe('ありがとうございます！');
    const postCall = mockFetch.mock.calls.find(
      ([url, options]) =>
        String(url) === 'https://api.twitter.com/2/tweets' && options.method === 'post'
    );
    const payload = JSON.parse(postCall[1].payload);
    expect(payload).toEqual({
      text: 'ありがとうございます！',
      quote_tweet_id: '1001',
    });
    expect(payload).not.toHaveProperty('reply');
    const row = inboxRowFor('1001');
    expect(row[8]).toBe('replied');
    expect(row[9]).toBe('2001');
  });

  it('rejects over-280 text in quote mode without calling X', () => {
    seedRow();
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_comment',
      comment_id: '1001',
      text: 'a'.repeat(281),
      as_quote: true,
    });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/too long/);
    expect(xCalls()).toHaveLength(0);
    expect(inboxRowFor('1001')[8]).toBe('new');
  });

  it('in dry-run mode quotes nothing and leaves the inbox row untouched', () => {
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
      as_quote: true,
    });
    expect(out).toMatchObject({
      ok: true,
      dry_run: true,
      id: 'dry-run',
      as_quote: true,
      quote_tweet_id: '1001',
    });
    expect(xCalls()).toHaveLength(0);
    const row = inboxRowFor('1001');
    expect(row[8]).toBe('new');
    expect(row[9]).toBe('');
  });
});
