import { describe, it, expect, vi, beforeEach } from 'vitest';
import crypto from 'node:crypto';
import * as Code from '../src/Code.js';

// --- GAS service mocks (same generic-sheet setup as inbox.spec.js) ------

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

// --- Fixtures -------------------------------------------------------------
// Event ids are 19-digit numeric strings: beyond Number's safe range
// on purpose, so id ordering only works with a BigInt comparison.

const ME = { data: { id: '111', username: 'freddiefujiwara', name: 'Freddie F' } };

const USERS = [
  { id: '111', username: 'freddiefujiwara', name: 'Freddie F' },
  { id: '222', username: 'alice', name: 'Alice A' },
  { id: '333', username: 'bob', name: 'Bob B' },
];

const dmEvent = (id, conversationId, senderId, text, createdAt, extra = {}) => ({
  id,
  event_type: 'MessageCreate',
  text,
  created_at: createdAt,
  sender_id: senderId,
  dm_conversation_id: conversationId,
  participant_ids: conversationId === '100-200' ? ['111', '222'] : ['111', '333'],
  ...extra,
});

const DM_PAGE1 = {
  data: [
    dmEvent('1800000000000000001', '100-200', '222', 'こんにちは、質問があります', '2026-11-09T01:00:00.000Z'),
    dmEvent('1800000000000000003', '100-200', '111', 'ありがとうございます', '2026-11-09T02:00:00.000Z'),
    dmEvent('1800000000000000002', '100-300', '333', 'DM失礼します', '2026-11-09T03:00:00.000Z'),
  ],
  includes: { users: USERS },
  meta: { result_count: 3, next_token: 'page2-token' },
};

const DM_PAGE2 = {
  data: [
    dmEvent('1800000000000000004', '100-200', '222', '追加の質問です', '2026-11-09T04:00:00.000Z'),
    // Not a message: must be filtered out even though it is newest.
    dmEvent('1800000000000000005', '100-200', '222', '', '2026-11-09T05:00:00.000Z', {
      event_type: 'ParticipantsJoin',
    }),
  ],
  includes: { users: USERS },
  // A third page exists upstream; list_dms must still stop at two.
  meta: { result_count: 2, next_token: 'page3-token' },
};

function routeFetch(overrides = {}) {
  mockFetch.mockImplementation((url, options) => {
    const u = String(url);
    if (u.includes('tokeninfo')) {
      return fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' });
    }
    if (u.includes('/dm_conversations/') && options && options.method === 'post') {
      return overrides.dmPost
        ? overrides.dmPost(u, options)
        : fetchResponse(201, {
            data: { dm_conversation_id: '100-200', dm_event_id: '1800000000000000099' },
          });
    }
    if (u.startsWith('https://api.twitter.com/2/dm_events')) {
      if (u.includes('pagination_token=page2-token')) {
        return fetchResponse(200, DM_PAGE2);
      }
      return fetchResponse(200, overrides.dmPage1 || DM_PAGE1);
    }
    if (u.startsWith('https://api.twitter.com/2/users/me')) {
      return fetchResponse(200, ME);
    }
    throw new Error('unexpected fetch: ' + u);
  });
}

const event = (obj) => ({ postData: { contents: JSON.stringify(obj) } });
const post = (obj) => JSON.parse(Code.doPost(event(obj)).text);
const xCalls = () =>
  mockFetch.mock.calls.filter(([url]) => String(url).includes('api.twitter.com'));
const dmEventsCalls = () =>
  mockFetch.mock.calls.filter(([url]) => String(url).includes('/dm_events'));

function seedSheet(name, headers, rows) {
  const sheet = makeSheet(name);
  sheet.values = [[...headers], ...rows.map((row) => [...row])];
  sheets.set(name, sheet);
  return sheet;
}
const dmRowFor = (eventId) =>
  (sheets.get('dm')?.values || []).find((row) => String(row[0]) === eventId);
const stateRowFor = (conversationId) =>
  (sheets.get('dm_state')?.values || []).find((row) => String(row[0]) === conversationId);

// --- id comparison / normalization ---------------------------------------

describe('compareDmEventIds', () => {
  it('compares numerically, not lexicographically', () => {
    expect(Code.compareDmEventIds('999', '1000')).toBe(-1);
    expect(Code.compareDmEventIds('1000', '999')).toBe(1);
    expect(Code.compareDmEventIds('1000', '1000')).toBe(0);
  });

  it('keeps ids distinct beyond the Number safe range', () => {
    // Number('1800000000000000001') === Number('1800000000000000002').
    expect(Code.compareDmEventIds('1800000000000000001', '1800000000000000002')).toBe(-1);
    expect(Code.compareDmEventIds('1800000000000000002', '1800000000000000001')).toBe(1);
  });
});

describe('normalizeDmEvent', () => {
  const users = { 222: { username: 'alice', name: 'Alice A' } };

  it('maps a raw event into a message with sender info and is_mine', () => {
    const message = Code.normalizeDmEvent(DM_PAGE1.data[0], users, '111');
    expect(message).toMatchObject({
      event_id: '1800000000000000001',
      conversation_id: '100-200',
      sender_id: '222',
      sender_username: 'alice',
      sender_name: 'Alice A',
      text: 'こんにちは、質問があります',
      is_mine: false,
    });
    const own = Code.normalizeDmEvent(DM_PAGE1.data[1], users, '111');
    expect(own.is_mine).toBe(true);
    // Unknown sender info stays null per the API contract.
    expect(own.sender_username).toBeNull();
  });

  it('rejects non-message events and events without a conversation', () => {
    expect(Code.normalizeDmEvent(DM_PAGE2.data[1], users, '111')).toBeNull();
    expect(
      Code.normalizeDmEvent({ id: '1', event_type: 'MessageCreate' }, users, '111')
    ).toBeNull();
    expect(Code.normalizeDmEvent(null, users, '111')).toBeNull();
  });
});

describe('buildDmConversations', () => {
  const msg = (id, conversationId, senderId, createdAt) => ({
    event_id: id,
    conversation_id: conversationId,
    sender_id: senderId,
    sender_username: null,
    sender_name: null,
    text: 't' + id,
    created_at: createdAt,
    is_mine: senderId === '111',
    participant_ids: ['111', senderId],
  });

  it('caps messages at the newest 20, oldest first', () => {
    const events = [];
    for (let i = 1; i <= 25; i += 1) {
      events.push(
        msg(String(i), 'c1', '222', `2026-11-09T00:${String(i).padStart(2, '0')}:00.000Z`)
      );
    }
    const [conversation] = Code.buildDmConversations(events, new Map(), {}, '111');
    expect(conversation.messages).toHaveLength(20);
    expect(conversation.messages[0].event_id).toBe('6');
    expect(conversation.latest_event.event_id).toBe('25');
    // The public message shape carries no participant_ids.
    expect(conversation.latest_event.participant_ids).toBeUndefined();
  });

  it('derives unread / needs_reply from the state markers', () => {
    const events = [msg('10', 'c1', '222', '2026-11-09T01:00:00.000Z')];
    const stateWith = (read, replied) =>
      new Map([['c1', { last_read_event_id: read, last_replied_event_id: replied }]]);
    const flags = (state) =>
      Code.buildDmConversations(events, state, {}, '111').map((c) => [
        c.unread,
        c.needs_reply,
      ])[0];
    expect(flags(new Map())).toEqual([true, true]);
    expect(flags(stateWith('10', null))).toEqual([false, true]);
    expect(flags(stateWith(null, '10'))).toEqual([true, false]);
    expect(flags(stateWith('10', '10'))).toEqual([false, false]);
    // An older marker (numeric comparison: '9' < '10') keeps both on.
    expect(flags(stateWith('9', '9'))).toEqual([true, true]);
  });

  it('reports neither flag when the latest event is our own', () => {
    const events = [
      msg('10', 'c1', '222', '2026-11-09T01:00:00.000Z'),
      msg('11', 'c1', '111', '2026-11-09T02:00:00.000Z'),
    ];
    const [conversation] = Code.buildDmConversations(events, new Map(), {}, '111');
    expect(conversation.unread).toBe(false);
    expect(conversation.needs_reply).toBe(false);
  });
});

// --- list_dms --------------------------------------------------------------

describe('list_dms action', () => {
  it('groups events into conversations, newest first, and stores them', () => {
    routeFetch();
    const out = post({ idToken: 'good', action: 'list_dms' });
    expect(out.ok).toBe(true);
    expect(out.conversations.map((c) => c.conversation_id)).toEqual(['100-200', '100-300']);

    const first = out.conversations[0];
    expect(first.participants).toEqual([{ id: '222', name: 'Alice A', username: 'alice' }]);
    // Chronological (oldest first); the ParticipantsJoin event is gone.
    expect(first.messages.map((m) => m.event_id)).toEqual([
      '1800000000000000001',
      '1800000000000000003',
      '1800000000000000004',
    ]);
    expect(first.latest_event).toMatchObject({
      event_id: '1800000000000000004',
      sender_username: 'alice',
      is_mine: false,
    });
    expect(first.messages[1].is_mine).toBe(true);
    expect(first.unread).toBe(true);
    expect(first.needs_reply).toBe(true);

    const second = out.conversations[1];
    expect(second.participants).toEqual([{ id: '333', name: 'Bob B', username: 'bob' }]);
    expect(second.unread).toBe(true);

    // The dm_events request follows the agreed query contract.
    const firstCallUrl = String(dmEventsCalls()[0][0]);
    expect(firstCallUrl).toContain('event_types=MessageCreate');
    expect(firstCallUrl).toContain('max_results=100');
    expect(firstCallUrl).toContain('expansions=sender_id%2Cparticipant_ids');
    expect(firstCallUrl).toContain('participant_ids');
    expect(firstCallUrl).toContain('user.fields=name%2Cusername');

    // Stored in the dm sheet with the full column layout, text kept.
    const values = sheets.get('dm').values;
    expect(values[0]).toEqual(Code.DM_HEADERS);
    expect(values).toHaveLength(5); // header + 4 message events
    expect(dmRowFor('1800000000000000001')[5]).toBe('こんにちは、質問があります');
    expect(dmRowFor('1800000000000000003')[7]).toBe(true); // is_mine
    expect(dmRowFor('1800000000000000005')).toBeUndefined();
  });

  it('stops pagination after two pages even when a third is offered', () => {
    routeFetch();
    post({ idToken: 'good', action: 'list_dms' });
    expect(dmEventsCalls()).toHaveLength(2);
    expect(String(dmEventsCalls()[1][0])).toContain('pagination_token=page2-token');
    expect(
      mockFetch.mock.calls.some(([url]) => String(url).includes('page3-token'))
    ).toBe(false);
  });

  it('does not fetch a second page when there is no next_token', () => {
    routeFetch({ dmPage1: { ...DM_PAGE1, meta: { result_count: 3 } } });
    const out = post({ idToken: 'good', action: 'list_dms' });
    expect(out.ok).toBe(true);
    expect(dmEventsCalls()).toHaveLength(1);
  });

  it('fails the action when dm_events itself fails', () => {
    mockFetch.mockImplementation((url) => {
      const u = String(url);
      if (u.includes('tokeninfo')) {
        return fetchResponse(200, { email: 'zegenvs@gmail.com', email_verified: 'true' });
      }
      if (u.startsWith('https://api.twitter.com/2/users/me')) {
        return fetchResponse(200, ME);
      }
      return fetchResponse(403, { title: 'Forbidden', detail: 'no DM access' });
    });
    const out = post({ idToken: 'good', action: 'list_dms' });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/dm_events failed/);
  });

  it('keeps dm_state markers across a re-fetch and reflects them in the flags', () => {
    seedSheet('dm_state', Code.DM_STATE_HEADERS, [
      ['100-200', '1800000000000000004', '1800000000000000004', '2026-11-09T04:30:00.000Z'],
    ]);
    routeFetch();
    const out = post({ idToken: 'good', action: 'list_dms' });
    const first = out.conversations.find((c) => c.conversation_id === '100-200');
    expect(first.unread).toBe(false);
    expect(first.needs_reply).toBe(false);
    // The stored markers were not reset by collecting again.
    expect(stateRowFor('100-200')[1]).toBe('1800000000000000004');
    expect(stateRowFor('100-200')[2]).toBe('1800000000000000004');
    expect(sheets.get('dm_state').values).toHaveLength(2);
  });

  it('returns an empty dry-run result without any X call or sheet write', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'DRY_RUN' ? '1' : (KEYS[name] ?? null)
    );
    routeFetch();
    const out = post({ idToken: 'good', action: 'list_dms' });
    expect(out).toEqual({ ok: true, dry_run: true, conversations: [] });
    expect(xCalls()).toHaveLength(0);
    expect(sheets.get('dm')).toBeUndefined();
    expect(sheets.get('dm_state')).toBeUndefined();
  });

  it('cannot be reached through the sync token', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'SYNC_TOKEN' ? 'sync-secret' : (KEYS[name] ?? null)
    );
    const out = post({ sync_token: 'sync-secret', action: 'list_dms' });
    expect(out).toMatchObject({ ok: false, status: 403 });
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('requires user authentication', () => {
    mockFetch.mockReturnValue(fetchResponse(400, {}));
    const out = post({ idToken: 'bad', action: 'list_dms' });
    expect(out).toEqual({ ok: false, error: 'unauthorized' });
  });
});

describe('upsertDmEvents', () => {
  it('updates in place: no duplicate row, first_seen_at preserved, text refreshed', () => {
    seedSheet('dm', Code.DM_HEADERS, [
      ['1800000000000000001', '100-200', '222', 'alice', 'Alice A', 'old text',
        '2026-11-09T01:00:00.000Z', false, '2026-01-01T00:00:00.000Z', '2026-01-01T00:00:00.000Z'],
    ]);
    Code.upsertDmEvents([
      {
        event_id: '1800000000000000001',
        conversation_id: '100-200',
        sender_id: '222',
        sender_username: 'alice',
        sender_name: 'Alice A',
        text: 'new text',
        created_at: '2026-11-09T01:00:00.000Z',
        is_mine: false,
        participant_ids: ['111', '222'],
      },
    ]);
    expect(sheets.get('dm').values).toHaveLength(2);
    const row = dmRowFor('1800000000000000001');
    expect(row[5]).toBe('new text');
    expect(row[8]).toBe('2026-01-01T00:00:00.000Z'); // first_seen_at kept
    expect(row[9]).not.toBe('2026-01-01T00:00:00.000Z'); // last_seen_at refreshed
  });
});

// --- mark_dm_read -----------------------------------------------------------

describe('mark_dm_read action', () => {
  it('creates the state row and never moves the marker backwards', () => {
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'mark_dm_read',
      conversation_id: '100-200',
      event_id: '1800000000000000004',
    });
    expect(out).toEqual({
      ok: true,
      conversation_id: '100-200',
      event_id: '1800000000000000004',
    });
    expect(sheets.get('dm_state').values[0]).toEqual(Code.DM_STATE_HEADERS);
    let row = stateRowFor('100-200');
    expect(row[1]).toBe('1800000000000000004');
    expect(row[3]).not.toBe(''); // updated_at

    post({
      idToken: 'good',
      action: 'mark_dm_read',
      conversation_id: '100-200',
      event_id: '1800000000000000002',
    });
    row = stateRowFor('100-200');
    expect(row[1]).toBe('1800000000000000004'); // older id did not regress it
    expect(sheets.get('dm_state').values).toHaveLength(2);

    post({
      idToken: 'good',
      action: 'mark_dm_read',
      conversation_id: '100-200',
      event_id: '1800000000000000009',
    });
    expect(stateRowFor('100-200')[1]).toBe('1800000000000000009');
  });

  it('rejects missing ids without touching the sheets', () => {
    routeFetch();
    const out = post({ idToken: 'good', action: 'mark_dm_read', conversation_id: '100-200' });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/event_id/);
    expect(sheets.get('dm_state')).toBeUndefined();
  });

  it('in dry-run mode reports dry_run and writes nothing', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'DRY_RUN' ? 'true' : (KEYS[name] ?? null)
    );
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'mark_dm_read',
      conversation_id: '100-200',
      event_id: '1800000000000000004',
    });
    expect(out).toEqual({
      ok: true,
      conversation_id: '100-200',
      event_id: '1800000000000000004',
      dry_run: true,
    });
    expect(sheets.get('dm_state')).toBeUndefined();
    expect(xCalls()).toHaveLength(0);
  });

  it('cannot be reached through the sync token', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'SYNC_TOKEN' ? 'sync-secret' : (KEYS[name] ?? null)
    );
    const out = post({
      sync_token: 'sync-secret',
      action: 'mark_dm_read',
      conversation_id: '100-200',
      event_id: '1',
    });
    expect(out).toMatchObject({ ok: false, status: 403 });
    expect(mockFetch).not.toHaveBeenCalled();
  });
});

// --- reply_dm ---------------------------------------------------------------

describe('validateDmText', () => {
  it('counts Unicode code points, not UTF-16 units or weighted length', () => {
    expect(Code.validateDmText('a'.repeat(10000))).toBe('a'.repeat(10000));
    // One emoji is one code point here (two UTF-16 units, weight 2
    // under the tweet rule): 10,000 of them are still acceptable.
    expect(Code.validateDmText('😀'.repeat(10000))).toBe('😀'.repeat(10000));
    expect(() => Code.validateDmText('a'.repeat(10001))).toThrow(/too long/);
    expect(() => Code.validateDmText('😀'.repeat(10001))).toThrow(/too long/);
  });

  it('rejects empty and non-string text', () => {
    expect(() => Code.validateDmText('   ')).toThrow(/empty/);
    expect(() => Code.validateDmText(42)).toThrow(/empty/);
    expect(() => Code.validateDmText(undefined)).toThrow(/empty/);
  });
});

describe('reply_dm action', () => {
  it('posts the payload, advances both state markers, and stores the sent event', () => {
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_dm',
      conversation_id: '100-200',
      text: 'ご連絡ありがとうございます',
    });
    expect(out).toEqual({
      ok: true,
      conversation_id: '100-200',
      event_id: '1800000000000000099',
    });
    const postCall = mockFetch.mock.calls.find(([url]) =>
      String(url).includes('/dm_conversations/')
    );
    expect(String(postCall[0])).toBe(
      'https://api.twitter.com/2/dm_conversations/100-200/messages'
    );
    expect(postCall[1].method).toBe('post');
    expect(postCall[1].headers.Authorization).toMatch(/^OAuth /);
    expect(JSON.parse(postCall[1].payload)).toEqual({
      text: 'ご連絡ありがとうございます',
    });
    const state = stateRowFor('100-200');
    expect(state[1]).toBe('1800000000000000099'); // last_read_event_id
    expect(state[2]).toBe('1800000000000000099'); // last_replied_event_id
    const stored = dmRowFor('1800000000000000099');
    expect(stored[1]).toBe('100-200');
    expect(stored[2]).toBe('111');
    expect(stored[5]).toBe('ご連絡ありがとうございます');
    expect(stored[7]).toBe(true); // is_mine
    // The reply_dm action log carries metadata only, never the text.
    const replyLog = (sheets.get('log')?.values || []).find((r) => r[1] === 'reply_dm');
    expect(replyLog).toBeTruthy();
    expect(replyLog[2]).not.toContain('ご連絡ありがとうございます');
    expect(JSON.parse(replyLog[2]).event_id).toBe('1800000000000000099');
  });

  it('returns event_id null and leaves state alone when the API omits dm_event_id', () => {
    routeFetch({
      dmPost: () => fetchResponse(201, { data: { dm_conversation_id: '100-200' } }),
    });
    const out = post({
      idToken: 'good',
      action: 'reply_dm',
      conversation_id: '100-200',
      text: 'hello',
    });
    expect(out).toEqual({ ok: true, conversation_id: '100-200', event_id: null });
    expect(sheets.get('dm_state')).toBeUndefined();
    expect(sheets.get('dm')).toBeUndefined();
  });

  it('rejects text over 10,000 code points without calling X', () => {
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_dm',
      conversation_id: '100-200',
      text: 'あ'.repeat(10001),
    });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/too long/);
    expect(xCalls()).toHaveLength(0);
    expect(sheets.get('dm_state')).toBeUndefined();
  });

  it('rejects a missing conversation_id without calling X', () => {
    routeFetch();
    const out = post({ idToken: 'good', action: 'reply_dm', text: 'hi' });
    expect(out.ok).toBe(false);
    expect(out.error).toMatch(/conversation_id/);
    expect(xCalls()).toHaveLength(0);
  });

  it('in dry-run mode sends nothing and changes no sheet', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'DRY_RUN' ? '1' : (KEYS[name] ?? null)
    );
    routeFetch();
    const out = post({
      idToken: 'good',
      action: 'reply_dm',
      conversation_id: '100-200',
      text: 'hello',
    });
    expect(out).toEqual({
      ok: true,
      conversation_id: '100-200',
      event_id: null,
      dry_run: true,
    });
    expect(xCalls()).toHaveLength(0);
    expect(sheets.get('dm')).toBeUndefined();
    expect(sheets.get('dm_state')).toBeUndefined();
  });

  it('cannot be reached through the sync token', () => {
    mockGetProperty.mockImplementation((name) =>
      name === 'SYNC_TOKEN' ? 'sync-secret' : (KEYS[name] ?? null)
    );
    const out = post({
      sync_token: 'sync-secret',
      action: 'reply_dm',
      conversation_id: '100-200',
      text: '送信されてはいけない',
    });
    expect(out).toMatchObject({ ok: false, status: 403 });
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('requires user authentication', () => {
    mockFetch.mockReturnValue(fetchResponse(400, {}));
    const out = post({
      idToken: 'bad',
      action: 'reply_dm',
      conversation_id: '100-200',
      text: 'hi',
    });
    expect(out).toEqual({ ok: false, error: 'unauthorized' });
  });
});
