const assert = require('node:assert/strict');
const test = require('node:test');

const { createUiServer } = require('../src/ui-server.js');

test('serves the UI entry page from GET /', async (t) => {
  const fixture = await startFixtureServer(t);
  const response = await fetch(`${fixture.origin}/`);
  const text = await response.text();

  assert.equal(response.status, 200);
  assert.match(response.headers.get('content-type'), /text\/html/);
  assert.match(text, /Outlook Mail Extractor/);
});

test('serves UI with browser storage security headers', async (t) => {
  const fixture = await startFixtureServer(t);
  const response = await fetch(`${fixture.origin}/`);

  assert.match(response.headers.get('content-security-policy'), /default-src 'self'/);
  assert.equal(response.headers.get('x-content-type-options'), 'nosniff');
  assert.equal(response.headers.get('referrer-policy'), 'no-referrer');
});

test('POST /api/code forwards normalized options and returns extracted code', async (t) => {
  const calls = [];
  const fixture = await startFixtureServer(t, {
    fetchMicrosoftVerificationCode: async (options) => {
      calls.push(options);
      return {
        code: '123456',
        emailTimestamp: Date.parse('2026-05-28T02:03:04.000Z'),
        messageId: 'msg-1',
        sender: 'noreply@example.com',
        subject: 'Login code',
        mailbox: 'Junk',
        nextRefreshToken: 'refresh-2',
        message: { id: 'msg-1', subject: 'Login code' },
      };
    },
  });

  const response = await postJson(`${fixture.origin}/api/code`, {
    clientId: 'client-1',
    refreshToken: 'refresh-1',
    mailboxes: ['INBOX', 'junk email'],
    senderFilters: ['openai.com'],
    subjectFilters: ['login'],
    requiredKeywords: ['continue'],
    excludeCodes: ['000000'],
    top: 10,
    maxRetries: 2,
    retryDelayMs: 25,
  });
  const payload = await response.json();

  assert.equal(response.status, 200);
  assert.equal(payload.ok, true);
  assert.equal(payload.data.code, '123456');
  assert.equal(payload.data.nextRefreshToken, 'refresh-2');
  assert.deepEqual(calls[0].mailboxes, ['INBOX', 'Junk']);
  assert.deepEqual(calls[0].senderFilters, ['openai.com']);
  assert.equal(calls[0].top, 10);
  assert.equal(calls[0].maxRetries, 2);
});

test('POST /api/messages fetches each mailbox and strips tokenData from response', async (t) => {
  const calls = [];
  const fixture = await startFixtureServer(t, {
    fetchMicrosoftMailboxMessages: async (options) => {
      calls.push(options);
      return {
        tokenData: { access_token: 'do-not-return' },
        nextRefreshToken: `${options.mailbox}-refresh`,
        transport: 'graph',
        tokenStrategy: 'entra-common-delegated',
        mailbox: options.mailbox,
        messages: [{ id: `${options.mailbox}-1`, subject: 'Hello' }],
      };
    },
  });

  const response = await postJson(`${fixture.origin}/api/messages`, {
    clientId: 'client-1',
    refreshToken: 'refresh-1',
    mailboxes: ['INBOX', 'Junk'],
    top: 3,
  });
  const payload = await response.json();

  assert.equal(response.status, 200);
  assert.equal(payload.ok, true);
  assert.deepEqual(calls.map((call) => call.mailbox), ['INBOX', 'Junk']);
  assert.equal(calls[1].refreshToken, 'INBOX-refresh');
  assert.equal(payload.data.mailboxes.length, 2);
  assert.equal(payload.data.mailboxes[0].tokenData, undefined);
  assert.equal(payload.data.mailboxes[0].messages[0].id, 'INBOX-1');
});

test('rejects malformed JSON with a 400 response', async (t) => {
  const fixture = await startFixtureServer(t);
  const response = await fetch(`${fixture.origin}/api/code`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: '{',
  });
  const payload = await response.json();

  assert.equal(response.status, 400);
  assert.equal(payload.ok, false);
  assert.match(payload.error.message, /Invalid JSON/i);
});

test('account API is not exposed because accounts are browser-local', async (t) => {
  const fixture = await startFixtureServer(t);

  const listResponse = await fetch(`${fixture.origin}/api/accounts`);
  const createResponse = await postJson(`${fixture.origin}/api/accounts`, {
    email: 'local.only@example.com',
    password: 'pass-123',
    clientId: 'client-1',
    refreshToken: 'refresh-1',
  });
  const detailResponse = await fetch(`${fixture.origin}/api/accounts/local.only%40example.com`);
  const deleteResponse = await fetch(`${fixture.origin}/api/accounts/local.only%40example.com`, {
    method: 'DELETE',
  });

  assert.equal(listResponse.status, 404);
  assert.equal(createResponse.status, 404);
  assert.equal(detailResponse.status, 404);
  assert.equal(deleteResponse.status, 404);
});

async function startFixtureServer(t, mailClient = {}, options = {}) {
  const server = createUiServer({
    accountsFile: options.accountsFile,
    mailClient: {
      normalizeMailboxLabel: (value) => /^junk/i.test(String(value || '')) ? 'Junk' : 'INBOX',
      fetchMicrosoftMailboxMessages: async () => ({ messages: [] }),
      fetchMicrosoftVerificationCode: async () => ({ code: '000000' }),
      ...mailClient,
    },
  });

  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise((resolve) => server.close(resolve)));

  const { port } = server.address();
  return { origin: `http://127.0.0.1:${port}` };
}

async function postJson(url, body) {
  return fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
}
