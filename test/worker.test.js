const assert = require('node:assert/strict');
const test = require('node:test');

let createWorker;

test.before(async () => {
  ({ createWorker } = await import('../src/worker.mjs'));
});

test('Worker serves static assets with security headers', async () => {
  const worker = createWorker();
  const response = await worker.fetch(new Request('https://extractor.example/'), {
    ASSETS: {
      fetch: async () => new Response('<html>ok</html>', {
        headers: { 'Content-Type': 'text/html; charset=utf-8' },
      }),
    },
  });

  assert.equal(response.status, 200);
  assert.match(await response.text(), /<html>ok/);
  assert.match(response.headers.get('content-security-policy'), /default-src 'self'/);
  assert.equal(response.headers.get('x-content-type-options'), 'nosniff');
  assert.equal(response.headers.get('cache-control'), 'no-store');
});

test('Worker can serve inline assets when deployed without an Assets binding', async () => {
  const worker = createWorker();
  const response = await worker.fetch(new Request('https://extractor.example/'), {
    INLINE_INDEX_HTML: '<html>inline</html>',
  });

  assert.equal(response.status, 200);
  assert.equal(response.headers.get('content-type'), 'text/html; charset=utf-8');
  assert.match(await response.text(), /inline/);
});

test('Worker code endpoint preserves existing response sanitization', async () => {
  const calls = [];
  const worker = createWorker({
    mailClient: {
      normalizeMailboxLabel: (value) => /^junk/i.test(String(value || '')) ? 'Junk' : 'INBOX',
      fetchMicrosoftVerificationCode: async (options) => {
        calls.push(options);
        return {
          code: '123456',
          nextRefreshToken: 'refresh-2',
          tokenData: { access_token: 'must-not-return' },
          message: { id: 'message-1' },
        };
      },
    },
  });

  const response = await worker.fetch(new Request('https://extractor.example/api/code', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      clientId: 'client-1',
      refreshToken: 'refresh-1',
      mailboxes: ['INBOX', 'junk email'],
      top: 10,
      maxRetries: 2,
    }),
  }), {});
  const payload = await response.json();

  assert.equal(response.status, 200);
  assert.equal(payload.ok, true);
  assert.equal(payload.data.code, '123456');
  assert.equal(payload.data.nextRefreshToken, 'refresh-2');
  assert.equal(payload.data.tokenData, undefined);
  assert.deepEqual(calls[0].mailboxes, ['INBOX', 'Junk']);
  assert.equal(calls[0].maxRetries, 2);
});

test('Worker messages endpoint updates refresh token between mailboxes', async () => {
  const calls = [];
  const worker = createWorker({
    mailClient: {
      normalizeMailboxLabel: (value) => /^junk/i.test(String(value || '')) ? 'Junk' : 'INBOX',
      fetchMicrosoftMailboxMessages: async (options) => {
        calls.push(options);
        return {
          tokenData: { access_token: 'must-not-return' },
          nextRefreshToken: `${options.mailbox}-refresh`,
          transport: 'graph',
          tokenStrategy: 'entra-common-delegated',
          mailbox: options.mailbox,
          messages: [{ id: `${options.mailbox}-1` }],
        };
      },
    },
  });

  const response = await worker.fetch(new Request('https://extractor.example/api/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      clientId: 'client-1',
      refreshToken: 'refresh-1',
      mailboxes: ['INBOX', 'Junk'],
    }),
  }), {});
  const payload = await response.json();

  assert.equal(response.status, 200);
  assert.deepEqual(calls.map((call) => call.mailbox), ['INBOX', 'Junk']);
  assert.equal(calls[1].refreshToken, 'INBOX-refresh');
  assert.equal(payload.data.nextRefreshToken, 'Junk-refresh');
  assert.equal(payload.data.mailboxes[0].tokenData, undefined);
});

test('Worker rejects malformed and oversized JSON requests', async () => {
  const worker = createWorker();
  const malformed = await worker.fetch(new Request('https://extractor.example/api/code', {
    method: 'POST',
    body: '{',
  }), {});
  const oversized = await worker.fetch(new Request('https://extractor.example/api/code', {
    method: 'POST',
    body: JSON.stringify({ value: 'x'.repeat(1024 * 1024) }),
  }), {});

  assert.equal(malformed.status, 400);
  assert.match((await malformed.json()).error.message, /Invalid JSON/i);
  assert.equal(oversized.status, 413);
  assert.match((await oversized.json()).error.message, /too large/i);
});

test('Worker keeps unknown API routes separate from static assets', async () => {
  const worker = createWorker();
  const response = await worker.fetch(new Request('https://extractor.example/api/accounts'), {
    ASSETS: {
      fetch: async () => new Response('asset should not be used'),
    },
  });
  const payload = await response.json();

  assert.equal(response.status, 404);
  assert.equal(payload.ok, false);
});
