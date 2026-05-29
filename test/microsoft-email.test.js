const assert = require('node:assert/strict');
const test = require('node:test');

const microsoftEmail = require('../src/microsoft-email.js');

test('normalizes Graph and Outlook messages into one mailbox shape', () => {
  const result = microsoftEmail.normalizeMessage({
    Id: 'msg-1',
    Subject: 'Your code',
    From: { EmailAddress: { Address: 'security@example.com', Name: 'Security' } },
    ToRecipients: [
      { EmailAddress: { Address: 'User@example.com' } },
      { EmailAddress: { Address: 'user@example.com' } },
    ],
    CcRecipients: [{ EmailAddress: { Address: 'audit@example.com' } }],
    BodyPreview: 'Use 123456',
    ReceivedDateTime: '2026-05-28T01:02:03.000Z',
  }, 'junk email');

  assert.equal(result.mailbox, 'Junk');
  assert.equal(result.id, 'msg-1');
  assert.equal(result.from.emailAddress.address, 'security@example.com');
  assert.deepEqual(result.recipients.to, ['User@example.com']);
  assert.deepEqual(result.recipients.all, ['User@example.com', 'audit@example.com']);
});

test('extracts the newest verification code that matches filters', () => {
  const result = microsoftEmail.extractVerificationCodeFromMessages([
    {
      id: 'old',
      mailbox: 'INBOX',
      subject: 'OpenAI code',
      from: { emailAddress: { address: 'noreply@example.com' } },
      bodyPreview: 'Your code is 111111',
      receivedDateTime: '2026-05-28T01:00:00.000Z',
    },
    {
      id: 'match',
      mailbox: 'Junk',
      subject: 'OpenAI login code',
      from: { emailAddress: { address: 'noreply@tm.openai.com' } },
      bodyPreview: 'Enter this code 654321 to continue',
      receivedDateTime: '2026-05-28T01:05:00.000Z',
    },
  ], {
    filterAfterTimestamp: Date.parse('2026-05-28T01:01:00.000Z'),
    senderFilters: ['openai.com'],
    subjectFilters: ['login'],
    excludeCodes: ['111111'],
  });

  assert.equal(result.code, '654321');
  assert.equal(result.messageId, 'match');
  assert.equal(result.mailbox, 'Junk');
});

test('fetches Graph message body so OpenAI codes beyond bodyPreview can be extracted', async () => {
  const calls = [];
  const result = await microsoftEmail.fetchMicrosoftVerificationCode({
    clientId: 'client-id',
    refreshToken: 'refresh-1',
    maxRetries: 1,
    retryDelayMs: 0,
    senderFilters: ['openai.com'],
    subjectFilters: ['login'],
    fetchImpl: async (url) => {
      const requestedUrl = String(url);
      calls.push(requestedUrl);
      if (requestedUrl.includes('/oauth2/v2.0/token')) {
        return jsonResponse({ access_token: 'access-1', refresh_token: 'refresh-2' });
      }

      const requestedBody = /\$select=[^&]*\bbody\b/i.test(requestedUrl);
      return jsonResponse({
        value: [{
          id: 'openai-login',
          subject: 'Your temporary OpenAI login code',
          from: { emailAddress: { address: 'noreply@tm.openai.com' } },
          bodyPreview: 'Enter this temporary verification code to continue:',
          body: requestedBody
            ? { content: 'Enter this temporary verification code to continue:\n\n\n924881\n\nBest,\nThe OpenAI team' }
            : undefined,
          receivedDateTime: '2026-05-29T04:16:50.000Z',
        }],
      });
    },
  });

  assert.equal(result.code, '924881');
  assert.equal(result.messageId, 'openai-login');
  assert.equal(result.nextRefreshToken, 'refresh-2');
  assert.equal(calls.some((url) => /\$select=[^&]*\bbody\b/i.test(url)), true);
});

test('falls back across Microsoft token and mailbox transports', async () => {
  const calls = [];
  const fetchImpl = async (url) => {
    calls.push(String(url));
    if (String(url).includes('/oauth2/v2.0/token') && calls.length === 1) {
      return jsonResponse({ error: { message: 'delegated denied' } }, false, 400);
    }
    if (String(url).includes('/oauth2/v2.0/token')) {
      return jsonResponse({ access_token: 'access-2', refresh_token: 'refresh-2' });
    }
    return jsonResponse({
      value: [
        {
          id: 'graph-msg',
          subject: 'Login code',
          from: { emailAddress: { address: 'sender@example.com' } },
          bodyPreview: 'code is 222333',
          receivedDateTime: '2026-05-28T02:00:00.000Z',
        },
      ],
    });
  };

  const result = await microsoftEmail.fetchMicrosoftMailboxMessages({
    clientId: 'client-id',
    refreshToken: 'refresh-1',
    mailbox: 'INBOX',
    fetchImpl,
  });

  assert.equal(result.transport, 'graph');
  assert.equal(result.tokenStrategy, 'entra-consumers-delegated');
  assert.equal(result.nextRefreshToken, 'refresh-2');
  assert.equal(result.messages[0].id, 'graph-msg');
});

function jsonResponse(payload, ok = true, status = 200) {
  return {
    ok,
    status,
    statusText: ok ? 'OK' : 'Bad Request',
    json: async () => payload,
    text: async () => JSON.stringify(payload),
  };
}
