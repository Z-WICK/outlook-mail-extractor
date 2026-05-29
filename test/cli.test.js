const assert = require('node:assert/strict');
const test = require('node:test');

const cli = require('../bin/outlook-mail-extractor.js');

test('parses CLI options with environment fallbacks', () => {
  const options = cli.parseCliOptions([
    'code',
    '--mailbox',
    'INBOX,Junk',
    '--sender',
    'openai.com',
    '--subject',
    'login',
    '--keyword',
    'continue',
    '--exclude-code',
    '111111',
    '--top',
    '10',
    '--max-retries',
    '2',
    '--pretty',
  ], {
    OUTLOOK_CLIENT_ID: 'env-client',
    OUTLOOK_REFRESH_TOKEN: 'env-refresh',
  });

  assert.equal(options.command, 'code');
  assert.equal(options.clientId, 'env-client');
  assert.equal(options.refreshToken, 'env-refresh');
  assert.deepEqual(options.mailboxes, ['INBOX', 'Junk']);
  assert.deepEqual(options.senderFilters, ['openai.com']);
  assert.deepEqual(options.subjectFilters, ['login']);
  assert.deepEqual(options.requiredKeywords, ['continue']);
  assert.deepEqual(options.excludeCodes, ['111111']);
  assert.equal(options.top, 10);
  assert.equal(options.maxRetries, 2);
  assert.equal(options.pretty, true);
});

test('builds mailbox request options for the messages command', () => {
  const options = cli.parseCliOptions([
    'messages',
    '--client-id',
    'client-1',
    '--refresh-token',
    'refresh-1',
    '--mailbox',
    'Junk',
  ], {});

  const request = cli.buildRequestOptions(options);

  assert.equal(request.command, 'messages');
  assert.equal(request.mailbox, 'Junk');
  assert.equal(request.clientId, 'client-1');
  assert.equal(request.refreshToken, 'refresh-1');
});
