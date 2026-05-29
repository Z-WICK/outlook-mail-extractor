const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const publicDir = path.join(__dirname, '..', 'public');

test('UI entry wires app assets and key controls', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');

  assert.match(html, /<link rel="stylesheet" href="\/styles\.css">/);
  assert.match(html, /<script src="\/app\.js" defer><\/script>/);
  assert.match(html, /data-testid="client-id"[^>]*type="hidden"/);
  assert.match(html, /data-testid="refresh-token"[^>]*type="hidden"/);
  assert.doesNotMatch(html, />Client ID</);
  assert.doesNotMatch(html, />Refresh Token</);
  assert.match(html, /data-testid="credential-raw"/);
  assert.match(html, /data-testid="parse-credential"/);
  assert.doesNotMatch(html, /data-testid="save-account"/);
  assert.match(html, /解析并保存/);
  assert.match(html, /data-testid="account-search"/);
  assert.match(html, /data-testid="account-list"/);
  assert.match(html, /data-testid="extract-code"/);
  assert.match(html, /data-testid="fetch-messages"/);
  assert.match(html, /data-testid="session-memory"/);
  assert.match(html, /data-testid="result-tabs"/);
});

test('UI script uses browser-local accounts and only talks to mail JSON API endpoints', () => {
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');

  assert.match(script, /\/api\/code/);
  assert.match(script, /\/api\/messages/);
  assert.doesNotMatch(script, /\/api\/accounts/);
  assert.match(script, /localStorage/);
  assert.match(script, /ACCOUNT_STORAGE_KEY/);
  assert.match(script, /parseCredentialRows/);
  assert.match(script, /await saveParsedCredentials\(parsedAccounts\)/);
  assert.doesNotMatch(script, /elements\.credentialRaw\.value = `\$\{account\.email\}/);
  assert.match(script, /sessionStorage/);
});

test('browser account store persists, searches, loads, and deletes locally', () => {
  const {
    createMemoryStorage,
    getStoredAccount,
    removeStoredAccount,
    searchStoredAccounts,
    upsertStoredAccounts,
  } = require('../public/app.js');
  const storage = createMemoryStorage();

  const saved = upsertStoredAccounts([
    {
      email: 'Local.One@Outlook.com',
      password: 'pass-1',
      clientId: 'client-1',
      refreshToken: 'refresh-1',
    },
    {
      email: 'local.two@outlook.com',
      password: 'pass-2',
      clientId: 'client-2',
      refreshToken: 'refresh-2',
    },
  ], storage);

  assert.equal(saved.length, 2);
  assert.equal(saved[0].email, 'local.one@outlook.com');
  assert.equal(saved[0].refreshToken, undefined);
  assert.equal(saved[0].refreshTokenMasked, 'refr****sh-1');

  const found = searchStoredAccounts('client-2', storage);
  assert.equal(found.length, 1);
  assert.equal(found[0].email, 'local.two@outlook.com');
  assert.equal(found[0].refreshToken, undefined);

  const detail = getStoredAccount('local.two@outlook.com', storage);
  assert.equal(detail.refreshToken, 'refresh-2');

  assert.equal(removeStoredAccount('local.two@outlook.com', storage), true);
  assert.equal(searchStoredAccounts('', storage).length, 1);
});

test('credential parser accepts pasted email password client token rows', () => {
  const { maskSecret, parseCredentialLine } = require('../public/app.js');
  const parsed = parseCredentialLine([
    '邮箱----密码----client_id----令牌',
    'demo@example.com----pass-123----00000000-0000-0000-0000-000000000000----M.C525_BAY.0.U.-token*with----tail',
  ].join('\n'));

  assert.deepEqual(parsed, {
    email: 'demo@example.com',
    password: 'pass-123',
    clientId: '00000000-0000-0000-0000-000000000000',
    refreshToken: 'M.C525_BAY.0.U.-token*with----tail',
  });
  assert.equal(maskSecret(parsed.password), 'pa****23');
  assert.equal(maskSecret(parsed.refreshToken).startsWith('M.'), true);
});

test('credential parser accepts batch pasted rows and skips headers', () => {
  const { parseCredentialRows } = require('../public/app.js');
  const parsed = parseCredentialRows([
    '邮箱----密码----client_id----令牌',
    'first@example.com----pass-1----client-1----token-1',
    '',
    'second@example.com----pass-2----client-2----token----with-tail',
  ].join('\n'));

  assert.deepEqual(parsed, [
    {
      email: 'first@example.com',
      password: 'pass-1',
      clientId: 'client-1',
      refreshToken: 'token-1',
    },
    {
      email: 'second@example.com',
      password: 'pass-2',
      clientId: 'client-2',
      refreshToken: 'token----with-tail',
    },
  ]);
});

test('mail list styles hide empty states and wrap long message content', () => {
  const styles = fs.readFileSync(path.join(publicDir, 'styles.css'), 'utf8');
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');

  assert.match(styles, /\[hidden\]\s*\{[^}]*display:\s*none\s*!important/s);
  assert.match(styles, /\.result-panel\s*\{[^}]*min-width:\s*0/s);
  assert.match(styles, /\.message-card\s*\{[^}]*overflow-wrap:\s*anywhere/s);
  assert.match(styles, /\.message-id\s*\{[^}]*overflow-wrap:\s*anywhere/s);
  assert.match(styles, /\.message-preview\s*\{[^}]*overflow-wrap:\s*anywhere/s);
  assert.match(script, /className = 'message-id'/);
});
