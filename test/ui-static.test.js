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

test('UI omits auxiliary console and local request labels', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');

  assert.doesNotMatch(html, /本地邮件提取控制台/);
  assert.doesNotMatch(html, /请求配置/);
  assert.doesNotMatch(html, /127\.0\.0\.1/);
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

test('browser account store updates returned refresh token locally', () => {
  const {
    createMemoryStorage,
    getStoredAccount,
    updateStoredAccountRefreshToken,
    upsertStoredAccounts,
  } = require('../public/app.js');
  const storage = createMemoryStorage();

  upsertStoredAccounts([{
    email: 'Token.Owner@Outlook.com',
    password: 'pass-1',
    clientId: 'client-1',
    refreshToken: 'old-refresh-token',
  }], storage);

  const updated = updateStoredAccountRefreshToken('token.owner@outlook.com', 'new-refresh-token', storage);

  assert.equal(updated.email, 'token.owner@outlook.com');
  assert.equal(updated.refreshToken, undefined);
  assert.equal(updated.refreshTokenMasked, 'new-****oken');
  assert.equal(getStoredAccount('token.owner@outlook.com', storage).refreshToken, 'new-refresh-token');
});

test('UI script applies returned refresh token to selected browser account after successful requests', () => {
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');

  assert.match(script, /function applyNextRefreshToken\(nextRefreshToken\)/);
  assert.match(script, /updateStoredAccountRefreshToken\(selectedAccountId, nextRefreshToken\)/);
  assert.match(script, /applyNextRefreshToken\(payload\.data\?\.nextRefreshToken\)/);
});

test('UI script clears selected browser credential state after deleting current account', () => {
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');

  assert.match(script, /function clearSelectedCredentialState\(\)/);
  assert.match(script, /pendingCredential = null/);
  assert.match(script, /elements\.clientId\.value = ''/);
  assert.match(script, /elements\.refreshToken\.value = ''/);
  assert.match(script, /if \(selectedAccountId === id\) \{\s*clearSelectedCredentialState\(\);\s*\}/s);
});

test('UI script automatically fetches messages after selecting a browser account', () => {
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');

  assert.match(script, /async function autoFetchSelectedAccountMessages\(\)/);
  assert.match(script, /await autoFetchSelectedAccountMessages\(\)/);
  assert.match(script, /await runMessageFetch\(\)/);
});

test('mail list pagination shows one newest message per page by default', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');
  const styles = fs.readFileSync(path.join(publicDir, 'styles.css'), 'utf8');
  const { getMessagePageState, sortMessagesByNewest } = require('../public/app.js');

  assert.match(html, /id="messagePagination"[^>]*data-testid="message-pagination"[^>]*hidden/);
  assert.match(html, /id="messagePrevPage"/);
  assert.match(html, /id="messagePageInfo"/);
  assert.match(html, /id="messageNextPage"/);
  assert.match(script, /const MESSAGE_PAGE_SIZE = 1/);
  assert.match(script, /let currentMessages = \[\]/);
  assert.match(script, /let currentMessagePage = 1/);
  assert.match(script, /function renderMessagePage\(\)/);
  assert.match(script, /currentMessages = sortMessagesByNewest\(messages\)/);
  assert.match(script, /currentMessagePage = 1/);
  assert.match(script, /visibleMessages\.map\(renderMessageCard\)/);
  assert.match(script, /elements\.messagePrevPage\.disabled = !pageState\.hasPreviousPage/);
  assert.match(script, /elements\.messageNextPage\.disabled = !pageState\.hasNextPage/);
  assert.match(styles, /\.message-pagination/);

  const messages = [
    { id: 'old', receivedDateTime: '2026-05-29T08:00:00.000Z' },
    { id: 'newest', receivedDateTime: '2026-05-29T10:00:00.000Z' },
    { id: 'middle', receivedDateTime: '2026-05-29T09:00:00.000Z' },
  ];
  const sortedMessages = sortMessagesByNewest(messages);
  assert.deepEqual(sortedMessages.map((message) => message.id), ['newest', 'middle', 'old']);

  const firstPage = getMessagePageState(sortedMessages, 1, 1);
  assert.deepEqual(firstPage.visibleMessages.map((message) => message.id), ['newest']);
  assert.equal(firstPage.currentPage, 1);
  assert.equal(firstPage.totalPages, 3);
  assert.equal(firstPage.hasPreviousPage, false);
  assert.equal(firstPage.hasNextPage, true);

  const secondPage = getMessagePageState(sortedMessages, 2, 1);
  assert.deepEqual(secondPage.visibleMessages.map((message) => message.id), ['middle']);
  assert.equal(secondPage.currentPage, 2);
  assert.equal(secondPage.totalPages, 3);
  assert.equal(secondPage.hasPreviousPage, true);
  assert.equal(secondPage.hasNextPage, true);

  const clampedPage = getMessagePageState(sortedMessages, 99, 1);
  assert.deepEqual(clampedPage.visibleMessages.map((message) => message.id), ['old']);
  assert.equal(clampedPage.currentPage, 3);
  assert.equal(clampedPage.hasNextPage, false);
});

test('UI moves browser account pool into a dense modal grid', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');
  const styles = fs.readFileSync(path.join(publicDir, 'styles.css'), 'utf8');

  assert.match(html, /id="accountPoolToggle"[^>]*data-testid="account-pool-toggle"/);
  assert.match(html, /id="accountPoolDialog"[^>]*data-testid="account-pool-dialog"[^>]*role="dialog"[^>]*hidden/);
  assert.match(html, /id="accountPoolPanel"/);
  assert.match(html, /id="accountPoolClose"[^>]*aria-label="关闭邮箱池"/);
  assert.match(html, /id="accountPoolDialog"[\s\S]*id="accountSearch"[\s\S]*id="accountList"[\s\S]*<\/div>\s*<\/div>\s*<input id="clientId"/);
  assert.doesNotMatch(html, /<section class="account-pool"/);
  assert.match(script, /accountPoolToggle: document\.getElementById\('accountPoolToggle'\)/);
  assert.match(script, /function openAccountPoolDialog\(\)/);
  assert.match(script, /function closeAccountPoolDialog\(\)/);
  assert.match(script, /elements\.accountPoolToggle\.addEventListener\('click', \(\) => openAccountPoolDialog\(\)\)/);
  assert.match(script, /event\.key === 'Escape' && !elements\.accountPoolDialog\.hidden/);
  assert.match(styles, /\.account-pool-overlay/);
  assert.match(styles, /\.account-pool-panel/);
  assert.match(styles, /\.account-list\s*\{[^}]*grid-template-columns:\s*repeat\(3,\s*minmax\(0,\s*1fr\)\)/s);
  assert.match(styles, /@media \(max-width: 720px\)[\s\S]*\.account-list\s*\{[^}]*grid-template-columns:\s*1fr/s);
});

test('UI shows the selected account in the result header instead of the import summary', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');
  const styles = fs.readFileSync(path.join(publicDir, 'styles.css'), 'utf8');

  assert.match(html, /class="result-title"[\s\S]*class="result-title-row"[\s\S]*id="resultMeta"[\s\S]*id="activeAccountMeta"/);
  assert.match(html, /class="result-state"[^>]*id="resultMeta"/);
  assert.match(html, /data-testid="active-account-meta"/);
  assert.match(script, /activeAccountMeta: document\.getElementById\('activeAccountMeta'\)/);
  assert.match(script, /function setActiveAccountMeta\(email\)/);
  assert.match(script, /setActiveAccountMeta\(primary\.email\)/);
  assert.match(script, /setActiveAccountMeta\(account\.email\)/);
  assert.match(script, /setActiveAccountMeta\(''\)/);
  assert.match(script, /正在使用 \$\{value\}/);
  assert.doesNotMatch(html, /parsedAccountSummary|parsed-account-summary|class="parsed-summary"/);
  assert.doesNotMatch(script, /parsedAccountSummary/);
  assert.doesNotMatch(styles, /\.parsed-summary/);
  assert.match(styles, /\.active-account-meta/);
  assert.match(styles, /\.result-title-row/);
  assert.match(styles, /\.result-state/);
  assert.match(styles, /\.result-state::before/);
});

test('UI puts the settings trigger beside status and keeps mailbox selection inside settings dialog', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');
  const styles = fs.readFileSync(path.join(publicDir, 'styles.css'), 'utf8');

  assert.match(html, /data-testid="settings-toggle"/);
  assert.match(html, /id="settingsDialog"[^>]*data-testid="settings-dialog"/);
  assert.match(html, /id="settingsPanel"/);
  assert.match(html, /aria-label="打开高级配置"/);
  assert.match(html, /aria-label="关闭高级配置"/);
  assert.match(html, /class="settings-overlay"[^>]*hidden/);
  assert.match(html, /class="topbar-actions"[\s\S]*id="statusText"[\s\S]*id="settingsToggle"[\s\S]*<\/div>\s*<\/header>/);
  assert.match(html, /id="settingsDialog"[\s\S]*<fieldset class="mailboxes"[\s\S]*name="mailbox"[\s\S]*id="senderFilters"[\s\S]*id="sessionMemory"[\s\S]*<\/div>\s*<\/div>\s*<div class="actions">/);
  assert.doesNotMatch(html, /class="mailbox-row"/);
  assert.ok(html.indexOf('id="settingsToggle"') < html.indexOf('<main class="workspace">'));
  assert.match(script, /settingsToggle: document\.getElementById\('settingsToggle'\)/);
  assert.match(script, /function openSettingsDialog\(\)/);
  assert.match(script, /function closeSettingsDialog\(\)/);
  assert.match(styles, /\.topbar-actions/);
  assert.match(styles, /\.settings-overlay/);
  assert.match(styles, /\.settings-panel/);
});

test('UI uses a polished Outlook inspired visual system', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');
  const styles = fs.readFileSync(path.join(publicDir, 'styles.css'), 'utf8');

  assert.match(html, /class="brand-mark"[\s\S]*<svg viewBox="0 0 24 24" role="img"/);
  assert.match(styles, /--primary:\s*#0078d4/);
  assert.match(styles, /--primary-strong:\s*#005a9e/);
  assert.match(styles, /--bg:\s*#f3f6fb/);
  assert.match(styles, /--surface:\s*#ffffff/);
  assert.match(styles, /--outlook-rail:\s*#eff6fc/);
  assert.match(styles, /\.topbar\s*\{[^}]*background:\s*rgba\(255,\s*255,\s*255,\s*0\.86\)/s);
  assert.match(styles, /\.brand-mark\s*\{[^}]*background:\s*linear-gradient\(135deg,\s*#0078d4,\s*#0f6cbd\)/s);
  assert.match(styles, /\.panel\s*\{[^}]*border:\s*1px solid var\(--border\)[^}]*box-shadow:\s*var\(--shadow\)/s);
  assert.match(styles, /\.import-box,\s*\.account-pool-entry\s*\{[^}]*background:\s*linear-gradient\(180deg,\s*#ffffff,\s*var\(--outlook-rail\)\)/s);
  assert.match(styles, /textarea\s*\{[^}]*background:\s*#ffffff[^}]*box-shadow:\s*inset 0 1px 2px rgba\(0,\s*0,\s*0,\s*0\.04\)/s);
  assert.match(styles, /\.result-header\s*\{[^}]*background:\s*linear-gradient\(180deg,\s*#ffffff,\s*#f7fbff\)/s);
  assert.match(styles, /\.code-line\s*\{[^}]*background:\s*linear-gradient\(180deg,\s*#eef7ff,\s*#ffffff\)/s);
  assert.match(styles, /\.message-card\s*\{[^}]*border-left:\s*3px solid var\(--primary\)/s);
  assert.match(styles, /\.message-card h3\s*\{[^}]*color:\s*#102a43/s);
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

test('code result omits raw message json debug output', () => {
  const html = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');
  const script = fs.readFileSync(path.join(publicDir, 'app.js'), 'utf8');
  const styles = fs.readFileSync(path.join(publicDir, 'styles.css'), 'utf8');

  assert.doesNotMatch(html, /id="codeJson"/);
  assert.doesNotMatch(html, /class="json-view"/);
  assert.doesNotMatch(script, /codeJson/);
  assert.doesNotMatch(script, /JSON\.stringify\(data\.message/);
  assert.doesNotMatch(styles, /\.json-view/);
});
