const SESSION_KEY = 'outlook-mail-extractor-ui';
const ACCOUNT_STORAGE_KEY = 'outlook-mail-extractor-accounts-v1';
const MESSAGE_PAGE_SIZE = 1;
let pendingCredential = null;
let selectedAccountId = '';
let currentMessages = [];
let currentMessagePage = 1;

const elements = typeof document === 'undefined' ? null : {
  form: document.getElementById('extractorForm'),
  credentialRaw: document.getElementById('credentialRaw'),
  parseCredentialButton: document.getElementById('parseCredentialButton'),
  accountPoolToggle: document.getElementById('accountPoolToggle'),
  accountPoolDialog: document.getElementById('accountPoolDialog'),
  accountPoolPanel: document.getElementById('accountPoolPanel'),
  accountPoolClose: document.getElementById('accountPoolClose'),
  accountSearch: document.getElementById('accountSearch'),
  accountList: document.getElementById('accountList'),
  accountCount: document.getElementById('accountCount'),
  settingsToggle: document.getElementById('settingsToggle'),
  settingsDialog: document.getElementById('settingsDialog'),
  settingsPanel: document.getElementById('settingsPanel'),
  settingsClose: document.getElementById('settingsClose'),
  clientId: document.getElementById('clientId'),
  refreshToken: document.getElementById('refreshToken'),
  senderFilters: document.getElementById('senderFilters'),
  subjectFilters: document.getElementById('subjectFilters'),
  requiredKeywords: document.getElementById('requiredKeywords'),
  excludeCodes: document.getElementById('excludeCodes'),
  top: document.getElementById('top'),
  maxRetries: document.getElementById('maxRetries'),
  retryDelayMs: document.getElementById('retryDelayMs'),
  filterAfter: document.getElementById('filterAfter'),
  sessionMemory: document.getElementById('sessionMemory'),
  extractCodeButton: document.getElementById('extractCodeButton'),
  fetchMessagesButton: document.getElementById('fetchMessagesButton'),
  statusText: document.getElementById('statusText'),
  errorBox: document.getElementById('errorBox'),
  resultMeta: document.getElementById('resultMeta'),
  activeAccountMeta: document.getElementById('activeAccountMeta'),
  codeTab: document.getElementById('codeTab'),
  messagesTab: document.getElementById('messagesTab'),
  codePanel: document.getElementById('codePanel'),
  messagesPanel: document.getElementById('messagesPanel'),
  codeEmpty: document.getElementById('codeEmpty'),
  codeResult: document.getElementById('codeResult'),
  codeValue: document.getElementById('codeValue'),
  copyCodeButton: document.getElementById('copyCodeButton'),
  codeMeta: document.getElementById('codeMeta'),
  codeJson: document.getElementById('codeJson'),
  messagesEmpty: document.getElementById('messagesEmpty'),
  messageList: document.getElementById('messageList'),
  messagePagination: document.getElementById('messagePagination'),
  messagePrevPage: document.getElementById('messagePrevPage'),
  messagePageInfo: document.getElementById('messagePageInfo'),
  messageNextPage: document.getElementById('messageNextPage'),
};

if (elements) {
  init();
}

function init() {
  restoreSessionConfig();
  loadAccounts();
  elements.parseCredentialButton.addEventListener('click', () => applyCredentialImport());
  elements.accountPoolToggle.addEventListener('click', () => openAccountPoolDialog());
  elements.accountPoolClose.addEventListener('click', () => closeAccountPoolDialog());
  elements.accountPoolDialog.addEventListener('click', (event) => {
    if (event.target === elements.accountPoolDialog) {
      closeAccountPoolDialog();
    }
  });
  elements.accountSearch.addEventListener('input', () => loadAccounts(elements.accountSearch.value));
  elements.credentialRaw.addEventListener('keydown', (event) => {
    if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
      event.preventDefault();
      applyCredentialImport();
    }
  });
  elements.extractCodeButton.addEventListener('click', () => runCodeExtraction());
  elements.fetchMessagesButton.addEventListener('click', () => runMessageFetch());
  elements.settingsToggle.addEventListener('click', () => openSettingsDialog());
  elements.settingsClose.addEventListener('click', () => closeSettingsDialog());
  elements.settingsDialog.addEventListener('click', (event) => {
    if (event.target === elements.settingsDialog) {
      closeSettingsDialog();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !elements.accountPoolDialog.hidden) {
      closeAccountPoolDialog();
      return;
    }
    if (event.key === 'Escape' && !elements.settingsDialog.hidden) {
      closeSettingsDialog();
    }
  });
  elements.copyCodeButton.addEventListener('click', () => copyCurrentCode());
  elements.codeTab.addEventListener('click', () => activateTab('code'));
  elements.messagesTab.addEventListener('click', () => activateTab('messages'));
  elements.messagePrevPage.addEventListener('click', () => setMessagePage(currentMessagePage - 1));
  elements.messageNextPage.addEventListener('click', () => setMessagePage(currentMessagePage + 1));
  elements.sessionMemory.addEventListener('change', () => {
    if (elements.sessionMemory.checked) {
      persistSessionConfig();
    } else {
      sessionStorage.removeItem(SESSION_KEY);
    }
  });
  elements.form.addEventListener('input', () => {
    if (elements.sessionMemory.checked) {
      persistSessionConfig();
    }
  });
}

function parseCredentialLine(value) {
  return parseCredentialRows(value)[0];
}

function parseCredentialRows(value) {
  const lines = selectCredentialDataLines(value);
  return lines.map(({ line, lineNumber }) => parseCredentialDataLine(line, lineNumber));
}

function parseCredentialDataLine(line, lineNumber) {
  const parts = line.split('----');
  const prefix = lineNumber ? `第 ${lineNumber} 行` : '账号行';
  if (parts.length < 4) {
    throw new Error(`${prefix}格式不正确，请使用：邮箱----密码----client_id----令牌`);
  }

  const [email, password, clientId, ...tokenParts] = parts;
  const result = {
    email: String(email || '').trim(),
    password: String(password || '').trim(),
    clientId: String(clientId || '').trim(),
    refreshToken: tokenParts.join('----').trim(),
  };

  if (!result.email || !result.password || !result.clientId || !result.refreshToken) {
    throw new Error(`${prefix}解析失败：邮箱、密码、client_id、令牌都不能为空。`);
  }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(result.email)) {
    throw new Error(`${prefix}解析失败：邮箱格式不正确。`);
  }

  return result;
}

function selectCredentialDataLine(value) {
  return selectCredentialDataLines(value)[0].line;
}

function selectCredentialDataLines(value) {
  const lines = String(value || '')
    .split(/\r?\n/)
    .map((line, index) => ({ line: line.trim(), lineNumber: index + 1 }))
    .filter(({ line }) => line && line.includes('----') && !isCredentialHeaderLine(line));
  if (!lines.length) {
    throw new Error('没有找到可解析的账号行。');
  }
  return lines;
}

function isCredentialHeaderLine(line) {
  const normalized = String(line || '').toLowerCase();
  return normalized.includes('邮箱----密码') || normalized.includes('email----password');
}

function maskSecret(value) {
  const text = String(value || '');
  if (text.length <= 4) {
    return '*'.repeat(text.length);
  }
  if (text.length <= 8) {
    return `${text.slice(0, 2)}****${text.slice(-2)}`;
  }
  return `${text.slice(0, 4)}****${text.slice(-4)}`;
}

function createMemoryStorage(initial = {}) {
  const values = new Map(Object.entries(initial));
  return {
    getItem: (key) => values.has(key) ? values.get(key) : null,
    setItem: (key, value) => values.set(key, String(value)),
    removeItem: (key) => values.delete(key),
    clear: () => values.clear(),
  };
}

function getBrowserAccountStorage() {
  return typeof localStorage === 'undefined' ? null : localStorage;
}

function readStoredAccounts(storage = getBrowserAccountStorage()) {
  if (!storage) return [];

  try {
    const parsed = JSON.parse(storage.getItem(ACCOUNT_STORAGE_KEY) || '{}');
    const accounts = Array.isArray(parsed?.accounts) ? parsed.accounts : [];
    return accounts.map(normalizeStoredAccount).filter(Boolean);
  } catch {
    return [];
  }
}

function writeStoredAccounts(accounts, storage = getBrowserAccountStorage()) {
  if (!storage) {
    throw new Error('当前浏览器不支持本地邮箱池。');
  }
  const normalizedAccounts = accounts.map(normalizeStoredAccount).filter(Boolean);
  storage.setItem(ACCOUNT_STORAGE_KEY, JSON.stringify({
    version: 1,
    accounts: normalizedAccounts,
  }));
  return normalizedAccounts;
}

function upsertStoredAccounts(inputs, storage = getBrowserAccountStorage()) {
  const now = new Date().toISOString();
  const existingAccounts = readStoredAccounts(storage);
  const accountMap = new Map(existingAccounts.map((account) => [account.id, account]));
  const savedAccounts = inputs.map((input) => {
    const nextAccount = normalizeStoredAccount({ ...input, updatedAt: now });
    if (!nextAccount) {
      throw new Error('邮箱账号格式不正确。');
    }
    const existing = accountMap.get(nextAccount.id);
    const merged = {
      ...existing,
      ...nextAccount,
      createdAt: existing?.createdAt || now,
      updatedAt: now,
    };
    accountMap.set(merged.id, merged);
    return merged;
  });

  writeStoredAccounts(Array.from(accountMap.values()), storage);
  return savedAccounts.map(toStoredAccountSummary);
}

function searchStoredAccounts(query = '', storage = getBrowserAccountStorage()) {
  const normalizedQuery = String(query || '').trim().toLowerCase();
  return readStoredAccounts(storage)
    .filter((account) => {
      if (!normalizedQuery) return true;
      return [account.email, account.clientId, account.id]
        .some((value) => String(value || '').toLowerCase().includes(normalizedQuery));
    })
    .sort((left, right) => String(right.updatedAt || '').localeCompare(String(left.updatedAt || '')))
    .map(toStoredAccountSummary);
}

function getStoredAccount(id, storage = getBrowserAccountStorage()) {
  const accountId = normalizeAccountId(id);
  const account = readStoredAccounts(storage).find((item) => item.id === accountId);
  return account ? { ...account } : null;
}

function removeStoredAccount(id, storage = getBrowserAccountStorage()) {
  const accountId = normalizeAccountId(id);
  const accounts = readStoredAccounts(storage);
  const nextAccounts = accounts.filter((account) => account.id !== accountId);
  writeStoredAccounts(nextAccounts, storage);
  return accounts.length !== nextAccounts.length;
}

function updateStoredAccountRefreshToken(id, nextRefreshToken, storage = getBrowserAccountStorage()) {
  const accountId = normalizeAccountId(id);
  const refreshToken = String(nextRefreshToken || '').trim();
  if (!accountId || !refreshToken) {
    return null;
  }

  const now = new Date().toISOString();
  let updatedAccount = null;
  const nextAccounts = readStoredAccounts(storage).map((account) => {
    if (account.id !== accountId) {
      return account;
    }
    updatedAccount = {
      ...account,
      refreshToken,
      updatedAt: now,
    };
    return updatedAccount;
  });
  if (!updatedAccount) {
    return null;
  }

  writeStoredAccounts(nextAccounts, storage);
  return toStoredAccountSummary(updatedAccount);
}

function normalizeStoredAccount(account = {}) {
  const email = normalizeEmail(account.email || account.id);
  if (!email) return null;
  return {
    id: normalizeAccountId(email),
    email,
    password: String(account.password || ''),
    clientId: String(account.clientId || ''),
    refreshToken: String(account.refreshToken || account.token || ''),
    createdAt: String(account.createdAt || ''),
    updatedAt: String(account.updatedAt || ''),
  };
}

function toStoredAccountSummary(account = {}) {
  return {
    id: account.id,
    email: account.email,
    clientId: account.clientId,
    passwordMasked: maskSecret(account.password),
    refreshTokenMasked: maskSecret(account.refreshToken),
    createdAt: account.createdAt || '',
    updatedAt: account.updatedAt || '',
  };
}

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeAccountId(value) {
  return normalizeEmail(value);
}

async function applyCredentialImport() {
  elements.parseCredentialButton.disabled = true;
  try {
    const parsedAccounts = parseCredentialRows(elements.credentialRaw.value);
    const primary = parsedAccounts[0];
    pendingCredential = primary;
    elements.clientId.value = primary.clientId;
    elements.refreshToken.value = primary.refreshToken;
    clearError();
    setStatus(`保存中 0/${parsedAccounts.length}`, 'busy');
    if (elements.sessionMemory.checked) {
      persistSessionConfig();
    }
    await saveParsedCredentials(parsedAccounts);
  } catch (error) {
    showError(error?.message || String(error));
    setStatus('导入失败', 'error');
  } finally {
    elements.parseCredentialButton.disabled = false;
  }
}

async function saveParsedCredentials(parsedAccounts) {
  setStatus(`保存中 ${parsedAccounts.length}/${parsedAccounts.length}`, 'busy');
  const savedAccounts = upsertStoredAccounts(parsedAccounts);

  const primary = savedAccounts[0];
  selectedAccountId = primary.id;
  setActiveAccountMeta(primary.email);
  setStatus(savedAccounts.length === 1 ? '已保存' : `已保存 ${savedAccounts.length} 个`, 'done');
  clearError();
  await loadAccounts(elements.accountSearch.value);
}

async function loadAccounts(query = '') {
  try {
    renderAccounts(searchStoredAccounts(query));
  } catch (error) {
    renderAccounts([], error?.message || String(error));
  }
}

function renderAccounts(accounts, errorMessage = '') {
  elements.accountList.replaceChildren();
  elements.accountCount.textContent = `${accounts.length} 个邮箱`;

  if (errorMessage) {
    const error = document.createElement('div');
    error.className = 'account-empty';
    error.textContent = errorMessage;
    elements.accountList.append(error);
    return;
  }

  if (!accounts.length) {
    const empty = document.createElement('div');
    empty.className = 'account-empty';
    empty.textContent = '暂无已保存邮箱';
    elements.accountList.append(empty);
    return;
  }

  elements.accountList.append(...accounts.map(renderAccountItem));
}

function renderAccountItem(account) {
  const item = document.createElement('div');
  item.className = account.id === selectedAccountId ? 'account-item active' : 'account-item';

  const body = document.createElement('button');
  body.className = 'account-select';
  body.type = 'button';
  body.setAttribute('aria-label', `切换邮箱 ${account.email}`);
  body.addEventListener('click', () => selectAccount(account.id));

  const email = document.createElement('strong');
  email.textContent = account.email;
  const meta = document.createElement('span');
  meta.textContent = `client ${account.clientId || '-'} · token ${account.refreshTokenMasked || '-'}`;
  body.append(email, meta);

  const remove = document.createElement('button');
  remove.className = 'account-delete';
  remove.type = 'button';
  remove.textContent = '删除';
  remove.addEventListener('click', () => deleteAccount(account.id));

  item.append(body, remove);
  return item;
}

async function selectAccount(id) {
  try {
    const account = getStoredAccount(id);
    if (!account) {
      throw new Error('账号不存在。');
    }
    selectedAccountId = account.id;
    pendingCredential = {
      email: account.email,
      password: account.password,
      clientId: account.clientId,
      refreshToken: account.refreshToken,
    };
    elements.clientId.value = account.clientId || '';
    elements.refreshToken.value = account.refreshToken || '';
    setActiveAccountMeta(account.email);
    setStatus('已切换', 'done');
    clearError();
    if (elements.sessionMemory.checked) {
      persistSessionConfig();
    }
    await loadAccounts(elements.accountSearch.value);
    closeAccountPoolDialog();
    await autoFetchSelectedAccountMessages();
  } catch (error) {
    showError(error?.message || String(error));
    setStatus('切换失败', 'error');
  }
}

async function deleteAccount(id) {
  try {
    if (!removeStoredAccount(id)) {
      throw new Error('账号不存在。');
    }
    if (selectedAccountId === id) {
      clearSelectedCredentialState();
    }
    setStatus('已删除', 'done');
    clearError();
    await loadAccounts(elements.accountSearch.value);
  } catch (error) {
    showError(error?.message || String(error));
    setStatus('删除失败', 'error');
  }
}

async function runCodeExtraction() {
  await runRequest({
    endpoint: '/api/code',
    busyText: '提取中',
    doneText: '已提取',
    onSuccess: renderCodeResult,
  });
}

async function runMessageFetch() {
  await runRequest({
    endpoint: '/api/messages',
    busyText: '拉取中',
    doneText: '已拉取',
    onSuccess: renderMessageResult,
  });
}

async function autoFetchSelectedAccountMessages() {
  if (elements.fetchMessagesButton.disabled) {
    return;
  }
  await runMessageFetch();
}

async function runRequest({ endpoint, busyText, doneText, onSuccess }) {
  setBusy(true, busyText);
  clearError();

  try {
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(collectFormPayload()),
    });
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      throw new Error(payload?.error?.message || `HTTP ${response.status}`);
    }
    await applyNextRefreshToken(payload.data?.nextRefreshToken);
    onSuccess(payload.data);
    setStatus(doneText, 'done');
    if (elements.sessionMemory.checked) {
      persistSessionConfig();
    }
  } catch (error) {
    showError(error?.message || String(error));
    setStatus('失败', 'error');
  } finally {
    setBusy(false);
  }
}

function collectFormPayload() {
  return {
    clientId: elements.clientId.value.trim(),
    refreshToken: elements.refreshToken.value.trim(),
    mailboxes: selectedMailboxes(),
    senderFilters: splitList(elements.senderFilters.value),
    subjectFilters: splitList(elements.subjectFilters.value),
    requiredKeywords: splitList(elements.requiredKeywords.value),
    excludeCodes: splitList(elements.excludeCodes.value),
    top: numberValue(elements.top.value),
    maxRetries: numberValue(elements.maxRetries.value),
    retryDelayMs: numberValue(elements.retryDelayMs.value),
    filterAfter: elements.filterAfter.value,
  };
}

async function applyNextRefreshToken(nextRefreshToken) {
  const refreshToken = String(nextRefreshToken || '').trim();
  if (!refreshToken) {
    return;
  }

  elements.refreshToken.value = refreshToken;
  if (pendingCredential) {
    pendingCredential = {
      ...pendingCredential,
      refreshToken,
    };
  }

  if (selectedAccountId) {
    updateStoredAccountRefreshToken(selectedAccountId, nextRefreshToken);
    await loadAccounts(elements.accountSearch.value);
  }
}

function clearSelectedCredentialState() {
  selectedAccountId = '';
  pendingCredential = null;
  elements.clientId.value = '';
  elements.refreshToken.value = '';
  setActiveAccountMeta('');
  if (elements.sessionMemory.checked) {
    persistSessionConfig();
  } else {
    sessionStorage.removeItem(SESSION_KEY);
  }
}

function selectedMailboxes() {
  return Array.from(document.querySelectorAll('input[name="mailbox"]:checked'))
    .map((input) => input.value);
}

function splitList(value) {
  return String(value || '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function numberValue(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function renderCodeResult(data) {
  activateTab('code');
  elements.codeEmpty.hidden = true;
  elements.codeResult.hidden = false;
  elements.codeValue.textContent = data.code || '------';
  elements.codeMeta.replaceChildren(
    metaItem('邮箱夹', data.mailbox || '-'),
    metaItem('发件人', data.sender || '-'),
    metaItem('主题', data.subject || '-'),
    metaItem('时间', formatTime(data.emailTimestamp)),
    metaItem('消息 ID', data.messageId || '-'),
    metaItem('Refresh Token', data.nextRefreshToken ? '已返回新 token' : '未返回')
  );
  elements.codeJson.textContent = JSON.stringify(data.message || {}, null, 2);
  elements.resultMeta.textContent = data.code ? `验证码 ${data.code}` : '未找到验证码';
}

function renderMessageResult(data) {
  activateTab('messages');
  const mailboxes = Array.isArray(data?.mailboxes) ? data.mailboxes : [];
  const messages = mailboxes.flatMap((mailbox) => {
    const mailboxName = mailbox.mailbox || 'INBOX';
    return (Array.isArray(mailbox.messages) ? mailbox.messages : [])
      .map((message) => ({ ...message, mailbox: message.mailbox || mailboxName }));
  });

  currentMessages = sortMessagesByNewest(messages);
  currentMessagePage = 1;
  renderMessagePage();
  elements.resultMeta.textContent = `${messages.length} 封邮件`;
}

function renderMessagePage() {
  const pageState = getMessagePageState(currentMessages, currentMessagePage, MESSAGE_PAGE_SIZE);
  currentMessagePage = pageState.currentPage;
  elements.messageList.replaceChildren(...pageState.visibleMessages.map(renderMessageCard));
  elements.messagesEmpty.hidden = pageState.totalMessages > 0;
  elements.messagePagination.hidden = pageState.totalPages <= 1;
  elements.messagePageInfo.textContent = `第 ${pageState.currentPage} / ${pageState.totalPages} 页`;
  elements.messagePrevPage.disabled = !pageState.hasPreviousPage;
  elements.messageNextPage.disabled = !pageState.hasNextPage;
}

function setMessagePage(page) {
  currentMessagePage = page;
  renderMessagePage();
}

function getMessagePageState(messages, page = 1, pageSize = MESSAGE_PAGE_SIZE) {
  const allMessages = Array.isArray(messages) ? messages : [];
  const safePageSize = Math.max(1, Number(pageSize) || MESSAGE_PAGE_SIZE);
  const totalMessages = allMessages.length;
  const totalPages = Math.max(1, Math.ceil(totalMessages / safePageSize));
  const currentPage = Math.min(Math.max(1, Number(page) || 1), totalPages);
  const start = (currentPage - 1) * safePageSize;
  const visibleMessages = allMessages.slice(start, start + safePageSize);

  return {
    currentPage,
    hasNextPage: currentPage < totalPages,
    hasPreviousPage: currentPage > 1,
    totalMessages,
    totalPages,
    visibleMessages,
  };
}

function sortMessagesByNewest(messages) {
  return (Array.isArray(messages) ? messages : [])
    .map((message, index) => ({ message, index }))
    .sort((left, right) => {
      const leftTime = getMessageTimestamp(left.message);
      const rightTime = getMessageTimestamp(right.message);
      if (rightTime !== leftTime) {
        return rightTime - leftTime;
      }
      return left.index - right.index;
    })
    .map(({ message }) => message);
}

function getMessageTimestamp(message = {}) {
  const timestamp = Date.parse(message.receivedDateTime || message.createdDateTime || '');
  return Number.isNaN(timestamp) ? 0 : timestamp;
}

function renderMessageCard(message) {
  const card = document.createElement('article');
  card.className = 'message-card';

  const title = document.createElement('h3');
  title.textContent = message.subject || '(无主题)';

  const meta = document.createElement('div');
  meta.className = 'message-meta';
  [
    `邮箱夹 ${message.mailbox || '-'}`,
    `发件人 ${message.from?.emailAddress?.address || '-'}`,
    `时间 ${formatTime(message.receivedDateTime)}`,
  ].forEach((value) => {
    const span = document.createElement('span');
    span.textContent = value;
    meta.append(span);
  });

  const messageId = document.createElement('div');
  messageId.className = 'message-id';
  messageId.textContent = `ID ${message.id || '-'}`;

  const preview = document.createElement('div');
  preview.className = 'message-preview';
  preview.textContent = message.bodyPreview || message.body?.content || '';

  card.append(title, meta, messageId, preview);
  return card;
}

function metaItem(label, value) {
  const wrapper = document.createElement('div');
  const term = document.createElement('dt');
  const detail = document.createElement('dd');
  term.textContent = label;
  detail.textContent = value;
  wrapper.append(term, detail);
  return wrapper;
}

function activateTab(name) {
  const isCode = name === 'code';
  elements.codeTab.classList.toggle('active', isCode);
  elements.messagesTab.classList.toggle('active', !isCode);
  elements.codeTab.setAttribute('aria-selected', String(isCode));
  elements.messagesTab.setAttribute('aria-selected', String(!isCode));
  elements.codePanel.hidden = !isCode;
  elements.messagesPanel.hidden = isCode;
}

async function copyCurrentCode() {
  const code = elements.codeValue.textContent.trim();
  if (!code || code === '------') return;
  await navigator.clipboard.writeText(code);
  setStatus('已复制', 'done');
}

function setBusy(isBusy, text = '') {
  elements.extractCodeButton.disabled = isBusy;
  elements.fetchMessagesButton.disabled = isBusy;
  if (isBusy) {
    setStatus(text, 'busy');
  }
}

function setStatus(text, state = '') {
  elements.statusText.textContent = text;
  elements.statusText.className = `status-pill ${state}`.trim();
}

function setActiveAccountMeta(email) {
  const value = String(email || '').trim();
  elements.activeAccountMeta.textContent = value ? `正在使用 ${value}` : '';
  elements.activeAccountMeta.hidden = !value;
}

function showError(message) {
  elements.errorBox.textContent = message;
  elements.errorBox.hidden = false;
}

function clearError() {
  elements.errorBox.textContent = '';
  elements.errorBox.hidden = true;
}

function openSettingsDialog() {
  elements.settingsDialog.hidden = false;
  elements.settingsToggle.setAttribute('aria-expanded', 'true');
  elements.settingsClose.focus();
}

function closeSettingsDialog() {
  elements.settingsDialog.hidden = true;
  elements.settingsToggle.setAttribute('aria-expanded', 'false');
  elements.settingsToggle.focus();
}

function openAccountPoolDialog() {
  elements.accountPoolDialog.hidden = false;
  elements.accountPoolToggle.setAttribute('aria-expanded', 'true');
  elements.accountSearch.focus();
}

function closeAccountPoolDialog() {
  elements.accountPoolDialog.hidden = true;
  elements.accountPoolToggle.setAttribute('aria-expanded', 'false');
  elements.accountPoolToggle.focus();
}

function persistSessionConfig() {
  sessionStorage.setItem(SESSION_KEY, JSON.stringify({
    ...collectFormPayload(),
    sessionMemory: true,
  }));
}

function restoreSessionConfig() {
  const raw = sessionStorage.getItem(SESSION_KEY);
  if (!raw) return;

  try {
    const data = JSON.parse(raw);
    elements.sessionMemory.checked = Boolean(data.sessionMemory);
    elements.clientId.value = data.clientId || '';
    elements.refreshToken.value = data.refreshToken || '';
    elements.senderFilters.value = (data.senderFilters || []).join(', ');
    elements.subjectFilters.value = (data.subjectFilters || []).join(', ');
    elements.requiredKeywords.value = (data.requiredKeywords || []).join(', ');
    elements.excludeCodes.value = (data.excludeCodes || []).join(', ');
    elements.top.value = data.top || 10;
    elements.maxRetries.value = data.maxRetries || 3;
    elements.retryDelayMs.value = data.retryDelayMs || 10000;
    elements.filterAfter.value = data.filterAfter || '';
    document.querySelectorAll('input[name="mailbox"]').forEach((input) => {
      input.checked = (data.mailboxes || ['INBOX', 'Junk']).includes(input.value);
    });
  } catch {
    sessionStorage.removeItem(SESSION_KEY);
  }
}

function formatTime(value) {
  if (!value) return '-';
  const date = typeof value === 'number' ? new Date(value) : new Date(String(value));
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString();
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    ACCOUNT_STORAGE_KEY,
    createMemoryStorage,
    getMessagePageState,
    getStoredAccount,
    maskSecret,
    parseCredentialLine,
    parseCredentialRows,
    removeStoredAccount,
    searchStoredAccounts,
    selectCredentialDataLine,
    selectCredentialDataLines,
    sortMessagesByNewest,
    updateStoredAccountRefreshToken,
    upsertStoredAccounts,
  };
}
