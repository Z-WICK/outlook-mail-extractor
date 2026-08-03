function requestError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function buildCodeRequestOptions(body = {}, mailClient) {
  const base = buildBaseMailRequestOptions(body, mailClient);
  return {
    ...base,
    maxRetries: parseInteger(body.maxRetries, 3, 1, 50),
    retryDelayMs: parseInteger(body.retryDelayMs, 10000, 0, 300000),
    filterAfterTimestamp: parseTimestamp(body.filterAfterTimestamp || body.filterAfter),
    senderFilters: normalizeList(body.senderFilters || body.sender),
    subjectFilters: normalizeList(body.subjectFilters || body.subject),
    requiredKeywords: normalizeList(body.requiredKeywords || body.keyword),
    excludeCodes: normalizeList(body.excludeCodes || body.excludeCode),
    codePatterns: normalizeCodePatterns(body.codePatterns || body.codePattern),
  };
}

function buildMessagesRequestOptions(body = {}, mailClient) {
  return buildBaseMailRequestOptions(body, mailClient);
}

function buildBaseMailRequestOptions(body = {}, mailClient) {
  const clientId = String(body.clientId || '').trim();
  const refreshToken = String(body.refreshToken || body.token || '').trim();
  if (!clientId) {
    throw requestError(400, 'Missing Microsoft client_id.');
  }
  if (!refreshToken) {
    throw requestError(400, 'Missing Microsoft refresh token.');
  }

  return {
    clientId,
    refreshToken,
    mailboxes: normalizeMailboxList(body.mailboxes || body.mailbox, mailClient),
    top: parseInteger(body.top, 5, 1, 30),
  };
}

function normalizeMailboxList(value, mailClient) {
  const fallbackNormalizer = (mailbox) => /^junk(?:\s*e-?mail|\s*email)?$/i.test(String(mailbox || ''))
    ? 'Junk'
    : 'INBOX';
  const normalizer = typeof mailClient?.normalizeMailboxLabel === 'function'
    ? mailClient.normalizeMailboxLabel
    : fallbackNormalizer;
  const rawList = normalizeList(value);
  const list = rawList.length ? rawList : ['INBOX'];
  return [...new Set(list.map((mailbox) => normalizer(mailbox)))];
}

function normalizeList(value) {
  const values = Array.isArray(value)
    ? value
    : String(value || '').split(',');
  const seen = new Set();
  return values
    .flatMap((item) => String(item || '').split(','))
    .map((item) => item.trim())
    .filter(Boolean)
    .filter((item) => {
      const key = item.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function normalizeCodePatterns(value) {
  if (Array.isArray(value)) {
    return value
      .map((item) => {
        if (typeof item === 'string') return { source: item, flags: 'i' };
        return item;
      })
      .filter((item) => String(item?.source || '').trim());
  }
  return normalizeList(value).map((source) => ({ source, flags: 'i' }));
}

function parseInteger(value, fallback, min, max) {
  if (value === undefined || value === null || value === '') {
    return fallback;
  }
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < min || parsed > max) {
    throw requestError(400, `Expected integer between ${min} and ${max}.`);
  }
  return parsed;
}

function parseTimestamp(value) {
  if (value === undefined || value === null || value === '') {
    return 0;
  }
  if (/^\d+$/.test(String(value))) {
    return Number(value);
  }
  const parsed = Date.parse(value);
  if (!Number.isFinite(parsed)) {
    throw requestError(400, 'Invalid filter-after timestamp.');
  }
  return parsed;
}

module.exports = {
  buildCodeRequestOptions,
  buildMessagesRequestOptions,
  normalizeCodePatterns,
  normalizeList,
  normalizeMailboxList,
  parseInteger,
  parseTimestamp,
  requestError,
};
