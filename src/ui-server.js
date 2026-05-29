const fs = require('node:fs/promises');
const http = require('node:http');
const path = require('node:path');

const defaultMailClient = require('./microsoft-email.js');

const DEFAULT_HOST = '127.0.0.1';
const DEFAULT_PORT = 8787;
const DEFAULT_PUBLIC_DIR = path.join(__dirname, '..', 'public');
const JSON_LIMIT_BYTES = 1024 * 1024;
const CONTENT_TYPES = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
};
const STATIC_SECURITY_HEADERS = {
  'Content-Security-Policy': [
    "default-src 'self'",
    "script-src 'self'",
    "style-src 'self'",
    "img-src 'self' data:",
    "connect-src 'self'",
    "form-action 'self'",
    "base-uri 'none'",
    "frame-ancestors 'none'",
  ].join('; '),
  'Referrer-Policy': 'no-referrer',
  'X-Content-Type-Options': 'nosniff',
};

function createUiServer(options = {}) {
  const mailClient = options.mailClient || defaultMailClient;
  const publicDir = options.publicDir || DEFAULT_PUBLIC_DIR;

  return http.createServer(async (request, response) => {
    try {
      await handleRequest(request, response, { mailClient, publicDir });
    } catch (error) {
      sendJson(response, error.statusCode || 500, {
        ok: false,
        error: {
          message: error?.message || String(error),
        },
      });
    }
  });
}

async function startUiServer(options = {}) {
  const host = options.host || DEFAULT_HOST;
  const port = Number.isInteger(Number(options.port)) ? Number(options.port) : DEFAULT_PORT;
  const server = createUiServer(options);

  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, host, () => {
      server.off('error', reject);
      resolve();
    });
  });

  const address = server.address();
  return {
    server,
    host,
    port: address?.port || port,
    url: `http://${host}:${address?.port || port}`,
  };
}

async function handleRequest(request, response, context) {
  const url = new URL(request.url || '/', `http://${request.headers.host || '127.0.0.1'}`);

  if (request.method === 'POST' && url.pathname === '/api/code') {
    const body = await parseJsonBody(request);
    const requestOptions = buildCodeRequestOptions(body, context.mailClient);
    const result = await context.mailClient.fetchMicrosoftVerificationCode(requestOptions);
    return sendJson(response, 200, {
      ok: true,
      data: sanitizeCodeResult(result),
    });
  }

  if (request.method === 'POST' && url.pathname === '/api/messages') {
    const body = await parseJsonBody(request);
    const requestOptions = buildMessagesRequestOptions(body, context.mailClient);
    const data = await fetchMailboxList(requestOptions, context.mailClient);
    return sendJson(response, 200, {
      ok: true,
      data,
    });
  }

  if (request.method === 'GET') {
    return serveStatic(request, response, context.publicDir, url.pathname);
  }

  sendJson(response, 404, {
    ok: false,
    error: { message: 'Not found.' },
  });
}

async function serveStatic(_request, response, publicDir, rawPathname) {
  const pathname = rawPathname === '/' ? '/index.html' : rawPathname;
  const requestedPath = path.normalize(decodeURIComponent(pathname)).replace(/^(\.\.[/\\])+/, '');
  const filePath = path.join(publicDir, requestedPath);
  const resolvedPublicDir = path.resolve(publicDir);
  const resolvedFilePath = path.resolve(filePath);

  if (!resolvedFilePath.startsWith(`${resolvedPublicDir}${path.sep}`) && resolvedFilePath !== resolvedPublicDir) {
    return sendJson(response, 403, {
      ok: false,
      error: { message: 'Forbidden.' },
    });
  }

  try {
    const content = await fs.readFile(resolvedFilePath);
    const contentType = CONTENT_TYPES[path.extname(resolvedFilePath)] || 'application/octet-stream';
    response.writeHead(200, {
      'Cache-Control': 'no-store',
      'Content-Type': contentType,
      ...STATIC_SECURITY_HEADERS,
    });
    response.end(content);
  } catch (error) {
    if (error?.code === 'ENOENT') {
      return sendJson(response, 404, {
        ok: false,
        error: { message: 'Not found.' },
      });
    }
    throw error;
  }
}

async function parseJsonBody(request) {
  const chunks = [];
  let totalBytes = 0;

  for await (const chunk of request) {
    totalBytes += chunk.length;
    if (totalBytes > JSON_LIMIT_BYTES) {
      throw httpError(413, 'Request body is too large.');
    }
    chunks.push(chunk);
  }

  const raw = Buffer.concat(chunks).toString('utf8').trim();
  if (!raw) {
    return {};
  }

  try {
    return JSON.parse(raw);
  } catch {
    throw httpError(400, 'Invalid JSON request body.');
  }
}

function buildCodeRequestOptions(body = {}, mailClient = defaultMailClient) {
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

function buildMessagesRequestOptions(body = {}, mailClient = defaultMailClient) {
  return buildBaseMailRequestOptions(body, mailClient);
}

function buildBaseMailRequestOptions(body = {}, mailClient = defaultMailClient) {
  const clientId = String(body.clientId || '').trim();
  const refreshToken = String(body.refreshToken || body.token || '').trim();
  if (!clientId) {
    throw httpError(400, 'Missing Microsoft client_id.');
  }
  if (!refreshToken) {
    throw httpError(400, 'Missing Microsoft refresh token.');
  }

  return {
    clientId,
    refreshToken,
    mailboxes: normalizeMailboxList(body.mailboxes || body.mailbox, mailClient),
    top: parseInteger(body.top, 5, 1, 30),
  };
}

async function fetchMailboxList(options, mailClient = defaultMailClient) {
  const mailboxes = [];
  let workingRefreshToken = options.refreshToken;

  for (const mailbox of options.mailboxes) {
    const result = await mailClient.fetchMicrosoftMailboxMessages({
      clientId: options.clientId,
      refreshToken: workingRefreshToken,
      mailbox,
      top: options.top,
    });
    if (result.nextRefreshToken) {
      workingRefreshToken = result.nextRefreshToken;
    }
    mailboxes.push(sanitizeMailboxResult(result));
  }

  return {
    nextRefreshToken: workingRefreshToken,
    mailboxes,
  };
}

function normalizeMailboxList(value, mailClient = defaultMailClient) {
  const normalizer = typeof mailClient.normalizeMailboxLabel === 'function'
    ? mailClient.normalizeMailboxLabel
    : defaultMailClient.normalizeMailboxLabel;
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
    throw httpError(400, `Expected integer between ${min} and ${max}.`);
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
    throw httpError(400, 'Invalid filter-after timestamp.');
  }
  return parsed;
}

function sanitizeCodeResult(result = {}) {
  return {
    code: result.code || '',
    emailTimestamp: result.emailTimestamp || 0,
    messageId: result.messageId || null,
    sender: result.sender || '',
    subject: result.subject || '',
    mailbox: result.mailbox || '',
    nextRefreshToken: result.nextRefreshToken || '',
    message: result.message || null,
  };
}

function sanitizeMailboxResult(result = {}) {
  return {
    source: 'microsoft-api',
    transport: result.transport || '',
    tokenStrategy: result.tokenStrategy || '',
    mailbox: result.mailbox || '',
    nextRefreshToken: result.nextRefreshToken || '',
    messages: Array.isArray(result.messages) ? result.messages : [],
  };
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    'Cache-Control': 'no-store',
    'Content-Type': 'application/json; charset=utf-8',
  });
  response.end(JSON.stringify(payload));
}

function httpError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

module.exports = {
  DEFAULT_HOST,
  DEFAULT_PORT,
  buildCodeRequestOptions,
  buildMessagesRequestOptions,
  createUiServer,
  normalizeMailboxList,
  parseJsonBody,
  startUiServer,
};
