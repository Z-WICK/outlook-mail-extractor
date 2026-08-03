import microsoftEmail from './microsoft-email.js';
import requestOptions from './request-options.js';

const {
  buildCodeRequestOptions,
  buildMessagesRequestOptions,
  requestError,
} = requestOptions;

const JSON_LIMIT_BYTES = 1024 * 1024;
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
const INLINE_ASSET_DEFINITIONS = {
  '/': {
    binding: 'INLINE_INDEX_HTML',
    contentType: 'text/html; charset=utf-8',
  },
  '/index.html': {
    binding: 'INLINE_INDEX_HTML',
    contentType: 'text/html; charset=utf-8',
  },
  '/app.js': {
    binding: 'INLINE_APP_JS',
    contentType: 'application/javascript; charset=utf-8',
  },
  '/styles.css': {
    binding: 'INLINE_STYLES_CSS',
    contentType: 'text/css; charset=utf-8',
  },
};

export function createWorker({ mailClient = microsoftEmail } = {}) {
  return {
    async fetch(request, env) {
      try {
        return await handleRequest(request, env, mailClient);
      } catch (error) {
        return jsonResponse(error.statusCode || 500, {
          ok: false,
          error: { message: error?.message || String(error) },
        });
      }
    },
  };
}

async function handleRequest(request, env, mailClient) {
  const url = new URL(request.url);

  if (request.method === 'POST' && url.pathname === '/api/code') {
    const body = await readJsonBody(request);
    const requestOptions = buildCodeRequestOptions(body, mailClient);
    const result = await mailClient.fetchMicrosoftVerificationCode(requestOptions);
    return jsonResponse(200, {
      ok: true,
      data: sanitizeCodeResult(result),
    });
  }

  if (request.method === 'POST' && url.pathname === '/api/messages') {
    const body = await readJsonBody(request);
    const requestOptions = buildMessagesRequestOptions(body, mailClient);
    const data = await fetchMailboxList(requestOptions, mailClient);
    return jsonResponse(200, {
      ok: true,
      data,
    });
  }

  if (url.pathname.startsWith('/api/')) {
    return jsonResponse(404, {
      ok: false,
      error: { message: 'Not found.' },
    });
  }

  return serveStatic(request, env);
}

async function serveStatic(request, env) {
  let assetResponse;
  if (env?.ASSETS && typeof env.ASSETS.fetch === 'function') {
    assetResponse = await env.ASSETS.fetch(request);
  } else {
    assetResponse = serveInlineAsset(request, env);
  }

  const headers = new Headers(assetResponse.headers);
  headers.set('Cache-Control', 'no-store');
  for (const [name, value] of Object.entries(STATIC_SECURITY_HEADERS)) {
    headers.set(name, value);
  }

  return new Response(assetResponse.body, {
    status: assetResponse.status,
    statusText: assetResponse.statusText,
    headers,
  });
}

function serveInlineAsset(request, env) {
  const pathname = new URL(request.url).pathname;
  const definition = INLINE_ASSET_DEFINITIONS[pathname];
  const content = definition ? env?.[definition.binding] : null;
  if (typeof content !== 'string') {
    return new Response('Not found.', { status: 404 });
  }

  return new Response(content, {
    status: 200,
    headers: { 'Content-Type': definition.contentType },
  });
}

async function readJsonBody(request) {
  const declaredLength = Number(request.headers.get('content-length'));
  if (Number.isFinite(declaredLength) && declaredLength > JSON_LIMIT_BYTES) {
    throw requestError(413, 'Request body is too large.');
  }

  if (!request.body) {
    return {};
  }

  const reader = request.body.getReader();
  const decoder = new TextDecoder();
  let raw = '';
  let totalBytes = 0;

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      totalBytes += value.byteLength;
      if (totalBytes > JSON_LIMIT_BYTES) {
        await reader.cancel();
        throw requestError(413, 'Request body is too large.');
      }
      raw += decoder.decode(value, { stream: true });
    }
    raw += decoder.decode();
  } finally {
    reader.releaseLock();
  }

  const trimmed = raw.trim();
  if (!trimmed) {
    return {};
  }

  try {
    return JSON.parse(trimmed);
  } catch {
    throw requestError(400, 'Invalid JSON request body.');
  }
}

async function fetchMailboxList(options, mailClient) {
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

function jsonResponse(statusCode, payload) {
  return new Response(JSON.stringify(payload), {
    status: statusCode,
    headers: {
      'Cache-Control': 'no-store',
      'Content-Type': 'application/json; charset=utf-8',
    },
  });
}

const worker = createWorker();

export default worker;
