#!/usr/bin/env node

const microsoftEmail = require('../src/microsoft-email.js');

const DEFAULT_MAILBOXES = ['INBOX'];
const COMMANDS = new Set(['code', 'messages']);

function usage() {
  return [
    'Usage:',
    '  outlook-mail-extractor code [options]',
    '  outlook-mail-extractor messages [options]',
    '',
    'Options:',
    '  --client-id <id>           Microsoft OAuth app client_id. Env: OUTLOOK_CLIENT_ID',
    '  --refresh-token <token>    Microsoft refresh token. Env: OUTLOOK_REFRESH_TOKEN',
    '  --mailbox <name[,name]>    Mailbox list. Supports INBOX and Junk. Repeatable.',
    '  --top <number>             Number of messages per mailbox, 1-30. Default: 5',
    '  --max-retries <number>     Verification-code polling attempts. Default: 3',
    '  --retry-delay-ms <ms>      Delay between polling attempts. Default: 10000',
    '  --filter-after <time>      Unix ms timestamp or parseable date string.',
    '  --sender <text>            Sender/search-text filter. Repeatable.',
    '  --subject <text>           Subject/search-text filter. Repeatable.',
    '  --keyword <text>           Required body/search-text keyword. Repeatable.',
    '  --exclude-code <code>      Code to ignore. Repeatable.',
    '  --code-pattern <regex>     Custom regex. First capture group is used when present.',
    '  --pretty                   Pretty-print JSON output.',
    '  --verbose                  Print fallback attempts to stderr.',
    '  -h, --help                 Show this help.',
  ].join('\n');
}

function parseCliOptions(argv = process.argv.slice(2), env = process.env) {
  const args = [...argv];
  const first = String(args[0] || '').trim();
  const command = first && !first.startsWith('-') ? args.shift() : 'code';
  if (!COMMANDS.has(command)) {
    throw new Error(`Unknown command "${command}". Use "code" or "messages".`);
  }

  const options = {
    command,
    clientId: String(env.OUTLOOK_CLIENT_ID || env.MICROSOFT_CLIENT_ID || '').trim(),
    refreshToken: String(env.OUTLOOK_REFRESH_TOKEN || env.MICROSOFT_REFRESH_TOKEN || '').trim(),
    mailboxes: [],
    top: 5,
    maxRetries: 3,
    retryDelayMs: 10000,
    filterAfterTimestamp: 0,
    senderFilters: [],
    subjectFilters: [],
    requiredKeywords: [],
    excludeCodes: [],
    codePatterns: [],
    pretty: false,
    verbose: false,
    help: false,
  };

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    switch (arg) {
      case '-h':
      case '--help':
        options.help = true;
        break;
      case '--client-id':
        options.clientId = readNextValue(args, ++index, arg);
        break;
      case '--refresh-token':
      case '--token':
        options.refreshToken = readNextValue(args, ++index, arg);
        break;
      case '--mailbox':
      case '--mailboxes':
        options.mailboxes.push(...splitList(readNextValue(args, ++index, arg)));
        break;
      case '--top':
        options.top = parseBoundedInteger(readNextValue(args, ++index, arg), arg, 1, 30);
        break;
      case '--max-retries':
        options.maxRetries = parseBoundedInteger(readNextValue(args, ++index, arg), arg, 1, 50);
        break;
      case '--retry-delay-ms':
        options.retryDelayMs = parseBoundedInteger(readNextValue(args, ++index, arg), arg, 0, 300000);
        break;
      case '--filter-after':
        options.filterAfterTimestamp = parseFilterAfter(readNextValue(args, ++index, arg));
        break;
      case '--sender':
        options.senderFilters.push(...splitList(readNextValue(args, ++index, arg)));
        break;
      case '--subject':
        options.subjectFilters.push(...splitList(readNextValue(args, ++index, arg)));
        break;
      case '--keyword':
        options.requiredKeywords.push(...splitList(readNextValue(args, ++index, arg)));
        break;
      case '--exclude-code':
        options.excludeCodes.push(...splitList(readNextValue(args, ++index, arg)));
        break;
      case '--code-pattern':
        options.codePatterns.push(parseCodePattern(readNextValue(args, ++index, arg)));
        break;
      case '--pretty':
        options.pretty = true;
        break;
      case '--verbose':
        options.verbose = true;
        break;
      default:
        throw new Error(`Unknown option "${arg}".`);
    }
  }

  options.mailboxes = normalizeList(options.mailboxes.length ? options.mailboxes : DEFAULT_MAILBOXES)
    .map(microsoftEmail.normalizeMailboxLabel);
  options.senderFilters = normalizeList(options.senderFilters);
  options.subjectFilters = normalizeList(options.subjectFilters);
  options.requiredKeywords = normalizeList(options.requiredKeywords);
  options.excludeCodes = normalizeList(options.excludeCodes);

  return options;
}

function readNextValue(args, index, optionName) {
  const value = args[index];
  if (!value || String(value).startsWith('--')) {
    throw new Error(`${optionName} requires a value.`);
  }
  return String(value).trim();
}

function splitList(value) {
  return String(value || '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function normalizeList(values) {
  const seen = new Set();
  return (Array.isArray(values) ? values : [])
    .map((value) => String(value || '').trim())
    .filter(Boolean)
    .filter((value) => {
      const key = value.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function parseBoundedInteger(value, optionName, min, max) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < min || parsed > max) {
    throw new Error(`${optionName} must be an integer between ${min} and ${max}.`);
  }
  return parsed;
}

function parseFilterAfter(value) {
  if (/^\d+$/.test(String(value))) {
    return Number(value);
  }
  const timestamp = Date.parse(value);
  if (!Number.isFinite(timestamp)) {
    throw new Error('--filter-after must be a Unix ms timestamp or parseable date string.');
  }
  return timestamp;
}

function parseCodePattern(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    throw new Error('--code-pattern requires a regex source.');
  }
  return { source: raw, flags: 'i' };
}

function buildRequestOptions(options) {
  const base = {
    command: options.command,
    clientId: options.clientId,
    refreshToken: options.refreshToken,
    top: options.top,
  };

  if (options.command === 'messages') {
    return {
      ...base,
      mailbox: options.mailboxes[0] || 'INBOX',
      mailboxes: options.mailboxes,
    };
  }

  return {
    ...base,
    mailboxes: options.mailboxes,
    maxRetries: options.maxRetries,
    retryDelayMs: options.retryDelayMs,
    filterAfterTimestamp: options.filterAfterTimestamp,
    senderFilters: options.senderFilters,
    subjectFilters: options.subjectFilters,
    requiredKeywords: options.requiredKeywords,
    excludeCodes: options.excludeCodes,
    codePatterns: options.codePatterns,
  };
}

async function run(argv = process.argv.slice(2), env = process.env, deps = {}) {
  const stdout = deps.stdout || process.stdout;
  const stderr = deps.stderr || process.stderr;

  try {
    const options = parseCliOptions(argv, env);
    if (options.help) {
      stdout.write(`${usage()}\n`);
      return 0;
    }

    assertRequiredAuth(options);
    const request = buildRequestOptions(options);
    const log = options.verbose ? (message) => stderr.write(`${message}\n`) : null;
    const fetchImpl = deps.fetchImpl || globalThis.fetch;

    // 这里是 CLI 命令路由扩展点：新增命令时只扩展分支，不改核心邮件提取库。
    const payload = options.command === 'messages'
      ? await fetchMessagesCommand(request, { fetchImpl, log })
      : await fetchCodeCommand(request, { fetchImpl, log });

    stdout.write(`${formatJson(payload, options.pretty)}\n`);
    return 0;
  } catch (error) {
    stderr.write(`${error?.message || String(error)}\n\n${usage()}\n`);
    return 1;
  }
}

function assertRequiredAuth(options) {
  if (!options.clientId) {
    throw new Error('Missing Microsoft client_id. Pass --client-id or OUTLOOK_CLIENT_ID.');
  }
  if (!options.refreshToken) {
    throw new Error('Missing Microsoft refresh token. Pass --refresh-token or OUTLOOK_REFRESH_TOKEN.');
  }
}

async function fetchMessagesCommand(request, deps = {}) {
  const results = [];
  let workingRefreshToken = request.refreshToken;

  for (const mailbox of request.mailboxes) {
    const result = await microsoftEmail.fetchMicrosoftMailboxMessages({
      clientId: request.clientId,
      refreshToken: workingRefreshToken,
      mailbox,
      top: request.top,
      fetchImpl: deps.fetchImpl,
      log: deps.log,
    });
    if (result.nextRefreshToken) {
      workingRefreshToken = result.nextRefreshToken;
    }
    results.push(sanitizeMailboxResult(result));
  }

  return {
    command: 'messages',
    nextRefreshToken: workingRefreshToken,
    mailboxes: results,
  };
}

async function fetchCodeCommand(request, deps = {}) {
  const result = await microsoftEmail.fetchMicrosoftVerificationCode({
    clientId: request.clientId,
    refreshToken: request.refreshToken,
    mailboxes: request.mailboxes,
    top: request.top,
    maxRetries: request.maxRetries,
    retryDelayMs: request.retryDelayMs,
    filterAfterTimestamp: request.filterAfterTimestamp,
    senderFilters: request.senderFilters,
    subjectFilters: request.subjectFilters,
    requiredKeywords: request.requiredKeywords,
    excludeCodes: request.excludeCodes,
    codePatterns: request.codePatterns,
    fetchImpl: deps.fetchImpl,
    log: deps.log,
  });

  return {
    command: 'code',
    code: result.code,
    emailTimestamp: result.emailTimestamp,
    messageId: result.messageId,
    sender: result.sender,
    subject: result.subject,
    mailbox: result.mailbox,
    nextRefreshToken: result.nextRefreshToken,
    message: result.message,
  };
}

function sanitizeMailboxResult(result) {
  return {
    source: 'microsoft-api',
    transport: result.transport,
    tokenStrategy: result.tokenStrategy,
    mailbox: result.mailbox,
    nextRefreshToken: result.nextRefreshToken,
    messages: result.messages,
  };
}

function formatJson(payload, pretty = false) {
  return JSON.stringify(payload, null, pretty ? 2 : 0);
}

if (require.main === module) {
  run().then((code) => {
    process.exitCode = code;
  });
}

module.exports = {
  buildRequestOptions,
  formatJson,
  parseCliOptions,
  run,
  splitList,
  usage,
};
