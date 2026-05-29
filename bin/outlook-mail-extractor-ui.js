#!/usr/bin/env node

const { DEFAULT_HOST, DEFAULT_PORT, startUiServer } = require('../src/ui-server.js');

function usage() {
  return [
    'Usage:',
    '  outlook-mail-extractor-ui [options]',
    '',
    'Options:',
    `  --host <host>      Host to bind. Default: ${DEFAULT_HOST}`,
    `  --port <port>      Port to bind. Default: ${DEFAULT_PORT}`,
    '  -h, --help         Show this help.',
  ].join('\n');
}

function parseArgs(argv = process.argv.slice(2), env = process.env) {
  const options = {
    host: env.OUTLOOK_EXTRACTOR_UI_HOST || DEFAULT_HOST,
    port: parsePort(env.OUTLOOK_EXTRACTOR_UI_PORT || DEFAULT_PORT),
    help: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    switch (arg) {
      case '-h':
      case '--help':
        options.help = true;
        break;
      case '--host':
        options.host = readNextValue(argv, ++index, arg);
        break;
      case '--port':
        options.port = parsePort(readNextValue(argv, ++index, arg));
        break;
      default:
        throw new Error(`Unknown option "${arg}".`);
    }
  }

  return options;
}

async function run(argv = process.argv.slice(2), env = process.env, deps = {}) {
  const stdout = deps.stdout || process.stdout;
  const stderr = deps.stderr || process.stderr;

  try {
    const options = parseArgs(argv, env);
    if (options.help) {
      stdout.write(`${usage()}\n`);
      return 0;
    }

    const runtime = await startUiServer({
      host: options.host,
      port: options.port,
    });
    stdout.write(`Outlook Mail Extractor UI listening at ${runtime.url}\n`);
    return runtime;
  } catch (error) {
    stderr.write(`${error?.message || String(error)}\n\n${usage()}\n`);
    return 1;
  }
}

function readNextValue(args, index, optionName) {
  const value = args[index];
  if (!value || String(value).startsWith('--')) {
    throw new Error(`${optionName} requires a value.`);
  }
  return String(value).trim();
}

function parsePort(value) {
  const port = Number(value);
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    throw new Error('Port must be an integer between 0 and 65535.');
  }
  return port;
}

if (require.main === module) {
  run().then((result) => {
    if (typeof result === 'number') {
      process.exitCode = result;
    }
  });
}

module.exports = {
  parseArgs,
  run,
  usage,
};
