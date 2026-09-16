/**
 * Outlook Assistant — Streamable HTTP entry point.
 *
 * Runs the same tool set as index.js (stdio) behind an HTTP endpoint so the
 * server can be deployed in a container and reached by remote MCP clients.
 *
 * Transport model: STATELESS. The MCP SDK forbids reusing a stateless
 * `StreamableHTTPServerTransport` across HTTP requests (it throws
 * "Stateless transport cannot be reused across requests"), so a fresh
 * `Server` + transport pair is created for every request and torn down when
 * the response finishes. Tool handlers themselves are stateless module-level
 * functions, so there is no per-session state to lose.
 *
 * Auth: a static bearer token (`MCP_AUTH_TOKEN`) is checked at the HTTP layer
 * before the request ever reaches the MCP transport. Doing it here — not in
 * the JSON-RPC fallback handler — means `initialize` (which the SDK handles
 * itself) is covered too, and unauthenticated callers get a proper 401.
 */
const { randomUUID, timingSafeEqual } = require('node:crypto');
const { createServer } = require('node:http');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const {
  StreamableHTTPServerTransport,
} = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const config = require('./config');
const { createRequestHandler } = require('./request-handler');

// Import module tools
const { authTools, setToolCount } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { contactsTools } = require('./contacts');
const { categoriesTools } = require('./categories');
const { settingsTools } = require('./settings');
const { advancedTools } = require('./advanced');

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER (HTTP)`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// F-1 / F-48: warn at startup when safety belts are unset. Mirrors the
// warning surfaced by `auth action=about`.
if (
  !process.env.OUTLOOK_MAX_EMAILS_PER_SESSION &&
  !process.env.OUTLOOK_ALLOWED_RECIPIENTS &&
  !config.USE_TEST_MODE
) {
  console.error(
    '⚠ Safety belts not configured. Consider setting OUTLOOK_MAX_EMAILS_PER_SESSION and OUTLOOK_ALLOWED_RECIPIENTS for safer AI-assisted sending. See `auth action=about` for details.'
  );
}

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...contactsTools,
  ...categoriesTools,
  ...settingsTools,
  ...advancedTools,
];

// Set dynamic tool count for auth about handler
setToolCount(TOOLS.length);

// Dispatch + error-shaping logic lives in request-handler.js so it is
// unit-testable without starting a transport. One handler instance is shared
// by every per-request Server below.
const requestHandler = createRequestHandler(TOOLS);

// ---------------------------------------------------------------------------
// Bearer auth
// ---------------------------------------------------------------------------

// `MCP_AUTH_TOKEN` matches docker-compose.yml. `AUTH_TOKEN` is accepted as a
// fallback for anyone who wired it up against the earlier name.
const AUTH_TOKEN = process.env.MCP_AUTH_TOKEN;

if (!AUTH_TOKEN) {
  console.error(
    'MCP_AUTH_TOKEN is not set. Refusing to start: this server binds to 0.0.0.0 and exposes the mailbox tools to anyone who can reach the port.'
  );
  process.exit(1);
}

const EXPECTED_AUTH = Buffer.from(`Bearer ${AUTH_TOKEN}`);

/** @returns {boolean} true when the Authorization header carries the token */
function isAuthorized(req) {
  const header = req.headers.authorization;
  if (typeof header !== 'string') return false;
  const given = Buffer.from(header);
  return (
    given.length === EXPECTED_AUTH.length &&
    timingSafeEqual(given, EXPECTED_AUTH)
  );
}

function sendJson(res, status, body) {
  res.writeHead(status, { 'content-type': 'application/json' });
  res.end(JSON.stringify(body));
}

// ---------------------------------------------------------------------------
// Per-request MCP server + transport
// ---------------------------------------------------------------------------

function createMcpServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    {
      capabilities: {
        tools: TOOLS.reduce((acc, tool) => {
          acc[tool.name] = {};
          return acc;
        }, {}),
      },
    }
  );
  server.fallbackRequestHandler = requestHandler;
  return server;
}

async function handleMcpRequest(req, res) {
  const server = createMcpServer();
  const transport = new StreamableHTTPServerTransport({
    // Stateless: no session IDs, one transport per HTTP request.
    sessionIdGenerator: undefined,
  });
  transport.onerror = (error) => {
    console.error(`MCP transport error [${req.method} ${req.url}]:`, error);
  };

  // Tear down once the client has the full response (or drops the socket)
  // so per-request objects don't accumulate.
  res.on('close', () => {
    transport.close().catch(() => {});
    server.close().catch(() => {});
  });

  await server.connect(transport);
  await transport.handleRequest(req, res);
}

const httpServer = createServer(async (req, res) => {
  const requestId = randomUUID().slice(0, 8);
  try {
    if (req.url === '/health' && req.method === 'GET') {
      sendJson(res, 200, { status: 'ok' });
      return;
    }

    if (req.url !== '/mcp') {
      sendJson(res, 404, { error: 'Not found' });
      return;
    }

    if (!isAuthorized(req)) {
      console.error(`[${requestId}] Unauthorized ${req.method} ${req.url}`);
      res.writeHead(401, {
        'content-type': 'application/json',
        'www-authenticate': 'Bearer',
      });
      res.end(
        JSON.stringify({
          jsonrpc: '2.0',
          error: { code: -32000, message: 'Unauthorized' },
          id: null,
        })
      );
      return;
    }

    await handleMcpRequest(req, res);
  } catch (error) {
    console.error(`[${requestId}] HTTP request error:`, error);

    if (!res.headersSent) {
      sendJson(res, 500, {
        jsonrpc: '2.0',
        error: { code: -32603, message: 'Internal server error' },
        id: null,
      });
    } else {
      res.end();
    }
  }
});

const PORT = Number(process.env.PORT) || 3000;

httpServer.listen(PORT, '0.0.0.0', () => {
  console.error(
    `${config.SERVER_NAME} listening on :${PORT} (POST /mcp, GET /health)`
  );
});

function shutdown(signal) {
  console.error(`${signal} received, shutting down...`);
  httpServer.close(() => process.exit(0));
  // Don't hang forever on lingering keep-alive connections.
  setTimeout(() => process.exit(0), 5000).unref();
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
