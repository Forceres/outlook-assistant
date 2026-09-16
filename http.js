const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const {
  StreamableHTTPServerTransport,
} = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const config = require('./config');
const { createRequestHandler } = require('./request-handler.js');
const { timingSafeEqual } = require('node:crypto');
const { createServer } = require('node:http');

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
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// F-1 / F-48: warn at startup when safety belts are unset. Mirrors the
// warning surfaced by `auth action=about`. Visible to operators reading
// stderr; AI clients reading the JSON-RPC stream are unaffected.
if (
  !process.env.OUTLOOK_MAX_EMAILS_PER_SESSION &&
  !process.env.OUTLOOK_ALLOWED_RECIPIENTS &&
  !config.USE_TEST_MODE
) {
  console.error(
    '⚠ Safety belts not configured. Consider setting OUTLOOK_MAX_EMAILS_PER_SESSION and OUTLOOK_ALLOWED_RECIPIENTS in your .mcp.json env block for safer AI-assisted sending. See `auth action=about` for details.'
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

// Create server with tools capabilities
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

const AUTH_TOKEN = process.env.AUTH_TOKEN;

const authenticate = (extra) => {
  const authorization = extra.requestInfo.headers.authorization;
  if (!authorization) throw new Error('Unauthorized!');

  const authBuf = Buffer.from(authorization);
  const validBuf = Buffer.from(`Bearer ${AUTH_TOKEN}`);

  if (
    authBuf.length === validBuf.length &&
    timingSafeEqual(authBuf, validBuf)
  ) {
    return;
  }
  throw new Error('Unauthorized!');
};

// Handle all requests. Dispatch + error-shaping logic lives in
// request-handler.js so it is unit-testable without starting the transport.
server.fallbackRequestHandler = createRequestHandler(TOOLS, authenticate);

const transport = new StreamableHTTPServerTransport({
  sessionIdGenerator: undefined,
});

const httpServer = createServer(async (req, res) => {
  try {
    if (req.url === '/health' && req.method === 'GET') {
      res.writeHead(200, {
        'content-type': 'application/json',
      });

      res.end(
        JSON.stringify({
          status: 'ok',
        })
      );

      return;
    }

    if (req.url !== '/mcp') {
      res.writeHead(404, {
        'content-type': 'application/json',
      });

      res.end(
        JSON.stringify({
          error: 'Not found',
        })
      );

      return;
    }

    await transport.handleRequest(req, res);
  } catch (error) {
    console.error('HTTP request error:', error);

    if (!res.headersSent) {
      res.writeHead(500, {
        'content-type': 'application/json',
      });

      res.end(
        JSON.stringify({
          error: 'Internal server error',
        })
      );
    }
  }
});

server
  .connect(transport)
  .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
  .catch((error) => {
    console.error(`Connection error: ${error.message}`);
    process.exit(1);
  });

const PORT = Number(process.env.PORT) || 3000;

httpServer.listen(PORT, '0.0.0.0', () => {
  console.error(`${config.SERVER_NAME} connected and listening on :${PORT}`);
});

async function shutdown(signal) {
  console.error(`${signal} received, shutting down...`);

  httpServer.close(async () => {
    await transport.close();
    await server.close();

    process.exit(0);
  });
}

process.on('SIGTERM', () => void shutdown('SIGTERM'));
process.on('SIGINT', () => void shutdown('SIGINT'));
