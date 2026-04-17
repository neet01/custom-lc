require('dotenv').config();

const fs = require('fs');
const path = require('path');
const { MongoMemoryServer } = require('mongodb-memory-server');

const stateFile = path.resolve(__dirname, '..', 'api', 'data', '.local-mongo.json');
const port = Number.parseInt(process.env.LOCAL_MONGO_PORT || '27017', 10);
const ip = process.env.LOCAL_MONGO_IP || '127.0.0.1';
const dbName = process.env.LOCAL_MONGO_DB || 'LibreChat';
const version = process.env.LOCAL_MONGO_VERSION;

/** @type {MongoMemoryServer | null} */
let mongoServer = null;
let shuttingDown = false;

function ensureStateDir() {
  fs.mkdirSync(path.dirname(stateFile), { recursive: true });
}

function readState() {
  try {
    return JSON.parse(fs.readFileSync(stateFile, 'utf8'));
  } catch {
    return null;
  }
}

function processAlive(pid) {
  if (!pid) {
    return false;
  }

  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function writeState(uri) {
  ensureStateDir();
  const state = {
    pid: process.pid,
    port,
    ip,
    dbName,
    uri,
    startedAt: new Date().toISOString(),
  };
  fs.writeFileSync(stateFile, `${JSON.stringify(state, null, 2)}\n`);
}

function clearState() {
  try {
    fs.unlinkSync(stateFile);
  } catch {
    // Best-effort cleanup for local dev state only.
  }
}

async function shutdown(signal) {
  if (shuttingDown) {
    return;
  }
  shuttingDown = true;

  console.log(`[local-mongo] Shutting down (${signal})`);
  clearState();

  if (mongoServer) {
    await mongoServer.stop();
  }

  process.exit(0);
}

async function main() {
  const state = readState();
  if (state?.pid && processAlive(state.pid)) {
    console.log(
      `[local-mongo] Already running at ${state.uri} (pid ${state.pid}). ` +
        `Use "npm run mongo:dev:stop" before starting a new instance.`,
    );
    return;
  }

  clearState();

  mongoServer = await MongoMemoryServer.create({
    binary: version ? { version } : undefined,
    instance: {
      dbName,
      ip,
      port,
      storageEngine: 'wiredTiger',
    },
  });

  const uri = mongoServer.getUri();
  writeState(uri);

  console.log(`[local-mongo] Ready at ${uri}`);
  console.log(`[local-mongo] State file: ${stateFile}`);

  process.on('SIGINT', () => {
    void shutdown('SIGINT');
  });
  process.on('SIGTERM', () => {
    void shutdown('SIGTERM');
  });
  process.on('SIGHUP', () => {
    void shutdown('SIGHUP');
  });

  setInterval(() => {
    if (!fs.existsSync(stateFile)) {
      void shutdown('missing-state-file');
    }
  }, 5_000);
}

void main().catch((error) => {
  console.error('[local-mongo] Failed to start:', error);
  clearState();
  process.exit(1);
});
