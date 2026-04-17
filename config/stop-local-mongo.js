require('dotenv').config();

const fs = require('fs');
const path = require('path');

const stateFile = path.resolve(__dirname, '..', 'api', 'data', '.local-mongo.json');

function readState() {
  try {
    return JSON.parse(fs.readFileSync(stateFile, 'utf8'));
  } catch {
    return null;
  }
}

function removeStateFile() {
  try {
    fs.unlinkSync(stateFile);
  } catch {
    // Best-effort cleanup for local dev state only.
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

async function waitForExit(pid) {
  const deadline = Date.now() + 10_000;

  while (Date.now() < deadline) {
    if (!processAlive(pid)) {
      return true;
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }

  return false;
}

async function main() {
  const state = readState();
  if (!state?.pid) {
    console.log('[local-mongo] No running local Mongo instance found.');
    removeStateFile();
    return;
  }

  if (!processAlive(state.pid)) {
    console.log(`[local-mongo] Stale state file found for pid ${state.pid}; cleaning up.`);
    removeStateFile();
    return;
  }

  process.kill(state.pid, 'SIGTERM');
  const stopped = await waitForExit(state.pid);
  removeStateFile();

  if (!stopped) {
    throw new Error(
      `Timed out waiting for local Mongo process ${state.pid} to stop. Kill it manually if needed.`,
    );
  }

  console.log(`[local-mongo] Stopped local Mongo process ${state.pid}.`);
}

void main().catch((error) => {
  console.error('[local-mongo] Failed to stop:', error);
  process.exit(1);
});
