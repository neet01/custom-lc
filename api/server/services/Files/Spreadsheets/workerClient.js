const path = require('path');
const { isEnabled } = require('@librechat/api');

const DEFAULT_WORKER_URL = 'http://spreadsheet_worker:8081';
const DEFAULT_TIMEOUT_MS = 120000;
const PYTHON_WORKER_EXTENSIONS = new Set(['.xlsx', '.xlsm', '.csv']);

class SpreadsheetWorkerError extends Error {
  constructor(message, options = {}) {
    super(message);
    this.name = 'SpreadsheetWorkerError';
    this.code = options.code ?? 'SPREADSHEET_WORKER_ERROR';
    this.status = options.status ?? 500;
    this.cause = options.cause;
  }
}

class SpreadsheetWorkerUnavailableError extends SpreadsheetWorkerError {
  constructor(message, options = {}) {
    super(message, {
      ...options,
      code: options.code ?? 'SPREADSHEET_WORKER_UNAVAILABLE',
      status: options.status ?? 503,
    });
    this.name = 'SpreadsheetWorkerUnavailableError';
  }
}

function isSpreadsheetWorkerEnabled() {
  return (
    isEnabled(process.env.SPREADSHEET_WORKER_ENABLED) ||
    isEnabled(process.env.ENABLE_SPREADSHEET_WORKER)
  );
}

function shouldFallbackToJs() {
  const value = process.env.SPREADSHEET_WORKER_FALLBACK_TO_JS;
  return value == null ? false : isEnabled(value);
}

function getSpreadsheetWorkerUrl() {
  return process.env.SPREADSHEET_WORKER_URL || DEFAULT_WORKER_URL;
}

function getSpreadsheetWorkerTimeoutMs() {
  const parsed = Number(process.env.SPREADSHEET_WORKER_TIMEOUT_MS);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : DEFAULT_TIMEOUT_MS;
}

function supportsPythonSpreadsheetWorker(sourceFilename) {
  const extension = path.extname(sourceFilename || '').toLowerCase();
  return PYTHON_WORKER_EXTENSIONS.has(extension);
}

function shouldUseSpreadsheetWorker(sourceFilename) {
  return isSpreadsheetWorkerEnabled() && supportsPythonSpreadsheetWorker(sourceFilename);
}

async function parseWorkerResponse(response) {
  const text = await response.text();
  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch (_error) {
    return text;
  }
}

async function requestSpreadsheetWorker(endpoint, payload) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), getSpreadsheetWorkerTimeoutMs());

  try {
    const response = await fetch(`${getSpreadsheetWorkerUrl()}${endpoint}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });

    const body = await parseWorkerResponse(response);
    if (response.ok) {
      return body;
    }

    const detail =
      body && typeof body === 'object' && 'detail' in body && body.detail && typeof body.detail === 'object'
        ? body.detail
        : {};
    throw new SpreadsheetWorkerError(
      detail.message || `Spreadsheet worker request failed with status ${response.status}`,
      {
        status: response.status,
        code: detail.code || 'SPREADSHEET_WORKER_REQUEST_FAILED',
      },
    );
  } catch (error) {
    if (error instanceof SpreadsheetWorkerError) {
      throw error;
    }

    if (error?.name === 'AbortError') {
      throw new SpreadsheetWorkerUnavailableError('Spreadsheet worker request timed out', {
        code: 'SPREADSHEET_WORKER_TIMEOUT',
        cause: error,
      });
    }

    throw new SpreadsheetWorkerUnavailableError('Spreadsheet worker is unavailable', {
      cause: error,
    });
  } finally {
    clearTimeout(timeout);
  }
}

async function inspectSpreadsheetWithWorker({ buffer, sourceFilename, maxPreviewRows }) {
  return requestSpreadsheetWorker('/inspect-workbook', {
    bufferBase64: buffer.toString('base64'),
    sourceFilename,
    maxPreviewRows,
  });
}

async function transformSpreadsheetWithWorker({
  buffer,
  sourceFilename,
  removeColumns,
  keepColumns,
  redactColumns,
  redactionText,
  sheetNames,
  outputFormat,
  operations,
}) {
  const response = await requestSpreadsheetWorker('/apply-plan', {
    bufferBase64: buffer.toString('base64'),
    sourceFilename,
    removeColumns,
    keepColumns,
    redactColumns,
    redactionText,
    sheetNames,
    outputFormat,
    operations,
  });

  return {
    ...response,
    buffer: Buffer.from(response.bufferBase64, 'base64'),
  };
}

module.exports = {
  SpreadsheetWorkerError,
  SpreadsheetWorkerUnavailableError,
  getSpreadsheetWorkerUrl,
  getSpreadsheetWorkerTimeoutMs,
  inspectSpreadsheetWithWorker,
  isSpreadsheetWorkerEnabled,
  shouldFallbackToJs,
  shouldUseSpreadsheetWorker,
  supportsPythonSpreadsheetWorker,
  transformSpreadsheetWithWorker,
};
