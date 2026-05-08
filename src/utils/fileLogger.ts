/**
 * Synchronous file-based logger for diagnostic purposes.
 *
 * Writes to /tmp/mcp-errors.log using writeFileSync so that data is flushed
 * immediately and is never lost due to buffering in the HTTP transport layer.
 * This is intentionally a fallback for situations where console.error output
 * is swallowed before it reaches Railway's log capture.
 */

import { writeFileSync, appendFileSync, existsSync } from 'fs';

const LOG_PATH = '/tmp/mcp-errors.log';

function formatEntry(level: string, message: string, data?: any): string {
  const timestamp = new Date().toISOString();
  const dataStr = data !== undefined ? `\n  data: ${JSON.stringify(data, null, 2)}` : '';
  return `[${timestamp}] [${level}] ${message}${dataStr}\n`;
}

/**
 * Append a line to the log file synchronously.
 * Safe to call from any context — never throws.
 */
function writeToFile(entry: string): void {
  try {
    appendFileSync(LOG_PATH, entry, { encoding: 'utf8' });
  } catch {
    // If we can't write to /tmp, there is nothing more we can do.
  }
}

/**
 * Log an informational message to the file.
 */
export function fileLog(message: string, data?: any): void {
  writeToFile(formatEntry('INFO', message, data));
}

/**
 * Log an error to the file. Accepts an Error object or any value.
 */
export function fileLogError(message: string, error?: unknown): void {
  let errorData: Record<string, unknown> | undefined;

  if (error !== undefined) {
    if (error instanceof Error) {
      errorData = {
        name: error.name,
        message: error.message,
        stack: error.stack,
      };
    } else {
      errorData = { raw: String(error) };
    }
  }

  writeToFile(formatEntry('ERROR', message, errorData));
}

/**
 * Write a separator + header so individual runs are easy to distinguish
 * when tailing the log file.
 */
export function fileLogSessionStart(): void {
  const entry =
    `\n${'='.repeat(72)}\n` +
    `SESSION START  ${new Date().toISOString()}\n` +
    `${'='.repeat(72)}\n`;
  writeToFile(entry);
}

/**
 * Return the path of the log file so callers can surface it in error messages.
 */
export function getLogPath(): string {
  return LOG_PATH;
}
