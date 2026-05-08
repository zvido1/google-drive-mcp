// ---------------------------------------------------------------------------
// External authentication modes: Service Account & pre-obtained OAuth tokens
// ---------------------------------------------------------------------------

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { OAuth2Client } from 'google-auth-library';
import { GoogleAuth } from 'google-auth-library';
import { DEFAULT_SCOPES } from './scopes.js';

// ---------------------------------------------------------------------------
// Service Account mode
// ---------------------------------------------------------------------------

/**
 * If `GOOGLE_DRIVE_MCP_SERVICE_ACCOUNT_JSON` is set, parse its contents as a
 * service-account JSON object, write it to a temporary file, and point
 * `GOOGLE_APPLICATION_CREDENTIALS` at that file.
 *
 * This allows Railway (and similar platforms) to supply the full service-account
 * JSON as a secret environment variable rather than mounting a file.
 *
 * Call this once, early in the auth flow, before `isServiceAccountMode()`.
 */
export function setupServiceAccountFromEnv(): void {
  const jsonContent = process.env.GOOGLE_DRIVE_MCP_SERVICE_ACCOUNT_JSON?.trim();
  if (!jsonContent) return;

  // If GOOGLE_APPLICATION_CREDENTIALS is already set, respect it and skip.
  if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    console.error(
      'GOOGLE_DRIVE_MCP_SERVICE_ACCOUNT_JSON is set but GOOGLE_APPLICATION_CREDENTIALS ' +
        'already points to a file — skipping env-var service account setup.'
    );
    return;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(jsonContent);
  } catch {
    throw new Error(
      'GOOGLE_DRIVE_MCP_SERVICE_ACCOUNT_JSON is set but its value is not valid JSON. ' +
        'Ensure the environment variable contains the raw JSON of your service account key file.'
    );
  }

  // Write to a temp file that persists for the lifetime of the process.
  const tmpFile = path.join(os.tmpdir(), `gcp-sa-${process.pid}.json`);
  fs.writeFileSync(tmpFile, JSON.stringify(parsed), { encoding: 'utf8', mode: 0o600 });

  process.env.GOOGLE_APPLICATION_CREDENTIALS = tmpFile;
  console.error(
    `Service account JSON loaded from GOOGLE_DRIVE_MCP_SERVICE_ACCOUNT_JSON ` +
      `and written to temporary file: ${tmpFile}`
  );
}

/** True when `GOOGLE_APPLICATION_CREDENTIALS` is set (standard Google convention). */
export function isServiceAccountMode(): boolean {
  return !!process.env.GOOGLE_APPLICATION_CREDENTIALS;
}

/**
 * Create an authorized client from a service account JSON key file.
 * `GoogleAuth` handles JWT signing and token refresh automatically.
 */
export async function createServiceAccountAuth(): Promise<any> {
  const keyFile = process.env.GOOGLE_APPLICATION_CREDENTIALS!;
  console.error(`Using service account credentials from ${keyFile}`);

  const auth = new GoogleAuth({
    keyFile,
    scopes: [...DEFAULT_SCOPES],
  });

  const client = await auth.getClient();
  console.error('Service account authentication successful');
  return client;
}

// ---------------------------------------------------------------------------
// Auth type detection
// ---------------------------------------------------------------------------

/**
 * Returns a human-readable string describing the active authentication mode.
 * Useful for diagnostics and error messages surfaced to the user.
 */
export function getAuthType(): 'service-account' | 'external-token' | 'oauth2' {
  if (isServiceAccountMode()) return 'service-account';
  if (isExternalTokenMode()) return 'external-token';
  return 'oauth2';
}

// ---------------------------------------------------------------------------
// External OAuth Token mode
// ---------------------------------------------------------------------------

/** True when `GOOGLE_DRIVE_MCP_ACCESS_TOKEN` is set. */
export function isExternalTokenMode(): boolean {
  return !!process.env.GOOGLE_DRIVE_MCP_ACCESS_TOKEN;
}

/**
 * Validate that the env-var combination makes sense.
 * Throws with an actionable message on mis-configuration.
 */
export function validateExternalTokenConfig(): void {
  const accessToken = process.env.GOOGLE_DRIVE_MCP_ACCESS_TOKEN?.trim();
  if (!accessToken) {
    throw new Error(
      'GOOGLE_DRIVE_MCP_ACCESS_TOKEN is set but empty. Provide a valid OAuth access token.'
    );
  }

  const refreshToken = process.env.GOOGLE_DRIVE_MCP_REFRESH_TOKEN?.trim();
  const clientId = process.env.GOOGLE_DRIVE_MCP_CLIENT_ID?.trim();
  const clientSecret = process.env.GOOGLE_DRIVE_MCP_CLIENT_SECRET?.trim();

  if (refreshToken) {
    if (!clientId || !clientSecret) {
      throw new Error(
        'GOOGLE_DRIVE_MCP_REFRESH_TOKEN is set but GOOGLE_DRIVE_MCP_CLIENT_ID and/or ' +
          'GOOGLE_DRIVE_MCP_CLIENT_SECRET are missing. All three are required for automatic token refresh.'
      );
    }
  }

  // Warn about partial client credential sets (one without the other)
  if ((clientId && !clientSecret) || (!clientId && clientSecret)) {
    throw new Error(
      'Both GOOGLE_DRIVE_MCP_CLIENT_ID and GOOGLE_DRIVE_MCP_CLIENT_SECRET must be provided together.'
    );
  }
}

/**
 * Create an OAuth2Client pre-loaded with externally-obtained credentials.
 * When a refresh token + client credentials are provided, the client will
 * auto-refresh transparently.
 */
export function createExternalOAuth2Client(): OAuth2Client {
  const accessToken = process.env.GOOGLE_DRIVE_MCP_ACCESS_TOKEN!.trim();
  const refreshToken = process.env.GOOGLE_DRIVE_MCP_REFRESH_TOKEN?.trim();
  const clientId = process.env.GOOGLE_DRIVE_MCP_CLIENT_ID?.trim();
  const clientSecret = process.env.GOOGLE_DRIVE_MCP_CLIENT_SECRET?.trim();

  const oauth2Client = new OAuth2Client(clientId, clientSecret);

  oauth2Client.setCredentials({
    access_token: accessToken,
    refresh_token: refreshToken || undefined,
  });

  if (!refreshToken) {
    console.error(
      'Warning: No refresh token provided. The access token will not auto-refresh when it expires.'
    );
  } else {
    console.error('External OAuth tokens configured with auto-refresh support.');
  }

  return oauth2Client;
}
