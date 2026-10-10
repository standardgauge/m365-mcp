/**
 * Session storage could not answer whether a token names a session. Thrown by
 * getSessionByToken so callers can tell "no such session" (undefined, a 401)
 * from "storage is down" (this, a 503). Fail closed: a request that hits this
 * is never authenticated. The message is fixed so handlers that echo
 * err.message never leak storage details; the storage error is the cause.
 *
 * Its own module so HTTP handlers can recognise it without importing the
 * session cache and, through it, Table Storage.
 */
export class SessionStoreUnavailableError extends Error {
  constructor(cause: unknown) {
    super('Session store unavailable', { cause });
    this.name = 'SessionStoreUnavailableError';
  }
}

/** Seconds a client should wait before retrying after a 503 from this. */
export const SESSION_STORE_RETRY_AFTER_S = 5;
