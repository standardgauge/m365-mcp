import { requireCryptoConfigured } from '../services/credentialCrypto.js';

/**
 * Backstop for the application-key startup gate.
 *
 * The image's CMD runs src/startup/checkKeys.ts before the Functions host, which
 * is what stops a misconfigured container. Any launch path that skips that CMD
 * (`func start` locally, a command override on the Container App) still loads
 * every file under dist/functions/ as an entry point, and this one throws while
 * loading, so the worker registers no functions and /health never answers.
 */
requireCryptoConfigured();
