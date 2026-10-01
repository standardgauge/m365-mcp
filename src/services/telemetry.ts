/**
 * Application Insights initialization — auto-patches console.log/warn/error
 * to forward to App Insights when APPLICATIONINSIGHTS_CONNECTION_STRING is set.
 *
 * Import this module early (before any console.log calls you want captured).
 * In local dev without the env var, this is a no-op and console.log works normally.
 */
import * as appInsights from 'applicationinsights';

const connectionString = process.env.APPLICATIONINSIGHTS_CONNECTION_STRING;

if (connectionString) {
  appInsights
    .setup(connectionString)
    .setAutoCollectConsole(true, true)   // patch console.log + console.error
    .setAutoCollectExceptions(true)       // unhandled exceptions
    .setAutoCollectDependencies(false)    // DISABLED — conflicts with Functions runtime gRPC
    .setAutoCollectRequests(false)        // Azure Functions handles request telemetry
    .setAutoCollectPerformance(false, false)  // DISABLED — Functions runtime handles this
    .setSendLiveMetrics(false)            // DISABLED — reduce overhead
    .start();

  console.log('[telemetry] Application Insights initialized — console output forwarded');
} else {
  console.log('[telemetry] No APPLICATIONINSIGHTS_CONNECTION_STRING — App Insights disabled');
}

export { appInsights };
