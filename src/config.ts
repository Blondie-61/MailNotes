// Wird von webpack zur Build-Zeit gesetzt. Keine Laufzeit-Erkennung per Hostname.
declare const __MAILNOTES_BUILD_MODE__: string;
declare const __MAILNOTES_AGENT_URL__: string;
declare const __MAILNOTES_BASE_URL__: string;
declare const __MAILNOTES_ENABLE_LOGGING__: boolean;

export const MailNotesConfig = Object.freeze({
  buildMode: __MAILNOTES_BUILD_MODE__,
  isProduction: __MAILNOTES_BUILD_MODE__ === "production",
  agentUrl: __MAILNOTES_AGENT_URL__,
  baseUrl: __MAILNOTES_BASE_URL__,
  enableLogging: __MAILNOTES_ENABLE_LOGGING__,
});
