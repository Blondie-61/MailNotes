module.exports = {
  version: "1.0.0.2",

  development: {
    baseUrl: "https://localhost:3000/",
    proxyTarget: "http://10.211.55.7:48571",
    agentUrl: "/api",
    enableLogging: false
  },

  production: {
    baseUrl: "https://localhost:48571/",
    agentUrl: "",
    enableLogging: false
  },

  manifest: {
    id: "DF769706-2458-41DB-B120-B45BA42CBB9C",
    displayName: "MailNotes",
    provider: "MailNotes",
    description: "MailNotes für Outlook.",
    supportUrl: "https://blondie-61.github.io",
    defaultLocale: "en-US"
  }
};
