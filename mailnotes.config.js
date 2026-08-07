// MailNotes Add-in: zentrale Build-Konfiguration
// ------------------------------------------------------------
// Nur HIER werden Unterschiede zwischen Entwicklung und Produktion gepflegt.
// Der aktive Block wird automatisch durch `webpack --mode development` bzw.
// `webpack --mode production` ausgewählt.

module.exports = {
  development: {
    modeName: "Development",
    baseUrl: "https://localhost:3000/",

    // Das Taskpane verwendet im Development-Build immer /api.
    // webpack leitet /api an dieses Ziel weiter.
    // Mac -> Windows-VM z.B.: http://10.211.55.7:48571
    // proxyTarget: "http://127.0.0.1:48571",
    proxyTarget: "http://10.211.55.7:48571",

    enableLogging: false,
  },

  production: {
    modeName: "Production",

    // Öffentliche HTTPS-Adresse, unter der dist/ bereitgestellt wird.
    baseUrl: "https://blondie-61.github.io/MailNotes/",

    // Im Produktivbetrieb läuft der Agent auf demselben Rechner wie Outlook.
    agentUrl: "http://127.0.0.1:48571",

    enableLogging: false,
  },
};
