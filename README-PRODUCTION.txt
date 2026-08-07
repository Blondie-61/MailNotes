MailNotes Outlook Add-in – eine Codebasis
=========================================

Zentrale Konfiguration
----------------------
Alle Unterschiede zwischen Test/Entwicklung und Echtbetrieb stehen in:

  mailnotes.config.js

Development
-----------
Start:

  npm run dev-server

Webpack wird mit --mode development gestartet. Das Taskpane verwendet /api;
der webpack-dev-server leitet die Aufrufe an development.proxyTarget weiter.
Wenn der Agent in der Windows-VM läuft, dort z.B. eintragen:

  proxyTarget: "http://10.211.55.7:48571"

Production
----------
Build:

  BuildProduction.cmd

oder:

  npm run build

Webpack wird mit --mode production gestartet. Das Taskpane verwendet direkt
production.agentUrl. Das Manifest wird beim Build automatisch von localhost
auf production.baseUrl umgeschrieben.

Es gibt KEINE zweite Quellcodekopie für Production.
