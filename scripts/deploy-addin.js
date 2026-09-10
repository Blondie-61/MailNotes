const fs = require("fs");
const os = require("os");
const path = require("path");

function fail(message) {
  console.error(`MailNotes: ${message}`);
  process.exit(1);
}

// Das automatische Deployment ist bewusst nur für macOS gedacht.
// Unter Windows wird der Production-Build separat in das Agent-Setup übernommen.
if (process.platform !== "darwin") {
  console.log("MailNotes: automatisches Add-in-Deployment wird nur unter macOS ausgeführt.");
  process.exit(0);
}

const projectRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(projectRoot, "dist");
const targetDir = path.join(
  os.homedir(),
  "Library",
  "Application Support",
  "MailNotes",
  "Addin"
);

if (!fs.existsSync(sourceDir)) {
  fail(`Build-Verzeichnis fehlt: ${sourceDir}`);
}

const taskpaneFile = path.join(sourceDir, "taskpane.html");
if (!fs.existsSync(taskpaneFile)) {
  fail(`Production-Build ist unvollständig: ${taskpaneFile} fehlt.`);
}

try {
  fs.mkdirSync(path.dirname(targetDir), { recursive: true });
  fs.rmSync(targetDir, { recursive: true, force: true });
  fs.mkdirSync(targetDir, { recursive: true });
  fs.cpSync(sourceDir, targetDir, { recursive: true });

  // Eine versehentliche lokale Sicherung darf nicht ausgeliefert werden.
  fs.rmSync(path.join(targetDir, "manifest copy.xml"), { force: true });
} catch (error) {
  fail(`Deployment nach ${targetDir} fehlgeschlagen: ${error.message}`);
}

console.log("");
console.log("MailNotes: Production-Add-in wurde aktualisiert:");
console.log(`  ${targetDir}`);
console.log("MailNotes: Outlook-Taskpane jetzt neu laden bzw. Outlook neu starten.");
