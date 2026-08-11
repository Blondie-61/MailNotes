#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const config = require("../mailnotes.config");

const manifestFile = path.join(__dirname, "..", "docs", "manifest.xml");

if (!fs.existsSync(manifestFile)) {
  console.error("Manifest nicht gefunden:");
  console.error(manifestFile);
  process.exit(1);
}

let xml = fs.readFileSync(manifestFile, "utf8");

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function replaceSingleTag(xmlText, tagName, value) {
  const re = new RegExp(`<${tagName}>[^<]*</${tagName}>`, "g");
  return xmlText.replace(
    re,
    `<${tagName}>${xmlEscape(value)}</${tagName}>`
  );
}

function replaceDefaultValueTag(xmlText, tagName, value) {
  const re = new RegExp(
    `(<${tagName}\\b[^>]*\\bDefaultValue=")[^"]*(")`,
    "g"
  );
  return xmlText.replace(re, `$1${xmlEscape(value)}$2`);
}

function replaceUrlResource(xmlText, resourceId, value) {
  const re = new RegExp(
    `(<bt:Url\\s+id="${resourceId}"\\s+DefaultValue=")[^"]*(")`,
    "g"
  );
  return xmlText.replace(re, `$1${xmlEscape(value)}$2`);
}

function replaceImageResource(xmlText, resourceId, value) {
  const re = new RegExp(
    `(<bt:Image\\s+id="${resourceId}"\\s+DefaultValue=")[^"]*(")`,
    "g"
  );
  return xmlText.replace(re, `$1${xmlEscape(value)}$2`);
}

const baseUrl = config.production.baseUrl.replace(/\/+$/, "");

xml = replaceSingleTag(xml, "Version", config.version);
xml = replaceSingleTag(xml, "ProviderName", config.manifest.provider);
xml = replaceSingleTag(xml, "DefaultLocale", config.manifest.defaultLocale);
xml = replaceDefaultValueTag(xml, "DisplayName", config.manifest.displayName);
xml = replaceDefaultValueTag(xml, "Description", config.manifest.description);
xml = replaceDefaultValueTag(xml, "SupportUrl", config.manifest.supportUrl);

xml = replaceDefaultValueTag(
  xml,
  "IconUrl",
  `${baseUrl}/assets/icon-64.png`
);

xml = replaceDefaultValueTag(
  xml,
  "HighResolutionIconUrl",
  `${baseUrl}/assets/icon-128.png`
);

xml = xml.replace(
  /<SourceLocation\s+DefaultValue="[^"]*"\s*\/>/g,
  `<SourceLocation DefaultValue="${xmlEscape(baseUrl + "/taskpane.html")}"/>`
);

xml = replaceUrlResource(
  xml,
  "Commands.Url",
  `${baseUrl}/commands.html`
);

xml = replaceUrlResource(
  xml,
  "Taskpane.Url",
  `${baseUrl}/taskpane.html`
);

xml = replaceImageResource(
  xml,
  "Icon.16x16",
  `${baseUrl}/assets/icon-16.png`
);

xml = replaceImageResource(
  xml,
  "Icon.32x32",
  `${baseUrl}/assets/icon-32.png`
);

xml = replaceImageResource(
  xml,
  "Icon.80x80",
  `${baseUrl}/assets/icon-80.png`
);

xml = xml.replace(
  /<AppDomains>[\s\S]*?<\/AppDomains>/,
  [
    "<AppDomains>",
    `  <AppDomain>${xmlEscape(baseUrl)}</AppDomain>`,
    "</AppDomains>"
  ].join("\n")
);

fs.writeFileSync(manifestFile, xml, "utf8");

console.log("Manifest erfolgreich aktualisiert.");
console.log("Version :", config.version);
console.log("BaseURL :", baseUrl);
console.log("Taskpane:", `${baseUrl}/taskpane.html`);
