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

function replaceTag(xml, tag, value) {
    const re = new RegExp(`<${tag}>[^<]*</${tag}>`, "g");
    return xml.replace(re, `<${tag}>${value}</${tag}>`);
}

function replaceAttribute(xml, attribute, value) {
    const re = new RegExp(`${attribute}="[^"]*"`, "g");
    return xml.replace(re, `${attribute}="${value}"`);
}

function replaceDefaultValue(xml, id, value) {
    const re = new RegExp(
        `(<[^>]+id="${id}"[^>]+DefaultValue=")[^"]*(")`,
        "g"
    );

    return xml.replace(re, `$1${value}$2`);
}

//
// Version
//
xml = replaceTag(xml, "Version", config.version);

//
// Grunddaten
//
xml = replaceTag(xml, "ProviderName", config.manifest.provider);
xml = replaceTag(xml, "DefaultLocale", config.manifest.defaultLocale);

xml = xml.replace(
    /<DisplayName DefaultValue="[^"]*"/g,
    `<DisplayName DefaultValue="${config.manifest.displayName}"`
);

xml = xml.replace(
    /<Description DefaultValue="[^"]*"/g,
    `<Description DefaultValue="${config.manifest.description}"`
);

xml = xml.replace(
    /<SupportUrl DefaultValue="[^"]*"/g,
    `<SupportUrl DefaultValue="${config.manifest.supportUrl}"`
);

//
// URLs
//
const baseUrl = config.production.baseUrl.replace(/\/$/, "");

xml = xml.replace(
    /https:\/\/blondie-61\.github\.io\/MailNotes/g,
    baseUrl
);

xml = xml.replace(
    /https:\/\/blondie-61\.github\.io"/g,
    config.manifest.supportUrl + '"'
);

fs.writeFileSync(manifestFile, xml, "utf8");

console.log("Manifest erfolgreich aktualisiert.");
console.log("Version :", config.version);
console.log("BaseURL :", baseUrl);