#!/bin/bash

set -e

cd "$(dirname "$0")"

echo
echo "========================================"
echo "  MailNotes Production Build"
echo "========================================"
echo

echo "[1/3] Webpack Production Build..."
npm run build

echo
echo "[2/3] GitHub-Pages-Verzeichnis aktualisieren..."
rm -rf docs
mkdir -p docs
cp -R dist/. docs/

echo
echo "[3/3] Manifest aktualisieren..."
node scripts/updateManifest.js

echo
echo "========================================"
echo "  Production Build erfolgreich"
echo "========================================"
echo
echo "Ausgabe:"
echo "  dist/"
echo "  docs/"
echo
