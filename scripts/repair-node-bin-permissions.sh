#!/usr/bin/env bash
set -euo pipefail

if [ ! -d node_modules/.bin ]; then
  echo "node_modules/.bin nicht gefunden. Bitte zuerst 'npm ci' ausführen."
  exit 1
fi

fixed=0
while IFS= read -r -d '' file; do
  if [ ! -x "$file" ]; then
    chmod +x "$file"
    fixed=$((fixed + 1))
  fi
done < <(find node_modules/.bin -type l -print0)

echo "Berechtigungen geprüft. Angepasste Einträge: $fixed"
