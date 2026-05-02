# MAC TESTING - Werkhaus Word Bild Tool

## Voraussetzungen
- macOS mit Microsoft Word
- Node.js + npm

## Setup
1. `npm install`
2. `npx office-addin-dev-certs install`
3. Word beenden und neu starten

## DEV starten
1. `npm run dev-server`
2. neues Terminal: `npm run start:dev`
3. In Word erscheint: **WERKHAUS TEST Word & Bild**

## Validierung
- `npm run validate:dev`
- `npm run validate:prod`

## Fehlerbehebung
- Word Cache loeschen: `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`
- Danach Word neu starten und `npm run stop:dev` + `npm run start:dev`

## Manifestdaten
- DEV Manifest: `manifests/manifest.dev.xml`
- PROD Manifest: `manifests/manifest.prod.xml`
- DEV Add-in ID: `9a3a0ed8-f310-4cf5-bf65-845bb9b94fd1`
- PROD Add-in ID: `c2c1d8a4-4d33-4e0d-b4c2-7b6d2f9a8e41`
- DEV URL: `https://localhost:3001`
- PROD URL: `https://tool2.wh-sv.de`
