# Azure Deployment Checklist

Tool: Werkhaus Word Bild Tool
Produktive URL: https://tool2.wh-sv.de
PROD Manifest: manifests/manifest.prod.xml

Diese Datei bereitet das Deployment vor. Sie loest kein Deployment aus.

## Vorbedingungen
- Lokaler DEV-Test in Word ist erfolgreich.
- Pull Request von `develop` nach `main` ist reviewed.
- `npm ci`, `npm run build`, `npm run lint` sind gruen.
- `npm run validate:prod` ist gruen oder ein externer Microsoft-Validator-Ausfall ist dokumentiert und spaeter erneut geprueft.
- `manifests/manifest.prod.xml` enthaelt kein `TEST`.
- `manifests/manifest.prod.xml` zeigt nicht auf `localhost`.
- `manifests/manifest.prod.xml` zeigt auf `https://tool2.wh-sv.de`.

## Azure Resource pruefen
- Azure Static Web App oder App Service fuer Werkhaus Word Bild Tool im Azure Portal identifizieren.
- Custom Domain `https://tool2.wh-sv.de` an genau diese Ressource binden.
- TLS-Zertifikat fuer die Custom Domain pruefen.
- Build-Ausgabe ist `dist`.
- PROD-Web-App liefert mindestens diese Dateien aus:
  - `/taskpane.html`
  - `/commands.html`
  - `/assets/...`
  - optional `/manifests/manifest.prod.xml` als Referenz, aber Microsoft 365 verwendet die lokal gepruefte Manifest-Datei.

## GitHub Actions / Secrets
- Erwartetes Deployment Secret: `AZURE_STATIC_WEB_APPS_API_TOKEN_VICTORIOUS_CLIFF_0407AA503`
- Secret niemals ins Repository schreiben.
- Deployment erst nach Merge in den stabilen PROD-Branch aktiv nutzen.
- Offen: Lokaler PROD-Ordner ist nicht automatisch Beweis fuer den GitHub-PROD-Branch. Branch-Strategie vor Deployment bestaetigen.

## Smoke-Test nach Azure Deployment
1. `curl -I https://tool2.wh-sv.de/taskpane.html`
2. `curl -I https://tool2.wh-sv.de/commands.html`
3. `curl -I https://tool2.wh-sv.de`
4. `npm run validate:prod` erneut ausfuehren.
5. PROD-Manifest in Word/Microsoft 365 nur mit Testgruppe pruefen.

## Rollback
1. Vorherigen GitHub Release Tag auswaehlen.
2. Azure Static Web App auf Artefakt/Commit dieses Tags zuruecksetzen.
3. Wenn Manifest-Metadaten geaendert wurden: vorheriges `manifest.prod.xml` erneut im Microsoft 365 Admin Center bereitstellen.
4. Word neu starten und mit Testgruppe gegenpruefen.

## Referenzen
- Microsoft Learn: Office Add-ins Manifest validieren.
- GitHub Docs: Azure Static Web Apps Deployment via `Azure/static-web-apps-deploy`.
