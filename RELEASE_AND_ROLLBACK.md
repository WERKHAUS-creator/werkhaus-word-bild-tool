# Release and Rollback

Tool: Werkhaus Word Bild Tool
PROD URL: https://tool2.wh-sv.de
PROD Manifest: manifests/manifest.prod.xml

## Branching
- `develop`: Arbeits- und Testversion.
- `main`: stabile Produktionsbasis, sofern nicht projektbezogen ein anderer Deploy-Branch bestaetigt wurde.
- Release nur per Pull Request von `develop` nach `main`.
- Direkte PROD-Aenderungen vermeiden.

## Pflichtchecks vor Merge
- `npm ci`
- `npm run build`
- `npm run lint`
- `npm run validate:dev`
- `npm run validate:prod`

Wenn der Microsoft-Validator extern nicht erreichbar ist, darf das nicht als Codefehler bewertet werden. Vor Microsoft-365-Rollout muss die PROD-Validierung aber erneut erfolgreich laufen.

## Release vorbereiten
1. DEV lokal in Word testen.
2. Develop-Stand auf GitHub sichern.
3. Pull Request `develop -> main` erstellen.
4. CI-Checks pruefen.
5. `manifests/manifest.prod.xml` manuell kontrollieren:
   - keine TEST-Bezeichnung
   - keine localhost-URL
   - PROD URL ist `https://tool2.wh-sv.de`
6. PR reviewen und erst dann mergen.

## Release veroeffentlichen
1. Merge nach `main`.
2. Release Tag setzen, z. B. `v1.0.0`.
3. Azure Deployment aus dem bestaetigten PROD-Branch ausfuehren.
4. Live-URL `https://tool2.wh-sv.de` pruefen.
5. PROD-Manifest validieren.
6. Microsoft 365 Admin Center: nur `manifests/manifest.prod.xml` mit kleiner Testgruppe bereitstellen.
7. Word Smoke-Test durchfuehren.
8. Danach breiteren Rollout freigeben.

## Rollback
1. Letzten funktionierenden Release Tag identifizieren.
2. Azure auf diesen Stand zuruecksetzen.
3. Wenn noetig vorheriges PROD-Manifest wiederherstellen.
4. Word und Office-Cache in Testumgebung bereinigen.
5. Testgruppe erneut pruefen.

## Dateien, die nicht direkt ueberschrieben werden duerfen
- `manifests/manifest.prod.xml`
- GitHub Actions Deployment Workflows
- Azure Secrets
- Microsoft 365 Admin Center App-Konfiguration

## Sicherheitsregeln
- Keine Secrets committen.
- Keine Zertifikate oder privaten Keys committen.
- `node_modules` und `dist` nicht versehentlich versionieren.
- Kein Azure Deployment ohne gruenen PROD-Check.
- Kein Microsoft-365-Rollout ohne validiertes PROD-Manifest.
