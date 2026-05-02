# Microsoft 365 Admin Deployment

Tool: Werkhaus Word Bild Tool
Produktive URL: https://tool2.wh-sv.de
Zu verwendendes Manifest: manifests/manifest.prod.xml

Diese Anleitung ist fuer die zentrale Bereitstellung im Microsoft 365 Admin Center. DEV-Manifeste niemals produktiv verteilen.

## Vorbedingungen
- PROD-Manifest ist validiert.
- PROD-URL `https://tool2.wh-sv.de` ist per HTTPS erreichbar.
- Build und Lint sind gruen.
- PROD-Manifest enthaelt kein `TEST`.
- PROD-Manifest zeigt nicht auf `localhost`.
- Tenant und Benutzer unterstuetzen Centralized Deployment.
- Benutzer sind fuer Microsoft 365 Apps angemeldet und gehoeren zur geplanten Testgruppe.

## Bereitstellung mit kleiner Testgruppe
1. Microsoft 365 Admin Center oeffnen.
2. `Settings > Integrated apps` oeffnen.
3. `Add-ins` bzw. `Upload custom apps` waehlen.
4. Ausschliesslich `manifests/manifest.prod.xml` hochladen.
5. Zuerst eine kleine Testgruppe oder `Just me` auswaehlen.
6. Deployment abschliessen.
7. Word komplett neu starten.
8. Im Ribbon pruefen, dass kein TEST-Add-in erscheint.
9. Erst nach erfolgreichem Test breiter ausrollen.

## Word-Funktionstest
- Ribbon-Name ist die PROD-Bezeichnung ohne `TEST`.
- Taskpane oeffnet ohne weisse Flaeche.
- Kernfunktion mit kleiner Testdatei pruefen.
- Fehlerfall pruefen.

## Rollback
- Wenn nur Web-Code defekt ist: Azure auf vorherigen Release Tag zuruecksetzen.
- Wenn Manifest-Metadaten defekt sind: vorheriges `manifest.prod.xml` erneut im Admin Center hochladen.
- Testgruppe erneut pruefen, bevor der breite Rollout aktiv bleibt.

## Sicherheitsregel
- DEV-Manifeste `manifests/manifest.dev.xml` niemals im Microsoft 365 Admin Center produktiv verteilen.
