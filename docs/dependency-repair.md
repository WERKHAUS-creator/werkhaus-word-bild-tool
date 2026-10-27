# Abhängigkeitsreparatur: Ursache und Absicherung

## Nachgewiesene Ursache

Untersuchter `main`-Stand: `44f5c3b1ee43ff1b8ce458d51773a76e0182fd47`.
Anwendungsversion: `1.0.0.10`.

[PR #2](https://github.com/WERKHAUS-creator/werkhaus-word-bild-tool/pull/2)
wurde am 08.05.2026 um 16:04:48 UTC in **develop** zusammengeführt.
Der Merge-Commit ist `81fb58b47e3051f44617975d6e064faa5c9a0122`;
die eigentliche Änderung stammt aus `ad0aa0964b812f72ffc51418f12e514dbb68d5f1`.
Sie entfernte das Löschen von `package-lock.json` und führte über
`clean:deps` / `reinstall:deps` zu `npm ci`.

`develop` steht bei der Untersuchung weiterhin auf diesem Merge-Commit.
Der gemeinsame Vorfahr mit `main` ist
`2d717ae1af21a316fbc4b2d5811b54c83ff6d3e5`. GitHub vergleicht beide Stände als
divergiert: `main` enthält 46 zusätzliche Commits, aber nicht die vier nur in
`develop` enthaltenen Commits. `git log main..origin/develop` bestätigt das.
`git log --all -S 'rm -rf node_modules package-lock.json' -- package.json`
zeigt die ursprüngliche Einführung (`ebc3305c`) und die Korrektur (`ad0aa096`).

Damit belegt die aktuelle Historie keine spätere Rücknahme der Korrektur auf
`main`. Die Korrektur wurde dort nie übernommen. Aus der Git-Historie allein
lässt sich nicht bestimmen, weshalb die Übernahme organisatorisch ausblieb.
Ein vollständiger Merge von `develop` würde auch alte Manifest-, Icon- und
Build-Änderungen übernehmen; deshalb wird nur der Reparaturablauf gezielt auf
dem aktuellen `main` korrigiert.

## Umsetzung

- `repair:deps` führt direkt `npm ci` aus. npm entfernt `node_modules` selbst;
  zusätzliche Löschbefehle oder ein Fallback auf `npm install` entfallen.
- Fehlende bzw. widersprüchliche Lockdateien und Installationsfehler führen zum
  Abbruch. Die Lockdatei wird nicht neu generiert.
- Es werden keine Paketversionen, Lockdatei, App-Quellen, Icons oder Manifeste
  geändert; die Anwendung bleibt bei `1.0.0.10`.
- Vier Integrationstests nutzen die tatsächliche Projektkonfiguration mit
  temporären Projekten, eigenen npm-Caches und lokalen Paketarchiven. Das
  installierte Testpaket und seine CLI-Version werden tatsächlich geprüft.
- Ein separater Workflow führt Tests, die echte Projekt-Reparatur,
  Manifest-Konsistenzprüfung und Entwicklungs-/Produktionsbuilds auf Linux und
  macOS aus. Ein abschließender Git-Diff erkennt Änderungen an den Paketdateien.
- Azure nutzt bereits `npm ci`; der Workflow erhält zusätzlich die Tests und
  eine Prüfung der Paketdateien. Entwurfs-PRs veröffentlichen keine Vorschau.
  Bestehende Push-/manuelle Produktionsauslöser werden nicht ausgeführt oder
  umkonfiguriert. Der Reparatur-PR bleibt ein Entwurf und wird nicht gemergt.

## Technische Grundlage

[npm-ci-Dokumentation](https://docs.npmjs.com/cli/v10/commands/npm-ci/):
`npm ci` setzt eine passende Lockdatei voraus, ersetzt vorhandene
`node_modules` und schreibt weder `package.json` noch die Lockdatei neu.
Für gleiche Entwicklungsumgebungen außerdem dieselbe Node-/npm-Version nutzen.
Native Module müssen je Betriebssystem installiert werden; ein zwischen Macs
kopierter `node_modules`-Ordner ist kein Ersatz für diesen Installationsschritt.

## Lokale Prüfung am 27.10.2026

Isolierter Checkout; Node `24.22.1`, npm `11.9.0`. Die CI verwendet wie bisher
Node 20; deren Betriebssystemprüfungen sind gesondert zu beurteilen.

| Prüfung | Ergebnis |
| --- | --- |
| `npm test` | 4/4 erfolgreich, ohne Projektinstallation und offline |
| Gegenprobe mit altem Reparaturbefehl in Testfixtures | 4/4 schlagen erwartungsgemäß fehl |
| `npm run lint` und ESLint für die neue Testdatei | erfolgreich |
| `npm run check:manifests` | erfolgreich |
| `npm run repair:bin` | erfolgreich, keine Rechteanpassung erforderlich |
| `npm run build:dev`, `npm run build`, `npm run build:production` | erfolgreich |
| `npm run validate:prod` | Microsoft meldet das Manifest als gültig |
| `npm run validate:dev` | nach vorübergehender Nichterreichbarkeit meldet Microsoft das Manifest beim erneuten Aufruf als gültig |
| `npm run repair:deps -- --ignore-scripts …` im echten Checkout | erfolgreich; native Installationsskripte nur für diese Umgebungsprüfung deaktiviert |
| Lockdatei vor/nach Reparatur und Builds | bytegleich, SHA-256 siehe unten |

SHA-256 von `package-lock.json`:
`eeab231a2a5cd661d0311f47b15e6c878cff80b5db2d7ccd81cae9cd0e2d6633`.

Die unveränderte Standardinstallation mit nativen Installationsskripten scheitert
in dieser Ausführungsumgebung an `keytar`/`node-gyp` mit `EINVAL` bei
`lchown`/`fchown`. Daher sind die lokalen Builds mit per `npm ci --ignore-scripts`
installierten Paketen geprüft. Dies ist keine Änderung des Reparaturbefehls und
keine Empfehlung für die Mac-Installation. Der neue CI-Workflow prüft den
normalen Reparaturbefehl einschließlich nativer Installationsskripte.

Von den bestehenden npm-Befehlen wurde nur `repair:deps` geändert; `test` kam
hinzu. Office-Start/-Stop, Entwicklungsserver und Watch-Modus wurden nicht
gestartet. Ein Word-/Office-Host-Test ist hier nicht möglich.
Es wurden weder bestehende Mac-Installationen verändert
noch eine Bereitstellung oder ein Merge ausgelöst.
