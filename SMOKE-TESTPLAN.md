# WERKHAUS Word-Bild-Tool - Smoke-Testplan

## Ziel

Dieser Smoke-Testplan prueft den aktuell vorhandenen Projektstand nach der Tooling-Stabilisierung. Er beschreibt nur Funktionen, die im Code derzeit aktiv verdrahtet sind.

Kontaktbogen, Mehrfachbild-Einfuegen und Tabellenlayout sind aktuell vorbereitet, aber noch nicht als bedienbare Word-Funktion umgesetzt. Diese Funktionen werden hier deshalb nicht als bestanden/nicht bestanden getestet.

## Aktueller Codebezug

- Taskpane-UI und Listenlogik: [src/taskpane/taskpane.ts](src/taskpane/taskpane.ts)
- Taskpane-HTML und Bedienelemente: [src/taskpane/taskpane.html](src/taskpane/taskpane.html)
- Datei-, Ordner- und Drop-Import: [src/taskpane/importImages.ts](src/taskpane/importImages.ts)
- Lokale Metadaten: [src/taskpane/persistence.ts](src/taskpane/persistence.ts)
- Bild- und Layouttypen: [src/taskpane/types.ts](src/taskpane/types.ts)
- Caption-Regeln fuer Einzelbild-Ausgabe: [src/taskpane/captions.ts](src/taskpane/captions.ts)
- Einzelbild-Einfuegen in Word: [src/taskpane/wordInsert.ts](src/taskpane/wordInsert.ts)
- Vorbereitete Layoutplanung, derzeit nicht in UI/Word-Ausgabe verdrahtet: [src/taskpane/layoutEngine.ts](src/taskpane/layoutEngine.ts), [src/taskpane/pageCalculator.ts](src/taskpane/pageCalculator.ts), [src/taskpane/constants.ts](src/taskpane/constants.ts)
- Manifest/Ribbon-Konfiguration: [manifest.xml](manifest.xml)

## Nicht mehr gueltige Alt-Verweise

Die frueher im Testplan genannten Dateien `src/taskpane/wordTable.ts` und `src/taskpane/layout.ts` existieren im aktuellen Projektstand nicht. Kontaktbogen-Tests muessen neu ergaenzt werden, sobald die entsprechende Word-Ausgabeschicht implementiert ist.

## Vorab-Checks

1. `npm run validate`
2. `./node_modules/.bin/tsc --noEmit`
3. `npm run lint`

Erwartung:

- Manifest ist gueltig.
- TypeScript-Konfiguration laeuft ohne Konfigurationsfehler.
- Lint ist technisch ausfuehrbar und meldet keine offenen Format-/Globalkonfigurationsfehler.

## Testumgebung

- Word Desktop oeffnen.
- Add-in ueber `npm run start` oder die vorhandene VS-Code-Office-Add-in-Konfiguration starten.
- Neues leeres Word-Dokument verwenden.
- 6 bis 12 Testbilder bereithalten:
  - Hochformat
  - Querformat
  - unterschiedliche Dateigroessen
  - mindestens 1 Bild mit langem Dateinamen
  - mindestens 1 echte Dublette derselben Datei
  - optional 1 JPEG mit EXIF-Daten
- Optional mindestens 2 Nicht-Bilddateien fuer negative Importtests bereithalten.

## Test 1 - Dateiimport

### Schritte

1. Add-in oeffnen.
2. `Dateien waehlen` anklicken.
3. 3 bis 5 einzelne Bilddateien auswaehlen.

### Erwartung

- Alle unterstuetzten Bilddateien erscheinen in der Bildliste.
- Vorschaubilder werden angezeigt.
- Dateiname, Dateigroesse und Dateityp erscheinen, wenn Bildinfos eingeblendet sind.
- Status meldet die Anzahl der hinzugefuegten Bilder.

## Test 2 - Ordnerimport

### Schritte

1. Add-in oeffnen.
2. `Ordner waehlen` anklicken.
3. Einen Ordner mit mehreren Bildern auswaehlen.

### Erwartung

- Unterstuetzte Bilddateien werden uebernommen.
- Nicht unterstuetzte Dateien werden ignoriert.
- Die Bildliste verwendet dieselbe Darstellung wie beim Dateiimport.
- Statusmeldung nennt importierte und ignorierte Dateien nachvollziehbar.

## Test 3 - Drag-and-Drop-Import

### Schritte

1. Mehrere gueltige Bilddateien in die Drop-Zone ziehen.
2. Danach eine Mischung aus Bilddateien und Nicht-Bilddateien in die Drop-Zone ziehen.
3. Danach nur Nicht-Bilddateien in die Drop-Zone ziehen.

### Erwartung

- Gueltige Bilder werden uebernommen.
- Nicht-Bilddateien werden ignoriert.
- Bei ausschliesslich ungueltigen Dateien bleibt die Bildliste unveraendert.
- Statusmeldung nennt die ignorierten Dateien beziehungsweise das Fehlen gueltiger Bilddateien.

## Test 4 - Dublettenerkennung

### Schritte

1. Einige Bilder importieren.
2. Dieselben Dateien erneut importieren.
3. Eine echte Dublette aus einem Ordnerimport testen.

### Erwartung

- Bereits importierte Bilder werden nicht erneut angelegt.
- Status meldet uebersprungene Dubletten.
- Vorhandene Captions und Positionen bleiben erhalten.

## Test 5 - Bildliste und Reihenfolge

### Schritte

1. Mindestens 5 Bilder importieren.
2. Reihenfolge per Drag-and-Drop veraendern.
3. Position ueber das Nummernfeld veraendern.
4. Bildinfos und Beschriftungen ein- und ausblenden.

### Erwartung

- Reihenfolge in der Liste aendert sich sichtbar.
- Positionsnummern werden neu durchnummeriert.
- Bildinfos und Beschriftungsbereiche lassen sich ein- und ausblenden.
- Die UI bleibt bedienbar.

## Test 6 - Beschriftungsmetadaten in der UI

### Schritte

1. Bei mehreren Bildern Beschriftungstexte eintragen.
2. Bei mindestens einem Bild die Option `Beschriftung ins Word-Dokument uebernehmen` deaktivieren.
3. Add-in neu laden, soweit im Host sinnvoll moeglich.
4. Dieselben Bilder erneut importieren.

### Erwartung

- Beschriftungstexte werden anhand des Bild-Hashes aus `localStorage` wiederhergestellt.
- Position und Caption-Aktivierung werden wiederhergestellt, soweit die Dateien identisch sind.
- Keine doppelten Bildobjekte entstehen.

## Test 7 - Einzelbild ohne Beschriftung einfuegen

### Schritte

1. Ein Bild importieren.
2. Cursor in ein leeres Word-Dokument setzen.
3. Groesse ueber `Lange Seite` einstellen.
4. Beim Bild `Einfuegen` anklicken.

### Erwartung

- Genau ein Inline-Bild wird an der aktuellen Auswahl eingefuegt.
- Es wird kein Beschriftungsabsatz eingefuegt.
- Die Bildgroesse entspricht plausibel dem eingestellten Wert.
- Der Cursor landet nach dem eingefuegten Bild.

## Test 8 - Einzelbild mit Beschriftung einfuegen

### Schritte

1. Ein Bild importieren.
2. Beschriftungstext eintragen.
3. Cursor in Word setzen.
4. `Einfuegen mit Beschriftung` anklicken.

### Erwartung

- Genau ein Inline-Bild wird eingefuegt.
- Darunter wird ein normaler Word-Absatz mit Text nach dem Muster `Abbildung N: ...` eingefuegt.
- Bei leerem Beschriftungstext wird der Dateiname verwendet.
- Die Nummer `N` entspricht der Position des Bildes in der Seitenleiste.

## Test 9 - Groessenskalierung

### Schritte

1. Hochformat- und Querformatbilder importieren.
2. Mehrere Werte am Slider `Lange Seite` testen.
3. Jeweils ein Einzelbild einfuegen.

### Erwartung

- Die laengste Bildseite wird innerhalb der konfigurierten Nutzflaeche skaliert.
- Seitenverhaeltnisse bleiben erhalten.
- Keine sichtbare Verzerrung.

## Test 10 - Liste leeren

### Schritte

1. Mehrere Bilder importieren.
2. `Liste leeren` anklicken.
3. Danach erneut Bilder importieren.

### Erwartung

- Die aktuelle In-Memory-Bildliste wird geleert.
- Importfelder werden zurueckgesetzt.
- Metadaten in `localStorage` werden dadurch nicht automatisch geloescht.
- Neuer Import funktioniert direkt wieder.

## Noch nicht testbare vorbereitete Bereiche

Die folgenden Bereiche sind als Code vorhanden, aber aktuell nicht in die Taskpane-UI oder Word-Ausgabe eingebunden:

- Mehrfachbild-Einfuegen
- Kontaktbogen
- Word-Tabellenlayout fuer Bilder
- Seitenumbrueche zwischen Kontaktbogen-Seiten
- Kontaktbogen-Captions in Tabellenzellen

Fuer diese Bereiche gilt im aktuellen Smoke-Test nur:

- TypeScript kompiliert die Module ohne Fehler.
- Lint akzeptiert die Module.
- Keine UI-Schaltflaeche verspricht diese Funktion als aktiv.

## Ergebnisprotokoll

Fuer jeden Test festhalten:

- Bestanden / Nicht bestanden
- beobachtetes Verhalten
- Word-Version und Plattform
- Browser-/WebView-Hinweis, falls sichtbar
- Anzahl und Typ der verwendeten Bilder
- ob Caption aktiv war
- ob das Problem reproduzierbar ist

## Fehlerklassifikation

### Kritisch

- Add-in startet nicht.
- Import scheitert vollstaendig.
- Word-Einfuegen scheitert vollstaendig.
- Word friert ein oder Dokument wird unbenutzbar.

### Mittel

- Dublettenerkennung falsch.
- Reihenfolge oder Positionen inkonsistent.
- Beschriftungstext falsch oder nicht wiederherstellbar.
- Bildgroesse deutlich falsch.

### Klein

- Statusmeldung ungenau.
- Optische Unsauberkeit in der Taskpane.
- Einzelne EXIF-Information fehlt, obwohl Import und Einfuegen funktionieren.
