# Word Add-in Cleanup (macOS)

Diese Anleitung entfernt veraltete Word-Add-in-Caches und Sideload-Reste, damit nur die aktuellen DEV-Manifeste geladen werden.

## 1) Word komplett beenden

1. In Word `Word > Beenden` waehlen.
2. Oder `Cmd + Q` verwenden.
3. Sicherstellen, dass kein `Microsoft Word` Prozess mehr offen ist.

## 2) Alte Add-ins in Word entfernen (falls sichtbar)

1. Word oeffnen.
2. `Einfügen > Meine Add-Ins` oeffnen.
3. Alte/doppelte lokale Add-ins entfernen.
4. Word wieder komplett beenden.

## 3) Word-Add-in-Cache loeschen

Nur nach komplett beendetem Word ausfuehren:

```bash
rm -rf ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches
rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef
```

Hinweis: Diese Befehle loeschen nur Add-in-/Cache-Daten, keine Word-Dokumente.

## 4) DEV-Server neu starten

### Bild Tool DEV

```bash
cd "/Users/hubenschmid/OneDrive - WERKHAUS/96. WERKHAUS Tools/Werkhaus-Word-Bild/WERKHAUS-Word-Bild-Tool Develop"
npm run dev-server
```

### Excel Tool DEV

```bash
cd "/Users/hubenschmid/OneDrive - WERKHAUS/96. WERKHAUS Tools/Werkhaus-Word-Excel/Werkhaus Word Excel Develop"
npm run dev-server
```

## 5) Nur die neuen DEV-Manifeste laden

Nur diese beiden Dateien in Word sideloaden:

- `/Users/hubenschmid/OneDrive - WERKHAUS/96. WERKHAUS Tools/Werkhaus-Word-Bild/WERKHAUS-Word-Bild-Tool Develop/manifests/manifest.dev.xml`
- `/Users/hubenschmid/OneDrive - WERKHAUS/96. WERKHAUS Tools/Werkhaus-Word-Excel/Werkhaus Word Excel Develop/manifests/manifest.dev.xml`

Erwartung im Ribbon:

- `WERKHAUS TEST Word & Bild` (blau markiertes DEV-Icon)
- `WERKHAUS TEST Word & Excel` (rot markiertes DEV-Icon)

## 6) Warum doppelte Piktogramme entstehen

Doppelte Ribbon-Icons entstehen fast immer durch:

- parallel geladene alte und neue Manifeste
- gecachte WEF-Eintraege in Word
- noch laufende alte DEV-Server mit alten Assets

Wenn wieder Doppelungen erscheinen, erneut mit dieser Anleitung bereinigen und danach nur die beiden DEV-Manifeste neu laden.
