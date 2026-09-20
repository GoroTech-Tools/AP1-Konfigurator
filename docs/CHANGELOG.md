# Changelog

## v1.0.33 - 2026-09-20

### Release und Dokumentation

- Release Notes enthalten direkte GitHub-Downloadlinks sowie automatisierte Informationen zu Buildzeit, Commit, ZIP-Größe und SHA-256-Prüfsumme.
- GitHub Actions verwendet die erzeugten Release Notes als sichtbaren Release-Body statt ausschließlich automatischer Changelog-Notizen.

## v1.0.16 - 2026-06-28

### GUI und Laufstatus

- Die EXE-GUI zeigt beim Start der AP1-Konfiguration jetzt einen sichtbaren Fortschrittsbalken.
- Der Ablaufstatus wird laufend aktualisiert; bei erfolgreichem Abschluss wird der Balken grün und zeigt `Fertig`.
- Bei Fehlern wird ein klarer Fehlerzustand mit roter Abschlussmarkierung ausgegeben.

### Dokumentation (v1.0.16)

- README sowie Anwender-/Technikdokumentation auf den aktuellen GUI- und Release-Stand synchronisiert.

## v1.0.14 - 2026-06-28

### Dokumentation und Release

- README, Anwender-, Technik- und Release-Prozess-Dokumentation auf `v1.0.14` aktualisiert.
- Neue Release-Notes-Vorlage `release/RELEASE_NOTES_TEMPLATE.md` ergänzt.
- Release-Notizen orientieren sich künftig am strukturierten Muster von `release/RELEASE_NOTES_v1.0.10.md`.
- Obsolete Weiterleitungsdatei `docs/Liesmich.txt` entfernt.

### Versionsstand

- Versionsquelle in `src/build_info.py` auf `1.0.14` angehoben.

## v1.0.11 - 2026-06-28

### Dokumentation (v1.0.11)

- Die kanonische Projektdokumentation wurde analog zu den Schwesterprojekten nach `docs/` überführt.
- Neue Standarddateien angelegt: `docs/DOKUMENTATION_ANWENDER.md`, `docs/DOKUMENTATION_TECHNIK.md`, `docs/KURZDOKUMENTATION.txt`, `docs/DOKUMENTATION_CHECKLISTE.md`.
- Bestehende Einstiegspunkte `DOCUMENTATION.md` und `docs/Liesmich.txt` in Weiterleitungen auf die kanonischen Dateien umgewandelt.

### Release

- Aktueller lokaler und GitHub-Release auf `v1.0.11` angehoben.
- Release-Unterlagen für das neue Paket ergänzt.

## v1.0.10 - 2026-06-28

### Behoben

- `Write-SafeOutput` unterstützt wieder `-ForegroundColor` und gibt Statusmeldungen sichtbar aus.
- Die dokumentierten Schalter `-Quiet` und `-RegistryOnly` sind nun im Hauptskript korrekt nutzbar.
- Die Hauptfunktion `Start-AP1Konfiguration` verarbeitet Parameter jetzt explizit statt indirekt und fehleranfällig.
- Die irreführende Meldung zum Registry-only-Modus bei erfolgreicher COM-Erkennung wurde korrigiert.
- Doppelte Fehlerausgabe beim fehlenden Excel-Pfad wurde entfernt.

### Bereinigt

- Doppelte Initialisierung für Encoding und Modulimporte im Hauptskript entfernt.
- README, technische Dokumentation und Kurzinfo auf den tatsächlichen Funktionsumfang aktualisiert.
- Release-Dokumentation für `v1.0.10` ergänzt.

## v1.0.9

- Vorheriger veröffentlichter Stand im Repository.

## v1.0.8

- Vorheriger veröffentlichter Stand im Repository.

## v1.0.7

- Vorheriger veröffentlichter Stand im Repository; siehe `release/RELEASE_NOTES_v1.0.7.md`.
