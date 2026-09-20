# AP1-Konfigurator vX.Y.Z

## Release-Übersicht

- Erstellungszeitpunkt: {{BUILD_DATE}}
- Quellstand: `{{COMMIT_SHA}}`
- Auslieferungsformat: Portable Onefile-EXE für Windows

## Download

- [Release-Seite vX.Y.Z](https://github.com/GoroTech-Tools/AP1-Konfigurator/releases/tag/vX.Y.Z)
- [ZIP direkt herunterladen](https://github.com/GoroTech-Tools/AP1-Konfigurator/releases/download/vX.Y.Z/AP1-Konfigurator-vX.Y.Z.zip)

## Lieferumfang

- `AP1-Konfigurator.exe`
- `data/` mit Eingabedaten, Vorlagen und Laufzeitkonfiguration
- `docs/` mit Anwender-, Technik- und Checklisten-Dokumentation
- `README.md`

## Highlights

- Die GitHub-Release-Seite zeigt diese strukturierten Release Notes direkt an.
- Direkte ZIP-Downloadlinks, Build-Metadaten und Integritätsinformationen sind enthalten.
- Die portable Onefile-EXE enthält die eingebetteten PowerShell-Ressourcen und wird mit `data/`, `docs/` und `README.md` ausgeliefert.

## Änderungen im Detail

### Release-Pipeline

- Version, Build-Datum und Dokumentationsstände werden beim Build synchronisiert.
- Ältere lokale Release-Artefakte werden in `release/_Archiv/` verschoben.

### Paketierung

- Das ZIP ist flach aufgebaut und enthält die EXE, `data/`, `docs/` und `README.md`.
- Die ZIP-Prüfsumme dient der optionalen Integritätskontrolle nach dem Download.

## Validierung

- PyInstaller-Onefile-Build erfolgreich erzeugt.
- Versionierte ZIP-Datei nach der Paketierung vorhanden.
- ZIP-Größe: {{ZIP_SIZE_MIB}} MiB
- ZIP-SHA-256: `{{ZIP_SHA256}}`

## Artefakte und GitHub-Hinweis

- Git-Tag: `vX.Y.Z`
- GitHub-Release: `vX.Y.Z`
- Lokales Release-ZIP: `release/{{ZIP_FILENAME}}`
- Commit zum Zeitpunkt der lokalen Paketierung: `{{COMMIT_SHA}}`
