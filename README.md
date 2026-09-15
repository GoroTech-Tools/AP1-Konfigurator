
# AP1-Konfigurator

Aktueller Stand: **v1.0.30** · Letzte Aktualisierung: **12. September 2026**

Der `AP1-Konfigurator` automatisiert die Vorbereitung von Prüfungsrechnern für die Abschlussprüfung Teil 1 (AP1) auf Windows-Systemen. Das Projekt setzt Office-Defaults, Standardordner, Proxy-Optionen, Kandidatenstruktur und Nuera-Dateien in einer nachvollziehbaren Reihenfolge auf, damit die Einrichtung ohne manuelle Fehlerquellen wiederholbar bleibt.

## Zweck des Projekts

Das Tool ist für die typische AP1-Vorbereitung gedacht:

- Standard-Speicherpfade für Word und Excel auf den Desktop des aktuellen Benutzers setzen
- persönliche Vorlagen und Office-Optionen im Registry-/COM-Workflow absichern
- Kandidatenordner aus einer Excel-Liste oder einer CSV-Datei erzeugen
- benötigte Nuera-Dateien automatisch bereitstellen
- Proxy-Konfigurationen für die Prüfungsumgebung sauber steuern
- Laufzeitzustände und Logs nachvollziehbar dokumentieren

## Was der Lauf konkret erledigt

- Office-Registrywerte für Word/Excel im modernen 16.0-Pfad aufsetzen
- Optionales COM-Handling mit sauberem Fallback in den Registry-only-Betrieb
- Standard-Speicherpfade und persönliche Word-Vorlagen definieren
- Desktop- und Kandidatenstruktur für AP1-Teilnehmer erzeugen
- Nuera-Dateien prüfen, kopieren und auf den Desktop bringen
- Proxy-Status steuern (`On`, `Off`, `Skip`)
- Protokolle nach `data/4. Logs` schreiben

## Schnellstart

Interaktiv per Batchdatei:

```powershell
.\src\AP1-Konfigurator.bat
```

Direkt per PowerShell:

```powershell
.\src\AP1-Konfigurator.ps1
```

Typische Aufrufe:

```powershell
.\src\AP1-Konfigurator.ps1 -Proxy Off
.\src\AP1-Konfigurator.ps1 -Proxy On -Quiet
.\src\AP1-Konfigurator.ps1 -RegistryOnly -CsvFallbackPath .\data\1. Anpassen\AP1-TN.csv
.\src\AP1-Konfigurator.ps1 -UseCom -ExcelListPath .\data\1. Anpassen\AP1-TN.xlsx
```

## Verfügbare Parameter

| Parameter | Typ | Standard | Zweck |
| --- | --- | --- | --- |
| `-Proxy` | `On`, `Off`, `Skip` | `Skip` | Proxy-Status steuern |
| `-ProxyServer` | `String` | `192.168.0.1:8080` | Server bei `-Proxy On` |
| `-ProxyBypass` | `String` | Office-/Microsoft-Bypassliste | Ausnahmen für den Proxy |
| `-ExcelListPath` | `String` | automatisch `data/1. Anpassen\AP1-TN.xlsx` | Teilnehmerliste für Kandidatenordner |
| `-CsvFallbackPath` | `String` | leer | CSV-Alternative, falls Excel/COM nicht verfügbar ist |
| `-MaxRows` | `Int` | `500` | maximale Zeilenanzahl beim Einlesen |
| `-Quiet` | `Switch` | aus | Proxy-Rückfrage unterdrücken |
| `-RegistryOnly` | `Switch` | aus | Betrieb ohne COM erzwingen |
| `-UseCom` | `Switch` | aus | COM-basierte Office-Änderungen explizit aktivieren |

## Projektstruktur

```text
AP1-Konfigurator/
├── src/
│   ├── AP1-Konfigurator.ps1
│   ├── AP1-Konfigurator.bat
│   ├── AP1-Konfigurator.spec
│   ├── build.ps1
│   ├── post_build.py
│   ├── setup.ps1
│   ├── publish_release.ps1
│   ├── requirements.txt
│   └── Skript-Module/
├── data/
│   ├── 1. Anpassen/
│   ├── 2. Bei Bedarf anpassen/
│   ├── 3. Nuera-Dateien/
│   └── 4. Logs/
├── docs/
├── release/
├── dist/
├── README.md
├── .gitignore
└── .github/
```

## Datenfluss und wichtige Hinweise

- `data/1. Anpassen` enthält die konfigurierbaren Eingabedaten, z. B. die Excel-Liste mit AP1-Teilnehmern.
- `data/2. Bei Bedarf anpassen` ist ein optionaler Bereich für manuelle Anpassungen; temporäre Hilfsordner werden nach Verwendung bereinigt, damit keine leeren Nebenordner im Projekt verbleiben.
- `data/3. Nuera-Dateien` nimmt die aktuell benötigten Dateien auf und hält sie für die AP1-Vorbereitung bereit.
- `data/4. Logs` sammelt die Laufzeitprotokolle, damit Fehler und Verläufe nachvollziehbar bleiben.
- Die EXE-Version nutzt eine eigene lokale Laufzeitkopie unter `%LOCALAPPDATA%\AP1-Konfigurator\vX.Y.Z` und besitzt zusätzlich einen aktuellen Linkpfad unter `%LOCALAPPDATA%\AP1-Konfigurator\current`.
- Alte `%LOCALAPPDATA%\AP1-Konfigurator\v*`-Ordner werden beim Start bereinigt, damit kein veralteter Versionsballast zurückbleibt.
- Wenn Word oder Excel nicht über COM erreichbar sind, fällt das Projekt automatisch auf den stabilen Registry-only-Fallback zurück.
- Die GUI zeigt Laufzeitstatus und beendet den Ablauf mit einer klaren Abschlussanzeige; bei erfolgreichem Lauf erscheint der Status als **Fertig**.

## EXE-Release

Das Release ist bewusst schlank gebaut und enthält in der Regel:

- `AP1-Konfigurator.exe`
- `data/`
- `docs/`
- `README.md`

Die PowerShell-Ressourcen und Module werden nicht separat mitgeliefert, sondern in die EXE integriert und beim Start in den lokalen Nutzerbereich kopiert.

## Build- und Release-Prozess

Das Build-System ist in `src/build.ps1` und `src/post_build.py` organisiert. Es erzeugt die versionierte Release-Struktur und das ZIP-Artefakt für GitHub.

Für lokale Builds und GitHub Actions wird Python 3.13 verwendet. Die lokale Umgebung wird mit `src/setup.ps1` als `.venv` angelegt.

- `src/build.ps1` prüft die Version, erzeugt die EXE und aktualisiert die Projekt-Doku
- `src/post_build.py` schafft den finalen Release-Ordner und das komprimierte Asset
- `release/` enthält die veröffentlichungsrelevanten Artefakte und Release-Notizen

## Dokumentation

- Anwender-Dokumentation: [`docs/DOKUMENTATION_ANWENDER.md`](./docs/DOKUMENTATION_ANWENDER.md)
- Technik-Dokumentation: [`docs/DOKUMENTATION_TECHNIK.md`](./docs/DOKUMENTATION_TECHNIK.md)
- Checkliste: [`docs/DOKUMENTATION_CHECKLISTE.md`](./docs/DOKUMENTATION_CHECKLISTE.md)
- Änderungsübersicht: [`docs/CHANGELOG.md`](./docs/CHANGELOG.md)
- Release-Hinweise: [`release/RELEASE_NOTES_v1.0.30.md`](./release/RELEASE_NOTES_v1.0.30.md)

## Hinweise zur Nutzung

- Für die normale Projektverwendung ist die EXE-Variante die bevorzugte Option.
- Für Tests in kontrollierter Umgebung kann `-RegistryOnly` genutzt werden, um Office-COM gezielt auszuschließen.
- `-Quiet` eignet sich für szenarien ohne Benutzerinteraktion, etwa bei automatisierten Setup-Läufen.
- Bei der Erstellung von Kandidatenordnern sollte die Excel-Datei mit den richtigen Spalten und Bezeichnungen vorliegen; CSV-Fallbacks eignen sich nur als Ersatz für testbare Standardfälle.

## Lizenz und Projektstatus

Das Projekt wird für die interne AP1-Vorbereitung in einer Windows-Office-Umgebung entwickelt. Der aktuelle Release-Stand ist `v1.0.30`. Für neue Releases und Änderungen wird die Version zentral über die Build-Informationen gepflegt; die Dokumentation wird im gleichen Schritt synchronisiert.
