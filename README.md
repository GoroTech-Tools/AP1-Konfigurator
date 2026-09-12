
# AP1-Konfigurator

Aktueller Stand: **v1.0.30** · Letzte Aktualisierung: **12. September 2026**

Der `AP1-Konfigurator` automatisiert die Einrichtung von Prüfungsrechnern für die Abschlussprüfung Teil 1 (AP1). Das PowerShell-Skript richtet Office, Explorer, Schnellzugriff, Proxy-Einstellungen, Kandidatenordner und die Nuera-Dateien in einer reproduzierbaren Reihenfolge ein.

## Was das Skript erledigt

- initialisiert Word und Excel robust über COM mit Registry-Fallback
- zeigt in der GUI einen Fortschrittsbalken mit Live-Status und Abschlusszustand
- setzt Standard-Speicherpfade auf den Desktop des aktuellen Benutzers
- übernimmt `Normal.dotm`, `Mappe.xltx` und Office-Schnellzugriff
- erzeugt Kandidatenordner aus `AP1-TN.xlsx` oder optional aus einer CSV-Datei
- lädt die neuesten Nuera-Dateien und kopiert sie auf den Desktop
- setzt Taskleisten- und optionale Proxy-Einstellungen
- schreibt Laufzeitprotokolle nach `data/4. Logs`

## Schnellstart

Interaktiv per Batch:

```powershell
.\src\AP1-Konfigurator.bat
```

Direkt per PowerShell:

```powershell
.\src\AP1-Konfigurator.ps1
```

## Verfügbare Parameter

| Parameter | Typ | Standard | Zweck |
| --- | --- | --- | --- |
| `-Proxy` | `On`, `Off`, `Skip` | `Skip` | Steuert die Proxy-Konfiguration |
| `-ProxyServer` | `String` | `192.168.0.1:8080` | Proxy-Server bei `-Proxy On` |
| `-ProxyBypass` | `String` | Office-/Microsoft-Bypassliste | Ausnahmeliste für den Proxy |
| `-ExcelListPath` | `String` | automatisch `data/1. Anpassen\AP1-TN.xlsx` | Teilnehmerliste für die Ordnererzeugung |
| `-CsvFallbackPath` | `String` | leer | CSV-Fallback, wenn Excel/COM nicht verfügbar ist |
| `-MaxRows` | `Int` | `500` | Obergrenze für gelesene Teilnehmerzeilen |
| `-Quiet` | `Switch` | aus | Unterdrückt die Proxy-Rückfrage |
| `-RegistryOnly` | `Switch` | aus | Erzwingt den Betrieb ohne COM |

## Typische Aufrufe

```powershell
.\src\AP1-Konfigurator.ps1 -Proxy Off
.\src\AP1-Konfigurator.ps1 -Proxy On -Quiet
.\src\AP1-Konfigurator.ps1 -RegistryOnly -CsvFallbackPath .\data\1. Anpassen\AP1-TN.csv
```

## Ordnerstruktur

```text
AP1-Konfigurator/
├── src/
│   ├── AP1-Konfigurator.ps1
│   ├── AP1-Konfigurator.bat
│   ├── Proxy-Deaktivieren.bat
│   ├── setup.ps1
│   ├── build.ps1
│   ├── publish_release.ps1
│   └── requirements.txt
│   └── Skript-Module/
├── data/
│   ├── 1. Anpassen/
│   ├── 2. Bei Bedarf anpassen/
│   ├── 3. Nuera-Dateien/
│   └── 4. Logs/
├── docs/
├── dist/
├── release/
└── _Archive/
```

## Wichtige Hinweise

- Die Nuera-Dateien werden im aktuellen Stand **automatisch** geprüft und bereitgestellt.
- Wird Word oder Excel per COM nicht verfügbar, wechselt das Skript automatisch in den Registry-Fallback.
- Für einen stillen Proxy-Lauf ohne Rückfrage sollte `-Quiet` mitgegeben werden.
- Die Laufzeit hängt stark davon ab, ob Office erstmals gestartet werden muss.
- Releases sollen grundsätzlich über die **EXE-Variante** erfolgen (`AP1-Konfigurator`).
- Die EXE entpackt bzw. synchronisiert ihre eingebetteten Laufzeitdateien beim Start nach `%LOCALAPPDATA%\AP1-Konfigurator\vX.Y.Z` und pflegt zusätzlich `%LOCALAPPDATA%\AP1-Konfigurator\current` als aktuelle Arbeitskopie.
- Ältere `%LOCALAPPDATA%\AP1-Konfigurator\v*`-Ordner werden beim Start automatisch bereinigt, damit kein veralteter Versionsballast liegen bleibt.
- In der GUI zeigt der Fortschrittsbalken den Ablaufstatus; nach erfolgreichem Lauf wechselt er auf **grün** und zeigt **„Fertig“**.

## EXE-Release-Struktur

Das EXE-Release ist bewusst schlank gehalten und enthält nur:

- `AP1-Konfigurator.exe`
- `data/`
- `docs/`
- `README.md`

Die PowerShell-Startskripte und `Skript-Module` werden nicht separat mit ausgeliefert, sondern befinden sich in der EXE und werden beim Start nach `%LOCALAPPDATA%` kopiert.

## Weiterführende Doku

- Anwender-Dokumentation: [`docs/DOKUMENTATION_ANWENDER.md`](./docs/DOKUMENTATION_ANWENDER.md)
- Technische Details: [`docs/DOKUMENTATION_TECHNIK.md`](./docs/DOKUMENTATION_TECHNIK.md)
- Checkliste: [`docs/DOKUMENTATION_CHECKLISTE.md`](./docs/DOKUMENTATION_CHECKLISTE.md)
- Änderungen: [`docs/CHANGELOG.md`](./docs/CHANGELOG.md)
- Release-Hinweise: [`release/RELEASE_NOTES_v1.0.30.md`](./release/RELEASE_NOTES_v1.0.30.md)
