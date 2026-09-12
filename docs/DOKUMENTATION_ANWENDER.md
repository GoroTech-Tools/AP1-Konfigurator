# DOKUMENTATION ANWENDER

Aktueller Stand: **v1.0.30** · Letzte Aktualisierung: **12. September 2026**

## Inhaltsverzeichnis

- [Zweck](#zweck)
- [Voraussetzungen](#voraussetzungen)
- [Vorbereitung](#vorbereitung)
- [Start des Skripts](#start-des-skripts)
- [Ablauf während der Ausführung](#ablauf-während-der-ausführung)
- [Ergebnis prüfen](#ergebnis-prüfen)
- [Häufige Probleme](#häufige-probleme)
- [Support-Hinweis](#support-hinweis)

## Zweck

Der `AP1-Konfigurator` richtet einen Prüfungsrechner für die Abschlussprüfung Teil 1 (AP1) standardisiert ein.

Dabei werden insbesondere vorbereitet:

- Office-Einstellungen ohne erforderliches COM
- Office-Vorlagen und Schnellzugriff
- Kandidatenordner aus der Teilnehmerliste
- Nuera-Dateien auf dem Desktop
- optionale Proxy-Einstellungen
- ein Laufzeitprotokoll zur Nachvollziehbarkeit

## Voraussetzungen

- Windows-PC mit Microsoft Office (Word und Excel) für die spätere Nutzung
- Schreibrechte im Benutzerprofil
- vorhandene Teilnehmerliste unter `data/1. Anpassen\AP1-TN.xlsx`
- Python-Abhängigkeit `openpyxl` für die Excel-Auswertung
- optional Internetzugang, falls keine lokale Nuera-Datei vorhanden ist

## Vorbereitung

1. Prüfen Sie, ob `AP1-TN.xlsx` im Ordner `data/1. Anpassen` vorhanden und aktuell ist.
2. Schließen Sie Word und Excel vor dem Start, damit Vorlagen sicher kopiert werden können.
3. Speichern Sie offene Arbeiten anderer Programme.
4. Entscheiden Sie, ob ein Proxy gesetzt werden soll.
5. Halten Sie den Desktop des aktuellen Benutzers frei, damit die erzeugten Ordner sichtbar bleiben.

## Start des Skripts

Empfohlener Start:

1. `dist/AP1-Konfigurator.exe` (oder im Release die `AP1-Konfigurator.exe`) starten.
2. In der GUI den gewünschten Proxy-Modus (`Skip`, `On`, `Off`) auswählen.
3. Auf **„AP1-Konfiguration starten“** klicken.
4. Den Fortschrittsbalken im unteren Bereich beobachten, bis **„Fertig“** angezeigt wird.

Alternativ per PowerShell:

- `./src/AP1-Konfigurator.ps1`
- `./src/AP1-Konfigurator.ps1 -Proxy Off`
- `./src/AP1-Konfigurator.ps1 -RegistryOnly -CsvFallbackPath .\data\1. Anpassen\AP1-TN.csv`
- optional `-UseCom`, wenn der COM-basierte Office-Modus ausdrücklich benötigt wird

## Ablauf während der Ausführung

Das Skript führt typischerweise folgende Schritte aus:

1. lokale Nuera-Dateien prüfen und nur bei Bedarf herunterladen
2. Nuera-Ordner auf dem Desktop bereitstellen
3. Office-Einstellungen über die Registry setzen
4. Vorlagen `Normal.dotm` und `Mappe.xltx` kopieren
5. Teilnehmerliste mit `openpyxl` lesen
6. Anmeldenamen in Spalte A suchen und genau den Ordnernamen aus Spalte B auf dem Desktop anlegen
7. Word- und Excel-Speicherpfade auf diesen Benutzerordner setzen
8. Taskleisten- und optional Proxy-Einstellungen anwenden
9. Logdatei in `data/4. Logs` schreiben

Hinweise während der Ausführung:

- Im Standardbetrieb werden Word und Excel nicht per COM gestartet.
- Der COM-Modus ist nur mit `-UseCom` aktiv.
- Eine vorhandene gültige Nuera wird lokal verwendet; ein Download erfolgt nur bei fehlendem Bestand.
- Die GUI zeigt den Fortschritt laufend an; bei erfolgreichem Ende wird der Balken grün und mit `Fertig` beschriftet.

## Ergebnis prüfen

Nach erfolgreicher Ausführung sollten Sie insbesondere Folgendes prüfen:

- aktueller Nuera-Ordner liegt auf dem Desktop
- genau ein Benutzerordner aus Spalte B wurde auf dem Desktop angelegt
- Word/Excel-Speicherpfade zeigen auf diesen Benutzerordner
- Schnellzugriff in Word/Excel wurde übernommen
- im Ordner `data/4. Logs` wurde eine aktuelle Logdatei angelegt

## Häufige Probleme

### Excel-Datei fehlt

- Prüfen, ob `data/1. Anpassen\AP1-TN.xlsx` vorhanden ist.
- Falls die Datei an anderem Ort liegt, `-ExcelListPath` verwenden.

### Word oder Excel startet nicht automatisch

- Das ist im Standardbetrieb unkritisch, da kein COM-Start erforderlich ist.
- Nur bei bewusst gesetztem `-UseCom`: Office einmal manuell öffnen und Hinweise bestätigen.

### Kandidatenordner wird nicht angelegt

- Anmeldename in Spalte A und gewünschte Ordnerbezeichnung in Spalte B prüfen.
- Bei Excel-Problemen optional einen CSV-Fallback verwenden.

### Proxy-Rückfrage stört einen automatisierten Lauf

- Skript mit `-Quiet` starten.

### Nuera-Dateien werden nicht geladen

- Prüfen, ob unter `data/3. Nuera-Dateien` ein gültiges ZIP oder ein gefüllter Nuera-Ordner vorhanden ist.
- Bei leerem Ordner wird das lokale ZIP erneut entpackt.
- Nur wenn kein lokaler Bestand vorhanden ist, wird eine Internetverbindung benötigt.

## Support-Hinweis

Bei Rückfragen bitte immer mitgeben:

- Datum/Uhrzeit der Ausführung
- Name der erzeugten Logdatei aus `data/4. Logs`
- kurze Beschreibung des letzten sichtbaren Schritts
- wenn möglich Screenshot einer Fehlermeldung
