# DOKUMENTATION CHECKLISTE

Aktueller Stand: **v1.0.32** · Letzte Aktualisierung: **15. September 2026**

## Vor dem Lauf

- [ ] `AP1-TN.xlsx` ist aktuell und im Ordner `data/1. Anpassen` vorhanden.
- [ ] Word und Excel sind geschlossen, damit Vorlagen sicher kopiert werden können.
- [ ] Benötigte Vorlagen liegen unter `data/2. Bei Bedarf anpassen` bereit.
- [ ] `AP1-TN.xlsx` enthält den Windows-Anmeldenamen in Spalte A und die gewünschte Ordnerbezeichnung in Spalte B.
- [ ] Lokaler Nuera-Bestand unter `data/3. Nuera-Dateien` wurde geprüft; Internet ist nur bei fehlendem Bestand nötig.
- [ ] Entscheidung zu Proxy `On`, `Off` oder `Skip` ist getroffen.

## Während des Laufs

- [ ] EXE/GUI oder `src/AP1-Konfigurator.bat` bzw. `src/AP1-Konfigurator.ps1` gestartet.
- [ ] Im Standardmodus wird kein Office-COM gestartet; `-UseCom` nur bei ausdrücklichem Bedarf verwenden.
- [ ] keine sichtbare Fehlermeldung im PowerShell-Fenster übersehen.

## Nach dem Lauf

- [ ] Aktueller Nuera-Ordner liegt auf dem Desktop.
- [ ] Genau ein Ordner aus Spalte B für den angemeldeten Benutzer liegt auf dem Desktop.
- [ ] Logdatei wurde unter `data/4. Logs` erstellt.
- [ ] Word-/Excel-Vorlagen wurden korrekt übernommen.
- [ ] Word- und Excel-Speicherpfade zeigen auf den Benutzerordner aus Spalte B.
- [ ] Taskleisten-/Proxy-Einstellungen entsprechen dem gewünschten Zustand.

## Bei Problemen

- [ ] aktuelle Logdatei aus `data/4. Logs` prüfen.
- [ ] Bei `-UseCom` Office einmal manuell starten und erneut testen; im Standardmodus ist dies nicht erforderlich.
- [ ] bei Excel-Problemen optional CSV-Fallback verwenden.
- [ ] bei leerem Nuera-Ordner prüfen, ob das zugehörige lokale ZIP vorhanden ist.
- [ ] Pfade und Schreibrechte im Benutzerprofil kontrollieren.
