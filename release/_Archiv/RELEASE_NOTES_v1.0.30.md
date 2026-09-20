# AP1-Konfigurator v1.0.30

## Download

- [Release-Seite v1.0.30](https://github.com/GoroTech-Tools/AP1-Konfigurator/releases/tag/v1.0.30)
- [ZIP direkt herunterladen](https://github.com/GoroTech-Tools/AP1-Konfigurator/releases/download/v1.0.30/AP1-Konfigurator-v1.0.30.zip)

## Highlights

- Stabilere Erststart-Initialisierung für Word und Excel
- Automatisches Ausblenden bzw. Bestätigen typischer Willkommens-/Erststartdialoge
- Bereinigung von veralteten temporären Kandidaten-Ordnern im Projektpfad
- Aktualisierte Projekt-Dokumentation und Release-Informationen nach dem aktuellen Stand

## Änderungen im Detail

### Office-Erststart robuster gemacht

- Beim Erkennen eines First-Run-Starts wurde die Initialisierung von Word und Excel weiter abgesichert.
- Typische Willkommensfenster werden jetzt automatisch mit Standardbestätigung versucht zu schließen, statt nur auf eine manuelle Tastatureingabe zu warten.
- Die COM-Initialisierung prüft jetzt gezielter, ob Word/Excel tatsächlich einsatzbereit sind und fällt bei Problemen sauber auf Registry-only-Betrieb zurück.

### Kandidatenordner wurden bereinigt

- Ein leerer bzw. temporärer Kandidaten-Root wie `Ordner` wurde nach der Bereitstellung auf dem Desktop automatisch bereinigt.
- Dadurch bleiben vorbereitete Projektordner sauber und es entstehen keine lästigen Nebenordner bei wiederholten Läufen.

### Dokumentation und Release-Stand synchronisiert

- Die Anwender- und Technikdokumentation wurde auf den aktuellen Release-Stand gebracht.
- Versionshinweise, Release-Notiz und README wurden auf die tatsächliche Projektversion `v1.0.30` synchronisiert.
- Die Release-Struktur wurde an den aktuellen Build-/Packaging-Prozess angepasst.

### Projekt- und Build-Qualität

- Die Release-Erzeugung wurde mit einer klareren Ordner-/ZIP-Struktur und konsistenter Versionspflege ergänzt.
- Die wichtigsten Laufzeitpfade wurden überprüft, damit alte Versionen nicht in lokalen `%LOCALAPPDATA%`-Ordnern weiterlaufen.

## Validierung

- Versionen und Release-Notizen wurden auf den aktuellen Stand abgeglichen.
- Projekt-Dokumentation und README wurden mit dem tatsächlichen Releasezustand konsistent gehalten.
- Die vorhandenen Release-Artefakte wurden auf die aktuelle Struktur und den aktuellen Versionsnamen geprüft.
- Die Erststart-Verbesserungen wurden in der Logik an der Office-Initialisierung verankert, damit sie beim ersten Start der Office-Anwendungen greift.

## Artefakte und GitHub-Hinweis

- Git-Tag: `v1.0.30`
- GitHub-Release: `v1.0.30`
- Lokales Release-Archiv: `release/AP1-Konfigurator-v1.0.30.zip`
- Releaseordner: `release/AP1-Konfigurator-v1.0.30/`
- Projekt-Dokumentation: `docs/`
- Wichtig: Das Release beschreibt den aktuellen Projektzustand und nicht nur den reinen Build-Stand, sondern auch die operative Stabilisierung im Office-Erststart.
