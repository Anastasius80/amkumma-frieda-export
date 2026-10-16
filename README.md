# friedaweb-export

Wandelt eine Microsoft-SQL-Server-Sicherung (`.bak`) der Friedhofsverwaltung **FriedaWeb** in zwei Dateien um, die ohne SQL Server lesbar bleiben:

- **Excel** (`.xlsx`): Gräber, Verstorbene, Nutzungsberechtigte, Nutzungsverlauf, Dokumente und Protokolle als einzelne Blätter. Nachschlage-IDs (Grabart, Reihe, Konfession, Bestatter …) sind durch Klartext ersetzt.
- **SQLite** (`.sqlite`): verlustfreie 1:1-Kopie aller Tabellen mit IDs, Primär- und Fremdschlüsseln.

Gedacht für die Archivierung, nachdem FriedaWeb durch ein anderes Programm abgelöst wurde.

## Verwendung

```bash
./export_fw.sh <backup.bak> [ausgabeordner]     # Standard-Ausgabeordner: ./export
```

Voraussetzungen: Linux mit Docker, `python3` mit `venv`, `openssl`, mindestens 2 GB freier Arbeitsspeicher. Beim ersten Lauf werden das SQL-Server-Image und die Python-Pakete `pymssql` und `openpyxl` geladen.

Das Skript stellt die Sicherung in einem kurzlebigen SQL-Server-Container (Express Edition, ohne veröffentlichten Port) wieder her, exportiert, prüft das Ergebnis und entfernt den Container wieder.

Die ausführliche Beschreibung steht am Anfang von [`export_fw.sh`](export_fw.sh) (Ablauf, Fehlerbehebung) und [`export_fw.py`](export_fw.py) (Aufbau von Excel und SQLite, Prüfungen, Eigenheiten der Daten).

> Die erzeugten Dateien enthalten personenbezogene Daten und Bankverbindungen. Ablageort und Zugriffsrechte entsprechend wählen.

## Lizenz

MIT, siehe [LICENSE](LICENSE).
