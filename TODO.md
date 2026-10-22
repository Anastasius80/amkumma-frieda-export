# TODO

## Offen

- [ ] Sichtprüfung durch den Benutzer: Excel, Archivansicht (`.html`, Doppelklick in Edge inkl. Dokument-Links), Druck (Strg+P bzw. „Grabblatt drucken“), Grabblätter (PDF), Dokumentordner.
- [ ] Belegung (Belegt / Frei / Möglich) und Grabstatus an einigen Gräbern mit der Anzeige im alten Programm vergleichen.

## Erledigt

- [x] PowerShell-Variante (`powershell/Export-FriedaWeb.ps1`) mit Excel, SQLite, Dokumentexport und HTML-Archivansicht; Endtest mit einer realen Sicherung bestanden (19.10.2026).
- [x] README der PowerShell-Variante inkl. Werkzeuge und Installationsanleitung (`powershell/README.md`).
- [x] Repository veröffentlicht (`main`), Beispieldaten neutral bzw. frei erfunden.
- [x] Grabansicht nach Fachvorgabe: Grabauswahl wie im alten Programm, Statusfilter, Grabinfo mit Belegung und Notiz, Verstorbene im Grab, Excel-Blatt „Verstorbene im Grab“; Personensuche entfernt (20./21.10.2026).
- [x] Druck nach Aufklappzustand; Grabblätter als PDF je Grab (`-Grabblaetter`, `Export-FriedaWebGrabblaetter.ps1`) (22.10.2026).
- [x] Musteransicht mit frei erfundenen Daten (`powershell/muster/`, `New-FriedaWebMuster.ps1`).

## Hinweis

Die Linux-Variante (`linux/`) wurde von Hohenems entwickelt (Alexander Waller) und ist hier unverändert aus dem ursprünglichen Repository übernommen: <https://github.com/walleralexander/friedaweb-export>. Sie ist nicht Gegenstand dieser Erweiterung; Änderungen daran erfolgen nur im Original.
