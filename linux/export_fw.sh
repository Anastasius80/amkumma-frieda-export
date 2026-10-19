#!/bin/bash
# export_fw.sh – FriedaWeb-MS-SQL-Backup (.bak) nach Excel und SQLite umwandeln
#
# Zweck
# -----
# Die Friedhofsverwaltung FriedaWeb speicherte ihre Daten in einer Microsoft-SQL-
# Server-Datenbank. Das Programm wurde 2024 abgelöst; übrig bleibt eine native
# SQL-Server-Sicherung (.bak). Eine solche Datei lässt sich nur von einem SQL
# Server selbst lesen. Dieses Skript stellt die Sicherung in einem kurzlebigen
# Docker-Container wieder her und erzeugt daraus zwei Dateien, die ohne SQL
# Server dauerhaft lesbar bleiben:
#
#   <backup-name>.xlsx    Für die tägliche Arbeit in Excel. Die fachlich
#                         relevanten Tabellen als einzelne Blätter, IDs von
#                         Nachschlagetabellen (Grabart, Reihe, Konfession,
#                         Bestatter …) durch ihren Klartext ersetzt.
#   <backup-name>.sqlite  Verlustfreie 1:1-Kopie aller Tabellen mit Original-
#                         Spaltennamen, IDs, Primär- und Fremdschlüsseln. Dient
#                         als Archiv und als Grundlage für spätere Auswertungen
#                         oder eine eigene Oberfläche.
#
# Aufruf
# ------
#   ./export_fw.sh <backup.bak> [ausgabeordner]
#
#   Beispiele:
#     ./export_fw.sh data/FW_Beispiel_20260101.bak             # Ausgabe nach ./export/
#     ./export_fw.sh /mnt/archiv/FW_Beispiel.bak /tmp/frieda   # eigener Ausgabeordner
#
#   Der Ausgabeordner wird bei Bedarf angelegt. Die Dateinamen richten sich nach
#   dem Namen der Sicherung (FW_Beispiel_20260101.bak -> FW_Beispiel_20260101.xlsx).
#   Vorhandene Dateien gleichen Namens werden ohne Rückfrage überschrieben.
#   Ein Lauf dauert bei einer Datenbank von 25 MB etwa 15 Sekunden.
#
# Voraussetzungen
# ---------------
#   - Linux mit Docker; der Benutzer muss "docker run" ausführen dürfen
#     (root oder Mitglied der Gruppe docker).
#   - python3 mit dem Modul venv, außerdem openssl.
#   - Mindestens 2 GB freier Arbeitsspeicher (Vorgabe von SQL Server) und
#     etwa 2 GB Speicherplatz für das Container-Image.
#   - Internetzugang nur beim ersten Lauf:
#       * das Image mcr.microsoft.com/mssql/server:2022-latest (~1,5 GB) wird geladen,
#       * die Python-Pakete pymssql und openpyxl werden in .venv neben diesem
#         Skript installiert. Danach läuft alles ohne Internet.
#
# Ablauf im Einzelnen
# -------------------
#   1. Beim ersten Lauf wird die virtuelle Python-Umgebung .venv angelegt.
#   2. Ein Zufallspasswort für den SQL-Server-Administrator (sa) wird erzeugt.
#      Es gilt nur für diesen einen Container und wird nirgends gespeichert.
#   3. Ein SQL-Server-2022-Container startet unter dem Namen fwrestore-<PID>:
#        - Express Edition (kostenlos, auch für produktive Nutzung zulässig;
#          die Lizenzbedingungen werden mit ACCEPT_EULA=Y akzeptiert),
#        - kein Port nach außen veröffentlicht; erreichbar ist er nur vom
#          Docker-Host über die interne Bridge-Adresse,
#        - der Ordner mit der Sicherung wird nur lesend unter /backup eingebunden,
#          die .bak-Datei wird also nicht verändert.
#   4. export_fw.py wartet, bis der Server bereit ist, stellt die Datenbank
#      wieder her, schreibt SQLite und Excel und prüft das Ergebnis (Details
#      siehe Beschreibung in export_fw.py).
#   5. Der Container wird am Ende in jedem Fall entfernt, auch wenn ein Schritt
#      fehlschlägt oder das Skript mit Strg+C abgebrochen wird. Die wieder-
#      hergestellte Datenbank liegt nur im Container und verschwindet mit ihm.
#
# Ausgabe am Bildschirm
# ---------------------
#   Nach erfolgreichem Lauf werden beide Dateipfade und die Zeilenzahl je
#   Excel-Blatt ausgegeben. Eine WARNUNG erscheint, wenn Datensätze auf
#   Nachschlagewerte verweisen, die es nicht gibt; diese Zellen enthalten in
#   Excel "?<ID>". Bei Abweichungen zwischen Quelle und SQLite bricht das
#   Skript mit einem AssertionError ab.
#
# Fehlerbehebung
# --------------
#   "permission denied ... docker.sock"   Benutzer darf Docker nicht verwenden.
#   "SQL Server nicht erreichbar"         Container startet nicht, meist zu wenig
#                                         Arbeitsspeicher; prüfen mit
#                                         docker logs fwrestore-<PID> während des Laufs.
#   Fehler beim RESTORE                   Datei ist keine SQL-Server-Vollsicherung
#                                         oder beschädigt.
#   KeyError / Tabelle fehlt              Die Datenbank stammt nicht aus FriedaWeb
#                                         oder aus einer stark abweichenden Version.
#
# Hinweise
# --------
#   - Nur für FriedaWeb-Datenbanken: Die Excel-Blätter sind auf deren Tabellen
#     (tbl_Grab, tbl_Verstorbener …) zugeschnitten.
#   - Excel und SQLite enthalten personenbezogene Daten (Verstorbene,
#     Nutzungsberechtigte mit Adressen) und Bankverbindungen (IBAN, BIC,
#     Mandatsnummern). Ablageort und Zugriffsrechte entsprechend wählen
#     (DSGVO).
#   - Bilder und PDFs sind nicht Teil der Datenbank; sie enthält nur deren
#     Pfade auf dem ehemaligen FriedaWeb-Server.
#
# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Alexander Waller
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
set -euo pipefail

BAK=$(realpath "${1:?Usage: $0 <backup.bak> [ausgabeordner]}")
OUT=$(realpath -m "${2:-export}")
DIR=$(cd "$(dirname "$0")" && pwd)
VENV="$DIR/.venv"
IMAGE=mcr.microsoft.com/mssql/server:2022-latest
NAME="fwrestore-$$"

[ -f "$BAK" ] || { echo "Datei nicht gefunden: $BAK" >&2; exit 1; }
mkdir -p "$OUT"

if [ ! -x "$VENV/bin/python" ]; then
  python3 -m venv "$VENV"
  "$VENV/bin/pip" -q install pymssql openpyxl
fi

export MSSQL_SA_PASSWORD="Tmp$(openssl rand -hex 12)!"
trap 'docker rm -f "$NAME" >/dev/null 2>&1' EXIT
# Express edition, no published port, backup mounted read-only
docker run -d --name "$NAME" -e ACCEPT_EULA=Y -e MSSQL_PID=Express -e MSSQL_SA_PASSWORD \
  -v "$(dirname "$BAK")":/backup:ro "$IMAGE" >/dev/null
export MSSQL_HOST=$(docker inspect -f '{{range .NetworkSettings.Networks}}{{.IPAddress}}{{end}}' "$NAME")

"$VENV/bin/python" -I "$DIR/export_fw.py" "$BAK" "$OUT"
