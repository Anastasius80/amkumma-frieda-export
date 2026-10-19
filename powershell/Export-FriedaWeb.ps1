#Requires -Version 5.1
<#
.SYNOPSIS
    Wandelt eine FriedaWeb-MS-SQL-Sicherung (.bak) unter Windows in eine Excel- und eine SQLite-Datei um.

.DESCRIPTION
    Windows-/PowerShell-Portierung von linux/export_fw.sh und linux/export_fw.py.

    Zweck
    -----
    Die Friedhofsverwaltung FriedaWeb speicherte ihre Daten in einer Microsoft-SQL-Server-
    Datenbank. Übrig bleibt eine native SQL-Server-Sicherung (.bak), die nur ein SQL Server
    lesen kann. Dieses Skript stellt die Sicherung in einer kurzlebigen SQL-Server-Express-
    LocalDB-Instanz wieder her und erzeugt daraus zwei Dateien, die ohne SQL Server dauerhaft
    lesbar bleiben:

      <backup-name>.xlsx    Für die tägliche Arbeit in Excel. Die fachlich relevanten
                            Tabellen als einzelne Blätter, IDs von Nachschlagetabellen
                            (Grabart, Reihe, Konfession, Bestatter …) durch Klartext ersetzt.
      <backup-name>.sqlite  Verlustfreie 1:1-Kopie aller Tabellen mit Original-Spaltennamen,
                            IDs, Primär- und Fremdschlüsseln.
      <backup-name>_Dokumente\
                            Alle Dokumente als einzelne Dateien (siehe "Dokumente" unten).
                            Excel verlinkt relativ darauf; Excel-Datei und Ordner daher
                            gemeinsam ablegen bzw. verschieben.
      <backup-name>.html    Archivansicht für den Browser (offline, ohne Installation, nur Lesen):
                            Einstieg über Gräber oder Personen, Suche und Filter, Detailseiten mit
                            allen verknüpften Daten und Dokument-Links. Enthält alle Excel-Blätter
                            außer Verlauf und Ereignisprotokoll. Vorlage: Export-FriedaWeb.Viewer.html.

    Ablauf
    ------
      1. SqlLocalDB.exe und sqlite3.exe werden gesucht.
      2. Eine eigene LocalDB-Instanz "fwrestore_<PID>" wird angelegt und gestartet. Sie läuft
         unter dem angemeldeten Windows-Benutzer, verwendet Windows-Anmeldung (kein Passwort)
         und ist nur lokal über Named Pipes erreichbar (kein Netzwerkport).
      3. Die Datenbank wird in einen temporären Arbeitsordner unter %TEMP% wiederhergestellt.
         Die .bak-Datei wird nur gelesen. Eine ältere Sicherung (z. B. SQL Server 2017) wird
         dabei automatisch auf das Format der LocalDB-Version aktualisiert; die Daten bleiben
         unverändert.
      4. SQLite und Excel werden im Arbeitsordner erzeugt und geprüft. Erst danach werden sie in
         den Ausgabeordner verschoben; vorhandene Dateien gleichen Namens werden überschrieben.
      5. LocalDB-Instanz und Arbeitsordner werden am Ende in jedem Fall entfernt, auch bei
         einem Fehler oder einem Abbruch mit Strg+C.

    SQLite – vollständige Kopie
    ---------------------------
    Jede Tabelle wird mit allen Zeilen übernommen, auch leere Tabellen. Spaltennamen,
    Primärschlüssel und Fremdschlüssel entsprechen dem Original. Typen:
        int, smallint, tinyint, bit   -> INTEGER (bit als 0/1)
        float                         -> REAL
        money                         -> NUMERIC
        varbinary                     -> BLOB
        datetime                      -> TEXT im Format "YYYY-MM-DD HH:MM:SS.mmm"
        nvarchar, varchar, char       -> TEXT

    Excel – Arbeitsdatei
    --------------------
    Blatt "Info": Quelle, Friedhof, letzter Sterbefall, letzte Änderung, Zeilenzahl je Blatt.
    Danach je ein Blatt für Gräber, Verstorbene, Verstorbene_Archiv, Berechtigte,
    Berechtigte_Archiv, Nutzungsverlauf, Dokumente, Verlauf und Ereignisprotokoll (wie Linux) sowie,
    sofern vorhanden, Grabstellen, Grab_Grabstelle, Personen, Berechtigungen, Abrechnungen, Bescheide
    und Dokumentablage. Blätter ohne Zeilen entfallen; das Info-Blatt nennt sie.
      - Verweise auf Personen (tbl_Person, tbl_Verstorbener, tbl_Berechtigter) behalten die ID und
        erhalten daneben eine Spalte mit "Nachname, Vorname".
      - Fremdschlüssel auf Nachschlagetabellen zeigen deren Klartext (erste befüllte Spalte aus
        Bezeichnung, Firmenname, Firma, Block); die Überschrift verliert Präfix und "_ID".
      - Verweise auf ein Grab behalten die ID und erhalten daneben die Spalte "Grab"
        ("Reihe / Grabnummer").
      - In den Blättern Dokumente (tbl_Bild) und Dokumentablage (tbl_Dokument, sofern vorhanden)
        ist "Titel#Pfad" in die Spalten Titel und Pfad geteilt. Die Binärspalte "Daten" wird
        durch "Datei" (Link auf die exportierte Datei) und "Dateiquelle" ersetzt.
      - Datumswerte sind echte Excel-Datumswerte (TT.MM.JJJJ, mit Uhrzeit nur wenn vorhanden);
        Werte vor 1900 stehen als Text in der Zelle.
      - Texte werden immer als Text gespeichert (kein Ausführen von "=..." als Formel).
      - Erste Zeile fixiert, Autofilter auf allen Spalten.

    Dokumente
    ---------
    Alle Dokumente aus tbl_Dokument und tbl_Bild landen genau einmal im Ordner
    <backup-name>_Dokumente (ohne Unterordner):
      - Vorrang hat die Datei im Dokumentordner (-DocumentPath, Standard backup\Dokumente). Sie wird
        zuerst über den Pfadteil nach "\<Ordnername>\", danach über einen eindeutigen Dateinamen
        gesucht und behält ihren Originalnamen.
      - Fehlt sie dort, wird die in der Spalte "Daten" eingebettete Fassung geschrieben. Die Endung
        richtet sich nach dem tatsächlichen Inhalt (z. B. .docx statt .doc bei Word-2007-Dateien).
      - Mehrere Einträge mit derselben Datei verweisen auf dieselbe Datei. Bei Namensgleichheit
        verschiedener Dateien wird die ID bzw. der Herkunftsordner vorangestellt.
      - Dateien im Dokumentordner ohne Bezug in der Datenbank werden ebenfalls übernommen.
    Nicht auffindbare Dokumente erscheinen in Excel mit Dateiquelle "fehlt" und als WARNUNG.

    Prüfungen
    ---------
    Abbruch, wenn die Zeilenzahl einer Tabelle in SQLite von der Quelle abweicht, SQLite
    Fremdschlüssel-Verletzungen meldet oder ein Text länger als 32.767 Zeichen ist. Ausnahme:
    Textschlüssel, die nur nach SQL-Server-Regeln (ohne Groß-/Kleinschreibung, ohne Leerzeichen
    am Ende) zusammenpassen, werden als HINWEIS ausgegeben; die Werte bleiben unverändert.
    Verweise ohne Nachschlagewert erscheinen in Excel als "?<ID>" und werden als WARNUNG gezählt.

    Hinweise
    --------
      - Nur für FriedaWeb-Datenbanken.
      - Excel und SQLite enthalten personenbezogene Daten (Verstorbene, Nutzungsberechtigte mit
        Adressen) und Bankverbindungen (IBAN, BIC, Mandatsnummern). Ablageort und Zugriffsrechte
        entsprechend wählen (DSGVO). Die Ordner backup/ und export/ sind per .gitignore
        vom Repository ausgeschlossen.
      - Je nach Installation sind Dokumente in der Datenbank eingebettet oder liegen nur als
        Dateien auf dem ehemaligen FriedaWeb-Server; diese müssen dann im Dokumentordner liegen.

.PARAMETER BackupPath
    Pfad zur .bak-Datei. Ohne Angabe wird die einzige .bak-Datei im Ordner "backup" neben
    diesem Skript verwendet.

.PARAMETER OutputPath
    Ausgabeordner. Standard ist der Ordner "export" neben diesem Skript. Er wird bei Bedarf
    angelegt.

.PARAMETER DocumentPath
    Ordner mit den Dokumentdateien des ehemaligen FriedaWeb-Servers. Standard ist der Ordner
    "Dokumente" im Ordner der Sicherung. Fehlt er, werden nur eingebettete Dokumente exportiert.

.PARAMETER Sqlite3Path
    Pfad zu sqlite3.exe, falls das Programm nicht über PATH oder eine winget-Installation
    gefunden wird.

.PARAMETER LocalDbVersion
    Zu verwendende LocalDB-Version (z. B. "15.0" für SQL Server 2019, "16.0" für 2022). Ohne
    Angabe wird die neueste installierte Version verwendet. Eine Sicherung kann nur von einer
    gleichen oder neueren SQL-Server-Version wiederhergestellt werden.

.PARAMETER KeepSqlite
    Die SQLite-Kopie (alle Tabellen 1:1) zusätzlich im Ausgabeordner behalten. Ohne diesen Schalter
    wird sie nur intern für die Prüfungen erzeugt und danach gelöscht; die vollständige Rohquelle
    ist dann die Original-Sicherung (.bak) samt Dokumentordner, die aufbewahrt werden sollte.

.EXAMPLE
    .\Export-FriedaWeb.ps1
    Verwendet die .bak-Datei aus .\backup und schreibt nach .\export.

.EXAMPLE
    .\Export-FriedaWeb.ps1 -BackupPath D:\Archiv\FW_Beispiel_20260101.bak -OutputPath D:\Export

.NOTES
    Voraussetzungen: Windows, PowerShell 5.1 oder 7, SQL Server Express LocalDB
    (SqlLocalDB.exe), sqlite3.exe. Keine Administratorrechte und kein Internetzugang
    während des Laufs erforderlich.

    SPDX-License-Identifier: MIT
    Copyright (c) 2026 Alexander Waller
    PowerShell-Portierung 2026 für den Gemeindeverband IT amKumma.
    Lizenztext siehe LICENSE im Hauptverzeichnis des Repositorys.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$BackupPath,

    [Parameter(Position = 1)]
    [string]$OutputPath,

    [string]$DocumentPath,

    [string]$Sqlite3Path,

    [string]$LocalDbVersion,

    [switch]$KeepSqlite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Fachliche Einstellungen (identisch zu linux/export_fw.py) ---------------------------------
# Optional = nur wenn die Tabelle existiert (Ergänzungen der Windows-Variante, z. B. für das
# personenbasierte Datenmodell anderer FriedaWeb-Installationen). Blätter ohne Zeilen entfallen.
$Sheets = @(
    @{ Name = 'Gräber'; Table = 'tbl_Grab' }
    @{ Name = 'Grabstellen'; Table = 'tbl_Grabstelle'; Optional = $true }
    @{ Name = 'Grab_Grabstelle'; Table = 'tblkett_Grab_Grabstelle'; Optional = $true }
    @{ Name = 'Verstorbene'; Table = 'tbl_Verstorbener' }
    @{ Name = 'Verstorbene_Archiv'; Table = 'tbl_VerstorbenerArchiv' }
    @{ Name = 'Personen'; Table = 'tbl_Person'; Optional = $true }
    @{ Name = 'Berechtigte'; Table = 'tbl_Berechtigter' }
    @{ Name = 'Berechtigte_Archiv'; Table = 'tbl_BerechtigterArchiv' }
    @{ Name = 'Berechtigungen'; Table = 'tblkett_Berechtigter_Objekt'; Optional = $true }
    @{ Name = 'Nutzungsverlauf'; Table = 'tbl_GrabNutzVerlauf' }
    @{ Name = 'Abrechnungen'; Table = 'tbl_Abrechnung'; Optional = $true }
    @{ Name = 'Bescheide'; Table = 'tbl_Bescheid'; Optional = $true }
    @{ Name = 'Dokumente'; Table = 'tbl_Bild' }
    @{ Name = 'Dokumentablage'; Table = 'tbl_Dokument'; Optional = $true }
    @{ Name = 'Verlauf'; Table = 'tbl_Verlauf' }
    @{ Name = 'Ereignisprotokoll'; Table = 'tbl_EventProtokoll' }
    # Weitere Fachtabellen (Secondary): im Excel als Blatt, im HTML als eingeklappter Unterpunkt.
    @{ Name = 'Protokoll'; Table = 'tbl_Protokoll'; Optional = $true; Secondary = $true }
    @{ Name = 'Sperrungen'; Table = 'tbl_Sperrung'; Optional = $true; Secondary = $true }
    @{ Name = 'Grabkontrollen'; Table = 'tbl_Grabkontrolle'; Optional = $true; Secondary = $true }
    @{ Name = 'Kontrollmängel'; Table = 'tblkett_Kontrolle_Mangel'; Optional = $true; Secondary = $true }
    @{ Name = 'Grabmalgenehmigungen'; Table = 'tbl_Grabmalgenehmigung'; Optional = $true; Secondary = $true }
    @{ Name = 'Grabmalauflagen'; Table = 'tblkett_GMauflage_GMgenehm'; Optional = $true; Secondary = $true }
    @{ Name = 'Rechnungsempfänger'; Table = 'tbl_EmpfaengerFreieRechnung'; Optional = $true; Secondary = $true }
    @{ Name = 'Bankverbindung'; Table = 'tbl_Bankverbindung'; Optional = $true; Secondary = $true }
    @{ Name = 'Berichtsvorlagen'; Table = 'tbl_Bericht'; Optional = $true; Secondary = $true }
)
# Referenz- bzw. Nachschlagetabellen: werden als Blatt "Ref_<Name>" ergänzt, sofern eine Exporttabelle
# auf sie verweist; im HTML sind ihre Klartexte als Link auf die Detailseite anklickbar.
$ReferenceTablePattern = '^(tblxx_|loktbl_)'
$ReferenceSheetPrefix = 'Ref_'
# Verknüpfungen, die in FriedaWeb ohne Fremdschlüssel gespeichert sind. Inhaltlich an einer realen
# Sicherung gegengeprüft; beim Export werden sie nur übernommen, wenn jede ID im Ziel existiert.
# Targets: erste vorhandene, nicht leere Zieltabelle (personenbasiertes bzw. älteres Datenmodell).
$ImplicitForeignKeys = @(
    @{ Table = 'tbl_Person'; Column = 'tbl_Grab_ID'; Targets = @('tbl_Grab') }
    @{ Table = 'tbl_Bescheid'; Column = 'tbl_Grab_ID'; Targets = @('tbl_Grab') }
    @{ Table = 'tbl_Abrechnung'; Column = 'tbl_Verstorbener_ID'; Targets = @('tbl_Person', 'tbl_Verstorbener') }
    @{ Table = 'tbl_Abrechnung'; Column = 'tbl_Berechtigter_ID'; Targets = @('tbl_Person', 'tbl_Berechtigter') }
    @{ Table = 'tbl_Dokument'; Column = 'tbl_Bezugs_ID'; Targets = @('tbl_Person') }
    @{ Table = 'tbl_Grabkontrolle'; Column = 'tbl_Grab_ID'; Targets = @('tbl_Grab') }
    @{ Table = 'tblkett_GMauflage_GMgenehm'; Column = 'tbl_Grabmalgenehmigung_ID'; Targets = @('tbl_Grabmalgenehmigung') }
    @{ Table = 'tblkett_GMauflage_GMgenehm'; Column = 'tbl_Grab_ID'; Targets = @('tbl_Grab') }
)
# Verweise auf diese Tabellen behalten die ID und erhalten daneben "Nachname, Vorname".
$PersonTables = @('tbl_Person', 'tbl_Verstorbener', 'tbl_Berechtigter')
$LabelColumns = @('Bezeichnung', 'Firmenname', 'Firma', 'Block')
$RequiredTables = @($Sheets | Where-Object { -not $_.ContainsKey('Optional') } | ForEach-Object { $_.Table }) +
    @('tblxx_Reihe', 'tblxx_Friedhof')
# Tabellen mit Dokumenten: Spalte im Format "Titel#Pfad" und optional eingebettete Datei in "Daten".
# Folder: Unterordner im Dokumentordner ('' = Hauptebene).
$DocumentTables = @{
    tbl_Dokument = @{ PathColumn = 'Dokumentpfad'; Folder = '' }
    tbl_Bild     = @{ PathColumn = 'Bildpfad'; Folder = '' }
    tbl_Bericht  = @{ PathColumn = 'Dokumentpfad'; Folder = 'Berichtsvorlagen' }
}
$DocumentDataColumn = 'Daten'
# HTML-Archivansicht: alle Excel-Blätter außer den reinen Protokollen.
$ViewerTemplateName = 'Export-FriedaWeb.Viewer.html'
$ViewerPlaceholder = '/*__FW_DATA__*/null'
$ViewerExcludedTables = @('tbl_Verlauf', 'tbl_EventProtokoll')
$DateTypes = @('datetime', 'datetime2', 'smalldatetime', 'date')
$CharacterTypes = @('nvarchar', 'varchar', 'nchar', 'char')
$AuditColumns = @('Creator', 'CreateDate', 'Updater', 'UpdateDate')
$SqliteTypes = @{
    int       = 'INTEGER'
    smallint  = 'INTEGER'
    tinyint   = 'INTEGER'
    bit       = 'INTEGER'
    float     = 'REAL'
    money     = 'NUMERIC'
    varbinary = 'BLOB'
}

# --- Technische Konstanten ---------------------------------------------------------------------
$Invariant = [Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
$IllegalXmlChars = [regex]'[\x00-\x08\x0B\x0C\x0E-\x1F]'
$SqliteUnsafeChars = [regex]'[\x00-\x08\x0B-\x1F]'
# Excel zählt fälschlich den 29.02.1900 mit; Seriennummern davor sind um 1 kleiner als OADate.
$Excel1900LeapBugEnd = New-Object DateTime -ArgumentList 1900, 3, 1
$NsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$NsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NsPkgRel = 'http://schemas.openxmlformats.org/package/2006/relationships'
$HasToHexString = $null -ne [Convert].GetMethod('ToHexString', [type[]]@([byte[]]))
# Excel-Formate (Index in styles.xml): 0 Standard, 1 fett, 2 Datum, 3 Datum mit Uhrzeit, 4 Link
$StyleBold = 1
$StyleLink = 4

#region Hilfsfunktionen

function Invoke-Tool {
    <#
    .SYNOPSIS
        Startet ein Kommandozeilenprogramm und bricht bei ExitCode ungleich 0 mit dessen Ausgabe ab.
    #>
    param([string]$FilePath, [string[]]$ArgumentList)

    # Unter Windows PowerShell 5.1 würde stderr-Ausgabe sonst bei 'Stop' sofort abbrechen.
    $ErrorActionPreference = 'Continue'
    $output = @(& $FilePath @ArgumentList 2>&1 | ForEach-Object { "$_" })
    if ($LASTEXITCODE -ne 0) {
        throw ('{0} {1} ist fehlgeschlagen (ExitCode {2}):{3}{4}' -f
            [IO.Path]::GetFileName($FilePath), ($ArgumentList -join ' '), $LASTEXITCODE,
            [Environment]::NewLine, ($output -join [Environment]::NewLine))
    }
    return $output
}

function Resolve-Tool {
    <#
    .SYNOPSIS
        Ermittelt den Pfad eines benötigten Programms (Parameter, PATH oder bekannte Ablageorte).
    #>
    param([string]$Name, [string]$ExplicitPath, [string[]]$Candidates, [string]$Hint)

    if ($ExplicitPath) {
        if (-not (Test-Path -LiteralPath $ExplicitPath -PathType Leaf)) {
            throw "$Name nicht gefunden: $ExplicitPath"
        }
        return (Resolve-Path -LiteralPath $ExplicitPath).ProviderPath
    }

    $command = Get-Command $Name -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($command) {
        return $command.Source
    }

    foreach ($pattern in $Candidates) {
        $hit = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($hit) {
            return $hit.FullName
        }
    }

    throw "$Name wurde nicht gefunden. $Hint"
}

function Open-SqlConnection {
    <#
    .SYNOPSIS
        Öffnet eine Verbindung zur temporären LocalDB-Instanz; wartet bis zu drei Minuten.
    #>
    param([string]$Instance, [string]$Database)

    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder['Data Source'] = "(localdb)\$Instance"
    $builder['Initial Catalog'] = $Database
    $builder['Integrated Security'] = $true
    $builder['Pooling'] = $false
    $builder['Connect Timeout'] = 15

    $lastError = $null
    $deadline = (Get-Date).AddMinutes(3)
    while ((Get-Date) -lt $deadline) {
        $connection = New-Object System.Data.SqlClient.SqlConnection -ArgumentList $builder.ConnectionString
        try {
            $connection.Open()
            return $connection
        }
        catch {
            $lastError = $_.Exception.Message
            $connection.Dispose()
            Start-Sleep -Seconds 2
        }
    }
    throw "SQL Server nicht erreichbar: $lastError"
}

function New-SqlCommand {
    <#
    .SYNOPSIS
        Erzeugt einen SQL-Befehl ohne Zeitlimit mit optionalen Parametern.
    #>
    param($Connection, [string]$Sql, [hashtable]$Parameters)

    $command = $Connection.CreateCommand()
    $command.CommandText = $Sql
    $command.CommandTimeout = 0
    if ($Parameters) {
        foreach ($key in $Parameters.Keys) {
            [void]$command.Parameters.AddWithValue($key, $Parameters[$key])
        }
    }
    return $command
}

function Invoke-SqlQuery {
    <#
    .SYNOPSIS
        Führt eine Abfrage aus und liefert Spaltennamen und alle Zeilen (DBNull als $null).
    #>
    param($Connection, [string]$Sql, [hashtable]$Parameters)

    $command = New-SqlCommand -Connection $Connection -Sql $Sql -Parameters $Parameters
    try {
        $reader = $command.ExecuteReader()
        try {
            $fieldCount = $reader.FieldCount
            $names = New-Object 'string[]' $fieldCount
            for ($i = 0; $i -lt $fieldCount; $i++) {
                $names[$i] = $reader.GetName($i)
            }

            $rows = New-Object 'System.Collections.Generic.List[object[]]'
            while ($reader.Read()) {
                $values = New-Object 'object[]' $fieldCount
                [void]$reader.GetValues($values)
                for ($i = 0; $i -lt $fieldCount; $i++) {
                    if ($values[$i] -is [DBNull]) {
                        $values[$i] = $null
                    }
                }
                $rows.Add($values)
            }

            return [pscustomobject]@{
                Columns = $names
                Rows    = $rows
            }
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $command.Dispose()
    }
}

function Invoke-SqlNonQuery {
    <#
    .SYNOPSIS
        Führt einen SQL-Befehl ohne Ergebnismenge aus.
    #>
    param($Connection, [string]$Sql, [hashtable]$Parameters)

    $command = New-SqlCommand -Connection $Connection -Sql $Sql -Parameters $Parameters
    try {
        [void]$command.ExecuteNonQuery()
    }
    finally {
        $command.Dispose()
    }
}

function Get-ColumnIndex {
    <#
    .SYNOPSIS
        Liefert die Position einer Spalte und bricht ab, wenn sie fehlt.
    #>
    param([string[]]$Names, [string]$Name, [string]$Context)

    $index = [array]::IndexOf($Names, $Name)
    if ($index -lt 0) {
        throw "Spalte '$Name' fehlt ($Context). Die Datenbank stammt nicht aus FriedaWeb oder aus einer stark abweichenden Version."
    }
    return $index
}

function ConvertTo-LookupKey {
    <#
    .SYNOPSIS
        Vereinheitlicht Schlüsselwerte (z. B. smallint und int) für Nachschlagetabellen.
    #>
    param($Value)

    if ($null -eq $Value) {
        return ''
    }
    return [Convert]::ToString($Value, $Invariant)
}

function Format-SqliteName {
    <#
    .SYNOPSIS
        Setzt einen Bezeichner für SQLite in doppelte Anführungszeichen.
    #>
    param([string]$Name)

    return '"' + $Name.Replace('"', '""') + '"'
}

function ConvertTo-HexString {
    <#
    .SYNOPSIS
        Wandelt Bytes in eine Hex-Zeichenkette um (schnell unter PowerShell 7, kompatibel mit 5.1).
    #>
    param([byte[]]$Bytes)

    if ($script:HasToHexString) {
        return [Convert]::ToHexString($Bytes)
    }
    return [BitConverter]::ToString($Bytes).Replace('-', '')
}

function ConvertTo-SqliteLiteral {
    <#
    .SYNOPSIS
        Wandelt einen SQL-Server-Wert in ein SQLite-Literal um.
    #>
    param($Value, [string]$SqlType)

    if ($null -eq $Value) {
        return 'NULL'
    }
    if ($Value -is [string]) {
        if ($SqliteUnsafeChars.IsMatch($Value)) {
            # sqlite3.exe liest die SQL-Datei unter Windows im Textmodus (CRLF wird LF, Strg+Z beendet
            # die Datei). Texte mit Steuerzeichen deshalb verlustfrei als Bytefolge übergeben.
            return "CAST(X'" + (ConvertTo-HexString -Bytes ([Text.Encoding]::UTF8.GetBytes($Value))) + "' AS TEXT)"
        }
        return "'" + $Value.Replace("'", "''") + "'"
    }
    if ($Value -is [bool]) {
        if ($Value) {
            return '1'
        }
        return '0'
    }
    if ($Value -is [datetime]) {
        if ($SqlType -eq 'date') {
            return "'" + $Value.ToString('yyyy-MM-dd', $Invariant) + "'"
        }
        return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss.fff', $Invariant) + "'"
    }
    if ($Value -is [byte[]]) {
        return "X'" + (ConvertTo-HexString -Bytes $Value) + "'"
    }
    if ($Value -is [decimal] -or $Value -is [double] -or $Value -is [single]) {
        return ([double]$Value).ToString('R', $Invariant)
    }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64]) {
        return $Value.ToString($Invariant)
    }
    return "'" + [Convert]::ToString($Value, $Invariant).Replace("'", "''") + "'"
}

function ConvertTo-ColumnLetter {
    <#
    .SYNOPSIS
        Wandelt eine Spaltennummer (1-basiert) in Excel-Buchstaben um (1 -> A, 27 -> AA).
    #>
    param([int]$Number)

    $name = ''
    while ($Number -gt 0) {
        $remainder = ($Number - 1) % 26
        $name = [char](65 + $remainder) + $name
        $Number = [int][math]::Floor(($Number - 1) / 26)
    }
    return $name
}

function Format-IsoDateTime {
    <#
    .SYNOPSIS
        Datum als Text wie Python datetime.isoformat(" ") (für Werte vor 1900).
    #>
    param([datetime]$Value)

    if ($Value.Ticks % [TimeSpan]::TicksPerSecond -ne 0) {
        return $Value.ToString('yyyy-MM-dd HH:mm:ss.ffffff', $Invariant)
    }
    return $Value.ToString('yyyy-MM-dd HH:mm:ss', $Invariant)
}

function ConvertTo-XlsxCell {
    <#
    .SYNOPSIS
        Erzeugt das XML einer Excel-Zelle. Texte werden immer als Text (inlineStr) geschrieben.
    #>
    param([string]$Reference, $Value, [int]$Style = 0)

    if ($null -eq $Value) {
        return ''
    }
    $styleAttribute = ''
    if ($Style -gt 0) {
        $styleAttribute = ' s="' + $Style + '"'
    }

    if ($Value -is [datetime]) {
        if ($Value.Year -lt 1900) {
            $Value = Format-IsoDateTime -Value $Value
        }
        else {
            $serial = $Value.ToOADate()
            if ($Value -lt $Excel1900LeapBugEnd) {
                $serial -= 1
            }
            $dateStyle = 3
            if ($Value.TimeOfDay -eq [TimeSpan]::Zero) {
                $dateStyle = 2
            }
            return '<c r="{0}" s="{1}"><v>{2}</v></c>' -f $Reference, $dateStyle, $serial.ToString('R', $Invariant)
        }
    }
    if ($Value -is [bool]) {
        return '<c r="{0}"{1} t="b"><v>{2}</v></c>' -f $Reference, $styleAttribute, [int]$Value
    }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or
        $Value -is [decimal]) {
        return '<c r="{0}"{1}><v>{2}</v></c>' -f $Reference, $styleAttribute, $Value.ToString($Invariant)
    }
    if ($Value -is [double] -or $Value -is [single]) {
        return '<c r="{0}"{1}><v>{2}</v></c>' -f $Reference, $styleAttribute, ([double]$Value).ToString('R', $Invariant)
    }
    if ($Value -is [byte[]]) {
        $Value = '0x' + (ConvertTo-HexString -Bytes $Value)
    }

    $text = $IllegalXmlChars.Replace([Convert]::ToString($Value, $Invariant), '')
    if ($text.Length -gt 32767) {
        throw "Text zu lang für Excel (Zelle $Reference, $($text.Length) Zeichen)."
    }
    # CR als Zeichenreferenz, da XML-Parser ein wörtliches CR beim Lesen in LF umwandeln.
    $escaped = [Security.SecurityElement]::Escape($text).Replace("`r", '&#13;')
    return '<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f
        $Reference, $styleAttribute, $escaped
}

function New-XlsxLink {
    <#
    .SYNOPSIS
        Beschreibt eine Link-Zelle: angezeigter Text und relatives Ziel (bereits URI-kodiert).
    #>
    param([string]$Text, [string]$Target)

    return [pscustomobject]@{
        PSTypeName = 'Npr.XlsxLink'
        Text       = $Text
        Target     = $Target
    }
}

function Write-XlsxRow {
    <#
    .SYNOPSIS
        Schreibt eine Tabellenzeile in ein Arbeitsblatt. Link-Zellen werden in $Hyperlinks gesammelt.
    #>
    param($Writer, [int]$RowNumber, [object[]]$Values, [string[]]$Letters, [switch]$Bold, $Hyperlinks)

    $baseStyle = 0
    if ($Bold) {
        $baseStyle = $StyleBold
    }
    $Writer.Write('<row r="' + $RowNumber + '">')
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $reference = $Letters[$i] + $RowNumber
        $value = $Values[$i]
        if ($value -is [System.Management.Automation.PSCustomObject] -and $value.PSObject.TypeNames -contains 'Npr.XlsxLink') {
            $Hyperlinks.Add([pscustomobject]@{
                    Reference = $reference
                    Target    = $value.Target
                })
            $Writer.Write((ConvertTo-XlsxCell -Reference $reference -Value $value.Text -Style $StyleLink))
        }
        else {
            $Writer.Write((ConvertTo-XlsxCell -Reference $reference -Value $value -Style $baseStyle))
        }
    }
    $Writer.Write('</row>')
}

function ConvertTo-RelativeUri {
    <#
    .SYNOPSIS
        Relativer, URI-kodierter Link aus Pfadsegmenten (für Excel-Hyperlinks).
    #>
    param([string[]]$Segments)

    return (@($Segments | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/')
}

function Write-ZipTextEntry {
    <#
    .SYNOPSIS
        Schreibt einen Text als Eintrag in das xlsx-Archiv.
    #>
    param($Archive, [string]$EntryName, [string]$Content)

    $entry = $Archive.CreateEntry($EntryName, [IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter -ArgumentList $entry.Open(), $Utf8NoBom
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
    }
}

function Open-XlsxSheet {
    <#
    .SYNOPSIS
        Legt ein Arbeitsblatt im Archiv an und schreibt den Kopf bis zum Beginn der Daten.
    #>
    param($Archive, [string]$EntryName, [switch]$Frozen, [switch]$Selected)

    $entry = $Archive.CreateEntry($EntryName, [IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter -ArgumentList $entry.Open(), $Utf8NoBom
    $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    $writer.Write("<worksheet xmlns=`"$NsMain`" xmlns:r=`"$NsRel`">")

    $selectedAttribute = ''
    if ($Selected) {
        $selectedAttribute = ' tabSelected="1"'
    }
    if ($Frozen) {
        $writer.Write("<sheetViews><sheetView workbookViewId=`"0`"$selectedAttribute>" +
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
    }
    else {
        $writer.Write("<sheetViews><sheetView workbookViewId=`"0`"$selectedAttribute/></sheetViews>")
    }
    $writer.Write('<sheetData>')
    return $writer
}

function Close-XlsxSheet {
    <#
    .SYNOPSIS
        Schließt ein Arbeitsblatt ab, optional mit Autofilter.
    #>
    param($Writer, [string]$AutoFilterReference, $Hyperlinks)

    $Writer.Write('</sheetData>')
    if ($AutoFilterReference) {
        $Writer.Write('<autoFilter ref="' + $AutoFilterReference + '"/>')
    }
    if ($Hyperlinks -and $Hyperlinks.Count -gt 0) {
        $Writer.Write('<hyperlinks>')
        for ($i = 0; $i -lt $Hyperlinks.Count; $i++) {
            $Writer.Write('<hyperlink ref="' + $Hyperlinks[$i].Reference + '" r:id="rId' + ($i + 1) + '"/>')
        }
        $Writer.Write('</hyperlinks>')
    }
    $Writer.Write('</worksheet>')
    $Writer.Dispose()
}

function Write-XlsxSheetRelationships {
    <#
    .SYNOPSIS
        Schreibt die Beziehungsdatei eines Arbeitsblatts mit den externen (relativen) Link-Zielen.
    #>
    param($Archive, [int]$SheetNumber, $Hyperlinks)

    if (-not $Hyperlinks -or $Hyperlinks.Count -eq 0) {
        return
    }
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$builder.Append("<Relationships xmlns=`"$NsPkgRel`">")
    for ($i = 0; $i -lt $Hyperlinks.Count; $i++) {
        [void]$builder.Append(('<Relationship Id="rId{0}" Type="{1}/hyperlink" Target="{2}" TargetMode="External"/>' -f
                ($i + 1), $NsRel, [Security.SecurityElement]::Escape($Hyperlinks[$i].Target)))
    }
    [void]$builder.Append('</Relationships>')
    Write-ZipTextEntry -Archive $Archive -EntryName "xl/worksheets/_rels/sheet$SheetNumber.xml.rels" -Content $builder.ToString()
}

function Get-SheetHeader {
    <#
    .SYNOPSIS
        Überschrift für Fremdschlüsselspalten: Präfix und "_ID" entfernen.
    #>
    param([string]$Column)

    return $Column -replace '^(tblxx_|loktbl_|tbl_)|_ID$', ''
}

function Get-LabelMap {
    <#
    .SYNOPSIS
        Liefert für eine Nachschlagetabelle die Zuordnung ID -> Klartext (mit Zwischenspeicher).
    #>
    param([string]$Table)

    if ($labelMapCache.ContainsKey($Table)) {
        return $labelMapCache[$Table]
    }

    $names = [string[]]@($columns[$Table] | ForEach-Object { $_.Name })
    $map = @{}
    if ($Table -eq 'tbl_Grab') {
        $reihe = Get-LabelMap -Table 'tblxx_Reihe'
        $iReihe = Get-ColumnIndex -Names $names -Name 'tblxx_Reihe_ID' -Context 'tbl_Grab'
        $iNummer = Get-ColumnIndex -Names $names -Name 'GrabNr' -Context 'tbl_Grab'
        foreach ($row in $data[$Table]) {
            $reiheLabel = ''
            $reiheKey = ConvertTo-LookupKey -Value $row[$iReihe]
            if ($null -ne $row[$iReihe] -and $reihe.ContainsKey($reiheKey)) {
                $reiheLabel = $reihe[$reiheKey]
            }
            $map[(ConvertTo-LookupKey -Value $row[0])] = '{0} / {1}' -f $reiheLabel, $row[$iNummer]
        }
    }
    elseif ($PersonTables -contains $Table) {
        # Personen: "Nachname, Vorname" (leere Teile entfallen); ohne Namen bleibt die ID.
        $iNachname = [array]::IndexOf($names, 'Nachname')
        $iVorname = [array]::IndexOf($names, 'Vorname')
        foreach ($row in $data[$Table]) {
            $parts = @()
            foreach ($index in @($iNachname, $iVorname)) {
                if ($index -ge 0 -and $null -ne $row[$index] -and ([string]$row[$index]).Trim()) {
                    $parts += ([string]$row[$index]).Trim()
                }
            }
            $label = $row[0]
            if ($parts.Count -gt 0) {
                $label = $parts -join ', '
            }
            $map[(ConvertTo-LookupKey -Value $row[0])] = $label
        }
    }
    else {
        # Nur Textspalten gelten als Bezeichnung (z. B. ist tblxx_Anreden.Firma eine Zahl).
        $indexes = @(
            foreach ($labelColumn in $LabelColumns) {
                $index = [array]::IndexOf($names, $labelColumn)
                if ($index -ge 0 -and $CharacterTypes -contains $columns[$Table][$index].Type) {
                    $index
                }
            }
        )
        if ($indexes.Count -eq 0 -and $Table -match $ReferenceTablePattern) {
            # Referenztabelle ohne Bezeichnungsspalte: das inhaltliche Feld statt der ID anzeigen –
            # bevorzugt die nach der Tabelle benannte Textspalte (z. B. tblxx_Anreden.Anrede), sonst
            # "Schluessel", sonst die erste Textspalte.
            $tableColumns = $columns[$Table]
            $cleanName = ($Table -replace $ReferenceTablePattern, '').ToLowerInvariant()
            $textIndexes = @(
                for ($i = 0; $i -lt $tableColumns.Count; $i++) {
                    $columnName = $tableColumns[$i].Name
                    if ($CharacterTypes -contains $tableColumns[$i].Type -and $columnName -notmatch '_ID$' -and $AuditColumns -notcontains $columnName) {
                        $i
                    }
                }
            )
            $fallback = @($textIndexes | Where-Object {
                    $lower = $tableColumns[$_].Name.ToLowerInvariant()
                    $cleanName.StartsWith($lower) -or $lower.StartsWith($cleanName)
                }) + @([array]::IndexOf($names, 'Schluessel')) + $textIndexes | Where-Object { $_ -ge 0 } | Select-Object -First 1
            if ($null -ne $fallback) {
                $indexes = @($fallback)
            }
        }
        foreach ($row in $data[$Table]) {
            $label = $row[0]
            foreach ($index in $indexes) {
                if ($null -ne $row[$index] -and '' -ne $row[$index]) {
                    $label = $row[$index]
                    break
                }
            }
            $map[(ConvertTo-LookupKey -Value $row[0])] = $label
        }
    }

    $labelMapCache[$Table] = $map
    return $map
}

function New-SheetColumn {
    <#
    .SYNOPSIS
        Beschreibt eine Excel-Spalte: Überschrift, Quellspalte und optionale Aufbereitung.
    #>
    param([string]$Header, [int]$Index, $Map, [scriptblock]$Transform, $Documents, [string]$DocumentField)

    return [pscustomobject]@{
        Header        = $Header
        Index         = $Index
        Map           = $Map
        Transform     = $Transform
        Documents     = $Documents
        DocumentField = $DocumentField
    }
}

function Get-CollationOnlyMismatches {
    <#
    .SYNOPSIS
        Zählt je Tabelle Text-Fremdschlüssel, die nur ohne Groß-/Kleinschreibung bzw. ohne
        Leerzeichen am Ende zu ihrem Ziel passen (in SQL Server gültig, in SQLite nicht).
    #>
    $result = @{}
    foreach ($child in $foreignKeys.Keys) {
        $childColumns = $columns[$child]
        $childNames = [string[]]@($childColumns | ForEach-Object { $_.Name })
        foreach ($column in $foreignKeys[$child].Keys) {
            $iChild = [array]::IndexOf($childNames, $column)
            if ($iChild -lt 0 -or $childColumns[$iChild].Type -notin $CharacterTypes) {
                continue
            }
            $parent = $foreignKeys[$child][$column][0]
            $parentNames = [string[]]@($columns[$parent] | ForEach-Object { $_.Name })
            $iParent = [array]::IndexOf($parentNames, $foreignKeys[$child][$column][1])
            $exact = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::Ordinal)
            $relaxed = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
            foreach ($row in $data[$parent]) {
                if ($null -ne $row[$iParent]) {
                    $text = [string]$row[$iParent]
                    [void]$exact.Add($text)
                    [void]$relaxed.Add($text.TrimEnd())
                }
            }
            foreach ($row in $data[$child]) {
                if ($null -eq $row[$iChild]) {
                    continue
                }
                $text = [string]$row[$iChild]
                if (-not $exact.Contains($text) -and $relaxed.Contains($text.TrimEnd())) {
                    $result[$child] = 1 + [int]$result[$child]
                }
            }
        }
    }
    return $result
}

function Move-ItemAcrossVolumes {
    <#
    .SYNOPSIS
        Verschiebt Datei oder Ordner; über Laufwerksgrenzen hinweg per Kopieren und Löschen.
    #>
    param([string]$Path, [string]$Destination)

    if ([IO.Path]::GetPathRoot($Path) -eq [IO.Path]::GetPathRoot($Destination)) {
        Move-Item -LiteralPath $Path -Destination $Destination -Force
    }
    else {
        Copy-Item -LiteralPath $Path -Destination $Destination -Recurse -Force
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function Split-DocumentPath {
    <#
    .SYNOPSIS
        Teilt "Titel#Pfad" wie die Linux-Variante: Titel vor dem ersten "#", Pfad danach
        (ohne "#" sind beide der Gesamtwert).
    #>
    param([string]$Value)

    $parts = $Value.Split([char[]]'#', 2)
    return [pscustomobject]@{
        Title = $parts[0]
        Path  = $parts[$parts.Count - 1]
    }
}

function Get-ContentExtension {
    <#
    .SYNOPSIS
        Ermittelt die Dateiendung anhand der ersten Bytes; behält passende Originalendungen bei.
    #>
    param([byte[]]$Bytes, [string]$OriginalExtension)

    $original = $OriginalExtension.ToLowerInvariant()
    if ($Bytes.Length -ge 4) {
        if ($Bytes[0] -eq 0xD0 -and $Bytes[1] -eq 0xCF -and $Bytes[2] -eq 0x11 -and $Bytes[3] -eq 0xE0) {
            # Altes Office-Format (OLE): .doc, .xls, .ppt, .msg …
            if ($original -in @('.doc', '.dot', '.xls', '.ppt', '.msg')) {
                return $original
            }
            return '.doc'
        }
        if ($Bytes[0] -eq 0x50 -and $Bytes[1] -eq 0x4B -and $Bytes[2] -eq 0x03 -and $Bytes[3] -eq 0x04) {
            # ZIP-Container: Office Open XML (.docx, .xlsx …), OpenDocument oder ZIP
            if ($original -in @('.docx', '.docm', '.dotx', '.xlsx', '.pptx', '.odt', '.ods', '.zip')) {
                return $original
            }
            return '.docx'
        }
        if ($Bytes[0] -eq 0x25 -and $Bytes[1] -eq 0x50 -and $Bytes[2] -eq 0x44 -and $Bytes[3] -eq 0x46) {
            return '.pdf'
        }
    }
    if ($original) {
        return $original
    }
    return '.bin'
}

function ConvertTo-JsonText {
    <#
    .SYNOPSIS
        Text als JSON-Zeichenkette; "<", ">" und "&" werden maskiert (sicher innerhalb von <script>).
    #>
    param([string]$Text)

    return [System.Web.HttpUtility]::JavaScriptStringEncode($Text, $true)
}

function ConvertTo-JsonValue {
    <#
    .SYNOPSIS
        Wandelt einen Excel-Zellwert in JSON um (Datum als ISO-Text, Link als {f, u}).
    #>
    param($Value)

    if ($null -eq $Value) {
        return 'null'
    }
    if ($Value -is [string]) {
        return ConvertTo-JsonText -Text $Value
    }
    if ($Value -is [bool]) {
        if ($Value) {
            return 'true'
        }
        return 'false'
    }
    if ($Value -is [datetime]) {
        if ($Value.TimeOfDay -eq [TimeSpan]::Zero) {
            return '"' + $Value.ToString('yyyy-MM-dd', $Invariant) + '"'
        }
        return '"' + $Value.ToString('yyyy-MM-ddTHH:mm:ss', $Invariant) + '"'
    }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64]) {
        return $Value.ToString($Invariant)
    }
    if ($Value -is [decimal] -or $Value -is [double] -or $Value -is [single]) {
        return ([double]$Value).ToString('R', $Invariant)
    }
    if ($Value -is [byte[]]) {
        return 'null'
    }
    if ($Value -is [System.Management.Automation.PSCustomObject] -and $Value.PSObject.TypeNames -contains 'Npr.XlsxLink') {
        return '{"f":' + (ConvertTo-JsonText -Text $Value.Text) + ',"u":' + (ConvertTo-JsonText -Text $Value.Target) + '}'
    }
    return ConvertTo-JsonText -Text ([Convert]::ToString($Value, $Invariant))
}

function Get-ViewerColumns {
    <#
    .SYNOPSIS
        Spaltenbeschreibung eines Blatts für die HTML-Ansicht. Unveränderte Spalten (r) können Datum (t)
        oder Verweis (fk) sein; fehlende Verweisschlüssel werden als versteckte Spalten (x) ergänzt.
    #>
    param([string]$Table, $PlanColumns, [string[]]$ViewerTables)

    $sourceColumns = $columns[$Table]
    $sourceNames = [string[]]@($sourceColumns | ForEach-Object { $_.Name })
    $tableForeignKeys = @{}
    if ($foreignKeys.ContainsKey($Table)) {
        $tableForeignKeys = $foreignKeys[$Table]
    }

    $descriptors = New-Object 'System.Collections.Generic.List[string]'
    $rawForeignKeys = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    $linkColumns = New-Object 'System.Collections.Generic.List[int]'
    # Quellspalten, die bereits unverändert (mit ID) als eigene Spalte vorkommen
    $rawSources = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    foreach ($planColumn in $PlanColumns) {
        if ($null -eq $planColumn.Map -and $null -eq $planColumn.Transform -and $null -eq $planColumn.Documents) {
            [void]$rawSources.Add($sourceNames[$planColumn.Index])
        }
    }
    for ($k = 0; $k -lt $PlanColumns.Count; $k++) {
        $planColumn = $PlanColumns[$k]
        $sourceName = $sourceNames[$planColumn.Index]
        $parts = @('"h":' + (ConvertTo-JsonText -Text $planColumn.Header))
        $isRaw = $null -eq $planColumn.Map -and $null -eq $planColumn.Transform -and $null -eq $planColumn.Documents
        $isLinkable = $tableForeignKeys.Contains($sourceName) -and $ViewerTables -contains $tableForeignKeys[$sourceName][0]
        if ($isRaw) {
            $parts += '"r":1'
            if ($DateTypes -contains $sourceColumns[$planColumn.Index].Type) {
                $parts += '"t":"d"'
            }
            if ($isLinkable) {
                $parts += '"fk":' + (ConvertTo-JsonText -Text $tableForeignKeys[$sourceName][0])
                [void]$rawForeignKeys.Add($sourceName)
            }
        }
        elseif ($null -ne $planColumn.Map -and $isLinkable -and -not $rawSources.Contains($sourceName)) {
            # Klartext einer Referenz: Wert und ID gemeinsam ({v, i}), damit der Klartext verlinkt werden kann
            $parts += '"fk":' + (ConvertTo-JsonText -Text $tableForeignKeys[$sourceName][0])
            $parts += '"o":1'
            [void]$rawForeignKeys.Add($sourceName)
            $linkColumns.Add($k)
        }
        $descriptors.Add('{' + ($parts -join ',') + '}')
    }

    $extra = New-Object 'System.Collections.Generic.List[int]'
    foreach ($foreignKeyColumn in $tableForeignKeys.Keys) {
        $parent = $tableForeignKeys[$foreignKeyColumn][0]
        if ($ViewerTables -notcontains $parent -or $rawForeignKeys.Contains($foreignKeyColumn)) {
            continue
        }
        $descriptors.Add('{"h":' + (ConvertTo-JsonText -Text $foreignKeyColumn) + ',"r":1,"x":1,"fk":' + (ConvertTo-JsonText -Text $parent) + '}')
        $extra.Add([array]::IndexOf($sourceNames, $foreignKeyColumn))
    }

    return [pscustomobject]@{
        Descriptors = $descriptors
        Extra       = $extra.ToArray()
        LinkColumns = $linkColumns.ToArray()
    }
}

function Get-UniqueDocumentName {
    <#
    .SYNOPSIS
        Liefert einen im Dokumentordner noch freien Dateinamen; bei Kollision mit vorangestelltem Präfix.
    #>
    param([string]$Desired, [string]$Prefix, $Taken)

    if ($Taken.Add($Desired)) {
        return $Desired
    }
    $candidate = '{0}_{1}' -f $Prefix, $Desired
    $counter = 2
    while (-not $Taken.Add($candidate)) {
        $candidate = '{0}_{1}_{2}' -f $Prefix, $counter, $Desired
        $counter++
    }
    return $candidate
}

function ConvertTo-SafeFileName {
    <#
    .SYNOPSIS
        Ersetzt unter Windows unzulässige Zeichen in Dateinamen.
    #>
    param([string]$Name)

    foreach ($character in [IO.Path]::GetInvalidFileNameChars()) {
        $Name = $Name.Replace($character, '_')
    }
    return $Name.Trim().TrimEnd('.')
}

#endregion

# --- Vorbedingungen und Pfade ------------------------------------------------------------------
if ($env:OS -ne 'Windows_NT') {
    throw 'Dieses Skript läuft nur unter Windows. Für Linux siehe linux/export_fw.sh.'
}

try {
    Add-Type -AssemblyName System.Data.SqlClient
}
catch {
    Add-Type -AssemblyName System.Data
}
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Web

$viewerTemplate = Join-Path $PSScriptRoot $ViewerTemplateName
if (-not (Test-Path -LiteralPath $viewerTemplate -PathType Leaf)) {
    throw "Vorlage für die HTML-Ansicht nicht gefunden: $viewerTemplate"
}

if (-not $BackupPath) {
    $backupFolder = Join-Path $PSScriptRoot 'backup'
    $candidates = @(Get-ChildItem -LiteralPath $backupFolder -Filter '*.bak' -File -ErrorAction SilentlyContinue)
    if ($candidates.Count -ne 1) {
        throw ("Im Ordner '{0}' muss genau eine .bak-Datei liegen (gefunden: {1}). " +
            'Alternativ den Pfad mit -BackupPath angeben.') -f $backupFolder, $candidates.Count
    }
    $BackupPath = $candidates[0].FullName
}
if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
    throw "Datei nicht gefunden: $BackupPath"
}
$bak = (Resolve-Path -LiteralPath $BackupPath).ProviderPath
$bakName = [IO.Path]::GetFileName($bak)
$stem = [IO.Path]::GetFileNameWithoutExtension($bak)

$documentFolder = $null
if ($DocumentPath) {
    if (-not (Test-Path -LiteralPath $DocumentPath -PathType Container)) {
        throw "Dokumentordner nicht gefunden: $DocumentPath"
    }
    $documentFolder = (Resolve-Path -LiteralPath $DocumentPath).ProviderPath.TrimEnd('\')
}
else {
    $defaultDocumentFolder = Join-Path ([IO.Path]::GetDirectoryName($bak)) 'Dokumente'
    if (Test-Path -LiteralPath $defaultDocumentFolder -PathType Container) {
        $documentFolder = $defaultDocumentFolder
    }
}

if (-not $OutputPath) {
    $OutputPath = Join-Path $PSScriptRoot 'export'
}
$out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
[void](New-Item -ItemType Directory -Path $out -Force)

$sqlLocalDb = Resolve-Tool -Name 'SqlLocalDB' `
    -Candidates @("$env:ProgramFiles\Microsoft SQL Server\*\Tools\Binn\SqlLocalDB.exe") `
    -Hint 'Bitte Microsoft SQL Server Express LocalDB installieren (SqlLocalDB.msi).'
$sqlite3 = Resolve-Tool -Name 'sqlite3' -ExplicitPath $Sqlite3Path `
    -Candidates @(
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links\sqlite3.exe"
        "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\SQLite.SQLite_*\sqlite3.exe"
    ) `
    -Hint 'Bitte installieren, z. B. mit "winget install SQLite.SQLite", oder -Sqlite3Path angeben.'

$instance = "fwrestore_$PID"
$work = Join-Path ([IO.Path]::GetTempPath()) $instance
$dataFolder = Join-Path $work 'data'
$instanceCreated = $false
$master = $null
$connection = $null
$exitCode = 0

try {
    # --- Temporäre LocalDB-Instanz ---------------------------------------------------------------
    if (Test-Path -LiteralPath $work) {
        Remove-Item -LiteralPath $work -Recurse -Force
    }
    [void](New-Item -ItemType Directory -Path $dataFolder -Force)

    $createArguments = @('create', $instance)
    if ($LocalDbVersion) {
        $createArguments += $LocalDbVersion
    }
    [void](Invoke-Tool -FilePath $sqlLocalDb -ArgumentList $createArguments)
    $instanceCreated = $true
    [void](Invoke-Tool -FilePath $sqlLocalDb -ArgumentList @('start', $instance))

    # --- Restore -------------------------------------------------------------------------------------
    $master = Open-SqlConnection -Instance $instance -Database 'master'
    $backupParameter = @{ '@bak' = $bak }

    $header = Invoke-SqlQuery -Connection $master -Sql 'RESTORE HEADERONLY FROM DISK = @bak' -Parameters $backupParameter
    if ($header.Rows.Count -eq 0) {
        throw "Die Datei enthält keine SQL-Server-Sicherung: $bak"
    }
    $db = [string]$header.Rows[0][(Get-ColumnIndex -Names $header.Columns -Name 'DatabaseName' -Context 'RESTORE HEADERONLY')]

    $fileList = Invoke-SqlQuery -Connection $master -Sql 'RESTORE FILELISTONLY FROM DISK = @bak' -Parameters $backupParameter
    $iLogicalName = Get-ColumnIndex -Names $fileList.Columns -Name 'LogicalName' -Context 'RESTORE FILELISTONLY'
    $iFileType = Get-ColumnIndex -Names $fileList.Columns -Name 'Type' -Context 'RESTORE FILELISTONLY'
    $moves = for ($i = 0; $i -lt $fileList.Rows.Count; $i++) {
        $file = $fileList.Rows[$i]
        $extension = 'mdf'
        if ([string]$file[$iFileType] -eq 'L') {
            $extension = 'ldf'
        }
        $target = Join-Path $dataFolder ('restore_{0}.{1}' -f $i, $extension)
        "MOVE N'{0}' TO N'{1}'" -f ([string]$file[$iLogicalName]).Replace("'", "''"), $target.Replace("'", "''")
    }

    Write-Host "Restore $db aus $bak …"
    $dbQuoted = '[' + $db.Replace(']', ']]') + ']'
    Invoke-SqlNonQuery -Connection $master `
        -Sql ("RESTORE DATABASE $dbQuoted FROM DISK = @bak WITH " + (@($moves) -join ', ')) `
        -Parameters $backupParameter

    $deadline = (Get-Date).AddMinutes(5)
    while ($true) {
        $state = Invoke-SqlQuery -Connection $master -Sql 'SELECT state_desc FROM sys.databases WHERE name = @db' `
            -Parameters @{ '@db' = $db }
        if ($state.Rows.Count -gt 0 -and [string]$state.Rows[0][0] -eq 'ONLINE') {
            break
        }
        if ((Get-Date) -gt $deadline) {
            throw "Die Datenbank $db ist nach dem Restore nicht ONLINE."
        }
        Start-Sleep -Seconds 1
    }

    $connection = Open-SqlConnection -Instance $instance -Database $db

    # --- Metadaten und Daten -------------------------------------------------------------------------
    $tables = @(
        (Invoke-SqlQuery -Connection $connection -Sql 'SELECT name FROM sys.tables ORDER BY name').Rows |
            ForEach-Object { [string]$_[0] }
    )

    $columns = @{}
    $columnRows = (Invoke-SqlQuery -Connection $connection -Sql (
            'SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS ' +
            'ORDER BY TABLE_NAME, ORDINAL_POSITION')).Rows
    foreach ($row in $columnRows) {
        $tableName = [string]$row[0]
        if (-not $columns.ContainsKey($tableName)) {
            $columns[$tableName] = New-Object 'System.Collections.Generic.List[object]'
        }
        $columns[$tableName].Add([pscustomobject]@{
                Name = [string]$row[1]
                Type = ([string]$row[2]).ToLowerInvariant()
            })
    }

    $primaryKeys = @{}
    $pkRows = (Invoke-SqlQuery -Connection $connection -Sql @'
SELECT t.name, c.name
FROM sys.indexes i
JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
JOIN sys.tables t ON t.object_id = i.object_id
WHERE i.is_primary_key = 1
ORDER BY t.name, ic.key_ordinal
'@).Rows
    foreach ($row in $pkRows) {
        $tableName = [string]$row[0]
        if (-not $primaryKeys.ContainsKey($tableName)) {
            $primaryKeys[$tableName] = New-Object 'System.Collections.Generic.List[string]'
        }
        $primaryKeys[$tableName].Add([string]$row[1])
    }

    $foreignKeys = @{}
    $fkRows = (Invoke-SqlQuery -Connection $connection -Sql @'
SELECT DISTINCT
    OBJECT_NAME(fkc.parent_object_id),
    COL_NAME(fkc.parent_object_id, fkc.parent_column_id),
    OBJECT_NAME(fkc.referenced_object_id),
    COL_NAME(fkc.referenced_object_id, fkc.referenced_column_id)
FROM sys.foreign_key_columns fkc
'@).Rows
    foreach ($row in $fkRows) {
        $child = [string]$row[0]
        $column = [string]$row[1]
        $parent = [string]$row[2]
        $parentColumn = [string]$row[3]
        if ($child -eq $parent -and $column -eq $parentColumn) {
            continue
        }
        if (-not $foreignKeys.ContainsKey($child)) {
            $foreignKeys[$child] = [ordered]@{}
        }
        $foreignKeys[$child][$column] = @($parent, $parentColumn)
    }

    $missing = @($RequiredTables | Where-Object { $tables -notcontains $_ })
    if ($missing.Count -gt 0) {
        throw ('Tabellen fehlen: {0}. Die Datenbank stammt nicht aus FriedaWeb oder aus einer stark abweichenden Version.' -f
            ($missing -join ', '))
    }

    $data = @{}
    foreach ($table in $tables) {
        $data[$table] = (Invoke-SqlQuery -Connection $connection -Sql ('SELECT * FROM [' + $table.Replace(']', ']]') + ']')).Rows
    }

    # --- SQLite: alle Tabellen, Originalnamen und -werte, PK und FK deklariert ------------------------
    $sqlFile = Join-Path $work 'export.sql'
    $sqlWriter = New-Object System.IO.StreamWriter -ArgumentList $sqlFile, $false, $Utf8NoBom
    try {
        $sqlWriter.WriteLine('PRAGMA foreign_keys = OFF;')
        $sqlWriter.WriteLine('BEGIN;')
        foreach ($table in $tables) {
            $definitions = New-Object 'System.Collections.Generic.List[string]'
            foreach ($column in $columns[$table]) {
                $sqliteType = 'TEXT'
                if ($SqliteTypes.ContainsKey($column.Type)) {
                    $sqliteType = $SqliteTypes[$column.Type]
                }
                $definitions.Add((Format-SqliteName -Name $column.Name) + ' ' + $sqliteType)
            }
            if ($primaryKeys.ContainsKey($table)) {
                $definitions.Add('PRIMARY KEY (' + (@($primaryKeys[$table] | ForEach-Object { Format-SqliteName -Name $_ }) -join ', ') + ')')
            }
            if ($foreignKeys.ContainsKey($table)) {
                foreach ($column in $foreignKeys[$table].Keys) {
                    $reference = $foreignKeys[$table][$column]
                    $definitions.Add(('FOREIGN KEY ({0}) REFERENCES {1} ({2})' -f
                            (Format-SqliteName -Name $column), (Format-SqliteName -Name $reference[0]),
                            (Format-SqliteName -Name $reference[1])))
                }
            }
            $sqlWriter.WriteLine('CREATE TABLE ' + (Format-SqliteName -Name $table) + " (`n  " + ($definitions -join ",`n  ") + "`n);")

            $types = [string[]]@($columns[$table] | ForEach-Object { $_.Type })
            $prefix = 'INSERT INTO ' + (Format-SqliteName -Name $table) + ' VALUES ('
            foreach ($row in $data[$table]) {
                $literals = New-Object 'string[]' $types.Count
                for ($i = 0; $i -lt $types.Count; $i++) {
                    $literals[$i] = ConvertTo-SqliteLiteral -Value $row[$i] -SqlType $types[$i]
                }
                $sqlWriter.WriteLine($prefix + ($literals -join ',') + ');')
            }
        }
        $sqlWriter.WriteLine('COMMIT;')
    }
    finally {
        $sqlWriter.Dispose()
    }

    # sqlite3 wird im Arbeitsordner mit relativen Dateinamen aufgerufen (unabhängig von Sonderzeichen im Pfad).
    Push-Location -LiteralPath $work
    try {
        [void](Invoke-Tool -FilePath $sqlite3 -ArgumentList @('-bail', '-batch', 'export.sqlite', '.read export.sql'))

        $checkSql = @($tables | ForEach-Object { 'SELECT COUNT(*) FROM ' + (Format-SqliteName -Name $_) + ';' }) -join "`n"
        [IO.File]::WriteAllText((Join-Path $work 'check.sql'), $checkSql, $Utf8NoBom)
        $counts = @(Invoke-Tool -FilePath $sqlite3 -ArgumentList @('-batch', 'export.sqlite', '.read check.sql') |
                Where-Object { $_ -ne '' })
        if ($counts.Count -ne $tables.Count) {
            throw "SQLite-Prüfung: $($counts.Count) Zählergebnisse für $($tables.Count) Tabellen."
        }
        for ($i = 0; $i -lt $tables.Count; $i++) {
            if ([long]$counts[$i] -ne $data[$tables[$i]].Count) {
                throw ('SQLite-Prüfung: Tabelle {0} hat {1} statt {2} Zeilen.' -f $tables[$i], $counts[$i], $data[$tables[$i]].Count)
            }
        }
        $violations = @(Invoke-Tool -FilePath $sqlite3 -ArgumentList @('-batch', 'export.sqlite', 'PRAGMA foreign_key_check;') |
                Where-Object { $_ -ne '' })
        $collationNotes = New-Object 'System.Collections.Generic.List[string]'
        if ($violations.Count -gt 0) {
            # SQL Server vergleicht Textschlüssel je nach Kollation ohne Groß-/Kleinschreibung und ohne
            # Leerzeichen am Ende, SQLite exakt. Nur so erklärbare Abweichungen sind zulässig.
            $violationsByTable = @{}
            foreach ($line in $violations) {
                $violationTable = ($line -split '\|')[0]
                $violationsByTable[$violationTable] = 1 + [int]$violationsByTable[$violationTable]
            }
            $explained = Get-CollationOnlyMismatches
            $unexplained = @($violationsByTable.Keys | Where-Object { $violationsByTable[$_] -ne [int]$explained[$_] })
            if ($unexplained.Count -gt 0) {
                throw "FK-Verletzungen in SQLite:`n$($violations -join "`n")"
            }
            foreach ($violationTable in ($violationsByTable.Keys | Sort-Object)) {
                $collationNotes.Add(('{0}: {1} Verweise passen nur ohne Groß-/Kleinschreibung bzw. Leerzeichen am Ende' -f
                        $violationTable, $violationsByTable[$violationTable]))
            }
        }
    }
    finally {
        Pop-Location
    }

    # --- Dokumente: jede Datei genau einmal in einem Ordner ----------------------------------------------
    # Vorrang hat die Datei im Dokumentordner (Dateisystem); fehlt sie dort, wird die in der Datenbank
    # eingebettete Fassung geschrieben. Mehrere Einträge mit derselben Datei verweisen auf dieselbe Datei.
    # Dateien im Dokumentordner ohne Datenbankbezug werden ebenfalls in denselben Ordner übernommen.
    $documentFolderName = "${stem}_Dokumente"
    $documentWork = Join-Path $work $documentFolderName
    $documentInfo = @{}
    $documentStats = [ordered]@{
        EintraegeDateisystem = 0
        EintraegeDatenbank   = 0
        EintraegeFehlt       = 0
        Dateien              = 0
        OhneBezug            = 0
        InUnterordnern       = 0
    }
    $hasDocuments = @($DocumentTables.Keys | Where-Object { $tables -contains $_ }).Count -gt 0
    if ($hasDocuments) {
        [void](New-Item -ItemType Directory -Path $documentWork -Force)
        $takenNames = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)

        # Index des Dokumentordners: relativer Pfad und (eindeutiger) Dateiname
        $sourceByRelative = @{}
        $sourceByName = @{}
        $sourceFiles = @()
        $documentMarker = $null
        if ($documentFolder) {
            $sourceFiles = @(Get-ChildItem -LiteralPath $documentFolder -Recurse -File | Sort-Object FullName)
            foreach ($file in $sourceFiles) {
                $relative = $file.FullName.Substring($documentFolder.Length).TrimStart('\')
                $sourceByRelative[$relative.ToLowerInvariant()] = $file
                $nameKey = $file.Name.ToLowerInvariant()
                if ($sourceByName.ContainsKey($nameKey)) {
                    $sourceByName[$nameKey] = $null
                }
                else {
                    $sourceByName[$nameKey] = $file
                }
            }
            $documentMarker = '\' + (Split-Path -Path $documentFolder -Leaf).ToLowerInvariant() + '\'
        }

        # 1. Jeden Eintrag auflösen: Datei im Dateisystem, sonst eingebettete Daten
        $resolved = New-Object 'System.Collections.Generic.List[object]'
        foreach ($documentTable in ($DocumentTables.Keys | Sort-Object)) {
            if ($tables -notcontains $documentTable) {
                continue
            }
            $names = [string[]]@($columns[$documentTable] | ForEach-Object { $_.Name })
            $iId = Get-ColumnIndex -Names $names -Name 'ID' -Context $documentTable
            $iPath = Get-ColumnIndex -Names $names -Name $DocumentTables[$documentTable].PathColumn -Context $documentTable
            $iData = [array]::IndexOf($names, $DocumentDataColumn)
            $targetFolder = $DocumentTables[$documentTable].Folder
            $documentInfo[$documentTable] = @{}

            foreach ($row in $data[$documentTable]) {
                $sourcePath = ''
                if ($null -ne $row[$iPath]) {
                    $sourcePath = (Split-DocumentPath -Value ([string]$row[$iPath])).Path
                }
                $leaf = ''
                if ($sourcePath) {
                    $segments = $sourcePath -split '[\\/]'
                    $leaf = ConvertTo-SafeFileName -Name $segments[$segments.Count - 1]
                }

                $file = $null
                # Dateien eines eigenen Unterordners (z. B. Berichtsvorlagen) stammen nur aus der Datenbank.
                if ($documentFolder -and $sourcePath -and -not $targetFolder) {
                    $markerIndex = $sourcePath.ToLowerInvariant().IndexOf($documentMarker)
                    if ($markerIndex -ge 0) {
                        $relativeKey = $sourcePath.Substring($markerIndex + $documentMarker.Length).ToLowerInvariant()
                        if ($sourceByRelative.ContainsKey($relativeKey)) {
                            $file = $sourceByRelative[$relativeKey]
                        }
                    }
                    if (-not $file -and $leaf -and $sourceByName.ContainsKey($leaf.ToLowerInvariant())) {
                        $file = $sourceByName[$leaf.ToLowerInvariant()]
                    }
                }
                $bytes = $null
                if (-not $file -and $iData -ge 0 -and $row[$iData] -is [byte[]] -and $row[$iData].Length -gt 0) {
                    $bytes = $row[$iData]
                }

                $resolved.Add([pscustomobject]@{
                        Table      = $documentTable
                        Id         = ConvertTo-LookupKey -Value $row[$iId]
                        SourcePath = $sourcePath
                        Leaf       = $leaf
                        File       = $file
                        Bytes      = $bytes
                        Folder     = $targetFolder
                    })
            }
        }

        # 2. Alle Dateien des Dokumentordners genau einmal kopieren (Originalname, bei Kollision mit Präfix)
        $nameBySourceFile = @{}
        $referencedFiles = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $resolved) {
            if ($entry.File) {
                [void]$referencedFiles.Add($entry.File.FullName)
            }
        }
        foreach ($file in $sourceFiles) {
            $prefix = Split-Path -Path (Split-Path -Path $file.FullName -Parent) -Leaf
            $fileName = Get-UniqueDocumentName -Desired $file.Name -Prefix $prefix -Taken $takenNames
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $documentWork $fileName)
            $nameBySourceFile[$file.FullName.ToLowerInvariant()] = $fileName
            $documentStats.Dateien++
            if (-not $referencedFiles.Contains($file.FullName)) {
                $documentStats.OhneBezug++
            }
        }

        # 3. Fehlende Dateien aus der Datenbank schreiben; gleicher Pfad und gleicher Inhalt nur einmal
        $nameByEmbedded = @{}
        $takenByFolder = @{}
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try {
            foreach ($entry in $resolved) {
                $fileName = $null
                $source = 'fehlt'
                if ($entry.File) {
                    $fileName = $nameBySourceFile[$entry.File.FullName.ToLowerInvariant()]
                    $source = 'Dateisystem'
                    $documentStats.EintraegeDateisystem++
                }
                elseif ($null -ne $entry.Bytes) {
                    $contentKey = $entry.Folder + '|' + $entry.SourcePath.ToLowerInvariant() + '|' + (ConvertTo-HexString -Bytes $sha256.ComputeHash($entry.Bytes))
                    if ($nameByEmbedded.ContainsKey($contentKey)) {
                        $fileName = $nameByEmbedded[$contentKey]
                    }
                    else {
                        $extension = Get-ContentExtension -Bytes $entry.Bytes -OriginalExtension ([IO.Path]::GetExtension($entry.Leaf))
                        $baseName = [IO.Path]::GetFileNameWithoutExtension($entry.Leaf)
                        if (-not $baseName) {
                            $baseName = $entry.Id
                        }
                        $taken = $takenNames
                        $targetDirectory = $documentWork
                        if ($entry.Folder) {
                            if (-not $takenByFolder.ContainsKey($entry.Folder)) {
                                $takenByFolder[$entry.Folder] = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
                                [void](New-Item -ItemType Directory -Path (Join-Path $documentWork $entry.Folder) -Force)
                            }
                            $taken = $takenByFolder[$entry.Folder]
                            $targetDirectory = Join-Path $documentWork $entry.Folder
                        }
                        $fileName = Get-UniqueDocumentName -Desired ($baseName + $extension) -Prefix $entry.Id -Taken $taken
                        [IO.File]::WriteAllBytes((Join-Path $targetDirectory $fileName), $entry.Bytes)
                        if ($entry.Folder) {
                            # Relativer Name mit Unterordner (für Links)
                            $fileName = $entry.Folder + '/' + $fileName
                            $documentStats.InUnterordnern++
                        }
                        $nameByEmbedded[$contentKey] = $fileName
                        $documentStats.Dateien++
                    }
                    $source = 'Datenbank'
                    $documentStats.EintraegeDatenbank++
                }
                else {
                    $documentStats.EintraegeFehlt++
                }
                $documentInfo[$entry.Table][$entry.Id] = [pscustomobject]@{
                    FileName = $fileName
                    Source   = $source
                }
            }
        }
        finally {
            $sha256.Dispose()
        }
    }

    # --- Ergänzte Verknüpfungen (in FriedaWeb ohne Fremdschlüssel gespeichert) -----------------------
    # Erst nach SQLite, damit die SQLite-Kopie das Originalschema behält. Übernahme nur, wenn jede ID im Ziel existiert.
    $implicitNotes = New-Object 'System.Collections.Generic.List[string]'
    foreach ($implicit in $ImplicitForeignKeys) {
        $childTable = $implicit.Table
        $childColumn = $implicit.Column
        if ($tables -notcontains $childTable -or $data[$childTable].Count -eq 0) {
            continue
        }
        $iChild = [array]::IndexOf([string[]]@($columns[$childTable] | ForEach-Object { $_.Name }), $childColumn)
        if ($iChild -lt 0 -or ($foreignKeys.ContainsKey($childTable) -and $foreignKeys[$childTable].Contains($childColumn))) {
            continue
        }
        $values = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($row in $data[$childTable]) {
            if ($null -ne $row[$iChild]) {
                [void]$values.Add((ConvertTo-LookupKey -Value $row[$iChild]))
            }
        }
        if ($values.Count -eq 0) {
            continue
        }
        foreach ($target in $implicit.Targets) {
            if ($tables -notcontains $target -or $data[$target].Count -eq 0) {
                continue
            }
            $iTargetId = [array]::IndexOf([string[]]@($columns[$target] | ForEach-Object { $_.Name }), 'ID')
            if ($iTargetId -lt 0) {
                continue
            }
            $ids = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($row in $data[$target]) {
                [void]$ids.Add((ConvertTo-LookupKey -Value $row[$iTargetId]))
            }
            $missingCount = @($values | Where-Object { -not $ids.Contains($_) }).Count
            if ($missingCount -eq 0) {
                if (-not $foreignKeys.ContainsKey($childTable)) {
                    $foreignKeys[$childTable] = [ordered]@{}
                }
                $foreignKeys[$childTable][$childColumn] = @($target, 'ID')
                $implicitNotes.Add(('{0}.{1} → {2}' -f $childTable, $childColumn, $target))
                break
            }
        }
    }

    # --- Referenztabellen, auf die eine Exporttabelle verweist, als zusätzliche Blätter ------------------
    $exportTables = @($Sheets | Where-Object { $tables -contains $_.Table -and $data[$_.Table].Count -gt 0 } | ForEach-Object { $_.Table })
    $referenceTables = New-Object 'System.Collections.Generic.List[string]'
    foreach ($exportTable in $exportTables) {
        if (-not $foreignKeys.ContainsKey($exportTable)) {
            continue
        }
        foreach ($foreignKeyColumn in $foreignKeys[$exportTable].Keys) {
            $parent = $foreignKeys[$exportTable][$foreignKeyColumn][0]
            if ($parent -match $ReferenceTablePattern -and $data[$parent].Count -gt 0 -and
                -not $referenceTables.Contains($parent) -and $exportTables -notcontains $parent) {
                $referenceTables.Add($parent)
            }
        }
    }
    $usedSheetNames = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    foreach ($sheet in $Sheets) {
        [void]$usedSheetNames.Add($sheet.Name)
    }
    $allSheets = @($Sheets)
    foreach ($referenceTable in ($referenceTables | Sort-Object)) {
        # Excel erlaubt höchstens 31 Zeichen je Blattname.
        $baseName = $ReferenceSheetPrefix + ($referenceTable -replace $ReferenceTablePattern, '')
        $baseName = $baseName.Substring(0, [math]::Min(31, $baseName.Length))
        $sheetName = $baseName
        $counter = 2
        while (-not $usedSheetNames.Add($sheetName)) {
            $suffix = "_$counter"
            $sheetName = $baseName.Substring(0, [math]::Min(31 - $suffix.Length, $baseName.Length)) + $suffix
            $counter++
        }
        $allSheets += @{ Name = $sheetName; Table = $referenceTable; Optional = $true; Reference = $true }
    }

    # --- Excel: Kerntabellen, Nachschlage-IDs durch Klartext ersetzt --------------------------------
    $labelMapCache = @{}
    $plans = @()
    $maxColumns = 3
    $emptySheets = New-Object 'System.Collections.Generic.List[string]'
    foreach ($sheet in $allSheets) {
        $table = $sheet.Table
        if ($sheet.ContainsKey('Optional') -and $tables -notcontains $table) {
            continue
        }
        if ($data[$table].Count -eq 0) {
            # Leere Tabellen erhalten kein Blatt; sie bleiben in der Sicherung (bzw. SQLite) enthalten.
            if (-not $sheet.ContainsKey('Secondary')) {
                $emptySheets.Add(('{0} ({1})' -f $sheet.Name, $table))
            }
            continue
        }
        $names = [string[]]@($columns[$table] | ForEach-Object { $_.Name })
        $isDocumentTable = $DocumentTables.ContainsKey($table)
        $tableForeignKeys = @{}
        if ($foreignKeys.ContainsKey($table)) {
            $tableForeignKeys = $foreignKeys[$table]
        }

        $maps = @{}
        foreach ($column in $tableForeignKeys.Keys) {
            $maps[$column] = Get-LabelMap -Table $tableForeignKeys[$column][0]
        }
        if ($table -eq 'tbl_Grab') {
            $maps['ID'] = Get-LabelMap -Table 'tbl_Grab'
        }

        $sheetColumns = New-Object 'System.Collections.Generic.List[object]'
        for ($i = 0; $i -lt $names.Count; $i++) {
            $column = $names[$i]
            $parent = $null
            if ($tableForeignKeys.Contains($column)) {
                $parent = $tableForeignKeys[$column][0]
            }

            if ($parent -eq 'tbl_Grab' -or ($table -eq 'tbl_Grab' -and $column -eq 'ID')) {
                $sheetColumns.Add((New-SheetColumn -Header $column -Index $i))
                $grabHeader = Get-SheetHeader -Column $column
                if ($column -in @('ID', 'tbl_Grab_ID', 'Grab_ID')) {
                    $grabHeader = 'Grab'
                }
                $sheetColumns.Add((New-SheetColumn -Header $grabHeader -Index $i -Map $maps[$column]))
            }
            elseif ($PersonTables -contains $parent) {
                # Personenverweis: ID für XVERWEIS behalten, daneben der Name.
                $sheetColumns.Add((New-SheetColumn -Header $column -Index $i))
                $sheetColumns.Add((New-SheetColumn -Header (Get-SheetHeader -Column $column) -Index $i -Map $maps[$column]))
            }
            elseif ($maps.ContainsKey($column)) {
                $sheetColumns.Add((New-SheetColumn -Header (Get-SheetHeader -Column $column) -Index $i -Map $maps[$column]))
            }
            elseif ($isDocumentTable -and $column -eq $DocumentDataColumn) {
                # Binärinhalt nicht ins Excel; stattdessen Link auf die exportierte Datei.
                $iId = Get-ColumnIndex -Names $names -Name 'ID' -Context $table
                $sheetColumns.Add((New-SheetColumn -Header 'Datei' -Index $iId -Documents $documentInfo[$table] -DocumentField 'Link'))
                $sheetColumns.Add((New-SheetColumn -Header 'Dateiquelle' -Index $iId -Documents $documentInfo[$table] -DocumentField 'Source'))
            }
            elseif ($isDocumentTable -and $column -eq $DocumentTables[$table].PathColumn) {
                $sheetColumns.Add((New-SheetColumn -Header 'Titel' -Index $i -Transform {
                            param($value)
                            ([string]$value).Split([char[]]'#', 2)[0]
                        }))
                $sheetColumns.Add((New-SheetColumn -Header 'Pfad' -Index $i -Transform {
                            param($value)
                            $parts = ([string]$value).Split([char[]]'#', 2)
                            $parts[$parts.Count - 1]
                        }))
            }
            else {
                $sheetColumns.Add((New-SheetColumn -Header $column -Index $i))
            }
        }

        $plans += , [pscustomobject]@{
            Name      = $sheet.Name
            Table     = $table
            Columns   = $sheetColumns
            Secondary = $sheet.ContainsKey('Secondary')
            Reference = $sheet.ContainsKey('Reference')
        }
        $maxColumns = [math]::Max($maxColumns, $sheetColumns.Count)
    }

    $letters = New-Object 'string[]' $maxColumns
    for ($i = 0; $i -lt $maxColumns; $i++) {
        $letters[$i] = ConvertTo-ColumnLetter -Number ($i + 1)
    }

    $xlsxFile = Join-Path $work 'export.xlsx'
    $sheetCounts = New-Object 'System.Collections.Generic.List[object]'
    $definedNames = New-Object 'System.Collections.Generic.List[string]'
    $dangling = 0
    $sheetTotal = $plans.Count + 1

    # Daten der HTML-Ansicht werden parallel zu den Excel-Blättern als JSON geschrieben.
    $viewerTables = [string[]]@($plans | Where-Object { $ViewerExcludedTables -notcontains $_.Table } | ForEach-Object { $_.Table })
    $viewerJsonFile = Join-Path $work 'viewer.json'
    $viewerWriter = New-Object System.IO.StreamWriter -ArgumentList $viewerJsonFile, $false, $Utf8NoBom
    $viewerWriter.Write('{"tables":{')
    $viewerFirstTable = $true

    $fileStream = [IO.File]::Open($xlsxFile, [IO.FileMode]::Create)
    $archive = New-Object System.IO.Compression.ZipArchive -ArgumentList $fileStream, ([IO.Compression.ZipArchiveMode]::Create)
    try {
        # Datenblätter (sheet2 … sheetN); das Info-Blatt (sheet1) folgt am Ende mit den Zeilenzahlen.
        for ($s = 0; $s -lt $plans.Count; $s++) {
            $plan = $plans[$s]
            $sheetNumber = $s + 2
            $rows = $data[$plan.Table]
            $planColumns = $plan.Columns
            $columnCount = $planColumns.Count

            $hyperlinks = New-Object 'System.Collections.Generic.List[object]'
            $writer = Open-XlsxSheet -Archive $archive -EntryName "xl/worksheets/sheet$sheetNumber.xml" -Frozen
            Write-XlsxRow -Writer $writer -RowNumber 1 -Values ([object[]]@($planColumns | ForEach-Object { $_.Header })) -Letters $letters -Bold

            $viewerColumns = $null
            if ($viewerTables -contains $plan.Table) {
                $viewerColumns = Get-ViewerColumns -Table $plan.Table -PlanColumns $planColumns -ViewerTables $viewerTables
                if (-not $viewerFirstTable) {
                    $viewerWriter.Write(',')
                }
                $viewerFirstTable = $false
                $viewerWriter.Write((ConvertTo-JsonText -Text $plan.Table) + ':{"name":' + (ConvertTo-JsonText -Text $plan.Name) +
                    ',"cols":[' + ($viewerColumns.Descriptors -join ',') + '],"rows":[')
            }

            $rowNumber = 1
            foreach ($row in $rows) {
                $rowNumber++
                $values = New-Object 'object[]' $columnCount
                for ($k = 0; $k -lt $columnCount; $k++) {
                    $sheetColumn = $planColumns[$k]
                    $value = $row[$sheetColumn.Index]
                    if ($null -ne $value) {
                        if ($null -ne $sheetColumn.Documents) {
                            $document = $sheetColumn.Documents[(ConvertTo-LookupKey -Value $value)]
                            $value = $null
                            if ($null -ne $document) {
                                if ($sheetColumn.DocumentField -eq 'Source') {
                                    $value = $document.Source
                                }
                                elseif ($document.FileName) {
                                    $value = New-XlsxLink -Text $document.FileName `
                                        -Target (ConvertTo-RelativeUri -Segments (@($documentFolderName) + $document.FileName.Split('/')))
                                }
                            }
                        }
                        elseif ($null -ne $sheetColumn.Transform) {
                            $value = & $sheetColumn.Transform $value
                        }
                        elseif ($null -ne $sheetColumn.Map) {
                            $key = ConvertTo-LookupKey -Value $value
                            if ($sheetColumn.Map.ContainsKey($key)) {
                                $value = $sheetColumn.Map[$key]
                            }
                            else {
                                $value = "?$value"
                                $dangling++
                            }
                        }
                    }
                    $values[$k] = $value
                }
                Write-XlsxRow -Writer $writer -RowNumber $rowNumber -Values $values -Letters $letters -Hyperlinks $hyperlinks

                if ($null -ne $viewerColumns) {
                    $extraCount = $viewerColumns.Extra.Count
                    $jsonValues = New-Object 'string[]' ($columnCount + $extraCount)
                    for ($k = 0; $k -lt $columnCount; $k++) {
                        $jsonValues[$k] = ConvertTo-JsonValue -Value $values[$k]
                    }
                    foreach ($k in $viewerColumns.LinkColumns) {
                        $rawValue = $row[$planColumns[$k].Index]
                        if ($null -ne $rawValue) {
                            $jsonValues[$k] = '{"v":' + $jsonValues[$k] + ',"i":' + (ConvertTo-JsonValue -Value $rawValue) + '}'
                        }
                    }
                    for ($k = 0; $k -lt $extraCount; $k++) {
                        $jsonValues[$columnCount + $k] = ConvertTo-JsonValue -Value $row[$viewerColumns.Extra[$k]]
                    }
                    if ($rowNumber -gt 2) {
                        $viewerWriter.Write(',')
                    }
                    $viewerWriter.Write('[' + ($jsonValues -join ',') + ']')
                }
            }
            if ($null -ne $viewerColumns) {
                $viewerWriter.Write(']}')
            }

            $lastLetter = $letters[$columnCount - 1]
            $lastRow = $rows.Count + 1
            Close-XlsxSheet -Writer $writer -AutoFilterReference ('A1:{0}{1}' -f $lastLetter, $lastRow) -Hyperlinks $hyperlinks
            Write-XlsxSheetRelationships -Archive $archive -SheetNumber $sheetNumber -Hyperlinks $hyperlinks

            $quotedName = "'" + $plan.Name.Replace("'", "''") + "'"
            $definedNames.Add(('<definedName name="_xlnm._FilterDatabase" localSheetId="{0}" hidden="1">{1}</definedName>' -f
                    ($sheetNumber - 1), [Security.SecurityElement]::Escape(('{0}!$A$1:${1}${2}' -f $quotedName, $lastLetter, $lastRow))))
            $sheetCounts.Add(@($plan.Name, $plan.Table, $rows.Count))
        }

        # Info-Blatt
        $friedhofRows = (Invoke-SqlQuery -Connection $connection -Sql 'SELECT Bezeichnung, Ort FROM tblxx_Friedhof').Rows
        $friedhof = @($friedhofRows | ForEach-Object { ('{0} {1}' -f $_[0], $_[1]).Trim() }) -join ', '
        # Letzter Sterbefall aus allen vorhandenen Personentabellen mit Spalte Sterbedatum.
        $lastDeath = $null
        foreach ($personTable in $PersonTables) {
            if ($tables -notcontains $personTable -or @($columns[$personTable] | Where-Object { $_.Name -eq 'Sterbedatum' }).Count -eq 0) {
                continue
            }
            $candidate = (Invoke-SqlQuery -Connection $connection -Sql "SELECT MAX(Sterbedatum) FROM [$personTable]").Rows[0][0]
            if ($null -ne $candidate -and ($null -eq $lastDeath -or $candidate -gt $lastDeath)) {
                $lastDeath = $candidate
            }
        }
        $lastChange = (Invoke-SqlQuery -Connection $connection -Sql 'SELECT MAX(Datum) FROM tbl_Verlauf').Rows[0][0]

        $writer = Open-XlsxSheet -Archive $archive -EntryName 'xl/worksheets/sheet1.xml' -Selected
        $infoRows = @(
            , @{ Bold = $true; Values = @("Export der Friedhofsverwaltung FriedaWeb – $friedhof") }
            , @{ Bold = $false; Values = @('Quelle', "MS-SQL-Backup $bakName (Datenbank $db)") }
            , @{ Bold = $false; Values = @('Letzter Sterbefall', $lastDeath) }
            , @{ Bold = $false; Values = @('Letzte Änderung', $lastChange) }
            , @{ Bold = $false; Values = @('Vollständige Rohdaten', $(if ($KeepSqlite) {
                            "$stem.sqlite enthält alle $($tables.Count) Tabellen unverändert mit IDs und Fremdschlüsseln"
                        }
                        else {
                            "Die Original-Sicherung $bakName (alle $($tables.Count) Tabellen) samt Dokumentordner aufbewahren; Excel und HTML enthalten die fachlich relevanten Daten."
                        })) }
            , @{ Bold = $false; Values = @('Hinweis', ("Nachschlage-IDs (Grabart, Reihe, Konfession, Firma …) sind hier durch ihre Bezeichnung ersetzt. " +
                        "Spalte 'Grab' = Reihe / Grabnummer; tbl_Grab_ID bzw. Grab_ID verweist auf die ID im Blatt Gräber.")) }
        )
        if ($hasDocuments) {
            $documentText = ("Ordner {0}: {1} Dateien, davon {2} ohne Datenbankbezug. Einträge: {3} aus dem Dateisystem, " +
                "{4} aus der Datenbank (im Dateisystem nicht vorhanden), {5} nicht gefunden. " +
                "Die Links in der Spalte 'Datei' sind relativ; Excel-Datei und Ordner gemeinsam ablegen.") -f
                $documentFolderName, $documentStats.Dateien, $documentStats.OhneBezug, $documentStats.EintraegeDateisystem,
                $documentStats.EintraegeDatenbank, $documentStats.EintraegeFehlt
            if ($documentStats.InUnterordnern -gt 0) {
                $folders = @($DocumentTables.Values | Where-Object { $_.Folder } | ForEach-Object { $_.Folder }) -join ', '
                $documentText += " Davon {0} Dateien in eigenen Unterordnern ({1})." -f $documentStats.InUnterordnern, $folders
            }
            $infoRows += , @{ Bold = $false; Values = @('Dokumente', $documentText) }
        }
        foreach ($note in $collationNotes) {
            $infoRows += , @{ Bold = $false; Values = @('Hinweis SQLite', "$note (in SQL Server gültig, Werte unverändert übernommen)") }
        }
        if ($emptySheets.Count -gt 0) {
            $infoRows += , @{ Bold = $false; Values = @('Ohne Blatt (leer)', (($emptySheets -join ', ') + ' – in dieser Sicherung ohne Daten')) }
        }
        if ($implicitNotes.Count -gt 0) {
            $infoRows += , @{ Bold = $false; Values = @('Ergänzte Verknüpfungen', (($implicitNotes -join ', ') + ' – in FriedaWeb ohne Fremdschlüssel gespeichert, jede ID im Ziel vorhanden')) }
        }
        if ($referenceTables.Count -gt 0) {
            $infoRows += , @{ Bold = $false; Values = @('Referenztabellen', "Blätter '$ReferenceSheetPrefix…' enthalten die verwendeten Nachschlagetabellen mit allen Feldern (z. B. Firmenadressen, Gebührenbeträge).") }
        }

        # Metadaten der HTML-Ansicht (Info-Zeilen wie im Excel, Datumswerte formatiert)
        $viewerInfo = New-Object 'System.Collections.Generic.List[string]'
        foreach ($infoRow in $infoRows) {
            $infoValues = @($infoRow.Values)
            if ($infoValues.Count -ne 2) {
                continue
            }
            $infoValue = $infoValues[1]
            if ($infoValue -is [datetime]) {
                $infoValue = $infoValue.ToString('dd.MM.yyyy', $Invariant)
            }
            $viewerInfo.Add('[' + (ConvertTo-JsonText -Text ([string]$infoValues[0])) + ',' + (ConvertTo-JsonText -Text ([string]$infoValue)) + ']')
        }
        $viewerPersonTables = @($PersonTables | Where-Object { $viewerTables -contains $_ } | ForEach-Object { ConvertTo-JsonText -Text $_ })
        $viewerDocumentTables = @($DocumentTables.Keys | Where-Object { $viewerTables -contains $_ } | ForEach-Object { ConvertTo-JsonText -Text $_ })
        $viewerWriter.Write('},"meta":{' +
            '"title":' + (ConvertTo-JsonText -Text "FriedaWeb Archiv – $friedhof") +
            ',"subtitle":' + (ConvertTo-JsonText -Text "Nur Lesen · Quelle $bakName") +
            ',"source":' + (ConvertTo-JsonText -Text "$bakName (Datenbank $db)") +
            ',"generated":"' + (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss', $Invariant) + '"' +
            ',"grabTable":"tbl_Grab"' +
            ',"personTables":[' + ($viewerPersonTables -join ',') + ']' +
            ',"documentTables":[' + ($viewerDocumentTables -join ',') + ']' +
            ',"secondaryTables":[' + (@($plans | Where-Object { $_.Secondary -and $viewerTables -contains $_.Table } | ForEach-Object { ConvertTo-JsonText -Text $_.Table }) -join ',') + ']' +
            ',"referenceTables":[' + (@($plans | Where-Object { $_.Reference -and $viewerTables -contains $_.Table } | ForEach-Object { ConvertTo-JsonText -Text $_.Table }) -join ',') + ']' +
            ',"info":[' + ($viewerInfo -join ',') + ']}}')
        $viewerWriter.Dispose()
        $viewerWriter = $null

        $infoRows += , @{ Bold = $false; Values = @() }
        $infoRows += , @{ Bold = $true; Values = @('Blatt', 'Quelltabelle', 'Zeilen') }
        foreach ($count in $sheetCounts) {
            $infoRows += , @{ Bold = $false; Values = $count }
        }
        for ($i = 0; $i -lt $infoRows.Count; $i++) {
            Write-XlsxRow -Writer $writer -RowNumber ($i + 1) -Values ([object[]]$infoRows[$i].Values) -Letters $letters -Bold:($infoRows[$i].Bold)
        }
        Close-XlsxSheet -Writer $writer

        # Paketstruktur
        $sheetNames = @('Info') + @($plans | ForEach-Object { $_.Name })
        $contentTypes = New-Object System.Text.StringBuilder
        [void]$contentTypes.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$contentTypes.Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
        [void]$contentTypes.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        [void]$contentTypes.Append('<Default Extension="xml" ContentType="application/xml"/>')
        [void]$contentTypes.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
        [void]$contentTypes.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
        for ($i = 1; $i -le $sheetTotal; $i++) {
            [void]$contentTypes.Append("<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
        }
        [void]$contentTypes.Append('</Types>')
        Write-ZipTextEntry -Archive $archive -EntryName '[Content_Types].xml' -Content $contentTypes.ToString()

        Write-ZipTextEntry -Archive $archive -EntryName '_rels/.rels' -Content (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            "<Relationships xmlns=`"$NsPkgRel`">" +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>')

        $workbook = New-Object System.Text.StringBuilder
        $workbookRels = New-Object System.Text.StringBuilder
        [void]$workbook.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$workbook.Append("<workbook xmlns=`"$NsMain`" xmlns:r=`"$NsRel`"><bookViews><workbookView activeTab=`"0`"/></bookViews><sheets>")
        [void]$workbookRels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$workbookRels.Append("<Relationships xmlns=`"$NsPkgRel`">")
        for ($i = 1; $i -le $sheetTotal; $i++) {
            [void]$workbook.Append(('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f [Security.SecurityElement]::Escape($sheetNames[$i - 1]), $i))
            [void]$workbookRels.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $i))
        }
        [void]$workbook.Append('</sheets><definedNames>' + ($definedNames -join '') + '</definedNames></workbook>')
        [void]$workbookRels.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' -f ($sheetTotal + 1)))
        [void]$workbookRels.Append('</Relationships>')
        Write-ZipTextEntry -Archive $archive -EntryName 'xl/workbook.xml' -Content $workbook.ToString()
        Write-ZipTextEntry -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels' -Content $workbookRels.ToString()

        # Formate: 0 Standard, 1 fett, 2 Datum, 3 Datum mit Uhrzeit, 4 Link
        Write-ZipTextEntry -Archive $archive -EntryName 'xl/styles.xml' -Content (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            "<styleSheet xmlns=`"$NsMain`">" +
            '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd.mm.yyyy"/><numFmt numFmtId="165" formatCode="dd.mm.yyyy hh:mm:ss"/></numFmts>' +
            '<fonts count="3"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
            '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
            '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/></font></fonts>' +
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="5">' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
            '</cellXfs>' +
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
            '</styleSheet>')
    }
    finally {
        $archive.Dispose()
        $fileStream.Dispose()
        if ($null -ne $viewerWriter) {
            $viewerWriter.Dispose()
        }
    }

    # --- HTML-Ansicht: Vorlage mit den Daten zusammensetzen ----------------------------------------
    $template = [IO.File]::ReadAllText($viewerTemplate, $Utf8NoBom)
    $placeholderIndex = $template.IndexOf($ViewerPlaceholder)
    if ($placeholderIndex -lt 0) {
        throw "Platzhalter $ViewerPlaceholder fehlt in der Vorlage $viewerTemplate."
    }
    $htmlFile = Join-Path $work 'export.html'
    $htmlStream = [IO.File]::Open($htmlFile, [IO.FileMode]::Create)
    try {
        $head = $Utf8NoBom.GetBytes($template.Substring(0, $placeholderIndex))
        $tail = $Utf8NoBom.GetBytes($template.Substring($placeholderIndex + $ViewerPlaceholder.Length))
        $htmlStream.Write($head, 0, $head.Length)
        $jsonStream = [IO.File]::OpenRead($viewerJsonFile)
        try {
            $jsonStream.CopyTo($htmlStream)
        }
        finally {
            $jsonStream.Dispose()
        }
        $htmlStream.Write($tail, 0, $tail.Length)
    }
    finally {
        $htmlStream.Dispose()
    }

    # --- Ergebnis in den Ausgabeordner verschieben -----------------------------------------------
    $sqlitePath = Join-Path $out "$stem.sqlite"
    $xlsxPath = Join-Path $out "$stem.xlsx"
    if ($KeepSqlite) {
        Move-Item -LiteralPath (Join-Path $work 'export.sqlite') -Destination $sqlitePath -Force
    }
    Move-Item -LiteralPath $xlsxFile -Destination $xlsxPath -Force
    $htmlPath = Join-Path $out "$stem.html"
    Move-Item -LiteralPath $htmlFile -Destination $htmlPath -Force
    if ($hasDocuments) {
        $documentPath = Join-Path $out $documentFolderName
        if (Test-Path -LiteralPath $documentPath) {
            # Nur der vom Skript erzeugte Ordner gleichen Namens wird ersetzt.
            Remove-Item -LiteralPath $documentPath -Recurse -Force
        }
        Move-ItemAcrossVolumes -Path $documentWork -Destination $documentPath
    }

    if ($KeepSqlite) {
        Write-Host "SQLite: $sqlitePath ($($tables.Count) Tabellen)"
    }
    else {
        Write-Host "Geprüft: $($tables.Count) Tabellen (SQLite-Kopie nur intern; mit -KeepSqlite behalten)"
    }
    Write-Host "Excel:  $xlsxPath"
    Write-Host "HTML:   $htmlPath (Archivansicht mit Suche, im Browser öffnen)"
    foreach ($count in $sheetCounts) {
        Write-Host ('  {0,-20} {1,6}' -f $count[0], $count[2])
    }
    if ($hasDocuments) {
        Write-Host ('Dokumente: {0} ({1} Dateien, davon {2} ohne Datenbankbezug)' -f
            $documentPath, $documentStats.Dateien, $documentStats.OhneBezug)
        Write-Host ('  Einträge: {0} aus Dateisystem, {1} aus Datenbank, {2} nicht gefunden' -f
            $documentStats.EintraegeDateisystem, $documentStats.EintraegeDatenbank, $documentStats.EintraegeFehlt)
        if (-not $documentFolder) {
            Write-Host 'HINWEIS: Kein Dokumentordner gefunden; nur eingebettete Dokumente exportiert.' -ForegroundColor Yellow
        }
        if ($documentStats.EintraegeFehlt) {
            Write-Host "WARNUNG: $($documentStats.EintraegeFehlt) Dokumente nicht gefunden (Dateiquelle 'fehlt')" -ForegroundColor Yellow
        }
    }
    foreach ($note in $collationNotes) {
        Write-Host "HINWEIS SQLite: $note (in SQL Server gültig, Werte unverändert)" -ForegroundColor Yellow
    }
    if ($dangling) {
        Write-Host "WARNUNG: $dangling Verweise ohne Nachschlagewert (als ?ID markiert)" -ForegroundColor Yellow
    }
}
catch {
    # Fehler sichtbar ausgeben und mit ExitCode 1 beenden (Aufräumen erfolgt im finally).
    Write-Error -ErrorRecord $_ -ErrorAction Continue
    $exitCode = 1
}
finally {
    # --- Aufräumen: Verbindungen, LocalDB-Instanz und Arbeitsordner --------------------------------
    foreach ($openConnection in @($connection, $master)) {
        if ($null -ne $openConnection) {
            $openConnection.Dispose()
        }
    }
    if ($instanceCreated) {
        $ErrorActionPreference = 'Continue'
        & $sqlLocalDb stop $instance -k 2>&1 | Out-Null
        & $sqlLocalDb delete $instance 2>&1 | Out-Null
        $instanceFolder = Join-Path $env:LOCALAPPDATA "Microsoft\Microsoft SQL Server Local DB\Instances\$instance"
        if (Test-Path -LiteralPath $instanceFolder) {
            Remove-Item -LiteralPath $instanceFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    if (Test-Path -LiteralPath $work) {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}
exit $exitCode
