#Requires -Version 5.1
<#
.SYNOPSIS
    Erzeugt die Musteransicht der FriedaWeb-Archivansicht mit frei erfundenen Daten.

.DESCRIPTION
    Setzt einen kleinen, erfundenen Datenbestand (Musterfriedhof, Abteilung B2, Reihe 03, vier
    Gräber) in die Vorlage Export-FriedaWeb.Viewer.html ein. Das Ergebnis zeigt Grabauswahl,
    Grabinfo mit Belegung und Notiz, Verstorbene im Grab und den Druck – ohne echte Personendaten.
    Gedacht zum Vorführen, für Screenshots und Dokumentation.

    Alle Namen, Daten, Grabnummern und der Friedhof sind frei erfunden.

    Nach Änderungen an der Vorlage das Skript erneut ausführen, damit die Musteransicht aktuell bleibt.

.PARAMETER TemplatePath
    Vorlage der Archivansicht. Standard: Export-FriedaWeb.Viewer.html neben diesem Skript.

.PARAMETER OutputPath
    Ziel der Musteransicht. Standard: muster\FW_Muster.html neben diesem Skript.

.PARAMETER Screenshot
    Erzeugt zusätzlich mit Microsoft Edge im Hintergrund: muster\FW_Muster_Grab.png (Grabseite, 1600 × 1000),
    muster\FW_Muster_Grabblatt.pdf (Grabblatt wie mit Export-FriedaWeb.ps1 -Grabblaetter) und dessen Vorschau
    muster\FW_Muster_Grabblatt.png.

.EXAMPLE
    .\New-FriedaWebMuster.ps1 -Screenshot

.NOTES
    SPDX-License-Identifier: MIT
#>
[CmdletBinding()]
param(
    [string]$TemplatePath = (Join-Path $PSScriptRoot 'Export-FriedaWeb.Viewer.html'),
    [string]$OutputPath = (Join-Path $PSScriptRoot 'muster\FW_Muster.html'),
    [switch]$Screenshot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Spaltenbeschreibung wie in Export-FriedaWeb.ps1 (Get-ViewerColumns): r = Rohwert, t = Datum, fk = Verweis
function Col($h, [switch]$Raw, [switch]$Date, $Fk) {
    $c = [ordered]@{ h = $h }
    if ($Raw) { $c.r = 1 }
    if ($Date) { $c.t = 'd' }
    if ($Fk) { $c.fk = $Fk }
    $c
}
$flagNames = @('E1', 'E2', 'E3', 'E4') + (1..12 | ForEach-Object { "U$_" })
function Flags([int]$e, [int]$u) { @($flagNames | ForEach-Object { if ($_ -like 'E*') { [int]$_.Substring(1) -le $e } else { [int]$_.Substring(1) -le $u } }) }

# Gräber einer Musterreihe
$graves = @(
    @{ ID = 1; Nr = '01-03'; Art = 'SEM3 (Sondergrab - Dreifach)'; Erst = '1958-01-01'; Nutz = '2033-05-31'; Ruhe = '2004-03-14'; Rueck = $null; Notiz = $null; Wichtig = $null; Stellen = @('01', '02', '03') }
    @{ ID = 2; Nr = '04-06'; Art = 'SEM3 (Sondergrab - Dreifach)'; Erst = '1960-01-01'; Nutz = '2035-02-07'; Ruhe = '2030-08-06'; Rueck = $null; Notiz = 'Grabeinfassung 2019 erneuert.'; Wichtig = 'In der Grabstelle 05 - Mitte - wurde ein Urnenschacht gesetzt'; Stellen = @('04', '05', '06') }
    @{ ID = 3; Nr = '07-08'; Art = 'SE2 (Sondergrab - Doppel)'; Erst = '1971-01-01'; Nutz = '2021-12-31'; Ruhe = '2016-09-30'; Rueck = $null; Notiz = $null; Wichtig = $null; Stellen = @('07', '08') }
    @{ ID = 4; Nr = '09-10'; Art = 'SE2 (Sondergrab - Doppel)'; Erst = '1966-01-01'; Nutz = '2024-06-30'; Ruhe = '2012-04-02'; Rueck = '2024-06-30'; Notiz = 'Vorzeitige Rückgabe durch die Nutzungsberechtigten.'; Wichtig = $null; Stellen = @('09', '10') }
)
$friedhof = 'Musterfriedhof'; $abteilung = 'Abteilung B2'; $reihe = 'Reihe 03'

$grabCols = @((Col 'ID' -Raw), (Col 'Grab'), (Col 'Friedhof'), (Col 'Block'), (Col 'Reihe'), (Col 'Grabart'), (Col 'GrabNr' -Raw),
    (Col 'Ersterwerb' -Raw -Date), (Col 'NutzEnde' -Raw -Date), (Col 'RuheEnde' -Raw -Date), (Col 'Rueckgabedatum' -Raw -Date),
    (Col 'AnzGrabstellen' -Raw), (Col 'Notiz' -Raw), (Col 'NotizWichtig' -Raw))
$grabRows = @($graves | ForEach-Object { , @($_.ID, "$reihe / $($_.Nr)", $friedhof, $abteilung, $reihe, $_.Art, $_.Nr, $_.Erst, $_.Nutz, $_.Ruhe, $_.Rueck, $_.Stellen.Count, $_.Notiz, $_.Wichtig) })

$stelleCols = @((Col 'ID' -Raw), (Col 'Bezeichnung' -Raw))
$stelleRows = @($graves | ForEach-Object { $_.Stellen } | ForEach-Object { , @((10 + [int]$_), $_) })

$kettCols = @((Col 'ID' -Raw), (Col 'tbl_Grab_ID' -Raw -Fk 'tbl_Grab'), (Col 'Grab'), (Col 'tbl_Grabstelle_ID' -Raw -Fk 'tbl_Grabstelle'), (Col 'Grabstelle'),
    (Col 'Nutzbeginn' -Raw -Date), (Col 'Nutzende' -Raw -Date)) + @($flagNames | ForEach-Object { Col $_ -Raw })
$kettRows = @(foreach ($g in $graves) { foreach ($st in $g.Stellen) { , (@((110 + [int]$st), $g.ID, "$reihe / $($g.Nr)", (10 + [int]$st), "$reihe / $st", $g.Erst, $null) + (Flags 1 3)) } })

# Personen (frei erfunden)
$people = @(
    @{ ID = 501; G = 1; N = 'Mustermann'; V = 'Josef'; Geb = '1890-02-11'; Tod = '1958-03-02'; Best = '1958-03-06'; Art = 'Erdbestattung'; Pos = 'E1'; Stelle = '01'; Ruhe = '1973-03-05' }
    @{ ID = 502; G = 1; N = 'Mustermann'; V = 'Rosa'; Geb = '1895-07-23'; Tod = '1989-03-10'; Best = '1989-03-14'; Art = 'Erdbestattung'; Pos = 'E1'; Stelle = '02'; Ruhe = '2004-03-14' }
    @{ ID = 511; G = 2; N = 'Beispiel'; V = 'Anton'; Geb = '1888-06-10'; Tod = '1977-01-29'; Best = '1977-02-02'; Art = 'Erdbestattung'; Pos = 'E1'; Stelle = '04'; Ruhe = '1992-02-01' }
    @{ ID = 512; G = 2; N = 'Muster'; V = 'Erwin'; Geb = '1934-03-23'; Tod = '1981-10-20'; Best = '1981-10-22'; Art = 'Erdbestattung'; Pos = 'E1'; Stelle = '05'; Ruhe = '1996-10-21' }
    @{ ID = 513; G = 2; N = 'Muster'; V = 'Karin'; Geb = '1961-01-08'; Tod = '2004-04-20'; Best = '2004-05-01'; Art = 'Urne'; Pos = 'U2'; Stelle = '05'; Ruhe = '2014-04-30' }
    @{ ID = 514; G = 2; N = 'Muster'; V = 'Erika'; Geb = '1933-04-04'; Tod = '2020-08-03'; Best = '2020-08-07'; Art = 'Urne'; Pos = 'U1'; Stelle = '05'; Ruhe = '2030-08-06' }
    @{ ID = 521; G = 3; N = 'Probst'; V = 'Hermine'; Geb = '1902-12-01'; Tod = '2001-09-27'; Best = '2001-10-01'; Art = 'Erdbestattung'; Pos = 'E1'; Stelle = '07'; Ruhe = '2016-09-30' }
    @{ ID = 531; G = 4; N = 'Fiktiv'; V = 'Walter'; Geb = '1910-05-17'; Tod = '1997-03-29'; Best = '1997-04-03'; Art = 'Erdbestattung'; Pos = 'E1'; Stelle = '09'; Ruhe = '2012-04-02' }
)
$holders = @(
    @{ ID = 601; N = 'Mustermann'; V = 'Franz'; G = 1; Von = '1989-03-11'; Bis = $null }
    @{ ID = 602; N = 'Muster'; V = 'Petra Maria'; G = 2; Von = '2020-08-04'; Bis = $null }
    @{ ID = 603; N = 'Muster'; V = 'Erika'; G = 2; Von = '1981-10-21'; Bis = '2020-08-03'; Person = 514 }
    @{ ID = 604; N = 'Probst'; V = 'Georg'; G = 3; Von = '2001-09-28'; Bis = '2021-12-31' }
    @{ ID = 605; N = 'Fiktiv'; V = 'Maria'; G = 4; Von = '1997-03-30'; Bis = '2024-06-30' }
)
$personCols = @((Col 'ID' -Raw), (Col 'Kategorie_B' -Raw), (Col 'tbl_Grab_ID' -Raw -Fk 'tbl_Grab'), (Col 'Nachname' -Raw), (Col 'Vorname' -Raw),
    (Col 'Geburtsdatum' -Raw -Date), (Col 'Sterbedatum' -Raw -Date), (Col 'Bestattungsdatum' -Raw -Date), (Col 'Bestattungsart'), (Col 'Position'), (Col 'Ort' -Raw))
$personRows = @($people | ForEach-Object { , @($_.ID, $true, $_.G, $_.N, $_.V, $_.Geb, $_.Tod, $_.Best, $_.Art, $_.Pos, 'Musterort') }) +
    @($holders | Where-Object { -not $_.ContainsKey('Person') } | ForEach-Object { , @($_.ID, $false, $null, $_.N, $_.V, $null, $null, $null, $null, $null, 'Musterort') })
$berCols = @((Col 'ID' -Raw), (Col 'tbl_Berechtigter_ID' -Raw -Fk 'tbl_Person'), (Col 'Berechtigter'), (Col 'tbl_Grab_ID' -Raw -Fk 'tbl_Grab'), (Col 'Grab'),
    (Col 'Berechtigungsbeginn' -Raw -Date), (Col 'Berechtigungsende' -Raw -Date))
$berRows = @($holders | ForEach-Object { $personId = if ($_.ContainsKey('Person')) { $_.Person } else { $_.ID }; $g = $graves | Where-Object ID -eq $_.G
        , @((700 + $_.ID), $personId, "$($_.N), $($_.V)", $_.G, "$reihe / $($g.Nr)", $_.Von, $_.Bis) })
$burialCols = @((Col 'Grab_ID' -Raw -Fk 'tbl_Grab'), (Col 'Friedhof' -Raw), (Col 'Abteilung' -Raw), (Col 'Reihe' -Raw), (Col 'GrabNr' -Raw),
    (Col 'Nachname' -Raw), (Col 'Vorname' -Raw), (Col 'Geburtsdatum' -Raw -Date), (Col 'Sterbedatum' -Raw -Date), (Col 'Bestattungsdatum' -Raw -Date),
    (Col 'Bestattungsart' -Raw), (Col 'Grabstelle' -Raw), (Col 'Position' -Raw), (Col 'Ruhezeit bis' -Raw -Date), (Col 'Person_ID' -Raw -Fk 'tbl_Person'))
$burialRows = @($people | Sort-Object G, Best | ForEach-Object { $g = $graves | Where-Object ID -eq $_.G
        , @($_.G, $friedhof, $abteilung, $reihe, $g.Nr, $_.N, $_.V, $_.Geb, $_.Tod, $_.Best, $_.Art, $_.Stelle, $_.Pos, $_.Ruhe, $_.ID) })
$verlaufCols = @((Col 'ID' -Raw), (Col 'tbl_Grab_ID' -Raw -Fk 'tbl_Grab'), (Col 'Beginn' -Raw -Date), (Col 'Ende' -Raw -Date), (Col 'Bemerkung' -Raw))
$verlaufRows = @($graves | ForEach-Object { , @((800 + $_.ID), $_.ID, $_.Erst, $_.Nutz, 'Nutzungsrecht (Musterdaten)') })

$data = [ordered]@{
    tables = [ordered]@{
        tbl_Grab                    = [ordered]@{ name = 'Gräber'; cols = $grabCols; rows = $grabRows }
        fw_VerstorbeneImGrab        = [ordered]@{ name = 'Verstorbene im Grab'; cols = $burialCols; rows = $burialRows }
        tbl_Grabstelle              = [ordered]@{ name = 'Grabstellen'; cols = $stelleCols; rows = $stelleRows }
        tblkett_Grab_Grabstelle     = [ordered]@{ name = 'Grab_Grabstelle'; cols = $kettCols; rows = $kettRows }
        tbl_Person                  = [ordered]@{ name = 'Personen'; cols = $personCols; rows = $personRows }
        tblkett_Berechtigter_Objekt = [ordered]@{ name = 'Berechtigungen'; cols = $berCols; rows = $berRows }
        tbl_GrabNutzVerlauf         = [ordered]@{ name = 'Nutzungsverlauf'; cols = $verlaufCols; rows = $verlaufRows }
    }
    meta   = [ordered]@{
        title = 'FriedaWeb Archiv – Musterfriedhof (MUSTERDATEN)'; subtitle = 'Musteransicht · alle Daten frei erfunden'
        source = 'Musterdaten (fiktiv)'; generated = '2026-01-01T00:00:00'
        grabTable = 'tbl_Grab'; burialTable = 'fw_VerstorbeneImGrab'; personTables = @('tbl_Person'); documentTables = @()
        secondaryTables = @(); referenceTables = @()
        info = @(, @('Hinweis', 'Musteransicht mit frei erfundenen Daten – Namen, Daten und Grablage sind fiktiv.'))
    }
}

$json = $data | ConvertTo-Json -Depth 12 -Compress
$utf8 = New-Object System.Text.UTF8Encoding -ArgumentList $false
$template = [IO.File]::ReadAllText($TemplatePath, $utf8)
$placeholder = '/*__FW_DATA__*/null'
if ($template.IndexOf($placeholder) -lt 0) {
    throw "Platzhalter $placeholder fehlt in $TemplatePath."
}
$folder = Split-Path -Parent $OutputPath
if ($folder -and -not (Test-Path -LiteralPath $folder)) {
    New-Item -ItemType Directory -Path $folder | Out-Null
}
[IO.File]::WriteAllText($OutputPath, $template.Replace($placeholder, $json), $utf8)
Write-Host "Musteransicht: $OutputPath"

if ($Screenshot) {
    $edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") |
        Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $edge) {
        throw 'Microsoft Edge wurde nicht gefunden; der Screenshot ist nur mit Edge möglich.'
    }
    $png = [IO.Path]::ChangeExtension($OutputPath, $null).TrimEnd('.') + '_Grab.png'
    $edgeProfile = Join-Path ([IO.Path]::GetTempPath()) ('fwmuster_' + [guid]::NewGuid().ToString('N'))
    $url = ([Uri](Resolve-Path -LiteralPath $OutputPath).ProviderPath).AbsoluteUri + '#d/tbl_Grab/2'
    try {
        & $edge --headless --disable-gpu --hide-scrollbars --window-size=1600,1000 "--screenshot=$png" `
            --virtual-time-budget=4000 "--user-data-dir=$edgeProfile" $url 2>&1 | Out-Null
        # Edge kehrt teils zurück, bevor das Bild geschrieben ist
        for ($i = 0; $i -lt 40 -and -not (Test-Path -LiteralPath $png); $i++) {
            Start-Sleep -Milliseconds 500
        }
    }
    finally {
        Remove-Item -LiteralPath $edgeProfile -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (-not (Test-Path -LiteralPath $png)) {
        throw "Screenshot wurde nicht erstellt: $png"
    }
    Write-Host "Screenshot:    $png"

    # Muster-Grabblatt (PDF wie mit -Grabblaetter) und Vorschaubild für die README
    $grabblattScript = Join-Path $PSScriptRoot 'Export-FriedaWebGrabblaetter.ps1'
    $sampleFolder = Join-Path ([IO.Path]::GetTempPath()) ('fwmuster_' + [guid]::NewGuid().ToString('N'))
    $samplePdf = [IO.Path]::ChangeExtension($OutputPath, $null).TrimEnd('.') + '_Grabblatt.pdf'
    $samplePng = [IO.Path]::ChangeExtension($samplePdf, '.png')
    try {
        [void](New-Item -ItemType Directory -Path $sampleFolder)
        $sampleHtml = Join-Path $sampleFolder 'FW_Muster.html'
        Copy-Item -LiteralPath $OutputPath -Destination $sampleHtml
        & $grabblattScript -HtmlPath $sampleHtml -KeepTemplate | Out-Null
        $pdf = Get-ChildItem -LiteralPath (Join-Path $sampleFolder 'FW_Muster_Grabblaetter') -Filter '*Grab 04-06.pdf' | Select-Object -First 1
        if (-not $pdf) {
            throw 'Muster-Grabblatt wurde nicht erstellt.'
        }
        Copy-Item -LiteralPath $pdf.FullName -Destination $samplePdf -Force
        Write-Host "Grabblatt:     $samplePdf"
        if (Test-Path -LiteralPath $samplePng) {
            Remove-Item -LiteralPath $samplePng
        }

        # Vorschaubild mit dem in Windows enthaltenen PDF-Renderer (WinRT, nur Windows PowerShell 5.1)
        $renderScript = Join-Path $sampleFolder 'render.ps1'
        Set-Content -LiteralPath $renderScript -Encoding UTF8 -Value @'
param([string]$Pdf, [string]$Png)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$methods = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
$asOperation = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
$asAction = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
function Wait-Operation($Operation, [type]$Type) { $task = $asOperation.MakeGenericMethod($Type).Invoke($null, @($Operation)); $task.Wait(); $task.Result }
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$file = Wait-Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Pdf)) ([Windows.Storage.StorageFile])
$document = Wait-Operation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
$stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
$options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
$options.DestinationWidth = 1600
$task = $asAction.Invoke($null, @($document.GetPage(0).RenderToStreamAsync($stream, $options)))
$task.Wait()
$source = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream.GetInputStreamAt(0))
$target = [IO.File]::Create($Png)
try { $source.CopyTo($target) } finally { $target.Dispose() }
'@
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $renderScript -Pdf $samplePdf -Png $samplePng 2>&1 | Out-Null
        if (Test-Path -LiteralPath $samplePng) {
            Write-Host "Vorschau:      $samplePng"
        }
        else {
            Write-Warning "Vorschaubild des Grabblatts konnte nicht erstellt werden: $samplePng"
        }
    }
    finally {
        Remove-Item -LiteralPath $sampleFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}