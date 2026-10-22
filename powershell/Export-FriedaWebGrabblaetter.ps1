#Requires -Version 5.1
<#
.SYNOPSIS
    Erzeugt aus einer FriedaWeb-Archivansicht (.html) für jedes Grab ein Grabblatt als PDF.

.DESCRIPTION
    Die PDFs entstehen aus derselben Druckvorlage wie "Grabblatt drucken" in der Archivansicht –
    ohne aufgeklappte Bereiche: Grabinfo, Status, Fristen, Nutzungsberechtigte, Notiz, Belegung und
    Verstorbene im Grab. Wird die Druckvorlage (Export-FriedaWeb.Viewer.html) geändert, genügt ein
    erneuter Aufruf dieses Skripts; die Archivansicht wird dabei auf die aktuelle Vorlage gebracht
    und alle PDFs werden neu erstellt.

    Ablage: Ordner "<backup-name>_Grabblaetter" neben der HTML-Datei (bzw. neben dem Dokumentordner).
    Dateiname = Titel des Grabblatts, z. B. "Friedhof Erlach · Abteilung A1 · Reihe 01 · Grab 01-03.pdf".
    Kommt ein Titel mehrfach vor (z. B. ein zurückgegebenes und ein neues Grab mit gleicher Nummer),
    erhält der Name den Status bzw. die Grab-ID als Zusatz.

    Technik: Microsoft Edge läuft einmalig unsichtbar im Hintergrund und wird über das DevTools-
    Protokoll gesteuert (nur lokal, auf einem freien Port von 127.0.0.1, eigenes temporäres Profil).
    Belegung und Ruhezeit-Kennzeichnung beziehen sich auf den Tag der Erzeugung ("gedruckt …").

.PARAMETER HtmlPath
    Archivansicht (<backup-name>.html). Ohne Angabe wird die einzige Archivansicht im Ordner
    "export" neben diesem Skript verwendet.

.PARAMETER KeepTemplate
    Die Archivansicht unverändert verwenden. Ohne diesen Schalter wird sie vorher auf die aktuelle
    Vorlage Export-FriedaWeb.Viewer.html gebracht (Daten bleiben unverändert), damit PDF und
    Druck aus der Archivansicht identisch aussehen.

.EXAMPLE
    .\Export-FriedaWebGrabblaetter.ps1
    Erstellt alle Grabblätter zur Archivansicht in .\export neu.

.EXAMPLE
    .\Export-FriedaWeb.ps1 -Grabblaetter
    Export und anschließend alle Grabblätter in einem Lauf.

.NOTES
    Voraussetzungen: Windows, PowerShell 5.1 oder 7, Microsoft Edge.
    Die PDFs enthalten personenbezogene Daten (DSGVO); Ablage wie Excel und HTML behandeln.

    SPDX-License-Identifier: MIT
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$HtmlPath,

    [switch]$KeepTemplate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$TemplateName = 'Export-FriedaWeb.Viewer.html'
$Placeholder = '/*__FW_DATA__*/null'
$FolderSuffix = '_Grabblaetter'
# Ladezeit der Archivansicht (große Bestände brauchen einige Sekunden)
$LoadTimeoutSeconds = 300

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

function Update-ViewerTemplate {
    <#
    .SYNOPSIS
        Setzt die Daten einer Archivansicht in die aktuelle Vorlage ein (Daten bleiben unverändert).
    #>
    param([string]$Path, [string]$Template)

    $utf8 = New-Object System.Text.UTF8Encoding -ArgumentList $false
    $html = [IO.File]::ReadAllText($Path, $utf8)
    $marker = 'const FW = '
    $start = $html.IndexOf($marker)
    $end = -1
    if ($start -ge 0) {
        $match = [regex]::new(';\r?\n</script>').Match($html, $start)
        if ($match.Success) {
            $end = $match.Index
        }
    }
    if ($start -lt 0 -or $end -lt 0) {
        throw "Datenblock in $Path nicht gefunden; ist das eine Archivansicht aus Export-FriedaWeb.ps1?"
    }
    $json = $html.Substring($start + $marker.Length, $end - $start - $marker.Length)
    $templateText = [IO.File]::ReadAllText($Template, $utf8)
    if ($templateText.IndexOf($Placeholder) -lt 0) {
        throw "Platzhalter $Placeholder fehlt in $Template."
    }
    [IO.File]::WriteAllText($Path, $templateText.Replace($Placeholder, $json), $utf8)
}

function Get-FreePort {
    <#
    .SYNOPSIS
        Liefert einen freien TCP-Port auf 127.0.0.1.
    #>
    $listener = New-Object System.Net.Sockets.TcpListener -ArgumentList ([Net.IPAddress]::Loopback), 0
    $listener.Start()
    try {
        return $listener.LocalEndpoint.Port
    }
    finally {
        $listener.Stop()
    }
}

function Invoke-Cdp {
    <#
    .SYNOPSIS
        Sendet einen Befehl über das DevTools-Protokoll und liefert dessen Ergebnis (Ereignisse werden übersprungen).
    #>
    param([string]$Method, [hashtable]$Params = @{})

    $script:cdpId++
    $message = @{ id = $script:cdpId; method = $Method; params = $Params } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($message)
    $none = [Threading.CancellationToken]::None
    [void]$script:socket.SendAsync((New-Object 'ArraySegment[byte]' -ArgumentList (, $bytes)),
        [Net.WebSockets.WebSocketMessageType]::Text, $true, $none).GetAwaiter().GetResult()

    $buffer = New-Object 'byte[]' 65536
    while ($true) {
        $stream = New-Object System.IO.MemoryStream
        try {
            do {
                $received = $script:socket.ReceiveAsync((New-Object 'ArraySegment[byte]' -ArgumentList (, $buffer)), $none).GetAwaiter().GetResult()
                if ($received.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) {
                    throw 'Edge hat die Verbindung beendet.'
                }
                $stream.Write($buffer, 0, $received.Count)
            } while (-not $received.EndOfMessage)
            $reply = [Text.Encoding]::UTF8.GetString($stream.ToArray()) | ConvertFrom-Json
        }
        finally {
            $stream.Dispose()
        }
        if ($reply.PSObject.Properties['id'] -and $reply.id -eq $script:cdpId) {
            if ($reply.PSObject.Properties['error']) {
                throw "Edge ($Method): $($reply.error.message)"
            }
            return $reply.result
        }
    }
}

function Invoke-PageScript {
    <#
    .SYNOPSIS
        Wertet einen JavaScript-Ausdruck in der Seite aus und liefert den Wert.
    #>
    param([string]$Expression)

    $result = Invoke-Cdp -Method 'Runtime.evaluate' -Params @{ expression = $Expression; returnByValue = $true; awaitPromise = $true }
    if ($result.PSObject.Properties['exceptionDetails']) {
        throw "Fehler in der Archivansicht: $($result.exceptionDetails.text)"
    }
    if ($result.result.PSObject.Properties['value']) {
        return $result.result.value
    }
    return $null
}

# --- Pfade ------------------------------------------------------------------------------------------
if (-not $HtmlPath) {
    $exportFolder = Join-Path $PSScriptRoot 'export'
    $candidates = @(Get-ChildItem -LiteralPath $exportFolder -Filter '*.html' -File -ErrorAction SilentlyContinue)
    if ($candidates.Count -ne 1) {
        throw ("Im Ordner '{0}' muss genau eine Archivansicht (.html) liegen (gefunden: {1}). " +
            'Alternativ den Pfad mit -HtmlPath angeben.') -f $exportFolder, $candidates.Count
    }
    $HtmlPath = $candidates[0].FullName
}
if (-not (Test-Path -LiteralPath $HtmlPath -PathType Leaf)) {
    throw "Archivansicht nicht gefunden: $HtmlPath"
}
$html = (Resolve-Path -LiteralPath $HtmlPath).ProviderPath
$stem = [IO.Path]::GetFileNameWithoutExtension($html)
$target = Join-Path ([IO.Path]::GetDirectoryName($html)) ($stem + $FolderSuffix)

$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") |
    Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $edge) {
    throw 'Microsoft Edge wurde nicht gefunden; die Grabblätter werden mit Edge als PDF gedruckt.'
}

if (-not $KeepTemplate) {
    $template = Join-Path $PSScriptRoot $TemplateName
    if (-not (Test-Path -LiteralPath $template -PathType Leaf)) {
        throw "Vorlage nicht gefunden: $template"
    }
    Update-ViewerTemplate -Path $html -Template $template
    Write-Host "Archivansicht auf aktuelle Vorlage gebracht: $html"
}

# --- Edge im Hintergrund starten ----------------------------------------------------------------------
$profileFolder = Join-Path ([IO.Path]::GetTempPath()) ('fwgrabblatt_' + [guid]::NewGuid().ToString('N'))
$port = Get-FreePort
$edgeProcess = $null
$script:socket = $null
$script:cdpId = 0
$work = $target + '.neu'
$watch = [Diagnostics.Stopwatch]::StartNew()
try {
    $edgeProcess = Start-Process -FilePath $edge -PassThru -WindowStyle Hidden -ArgumentList @(
        '--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions',
        "--remote-debugging-port=$port", "--user-data-dir=`"$profileFolder`"", 'about:blank')

    $pageSocketUrl = $null
    for ($i = 0; $i -lt 60 -and -not $pageSocketUrl; $i++) {
        Start-Sleep -Milliseconds 500
        try {
            $targets = Invoke-RestMethod -Uri "http://127.0.0.1:$port/json/list" -TimeoutSec 5
            $pages = @(@($targets) | Where-Object { $_.type -eq 'page' })
            if ($pages.Count -gt 0) {
                $pageSocketUrl = $pages[0].webSocketDebuggerUrl
            }
        }
        catch {
            # Edge startet noch
        }
    }
    if (-not $pageSocketUrl) {
        throw 'Edge ließ sich nicht im Hintergrund starten (keine DevTools-Verbindung).'
    }
    $script:socket = New-Object System.Net.WebSockets.ClientWebSocket
    [void]$script:socket.ConnectAsync([Uri]$pageSocketUrl, [Threading.CancellationToken]::None).GetAwaiter().GetResult()

    # Archivansicht laden (Einstieg "Info", damit keine Trefferliste aufgebaut wird)
    $url = ([Uri]$html).AbsoluteUri + '#info'
    [void](Invoke-Cdp -Method 'Page.navigate' -Params @{ url = $url })
    $ready = $false
    for ($i = 0; $i -lt $LoadTimeoutSeconds * 2 -and -not $ready; $i++) {
        Start-Sleep -Milliseconds 500
        try {
            $ready = [bool](Invoke-PageScript -Expression "document.readyState === 'complete' && typeof window.FW_PDF === 'object' && window.FW_PDF !== null")
        }
        catch {
            # Seite lädt noch
        }
    }
    if (-not $ready) {
        throw "Die Archivansicht wurde nicht geladen oder enthält keine Gräber (Zeitlimit $LoadTimeoutSeconds s): $html"
    }

    # Windows PowerShell 5.1 liefert ein JSON-Array als ein Objekt; @() über die Variable löst es auf
    $parsed = Invoke-PageScript -Expression 'JSON.stringify(window.FW_PDF.list())' | ConvertFrom-Json
    $graves = @($parsed)
    Write-Host ("Grabblätter: {0} Gräber, Ziel {1}" -f $graves.Count, $target)

    # In einen Arbeitsordner schreiben; der bisherige Ordner wird erst am Ende ersetzt.
    if (Test-Path -LiteralPath $work) {
        Remove-Item -LiteralPath $work -Recurse -Force
    }
    [void](New-Item -ItemType Directory -Path $work)
    # Dateinamen: bei gleichem Titel behält das aktuelle Grab den Titel, die übrigen erhalten den Status
    # (bzw. zusätzlich die Grab-ID) als Zusatz.
    $taken = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    $names = @{}
    foreach ($grave in (@($graves | Where-Object { $_.status -eq 'aktuell' }) + @($graves | Where-Object { $_.status -ne 'aktuell' }))) {
        $name = ConvertTo-SafeFileName -Name $grave.title
        if (-not $taken.Add($name)) {
            $name = ConvertTo-SafeFileName -Name ('{0} ({1})' -f $grave.title, $grave.status)
            if (-not $taken.Add($name)) {
                $name = ConvertTo-SafeFileName -Name ('{0} (ID {1})' -f $grave.title, $grave.id)
                [void]$taken.Add($name)
            }
        }
        $names[$grave.id] = $name
    }
    $printOptions = @{ printBackground = $true; preferCSSPageSize = $true; displayHeaderFooter = $false }
    $count = 0
    foreach ($grave in $graves) {
        $name = $names[$grave.id]
        $idLiteral = $grave.id | ConvertTo-Json -Compress
        if (-not (Invoke-PageScript -Expression "window.FW_PDF.render($idLiteral)")) {
            throw "Grab $($grave.id) nicht gefunden."
        }
        $pdf = Invoke-Cdp -Method 'Page.printToPDF' -Params $printOptions
        [IO.File]::WriteAllBytes((Join-Path $work ($name + '.pdf')), [Convert]::FromBase64String($pdf.data))
        $count++
        if ($count % 50 -eq 0 -or $count -eq $graves.Count) {
            Write-Progress -Activity 'Grabblätter als PDF' -Status ('{0} von {1}' -f $count, $graves.Count) -PercentComplete ([int](100 * $count / [math]::Max(1, $graves.Count)))
        }
        if ($count % 500 -eq 0) {
            Write-Host ("  {0} von {1} erstellt ({2:N0} s)" -f $count, $graves.Count, $watch.Elapsed.TotalSeconds)
        }
    }
    Write-Progress -Activity 'Grabblätter als PDF' -Completed

    # Nur der von diesem Skript erzeugte Ordner gleichen Namens wird ersetzt.
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
    Move-Item -LiteralPath $work -Destination $target
    Write-Host ("Grabblätter: {0} ({1} PDF-Dateien, {2:N0} s)" -f $target, $count, $watch.Elapsed.TotalSeconds)
}
finally {
    if ($null -ne $script:socket) {
        try {
            [void](Invoke-Cdp -Method 'Browser.close')
        }
        catch {
            # Edge ist bereits beendet
        }
        $script:socket.Dispose()
    }
    # Edge übergibt an weitere Prozesse; alle mit dem temporären Profil beenden (das normale Edge bleibt unberührt).
    if ($null -ne $edgeProcess) {
        for ($i = 0; $i -lt 10; $i++) {
            $leftover = @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'msedge.exe'" |
                    Where-Object { $_.CommandLine -and $_.CommandLine.Contains($profileFolder) })
            if ($leftover.Count -eq 0) {
                break
            }
            if ($i -gt 0) {
                $leftover | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
            }
            Start-Sleep -Milliseconds 500
        }
    }
    foreach ($folder in @($work, $profileFolder)) {
        if ((Test-Path -LiteralPath $folder) -and $folder -ne $target) {
            Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
