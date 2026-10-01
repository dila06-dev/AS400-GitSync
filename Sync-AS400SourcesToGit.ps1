#requires -Version 5.1
<#
.SYNOPSIS
Exportiert IBM-i-Quellmember laut CSV und uebertraegt Aenderungen nach Git.
.DESCRIPTION
Ohne -Execute: Abfrage und Pruefung, aber kein CL-Export, Commit oder Push.
Mit -Execute: Abfrage -> CPYTOSTMF ueber ODBC -> ggf. Kopie -> Commit -> Push.
Die vorhandene Abfrage muss die konfigurierte CSV bei JEDEM Erfolg neu schreiben.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'AS400-GitSync.psd1'),
    [switch]$Execute,
    [switch]$SkipQuery,
    [switch]$NoPush,
    [switch]$AllowModifiedTargets,
    [string]$CommitMessage
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:LogFile = $null
$script:Config = $null
$script:Repo = $null
$script:Git = $null
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$connection = $null
$lock = $null
$exitCode = 0

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    if ($script:LogFile) {
        [IO.File]::AppendAllText($script:LogFile, $line + [Environment]::NewLine, $script:Utf8)
    }
}

function Quote-NativeArgument {
    param([AllowEmptyString()][string]$Value)
    # Windows CommandLineToArgvW: Backslashes vor Anfuehrungszeichen/Ende verdoppeln.
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function Invoke-Process {
    param(
        [string]$FilePath, [string[]]$Arguments, [string]$WorkingDirectory,
        [int]$TimeoutSec, [hashtable]$Environment = @{}
    )
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $FilePath
    $info.WorkingDirectory = $WorkingDirectory
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = $script:Utf8
    $info.StandardErrorEncoding = $script:Utf8
    if ($info.PSObject.Properties['ArgumentList']) {
        foreach ($argument in $Arguments) { [void]$info.ArgumentList.Add($argument) }
    } else {
        $info.Arguments = (($Arguments | ForEach-Object { Quote-NativeArgument $_ }) -join ' ')
    }
    foreach ($key in $Environment.Keys) { $info.EnvironmentVariables[$key] = [string]$Environment[$key] }
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        if (-not $process.Start()) { throw "Prozess konnte nicht gestartet werden: $FilePath" }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSec * 1000)) {
            try { $process.Kill($true) } catch { try { $process.Kill() } catch {} }
            throw "Zeitlimit ($TimeoutSec s) bei $FilePath erreicht. Remote-Arbeit kann noch laufen; Log pruefen."
        }
        $result = [pscustomobject]@{
            ExitCode = $process.ExitCode
            Out = $stdout.GetAwaiter().GetResult()
            Err = $stderr.GetAwaiter().GetResult()
        }
        return $result
    } finally { $process.Dispose() }
}

function Invoke-Git {
    param([string[]]$Arguments, [int[]]$AllowedExitCodes = @(0), [switch]$Quiet)
    $result = Invoke-Process -FilePath $script:Git -Arguments (@('-C', $script:Repo) + $Arguments) `
        -WorkingDirectory $script:Repo -TimeoutSec $script:Config.GitTimeoutSec `
        -Environment @{ GIT_TERMINAL_PROMPT = '0'; GCM_INTERACTIVE = 'Never' }
    if (-not $Quiet -and $result.Out.Trim()) { Write-Log $result.Out.TrimEnd() }
    if ($result.Err.Trim()) { Write-Log $result.Err.TrimEnd() }
    if ($AllowedExitCodes -notcontains $result.ExitCode) {
        throw "git $($Arguments[0]) fehlgeschlagen (Exitcode $($result.ExitCode)): $($result.Err.Trim()) $($result.Out.Trim())"
    }
    return $result
}

function Get-FullPath {
    param([string]$Path)
    return [IO.Path]::GetFullPath($Path).TrimEnd([char[]]@('\', '/'))
}

function Assert-NoLinks {
    param([string]$Root, [string]$RelativePath)
    $current = $Root
    foreach ($part in @('') + ($RelativePath -split '/')) {
        if ($part) { $current = Join-Path $current $part }
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Link/Junction im Quell- oder Zielpfad nicht unterstuetzt: $current"
            }
        }
    }
}

function Get-ExportPlan {
    param([string]$CsvPath)
    $encoding = [Text.Encoding]::GetEncoding($script:Config.CsvEncoding)
    $text = [IO.File]::ReadAllText($CsvPath, $encoding)
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw 'CSV ist leer. Die Abfrage muss auch ohne Treffer mindestens die Kopfzeile schreiben.'
    }
    $header = ($text -split '\r?\n', 2)[0]
    $columns = @($header.Split(';') | ForEach-Object { $_.Trim().Trim('"') })
    foreach ($required in @('SRC_LIB', 'SRC_FILE', 'SRC_MEMBER', 'ZIELPFAD', 'COMMAND')) {
        if ($columns -notcontains $required) { throw "CSV-Spalte fehlt: $required" }
    }
    $rows = @($text | ConvertFrom-Csv -Delimiter ';')
    $seen = @{}
    $root = ([string]$script:Config.IfsRoot).TrimEnd('/')
    if ($root -notmatch '^/(?:[A-Za-z0-9_@#$.-]+/)*[A-Za-z0-9_@#$.-]+$' -or $root -match '(?:^|/)\.\.?(/|$)') {
        throw 'IfsRoot muss ein absoluter IFS-Pfad ohne Leerzeichen oder .. sein.'
    }
    foreach ($row in $rows) {
        foreach ($field in @('SRC_LIB', 'SRC_FILE', 'SRC_MEMBER')) {
            if ([string]$row.$field -notmatch '^[A-Za-z_@#$][A-Za-z0-9_@#$]{0,9}$') {
                throw "Ungueltiger IBM-i-Objektname in ${field}: $($row.$field)"
            }
        }
        if ($row.PSObject.Properties['SRC_IFS_PATH'] -and -not [string]::IsNullOrWhiteSpace($row.SRC_IFS_PATH)) {
            throw "IFS-Quelle ist kein QSYS-Quellmember: $($row.SRC_IFS_PATH). Dieser Ablauf verarbeitet Member."
        }
        $target = [string]$row.ZIELPFAD
        if (-not $target.StartsWith($root + '/', [StringComparison]::Ordinal)) {
            throw "ZIELPFAD liegt nicht unter IfsRoot: $target"
        }
        $relative = $target.Substring($root.Length + 1)
        $expected = '^{0}/{1}/{2}\.[A-Za-z0-9]{{1,16}}$' -f [regex]::Escape($row.SRC_LIB), [regex]::Escape($row.SRC_FILE), [regex]::Escape($row.SRC_MEMBER)
        if ($relative -notmatch $expected) { throw "ZIELPFAD passt nicht zu LIB/FILE/MEMBER: $target" }
        $source = '/QSYS.LIB/{0}.LIB/{1}.FILE/{2}.MBR' -f $row.SRC_LIB, $row.SRC_FILE, $row.SRC_MEMBER

        # Nur CPYTOSTMF mit dem vereinbarten Parametersatz zulassen. Reihenfolge ist beliebig.
        $command = ([string]$row.COMMAND).Trim()
        $start = [regex]::Match($command, '^(?i:(?:QSYS/)?CPYTOSTMF)\b')
        if (-not $start.Success) { throw "Nur CPYTOSTMF ist erlaubt: $($row.SRC_MEMBER)" }
        $parameters = @{}
        $offset = $start.Length
        while ($offset -lt $command.Length) {
            $match = [regex]::Match($command.Substring($offset), "^\s+([A-Za-z0-9]+)\(('[^'\r\n]*'|\*?[A-Za-z0-9]+)\)")
            if (-not $match.Success) { throw "Ungueltiger COMMAND fuer $($row.SRC_MEMBER)." }
            $name = $match.Groups[1].Value.ToUpperInvariant()
            if ($parameters.ContainsKey($name)) { throw "Doppelter CL-Parameter: $name" }
            $parameters[$name] = $match.Groups[2].Value
            $offset += $match.Length
        }
        $requiredValues = @{ FROMMBR = "'$source'"; TOSTMF = "'$target'"; CVTDTA = '*AUTO'; STMFCCSID = '1208'; ENDLINFMT = '*LF'; STMFOPT = '*REPLACE' }
        if ($parameters.Count -ne $requiredValues.Count) { throw "Unerwartete CL-Parameter fuer $target" }
        foreach ($key in $requiredValues.Keys) {
            if (-not $parameters.ContainsKey($key) -or $parameters[$key] -cne $requiredValues[$key]) {
                # CL-Objekte und Schluesselwerte duerfen unterschiedliche Gross-/Kleinschreibung haben.
                if ($key -eq 'TOSTMF' -or -not $parameters.ContainsKey($key) -or $parameters[$key] -ine $requiredValues[$key]) {
                    throw "COMMAND-Parameter $key passt nicht zu CSV/Vorgaben: $target"
                }
            }
        }
        if ($seen.ContainsKey($relative)) {
            if ($seen[$relative] -cne $target) { throw "Pfadkollision durch Gross-/Kleinschreibung: $target" }
            continue
        }
        $seen[$relative] = $target
        # Qualifizierter Befehl verhindert die Aufloesung eines gleichnamigen Befehls aus *LIBL.
        $safeCommand = "QSYS/CPYTOSTMF FROMMBR('$source') TOSTMF('$target') CVTDTA(*AUTO) STMFCCSID(1208) ENDLINFMT(*LF) STMFOPT(*REPLACE)"
        [pscustomobject]@{
            RelativePath = $relative
            IfsPath = $target
            WindowsSource = Join-Path $script:Config.IfsWindowsRoot $relative.Replace('/', [IO.Path]::DirectorySeparatorChar)
            LocalTarget = Join-Path $script:Repo $relative.Replace('/', [IO.Path]::DirectorySeparatorChar)
            Command = $safeCommand
        }
    }
}

function Invoke-Query {
    $query = Get-FullPath $script:Config.QueryScript
    if (-not (Test-Path -LiteralPath $query -PathType Leaf)) { throw "Abfrageskript fehlt: $query" }
    if (Test-Path -LiteralPath $script:Config.CsvPath) {
        $archive = Join-Path $script:Config.LogDirectory ('previous-{0}-{1}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N').Substring(0,8))
        Move-Item -LiteralPath $script:Config.CsvPath -Destination $archive
        Write-Log "Vorherige CSV archiviert: $archive"
    }
    $quoted = "'" + $query.Replace("'", "''") + "'"
    $bootstrap = @'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$ErrorActionPreference = 'Stop'
$Error.Clear()
$global:LASTEXITCODE = 0
try {
    & __QUERY__
    $success = $?
    if (-not $success -or $Error.Count -gt 0 -or $global:LASTEXITCODE -ne 0) { exit 1 }
    exit 0
} catch {
    [Console]::Error.WriteLine($_.ToString())
    exit 1
}
'@
    $bootstrap = $bootstrap.Replace('__QUERY__', $quoted)
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($bootstrap))
    $shellExe = $null
    foreach ($name in @('powershell.exe', 'pwsh.exe', 'pwsh')) {
        $candidate = Join-Path $PSHOME $name
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $shellExe = $candidate; break }
    }
    if (-not $shellExe) { throw 'PowerShell-Programm fuer den Abfrage-Unterprozess nicht gefunden.' }
    Write-Log "Starte Abfrage: $query"
    $result = Invoke-Process -FilePath $shellExe -Arguments @('-NoLogo', '-NoProfile', '-NonInteractive', '-EncodedCommand', $encoded) `
        -WorkingDirectory (Split-Path -Parent $query) -TimeoutSec $script:Config.QueryTimeoutSec
    if ($result.Out.Trim()) { Write-Log $result.Out.TrimEnd() }
    if ($result.Err.Trim()) { Write-Log $result.Err.TrimEnd() }
    if ($result.ExitCode -ne 0) { throw "Abfrageskript fehlgeschlagen (Exitcode $($result.ExitCode)). Kein Export/Commit/Push." }
    if (-not (Test-Path -LiteralPath $script:Config.CsvPath -PathType Leaf)) {
        throw 'Abfrage hat keine neue CSV erzeugt. Bei null Treffern bitte eine CSV mit Kopfzeile schreiben.'
    }
}

function Open-IbmConnection {
    if (-not (Test-Path -LiteralPath $script:Config.CredentialPath -PathType Leaf)) {
        throw "Credential-Datei fehlt: $($script:Config.CredentialPath). Einrichtung siehe Anleitung."
    }
    $credential = Import-Clixml -LiteralPath $script:Config.CredentialPath
    if ($credential -isnot [Management.Automation.PSCredential]) { throw 'Credential-Datei enthaelt kein PSCredential.' }
    $builder = New-Object System.Data.Odbc.OdbcConnectionStringBuilder
    foreach ($key in $script:Config.OdbcOptions.Keys) { $builder[$key] = $script:Config.OdbcOptions[$key] }
    if ($script:Config.OdbcDsn) {
        $builder['DSN'] = $script:Config.OdbcDsn
    } else {
        $builder['Driver'] = $script:Config.OdbcDriver
        $builder['System'] = $script:Config.IbmSystem
    }
    # SQL-Systemnamenskonvention: QSYS2.QCMDEXC; keine Kennwoerter in Prozessargumenten/Logs.
    $builder['NAM'] = '0'
    $builder['UID'] = $credential.UserName
    $builder['PWD'] = $credential.GetNetworkCredential().Password
    $db = New-Object System.Data.Odbc.OdbcConnection($builder.ConnectionString)
    $builder.Clear()
    try { $db.Open(); return $db } catch { $db.Dispose(); throw }
}

function Invoke-IbmCommand {
    param([System.Data.Odbc.OdbcConnection]$Connection, [string]$Command)
    $sql = $Connection.CreateCommand()
    try {
        $sql.CommandText = 'CALL QSYS2.QCMDEXC(?)'
        $sql.CommandTimeout = $script:Config.CommandTimeoutSec
        $parameter = $sql.Parameters.Add('command', [Data.Odbc.OdbcType]::VarChar, 32702)
        $parameter.Value = $Command
        [void]$sql.ExecuteNonQuery()
    } finally { $sql.Dispose() }
}

function Assert-IfsMapping {
    param([System.Data.Odbc.OdbcConnection]$Connection, $FirstItem)
    # Ein einmaliger Name beweist, dass Windows-Zugriff und IBM-i-Export auf dieselben Daten zeigen.
    $name = '.as400-sync-probe-' + [guid]::NewGuid().ToString('N') + '.tmp'
    $ifsProbe = ([string]$script:Config.IfsRoot).TrimEnd('/') + '/' + $name
    $windowsProbe = Join-Path $script:Config.IfsWindowsRoot $name
    $probeCommand = $FirstItem.Command.Replace("TOSTMF('$($FirstItem.IfsPath)')", "TOSTMF('$ifsProbe')").Replace('STMFOPT(*REPLACE)', 'STMFOPT(*NONE)')
    $created = $false
    try {
        Invoke-IbmCommand -Connection $Connection -Command $probeCommand
        $created = $true
        if (-not (Test-Path -LiteralPath $windowsProbe -PathType Leaf)) {
            throw 'IFS-Pfadzuordnung stimmt nicht: die neu erzeugte Pruefdatei ist ueber IfsWindowsRoot nicht sichtbar.'
        }
        Write-Log 'Zuordnung IfsRoot -> IfsWindowsRoot erfolgreich geprueft.'
    } finally {
        if ($created) { Invoke-IbmCommand -Connection $Connection -Command "QSYS/RMVLNK OBJLNK('$ifsProbe')" }
    }
}

try {
    $script:Config = Import-PowerShellDataFile -LiteralPath $ConfigPath
    foreach ($key in @('QueryScript','CsvPath','CsvEncoding','IbmSystem','OdbcDriver','CredentialPath','OdbcDsn','OdbcOptions','IfsRoot','IfsWindowsRoot','RepoPath','Remote','Branch','GitExe','LogDirectory','QueryTimeoutSec','CommandTimeoutSec','GitTimeoutSec')) {
        if (-not $script:Config.ContainsKey($key)) { throw "Konfiguration unvollstaendig: $key" }
    }
    foreach ($key in @('QueryTimeoutSec','CommandTimeoutSec','GitTimeoutSec')) {
        if ([int]$script:Config[$key] -lt 1 -or [int]$script:Config[$key] -gt 86400) { throw "Ungueltiger Timeout: $key" }
    }
    $script:Repo = Get-FullPath $script:Config.RepoPath
    $script:Config.IfsWindowsRoot = Get-FullPath $script:Config.IfsWindowsRoot
    $script:Config.CsvPath = Get-FullPath $script:Config.CsvPath
    $script:Config.LogDirectory = Get-FullPath $script:Config.LogDirectory
    [void][IO.Directory]::CreateDirectory($script:Config.LogDirectory)
    $script:LogFile = Join-Path $script:Config.LogDirectory ('sync-{0}-{1}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $PID)
    Write-Log "Start. Execute=$Execute, SkipQuery=$SkipQuery, NoPush=$NoPush"
    if (-not (Test-Path -LiteralPath $script:Repo -PathType Container)) { throw "Git-Verzeichnis fehlt: $script:Repo" }
    if (-not (Test-Path -LiteralPath $script:Config.IfsWindowsRoot -PathType Container)) {
        throw "IFS-Windows-Pfad nicht erreichbar: $($script:Config.IfsWindowsRoot)"
    }
    $script:Git = (Get-Command $script:Config.GitExe -CommandType Application -ErrorAction Stop).Source
    $top = (Invoke-Git -Arguments @('rev-parse', '--show-toplevel') -Quiet).Out.Trim()
    if ((Get-FullPath $top) -ine $script:Repo) { throw 'RepoPath muss auf das Hauptverzeichnis des Git-Arbeitsverzeichnisses zeigen.' }
    $gitDirectory = (Invoke-Git -Arguments @('rev-parse', '--absolute-git-dir') -Quiet).Out.Trim()
    # Exklusive Datei im Git-Verzeichnis: auch mehrere Windows-Tasks teilen dieselbe Sperre.
    try { $lock = [IO.File]::Open((Join-Path $gitDirectory 'as400-source-sync.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
    catch { throw 'Ein anderer Sync laeuft oder die Sperrdatei ist nicht beschreibbar.' }
    $branch = (Invoke-Git -Arguments @('symbolic-ref', '--short', 'HEAD') -Quiet).Out.Trim()
    if ($branch -cne $script:Config.Branch) { throw "Aktueller Branch '$branch', erwartet '$($script:Config.Branch)'." }
    [void](Invoke-Git -Arguments @('rev-parse', '--verify', 'HEAD') -Quiet)
    foreach ($marker in @('MERGE_HEAD','CHERRY_PICK_HEAD','REVERT_HEAD','rebase-apply','rebase-merge','sequencer')) {
        if (Test-Path -LiteralPath (Join-Path $gitDirectory $marker)) { throw "Git-Vorgang noch offen: $marker" }
    }
    $staged = Invoke-Git -Arguments @('diff', '--cached', '--quiet', '--exit-code') -AllowedExitCodes @(0,1) -Quiet
    if ($staged.ExitCode -eq 1) { throw 'Bereits gestagte Aenderungen vorhanden. Erst separat abschliessen oder aus dem Index nehmen.' }
    $unmerged = (Invoke-Git -Arguments @('ls-files', '--unmerged') -Quiet).Out
    if ($unmerged.Trim()) { throw 'Nicht aufgeloeste Git-Konflikte vorhanden.' }
    $remotes = ((Invoke-Git -Arguments @('remote') -Quiet).Out.Trim() -split '\r?\n')
    if (-not $NoPush -and $remotes -cnotcontains $script:Config.Remote) { throw 'Konfiguriertes Git-Remote fehlt.' }
    if (-not $SkipQuery) { Invoke-Query } else { Write-Log 'SkipQuery: vorhandene CSV wird bewusst wiederverwendet.' 'WARN' }
    $plan = @(Get-ExportPlan -CsvPath $script:Config.CsvPath)
    Write-Log "$($plan.Count) eindeutige Quellmember in der CSV."

    $tracked = @{}
    foreach ($path in ((Invoke-Git -Arguments @('ls-files', '-z') -Quiet).Out -split "`0")) {
        if ($path) { $tracked[$path] = $path }
    }
    foreach ($item in $plan) {
        Assert-NoLinks -Root $script:Config.IfsWindowsRoot -RelativePath $item.RelativePath
        Assert-NoLinks -Root $script:Repo -RelativePath $item.RelativePath
        if ($tracked.ContainsKey($item.RelativePath) -and $tracked[$item.RelativePath] -cne $item.RelativePath) {
            throw "Pfadschreibung weicht von Git ab: '$($item.RelativePath)' / '$($tracked[$item.RelativePath])'. CSV/Exportpfad angleichen."
        }
        $ignored = Invoke-Git -Arguments @('check-ignore', '--quiet', '--', $item.RelativePath) -AllowedExitCodes @(0,1) -Quiet
        if ($ignored.ExitCode -eq 0) { throw "Git ignoriert die Exportdatei: $($item.RelativePath)" }
        if (-not $AllowModifiedTargets) {
            $dirty = (Invoke-Git -Arguments @('status', '--porcelain=v1', '--untracked-files=all', '--', $item.RelativePath) -Quiet).Out
            if ($dirty.Trim()) { throw "Zieldatei bereits veraendert: $($item.RelativePath). Pruefen; ggf. mit -AllowModifiedTargets erneut starten." }
        }
        # Fehlende Unterordner im IFS ueber den vorhandenen Windows-Zugriff anlegen.
        Write-Log $item.Command
    }
    if (-not $Execute) {
        Write-Log 'Testlauf beendet. Keine CL-Befehle, Dateiexporte, Commits oder Pushes ausgefuehrt.'
    } else {
        if ($plan.Count -gt 0) {
            Write-Log "Verbinde mit IBM i ueber ODBC: $($script:Config.IbmSystem)"
            $connection = Open-IbmConnection
            Assert-IfsMapping -Connection $connection -FirstItem $plan[0]
            foreach ($item in $plan) {
                [void][IO.Directory]::CreateDirectory((Split-Path -Parent $item.WindowsSource))
                Invoke-IbmCommand -Connection $connection -Command $item.Command
                if (-not (Test-Path -LiteralPath $item.WindowsSource -PathType Leaf)) {
                    throw "Export meldet Erfolg, aber Datei ist ueber Windows nicht sichtbar: $($item.WindowsSource). Pfadzuordnung pruefen."
                }
            }
            $connection.Close(); $connection.Dispose(); $connection = $null
            # Erst nach Erfolg ALLER IBM-i-Exporte lokale Kopien anfertigen.
            foreach ($item in $plan) {
                if ((Get-FullPath $item.WindowsSource) -ine (Get-FullPath $item.LocalTarget)) {
                    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $item.LocalTarget))
                    Copy-Item -LiteralPath $item.WindowsSource -Destination $item.LocalTarget -Force
                }
                $sourceHash = (Get-FileHash -LiteralPath $item.WindowsSource -Algorithm SHA256).Hash
                $targetHash = (Get-FileHash -LiteralPath $item.LocalTarget -Algorithm SHA256).Hash
                if ($sourceHash -ne $targetHash) { throw "Kopierpruefung fehlgeschlagen: $($item.RelativePath)" }
                if ((Get-Item -LiteralPath $item.LocalTarget).Length -ge 100MB) { throw "Quelldatei >= 100 MiB, Export pruefen: $($item.RelativePath)" }
            }
            # Nochmals pruefen, ob waehrend des Exports jemand etwas gestagt hat.
            $staged = Invoke-Git -Arguments @('diff', '--cached', '--quiet', '--exit-code') -AllowedExitCodes @(0,1) -Quiet
            if ($staged.ExitCode -eq 1) { throw 'Waehrend des Exports wurden andere Aenderungen gestagt. Abbruch vor git add.' }
            foreach ($item in $plan) { [void](Invoke-Git -Arguments @('add', '--', $item.RelativePath)) }
            $hasChanges = Invoke-Git -Arguments @('diff', '--cached', '--quiet', '--exit-code') -AllowedExitCodes @(0,1) -Quiet
            if ($hasChanges.ExitCode -eq 1) {
                $allowed = @{}
                foreach ($item in $plan) { $allowed[$item.RelativePath] = $true }
                $changed = @((Invoke-Git -Arguments @('diff', '--cached', '--name-only', '-z', '--no-renames') -Quiet).Out -split "`0" | Where-Object { $_ })
                foreach ($path in $changed) { if (-not $allowed.ContainsKey($path)) { throw "Unerwartete Datei im Index: $path" } }
                if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
                    $CommitMessage = 'IBM i Quellen: {0} Datei(en), {1}' -f $changed.Count, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                }
                [void](Invoke-Git -Arguments @('commit', '-m', $CommitMessage))
                Write-Log 'Git-Commit erstellt.'
            } else { Write-Log 'Exportdateien sind unveraendert; kein neuer Commit.' }
        } else { Write-Log 'Keine Quellmember zu exportieren.' }
        if (-not $NoPush) {
            # Auch ohne neuen Commit: vorher fehlgeschlagene Pushes erneut versuchen.
            [void](Invoke-Git -Arguments @('push', $script:Config.Remote, ('HEAD:refs/heads/' + $script:Config.Branch)))
            Write-Log 'Git-Push erfolgreich.'
        } else { Write-Log 'NoPush: Push ausgelassen.' }
        Write-Log 'Ablauf erfolgreich abgeschlossen.'
    }
} catch {
    $exitCode = 1
    Write-Log $_.Exception.Message 'ERROR'
    if ($_.Exception -is [Data.Odbc.OdbcException]) {
        foreach ($errorDetail in $_.Exception.Errors) {
            Write-Log ("ODBC SQLSTATE={0}; NativeError={1}; {2}" -f $errorDetail.SQLState, $errorDetail.NativeError, $errorDetail.Message) 'ERROR'
        }
    }
} finally {
    if ($null -ne $connection) { $connection.Dispose() }
    if ($null -ne $lock) { $lock.Dispose() }
}
exit $exitCode
