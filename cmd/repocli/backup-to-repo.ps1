<#
.SYNOPSIS
    Resumable, folder-by-folder backup of a large local directory tree to a
    WebDAV collection (e.g. the Radboud Data Repository) using repocli.exe.

.DESCRIPTION
    Instead of handing the whole tree to `repocli put` in one call, this script
    splits the source into "units" (sub-folders at a given depth) and uploads
    them one at a time.  Every completed unit is recorded in a state directory,
    so an interrupted run (crash, reboot, network loss, Ctrl+C) can simply be
    started again: finished units are skipped and the rest continues.

    Within a unit, `repocli put` itself skips files that already exist in the
    repository with the same size and a newer modification time, so a re-run of
    a partially uploaded unit only transfers what is missing.

    Success of a unit is checked with three signals, because `repocli put` on a
    directory exits 0 even when individual files failed:
      1. process exit code
      2. the per-unit error file (-e) must be empty
      3. the log must not contain "cannot create repo dir"

.PARAMETER Source
    Local root to back up, e.g. E:\ or E:\projectdata.

.PARAMETER Dest
    Destination directory in the repository, relative to -Root.  If the baseURL
    in .repocli.yml is the full collection URL (i.e. `repocli ls /` already
    lists the collection content), this is just the path inside the collection,
    e.g. HM_neurons.  If the baseURL is only https://webdav.data.ru.nl, name the
    collection too, e.g. dcn.DSC_626830_0003_227/HM_neurons.

.PARAMETER Root
    Optional folder in the repository to place -Dest under.  Empty by default,
    so a destination stands on its own and a dataset lands in the same place
    whichever drive it was read from -- the repository mirrors the experiment,
    not the disks it happens to be spread over.  Set it only to park a copy
    away from the main tree, e.g. -Root scratch.  The recorded state is per
    root+destination, so aiming a finished backup at a different root makes the
    next run re-check every file instead of skipping it.

.PARAMETER Depth
    How deep to split the tree into upload units. 1 = one unit per top-level
    sub-folder (default). Use 2 if a top-level folder is itself many TB.

.PARAMETER Threads
    Passed to repocli as -n: how many files are uploaded at the same time
    within a unit.  Default 1: one file at a time, the safest setting against
    WriteStream 405 errors from a busy server.  Raise it on a link that has
    room to spare.

.PARAMETER Attempts
    How many times the script retries a whole unit before giving up on it and
    moving to the next one.

.PARAMETER FileRetry
    Passed to repocli as -r: per-file retries inside a unit.

.PARAMETER Redo
    Ignore the recorded state and re-run units that were already marked done.

.PARAMETER ListOnly
    Show the units that would be uploaded and exit.

.PARAMETER ShowProgress
    Let repocli draw its own per-file progress bar on the console instead of
    running it in silent mode.  Nothing is written to the per-unit log in this
    mode, so a failed file can then only be found through the error file.  The
    overall unit-by-unit bar is drawn either way.

.EXAMPLE
    # uploads to /HM_neurons, one unit per <rat>\<date> folder
    .\backup-to-repo.ps1 -Source E:\HM_neurons -Dest HM_neurons -Depth 2

.EXAMPLE
    # preview the split first
    .\backup-to-repo.ps1 -Source E:\HM_neurons -Dest HM_neurons -Depth 2 -ListOnly

.EXAMPLE
    # a rat kept on another drive still lands beside the others
    .\backup-to-repo.ps1 -Source F:\HM_neurons -Dest HM_neurons `
        -Depth 2 -Include 'Rat7_491392*'

.NOTES
    Prerequisites:
      1. repocli.exe on PATH (or pass -Repocli C:\tools\repocli.exe)
      2. run `repocli config` once to store baseURL + data-access credentials
         in %USERPROFILE%\.repocli.yml
    Throughput: -Threads sets the parallel transfers per directory upload
    (repocli -n). To push harder still, start 2-3 instances of this script
    with different -Include filters; unit locking makes that safe.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Source,

    [Parameter(Mandatory = $true)]
    [string]$Dest,

    [string]$Root = '',

    [string]$Repocli = 'repocli.exe',
    [string]$Config = (Join-Path $env:USERPROFILE '.repocli.yml'),
    [string]$Url = '',
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'repocli-backup'),

    [int]$Depth = 1,
    [int]$Threads = 1,
    [int]$FileRetry = 3,
    [int]$Attempts = 3,
    [int]$RetryDelaySeconds = 60,

    [string[]]$Include,
    [string[]]$Exclude,

    [switch]$MeasureSize,
    [switch]$Redo,
    [switch]$DryRun,
    [switch]$ListOnly,
    [switch]$ShowProgress,
    [switch]$StopOnError
)

# Native tools write to stderr for ordinary logging; 'Stop' would turn that
# into a terminating error, so keep the default and check exit codes by hand.
$ErrorActionPreference = 'Continue'

#--------------------------------------------------------------------------
# helpers
#--------------------------------------------------------------------------

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        default { Write-Host $line }
    }
    if ($script:RunLog) {
        Add-Content -LiteralPath $script:RunLog -Value $line -Encoding UTF8
    }
}

function Join-RemotePath {
    param([string]$Parent, [string]$Name)

    $p = $Parent.TrimEnd('/')
    if (-not $p) { return "/$Name" }
    return "$p/$Name"
}

# Stable, filesystem-safe key for a unit: readable prefix + hash of the full
# relative path *and* where it is being sent, so neither two different folders
# nor the same folder aimed at two different roots share a state file.
function Get-UnitKey {
    param([string]$Rel, [string]$Dest)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("$Dest|$Rel")
        $hash = ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join ''
    }
    finally {
        $sha.Dispose()
    }

    $safe = $Rel -replace '[^A-Za-z0-9._-]', '_'
    if ($safe.Length -gt 60) { $safe = $safe.Substring(0, 60) }
    if (-not $safe) { $safe = 'root' }
    return '{0}_{1}' -f $safe, $hash.Substring(0, 10)
}

function Format-Bytes {
    param([double]$Bytes)

    $units = 'B', 'KB', 'MB', 'GB', 'TB', 'PB'
    $i = 0
    while ($Bytes -ge 1024 -and $i -lt ($units.Count - 1)) {
        $Bytes = $Bytes / 1024
        $i++
    }
    return '{0:N2} {1}' -f $Bytes, $units[$i]
}

#--------------------------------------------------------------------------
# unit discovery
#--------------------------------------------------------------------------

$script:Units = New-Object System.Collections.ArrayList

# Windows volume internals that exist when $Source is a drive root; they are not
# readable and are never part of a data backup. Junctions/symlinks are skipped
# too, to avoid walking the same data twice or looping.
$script:SkipDirNames = @(
    'System Volume Information',
    '$RECYCLE.BIN',
    'RECYCLER',
    'Config.Msi',
    '$WinREAgent',
    'Recovery'
)

function Add-Unit {
    param([hashtable]$Unit)

    $Unit['Key'] = Get-UnitKey -Rel $Unit['Rel'] -Dest $Unit['DestArg']
    [void]$script:Units.Add([pscustomobject]$Unit)
}

function Find-Units {
    param(
        [string]$LocalDir,      # local directory being examined
        [string]$RemoteDir,     # repository path mirroring $LocalDir
        [string]$RemoteParent,  # repository path mirroring $LocalDir's parent ('' at the root)
        [int]$Level,
        [string]$Rel            # path of $LocalDir relative to $Source ('' at the root)
    )

    try {
        $children = Get-ChildItem -LiteralPath $LocalDir -Force -ErrorAction Stop
    }
    catch {
        Write-Log "cannot read $LocalDir : $($_.Exception.Message)" 'WARN'
        return
    }

    $subdirs = @($children | Where-Object {
        $_.PSIsContainer -and
        ($script:SkipDirNames -notcontains $_.Name) -and
        -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
    })
    $files = @($children | Where-Object { -not $_.PSIsContainer })

    # A leaf of the split: upload this whole directory as one unit. `repocli put
    # <dir> <parent>` (no trailing separator) recreates <dir> by name under
    # <parent>, which is what we want.
    if ($RemoteParent -and ($Level -ge $Depth -or $subdirs.Count -eq 0)) {
        Add-Unit @{
            Kind    = 'dir'
            Local   = $LocalDir
            DestArg = $RemoteParent
            Rel     = $Rel
        }
        return
    }

    foreach ($d in $subdirs) {
        $childRel = if ($Rel) { Join-Path $Rel $d.Name } else { $d.Name }
        Find-Units -LocalDir $d.FullName `
                   -RemoteDir (Join-RemotePath $RemoteDir $d.Name) `
                   -RemoteParent $RemoteDir `
                   -Level ($Level + 1) `
                   -Rel $childRel
    }

    # Loose files at a level we are still splitting through are grouped into one
    # unit and uploaded individually.
    if ($files.Count -gt 0) {
        $relLabel = if ($Rel) { "$Rel\*" } else { '*' }
        Add-Unit @{
            Kind    = 'files'
            Local   = $LocalDir
            Files   = @($files | ForEach-Object { $_.FullName })
            DestArg = $RemoteDir
            Rel     = $relLabel
        }
    }
}

#--------------------------------------------------------------------------
# repocli invocation
#--------------------------------------------------------------------------

function Invoke-Repocli {
    param([string[]]$Arguments, [string]$LogFile)

    Write-Verbose ("running: {0} {1}" -f $script:RepocliPath, ($Arguments -join ' '))

    if ($ShowProgress) {
        # Call repocli straight, with no pipeline: it redraws its bar with
        # carriage returns, which a pipeline would turn into one line per update.
        # Running it in this console rather than through Start-Process also means
        # Ctrl+C reaches repocli, instead of merely ending the wait and leaving it
        # uploading in the background. PowerShell quotes the arguments itself, so
        # paths containing spaces need no help here. Out-Host keeps stdout off
        # this function's output stream -- whatever a function writes there
        # becomes part of its return value, and the exit code would come back as
        # an array. Only stdout is piped; the bar is on stderr and still reaches
        # the console untouched.
        & $script:RepocliPath @Arguments | Out-Host
        if ($null -eq $LASTEXITCODE) { return 0 }
        return [int]$LASTEXITCODE
    }

    # Write-Host (not the pipeline) so repocli's output stays visible live
    # without becoming part of this function's return value.
    & $script:RepocliPath @Arguments 2>&1 | ForEach-Object {
        $line = "$_"
        Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
        Write-Host "    $line" -ForegroundColor DarkGray
    }
    if ($null -eq $LASTEXITCODE) { return 0 }
    return [int]$LASTEXITCODE
}

function Test-UnitSucceeded {
    param([int]$ExitCode, [string]$LogFile, [string]$ErrFile)

    if ($ExitCode -ne 0) {
        Write-Log "  exit code $ExitCode" 'WARN'
        return $false
    }

    if (Test-Path -LiteralPath $ErrFile) {
        $len = (Get-Item -LiteralPath $ErrFile).Length
        if ($len -gt 0) {
            $first = (Get-Content -LiteralPath $ErrFile -TotalCount 3) -join '; '
            Write-Log "  repocli reported file errors, see $ErrFile -> $first" 'WARN'
            return $false
        }
    }

    if ((-not $ShowProgress) -and (Test-Path -LiteralPath $LogFile)) {
        $log = Get-Content -LiteralPath $LogFile -Raw
        if ($log -and $log -match 'cannot create repo dir') {
            Write-Log '  repocli could not create a remote directory' 'WARN'
            return $false
        }
    }

    return $true
}

function Invoke-Unit {
    param([psobject]$Unit, [string]$LogFile, [string]$ErrFile)

    $common = @('-c', $Config, '-n', "$Threads")
    if ($Url) { $common += @('-u', $Url) }
    if (-not $ShowProgress) { $common += '-s' }

    if ($Unit.Kind -eq 'dir') {
        $cliArgs = @('put', $Unit.Local, $Unit.DestArg, '-r', "$FileRetry", '-e', $ErrFile) + $common
        $code = Invoke-Repocli -Arguments $cliArgs -LogFile $LogFile
        return (Test-UnitSucceeded -ExitCode $code -LogFile $LogFile -ErrFile $ErrFile)
    }

    # 'files': make sure the destination directory exists, otherwise `put` of a
    # single file would create a *file* at the destination path.
    $code = Invoke-Repocli -Arguments (@('mkdir', $Unit.DestArg) + $common) -LogFile $LogFile
    if ($code -ne 0) {
        Write-Log "  mkdir $($Unit.DestArg) failed (exit $code)" 'WARN'
        return $false
    }

    $ok = $true
    foreach ($f in $Unit.Files) {
        $cliArgs = @('put', $f, $Unit.DestArg, '-r', "$FileRetry") + $common
        $code = Invoke-Repocli -Arguments $cliArgs -LogFile $LogFile
        if ($code -ne 0) {
            Write-Log "  failed: $f (exit $code)" 'WARN'
            Add-Content -LiteralPath $ErrFile -Value "$f error: exit $code" -Encoding UTF8
            $ok = $false
        }
    }
    return $ok
}

#--------------------------------------------------------------------------
# preflight
#--------------------------------------------------------------------------

$cmd = Get-Command $Repocli -ErrorAction SilentlyContinue
if (-not $cmd) {
    $local = Join-Path $PSScriptRoot 'repocli.exe'
    if (Test-Path -LiteralPath $local) {
        $cmd = Get-Command $local
    }
    else {
        throw "repocli not found: '$Repocli'. Put repocli.exe on your PATH or pass -Repocli <path>."
    }
}
$script:RepocliPath = $cmd.Source

if (-not (Test-Path -LiteralPath $Source)) {
    throw "source not found: $Source"
}
$Source = (Resolve-Path -LiteralPath $Source).ProviderPath
# Keep the trailing backslash of a drive root ("E:\"); "E:" would be read as
# "current directory on E:" instead.
if ($Source.Length -gt 3) { $Source = $Source.TrimEnd('\') }

if (-not (Test-Path -LiteralPath $Config)) {
    throw "repocli configuration not found: $Config. Run 'repocli config' first."
}

# Everything lands under one root folder, so backing up a second drive is a
# matter of changing -Root rather than every -Dest.
$segments = @($Root, $Dest) | ForEach-Object { $_.Trim('/') } | Where-Object { $_ }
$Dest = '/' + ($segments -join '/')

$stateDir = Join-Path $WorkDir 'state'
$logDir = Join-Path $WorkDir 'logs'
$errDir = Join-Path $WorkDir 'errors'
foreach ($d in @($WorkDir, $stateDir, $logDir, $errDir)) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$script:RunLog = Join-Path $WorkDir ('run-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

Write-Log "source      : $Source"
Write-Log "destination : $Dest"
Write-Log "repocli     : $($script:RepocliPath)"
Write-Log "work dir    : $WorkDir"
Write-Log "split depth : $Depth"

# Fail fast on a wrong destination or bad credentials instead of discovering it
# after hours of transfer.
if (-not $ListOnly -and -not $DryRun) {
    $baseArgs = @('-c', $Config)
    if ($Url) { $baseArgs += @('-u', $Url) }

    & $script:RepocliPath @(@('ls', $Dest) + $baseArgs) 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        # The destination may simply not exist yet; try to create it.
        & $script:RepocliPath @(@('mkdir', $Dest) + $baseArgs) 2>&1 | Out-Null
        & $script:RepocliPath @(@('ls', $Dest) + $baseArgs) 2>&1 | Out-Null
    }
    if ($LASTEXITCODE -ne 0) {
        Write-Log "cannot reach $Dest in the repository." 'ERROR'
        Write-Log "  If baseURL in $Config is https://webdav.data.ru.nl, -Dest must be the" 'ERROR'
        Write-Log '  collection path, e.g. /dcn.DSC_626830_0003_227' 'ERROR'
        Write-Log '  If baseURL is the full collection URL, use -Dest /' 'ERROR'
        Write-Log '  Also check the data-access credentials with: repocli ls /' 'ERROR'
        throw "destination check failed: $Dest"
    }
    Write-Log 'destination reachable'
}

#--------------------------------------------------------------------------
# build the work list
#--------------------------------------------------------------------------

Write-Log 'scanning source tree ...'
Find-Units -LocalDir $Source -RemoteDir $Dest -RemoteParent '' -Level 0 -Rel ''

if ($Include) {
    $script:Units = @($script:Units | Where-Object {
        $u = $_
        @($Include | Where-Object { $u.Rel -like $_ }).Count -gt 0
    })
}
if ($Exclude) {
    $script:Units = @($script:Units | Where-Object {
        $u = $_
        -not (@($Exclude | Where-Object { $u.Rel -like $_ }).Count -gt 0)
    })
}

$total = @($script:Units).Count
Write-Log "$total unit(s) to consider"

if ($ListOnly) {
    $script:Units | ForEach-Object {
        $done = Test-Path -LiteralPath (Join-Path $stateDir "$($_.Key).done")
        '{0,-6} {1,-5} {2}  ->  {3}' -f $(if ($done) { 'done' } else { '' }), $_.Kind, $_.Rel, $_.DestArg
    }
    return
}

#--------------------------------------------------------------------------
# upload
#--------------------------------------------------------------------------

$stats = [pscustomobject]@{
    Done    = 0
    Skipped = 0
    Failed  = 0
    Locked  = 0
    Bytes   = 0L
}
$failedUnits = New-Object System.Collections.ArrayList
$runStart = Get-Date
$index = 0
$activity = "backup to $Dest"

try {
    foreach ($unit in $script:Units) {
        $index++

        # Overall progress. Write-Progress draws in the host's own progress area
        # instead of on stdout/stderr, so it shows in both modes: repocli writes
        # its per-file bar to stderr with carriage returns, which the pipeline
        # this script normally reads it through would shred into log lines.
        $progress = @{
            Activity        = $activity
            Status          = "[$index/$total] $($unit.Rel)"
            PercentComplete = [int](100 * ($index - 1) / $total)
        }
        if ($stats.Done -gt 0) {
            $perUnit = ((Get-Date) - $runStart).TotalSeconds / $stats.Done
            $progress['SecondsRemaining'] = [int]($perUnit * ($total - $index + 1))
        }
        Write-Progress @progress

        $stateFile = Join-Path $stateDir "$($unit.Key).done"
        $lockFile = Join-Path $stateDir "$($unit.Key).lock"
        $logFile = Join-Path $logDir "$($unit.Key).log"
        $errFile = Join-Path $errDir "$($unit.Key).err"

        if ((Test-Path -LiteralPath $stateFile) -and (-not $Redo)) {
            Write-Log "[$index/$total] skip (already done): $($unit.Rel)"
            $stats.Skipped++
            continue
        }

        # Exclusive lock so several instances of this script can run side by side.
        $lock = $null
        try {
            $lock = [System.IO.File]::Open($lockFile, 'OpenOrCreate', 'ReadWrite', 'None')
        }
        catch {
            Write-Log "[$index/$total] skip (locked by another run): $($unit.Rel)" 'WARN'
            $stats.Locked++
            continue
        }

        try {
            $size = 0L
            if ($MeasureSize) {
                $m = Get-ChildItem -LiteralPath $unit.Local -Recurse -Force -ErrorAction SilentlyContinue |
                     Where-Object { -not $_.PSIsContainer } | Measure-Object -Property Length -Sum
                if ($m) { $size = [int64]$m.Sum }
                Write-Log "[$index/$total] $($unit.Rel) -> $($unit.DestArg)  ($(Format-Bytes $size))"
            }
            else {
                Write-Log "[$index/$total] $($unit.Rel) -> $($unit.DestArg)"
            }

            if ($DryRun) {
                Write-Log '  (dry run, nothing transferred)'
                continue
            }

            # Start from no error file at all. Writing an "empty" file with
            # Set-Content -Encoding UTF8 would emit a 3-byte BOM, and the
            # non-empty check below would then flag every unit as failed.
            Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
            $unitStart = Get-Date
            $ok = $false

            for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
                if ($attempt -gt 1) {
                    Write-Log "  retry $attempt/$Attempts in $RetryDelaySeconds s ..." 'WARN'
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
                $ok = Invoke-Unit -Unit $unit -LogFile $logFile -ErrFile $errFile
                if ($ok) { break }
            }

            $elapsed = (Get-Date) - $unitStart

            if ($ok) {
                $record = [pscustomobject]@{
                    rel       = $unit.Rel
                    local     = $unit.Local
                    dest      = $unit.DestArg
                    bytes     = $size
                    seconds   = [int]$elapsed.TotalSeconds
                    completed = (Get-Date -Format 'o')
                }
                $record | ConvertTo-Json -Compress | Set-Content -LiteralPath $stateFile -Encoding UTF8
                $stats.Done++
                $stats.Bytes += $size
                Write-Log ("  ok in {0:hh\:mm\:ss}" -f $elapsed) 'OK'
            }
            else {
                $stats.Failed++
                [void]$failedUnits.Add($unit.Rel)
                Write-Log "  FAILED after $Attempts attempt(s); log: $logFile" 'ERROR'
                if ($StopOnError) { throw "stopping on error in unit: $($unit.Rel)" }
            }
        }
        finally {
            if ($lock) { $lock.Close() }
            Remove-Item -LiteralPath $lockFile -Force -ErrorAction SilentlyContinue
        }
    }
}
finally {
    Write-Progress -Activity $activity -Completed

    $runElapsed = (Get-Date) - $runStart
    Write-Log '---------------------------------------------------------------'
    Write-Log ("finished in {0:d\.hh\:mm\:ss}" -f $runElapsed)
    Write-Log ("uploaded {0} unit(s), skipped {1}, locked {2}, failed {3}" -f `
        $stats.Done, $stats.Skipped, $stats.Locked, $stats.Failed)
    if ($MeasureSize) { Write-Log ("transferred (unit size sum): {0}" -f (Format-Bytes $stats.Bytes)) }

    if ($failedUnits.Count -gt 0) {
        $failFile = Join-Path $WorkDir 'failed.txt'
        $failedUnits | Set-Content -LiteralPath $failFile -Encoding UTF8
        Write-Log "failed units listed in $failFile" 'ERROR'
        Write-Log 'Re-run the same command to retry only what is missing.' 'WARN'
    }
    Write-Log "run log: $($script:RunLog)"
}

if ($stats.Failed -gt 0) { exit 1 }
exit 0
