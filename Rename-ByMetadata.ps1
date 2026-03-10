[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Path,
    [switch]$Apply,
    [switch]$ImmediateRename,
    [switch]$QuickPreview,
    [int]$QuickPreviewLimit = 200,
    [switch]$NoRecurse,
    [switch]$InteractiveSubfolders,
    [switch]$IncludeHidden,
    [switch]$Force,
    [switch]$OnlyGeneric,
    [string]$ReportPath,
    [string]$CsvPath,
    [string]$LogPath,
    [switch]$Resume,
    [string]$ResumeLogPath,
    [switch]$RetryErrors,
    [string]$RetryErrorsLogPath,
    [switch]$IntelligentScan,
    [switch]$IncludeOtherFiles,
    [switch]$ScanOtherFilesAtEnd,
    [switch]$ListUncommonFiles,
    [switch]$NoAutoResumePrompt,
    [bool]$SkipLikelyAlreadyRenamed = $true,
    [int]$PerFileTimeoutSeconds = 20,
    [int]$BatchSize = 500,
    [switch]$SkipMetadataUpgrade,
    [switch]$Unattended,
    [switch]$UseWordCom,
    [int]$MaxFileReadMB = 50,
    [int]$ComRecycleInterval = 200,
    [switch]$BackfillRenameSuffix,
    [string]$BackfillLogPath,
    [switch]$Help,
    [switch]$DebugOutput
)

<#
.SYNOPSIS
Renames files to meaningful names using embedded metadata with a preview-first workflow.

.DESCRIPTION
Scans a target folder (optionally recursive), extracts metadata from supported files, builds a
rename plan, and emits live color-coded output with confidence scoring.


# Version: 1.8.0

This script exists to rebuild meaningful filenames from chaotic recovery sets where original
names are lost. It extracts deep metadata from documents, images, audio, and video, ranks it by
reliability, and generates a safe, preview-first rename plan with confidence scoring. Traditional
renamers cannot interpret DOCX/PDF/XMP metadata or handle corrupted recovery trees, so this tool
fills that gap by acting as a forensic-grade metadata reconstruction engine rather than a
pattern-based renamer.

Default behavior is preview-first:
- Prints live preview lines: PREVIEW [confidence] old -> new [source]
- Writes a TXT preview report
- Writes a CSV preview report
- Prompts at end to optionally apply the generated rename plan immediately

Quick preview mode:
- Use -QuickPreview to preview only a limited number of files for rapid validation.
- Default quick preview limit is 200 files (customizable with -QuickPreviewLimit).

Actual rename occurs when -Apply is used, when startup prompt selects rename mode, or when
confirmed at the end-of-preview prompt. In apply mode, live RENAMED output is shown for each
successful rename.


Interactive flow can prompt for target path, preview/apply mode, quick preview mode/limit.
Shell metadata queries and embedded media metadata extraction can be slower on large trees.
Run preview first for large datasets, then execute apply mode after review.

RENAME PLAN FIELDS
CSV fields:
- RelativePath
- OldName
- NewName
- Source
- ConfidenceScore
- ConfidenceLevel
- ScoreReason
- FullPath

TXT report includes:
- Human-readable old -> new mappings
- Source and confidence context
- Candidate and skipped totals

MODULE INTEGRATION
This help block is module-ready and suitable for migration into an advanced function/cmdlet
with minimal updates to parameter metadata and examples.

if (-not $isImmediateRenameMode) {
Maximum number of files to evaluate in -QuickPreview mode.
Default is 200.

.PARAMETER LogPath
Optional path for the run log file.
If omitted, a timestamped log file is created in the target root folder as Rename-Run-YYYYMMDD-HHMMSS.log.
The run log includes startup settings, per-file skip diagnostics (including timeout skips),
structured RENAMED entries with From/To paths, and an end-of-log completion marker.

.PARAMETER Resume
When specified, loads prior structured RENAMED entries from a previous run log and skips files
that were already renamed in that log (matching either original or target path).
Also restores prior job settings from that same log before interactive prompts so values do not
need to be re-entered (explicit command-line arguments still take precedence).

.PARAMETER ResumeLogPath
Optional path to a previous run log used for resume state.
If omitted with -Resume, the script auto-selects the newest Rename-Run-*.log in the target root.
When used for resume, the script can also restore prior job settings recorded in that log.

.PARAMETER RetryErrors
Retries only files previously logged as ERROR entries from a prior run log.
If used without -RetryErrorsLogPath, the script auto-selects the newest Rename-Run-*.log in the target root.
When retrying from a log, prior job settings from that log can also be restored before prompts
(explicit command-line arguments still take precedence).

.PARAMETER RetryErrorsLogPath
Optional path to a previous run log used for retrying ERROR files.
If omitted with -RetryErrors, the script auto-selects the newest Rename-Run-*.log in the target root.
When used with retry mode, the selected log may also provide prior job settings.

.PARAMETER IntelligentScan
Limits discovery to common end-user file types first (documents, photos, videos, audio, XML)
and excludes uncommon/system/archive-like file types unless explicitly included.

.PARAMETER IncludeOtherFiles
Disables intelligent filtering and includes uncommon/other file types in the same run.
Cannot be combined with -ScanOtherFilesAtEnd.

.PARAMETER ScanOtherFilesAtEnd
When intelligent scan is enabled, queues uncommon/other file types and scans them after
common file types complete.
Cannot be combined with -IncludeOtherFiles.

.PARAMETER ListUncommonFiles
When intelligent scan excludes uncommon/other file types (without -ScanOtherFilesAtEnd),
prints and logs each excluded file with an [Uncommon] prefix.
Use this for auditing exactly which uncommon files were skipped.

.PARAMETER NoAutoResumePrompt
Disables the automatic prompt that appears when an eligible interrupted prior run log is detected
(log has structured RENAMED entries and no end-of-log completion marker).
Use this for fully non-interactive runs unless resume is explicitly requested via -Resume or -ResumeLogPath.

.PARAMETER SkipLikelyAlreadyRenamed
Optional performance switch that skips files whose existing names already match common metadata-based
output patterns used by this script (for example date-labeled media names or collision-suffixed variants).
Use with care: this is heuristic and may skip some files you still want to evaluate.

.PARAMETER PerFileTimeoutSeconds
Maximum time budget per file for metadata evaluation before the file is skipped.
Default is 20 seconds.
Timeout is evaluated after metadata extraction returns; if elapsed time exceeds this limit,
the file is counted as [TIMEOUT], logged, and excluded from planning/rename.

.PARAMETER BatchSize
Number of files to process per batch during long runs.
Default is 500.
Batching is always enabled and does not prompt; lower values can reduce UI stalls on very large trees.

.NOTES
- Word COM title extraction requires Microsoft Word installed for .doc/.dot/.docx/.docm/.dotx/.dotm/.rtf COM scenarios.
- COM extraction may be slower than native OpenXML/RTF parsing on large file sets.
- If Word COM is unavailable, the script continues with non-COM metadata sources.

.EXAMPLE
.\Rename-ByMetadata.ps1
Runs a basic preview with live output and writes TXT/CSV plan files.
If no path is entered at the prompt, the current working directory is used.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed"
Runs recursive preview for the target path (recursive by default).

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -Apply
Applies planned renames and prints live RENAMED output.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -Apply -ImmediateRename
Starts renaming immediately during processing (no separate planned-preview pass).

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -ReportPath "D:\Reconstructed\rename-report.txt" -CsvPath "D:\Reconstructed\rename-report.csv"
Writes report artifacts to custom log/report file paths.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -LogPath "D:\Reconstructed\rename-run.log"
Runs preview with a custom run log file.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -QuickPreview
Runs a fast sample preview against the first 200 discovered files.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -QuickPreview -QuickPreviewLimit 50
Runs a fast sample preview against the first 50 discovered files.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -QuickPreview -QuickPreviewLimit 50 -Apply
Applies renames only within the sampled quick preview set.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -InteractiveSubfolders
Prompts y/n for each subfolder before including it in the scan.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -NoAutoResumePrompt
Runs without auto-prompting to resume from detected prior run logs.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -Resume
Resumes using the newest Rename-Run-*.log in the target root and skips files already renamed in that prior run.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -ResumeLogPath "D:\Reconstructed\Rename-Run-20260226-213000.log"
Resumes from a specific prior run log path.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -RetryErrors
Retries files logged as ERROR from the newest run log in the target root.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -RetryErrorsLogPath "D:\Reconstructed\Rename-Run-20260226-213000.log"
Retries files logged as ERROR from a specific prior run log.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -IntelligentScan
Scans only common end-user file types.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -IntelligentScan -ScanOtherFilesAtEnd
Scans common end-user file types first, then scans uncommon file types at the end.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -IncludeOtherFiles
Includes uncommon/other file types in the same run (disables intelligent filtering).

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -IntelligentScan -ListUncommonFiles
Scans common file types and emits an [Uncommon] line for each excluded uncommon file.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -PerFileTimeoutSeconds 45
Uses a 45-second per-file metadata evaluation budget before counting slow files as [TIMEOUT] and skipping them.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Path "D:\Reconstructed" -BatchSize 500
Processes files in batches of 500 to reduce long-run UI stalls.

.EXAMPLE
.\Rename-ByMetadata.ps1 -Help
Displays command-line parameter help and exits.

.EXAMPLE
Rename-ByMetadata.exe /?
Displays command-line parameter help in the compiled EXE.

.EXAMPLE
Rename-ByMetadata.exe /help
Displays command-line parameter help in the compiled EXE.

.EXAMPLE
Pipeline input is not currently supported.
Use -Path to specify the scan root for this script version.
#>

# Script-wide confidence score matrix
$script:ScoreMatrix = @{
    'DOCX core title' = [PSCustomObject]@{ Score = 98; Reason = 'Embedded canonical document title' }
    'DOCX heading title' = [PSCustomObject]@{ Score = 92; Reason = 'Document heading structure from app.xml' }
    'Word OpenXML core title' = [PSCustomObject]@{ Score = 98; Reason = 'Embedded Word OpenXML canonical document title' }
    'Word heading title' = [PSCustomObject]@{ Score = 92; Reason = 'Word heading structure from app.xml TitlesOfParts' }
    'PPTX core title' = [PSCustomObject]@{ Score = 97; Reason = 'Embedded PowerPoint OpenXML document title' }
    'PPTX heading title' = [PSCustomObject]@{ Score = 91; Reason = 'PowerPoint heading from app.xml TitlesOfParts' }
    'XLSX core title' = [PSCustomObject]@{ Score = 96; Reason = 'Embedded Excel OpenXML document title' }
    'XLSX heading title' = [PSCustomObject]@{ Score = 90; Reason = 'Excel heading from app.xml TitlesOfParts' }
    'RTF title control word' = [PSCustomObject]@{ Score = 96; Reason = 'Embedded RTF \\title metadata control word' }
    'Word COM title' = [PSCustomObject]@{ Score = 90; Reason = 'Word built-in document title via COM automation' }
    'PDF XMP title' = [PSCustomObject]@{ Score = 94; Reason = 'Embedded PDF XMP title metadata' }
    'PDF Info title' = [PSCustomObject]@{ Score = 92; Reason = 'PDF /Info dictionary /Title field' }
    'PDF bookmark title' = [PSCustomObject]@{ Score = 89; Reason = 'First PDF outline/bookmark title' }
    'XML title node' = [PSCustomObject]@{ Score = 93; Reason = 'Structured XML title element' }
    'Image embedded XMP title' = [PSCustomObject]@{ Score = 95; Reason = 'Embedded authored image title metadata' }
    'Windows metadata title' = [PSCustomObject]@{ Score = 88; Reason = 'Windows semantic metadata title field' }
    'Audio composite metadata' = [PSCustomObject]@{ Score = 91; Reason = 'Artist/title/album/track composite metadata' }
    'Audio title metadata' = [PSCustomObject]@{ Score = 85; Reason = 'Audio title metadata only' }
    'Video episode metadata' = [PSCustomObject]@{ Score = 87; Reason = 'Season/episode/media title metadata' }
    'GPS location metadata' = [PSCustomObject]@{ Score = 80; Reason = 'Location metadata from GPS fields' }
    'Media date metadata' = [PSCustomObject]@{ Score = 72; Reason = 'Date-based media fallback naming' }
	'Author/company metadata' = [PSCustomObject]@{ Score = 62; Reason = 'Attribution fallback metadata' }
}


$script:LogPath = $null
$script:StopRequested = $false
$script:LongRunThreshold = 500
$script:ResumeActive = $false
$script:ResumeLogSource = $null
$script:ResumeRenamedFrom = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:ResumeRenamedTo = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:ResumeRenamedScores = [System.Collections.Generic.Dictionary[string,PSCustomObject]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:EndOfLogWritten = $false
$script:LastUsedPathStateFile = Join-Path -Path $env:APPDATA -ChildPath 'Rename-ByMetadata\last-scan-path.txt'
$script:UnattendedMode = $false
$script:MaxFileReadBytes = 50 * 1024 * 1024
$script:ComRecycleInterval = 200
$script:NoWordComMode = $true
$script:OriginalWindowTitle = $null

function Update-WindowTitle {
    param(
        [string]$Phase,
        [int]$Processed,
        [int]$Total,
        [int]$Renamed,
        [int]$Errors,
        [double]$ElapsedSeconds
    )

    $elapsed = [TimeSpan]::FromSeconds($ElapsedSeconds)
    $elapsedText = if ($elapsed.TotalHours -ge 1) {
        '{0:N0}h {1:N0}m' -f [Math]::Floor($elapsed.TotalHours), $elapsed.Minutes
    } else {
        '{0:N0}m {1:N0}s' -f [Math]::Floor($elapsed.TotalMinutes), $elapsed.Seconds
    }

    $pct = if ($Total -gt 0) { [int](($Processed / $Total) * 100) } else { 0 }
    $title = '[Rename-{0}] {1}/{2} ({3}%) | Renamed:{4} Err:{5} | {6}' -f $Phase, $Processed, $Total, $pct, $Renamed, $Errors, $elapsedText

    try { $host.UI.RawUI.WindowTitle = $title } catch {}
}

function Restore-WindowTitle {
    if (-not [string]::IsNullOrWhiteSpace($script:OriginalWindowTitle)) {
        try { $host.UI.RawUI.WindowTitle = $script:OriginalWindowTitle } catch {}
    }
}


function Get-ConfidenceLevel {
    param(
        [Parameter(Mandatory)]
        [int]$ConfidenceScore
    )

    if ($ConfidenceScore -ge 85) { return 'High' }
    if ($ConfidenceScore -ge 60) { return 'Medium' }
    return 'Low'
}

function Get-ScoreFromMatrix {
    param(
        [Parameter(Mandatory)]
        [string]$Source
    )

    if ($script:ScoreMatrix.Contains($Source)) {
        return $script:ScoreMatrix[$Source]
    }

    return [PSCustomObject]@{ Score = 55; Reason = 'Unmapped source fallback score' }
}

function Write-RunLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($script:LogPath)) {
        return
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -LiteralPath $script:LogPath -Value ("[{0}] {1}" -f $timestamp, $Message) -Encoding UTF8
}

function Get-LastUsedScanPath {
    if ([string]::IsNullOrWhiteSpace($script:LastUsedPathStateFile)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $script:LastUsedPathStateFile)) {
        return $null
    }

    try {
        $storedPath = Get-Content -LiteralPath $script:LastUsedPathStateFile -ErrorAction Stop | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace($storedPath)) {
            return $null
        }

        $candidate = $storedPath.Trim()
        if (-not (Test-Path -LiteralPath $candidate -PathType Container)) {
            return $null
        }

        $resolved = Resolve-Path -LiteralPath $candidate -ErrorAction Stop
        return $resolved.Path
    }
    catch {
        return $null
    }
}

function Save-LastUsedScanPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($script:LastUsedPathStateFile) -or [string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    try {
        $stateDir = Split-Path -Path $script:LastUsedPathStateFile -Parent
        if (-not [string]::IsNullOrWhiteSpace($stateDir) -and -not (Test-Path -LiteralPath $stateDir)) {
            New-Item -Path $stateDir -ItemType Directory -Force | Out-Null
        }

        Set-Content -LiteralPath $script:LastUsedPathStateFile -Value $Path -Encoding UTF8
    }
    catch {
    }
}

function Initialize-RunLog {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [switch]$Recurse,

        [Parameter(Mandatory)]
        [string]$Mode,

        [Parameter(Mandatory)]
        [switch]$Admin,

        [string]$RequestedLogPath,

        [hashtable]$Settings
    )

    if ([string]::IsNullOrWhiteSpace($RequestedLogPath)) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $script:LogPath = Join-Path $Path ("Rename-Run-{0}.log" -f $stamp)
    }
    else {
        $script:LogPath = $RequestedLogPath
    }

    $logDir = Split-Path -Path $script:LogPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }

    $headerLines = [System.Collections.Generic.List[string]]::new()
    $headerLines.Add('Rename-ByMetadata run log') | Out-Null
    $headerLines.Add(('Started: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))) | Out-Null
    $headerLines.Add(('Path: {0}' -f $Path)) | Out-Null
    $headerLines.Add(('Recurse: {0}' -f $Recurse)) | Out-Null
    $headerLines.Add(('Mode: {0}' -f $Mode)) | Out-Null
    $headerLines.Add(('Admin: {0}' -f $Admin)) | Out-Null

    if ($Settings -and $Settings.Count -gt 0) {
        $headerLines.Add('') | Out-Null
        $headerLines.Add('Selected settings:') | Out-Null
        foreach ($key in ($Settings.Keys | Sort-Object)) {
            $value = $Settings[$key]
            $headerLines.Add(('- {0}: {1}' -f $key, $value)) | Out-Null
        }
    }

    $headerLines.Add('') | Out-Null

    Set-Content -LiteralPath $script:LogPath -Value $headerLines -Encoding UTF8
}

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Read-YesNo {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [bool]$DefaultYes = $true
    )

    if ($script:UnattendedMode) {
        return $DefaultYes
    }

    $suffix = if ($DefaultYes) { '[Y/n]' } else { '[y/N]' }

    while ($true) {
        $inputValue = Read-Host "$Prompt $suffix"

        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            return $DefaultYes
        }

        switch ($inputValue.Trim().ToLowerInvariant()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
            default {
                Write-Host 'Please answer y or n.' -ForegroundColor Yellow
            }
        }
    }
}

function Get-DiscoveredFiles {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [switch]$IncludeHidden,

        [Parameter(Mandatory)]
        [switch]$Recurse,

        [Parameter(Mandatory)]
        [switch]$InteractiveSubfolders
    )

    if (-not $InteractiveSubfolders) {
        return @(Get-ChildItem -LiteralPath $Root -File -ErrorAction Stop -Force:$IncludeHidden -Recurse:$Recurse)
    }

    $discovered = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

    try {
        $rootFiles = Get-ChildItem -LiteralPath $Root -File -ErrorAction Stop -Force:$IncludeHidden
        foreach ($file in $rootFiles) {
            $discovered.Add($file) | Out-Null
        }
    }
    catch {
        Write-Host ("Warning: unable to read files in '{0}'" -f $Root) -ForegroundColor Yellow
        Write-RunLog -Message ("Warning: unable to read files in '{0}'" -f $Root)
    }

    $topLevelSubdirs = @()
    try {
        $topLevelSubdirs = @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction Stop -Force:$IncludeHidden | Sort-Object FullName)
    }
    catch {
        Write-Host ("Warning: unable to enumerate subfolders in '{0}'" -f $Root) -ForegroundColor Yellow
        Write-RunLog -Message ("Warning: unable to enumerate subfolders in '{0}'" -f $Root)
        return @($discovered)
    }

    foreach ($subdir in $topLevelSubdirs) {
        $includeFolder = Read-YesNo -Prompt ("Include subfolder '{0}'?" -f $subdir.FullName) -DefaultYes $true
        if ($includeFolder) {
            Write-RunLog -Message ("Interactive subfolder include (top-level): {0}" -f $subdir.FullName)
            try {
                $selectedFiles = Get-ChildItem -LiteralPath $subdir.FullName -File -ErrorAction Stop -Force:$IncludeHidden -Recurse
                foreach ($file in $selectedFiles) {
                    $discovered.Add($file) | Out-Null
                }
            }
            catch {
                Write-Host ("Warning: unable to read files in included subtree '{0}'" -f $subdir.FullName) -ForegroundColor Yellow
                Write-RunLog -Message ("Warning: unable to read files in included subtree '{0}'" -f $subdir.FullName)
            }
        }
        else {
            Write-RunLog -Message ("Interactive subfolder skip (top-level): {0}" -f $subdir.FullName)
        }
    }

    return @($discovered)
}

function Test-RunControl {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('preview', 'planning', 'rename')]
        [string]$Phase
    )

    if ($script:StopRequested) {
        return $false
    }

    if ($script:UnattendedMode) {
        return $true
    }

    try {
        $consoleAvailable = $false
        try { $null = [Console]::KeyAvailable; $consoleAvailable = $true } catch { $consoleAvailable = $false }
        if (-not $consoleAvailable) {
            # Fallback: prompt user for action
            $input = Read-Host "Type P to pause, S to stop, or Enter to continue ($Phase phase)"
            switch ($input.Trim().ToUpperInvariant()) {
                'P' {
                    Write-Host ("Paused during {0}. Type C to continue or S to stop." -f $Phase) -ForegroundColor Yellow
                    Write-RunLog -Message ("Paused during {0} phase by fallback prompt." -f $Phase)
                    while ($true) {
                        $resumeInput = Read-Host "Type C to continue or S to stop ($Phase phase)"
                        switch ($resumeInput.Trim().ToUpperInvariant()) {
                            'C' {
                                Write-Host ("Resuming {0}..." -f $Phase) -ForegroundColor Green
                                Write-RunLog -Message ("Resumed {0} phase by fallback prompt." -f $Phase)
                                return $true
                            }
                            'S' {
                                $script:StopRequested = $true
                                Write-Host ("Stop requested during paused {0} phase. Ending run gracefully..." -f $Phase) -ForegroundColor Yellow
                                Write-RunLog -Message ("Stop requested during paused {0} phase by fallback prompt." -f $Phase)
                                return $false
                            }
                            default { continue }
                        }
                    }
                }
                'S' {
                    $script:StopRequested = $true
                    Write-Host ("Stop requested during {0} phase. Ending run gracefully..." -f $Phase) -ForegroundColor Yellow
                    Write-RunLog -Message ("Stop requested during {0} phase by fallback prompt." -f $Phase)
                    return $false
                }
                default { return $true }
            }
        }
        # Console available: use hotkeys
        if (-not [Console]::KeyAvailable) {
            return $true
        }
        $key = [Console]::ReadKey($true)
        switch ($key.Key) {
            'P' {
                Write-Host ("Paused during {0}. Press C to continue or S to stop." -f $Phase) -ForegroundColor Yellow
                Write-RunLog -Message ("Paused during {0} phase by user hotkey." -f $Phase)
                while ($true) {
                    $resumeKey = [Console]::ReadKey($true)
                    switch ($resumeKey.Key) {
                        'C' {
                            Write-Host ("Resuming {0}..." -f $Phase) -ForegroundColor Green
                            Write-RunLog -Message ("Resumed {0} phase by user hotkey." -f $Phase)
                            return $true
                        }
                        'S' {
                            $script:StopRequested = $true
                            Write-Host ("Stop requested during paused {0} phase. Ending run gracefully..." -f $Phase) -ForegroundColor Yellow
                            Write-RunLog -Message ("Stop requested during paused {0} phase by user hotkey." -f $Phase)
                            return $false
                        }
                        default { continue }
                    }
                }
            }
            'S' {
                $script:StopRequested = $true
                Write-Host ("Stop requested during {0} phase. Ending run gracefully..." -f $Phase) -ForegroundColor Yellow
                Write-RunLog -Message ("Stop requested during {0} phase by user hotkey." -f $Phase)
                return $false
            }
            default { return $true }
        }
    }
    catch {
        return $true
    }
}

function ConvertTo-NormalizedMetadataTitle {
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )

    $normalized = $Title.Normalize([Text.NormalizationForm]::FormKC)
    $normalized = $normalized -replace '[\u2018\u2019]', "'"
    $normalized = $normalized -replace '[\u201C\u201D]', '"'
    $normalized = $normalized -replace '[\u2013\u2014]', '-'
    $normalized = $normalized -replace '[\|]+', ' - '
    $normalized = $normalized -replace '[_]+', ' '

    $boilerplatePatterns = @(
        '(?i)^copy\s+of\s+',
        '(?i)\bscanned\s+document\b',
        '(?i)\bnew\s+document\b',
        '(?i)\buntitled\b',
        '(?i)\bdraft\b',
        '(?i)\bfinal\s+version\b'
    )
    foreach ($pattern in $boilerplatePatterns) {
        $normalized = $normalized -replace $pattern, ' '
    }

    $normalized = $normalized -replace '\s+', ' '
    $normalized = $normalized.Trim(' ', '.', '-', '_')
    $normalized = $normalized -replace '[\p{P}\s]+$', ''

    $tokens = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($token in ($normalized -split '\s+')) {
        if ([string]::IsNullOrWhiteSpace($token)) { continue }
        if ($seen.Add($token)) {
            $tokens.Add($token) | Out-Null
        }
    }
    $normalized = ($tokens -join ' ').Trim()

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $result = $normalized
    if ($normalized -ceq $normalized.ToUpperInvariant()) {
        $result = (Get-Culture).TextInfo.ToTitleCase($normalized.ToLowerInvariant())
    }

    $result = $result -replace '\bPdf\b', 'PDF'
    $result = $result -replace '\bGps\b', 'GPS'
    $result = $result -replace '\bUsa\b', 'USA'
    $result = $result -replace '\bUk\b', 'UK'
    return $result
}

function Test-IsJunkTitle {
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )

    $candidate = $Title.Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) { return $true }
    if ($candidate.Length -lt 3) { return $true }
    if ($candidate -match '^\d+$') { return $true }

    if ($candidate -match '^(Photo|Video|Audio|Document|File)\s+\d{4}-\d{2}-\d{2}') {
        return $false
    }

    $lower = $candidate.ToLowerInvariant()
    if ($lower -match '^(file|img|dsc|scan|document|image|photo|untitled|new document|scanned document)[-_ ]?\d*$') { return $true }
    if ($lower -match '^(copy|final|new)$') { return $true }
    if ($lower -match '^(jpg|jpeg|png|pdf|audio|video|document)$') { return $true }

    $letters = ([regex]::Matches($candidate, '[A-Za-z]')).Count
    $nonLetters = [Math]::Max(0, $candidate.Length - $letters)
    if ($candidate.Length -ge 6 -and $letters -eq 0) { return $true }
    if ($candidate.Length -gt 0 -and (($nonLetters / $candidate.Length) -gt 0.60) -and ($candidate -notmatch '^\d{4}-\d{2}-\d{2}')) { return $true }

    return $false
}

function New-MetadataNameResult {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Source,

        [int]$ConfidenceScore,

        [string]$ScoreReason
    )

    if ($ConfidenceScore -le 0 -or [string]::IsNullOrWhiteSpace($ScoreReason)) {
        $scoreEntry = Get-ScoreFromMatrix -Source $Source
        if ($ConfidenceScore -le 0) {
            $ConfidenceScore = [int]$scoreEntry.Score
        }
        if ([string]::IsNullOrWhiteSpace($ScoreReason)) {
            $ScoreReason = [string]$scoreEntry.Reason
        }
    }

    $cleanName = ConvertTo-NormalizedMetadataTitle -Title $Name
    if ([string]::IsNullOrWhiteSpace($cleanName)) {
        return $null
    }

    if (Test-IsJunkTitle -Title $cleanName) {
        return $null
    }

    $confidenceLevel = Get-ConfidenceLevel -ConfidenceScore $ConfidenceScore

    return [PSCustomObject]@{
        Name            = $cleanName
        Source          = $Source
        ConfidenceScore = $ConfidenceScore
        ConfidenceLevel = $confidenceLevel
        ScoreReason     = $ScoreReason
    }
}

function Get-ConfidenceColor {
    param(
        [Parameter(Mandatory)]
        [string]$ConfidenceLevel
    )

    switch ($ConfidenceLevel) {
        'High' { return 'Green' }
        'Medium' { return 'Yellow' }
        default { return 'DarkYellow' }
    }
}

function Write-RenameLiveLine {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('PREVIEW', 'PLANNED', 'RENAMED')]
        [string]$Action,

        [Parameter(Mandatory)]
        [psobject]$Entry
    )

    $confidenceColor = Get-ConfidenceColor -ConfidenceLevel $Entry.ConfidenceLevel
    $confidenceTag = '[{0} {1}%]' -f $Entry.ConfidenceLevel.ToUpperInvariant(), [int]$Entry.ConfidenceScore
    $reasonText = if ($Entry.PSObject.Properties.Name -contains 'ScoreReason') { $Entry.ScoreReason } else { '' }

    Write-Host -NoNewline ('[{0}] ' -f $Action) -ForegroundColor DarkCyan
    Write-Host -NoNewline ($confidenceTag + ' ') -ForegroundColor $confidenceColor
    Write-Host ("{0} -> {1} [{2}] {3}" -f $Entry.OldName, $Entry.NewName, $Entry.Source, $reasonText)

    $logMessage = "{0} {1} {2} -> {3} [{4}] ({5})" -f $Action, $confidenceTag, $Entry.OldName, $Entry.NewName, $Entry.Source, $reasonText
    if ($Action -eq 'RENAMED' -and ($Entry.PSObject.Properties.Name -contains 'FullPath')) {
        $sourcePath = [string]$Entry.FullPath
        $targetPath = Join-Path -Path (Split-Path -Path $Entry.FullPath -Parent) -ChildPath $Entry.NewName
        $logMessage = "{0} | From={1} | To={2}" -f $logMessage, $sourcePath, $targetPath
    }
    Write-RunLog -Message $logMessage
}

function Write-SkippedLiveLine {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [string]$Reason
    )

    Write-Host -NoNewline '[SKIPPED] ' -ForegroundColor DarkYellow
    Write-Host ("{0} [{1}]" -f $FileName, $Reason)
}

function Write-ErrorLiveLine {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [string]$Reason
    )

    Write-Host -NoNewline '[ERROR] ' -ForegroundColor Red
    Write-Host ("{0} [{1}]" -f $FileName, $Reason)
}

function Write-TimeoutLiveLine {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [double]$ElapsedSeconds,

        [Parameter(Mandatory)]
        [int]$TimeoutSeconds
    )

    Write-Host -NoNewline '[SKIPPED] ' -ForegroundColor DarkYellow
    Write-Host (
        "{0} [Per-file timeout exceeded: {1:N1}s > {2}s]" -f $FileName,
        $ElapsedSeconds,
        $TimeoutSeconds
    )
}

function Resolve-ResumeLogPath {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [string]$RequestedPath
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        try {
            return (Resolve-Path -LiteralPath $RequestedPath -ErrorAction Stop).Path
        }
        catch {
            return $null
        }
    }

    try {
        $candidates = Get-ChildItem -LiteralPath $Root -File -Filter 'Rename-Run-*.log' -ErrorAction Stop |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 10
        foreach ($candidate in $candidates) {
            # Skip empty/zero-rename logs that completed normally — they have no useful resume data
            $quickCheck = Get-ResumeCandidateStatus -Path $candidate.FullName
            if ($quickCheck.Eligible -or $quickCheck.StructuredRenamedEntries -gt 0) {
                return $candidate.FullName
            }
        }
        # Fall back to the most recent log if none had rename data
        if ($candidates) {
            return $candidates[0].FullName
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-ResumeCandidateStatus {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $status = [PSCustomObject]@{
        Exists = $false
        StructuredRenamedEntries = 0
        SelectedSettingsDetected = $false
        StopDetected = $false
        EndOfLogDetected = $false
        EndOfLogStatus = $null
        Eligible = $false
    }

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $status
    }

    $status.Exists = $true

    try {
        $content = Get-Content -LiteralPath $Path -ErrorAction Stop
    }
    catch {
        return $status
    }

    foreach ($line in $content) {
        if ($line -match '^\s*Selected settings:\s*$') {
            $status.SelectedSettingsDetected = $true
        }

        if ($line -match 'RENAMED\b.*\|\s*From=.*\|\s*To=.*$') {
            $status.StructuredRenamedEntries++
        }

        if ($line -match '(?i)Stop requested|stopped by user') {
            $status.StopDetected = $true
        }

        if ($line -match '(?i)^\[[0-9]{4}-[0-9]{2}-[0-9]{2}\s+[0-9]{2}:[0-9]{2}:[0-9]{2}\]\s+END_OF_LOG_CODE:\s*(?<status>.*)$') {
            $status.EndOfLogDetected = $true
            $status.EndOfLogStatus = $Matches['status'].Trim()
        }
    }

    $isAbnormalEndMarker = $false
    if ($status.EndOfLogDetected -and -not [string]::IsNullOrWhiteSpace($status.EndOfLogStatus)) {
        if ($status.EndOfLogStatus -match '(?i)abnormal|unhandled exception|crash|terminated|run stopped by user|stop requested') {
            $isAbnormalEndMarker = $true
        }
    }

    $hasResumeData = ($status.StructuredRenamedEntries -gt 0) -or $status.SelectedSettingsDetected
    $status.Eligible = $hasResumeData -and ((-not $status.EndOfLogDetected) -or $isAbnormalEndMarker)
    return $status
}

function Write-EndOfLogCode {
    param(
        [Parameter(Mandatory)]
        [string]$Status
    )

    if ($script:EndOfLogWritten) {
        return
    }

    Write-RunLog -Message ("END_OF_LOG_CODE: {0}" -f $Status)
    $script:EndOfLogWritten = $true
}

trap {
    $errorRecord = $_
    $exceptionMessage = if ($errorRecord.Exception) {
        $errorRecord.Exception.Message
    }
    else {
        $errorRecord.ToString()
    }

    Write-Host ("Unhandled terminating error: {0}" -f $exceptionMessage) -ForegroundColor Red

    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        try {
            Write-RunLog -Message ("UNHANDLED_EXCEPTION: {0}" -f $exceptionMessage)

            if ($errorRecord.InvocationInfo) {
                $invocation = $errorRecord.InvocationInfo
                Write-RunLog -Message (
                    "UNHANDLED_EXCEPTION_LOCATION: Script={0}; Line={1}; Offset={2}; Statement={3}" -f
                    $invocation.ScriptName,
                    $invocation.ScriptLineNumber,
                    $invocation.OffsetInLine,
                    $invocation.Line
                )
            }

            if (-not [string]::IsNullOrWhiteSpace($errorRecord.ScriptStackTrace)) {
                $stackFlat = ($errorRecord.ScriptStackTrace -replace "`r?`n", ' | ')
                Write-RunLog -Message ("UNHANDLED_EXCEPTION_STACK: {0}" -f $stackFlat)
            }

            Write-EndOfLogCode -Status 'Abnormal termination (unhandled exception)'
        }
        catch {
            try {
                $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Add-Content -LiteralPath $script:LogPath -Value ("[{0}] UNHANDLED_EXCEPTION: logging failure while writing crash details." -f $timestamp) -Encoding UTF8
            }
            catch {
            }
        }
    }

    break
}

function Import-ResumeStateFromLog {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $state = [PSCustomObject]@{
        Path         = $Path
        ImportedFrom = 0
        ImportedTo   = 0
    }

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $state
    }

    $pattern = 'RENAMED\b\s*\[(?:[^/]+)/(?<score>\d+)%?\].*?\[(?<source>[^\]]+)\].*\|\s*From=(?<from>.+?)\s*\|\s*To=(?<to>.+?)\s*$'
    $fallbackPattern = 'RENAMED\b.*\|\s*From=(?<from>.+?)\s*\|\s*To=(?<to>.+?)\s*$'

    try {
        foreach ($line in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
            $score = $null
            $source = $null

            if ($line -match $pattern) {
                $score = [int]$Matches['score']
                $source = $Matches['source'].Trim()
            }
            elseif ($line -match $fallbackPattern) {
                # Older log format without parseable score
            }
            else {
                continue
            }

            $fromPath = $Matches['from'].Trim()
            $toPath = $Matches['to'].Trim()

            if (-not [string]::IsNullOrWhiteSpace($fromPath)) {
                $null = $script:ResumeRenamedFrom.Add($fromPath)
            }
            if (-not [string]::IsNullOrWhiteSpace($toPath)) {
                $null = $script:ResumeRenamedTo.Add($toPath)
                if ($null -ne $score) {
                    $script:ResumeRenamedScores[$toPath] = [PSCustomObject]@{ Score = $score; Source = $source }
                }
            }
        }
    }
    catch {
        return $state
    }

    $state.ImportedFrom = $script:ResumeRenamedFrom.Count
    $state.ImportedTo = $script:ResumeRenamedTo.Count
    return $state
}

function Import-SelectedSettingsFromLog {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $settings = [ordered]@{}

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $settings
    }

    $inSettingsSection = $false

    try {
        foreach ($line in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
            if (-not $inSettingsSection) {
                if ($line -match '^\s*Selected settings:\s*$') {
                    $inSettingsSection = $true
                }
                continue
            }

            if ($line -match '^\s*-\s*(?<key>[^:]+):\s*(?<value>.*)$') {
                $key = $Matches['key'].Trim()
                $value = $Matches['value'].Trim()
                if (-not [string]::IsNullOrWhiteSpace($key)) {
                    $settings[$key] = $value
                }
                continue
            }

            if ($line -match '^\s*$' -or $line -match '^\[') {
                break
            }
        }
    }
    catch {
        return [ordered]@{}
    }

    return $settings
}

function Convert-LogSettingToBool {
    param(
        [string]$Value,

        [bool]$Default = $false
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    switch ($Value.Trim().ToLowerInvariant()) {
        'true' { return $true }
        'false' { return $false }
        'yes' { return $true }
        'no' { return $false }
        '1' { return $true }
        '0' { return $false }
        default { return $Default }
    }
}

function Convert-LogSettingToInt {
    param(
        [string]$Value,

        [int]$Default = 0,

        [int]$Max = [int]::MaxValue
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    $parsed = 0
    if ([int]::TryParse($Value.ToString().Trim(), [ref]$parsed)) {
        if ($parsed -ge 1 -and $parsed -le $Max) {
            return $parsed
        }
    }

    return $Default
}

function Set-SelectedSettingsFromLog {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Settings,

        [Parameter(Mandatory)]
        [hashtable]$BoundParameters
    )

    $appliedKeys = [System.Collections.Generic.List[string]]::new()

    if ($Settings.Count -eq 0) {
        return @($appliedKeys)
    }

    if (-not $BoundParameters.ContainsKey('Apply') -and $Settings.Contains('Apply')) {
        $script:Apply = Convert-LogSettingToBool -Value $Settings['Apply'] -Default $false
        $appliedKeys.Add('Apply') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('ImmediateRename') -and $Settings.Contains('ImmediateRename')) {
        $script:ImmediateRename = Convert-LogSettingToBool -Value $Settings['ImmediateRename'] -Default $false
        $appliedKeys.Add('ImmediateRename') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('QuickPreview') -and $Settings.Contains('QuickPreview')) {
        $script:QuickPreview = Convert-LogSettingToBool -Value $Settings['QuickPreview'] -Default $false
        $appliedKeys.Add('QuickPreview') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('QuickPreviewLimit') -and $Settings.Contains('QuickPreviewLimit')) {
        $limitRaw = $Settings['QuickPreviewLimit']
        $parsedLimit = Convert-LogSettingToInt -Value $limitRaw -Default 200 -Max 1000000
        if ($parsedLimit -ge 1 -and $parsedLimit -le 1000000) {
            $script:QuickPreviewLimit = $parsedLimit
            $appliedKeys.Add('QuickPreviewLimit') | Out-Null
        }
    }

    if (-not $BoundParameters.ContainsKey('NoRecurse') -and $Settings.Contains('NoRecurse')) {
        $script:NoRecurse = Convert-LogSettingToBool -Value $Settings['NoRecurse'] -Default $false
        $appliedKeys.Add('NoRecurse') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('InteractiveSubfolders') -and $Settings.Contains('InteractiveSubfolders')) {
        $script:InteractiveSubfolders = Convert-LogSettingToBool -Value $Settings['InteractiveSubfolders'] -Default $false
        $appliedKeys.Add('InteractiveSubfolders') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('IncludeHidden') -and $Settings.Contains('IncludeHidden')) {
        $script:IncludeHidden = Convert-LogSettingToBool -Value $Settings['IncludeHidden'] -Default $false
        $appliedKeys.Add('IncludeHidden') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('Force') -and $Settings.Contains('Force')) {
        $script:Force = Convert-LogSettingToBool -Value $Settings['Force'] -Default $false
        $appliedKeys.Add('Force') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('OnlyGeneric') -and $Settings.Contains('OnlyGeneric')) {
        $script:OnlyGeneric = Convert-LogSettingToBool -Value $Settings['OnlyGeneric'] -Default $false
        $appliedKeys.Add('OnlyGeneric') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('IntelligentScan') -and $Settings.Contains('IntelligentScan')) {
        $script:IntelligentScan = Convert-LogSettingToBool -Value $Settings['IntelligentScan'] -Default $false
        $appliedKeys.Add('IntelligentScan') | Out-Null
    }

    if ($Settings.Contains('UseExtensionFolders')) {
        $script:UseExtensionFolders = Convert-LogSettingToBool -Value $Settings['UseExtensionFolders'] -Default $false
        $appliedKeys.Add('UseExtensionFolders') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('IncludeOtherFiles') -and $Settings.Contains('IncludeOtherFiles')) {
        $script:IncludeOtherFiles = Convert-LogSettingToBool -Value $Settings['IncludeOtherFiles'] -Default $false
        $appliedKeys.Add('IncludeOtherFiles') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('ScanOtherFilesAtEnd') -and $Settings.Contains('ScanOtherFilesAtEnd')) {
        $script:ScanOtherFilesAtEnd = Convert-LogSettingToBool -Value $Settings['ScanOtherFilesAtEnd'] -Default $false
        $appliedKeys.Add('ScanOtherFilesAtEnd') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('ListUncommonFiles') -and $Settings.Contains('ListUncommonFiles')) {
        $script:ListUncommonFiles = Convert-LogSettingToBool -Value $Settings['ListUncommonFiles'] -Default $false
        $appliedKeys.Add('ListUncommonFiles') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('SkipLikelyAlreadyRenamed') -and $Settings.Contains('SkipLikelyAlreadyRenamed')) {
        $script:SkipLikelyAlreadyRenamed = Convert-LogSettingToBool -Value $Settings['SkipLikelyAlreadyRenamed'] -Default $true
        $appliedKeys.Add('SkipLikelyAlreadyRenamed') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('PerFileTimeoutSeconds') -and $Settings.Contains('PerFileTimeoutSeconds')) {
        $timeoutRaw = $Settings['PerFileTimeoutSeconds']
        $parsedTimeout = Convert-LogSettingToInt -Value $timeoutRaw -Default 20 -Max 600
        if ($parsedTimeout -ge 5 -and $parsedTimeout -le 600) {
            $script:PerFileTimeoutSeconds = $parsedTimeout
            $appliedKeys.Add('PerFileTimeoutSeconds') | Out-Null
        }
    }

    if ($script:InteractiveSubfolders -and $script:NoRecurse) {
        $script:NoRecurse = $false
    }

    if ($script:IncludeOtherFiles) {
        $script:IntelligentScan = $false
        $script:ScanOtherFilesAtEnd = $false
    }

    return @($appliedKeys)
}

function Import-ErrorStateFromLog {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $state = [PSCustomObject]@{
        Path = $Path
        ImportedFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        ImportedCount = 0
    }

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $state
    }

    $structuredPattern = 'ERROR\s*\|\s*Phase=(?<phase>[^|]+)\|\s*File=(?<file>[^|]+)\|\s*NewName=(?<newname>[^|]*)\|\s*Message=' 
    $legacyProcessingPattern = 'ErrorSkip\s+File\s+processing\s+failed\s+for\s+(?<file>.+?)\.\s+Error='
    $legacyRenamePattern = 'ErrorSkip\s+Rename\s+failed\s+for\s+(?<file>.+?)\s+->\s+.+?\.\s+Error='

    try {
        foreach ($line in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
            $filePath = $null

            if ($line -match $structuredPattern) {
                $filePath = $Matches['file'].Trim()
            }
            elseif ($line -match $legacyProcessingPattern) {
                $filePath = $Matches['file'].Trim()
            }
            elseif ($line -match $legacyRenamePattern) {
                $filePath = $Matches['file'].Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($filePath)) {
                $null = $state.ImportedFiles.Add($filePath)
            }
        }
    }
    catch {
        return $state
    }

    $state.ImportedCount = $state.ImportedFiles.Count
    return $state
}

function Test-LikelyAlreadyRenamedName {
    param(
        [Parameter(Mandatory)]
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return $false
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        return $false
    }

    $patterns = @(
        '^(Audio|Video|Photo|Document)\s+\d{4}-\d{2}-\d{2}(?:\s+\d{2}-\d{2}(?:-\d{2})?)?(?:\s+\(\d+\))?$',
        '.*\s+\(\d+\)$',
        '^(Location\s+.+|.+\s+-\s+Location\s+.+)$',
        '.*_renamed(?:\s+\(\d+\))?$'
    )

    foreach ($pattern in $patterns) {
        if ($baseName -match $pattern) {
            return $true
        }
    }

    return $false
}

function Test-IsRunArtifactFile {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,

        [string]$CurrentRunLogPath,

        [string]$CurrentReportPath,

        [string]$CurrentCsvPath,

        [string]$ResumeLogPath
    )

    if ($CurrentRunLogPath -and ($File.FullName -ieq $CurrentRunLogPath)) { return $true }
    if ($CurrentReportPath -and ($File.FullName -ieq $CurrentReportPath)) { return $true }
    if ($CurrentCsvPath -and ($File.FullName -ieq $CurrentCsvPath)) { return $true }
    if ($ResumeLogPath -and ($File.FullName -ieq $ResumeLogPath)) { return $true }

    return (
        ($File.Name -match '^Rename-Run-\d{8}-\d{6}\.log$') -or
        ($File.Name -match '^Rename-Report-\d{8}-\d{6}\.txt$') -or
        ($File.Name -match '^Rename-Report-\d{8}-\d{6}\.csv$')
    )
}

function Test-IsCommonUserFileType {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    $ext = $File.Extension.ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($ext)) {
        return $false
    }

    $commonExtensions = @(
        '.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.pdf', '.ppt', '.pptx', '.xls', '.xlsx', '.txt', '.rtf', '.odt', '.ods', '.odp', '.xml',
        '.jpg', '.jpeg', '.jfif', '.png', '.heic', '.tif', '.tiff', '.gif', '.bmp', '.webp', '.raw', '.cr2', '.nef', '.arw', '.dng', '.ico', '.avif',
        '.mp3', '.wav', '.flac', '.aac', '.m4a', '.ogg', '.wma', '.aiff', '.alac', '.opus',
        '.mp4', '.mov', '.m4v', '.avi', '.mkv', '.wmv', '.flv', '.webm', '.mpg', '.mpeg', '.3gp', '.mts', '.m2ts', '.ts'
    )

    return ($ext -in $commonExtensions)
}

function Get-SafeFileName {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalidChars) {
        $Name = $Name.Replace($char, '_')
    }

    $Name = ($Name -replace '\s+', ' ').Trim()
    $Name = $Name.Trim('.', ' ')

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    return $Name
}

function Show-CommandLineHelp {
    $scriptName = Split-Path -Leaf $PSCommandPath
    $exeName = [System.IO.Path]::ChangeExtension($scriptName, '.exe')

    Write-Host ''
    Write-Host 'Rename-ByMetadata - Parameter Help' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'PowerShell usage:' -ForegroundColor Yellow
    Write-Host ("  .\{0} [-Path <folder>] [-Apply] [-ImmediateRename] [-QuickPreview] [-QuickPreviewLimit <n>]" -f $scriptName)
    Write-Host '               [-NoRecurse] [-InteractiveSubfolders] [-IncludeHidden] [-OnlyGeneric] [-Force]'
    Write-Host '               [-ReportPath <txt>] [-CsvPath <csv>] [-LogPath <log>]'
    Write-Host '               [-Resume] [-ResumeLogPath <log>] [-RetryErrors] [-RetryErrorsLogPath <log>]'
    Write-Host '               [-IntelligentScan] [-IncludeOtherFiles] [-ScanOtherFilesAtEnd] [-ListUncommonFiles] [-NoAutoResumePrompt]'
    Write-Host '               [-SkipLikelyAlreadyRenamed <$true/$false>] [-PerFileTimeoutSeconds <5-600>] [-BatchSize <100-50000>]'
    Write-Host '               [-Unattended] [-UseWordCom] [-MaxFileReadMB <1-500>] [-ComRecycleInterval <50-10000>]'
    Write-Host ''
    Write-Host 'EXE help shortcuts:' -ForegroundColor Yellow
    Write-Host ("  .\{0} /?" -f $exeName)
    Write-Host ("  .\{0} /help" -f $exeName)
    Write-Host ''
    Write-Host 'Common examples:' -ForegroundColor Yellow
    Write-Host ("  .\{0} -Path `"D:\Reconstructed`"" -f $scriptName)
    Write-Host ("  .\{0} -Path `"D:\Reconstructed`" -Apply" -f $scriptName)
    Write-Host ("  .\{0} -Path `"D:\Reconstructed`" -QuickPreview -QuickPreviewLimit 50" -f $scriptName)
    Write-Host ("  .\{0} -Path `"D:\Reconstructed`" -BatchSize 500" -f $scriptName)
    Write-Host ''
    Write-Host 'Unattended weekend run (hardened, no prompts):' -ForegroundColor Yellow
    Write-Host ("  .\{0} -Path `"D:\CarolynR\Reconstructed`" -Apply -ImmediateRename -Resume -NoAutoResumePrompt -IntelligentScan -Unattended" -f $scriptName)
    Write-Host ''
    Write-Host 'Notes:' -ForegroundColor DarkGray
    Write-Host '  Word COM is disabled by default (freeze-prone). Use -UseWordCom to opt back in.' -ForegroundColor DarkGray
    Write-Host '  Already-renamed file skipping is on by default. Use -SkipLikelyAlreadyRenamed:$false to disable.' -ForegroundColor DarkGray
    Write-Host ''
}

$helpTokens = @('/?', '/help', '-?', '-help', '--help')
$pathLooksLikeHelp = -not [string]::IsNullOrWhiteSpace($Path) -and ($helpTokens -contains $Path.Trim().ToLowerInvariant())
$rawArgsContainSlashHelp = $false
$commandLineContainsHelp = $false
$invocationLineContainsHelp = $false

foreach ($arg in [System.Environment]::GetCommandLineArgs()) {
    if ($helpTokens -contains $arg.Trim().ToLowerInvariant()) {
        $rawArgsContainSlashHelp = $true
        break
    }
}

try {
    $commandLineText = [System.Environment]::CommandLine
    if (-not [string]::IsNullOrWhiteSpace($commandLineText) -and ($commandLineText -match '(?i)(?:^|\s)/(?:\?|help)(?=\s|$)')) {
        $commandLineContainsHelp = $true
    }
}
catch {
    $commandLineContainsHelp = $false
}

try {
    $invocationLine = $MyInvocation.Line
    if (-not [string]::IsNullOrWhiteSpace($invocationLine) -and ($invocationLine -match '(?i)(?:^|\s)/(?:\?|help)(?=\s|$)')) {
        $invocationLineContainsHelp = $true
    }
}
catch {
    $invocationLineContainsHelp = $false
}

if ($Help -or $pathLooksLikeHelp -or $rawArgsContainSlashHelp -or $commandLineContainsHelp -or $invocationLineContainsHelp) {
    Show-CommandLineHelp
    return
}

# --- BackfillRenameSuffix: add _renamed to files from a prior run log ---
if ($BackfillRenameSuffix) {
    $backfillLog = $BackfillLogPath
    if ([string]::IsNullOrWhiteSpace($backfillLog) -and -not [string]::IsNullOrWhiteSpace($Path)) {
        $backfillLog = (Resolve-ResumeLogPath -Root (Resolve-Path -LiteralPath $Path).Path -RequestedPath $null)
    }
    if ([string]::IsNullOrWhiteSpace($backfillLog) -or -not (Test-Path -LiteralPath $backfillLog)) {
        Write-Host 'BackfillRenameSuffix: No log file found. Use -BackfillLogPath to specify one.' -ForegroundColor Red
        return
    }
    Write-Host ("Backfilling _renamed suffix from log: {0}" -f $backfillLog) -ForegroundColor Cyan
    $pattern = 'RENAMED\b.*\|\s*To=(?<to>.+?)\s*$'
    $backfillRenamed = 0
    $backfillSkipped = 0
    $backfillMissing = 0
    foreach ($line in (Get-Content -LiteralPath $backfillLog -Encoding UTF8)) {
        if ($line -match $pattern) {
            $toPath = $Matches['to']
            if (-not (Test-Path -LiteralPath $toPath)) {
                $backfillMissing++
                continue
            }
            $toBase = [System.IO.Path]::GetFileNameWithoutExtension($toPath)
            $toExt  = [System.IO.Path]::GetExtension($toPath)
            if ($toBase -match '_renamed(\s+\(\d+\))?$') {
                $backfillSkipped++
                continue
            }
            $newBase = $toBase + '_renamed'
            $dir = [System.IO.Path]::GetDirectoryName($toPath)
            $candidate = Join-Path $dir ($newBase + $toExt)
            $counter = 1
            while (Test-Path -LiteralPath $candidate) {
                $candidate = Join-Path $dir ("{0} ({1}){2}" -f $newBase, $counter, $toExt)
                $counter++
            }
            $newName = [System.IO.Path]::GetFileName($candidate)
            if ($PSCmdlet.ShouldProcess($toPath, "Backfill rename to '$newName'")) {
                try {
                    Rename-Item -LiteralPath $toPath -NewName $newName -ErrorAction Stop
                    $backfillRenamed++
                }
                catch {
                    Write-Host ("  ERROR backfilling {0}: {1}" -f $toPath, $_.Exception.Message) -ForegroundColor Red
                }
            }
        }
    }
    Write-Host ("Backfill complete. Renamed: {0} | AlreadySuffixed: {1} | Missing: {2}" -f $backfillRenamed, $backfillSkipped, $backfillMissing) -ForegroundColor Green
    return
}

# Initialize hardening modes from parameters
if ($Unattended) {
    $script:UnattendedMode = $true
    Write-Host 'Unattended mode enabled. All interactive prompts will auto-answer with defaults.' -ForegroundColor Yellow
}
if ($UseWordCom) {
    $script:NoWordComMode = $false
    Write-Host 'Word COM automation enabled by -UseWordCom. This may cause hangs on corrupt files.' -ForegroundColor Yellow
}
$script:MaxFileReadBytes = [long]$MaxFileReadMB * 1024 * 1024
$script:ComRecycleInterval = $ComRecycleInterval
$script:PerFileTimeoutSeconds = $PerFileTimeoutSeconds

$retryErrorsRequested = $RetryErrors.IsPresent -or -not [string]::IsNullOrWhiteSpace($RetryErrorsLogPath)

$defaultPath = Get-LastUsedScanPath
if ([string]::IsNullOrWhiteSpace($defaultPath)) {
    $defaultPath = $PWD.Path
}

if ([string]::IsNullOrWhiteSpace($Path)) {
    if ($script:UnattendedMode) {
        $Path = $defaultPath
    } else {
    $enteredPath = Read-Host "Enter folder path to scan (press Enter for default: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($enteredPath)) {
        $Path = $defaultPath
    }
    else {
        $Path = $enteredPath.Trim()
    }
    }
}

$resolvedPath = Resolve-Path -LiteralPath $Path
$root = $resolvedPath.Path
Save-LastUsedScanPath -Path $root
$latestRunLog = Resolve-ResumeLogPath -Root $root -RequestedPath $null
$latestRunLogStatus = $null

if (-not [string]::IsNullOrWhiteSpace($latestRunLog)) {
    $latestRunLogStatus = Get-ResumeCandidateStatus -Path $latestRunLog
    $isAbnormalPriorRun = ($latestRunLogStatus.EndOfLogDetected -and -not [string]::IsNullOrWhiteSpace($latestRunLogStatus.EndOfLogStatus) -and ($latestRunLogStatus.EndOfLogStatus -match '(?i)abnormal|unhandled exception|crash|terminated'))
    if ($latestRunLogStatus.Exists -and ((-not $latestRunLogStatus.EndOfLogDetected) -or $isAbnormalPriorRun)) {
        $terminationHint = if ($latestRunLogStatus.StopDetected) {
            'Previous run log did not include a completion marker and appears user-stopped.'
        }
        elseif ($isAbnormalPriorRun) {
            'Previous run ended abnormally based on end-of-log status.'
        }
        else {
            'Previous run log did not include a completion marker and appears abnormally terminated.'
        }

        Write-Host $terminationHint -ForegroundColor Yellow
        Write-Host ("Previous run log: {0}" -f $latestRunLog) -ForegroundColor DarkYellow
    }
}

$resumeRequested = $Resume.IsPresent -or -not [string]::IsNullOrWhiteSpace($ResumeLogPath)
$resumeRequestedViaPrompt = $false
$resumeSettingsApplied = $false
$retrySettingsApplied = $false

if (-not $resumeRequested -and -not $NoAutoResumePrompt.IsPresent) {
    if (-not [string]::IsNullOrWhiteSpace($latestRunLog)) {
        $candidateStatus = if ($latestRunLogStatus) { $latestRunLogStatus } else { Get-ResumeCandidateStatus -Path $latestRunLog }
        if ($candidateStatus.Exists) {
            if ($candidateStatus.Eligible) {
                $promptMsg = "Previous run appears interrupted ({0} completed renames). Resume from '{1}'?" -f $candidateStatus.StructuredRenamedEntries, $latestRunLog
            }
            else {
                $promptMsg = "Previous run log found ({0} completed renames). Resume from '{1}'?" -f $candidateStatus.StructuredRenamedEntries, $latestRunLog
            }

            $resumePrompt = Read-YesNo -Prompt ($promptMsg) -DefaultYes $true

            if ($resumePrompt) {
                $resumeRequested = $true
                $resumeRequestedViaPrompt = $true
                $ResumeLogPath = $latestRunLog
            }
        }
    }
}

if ($resumeRequested) {
    $resolvedResumeSettingsLog = Resolve-ResumeLogPath -Root $root -RequestedPath $ResumeLogPath
    if (-not [string]::IsNullOrWhiteSpace($resolvedResumeSettingsLog)) {
        $selectedSettings = Import-SelectedSettingsFromLog -Path $resolvedResumeSettingsLog
        if ($selectedSettings.Count -gt 0) {
            # Display previous settings so the user can decide whether to keep them
            Write-Host ''
            Write-Host 'Previous run settings:' -ForegroundColor Cyan

            # Show human-readable scan mode summary
            $scanModeLabel = 'All files (no filtering)'
            if ($selectedSettings.Contains('UseExtensionFolders') -and (Convert-LogSettingToBool -Value $selectedSettings['UseExtensionFolders'] -Default $false)) {
                $scanModeLabel = 'Folder names as extensions'
            }
            elseif ($selectedSettings.Contains('IncludeOtherFiles') -and (Convert-LogSettingToBool -Value $selectedSettings['IncludeOtherFiles'] -Default $false)) {
                $scanModeLabel = 'All files (including uncommon types)'
            }
            elseif ($selectedSettings.Contains('IntelligentScan') -and (Convert-LogSettingToBool -Value $selectedSettings['IntelligentScan'] -Default $false)) {
                $endScan = $selectedSettings.Contains('ScanOtherFilesAtEnd') -and (Convert-LogSettingToBool -Value $selectedSettings['ScanOtherFilesAtEnd'] -Default $false)
                $scanModeLabel = if ($endScan) { 'Intelligent scan + uncommon at end' } else { 'Intelligent scan (common files only)' }
            }
            Write-Host ("  Scan mode: {0}" -f $scanModeLabel) -ForegroundColor Gray

            # Show remaining key settings
            $skipKeys = @('IntelligentScan', 'UseExtensionFolders', 'IncludeOtherFiles', 'ScanOtherFilesAtEnd',
                          'ResumeRequested', 'ResumeActive', 'ResumeLogPath', 'RetryErrorsRequested',
                          'RetryErrorsActive', 'RetryErrorsLogPath', 'ReportPath', 'CsvPath', 'RequestedLogPath')
            foreach ($k in $selectedSettings.Keys) {
                if ($k -in $skipKeys) { continue }
                Write-Host ("  {0}: {1}" -f $k, $selectedSettings[$k]) -ForegroundColor DarkGray
            }
            Write-Host ''

            $usePreviousSettings = Read-YesNo -Prompt (
                "Keep these settings? (No = choose new settings, resume data is kept either way)"
            ) -DefaultYes $true

            if ($usePreviousSettings) {
                $appliedSettings = Set-SelectedSettingsFromLog -Settings $selectedSettings -BoundParameters $PSBoundParameters
                $resumeSettingsApplied = @($appliedSettings).Count -gt 0
                if ($resumeSettingsApplied) {
                    Write-Host (
                        "Loaded {0} setting(s) from resume log: {1}" -f
                        @($appliedSettings).Count,
                        $resolvedResumeSettingsLog
                    ) -ForegroundColor Yellow
                }
            }
            else {
                Write-Host 'Previous settings declined. You will be prompted for new settings.' -ForegroundColor Yellow
            }
        }
    }
}

if ($retryErrorsRequested -and -not $resumeRequested) {
    $resolvedRetrySettingsLog = Resolve-ResumeLogPath -Root $root -RequestedPath $RetryErrorsLogPath
    if (-not [string]::IsNullOrWhiteSpace($resolvedRetrySettingsLog)) {
        $retrySettings = Import-SelectedSettingsFromLog -Path $resolvedRetrySettingsLog
        if ($retrySettings.Count -gt 0) {
            $applyRetrySettings = $true
            if ([string]::IsNullOrWhiteSpace($RetryErrorsLogPath)) {
                $applyRetrySettings = Read-YesNo -Prompt (
                    "RetryErrors log includes previous job settings. Continue with those settings?"
                ) -DefaultYes $true
            }

            if ($applyRetrySettings) {
                $appliedRetrySettings = Set-SelectedSettingsFromLog -Settings $retrySettings -BoundParameters $PSBoundParameters
                $retrySettingsApplied = @($appliedRetrySettings).Count -gt 0
                if ($retrySettingsApplied) {
                    Write-Host (
                        "Loaded {0} setting(s) from retry log: {1}" -f
                        @($appliedRetrySettings).Count,
                        $resolvedRetrySettingsLog
                    ) -ForegroundColor Yellow
                }
            }
        }
    }
}

$logSettingsApplied = $resumeSettingsApplied -or $retrySettingsApplied

if ($InteractiveSubfolders -and $NoRecurse) {
    throw 'Cannot use -InteractiveSubfolders together with -NoRecurse.'
}

if (-not $PSBoundParameters.ContainsKey('Apply') -and -not $logSettingsApplied) {
    $previewFirst = Read-YesNo -Prompt 'Start in preview mode first (recommended)?' -DefaultYes $true
    if ($previewFirst) {
        $Apply = $false
    }
    else {
        $Apply = $true
    }
}

if ($Apply -and -not $PSBoundParameters.ContainsKey('ImmediateRename') -and -not $logSettingsApplied) {
    $ImmediateRename = Read-YesNo -Prompt 'In rename mode, start renaming immediately (skip planned list first)?' -DefaultYes $false
}

if ($ImmediateRename -and -not $Apply) {
    throw 'Cannot use -ImmediateRename unless apply mode is selected.'
}

if ($IncludeOtherFiles -and $ScanOtherFilesAtEnd) {
    throw 'Cannot combine -IncludeOtherFiles with -ScanOtherFilesAtEnd.'
}

$isApplyMode = [bool]$Apply
$isImmediateRenameMode = $isApplyMode -and [bool]$ImmediateRename

if (-not $PSBoundParameters.ContainsKey('QuickPreview') -and -not $logSettingsApplied) {
    if ($script:UnattendedMode) {
        $QuickPreview = $false
    } else {
    $useQuickPreview = Read-YesNo -Prompt 'Use quick preview sample mode?' -DefaultYes $true
    if ($useQuickPreview) {
        $QuickPreview = $true

        if (-not $PSBoundParameters.ContainsKey('QuickPreviewLimit')) {
            $enteredLimit = Read-Host "Quick preview file limit (press Enter for default: $QuickPreviewLimit)"
            if (-not [string]::IsNullOrWhiteSpace($enteredLimit)) {
                $parsedLimit = 0
                if ([int]::TryParse($enteredLimit.Trim(), [ref]$parsedLimit) -and $parsedLimit -ge 1 -and $parsedLimit -le 100000) {
                    $QuickPreviewLimit = $parsedLimit
                }
                else {
                    Write-Host 'Invalid quick preview limit. Using default.' -ForegroundColor Yellow
                }
            }
        }
    }
    else {
        $QuickPreview = $false
    }
    }
}

$intelligentScanEnabled = [bool]$IntelligentScan
$includeOtherFilesNow = [bool]$IncludeOtherFiles
$scanOtherFilesAtEndEnabled = [bool]$ScanOtherFilesAtEnd

if ($includeOtherFilesNow) {
    $intelligentScanEnabled = $false
    $scanOtherFilesAtEndEnabled = $false
}



# Add a new autodiscovery mode: F = Use folder names as extensions
$useExtensionFolders = if ($logSettingsApplied -and $script:UseExtensionFolders) { $true } else { $false }

# If extension-folder mode was restored from prior settings but the current root
# contains no matching extension-named subfolders, confirm with the user before
# continuing; default to NO to avoid an empty scan.
if ($useExtensionFolders -and $logSettingsApplied) {
    $commonExtFolders = @(
        'doc','docx','docm','dot','dotx','dotm','pdf','ppt','pptx','xls','xlsx','rtf','odt','ods','odp','xml',
        'jpg','jpeg','jfif','png','heic','tif','tiff','gif','bmp','webp','raw','cr2','nef','arw','dng','ico','avif',
        'mp3','wav','flac','aac','m4a','ogg','wma','aiff','alac','opus',
        'mp4','mov','m4v','avi','mkv','wmv','flv','webm','mpg','mpeg','3gp','mts','m2ts','ts'
    )
    try {
        $foundExtFolders = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $commonExtFolders -contains $_.Name.ToLowerInvariant() }
    } catch { $foundExtFolders = @() }

    if (-not $foundExtFolders -or $foundExtFolders.Count -eq 0) {
        $confirmExtMode = Read-YesNo -Prompt ("No extension-named subfolders found in '{0}'. Continue with folder-name-as-extension scan?" -f $root) -DefaultYes $false
        if (-not $confirmExtMode) { $useExtensionFolders = $false }
    }
}
if ($script:UnattendedMode -and -not $PSBoundParameters.ContainsKey('IntelligentScan') -and -not $PSBoundParameters.ContainsKey('IncludeOtherFiles') -and -not $logSettingsApplied) {
    $intelligentScanEnabled = $true
    $useExtensionFolders = $false
} elseif (-not $retryErrorsRequested -and -not $PSBoundParameters.ContainsKey('IntelligentScan') -and -not $PSBoundParameters.ContainsKey('IncludeOtherFiles') -and -not $logSettingsApplied) {
    while ($true) {
        $scanPrompt = Read-Host 'File type scan mode: [Y]es (intelligent scan), [N]o (all files), or [F]older names as extensions? [Y/n/f]'
        if ([string]::IsNullOrWhiteSpace($scanPrompt)) { $scanPrompt = 'y' }
        switch ($scanPrompt.Trim().ToLowerInvariant()) {
            'y' { $intelligentScanEnabled = $true; $useExtensionFolders = $false; break }
            'yes' { $intelligentScanEnabled = $true; $useExtensionFolders = $false; break }
            'n' { $intelligentScanEnabled = $false; $useExtensionFolders = $false; break }
            'no' { $intelligentScanEnabled = $false; $useExtensionFolders = $false; break }
            'f' { $intelligentScanEnabled = $true; $useExtensionFolders = $true; break }
            'folder' { $intelligentScanEnabled = $true; $useExtensionFolders = $true; break }
            default { Write-Host 'Please answer y, n, or f.' -ForegroundColor Yellow }
        }
        if ($scanPrompt.Trim().ToLowerInvariant() -in @('y','yes','n','no','f','folder')) { break }
    }
}

# Streamlined F mode: always recurse, suppress subfolder prompts
if ($useExtensionFolders) {
    $NoRecurse = $false
    $InteractiveSubfolders = $false
} else {
    if ($intelligentScanEnabled -and -not $includeOtherFilesNow -and -not $PSBoundParameters.ContainsKey('ScanOtherFilesAtEnd') -and -not $logSettingsApplied -and -not $useExtensionFolders) {
        $scanOtherFilesAtEndEnabled = Read-YesNo -Prompt 'After common files, scan uncommon file types at end?' -DefaultYes $false
    }

    if (-not $PSBoundParameters.ContainsKey('NoRecurse') -and -not $logSettingsApplied) {
        if ($script:UnattendedMode) {
            $NoRecurse = $false
            $InteractiveSubfolders = $false
        } elseif (-not $PSBoundParameters.ContainsKey('InteractiveSubfolders')) {
            :SubfolderModePrompt while ($true) {
                $subfolderMode = Read-Host 'Subfolder scan mode: [Y]es all, [N]o, or [I]nteractive per folder? [Y/n/i]'

                if ([string]::IsNullOrWhiteSpace($subfolderMode)) {
                    $subfolderMode = 'y'
                }

                switch ($subfolderMode.Trim().ToLowerInvariant()) {
                    'y' {
                        $NoRecurse = $false
                        $InteractiveSubfolders = $false
                        break SubfolderModePrompt
                    }
                    'yes' {
                        $NoRecurse = $false
                        $InteractiveSubfolders = $false
                        break SubfolderModePrompt
                    }
                    'n' {
                        $NoRecurse = $true
                        $InteractiveSubfolders = $false
                        break SubfolderModePrompt
                    }
                    'no' {
                        $NoRecurse = $true
                        $InteractiveSubfolders = $false
                        break SubfolderModePrompt
                    }
                    'i' {
                        $NoRecurse = $false
                        $InteractiveSubfolders = $true
                        break SubfolderModePrompt
                    }
                    'interactive' {
                        $NoRecurse = $false
                        $InteractiveSubfolders = $true
                        break SubfolderModePrompt
                    }
                    default {
                        Write-Host 'Please answer y, n, or i.' -ForegroundColor Yellow
                    }
                }
            }
        }
    }
}

if ($InteractiveSubfolders) {
    $NoRecurse = $false
}

$isAdmin = Test-IsAdministrator
Write-Host ("Current session is running as admin: {0}" -f $isAdmin)

$wantsAdmin = Read-YesNo -Prompt 'Run with administrator privileges? (Usually only needed for protected folders)' -DefaultYes $false
if ($wantsAdmin -and -not $isAdmin) {
    $argsList = [System.Collections.Generic.List[string]]::new()
    $argsList.Add('-NoProfile') | Out-Null
    $argsList.Add('-ExecutionPolicy') | Out-Null
    $argsList.Add('Bypass') | Out-Null
    $argsList.Add('-File') | Out-Null
    $argsList.Add(('"{0}"' -f $PSCommandPath)) | Out-Null
    $argsList.Add('-Path') | Out-Null
    $argsList.Add(('"{0}"' -f $Path)) | Out-Null

    if ($NoRecurse) { $argsList.Add('-NoRecurse') | Out-Null }
    if ($InteractiveSubfolders) { $argsList.Add('-InteractiveSubfolders') | Out-Null }
    if ($IncludeHidden) { $argsList.Add('-IncludeHidden') | Out-Null }
    if ($Force) { $argsList.Add('-Force') | Out-Null }
    if ($OnlyGeneric) { $argsList.Add('-OnlyGeneric') | Out-Null }
    if ($Apply) { $argsList.Add('-Apply') | Out-Null }
    if ($ImmediateRename) { $argsList.Add('-ImmediateRename') | Out-Null }
    if ($QuickPreview) {
        $argsList.Add('-QuickPreview') | Out-Null
        $argsList.Add('-QuickPreviewLimit') | Out-Null
        $argsList.Add([string]$QuickPreviewLimit) | Out-Null
    }
    if ($ReportPath) {
        $argsList.Add('-ReportPath') | Out-Null
        $argsList.Add(('"{0}"' -f $ReportPath)) | Out-Null
    }
    if ($CsvPath) {
        $argsList.Add('-CsvPath') | Out-Null
        $argsList.Add(('"{0}"' -f $CsvPath)) | Out-Null
    }
    if ($LogPath) {
        $argsList.Add('-LogPath') | Out-Null
        $argsList.Add(('"{0}"' -f $LogPath)) | Out-Null
    }
    if ($Resume.IsPresent) {
        $argsList.Add('-Resume') | Out-Null
    }
    if ($ResumeLogPath) {
        $argsList.Add('-ResumeLogPath') | Out-Null
        $argsList.Add(('"{0}"' -f $ResumeLogPath)) | Out-Null
    }
    if ($NoAutoResumePrompt.IsPresent) {
        $argsList.Add('-NoAutoResumePrompt') | Out-Null
    }
    if ($SkipLikelyAlreadyRenamed) {
        $argsList.Add('-SkipLikelyAlreadyRenamed') | Out-Null
        $argsList.Add('$true') | Out-Null
    } else {
        $argsList.Add('-SkipLikelyAlreadyRenamed') | Out-Null
        $argsList.Add('$false') | Out-Null
    }
    if ($SkipMetadataUpgrade) {
        $argsList.Add('-SkipMetadataUpgrade') | Out-Null
    }
    if ($UseWordCom) {
        $argsList.Add('-UseWordCom') | Out-Null
    }
    if ($PSBoundParameters.ContainsKey('BatchSize')) {
        $argsList.Add('-BatchSize') | Out-Null
        $argsList.Add([string]$BatchSize) | Out-Null
    }

    Write-Host 'Relaunching elevated...' -ForegroundColor Yellow
    Start-Process -FilePath 'powershell.exe' -ArgumentList ($argsList -join ' ') -Verb RunAs | Out-Null
    return
}

function Get-UniqueTargetPath {
    param(
        [Parameter(Mandatory)]
        [string]$Directory,

        [Parameter(Mandatory)]
        [string]$BaseName,

        [Parameter(Mandatory)]
        [string]$Extension,

        [Parameter(Mandatory)]
        [string]$CurrentFullName,

        [AllowNull()]
        [System.Collections.Generic.HashSet[string]]$ReservedPaths
    )

    if ($null -eq $ReservedPaths) {
        $ReservedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    # Append _renamed suffix so files are identifiable even without the log
    if ($BaseName -notmatch '_renamed$') {
        $BaseName = $BaseName + '_renamed'
    }

    $candidate = Join-Path $Directory ($BaseName + $Extension)

    if (($candidate -ieq $CurrentFullName) -and -not $ReservedPaths.Contains($candidate)) {
        $null = $ReservedPaths.Add($candidate)
        return $candidate
    }

    # Use a shared counter dictionary to skip O(n) probing for repeated base names
    if ($null -eq $script:_uniquePathCounters) {
        $script:_uniquePathCounters = [System.Collections.Generic.Dictionary[string,int]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    $counterKey = (Join-Path $Directory ($BaseName + $Extension))
    $counter = 1
    if ($script:_uniquePathCounters.ContainsKey($counterKey)) {
        $counter = $script:_uniquePathCounters[$counterKey]
    }
    if (-not ((Test-Path -LiteralPath $candidate) -or $ReservedPaths.Contains($candidate))) {
        # Base name (no counter) is available
    }
    else {
        $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $BaseName, $counter, $Extension)
        while ((Test-Path -LiteralPath $candidate) -or $ReservedPaths.Contains($candidate)) {
            $counter++
            $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $BaseName, $counter, $Extension)
        }
    }
    $script:_uniquePathCounters[$counterKey] = $counter + 1

    $null = $ReservedPaths.Add($candidate)
    return $candidate
}

function Get-DocxTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        try {
            $entry = $zip.GetEntry('docProps/core.xml')
            if (-not $entry) {
                return $null
            }

            $stream = $entry.Open()
            try {
                $reader = [System.IO.StreamReader]::new($stream)
                $xmlText = $reader.ReadToEnd()
            }
            finally {
                if ($reader) { $reader.Dispose() }
                if ($stream) { $stream.Dispose() }
            }

            if ([string]::IsNullOrWhiteSpace($xmlText)) {
                return $null
            }

            [xml]$xml = $xmlText
            $node = $xml.SelectSingleNode('/*[local-name()="coreProperties"]/*[local-name()="title"]')

            if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                return $node.InnerText.Trim()
            }

            return $null
        }
        finally {
            if ($zip) { $zip.Dispose() }
        }
    }
    catch {
        return $null
    }
}

function Get-DocxHeadingTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        try {
            $entry = $zip.GetEntry('docProps/app.xml')
            if (-not $entry) { return $null }

            $stream = $entry.Open()
            try {
                $reader = [System.IO.StreamReader]::new($stream)
                $xmlText = $reader.ReadToEnd()
            }
            finally {
                if ($reader) { $reader.Dispose() }
                if ($stream) { $stream.Dispose() }
            }

            if ([string]::IsNullOrWhiteSpace($xmlText)) { return $null }

            [xml]$xml = $xmlText
            $titles = $xml.SelectNodes('//*[local-name()="TitlesOfParts"]//*[local-name()="lpstr"]')
            if ($titles) {
                foreach ($titleNode in $titles) {
                    if ($titleNode -and -not [string]::IsNullOrWhiteSpace($titleNode.InnerText)) {
                        return $titleNode.InnerText.Trim()
                    }
                }
            }

            return $null
        }
        finally {
            if ($zip) { $zip.Dispose() }
        }
    }
    catch {
        return $null
    }
}

function Get-RtfTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        $fileInfo = [System.IO.FileInfo]::new($FilePath)
        if ($fileInfo.Length -eq 0) { return $null }
        if ($fileInfo.Length -gt $script:MaxFileReadBytes) { return $null }
        $rtf = Get-Content -LiteralPath $FilePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($rtf)) { return $null }

        $match = [regex]::Match($rtf, '\\\\title\s+(?<title>[^\\{}]+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $match.Success) { return $null }

        $title = $match.Groups['title'].Value
        if ([string]::IsNullOrWhiteSpace($title)) { return $null }

        $title = $title -replace '\\\\par\b', ' '
        $title = $title -replace '\\\\[a-z]+\d*\s?', ' '
        $title = $title -replace '[{}]', ' '
        $title = $title -replace '\s+', ' '
        $title = $title.Trim()

        if ([string]::IsNullOrWhiteSpace($title)) { return $null }
        return $title
    }
    catch {
        return $null
    }
}

function Get-WordComTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    if ($script:NoWordComMode) {
        return $null
    }

    $timeoutMs = $script:PerFileTimeoutSeconds * 1000
    if ($timeoutMs -le 0) { $timeoutMs = 20000 }

    $runspace = $null
    $ps = $null

    try {
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $runspace.Open()

        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        $null = $ps.AddScript({
            param($path)
            $wordApp = $null
            $document = $null
            try {
                $wordApp = New-Object -ComObject Word.Application
                $wordApp.Visible = $false
                $wordApp.DisplayAlerts = 0
                $document = $wordApp.Documents.Open($path, $false, $true, $false)
                if (-not $document) { return $null }
                $title = $null
                try { $title = [string]$document.BuiltInDocumentProperties('Title').Value } catch {}
                if (-not [string]::IsNullOrWhiteSpace($title)) { return $title.Trim() }
                return $null
            }
            catch { return $null }
            finally {
                if ($document) {
                    try { $document.Close($false) | Out-Null } catch {}
                    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) } catch {}
                }
                if ($wordApp) {
                    try { $wordApp.Quit() | Out-Null } catch {}
                    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) } catch {}
                }
            }
        }).AddArgument($FilePath)

        $handle = $ps.BeginInvoke()

        if ($handle.AsyncWaitHandle.WaitOne($timeoutMs)) {
            try {
                $result = $ps.EndInvoke($handle)
                return ($result | Select-Object -First 1)
            }
            catch { return $null }
        }

        # Timed out - kill the hung runspace and any orphan WINWORD processes
        try { $ps.Stop() } catch {}
        try {
            Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        } catch {}
        return $null
    }
    catch {
        return $null
    }
    finally {
        if ($ps) { try { $ps.Dispose() } catch {} }
        if ($runspace) {
            try { $runspace.Close() } catch {}
            try { $runspace.Dispose() } catch {}
        }
    }
}

function Get-EmbeddedXmpTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        $fileInfo = [System.IO.FileInfo]::new($FilePath)
        if ($fileInfo.Length -eq 0) { return $null }
        $readSize = [int][Math]::Min($fileInfo.Length, $script:MaxFileReadBytes)
        $stream = [System.IO.File]::OpenRead($FilePath)
        try {
            $bytes = New-Object byte[] $readSize
            $null = $stream.Read($bytes, 0, $readSize)
        }
        finally {
            $stream.Dispose()
        }

        $text = [System.Text.Encoding]::UTF8.GetString($bytes)
        $bytes = $null
        $startToken = '<x:xmpmeta'
        $endToken = '</x:xmpmeta>'

        $start = $text.IndexOf($startToken, [System.StringComparison]::OrdinalIgnoreCase)
        if ($start -lt 0) { return $null }

        $end = $text.IndexOf($endToken, $start, [System.StringComparison]::OrdinalIgnoreCase)
        if ($end -lt 0) { return $null }

        $xmpXml = $text.Substring($start, ($end - $start) + $endToken.Length)
        $text = $null
        [xml]$xmp = $xmpXml

        $titleNode = $xmp.SelectSingleNode('//*[local-name()="title"]//*[local-name()="li"]')
        if ($titleNode -and -not [string]::IsNullOrWhiteSpace($titleNode.InnerText)) {
            return $titleNode.InnerText.Trim()
        }

        $fallbackNode = $xmp.SelectSingleNode('//*[local-name()="title" or local-name()="Title"]')
        if ($fallbackNode -and -not [string]::IsNullOrWhiteSpace($fallbackNode.InnerText)) {
            return $fallbackNode.InnerText.Trim()
        }
    }
    catch {
        return $null
    }

    return $null
}

function ConvertFrom-PdfLiteralString {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $result = $Value -replace '\\\\', '\\'
    $result = $result -replace '\\n', ' '
    $result = $result -replace '\\r', ' '
    $result = $result -replace '\\t', ' '
    $result = $result -replace '\\\(', '('
    $result = $result -replace '\\\)', ')'
    return ($result -replace '\s+', ' ').Trim()
}

function Get-PdfXmpTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    return Get-EmbeddedXmpTitle -FilePath $FilePath
}

function Get-PdfBookmarkTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        $fileInfo = [System.IO.FileInfo]::new($FilePath)
        if ($fileInfo.Length -eq 0) { return $null }
        $readSize = [int][Math]::Min($fileInfo.Length, $script:MaxFileReadBytes)
        $stream = [System.IO.File]::OpenRead($FilePath)
        try {
            $bytes = New-Object byte[] $readSize
            $null = $stream.Read($bytes, 0, $readSize)
        }
        finally {
            $stream.Dispose()
        }

        $pdfText = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
        $bytes = $null
        $titleMatches = [regex]::Matches($pdfText, '/Title\s*\((?<title>(?:\\.|[^\\)])+)\)')

        foreach ($match in $titleMatches) {
            $rawTitle = $match.Groups['title'].Value
            if (-not [string]::IsNullOrWhiteSpace($rawTitle)) {
                $decoded = ConvertFrom-PdfLiteralString -Value $rawTitle
                if (-not [string]::IsNullOrWhiteSpace($decoded)) {
                    return $decoded
                }
            }
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-PdfInfoTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        $fileInfo = [System.IO.FileInfo]::new($FilePath)
        if ($fileInfo.Length -eq 0) { return $null }
        # Read only the last 4KB + first 8KB where /Info dict usually lives
        $tailSize = [int][Math]::Min($fileInfo.Length, 4096)
        $headSize = [int][Math]::Min($fileInfo.Length, 8192)
        $stream = [System.IO.File]::OpenRead($FilePath)
        try {
            $headBytes = New-Object byte[] $headSize
            $null = $stream.Read($headBytes, 0, $headSize)
            $tailBytes = New-Object byte[] $tailSize
            $null = $stream.Seek(-$tailSize, [System.IO.SeekOrigin]::End)
            $null = $stream.Read($tailBytes, 0, $tailSize)
        }
        finally {
            $stream.Dispose()
        }

        $combined = [System.Text.Encoding]::GetEncoding(28591).GetString($headBytes) + [System.Text.Encoding]::GetEncoding(28591).GetString($tailBytes)
        $headBytes = $null
        $tailBytes = $null

        # Match /Title inside /Info dictionary (literal string or hex string)
        $infoMatch = [regex]::Match($combined, '/Title\s*\((?<title>(?:\\.|[^\\)])+)\)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($infoMatch.Success) {
            $rawTitle = $infoMatch.Groups['title'].Value
            if (-not [string]::IsNullOrWhiteSpace($rawTitle)) {
                $decoded = ConvertFrom-PdfLiteralString -Value $rawTitle
                if (-not [string]::IsNullOrWhiteSpace($decoded)) {
                    return $decoded
                }
            }
        }

        # Try hex string variant: /Title <hex>
        $hexMatch = [regex]::Match($combined, '/Title\s*<(?<hex>[0-9A-Fa-f]+)>')
        if ($hexMatch.Success) {
            $hexStr = $hexMatch.Groups['hex'].Value
            if ($hexStr.Length -ge 4 -and ($hexStr.Length % 2) -eq 0) {
                $decodedBytes = New-Object byte[] ($hexStr.Length / 2)
                for ($hi = 0; $hi -lt $hexStr.Length; $hi += 2) {
                    $decodedBytes[$hi / 2] = [Convert]::ToByte($hexStr.Substring($hi, 2), 16)
                }
                # Detect UTF-16 BOM (FEFF)
                $titleText = if ($decodedBytes.Length -ge 2 -and $decodedBytes[0] -eq 0xFE -and $decodedBytes[1] -eq 0xFF) {
                    [System.Text.Encoding]::BigEndianUnicode.GetString($decodedBytes, 2, $decodedBytes.Length - 2)
                } else {
                    [System.Text.Encoding]::UTF8.GetString($decodedBytes)
                }
                $titleText = ($titleText -replace '\s+', ' ').Trim()
                if (-not [string]::IsNullOrWhiteSpace($titleText)) {
                    return $titleText
                }
            }
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-XmlFileTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        [xml]$xml = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8
    }
    catch {
        try {
            [xml]$xml = Get-Content -LiteralPath $FilePath -Raw
        }
        catch {
            return $null
        }
    }

    if (-not $xml) {
        return $null
    }

    $titleNode = $xml.SelectSingleNode('//*[local-name()="title" or local-name()="Title"]')
    if ($titleNode -and -not [string]::IsNullOrWhiteSpace($titleNode.InnerText)) {
        return $titleNode.InnerText.Trim()
    }

    $dcNode = $xml.SelectSingleNode('//*[local-name()="title"]/*[local-name()="li" or local-name()="Alt"]')
    if ($dcNode -and -not [string]::IsNullOrWhiteSpace($dcNode.InnerText)) {
        return $dcNode.InnerText.Trim()
    }

    return $null
}

function Get-JpegXmpTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    return Get-EmbeddedXmpTitle -FilePath $FilePath
}

function Get-VideoMetadataName {
    param(
        $ShellFolder,

        $ShellItem
    )

    $mediaTitle = Get-ShellMetadataValue -ShellFolder $ShellFolder -ShellItem $ShellItem -PropertyNames @(
        'System.Media.Title',
        'System.Title',
        'System.Media.SubTitle'
    )

    $season = Get-ShellMetadataValue -ShellFolder $ShellFolder -ShellItem $ShellItem -PropertyNames @(
        'System.Media.SeasonNumber'
    )

    $episode = Get-ShellMetadataValue -ShellFolder $ShellFolder -ShellItem $ShellItem -PropertyNames @(
        'System.Media.EpisodeNumber'
    )

    if ($mediaTitle -and $season -and $episode) {
        return ("{0} S{1}E{2}" -f $mediaTitle, $season, $episode)
    }
    if ($mediaTitle -and $episode) {
        return ("{0} Episode {1}" -f $mediaTitle, $episode)
    }
    if ($mediaTitle) {
        return $mediaTitle
    }

    return $null
}

function Get-GpsLocationLabel {
    param(
        $ShellFolder,

        $ShellItem
    )

    $placeName = Get-ShellMetadataValue -ShellFolder $ShellFolder -ShellItem $ShellItem -PropertyNames @(
        'System.Location.PlaceName',
        'System.GPS.AreaInformation'
    )

    if ($placeName) {
        return "Location $placeName"
    }

    $latitude = Get-ShellMetadataValue -ShellFolder $ShellFolder -ShellItem $ShellItem -PropertyNames @(
        'System.GPS.Latitude',
        'System.GPS.DestinationLatitude'
    )

    $longitude = Get-ShellMetadataValue -ShellFolder $ShellFolder -ShellItem $ShellItem -PropertyNames @(
        'System.GPS.Longitude',
        'System.GPS.DestinationLongitude'
    )

    if ($latitude -and $longitude) {
        return ("Location {0}, {1}" -f $latitude, $longitude)
    }

    return $null
}

# All Shell property names that Get-RenameBaseNameFromMetadata and its helpers may query.
# Pre-fetching them in one batch inside a timeout-protected runspace prevents COM deadlocks.
$script:_allShellPropertyNames = @(
    'System.Title', 'System.Document.Title', 'System.Music.Title',
    'System.Media.SubTitle', 'System.Subject', 'System.Keywords', 'System.Comment',
    'System.Music.Artist', 'System.Author',
    'System.Music.AlbumTitle', 'System.Music.TrackNumber',
    'System.Media.Title', 'System.Media.SeasonNumber', 'System.Media.EpisodeNumber',
    'System.Photo.DateTaken', 'System.Media.DateEncoded', 'System.Media.DateReleased',
    'System.RecordedTV.OriginalBroadcastDate', 'System.ItemDate', 'System.DateCreated',
    'System.Location.PlaceName', 'System.GPS.AreaInformation',
    'System.GPS.Latitude', 'System.GPS.DestinationLatitude',
    'System.GPS.Longitude', 'System.GPS.DestinationLongitude',
    'System.Company', 'System.ApplicationName'
)

$script:_shellMetadataRunspace = $null
$script:_shellPropertyCache = $null

function Initialize-ShellMetadataRunspace {
    if ($script:_shellMetadataRunspace) { return }
    $script:_shellMetadataRunspace = [runspacefactory]::CreateRunspace()
    $script:_shellMetadataRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $script:_shellMetadataRunspace.Open()
}

function Reset-ShellMetadataRunspace {
    if ($script:_shellMetadataRunspace) {
        try { $script:_shellMetadataRunspace.Close() } catch {}
        try { $script:_shellMetadataRunspace.Dispose() } catch {}
    }
    $script:_shellMetadataRunspace = [runspacefactory]::CreateRunspace()
    $script:_shellMetadataRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $script:_shellMetadataRunspace.Open()
}

function Get-AllShellFileProperties {
    param(
        [Parameter(Mandatory)]
        [string]$DirectoryPath,

        [Parameter(Mandatory)]
        [string]$FileName,

        [int]$TimeoutMs = 15000
    )

    Initialize-ShellMetadataRunspace

    $ps = [powershell]::Create()
    $ps.Runspace = $script:_shellMetadataRunspace

    $null = $ps.AddScript({
        param($dirPath, $fileName, $propNames)
        if (-not $global:_shellCom) {
            $global:_shellCom = New-Object -ComObject Shell.Application
        }
        if ($global:_cachedMetaDir -ne $dirPath -or -not $global:_cachedMetaFolder) {
            if ($global:_cachedMetaFolder) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($global:_cachedMetaFolder) } catch {}
            }
            $global:_cachedMetaFolder = $global:_shellCom.Namespace($dirPath)
            $global:_cachedMetaDir = $dirPath
        }
        $folder = $global:_cachedMetaFolder
        if (-not $folder) { return @{} }
        $item = $folder.ParseName($fileName)
        if (-not $item) { return @{} }
        $result = @{}
        foreach ($prop in $propNames) {
            try {
                $val = $item.ExtendedProperty($prop)
                if ($null -ne $val) {
                    $text = [string]$val
                    if (-not [string]::IsNullOrWhiteSpace($text)) {
                        $result[$prop] = $text.Trim()
                    }
                }
            } catch {}
        }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) } catch {}
        return $result
    }).AddArgument($DirectoryPath).AddArgument($FileName).AddArgument($script:_allShellPropertyNames)

    try {
        $handle = $ps.BeginInvoke()

        if ($handle.AsyncWaitHandle.WaitOne($TimeoutMs)) {
            try {
                $result = $ps.EndInvoke($handle)
                return ($result | Select-Object -First 1)
            }
            catch { return @{} }
        }

        # Timed out - the Shell COM call is stuck; kill and recreate the runspace
        try { $ps.Stop() } catch {}
        Reset-ShellMetadataRunspace
        return $null
    }
    finally {
        try { $ps.Dispose() } catch {}
    }
}

function Get-ShellMetadataValue {
    param(
        $ShellFolder,

        $ShellItem,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames
    )

    # Use pre-fetched property cache when available (avoids blocking COM calls)
    if ($script:_shellPropertyCache) {
        foreach ($prop in $PropertyNames) {
            if ($script:_shellPropertyCache.ContainsKey($prop)) {
                return $script:_shellPropertyCache[$prop]
            }
        }
        return $null
    }

    foreach ($prop in $PropertyNames) {
        try {
            $value = $ShellItem.ExtendedProperty($prop)
            if ($null -ne $value) {
                $text = [string]$value
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    return $text.Trim()
                }
            }
        }
        catch {
            continue
        }
    }

    return $null
}

function Get-DateLabel {
    param(
        $ShellFolder,

        $ShellItem,

        [string]$Prefix
    )

    $dateRaw = Get-ShellMetadataValue -ShellFolder $ShellFolder -ShellItem $ShellItem -PropertyNames @(
        'System.Photo.DateTaken',
        'System.Media.DateEncoded',
        'System.Media.DateReleased',
        'System.RecordedTV.OriginalBroadcastDate',
        'System.ItemDate',
        'System.DateCreated'
    )

    if (-not $dateRaw) {
        return $null
    }

    try {
        $dateValue = [DateTime]::Parse($dateRaw)
        if ([string]::IsNullOrWhiteSpace($Prefix)) {
            return ('{0:yyyy-MM-dd}' -f $dateValue)
        }

        return ('{0} {1:yyyy-MM-dd HH-mm-ss}' -f $Prefix, $dateValue)
    }
    catch {
        $normalized = ($dateRaw -replace '[:/\\]+', '-' -replace '\s+', ' ').Trim()
        if ($normalized) {
            if ([string]::IsNullOrWhiteSpace($Prefix)) {
                return $normalized
            }
            return "$Prefix $normalized"
        }
    }

    return $null
}

function Add-ContextualEnhancers {
    param(
        [Parameter(Mandatory)]
        [psobject]$BaseResult,

        $ShellFolder,

        $ShellItem,

        [Parameter(Mandatory)]
        [string]$Extension
    )

    $dateSuffix = Get-DateLabel -ShellFolder $ShellFolder -ShellItem $ShellItem
    $locationLabel = Get-GpsLocationLabel -ShellFolder $ShellFolder -ShellItem $ShellItem
    $locationSuffix = if ($locationLabel) { ($locationLabel -replace '^Location\s+', '').Trim() } else { $null }

    if ([string]::IsNullOrWhiteSpace($dateSuffix) -and [string]::IsNullOrWhiteSpace($locationSuffix)) {
        return $BaseResult
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add($BaseResult.Name) | Out-Null

    if ($dateSuffix -and ($BaseResult.Name -notmatch [regex]::Escape($dateSuffix))) {
        $parts.Add($dateSuffix) | Out-Null
    }

    if ($locationSuffix -and ($BaseResult.Name -notmatch [regex]::Escape($locationSuffix))) {
        $parts.Add($locationSuffix) | Out-Null
    }

    if ($parts.Count -le 1) {
        return $BaseResult
    }

    $enhancedName = $parts -join ' - '
    $enhancedReason = if ([string]::IsNullOrWhiteSpace($BaseResult.ScoreReason)) {
        'Contextual enhancers appended'
    }
    else {
        '{0}; contextual enhancers appended (date/location)' -f $BaseResult.ScoreReason
    }

    $enhanced = New-MetadataNameResult -Name $enhancedName -Source $BaseResult.Source -ConfidenceScore $BaseResult.ConfidenceScore -ScoreReason $enhancedReason
    if ($enhanced) {
        return $enhanced
    }

    return $BaseResult
}

function Get-MediaPrefix {
    param(
        [Parameter(Mandatory)]
        [string]$Extension
    )

    $imageExt = @('.jpg', '.jpeg', '.png', '.heic', '.tif', '.tiff', '.gif', '.bmp', '.webp', '.raw', '.cr2', '.nef', '.arw', '.dng', '.ico', '.jfif', '.avif')
    $videoExt = @('.mp4', '.mov', '.m4v', '.avi', '.mkv', '.wmv', '.flv', '.webm', '.mpg', '.mpeg', '.3gp', '.mts', '.m2ts', '.ts')
    $audioExt = @('.mp3', '.wav', '.flac', '.aac', '.m4a', '.ogg', '.wma', '.aiff', '.alac', '.opus')
    $docExt = @('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.pdf', '.ppt', '.pptx', '.xls', '.xlsx', '.txt', '.rtf', '.odt', '.ods', '.odp')

    if ($Extension -in $imageExt) { return 'Photo' }
    if ($Extension -in $videoExt) { return 'Video' }
    if ($Extension -in $audioExt) { return 'Audio' }
    if ($Extension -in $docExt) { return 'Document' }
    return 'File'
}

function Get-RenameBaseNameFromMetadata {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory)]
        [System.__ComObject]$Shell,

        [switch]$Force,

        [switch]$OnlyGeneric
    )

    if ($OnlyGeneric.IsPresent -and -not $Force.IsPresent) {
        if ($File.BaseName -notmatch '^(file\d+|img_?\d+|dsc_?\d+|p\d+|scan\d+)$') {
            return $null
        }
    }

    $extension = $File.Extension.ToLowerInvariant()
    $wordOpenXmlExtensions = @('.docx', '.docm', '.dotx', '.dotm')
    $officeOpenXmlExtensions = @('.pptx', '.pptm', '.xlsx', '.xlsm')
    $wordComExtensions = @('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.rtf')

    $primaryResult = $null

    if ($extension -in $wordOpenXmlExtensions) {
        $docxTitle = Get-DocxTitle -FilePath $File.FullName
        if ($docxTitle) {
            $sourceName = if ($extension -eq '.docx') { 'DOCX core title' } else { 'Word OpenXML core title' }
            $result = New-MetadataNameResult -Name $docxTitle -Source $sourceName
            if ($result) { $primaryResult = $result }
        }

        if (-not $primaryResult) {
            $docxHeading = Get-DocxHeadingTitle -FilePath $File.FullName
            if ($docxHeading) {
                $sourceName = if ($extension -eq '.docx') { 'DOCX heading title' } else { 'Word heading title' }
                $result = New-MetadataNameResult -Name $docxHeading -Source $sourceName
                if ($result) { $primaryResult = $result }
            }
        }
    }

    # PPTX/XLSX/PPTM/XLSM use the same Office OpenXML ZIP structure
    if (-not $primaryResult -and $extension -in $officeOpenXmlExtensions) {
        $officeTitle = Get-DocxTitle -FilePath $File.FullName
        if ($officeTitle) {
            $sourceName = if ($extension -in @('.pptx', '.pptm')) { 'PPTX core title' } else { 'XLSX core title' }
            $result = New-MetadataNameResult -Name $officeTitle -Source $sourceName
            if ($result) { $primaryResult = $result }
        }

        if (-not $primaryResult) {
            $officeHeading = Get-DocxHeadingTitle -FilePath $File.FullName
            if ($officeHeading) {
                $sourceName = if ($extension -in @('.pptx', '.pptm')) { 'PPTX heading title' } else { 'XLSX heading title' }
                $result = New-MetadataNameResult -Name $officeHeading -Source $sourceName
                if ($result) { $primaryResult = $result }
            }
        }
    }

    if (-not $primaryResult -and $extension -eq '.rtf') {
        $rtfTitle = Get-RtfTitle -FilePath $File.FullName
        if ($rtfTitle) {
            $result = New-MetadataNameResult -Name $rtfTitle -Source 'RTF title control word'
            if ($result) { $primaryResult = $result }
        }
    }

    if (-not $primaryResult -and -not $script:NoWordComMode -and $extension -in $wordComExtensions) {
        $wordComTitle = Get-WordComTitle -FilePath $File.FullName
        if ($wordComTitle) {
            $result = New-MetadataNameResult -Name $wordComTitle -Source 'Word COM title'
            if ($result) { $primaryResult = $result }
        }
    }

    if (-not $primaryResult -and $extension -eq '.pdf') {
        $pdfXmpTitle = Get-PdfXmpTitle -FilePath $File.FullName
        if ($pdfXmpTitle) {
            $result = New-MetadataNameResult -Name $pdfXmpTitle -Source 'PDF XMP title'
            if ($result) { $primaryResult = $result }
        }

        if (-not $primaryResult) {
            $pdfInfoTitle = Get-PdfInfoTitle -FilePath $File.FullName
            if ($pdfInfoTitle) {
                $result = New-MetadataNameResult -Name $pdfInfoTitle -Source 'PDF Info title'
                if ($result) { $primaryResult = $result }
            }
        }

        if (-not $primaryResult) {
            $pdfBookmarkTitle = Get-PdfBookmarkTitle -FilePath $File.FullName
            if ($pdfBookmarkTitle) {
                $result = New-MetadataNameResult -Name $pdfBookmarkTitle -Source 'PDF bookmark title'
                if ($result) { $primaryResult = $result }
            }
        }
    }

    if (-not $primaryResult -and $extension -in @('.jpg', '.jpeg', '.jfif', '.png')) {
        $xmpTitle = Get-JpegXmpTitle -FilePath $File.FullName
        if ($xmpTitle) {
            $result = New-MetadataNameResult -Name $xmpTitle -Source 'Image embedded XMP title'
            if ($result) { $primaryResult = $result }
        }
    }

    if (-not $primaryResult -and $extension -eq '.xml') {
        $xmlTitle = Get-XmlFileTitle -FilePath $File.FullName
        if ($xmlTitle) {
            $result = New-MetadataNameResult -Name $xmlTitle -Source 'XML title node'
            if ($result) { $primaryResult = $result }
        }
    }

    # Pre-fetch all Shell metadata in a timeout-protected runspace to prevent COM deadlocks
    $shellTimeoutMs = $script:PerFileTimeoutSeconds * 1000
    if ($shellTimeoutMs -le 0) { $shellTimeoutMs = 20000 }
    $script:_shellPropertyCache = Get-AllShellFileProperties -DirectoryPath $File.DirectoryName -FileName $File.Name -TimeoutMs $shellTimeoutMs
    if ($null -eq $script:_shellPropertyCache) {
        Write-RunLog -Message ("ShellTimeout Shell metadata extraction timed out for {0}" -f $File.FullName)
        return $null
    }

    $shellFolder = $null
    $shellItem = $null

    try {

    $mediaPrefix = Get-MediaPrefix -Extension $extension

    if (-not $primaryResult) {
        $title = Get-ShellMetadataValue -ShellFolder $shellFolder -ShellItem $shellItem -PropertyNames @(
            'System.Title',
            'System.Document.Title',
            'System.Music.Title',
            'System.Media.SubTitle',
            'System.Subject',
            'System.Keywords',
            'System.Comment'
        )

        if ($title) {
            $result = New-MetadataNameResult -Name $title -Source 'Windows metadata title'
            if ($result) { $primaryResult = $result }
        }
    }

    if (-not $primaryResult -and $mediaPrefix -eq 'Audio') {
        $artist = Get-ShellMetadataValue -ShellFolder $shellFolder -ShellItem $shellItem -PropertyNames @(
            'System.Music.Artist',
            'System.Author'
        )
        $trackTitle = Get-ShellMetadataValue -ShellFolder $shellFolder -ShellItem $shellItem -PropertyNames @(
            'System.Music.Title',
            'System.Title'
        )
        $album = Get-ShellMetadataValue -ShellFolder $shellFolder -ShellItem $shellItem -PropertyNames @(
            'System.Music.AlbumTitle'
        )
        $trackNumber = Get-ShellMetadataValue -ShellFolder $shellFolder -ShellItem $shellItem -PropertyNames @(
            'System.Music.TrackNumber'
        )

        if ($artist -and $trackTitle) {
            $audioName = if ($trackNumber -and $album) {
                "{0} - {1} - {2} ({3})" -f $trackNumber, $artist, $trackTitle, $album
            }
            elseif ($trackNumber) {
                "{0} - {1} - {2}" -f $trackNumber, $artist, $trackTitle
            }
            elseif ($album) {
                "{0} - {1} ({2})" -f $artist, $trackTitle, $album
            }
            else {
                "$artist - $trackTitle"
            }

            $result = New-MetadataNameResult -Name $audioName -Source 'Audio composite metadata'
            if ($result) { $primaryResult = $result }
        }

        if (-not $primaryResult -and $trackTitle) {
            $result = New-MetadataNameResult -Name $trackTitle -Source 'Audio title metadata'
            if ($result) { $primaryResult = $result }
        }
    }

    if (-not $primaryResult -and $mediaPrefix -eq 'Video') {
        $videoName = Get-VideoMetadataName -ShellFolder $shellFolder -ShellItem $shellItem
        if ($videoName) {
            $result = New-MetadataNameResult -Name $videoName -Source 'Video episode metadata'
            if ($result) { $primaryResult = $result }
        }
    }

    if ($primaryResult) {
        return Add-ContextualEnhancers -BaseResult $primaryResult -ShellFolder $shellFolder -ShellItem $shellItem -Extension $extension
    }

    $gpsName = Get-GpsLocationLabel -ShellFolder $shellFolder -ShellItem $shellItem
    $dateLabel = Get-DateLabel -ShellFolder $shellFolder -ShellItem $shellItem -Prefix $mediaPrefix

    if ($dateLabel -and $gpsName) {
        $gpsSuffix = ($gpsName -replace '^Location\s+', '').Trim()
        $compositeContext = if ($gpsSuffix) { "{0} - {1}" -f $dateLabel, $gpsSuffix } else { $dateLabel }
        $result = New-MetadataNameResult -Name $compositeContext -Source 'Media date metadata' -ScoreReason 'Date-based media naming with GPS context'
        if ($result) { return $result }
    }

    if ($gpsName) {
        $result = New-MetadataNameResult -Name $gpsName -Source 'GPS location metadata'
        if ($result) { return $result }
    }

    if ($dateLabel) {
        $result = New-MetadataNameResult -Name $dateLabel -Source 'Media date metadata'
        if ($result) { return $result }
    }

    $authorLike = Get-ShellMetadataValue -ShellFolder $shellFolder -ShellItem $shellItem -PropertyNames @(
        'System.Author',
        'System.Company',
        'System.ApplicationName'
    )

    if ($authorLike) {
        $result = New-MetadataNameResult -Name "$mediaPrefix $authorLike" -Source 'Author/company metadata'
        if ($result) { return $result }
    }

    return $null

    } # end try
    finally {
        # Clear per-file property cache; COM objects live in the metadata runspace
        $script:_shellPropertyCache = $null
    }
}

$retryErrorActive = $false
$retryErrorSourceLog = $null
$retryErrorImportedFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$retryErrorsRequested = $RetryErrors.IsPresent -or -not [string]::IsNullOrWhiteSpace($RetryErrorsLogPath)
if ($retryErrorsRequested) {
    $resolvedRetryErrorLog = Resolve-ResumeLogPath -Root $root -RequestedPath $RetryErrorsLogPath
    if (-not [string]::IsNullOrWhiteSpace($resolvedRetryErrorLog)) {
        $errorState = Import-ErrorStateFromLog -Path $resolvedRetryErrorLog
        if ($errorState.ImportedCount -gt 0) {
            $retryErrorActive = $true
            $retryErrorSourceLog = $resolvedRetryErrorLog
            foreach ($importedFile in $errorState.ImportedFiles) {
                $null = $retryErrorImportedFiles.Add($importedFile)
            }
        }
    }

    if (-not $retryErrorActive) {
        if (-not [string]::IsNullOrWhiteSpace($RetryErrorsLogPath) -and -not (Test-Path -LiteralPath $RetryErrorsLogPath)) {
            Write-Host ("RetryErrors requested, but log was not found: {0}" -f $RetryErrorsLogPath) -ForegroundColor Yellow
        }
        else {
            Write-Host 'RetryErrors requested, but no parseable ERROR entries were found in the selected log.' -ForegroundColor Yellow
        }
    }
}

if ($resumeRequested) {
    $resolvedResumeLog = Resolve-ResumeLogPath -Root $root -RequestedPath $ResumeLogPath
    if (-not [string]::IsNullOrWhiteSpace($resolvedResumeLog)) {
        $importState = Import-ResumeStateFromLog -Path $resolvedResumeLog
        $script:ResumeActive = (($importState.ImportedFrom -gt 0) -or ($importState.ImportedTo -gt 0))
        $script:ResumeLogSource = $resolvedResumeLog
    }
}

if (-not $ReportPath) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $ReportPath = Join-Path $root ("Rename-Report-{0}.txt" -f $stamp)
}

if (-not $CsvPath) {
    $CsvPath = [System.IO.Path]::ChangeExtension($ReportPath, '.csv')
}

$shouldRecurse = -not $NoRecurse
$runMode = if ($isApplyMode) { 'Apply (rename)' } else { 'Preview only' }
if ($isImmediateRenameMode) {
    $runMode = 'Apply (rename immediate)'
}
if ($InteractiveSubfolders) {
    $runMode = $runMode + ' + InteractiveSubfolders'
}
if ($QuickPreview) {
    $runMode = if ($isApplyMode) {
        "Apply (rename) + QuickPreview($QuickPreviewLimit)"
    }
    else {
        "Preview only + QuickPreview($QuickPreviewLimit)"
    }

    if ($isImmediateRenameMode) {
        $runMode = "Apply (rename immediate) + QuickPreview($QuickPreviewLimit)"
    }

    if ($InteractiveSubfolders) {
        $runMode = $runMode + ' + InteractiveSubfolders'
    }
}

$startupSettings = [ordered]@{
    Apply = $isApplyMode
    ImmediateRename = $isImmediateRenameMode
    QuickPreview = [bool]$QuickPreview
    QuickPreviewLimit = if ($QuickPreview) { $QuickPreviewLimit } else { 'N/A' }
    NoRecurse = [bool]$NoRecurse
    InteractiveSubfolders = [bool]$InteractiveSubfolders
    IncludeHidden = [bool]$IncludeHidden
    Force = [bool]$Force
    OnlyGeneric = [bool]$OnlyGeneric
    ResumeRequested = $resumeRequested
    ResumeActive = $script:ResumeActive
    ResumeLogPath = if ($script:ResumeLogSource) { $script:ResumeLogSource } elseif ($ResumeLogPath) { $ResumeLogPath } else { 'N/A' }
    RetryErrorsRequested = $retryErrorsRequested
    RetryErrorsActive = $retryErrorActive
    RetryErrorsLogPath = if ($retryErrorSourceLog) { $retryErrorSourceLog } elseif ($RetryErrorsLogPath) { $RetryErrorsLogPath } else { 'N/A' }
    IntelligentScan = $intelligentScanEnabled
    UseExtensionFolders = $useExtensionFolders
    IncludeOtherFiles = $includeOtherFilesNow
    ScanOtherFilesAtEnd = $scanOtherFilesAtEndEnabled
    ListUncommonFiles = [bool]$ListUncommonFiles
    NoAutoResumePrompt = [bool]$NoAutoResumePrompt
    SkipLikelyAlreadyRenamed = [bool]$SkipLikelyAlreadyRenamed
    SkipMetadataUpgrade = [bool]$SkipMetadataUpgrade
    PerFileTimeoutSeconds = $PerFileTimeoutSeconds
    BatchSize = $BatchSize
    Unattended = $script:UnattendedMode
    UseWordCom = -not $script:NoWordComMode
    WordComDisabled = $script:NoWordComMode
    MaxFileReadMB = $MaxFileReadMB
    ComRecycleInterval = $script:ComRecycleInterval
    ReportPath = $ReportPath
    CsvPath = $CsvPath
    RequestedLogPath = if ([string]::IsNullOrWhiteSpace($LogPath)) { 'Auto' } else { $LogPath }
}

Initialize-RunLog -Path $root -Recurse:$shouldRecurse -Mode $runMode -Admin:$isAdmin -RequestedLogPath $LogPath -Settings $startupSettings

Write-Host (
    "Startup: Path='{0}' | Recurse={1} | Mode={2} | Admin={3}" -f `
    $root,
    $shouldRecurse,
    $runMode,
    $isAdmin
) -ForegroundColor Cyan
Write-Host ("Run log: {0}" -f $script:LogPath) -ForegroundColor DarkGray
if ($resumeRequested) {
    if ($script:ResumeActive) {
        Write-Host (
            "Resume mode: loaded {0} completed rename entries from {1}" -f $script:ResumeRenamedTo.Count,
            $script:ResumeLogSource
        ) -ForegroundColor Yellow
        Write-RunLog -Message (
            "Resume enabled. LoadedFrom={0}; LoadedTo={1}; SourceLog={2}" -f
            $script:ResumeRenamedFrom.Count,
            $script:ResumeRenamedTo.Count,
            $script:ResumeLogSource
        )
    }
    else {
        if (-not [string]::IsNullOrWhiteSpace($ResumeLogPath) -and -not (Test-Path -LiteralPath $ResumeLogPath)) {
            Write-Host ("Resume mode requested, but log was not found: {0}" -f $ResumeLogPath) -ForegroundColor Yellow
            Write-RunLog -Message ("Resume requested but log not found: {0}" -f $ResumeLogPath)
        }
        else {
            Write-Host 'Resume mode requested, but no parseable completed rename entries were found.' -ForegroundColor Yellow
            Write-RunLog -Message 'Resume requested but no parseable completed rename entries were found.'
        }
    }
}
if (-not $QuickPreview) {
    Write-Host 'Runtime controls: press P to pause, C to continue (when paused), or S to stop gracefully.' -ForegroundColor DarkYellow
}
Write-RunLog -Message ("Startup complete. Report path: {0}; CSV path: {1}" -f $ReportPath, $CsvPath)


$items = Get-DiscoveredFiles -Root $root -IncludeHidden:([bool]$IncludeHidden) -Recurse:$shouldRecurse -InteractiveSubfolders:([bool]$InteractiveSubfolders)
if ($DebugOutput.IsPresent) {
    Write-Host ("DEBUG: After Get-DiscoveredFiles, items count: {0}" -f (@($items).Count)) -ForegroundColor Cyan
    if (@($items).Count -eq 0) {
        Write-Host "DEBUG: No files discovered. Check scan path, permissions, and filters." -ForegroundColor Red
    } else {
        Write-Host ("DEBUG: First 5 items: {0}" -f ((@($items)[0..([Math]::Min(4,(@($items).Count-1)))] | ForEach-Object { $_.FullName }) -join ', ')) -ForegroundColor Cyan
    }
}

    # If using extension folders, override $items with files from subfolders named as common Office, image, audio, or video extensions
if ($useExtensionFolders) {
    $commonExtFolders = @(
        # Office/document extensions
        'doc','docx','docm','dot','dotx','dotm','pdf','ppt','pptx','xls','xlsx','rtf','odt','ods','odp','xml',
        # Image extensions
        'jpg','jpeg','jfif','png','heic','tif','tiff','gif','bmp','webp','raw','cr2','nef','arw','dng','ico','avif',
        # Audio extensions
        'mp3','wav','flac','aac','m4a','ogg','wma','aiff','alac','opus',
        # Video extensions
        'mp4','mov','m4v','avi','mkv','wmv','flv','webm','mpg','mpeg','3gp','mts','m2ts','ts'
    )
    $extFolders = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $commonExtFolders -contains $_.Name.ToLowerInvariant() }
    $filteredFiles = @()
    foreach ($folder in $extFolders) {
        try {
            $files = Get-ChildItem -LiteralPath $folder.FullName -File -Recurse -Force:([bool]$IncludeHidden) -ErrorAction SilentlyContinue
            $filteredFiles += $files
        } catch {}
    }
    $items = $filteredFiles
    Write-Host ("Extension-folder scan: found {0} files in {1} extension-named folders." -f $items.Count, $extFolders.Count) -ForegroundColor Yellow
    Write-RunLog -Message ("ExtensionFolderScan enabled. Folders={0}; Files={1}" -f $extFolders.Count, $items.Count)
}
$uncommonSkipped = 0

if ($retryErrorActive) {
    $items = @($items | Where-Object { $retryErrorImportedFiles.Contains($_.FullName) })
    Write-Host ("RetryErrors mode: retrying files from log ({0}). Matched files in current scan: {1}" -f $retryErrorSourceLog, @($items).Count) -ForegroundColor Yellow
    Write-RunLog -Message ("RetryErrors mode enabled. SourceLog={0}; MatchedFiles={1}" -f $retryErrorSourceLog, @($items).Count)
}

if ($intelligentScanEnabled -and -not $includeOtherFilesNow -and -not $retryErrorActive) {
    $primaryItems = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    $deferredItems = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

    foreach ($scanItem in @($items)) {
        if (Test-IsCommonUserFileType -File $scanItem) {
            $primaryItems.Add($scanItem) | Out-Null
        }
        else {
            $deferredItems.Add($scanItem) | Out-Null
        }
    }

    if ($scanOtherFilesAtEndEnabled) {
        $items = @($primaryItems) + @($deferredItems)
        Write-Host (
            "Intelligent scan enabled: common files first ({0}), uncommon files deferred to end ({1})." -f
            $primaryItems.Count,
            $deferredItems.Count
        ) -ForegroundColor Yellow
        Write-RunLog -Message (
            "IntelligentScan enabled. CommonFiles={0}; DeferredUncommonFiles={1}; DeferredScanAtEnd=True" -f
            $primaryItems.Count,
            $deferredItems.Count
        )
    }
    else {
        $items = @($primaryItems)
        $uncommonSkipped = $deferredItems.Count
        if ($ListUncommonFiles -and $deferredItems.Count -gt 0) {
            foreach ($uncommonItem in $deferredItems) {
                Write-Host ("[Uncommon] {0}" -f $uncommonItem.FullName) -ForegroundColor DarkYellow
                Write-RunLog -Message ("UNCOMMON_SKIPPED | Path={0}" -f $uncommonItem.FullName)
            }
        }
        Write-Host (
            "Intelligent scan enabled: scanning common files only ({0}). [Uncommon] skipped for now ({1})." -f
            $primaryItems.Count,
            $deferredItems.Count
        ) -ForegroundColor Yellow
        Write-Host 'Tip: use -ScanOtherFilesAtEnd to process uncommon files after common files, or -IncludeOtherFiles to include all files now.' -ForegroundColor DarkYellow
        Write-RunLog -Message (
            "IntelligentScan enabled. CommonFiles={0}; SkippedUncommonFiles={1}; DeferredScanAtEnd=False" -f
            $primaryItems.Count,
            $deferredItems.Count
        )
    }
}

if ($QuickPreview) {
    $items = $items | Select-Object -First $QuickPreviewLimit
    Write-Host ("Quick preview enabled: evaluating first {0} files." -f $QuickPreviewLimit) -ForegroundColor Yellow
    Write-RunLog -Message ("QuickPreview enabled. Limit={0}" -f $QuickPreviewLimit)
}

$discoveredCount = @($items).Count
if (-not $QuickPreview -and $discoveredCount -ge $script:LongRunThreshold) {
    $longScanContext = if ($isApplyMode) {
        if ($isImmediateRenameMode) { 'before immediate rename' } else { 'before rename' }
    }
    else { 'for preview planning' }
    Write-Host (
        "Large scan detected ({0} files, {1}). Runtime controls: P = pause, C = continue (while paused), S = stop gracefully." -f $discoveredCount, $longScanContext
    ) -ForegroundColor Yellow

    $continueLongPreview = if ($isApplyMode) {
        if ($isImmediateRenameMode) {
            Read-YesNo -Prompt 'Proceed with this long run in immediate rename mode?' -DefaultYes $true
        }
        else {
            Read-YesNo -Prompt 'Proceed with this long run in rename mode (planning scan)?' -DefaultYes $true
        }
    }
    else {
        Read-YesNo -Prompt 'Proceed with this long scan for preview planning?' -DefaultYes $true
    }

    if (-not $continueLongPreview) {
        Write-Host 'Run canceled before long scan.' -ForegroundColor Yellow
        Write-RunLog -Message ("User canceled long scan before processing. DiscoveredFiles={0}; ApplyMode={1}" -f $discoveredCount, $isApplyMode)
        Write-EndOfLogCode -Status 'Run canceled by user before long scan'
        return
    }
}
$shell = New-Object -ComObject Shell.Application
Initialize-ShellMetadataRunspace

$reservedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:_uniquePathCounters = $null
$script:_cachedShellFolder = $null
$script:_cachedShellFolderDir = $null
$plan = [System.Collections.Generic.List[object]]::new()

$renamed = 0
$skipped = 0
$resumeSkipped = 0
$heuristicSkipped = 0
$timeoutSkipped = 0
$artifactSkipped = 0
$errorCount = 0
$consecutiveErrors = 0
$heartbeatEvery = 100
$scanProcessed = 0
$renameProcessed = 0
$scanHeartbeatStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$renameHeartbeatStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$processingErrorFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$renameErrorEntries = [System.Collections.Generic.List[object]]::new()

$itemsArray = @($items)
$totalItems = $itemsArray.Count
$safeBatchSize = Convert-LogSettingToInt $BatchSize 100 1000000

if ($totalItems -eq 0) {
    Write-Host "No files discovered. Exiting." -ForegroundColor Red
    Write-RunLog -Message "No files discovered. Exiting."
    Write-EndOfLogCode -Status 'No files discovered.'
    return
}

$totalItemBatches = [int][Math]::Ceiling($totalItems / [double]$safeBatchSize)
Write-Host ("Discovered {0} files. Starting batch processing in {1} batches..." -f $totalItems, $totalItemBatches) -ForegroundColor Green

try { $script:OriginalWindowTitle = $host.UI.RawUI.WindowTitle } catch {}

for ($batchStart = 0; $batchStart -lt $totalItems; $batchStart += $safeBatchSize) {
    $batchEnd = [Math]::Min($batchStart + $safeBatchSize - 1, $totalItems - 1)
    $currentBatchNumber = [int]([Math]::Floor($batchStart / [double]$safeBatchSize) + 1)
    $currentBatchCount = ($batchEnd - $batchStart + 1)

    if ($totalItemBatches -gt 1) {
        Write-Host (
            "Processing batch {0}/{1} ({2} files)..." -f
            $currentBatchNumber,
            $totalItemBatches,
            $currentBatchCount
        ) -ForegroundColor DarkYellow
        Write-RunLog -Message (
            "BatchStart | Phase=scan | Batch={0}/{1} | Size={2}" -f
            $currentBatchNumber,
            $totalItemBatches,
            $currentBatchCount
        )
    }

    for ($i = $batchStart; $i -le $batchEnd; $i++) {
        $file = $itemsArray[$i]
        $fileStopwatch = $null
        try {
            $activePhase = if ($isImmediateRenameMode) { 'rename' } elseif ($isApplyMode) { 'planning' } else { 'preview' }
            if (-not (Test-RunControl -Phase $activePhase)) {
                break
            }

            $scanProcessed++

            # Update window title on every file so taskbar always shows live progress
            Update-WindowTitle -Phase 'Scan' -Processed $scanProcessed -Total $totalItems -Renamed $renamed -Errors $errorCount -ElapsedSeconds $scanHeartbeatStopwatch.Elapsed.TotalSeconds

            if (($scanProcessed % $heartbeatEvery) -eq 0) {
                # Visible console heartbeat so scrolling output proves the script is alive
                $hbElapsed = $scanHeartbeatStopwatch.Elapsed
                $hbPct = if ($totalItems -gt 0) { [int](($scanProcessed / $totalItems) * 100) } else { 0 }
                Write-Host (
                    "  [{0}] Scan progress: {1}/{2} ({3}%) | Candidates:{4} Renamed:{5} Skipped:{6} Errors:{7} | Elapsed:{8:N0}m {9:N0}s" -f
                    (Get-Date -Format 'HH:mm:ss'),
                    $scanProcessed,
                    $totalItems,
                    $hbPct,
                    $plan.Count,
                    $renamed,
                    $skipped,
                    $errorCount,
                    [Math]::Floor($hbElapsed.TotalMinutes),
                    $hbElapsed.Seconds
                ) -ForegroundColor DarkGray

                Write-RunLog -Message (
                    "HEARTBEAT | Phase=scan | Processed={0}/{1} | Candidates={2} | Renamed={3} | Skipped={4} | Errors={5} | ElapsedSeconds={6:N1}" -f
                    $scanProcessed,
                    $totalItems,
                    $plan.Count,
                    $renamed,
                    $skipped,
                    $errorCount,
                    $scanHeartbeatStopwatch.Elapsed.TotalSeconds
                )
            }

            # Periodic COM recycling and GC to prevent handle exhaustion and memory pressure
            if ($script:ComRecycleInterval -gt 0 -and ($scanProcessed % $script:ComRecycleInterval) -eq 0 -and $scanProcessed -gt 0) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
                # Clear the per-directory cache since we're recycling the parent COM object
                if ($null -ne $script:_cachedShellFolder) {
                    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:_cachedShellFolder) } catch {}
                    $script:_cachedShellFolder = $null
                    $script:_cachedShellFolderDir = $null
                }
                Reset-ShellMetadataRunspace
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
                $shell = New-Object -ComObject Shell.Application
                $memMB = [int]([System.Diagnostics.Process]::GetCurrentProcess().WorkingSet64 / 1MB)
                Write-RunLog -Message ("COM_RECYCLE | Recycled Shell.Application at file #{0} | MemoryMB={1}" -f $scanProcessed, $memMB)
                if ($memMB -gt 1500) {
                    Write-RunLog -Message ("MEMORY_WARNING | WorkingSet={0}MB exceeds 1500MB safety threshold. Forcing aggressive GC." -f $memMB)
                    [GC]::Collect(2, [System.GCCollectionMode]::Forced, $true)
                    [GC]::WaitForPendingFinalizers()
                    [GC]::Collect()
                    $memAfter = [int]([System.Diagnostics.Process]::GetCurrentProcess().WorkingSet64 / 1MB)
                    Write-RunLog -Message ("MEMORY_WARNING | Post-aggressive-GC MemoryMB={0}" -f $memAfter)
                    if ($memAfter -gt 1800) {
                        Write-Host ("  [!] MEMORY CRITICAL: {0}MB. Saving progress and exiting to prevent crash." -f $memAfter) -ForegroundColor Red
                        Write-RunLog -Message ("MEMORY_CRITICAL | WorkingSet={0}MB. Graceful exit to prevent OS kill." -f $memAfter)
                        Write-EndOfLogCode -Status 'Graceful exit (memory critical)'
                        Restore-WindowTitle
                        return
                    }
                }
            }

            if ($script:ResumeActive -and $script:ResumeRenamedFrom.Contains($file.FullName)) {
                $resumeSkipped++
                continue
            }

            if ($script:ResumeActive -and $script:ResumeRenamedTo.Contains($file.FullName) -and $SkipMetadataUpgrade) {
                $resumeSkipped++
                continue
            }

            if (Test-IsRunArtifactFile -File $file -CurrentRunLogPath $script:LogPath -CurrentReportPath $ReportPath -CurrentCsvPath $CsvPath -ResumeLogPath $script:ResumeLogSource) {
                $artifactSkipped++
                Write-SkippedLiveLine -FileName $file.Name -Reason 'Script run/report artifact'
                Write-RunLog -Message ("ArtifactSkip Run/report artifact excluded from rename candidates: {0}" -f $file.FullName)
                continue
            }

            if ($SkipLikelyAlreadyRenamed -and (Test-LikelyAlreadyRenamedName -FileName $file.Name)) {
                $heuristicSkipped++
                Write-SkippedLiveLine -FileName $file.Name -Reason 'Likely already renamed (heuristic)'
                Write-RunLog -Message ("HeuristicSkip Name appears already-renamed: {0}" -f $file.FullName)
                continue
            }

            $fileStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $metadataName = Get-RenameBaseNameFromMetadata -File $file -Shell $shell -Force:([bool]$Force) -OnlyGeneric:([bool]$OnlyGeneric)
            $fileStopwatch.Stop()

            # Hard-kill the file if the COM call takes longer than the budget
            # (The stopwatch check below handles this since the COM release fixes prevent true hangs;
            #  but if a COM call still manages to hang in the future, the memory watchdog
            #  will trigger a graceful exit rather than letting the OS kill the process.)

            if ($fileStopwatch.Elapsed.TotalSeconds -gt $PerFileTimeoutSeconds) {
                $timeoutSkipped++
                Write-TimeoutLiveLine -FileName $file.Name -ElapsedSeconds $fileStopwatch.Elapsed.TotalSeconds -TimeoutSeconds $PerFileTimeoutSeconds
                Write-RunLog -Message (
                    "TimeoutSkip Metadata evaluation exceeded budget for {0}. Elapsed={1:N3}s; Limit={2}s" -f
                    $file.FullName,
                    $fileStopwatch.Elapsed.TotalSeconds,
                    $PerFileTimeoutSeconds
                )
                continue
            }

            if (-not $metadataName) {
                $skipped++
                $consecutiveErrors = 0
                continue
            }

            $safeBaseName = Get-SafeFileName -Name $metadataName.Name
            if (-not $safeBaseName) {
                $skipped++
                continue
            }

            $targetPath = Get-UniqueTargetPath -Directory $file.DirectoryName -BaseName $safeBaseName -Extension $file.Extension -CurrentFullName $file.FullName -ReservedPaths $reservedPaths

            if ($targetPath -ieq $file.FullName) {
                $skipped++
                continue
            }

            # Metadata upgrade check: only re-rename if new score is strictly higher
            if ($script:ResumeActive -and $script:ResumeRenamedTo.Contains($file.FullName) -and -not $SkipMetadataUpgrade) {
                $prevInfo = $null
                if ($script:ResumeRenamedScores.TryGetValue($file.FullName, [ref]$prevInfo)) {
                    if ($metadataName.ConfidenceScore -le $prevInfo.Score) {
                        $skipped++
                        Write-RunLog -Message ("UpgradeSkip No improvement for {0}. New={1}/{2}; Previous={3}/{4}" -f $file.FullName, $metadataName.Source, $metadataName.ConfidenceScore, $prevInfo.Source, $prevInfo.Score)
                        continue
                    }
                    Write-RunLog -Message ("MetadataUpgrade Upgrading {0}. New={1}/{2}; Previous={3}/{4}" -f $file.FullName, $metadataName.Source, $metadataName.ConfidenceScore, $prevInfo.Source, $prevInfo.Score)
                }
            }

            $targetName = [System.IO.Path]::GetFileName($targetPath)
            $plan.Add([PSCustomObject]@{
                FullPath     = $file.FullName
                RelativePath = $file.FullName.Substring($root.Length).TrimStart('\\')
                OldName      = $file.Name
                NewName      = $targetName
                Source       = $metadataName.Source
                ConfidenceScore = $metadataName.ConfidenceScore
                ConfidenceLevel = $metadataName.ConfidenceLevel
                ScoreReason  = $metadataName.ScoreReason
            }) | Out-Null

            $lastAdded = $plan[$plan.Count - 1]
            $consecutiveErrors = 0
            if ($isImmediateRenameMode) {
                $immediateTargetPath = Join-Path -Path (Split-Path -Path $lastAdded.FullPath -Parent) -ChildPath $lastAdded.NewName
                if ($script:ResumeActive -and $script:ResumeRenamedFrom.Contains($lastAdded.FullPath)) {
                    $resumeSkipped++
                    continue
                }
                if ($script:ResumeActive -and $SkipMetadataUpgrade -and $script:ResumeRenamedTo.Contains($immediateTargetPath)) {
                    $resumeSkipped++
                    continue
                }

                if ($PSCmdlet.ShouldProcess($lastAdded.FullPath, "Rename to '$($lastAdded.NewName)'")) {
                    Rename-Item -LiteralPath $lastAdded.FullPath -NewName $lastAdded.NewName -ErrorAction Stop
                    $renamed++
                    Write-RenameLiveLine -Action 'RENAMED' -Entry $lastAdded
                }
            }
            else {
                $planAction = if ($isApplyMode) { 'PLANNED' } else { 'PREVIEW' }
                Write-RenameLiveLine -Action $planAction -Entry $lastAdded
            }
        }
        catch {
            if ($fileStopwatch -and $fileStopwatch.IsRunning) {
                $fileStopwatch.Stop()
            }

            $errorCount++
            $consecutiveErrors++
            $errorMessage = if ($_.Exception -and -not [string]::IsNullOrWhiteSpace($_.Exception.Message)) { $_.Exception.Message } else { $_.ToString() }
            Write-ErrorLiveLine -FileName $file.Name -Reason 'Processing error (see run log)'
            $null = $processingErrorFiles.Add($file.FullName)
            $errorMessageFlat = (($errorMessage -replace '\|', '/') -replace "`r?`n", ' ').Trim()
            Write-RunLog -Message ("ERROR | Phase=processing | File={0} | NewName= | Message={1}" -f $file.FullName, $errorMessageFlat)
            Write-RunLog -Message ("ErrorSkip File processing failed for {0}. Error={1}" -f $file.FullName, $errorMessage)

            if ($consecutiveErrors -ge 50) {
                Write-RunLog -Message ("CIRCUIT BREAKER: {0} consecutive errors - force-recycling COM and running GC" -f $consecutiveErrors)
                Write-Host ("  [!] Circuit breaker: {0} consecutive errors - recycling COM..." -f $consecutiveErrors) -ForegroundColor Red
                try {
                    Reset-ShellMetadataRunspace
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    $shell = New-Object -ComObject Shell.Application
                    Write-RunLog -Message 'Circuit breaker: COM recycled successfully.'
                }
                catch {
                    Write-RunLog -Message 'Circuit breaker: Failed to recycle COM.'
                }
                $consecutiveErrors = 0
            }
            elseif ($errorMessage -match '(?i)com object|rpc server|disconnected|invalid|out of memory|e_fail') {
                try {
                    Reset-ShellMetadataRunspace
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    $shell = New-Object -ComObject Shell.Application
                    Write-RunLog -Message 'Recovered Shell.Application COM object after processing error.'
                }
                catch {
                    Write-RunLog -Message 'Failed to recover Shell.Application COM object after processing error.'
                }
            }

            continue
        }
    }

    if ($script:StopRequested) {
        break
    }

    if ($totalItemBatches -gt 1) {
        Write-RunLog -Message (
            "BatchComplete | Phase=scan | Batch={0}/{1}" -f
            $currentBatchNumber,
            $totalItemBatches
        )

        # Force GC at batch boundaries to reduce memory pressure
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Write-RunLog -Message (
    "HEARTBEAT_FINAL | Phase=scan | Processed={0}/{1} | Candidates={2} | Renamed={3} | Skipped={4} | ResumeSkipped={5} | Errors={6} | Timeouts={7} | ArtifactSkipped={8} | HeuristicSkipped={9} | ElapsedSeconds={10:N1}" -f
    $scanProcessed,
    $totalItems,
    $plan.Count,
    $renamed,
    $skipped,
    $resumeSkipped,
    $errorCount,
    $timeoutSkipped,
    $artifactSkipped,
    $heuristicSkipped,
    $scanHeartbeatStopwatch.Elapsed.TotalSeconds
)

if ($processingErrorFiles.Count -gt 0) {
    $retryCurrentErrorsNow = Read-YesNo -Prompt ("Retry current ERROR files now before finalizing report? ({0} files)" -f $processingErrorFiles.Count) -DefaultYes $true
    if ($retryCurrentErrorsNow) {
        Write-RunLog -Message ("Retrying current processing ERROR files. Count={0}" -f $processingErrorFiles.Count)
        foreach ($retryFilePath in @($processingErrorFiles)) {
            try {
                if (-not (Test-Path -LiteralPath $retryFilePath)) {
                    Write-RunLog -Message ("ERROR retry skipped, file no longer exists: {0}" -f $retryFilePath)
                    continue
                }

                $retryFile = Get-Item -LiteralPath $retryFilePath -ErrorAction Stop
                if (-not ($retryFile -is [System.IO.FileInfo])) {
                    continue
                }

                $fileStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $metadataName = Get-RenameBaseNameFromMetadata -File $retryFile -Shell $shell -Force:([bool]$Force) -OnlyGeneric:([bool]$OnlyGeneric)
                $fileStopwatch.Stop()

                if ($fileStopwatch.Elapsed.TotalSeconds -gt $PerFileTimeoutSeconds) {
                    $timeoutSkipped++
                    Write-TimeoutLiveLine -FileName $retryFile.Name -ElapsedSeconds $fileStopwatch.Elapsed.TotalSeconds -TimeoutSeconds $PerFileTimeoutSeconds
                    continue
                }

                if (-not $metadataName) {
                    continue
                }

                $safeBaseName = Get-SafeFileName -Name $metadataName.Name
                if (-not $safeBaseName) {
                    continue
                }

                $targetPath = Get-UniqueTargetPath -Directory $retryFile.DirectoryName -BaseName $safeBaseName -Extension $retryFile.Extension -CurrentFullName $retryFile.FullName -ReservedPaths $reservedPaths
                if ($targetPath -ieq $retryFile.FullName) {
                    continue
                }

                $targetName = [System.IO.Path]::GetFileName($targetPath)
                $retryEntry = [PSCustomObject]@{
                    FullPath     = $retryFile.FullName
                    RelativePath = $retryFile.FullName.Substring($root.Length).TrimStart('\\')
                    OldName      = $retryFile.Name
                    NewName      = $targetName
                    Source       = $metadataName.Source
                    ConfidenceScore = $metadataName.ConfidenceScore
                    ConfidenceLevel = $metadataName.ConfidenceLevel
                    ScoreReason  = $metadataName.ScoreReason
                }

                if ($isImmediateRenameMode) {
                    if ($PSCmdlet.ShouldProcess($retryEntry.FullPath, "Retry rename to '$($retryEntry.NewName)'")) {
                        Rename-Item -LiteralPath $retryEntry.FullPath -NewName $retryEntry.NewName -ErrorAction Stop
                        $renamed++
                        $errorCount = [Math]::Max(0, $errorCount - 1)
                        Write-RenameLiveLine -Action 'RENAMED' -Entry $retryEntry
                        Write-RunLog -Message ("ERROR_RETRY_SUCCESS | Phase=processing | File={0}" -f $retryFile.FullName)
                    }
                }
                else {
                    $plan.Add($retryEntry) | Out-Null
                    $planAction = if ($isApplyMode) { 'PLANNED' } else { 'PREVIEW' }
                    Write-RenameLiveLine -Action $planAction -Entry $retryEntry
                    $errorCount = [Math]::Max(0, $errorCount - 1)
                    Write-RunLog -Message ("ERROR_RETRY_SUCCESS | Phase=processing | File={0}" -f $retryFile.FullName)
                }
            }
            catch {
                $retryMessage = if ($_.Exception -and -not [string]::IsNullOrWhiteSpace($_.Exception.Message)) { $_.Exception.Message } else { $_.ToString() }
                Write-ErrorLiveLine -FileName ([System.IO.Path]::GetFileName($retryFilePath)) -Reason 'Retry failed (see run log)'
                $retryMessageFlat = (($retryMessage -replace '\|', '/') -replace "`r?`n", ' ').Trim()
                Write-RunLog -Message ("ERROR_RETRY_FAILED | Phase=processing | File={0} | Message={1}" -f $retryFilePath, $retryMessageFlat)
            }
        }
    }
}

$lines = [System.Collections.Generic.List[string]]::new()
$reportKind = ''
if ($isImmediateRenameMode) {
    $reportKind = 'Rename'
} elseif ($isApplyMode) {
    $reportKind = 'Plan'
} else {
    $reportKind = 'Preview'
}
$lines.Add("$reportKind generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
$lines.Add("Root path: $root") | Out-Null
$lines.Add("") | Out-Null

if ($plan.Count -eq 0) {
    if ($isImmediateRenameMode) {
        $lines.Add('No files were renamed.') | Out-Null
    }
    else {
        $lines.Add('No rename candidates found.') | Out-Null
    }
}
else {
    if ($isImmediateRenameMode) {
        $lines.Add('Renamed files:') | Out-Null
    }
    else {
        $lines.Add('Planned renames:') | Out-Null
    }
    foreach ($entry in $plan) {
        $line = "{0} -> {1} [{2}] [Confidence: {3}%/{4}] [{5}]" -f $entry.OldName, $entry.NewName, $entry.Source, $entry.ConfidenceScore, $entry.ConfidenceLevel, $entry.ScoreReason
        $lines.Add($line) | Out-Null
    }
}

$lines.Add("") | Out-Null
if ($isImmediateRenameMode) {
    $lines.Add(("Renamed: {0} | Skipped: {1} | [Uncommon]: {2} | ResumeSkipped: {3} | ArtifactSkipped: {4} | HeuristicSkipped: {5} | ERROR: {6} | [TIMEOUT]: {7}" -f $renamed, $skipped, $uncommonSkipped, $resumeSkipped, $artifactSkipped, $heuristicSkipped, $errorCount, $timeoutSkipped)) | Out-Null
}
else {
    $lines.Add(("Candidates: {0} | Skipped: {1} | [Uncommon]: {2} | ResumeSkipped: {3} | ArtifactSkipped: {4} | HeuristicSkipped: {5} | ERROR: {6} | [TIMEOUT]: {7}" -f $plan.Count, $skipped, $uncommonSkipped, $resumeSkipped, $artifactSkipped, $heuristicSkipped, $errorCount, $timeoutSkipped)) | Out-Null
}

$lines.Add('') | Out-Null
$lines.Add('Score Matrix:') | Out-Null
foreach ($scoreSource in $script:ScoreMatrix.Keys) {
    $matrixItem = $script:ScoreMatrix[$scoreSource]
    $matrixLine = "- {0}: {1}% ({2})" -f $scoreSource, [int]$matrixItem.Score, $matrixItem.Reason
    $lines.Add($matrixLine) | Out-Null
}

Set-Content -LiteralPath $ReportPath -Value $lines -Encoding UTF8
$plan | Select-Object RelativePath, OldName, NewName, Source, ConfidenceScore, ConfidenceLevel, ScoreReason, FullPath | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Host ""
if ($isImmediateRenameMode) {
    Write-Host ("Rename report written: {0}" -f $ReportPath)
    Write-Host ("Rename CSV written: {0}" -f $CsvPath)
}
elseif ($isApplyMode) {
    Write-Host ("Plan report written: {0}" -f $ReportPath)
    Write-Host ("Plan CSV written: {0}" -f $CsvPath)
}
else {
    Write-Host ("Preview report written: {0}" -f $ReportPath)
    Write-Host ("Preview CSV written: {0}" -f $CsvPath)
}
Write-Host ("Run log written: {0}" -f $script:LogPath)
if ($isImmediateRenameMode) {
    Write-RunLog -Message ("Immediate rename complete. Renamed={0}; Skipped={1}; UncommonSkipped={2}; ResumeSkipped={3}; ArtifactSkipped={4}; HeuristicSkipped={5}; ERROR={6}; TIMEOUT={7}" -f $renamed, $skipped, $uncommonSkipped, $resumeSkipped, $artifactSkipped, $heuristicSkipped, $errorCount, $timeoutSkipped)
}
elseif ($isApplyMode) {
    Write-RunLog -Message ("Planning complete for apply mode. Candidates={0}; Skipped={1}; UncommonSkipped={2}; ResumeSkipped={3}; ArtifactSkipped={4}; HeuristicSkipped={5}; ERROR={6}; TIMEOUT={7}" -f $plan.Count, $skipped, $uncommonSkipped, $resumeSkipped, $artifactSkipped, $heuristicSkipped, $errorCount, $timeoutSkipped)
}
else {
    Write-RunLog -Message ("Preview complete. Candidates={0}; Skipped={1}; UncommonSkipped={2}; ResumeSkipped={3}; ArtifactSkipped={4}; HeuristicSkipped={5}; ERROR={6}; TIMEOUT={7}" -f $plan.Count, $skipped, $uncommonSkipped, $resumeSkipped, $artifactSkipped, $heuristicSkipped, $errorCount, $timeoutSkipped)
}

if ($script:StopRequested) {
    if ($isImmediateRenameMode) {
        Write-Host 'Run stopped by user during immediate rename phase.' -ForegroundColor Yellow
        Write-RunLog -Message 'Run stopped by user during immediate rename phase.'
    }
    elseif ($isApplyMode) {
        Write-Host 'Run stopped by user during planning phase.' -ForegroundColor Yellow
        Write-RunLog -Message 'Run stopped by user during planning phase.'
    }
    else {
        Write-Host 'Run stopped by user during preview phase.' -ForegroundColor Yellow
        Write-RunLog -Message 'Run stopped by user during preview phase.'
    }
    Write-EndOfLogCode -Status 'Run stopped by user'
    return
}

if (-not $isApplyMode) {
    if ($plan.Count -eq 0) {
        Write-Host 'Preview only. No files were renamed.'
        Write-RunLog -Message 'Preview-only mode ended without rename operations (no candidates).' 
        Write-EndOfLogCode -Status 'Preview ended without candidates'
        return
    }

    $applyPrompt = if ($QuickPreview) {
        'Quick preview complete. Continue to rename these sampled files now?'
    }
    else {
        'Apply these planned renames now?'
    }

    $applyAfterPreview = Read-YesNo -Prompt $applyPrompt -DefaultYes $false
    if (-not $applyAfterPreview) {
        Write-Host 'Preview only. No files were renamed.'
        Write-Host "Tip: run with -Apply to execute renames immediately next time."
        Write-RunLog -Message 'Preview-only mode ended without rename operations.'
        Write-EndOfLogCode -Status 'Preview ended without apply confirmation'
        return
    }

    Write-Host 'Applying planned renames from preview...' -ForegroundColor Yellow
    Write-RunLog -Message 'User confirmed apply at end-of-preview prompt.'
}

if (-not $QuickPreview -and -not $isImmediateRenameMode -and $plan.Count -ge $script:LongRunThreshold) {
    Write-Host (
        "Large rename set detected ({0} files). Runtime controls: P = pause, C = continue (while paused), S = stop gracefully." -f $plan.Count
    ) -ForegroundColor Yellow

    $continueLongRename = Read-YesNo -Prompt 'Proceed with this long rename run?' -DefaultYes $true
    if (-not $continueLongRename) {
        Write-Host 'Rename phase canceled before execution. Preview files are still available.' -ForegroundColor Yellow
        Write-RunLog -Message ("User canceled long rename before apply. PlannedRenames={0}" -f $plan.Count)
        Write-EndOfLogCode -Status 'Rename canceled by user before apply'
        return
    }
}

if (-not $isImmediateRenameMode) {
    $planEntries = @($plan)
    $totalPlanEntries = $planEntries.Count
    $safeBatchSize = Convert-LogSettingToInt $BatchSize 100 1000000
    $totalRenameBatches = if ($totalPlanEntries -gt 0) { [int][Math]::Ceiling($totalPlanEntries / [double]$safeBatchSize) } else { 0 }

    for ($renameBatchStart = 0; $renameBatchStart -lt $totalPlanEntries; $renameBatchStart += $safeBatchSize) {
        $renameBatchEnd = [Math]::Min($renameBatchStart + $safeBatchSize - 1, $totalPlanEntries - 1)
        $currentRenameBatch = [int]([Math]::Floor($renameBatchStart / [double]$safeBatchSize) + 1)
        if ($totalRenameBatches -gt 1) {
            Write-Host ("Applying batch {0}/{1} renames..." -f $currentRenameBatch, $totalRenameBatches) -ForegroundColor DarkYellow
            Write-RunLog -Message ("BatchStart | Phase=rename | Batch={0}/{1}" -f $currentRenameBatch, $totalRenameBatches)
        }

        for ($ri = $renameBatchStart; $ri -le $renameBatchEnd; $ri++) {
            $entry = $planEntries[$ri]

            if (-not (Test-RunControl -Phase 'rename')) {
                break
            }

            $renameProcessed++

            # Update window title on every rename
            Update-WindowTitle -Phase 'Apply' -Processed $renameProcessed -Total $totalPlanEntries -Renamed $renamed -Errors $errorCount -ElapsedSeconds $renameHeartbeatStopwatch.Elapsed.TotalSeconds

            if (($renameProcessed % $heartbeatEvery) -eq 0) {
                # Visible console heartbeat for rename phase
                $hbElapsed = $renameHeartbeatStopwatch.Elapsed
                $hbPct = if ($totalPlanEntries -gt 0) { [int](($renameProcessed / $totalPlanEntries) * 100) } else { 0 }
                Write-Host (
                    "  [{0}] Rename progress: {1}/{2} ({3}%) | Renamed:{4} Errors:{5} | Elapsed:{6:N0}m {7:N0}s" -f
                    (Get-Date -Format 'HH:mm:ss'),
                    $renameProcessed,
                    $totalPlanEntries,
                    $hbPct,
                    $renamed,
                    $errorCount,
                    [Math]::Floor($hbElapsed.TotalMinutes),
                    $hbElapsed.Seconds
                ) -ForegroundColor DarkGray

                Write-RunLog -Message (
                    "HEARTBEAT | Phase=rename | Processed={0}/{1} | Renamed={2} | Errors={3} | ElapsedSeconds={4:N1}" -f
                    $renameProcessed,
                    $totalPlanEntries,
                    $renamed,
                    $errorCount,
                    $renameHeartbeatStopwatch.Elapsed.TotalSeconds
                )
            }

            if ($script:ResumeActive) {
                $renameTargetPath = Join-Path -Path (Split-Path -Path $entry.FullPath -Parent) -ChildPath $entry.NewName
                if ($script:ResumeRenamedFrom.Contains($entry.FullPath) -or $script:ResumeRenamedTo.Contains($renameTargetPath)) {
                    $resumeSkipped++
                    continue
                }
            }

            try {
                if (-not (Test-Path -LiteralPath $entry.FullPath)) {
                    Write-RunLog -Message ("RenameSkip File no longer exists: {0}" -f $entry.FullPath)
                    continue
                }

                if ($PSCmdlet.ShouldProcess($entry.FullPath, "Rename to '$($entry.NewName)'")) {
                    Rename-Item -LiteralPath $entry.FullPath -NewName $entry.NewName -ErrorAction Stop
                    $renamed++
                    Write-RenameLiveLine -Action 'RENAMED' -Entry $entry
                }
            }
            catch {
                $errorCount++
                $renameMessage = if ($_.Exception -and -not [string]::IsNullOrWhiteSpace($_.Exception.Message)) { $_.Exception.Message } else { $_.ToString() }
                Write-ErrorLiveLine -FileName $entry.OldName -Reason 'Rename failed (see run log)'
                $renameMessageFlat = (($renameMessage -replace '\|', '/') -replace "`r?`n", ' ').Trim()
                Write-RunLog -Message ("ERROR | Phase=rename | File={0} | NewName={1} | Message={2}" -f $entry.FullPath, $entry.NewName, $renameMessageFlat)
                $renameErrorEntries.Add($entry) | Out-Null
            }
        }

        if ($script:StopRequested) {
            break
        }

        if ($totalRenameBatches -gt 1) {
            Write-RunLog -Message ("BatchComplete | Phase=rename | Batch={0}/{1}" -f $currentRenameBatch, $totalRenameBatches)

            # Force GC at rename batch boundaries
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    Write-RunLog -Message (
        "HEARTBEAT_FINAL | Phase=rename | Processed={0}/{1} | Renamed={2} | ResumeSkipped={3} | Errors={4} | ElapsedSeconds={5:N1}" -f
        $renameProcessed,
        $totalPlanEntries,
        $renamed,
        $resumeSkipped,
        $errorCount,
        $renameHeartbeatStopwatch.Elapsed.TotalSeconds
    )
}

if ($renameErrorEntries.Count -gt 0) {
    $retryRenameErrorsNow = Read-YesNo -Prompt ("Retry rename ERROR files now? ({0} files)" -f $renameErrorEntries.Count) -DefaultYes $true
    if ($retryRenameErrorsNow) {
        foreach ($failedEntry in $renameErrorEntries) {
            try {
                if (-not (Test-Path -LiteralPath $failedEntry.FullPath)) {
                    continue
                }

                if ($PSCmdlet.ShouldProcess($failedEntry.FullPath, "Retry rename to '$($failedEntry.NewName)'")) {
                    Rename-Item -LiteralPath $failedEntry.FullPath -NewName $failedEntry.NewName -ErrorAction Stop
                    $renamed++
                    $errorCount = [Math]::Max(0, $errorCount - 1)
                    Write-RunLog -Message ("ERROR_RETRY_SUCCESS | Phase=rename | File={0} | NewName={1}" -f $failedEntry.FullPath, $failedEntry.NewName)
                }
            }
            catch {
                $retryRenameMessage = if ($_.Exception -and -not [string]::IsNullOrWhiteSpace($_.Exception.Message)) { $_.Exception.Message } else { $_.ToString() }
                $retryRenameFlat = (($retryRenameMessage -replace '\|', '/') -replace "`r?`n", ' ').Trim()
                Write-RunLog -Message ("ERROR_RETRY_FAILED | Phase=rename | File={0} | NewName={1} | Message={2}" -f $failedEntry.FullPath, $failedEntry.NewName, $retryRenameFlat)
            }
        }
    }
}

Write-Host ""
Write-Host ("Done. Renamed: {0} | Skipped: {1} | [Uncommon]: {2} | ResumeSkipped: {3} | ArtifactSkipped: {4} | HeuristicSkipped: {5} | ERROR: {6} | [TIMEOUT]: {7}" -f $renamed, $skipped, $uncommonSkipped, $resumeSkipped, $artifactSkipped, $heuristicSkipped, $errorCount, $timeoutSkipped)
if ($script:StopRequested) {
    Write-Host 'Rename phase stopped early by user request.' -ForegroundColor Yellow
    Write-RunLog -Message 'Rename phase stopped early by user request.'
}
if ($isImmediateRenameMode) {
    Write-Host ("Rename report: {0}" -f $ReportPath)
    Write-Host ("Rename CSV: {0}" -f $CsvPath)
}
elseif ($isApplyMode) {
    Write-Host ("Plan report: {0}" -f $ReportPath)
    Write-Host ("Plan CSV: {0}" -f $CsvPath)
}
else {
    Write-Host ("Preview report: {0}" -f $ReportPath)
    Write-Host ("Preview CSV: {0}" -f $CsvPath)
}
Write-Host ("Run log: {0}" -f $script:LogPath)
Write-RunLog -Message ("Apply complete. Renamed={0}; Skipped={1}; UncommonSkipped={2}; ResumeSkipped={3}; ArtifactSkipped={4}; HeuristicSkipped={5}; ERROR={6}; TIMEOUT={7}" -f $renamed, $skipped, $uncommonSkipped, $resumeSkipped, $artifactSkipped, $heuristicSkipped, $errorCount, $timeoutSkipped)
if ($script:_shellMetadataRunspace) {
    try { $script:_shellMetadataRunspace.Close() } catch {}
    try { $script:_shellMetadataRunspace.Dispose() } catch {}
    $script:_shellMetadataRunspace = $null
}
Restore-WindowTitle
Write-EndOfLogCode -Status 'Run completed'
