
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

Interactive flow can prompt for target path, preview/apply mode, quick preview mode/limit,
if (-not $isImmediateRenameMode) {
- Shell metadata queries and embedded media metadata extraction can be slower on large trees.
- Run preview first for large datasets, then execute apply mode after review.

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

.PARAMETER Path
Root folder to scan for rename candidates.
If omitted, interactive mode defaults to the current working directory ($PWD.Path).

.PARAMETER Help
Displays command-line parameter help and exits.
Useful for both script and EXE usage.

.PARAMETER NoRecurse
Disables recursive scanning.
By default, recursion is enabled unless -NoRecurse is specified.

.PARAMETER InteractiveSubfolders
Prompts for each discovered subfolder to include or skip during scanning.
Cannot be combined with -NoRecurse.

.PARAMETER Apply
Executes renames from the generated plan immediately after planning.
If omitted, preview-first mode runs and can optionally confirm apply at end.

.PARAMETER ImmediateRename
When used with -Apply, renames files immediately during processing instead of showing
planned entries first and then running a separate rename pass.

.PARAMETER IncludeHidden
Includes hidden files during scanning.

.PARAMETER Force
Forces metadata naming attempts regardless of generic filename heuristics.

.PARAMETER OnlyGeneric
Only attempts renames for filenames that look generic (for example file123, img_0001, dsc0002).

.PARAMETER ReportPath
Optional path for TXT report output.
If omitted, a timestamped TXT report file is created in the scan root.
Backward-compatible alias: -PreviewReportPath.

.PARAMETER CsvPath
Optional path for CSV report output.
If omitted, a CSV path is derived from the TXT report path.
Backward-compatible alias: -PreviewCsvPath.

.PARAMETER QuickPreview
Limits discovery to a small sample set for faster preview validation.
By default this previews the sampled set and then offers an end-of-preview apply prompt.
If used with -Apply, renames are applied to the sampled file set without the preview-only stop.

.PARAMETER QuickPreviewLimit
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
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Position = 0)]
    [string]$Path,

    [Alias('?')]
    [switch]$Help,

    [switch]$NoRecurse,

    [switch]$InteractiveSubfolders,

    [switch]$IncludeHidden,

    [switch]$Force,

    [switch]$OnlyGeneric,

    [switch]$Apply,

    [switch]$ImmediateRename,

    [switch]$QuickPreview,

    [ValidateRange(1, 100000)]
    [int]$QuickPreviewLimit = 200,

    [Alias('PreviewReportPath')]
    [string]$ReportPath,

    [Alias('PreviewCsvPath')]
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

    [switch]$SkipLikelyAlreadyRenamed,

    [ValidateRange(5, 600)]
    [int]$PerFileTimeoutSeconds = 20,

    [ValidateRange(100, 50000)]
    [int]$BatchSize = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScoreMatrix = [ordered]@{
    'DOCX core title'              = [PSCustomObject]@{ Score = 98; Reason = 'Embedded canonical document title' }
    'DOCX heading title'           = [PSCustomObject]@{ Score = 92; Reason = 'Document heading structure from app.xml' }
    'Word OpenXML core title'      = [PSCustomObject]@{ Score = 98; Reason = 'Embedded Word OpenXML canonical document title' }
    'Word heading title'           = [PSCustomObject]@{ Score = 92; Reason = 'Word heading structure from app.xml TitlesOfParts' }
    'RTF title control word'       = [PSCustomObject]@{ Score = 96; Reason = 'Embedded RTF \title metadata control word' }
    'Word COM title'               = [PSCustomObject]@{ Score = 90; Reason = 'Word built-in document title via COM automation' }
    'PDF XMP title'                = [PSCustomObject]@{ Score = 94; Reason = 'Embedded PDF XMP title metadata' }
    'PDF bookmark title'           = [PSCustomObject]@{ Score = 89; Reason = 'First PDF outline/bookmark title' }
    'XML title node'               = [PSCustomObject]@{ Score = 93; Reason = 'Structured XML title element' }
    'Image embedded XMP title'     = [PSCustomObject]@{ Score = 95; Reason = 'Embedded authored image title metadata' }
    'Windows metadata title'       = [PSCustomObject]@{ Score = 88; Reason = 'Windows semantic metadata title field' }
    'Audio composite metadata'     = [PSCustomObject]@{ Score = 91; Reason = 'Artist/title/album/track composite metadata' }
    'Audio title metadata'         = [PSCustomObject]@{ Score = 85; Reason = 'Audio title metadata only' }
    'Video episode metadata'       = [PSCustomObject]@{ Score = 87; Reason = 'Season/episode/media title metadata' }
    'GPS location metadata'        = [PSCustomObject]@{ Score = 80; Reason = 'Location metadata from GPS fields' }
    'Media date metadata'          = [PSCustomObject]@{ Score = 72; Reason = 'Date-based media fallback naming' }
    'Author/company metadata'      = [PSCustomObject]@{ Score = 62; Reason = 'Attribution fallback metadata' }
}

$script:LogPath = $null
$script:StopRequested = $false
$script:LongRunThreshold = 500
$script:ResumeActive = $false
$script:ResumeLogSource = $null
$script:ResumeRenamedFrom = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:ResumeRenamedTo = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:EndOfLogWritten = $false
$script:LastUsedPathStateFile = Join-Path -Path $env:APPDATA -ChildPath 'Rename-ByMetadata\last-scan-path.txt'

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
        [bool]$Recurse,

        [Parameter(Mandatory)]
        [string]$Mode,

        [Parameter(Mandatory)]
        [bool]$Admin,

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
        [bool]$IncludeHidden,

        [Parameter(Mandatory)]
        [bool]$Recurse,

        [Parameter(Mandatory)]
        [bool]$InteractiveSubfolders
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

    try {
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
                        default {
                            continue
                        }
                    }
                }
            }
            'S' {
                $script:StopRequested = $true
                Write-Host ("Stop requested during {0} phase. Ending run gracefully..." -f $Phase) -ForegroundColor Yellow
                Write-RunLog -Message ("Stop requested during {0} phase by user hotkey." -f $Phase)
                return $false
            }
            default {
                return $true
            }
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

    $cleanName = Normalize-MetadataTitle -Title $Name
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
        $candidate = Get-ChildItem -LiteralPath $Root -File -Filter 'Rename-Run-*.log' -ErrorAction Stop |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1
        if ($candidate) {
            return $candidate.FullName
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

    $pattern = 'RENAMED\b.*\|\s*From=(?<from>.+?)\s*\|\s*To=(?<to>.+?)\s*$'

    try {
        foreach ($line in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
            if ($line -notmatch $pattern) {
                continue
            }

            $fromPath = $Matches['from'].Trim()
            $toPath = $Matches['to'].Trim()

            if (-not [string]::IsNullOrWhiteSpace($fromPath)) {
                $null = $script:ResumeRenamedFrom.Add($fromPath)
            }
            if (-not [string]::IsNullOrWhiteSpace($toPath)) {
                $null = $script:ResumeRenamedTo.Add($toPath)
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
        $parsedLimit = 0
        if (-not [string]::IsNullOrWhiteSpace($limitRaw) -and $limitRaw -ne 'N/A' -and [int]::TryParse($limitRaw, [ref]$parsedLimit) -and $parsedLimit -ge 1 -and $parsedLimit -le 100000) {
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
        $script:SkipLikelyAlreadyRenamed = Convert-LogSettingToBool -Value $Settings['SkipLikelyAlreadyRenamed'] -Default $false
        $appliedKeys.Add('SkipLikelyAlreadyRenamed') | Out-Null
    }

    if (-not $BoundParameters.ContainsKey('PerFileTimeoutSeconds') -and $Settings.Contains('PerFileTimeoutSeconds')) {
        $timeoutRaw = $Settings['PerFileTimeoutSeconds']
        $parsedTimeout = 0
        if (-not [string]::IsNullOrWhiteSpace($timeoutRaw) -and [int]::TryParse($timeoutRaw, [ref]$parsedTimeout) -and $parsedTimeout -ge 5 -and $parsedTimeout -le 600) {
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
        '^(Location\s+.+|.+\s+-\s+Location\s+.+)$'
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
    Write-Host '               [-SkipLikelyAlreadyRenamed] [-PerFileTimeoutSeconds <5-600>] [-BatchSize <100-50000>]'
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

$retryErrorsRequested = $RetryErrors.IsPresent -or -not [string]::IsNullOrWhiteSpace($RetryErrorsLogPath)

$defaultPath = Get-LastUsedScanPath
if ([string]::IsNullOrWhiteSpace($defaultPath)) {
    $defaultPath = $PWD.Path
}

if ([string]::IsNullOrWhiteSpace($Path)) {
    $enteredPath = Read-Host "Enter folder path to scan (press Enter for default: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($enteredPath)) {
        $Path = $defaultPath
    }
    else {
        $Path = $enteredPath.Trim()
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
        if ($candidateStatus.Eligible) {
            $resumePrompt = Read-YesNo -Prompt (
                "Previous run appears interrupted ({0} completed renames). Resume from '{1}'?" -f
                $candidateStatus.StructuredRenamedEntries,
                $latestRunLog
            ) -DefaultYes $true

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
            $usePreviousSettings = $true
            if ($resumeRequestedViaPrompt) {
                $usePreviousSettings = Read-YesNo -Prompt (
                    "Resume log includes previous job settings. Continue with those settings?" 
                ) -DefaultYes $true
            }

            if ($usePreviousSettings) {
                $appliedSettings = Apply-SelectedSettingsFromLog -Settings $selectedSettings -BoundParameters $PSBoundParameters
                $resumeSettingsApplied = @($appliedSettings).Count -gt 0
                if ($resumeSettingsApplied) {
                    Write-Host (
                        "Loaded {0} setting(s) from resume log: {1}" -f
                        @($appliedSettings).Count,
                        $resolvedResumeSettingsLog
                    ) -ForegroundColor Yellow
                }
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
                $appliedRetrySettings = Apply-SelectedSettingsFromLog -Settings $retrySettings -BoundParameters $PSBoundParameters
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

if ($InteractiveSubfolders.IsPresent -and $NoRecurse.IsPresent) {
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

if ($IncludeOtherFiles.IsPresent -and $ScanOtherFilesAtEnd.IsPresent) {
    throw 'Cannot combine -IncludeOtherFiles with -ScanOtherFilesAtEnd.'
}

$isApplyMode = [bool]$Apply
$isImmediateRenameMode = $isApplyMode -and [bool]$ImmediateRename

if (-not $PSBoundParameters.ContainsKey('QuickPreview') -and -not $logSettingsApplied) {
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

$intelligentScanEnabled = [bool]$IntelligentScan
$includeOtherFilesNow = [bool]$IncludeOtherFiles
$scanOtherFilesAtEndEnabled = [bool]$ScanOtherFilesAtEnd

if ($includeOtherFilesNow) {
    $intelligentScanEnabled = $false
    $scanOtherFilesAtEndEnabled = $false
}



# Add a new autodiscovery mode: F = Use folder names as extensions
$useExtensionFolders = $false
if (-not $retryErrorsRequested -and -not $PSBoundParameters.ContainsKey('IntelligentScan') -and -not $PSBoundParameters.ContainsKey('IncludeOtherFiles') -and -not $logSettingsApplied) {
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
        if (-not $PSBoundParameters.ContainsKey('InteractiveSubfolders')) {
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

if ($InteractiveSubfolders.IsPresent) {
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
    if ($InteractiveSubfolders.IsPresent) { $argsList.Add('-InteractiveSubfolders') | Out-Null }
    if ($IncludeHidden.IsPresent) { $argsList.Add('-IncludeHidden') | Out-Null }
    if ($Force.IsPresent) { $argsList.Add('-Force') | Out-Null }
    if ($OnlyGeneric.IsPresent) { $argsList.Add('-OnlyGeneric') | Out-Null }
    if ($Apply.IsPresent) { $argsList.Add('-Apply') | Out-Null }
    if ($ImmediateRename.IsPresent) { $argsList.Add('-ImmediateRename') | Out-Null }
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
    if ($SkipLikelyAlreadyRenamed.IsPresent) {
        $argsList.Add('-SkipLikelyAlreadyRenamed') | Out-Null
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

    $candidate = Join-Path $Directory ($BaseName + $Extension)

    if (($candidate -ieq $CurrentFullName) -and -not $ReservedPaths.Contains($candidate)) {
        $null = $ReservedPaths.Add($candidate)
        return $candidate
    }

    $counter = 1
    while ((Test-Path -LiteralPath $candidate) -or $ReservedPaths.Contains($candidate)) {
        $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $BaseName, $counter, $Extension)
        $counter++
    }

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

    $wordApp = $null
    $document = $null

    try {
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        $wordApp.DisplayAlerts = 0

        $document = $wordApp.Documents.Open($FilePath, $false, $true, $false)
        if (-not $document) { return $null }

        $title = $null
        try {
            $title = [string]$document.BuiltInDocumentProperties('Title').Value
        }
        catch {
            $title = $null
        }

        if (-not [string]::IsNullOrWhiteSpace($title)) {
            return $title.Trim()
        }
    }
    catch {
        return $null
    }
    finally {
        if ($document) {
            try { $document.Close($false) | Out-Null } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
        if ($wordApp) {
            try { $wordApp.Quit() | Out-Null } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    return $null
}

function Get-EmbeddedXmpTitle {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if (-not $bytes -or $bytes.Length -eq 0) { return $null }

        $text = [System.Text.Encoding]::UTF8.GetString($bytes)
        $startToken = '<x:xmpmeta'
        $endToken = '</x:xmpmeta>'

        $start = $text.IndexOf($startToken, [System.StringComparison]::OrdinalIgnoreCase)
        if ($start -lt 0) { return $null }

        $end = $text.IndexOf($endToken, $start, [System.StringComparison]::OrdinalIgnoreCase)
        if ($end -lt 0) { return $null }

        $xmpXml = $text.Substring($start, ($end - $start) + $endToken.Length)
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
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if (-not $bytes -or $bytes.Length -eq 0) { return $null }

        $pdfText = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
        $titleMatches = [regex]::Matches($pdfText, '/Title\s*\((?<title>(?:\\.|[^\\)])+)\)')

        foreach ($match in $titleMatches) {
            $rawTitle = $match.Groups['title'].Value
            if (-not [string]::IsNullOrWhiteSpace($rawTitle)) {
                $decoded = Unescape-PdfLiteralString -Value $rawTitle
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
        [Parameter(Mandatory)]
        [System.__ComObject]$ShellFolder,

        [Parameter(Mandatory)]
        [System.__ComObject]$ShellItem
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
        [Parameter(Mandatory)]
        [System.__ComObject]$ShellFolder,

        [Parameter(Mandatory)]
        [System.__ComObject]$ShellItem
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

function Get-ShellMetadataValue {
    param(
        [Parameter(Mandatory)]
        [System.__ComObject]$ShellFolder,

        [Parameter(Mandatory)]
        [System.__ComObject]$ShellItem,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames
    )

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
        [Parameter(Mandatory)]
        [System.__ComObject]$ShellFolder,

        [Parameter(Mandatory)]
        [System.__ComObject]$ShellItem,

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

        [Parameter(Mandatory)]
        [System.__ComObject]$ShellFolder,

        [Parameter(Mandatory)]
        [System.__ComObject]$ShellItem,

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

    if (-not $primaryResult -and $extension -eq '.rtf') {
        $rtfTitle = Get-RtfTitle -FilePath $File.FullName
        if ($rtfTitle) {
            $result = New-MetadataNameResult -Name $rtfTitle -Source 'RTF title control word'
            if ($result) { $primaryResult = $result }
        }
    }

    if (-not $primaryResult -and $extension -in $wordComExtensions) {
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

    $shellFolder = $Shell.Namespace($File.DirectoryName)
    if (-not $shellFolder) {
        return $null
    }

    $shellItem = $shellFolder.ParseName($File.Name)
    if (-not $shellItem) {
        return $null
    }

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

$shouldRecurse = -not $NoRecurse.IsPresent
$runMode = if ($isApplyMode) { 'Apply (rename)' } else { 'Preview only' }
if ($isImmediateRenameMode) {
    $runMode = 'Apply (rename immediate)'
}
if ($InteractiveSubfolders.IsPresent) {
    $runMode = $runMode + ' + InteractiveSubfolders'
}
if ($QuickPreview.IsPresent) {
    $runMode = if ($isApplyMode) {
        "Apply (rename) + QuickPreview($QuickPreviewLimit)"
    }
    else {
        "Preview only + QuickPreview($QuickPreviewLimit)"
    }

    if ($isImmediateRenameMode) {
        $runMode = "Apply (rename immediate) + QuickPreview($QuickPreviewLimit)"
    }

    if ($InteractiveSubfolders.IsPresent) {
        $runMode = $runMode + ' + InteractiveSubfolders'
    }
}

$startupSettings = [ordered]@{
    Apply = $isApplyMode
    ImmediateRename = $isImmediateRenameMode
    QuickPreview = [bool]$QuickPreview
    QuickPreviewLimit = if ($QuickPreview.IsPresent) { $QuickPreviewLimit } else { 'N/A' }
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
    IncludeOtherFiles = $includeOtherFilesNow
    ScanOtherFilesAtEnd = $scanOtherFilesAtEndEnabled
    ListUncommonFiles = [bool]$ListUncommonFiles
    NoAutoResumePrompt = [bool]$NoAutoResumePrompt
    SkipLikelyAlreadyRenamed = [bool]$SkipLikelyAlreadyRenamed
    PerFileTimeoutSeconds = $PerFileTimeoutSeconds
    BatchSize = $BatchSize
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
if (-not $QuickPreview.IsPresent) {
    Write-Host 'Runtime controls: press P to pause, C to continue (when paused), or S to stop gracefully.' -ForegroundColor DarkYellow
}
Write-RunLog -Message ("Startup complete. Report path: {0}; CSV path: {1}" -f $ReportPath, $CsvPath)

$items = Get-DiscoveredFiles -Root $root -IncludeHidden:$IncludeHidden.IsPresent -Recurse:$shouldRecurse -InteractiveSubfolders:$InteractiveSubfolders.IsPresent

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
            $files = Get-ChildItem -LiteralPath $folder.FullName -File -Recurse -Force:$IncludeHidden.IsPresent -ErrorAction SilentlyContinue
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
        if ($ListUncommonFiles.IsPresent -and $deferredItems.Count -gt 0) {
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

if ($QuickPreview.IsPresent) {
    $items = $items | Select-Object -First $QuickPreviewLimit
    Write-Host ("Quick preview enabled: evaluating first {0} files." -f $QuickPreviewLimit) -ForegroundColor Yellow
    Write-RunLog -Message ("QuickPreview enabled. Limit={0}" -f $QuickPreviewLimit)
}

$discoveredCount = @($items).Count
if (-not $QuickPreview.IsPresent -and $discoveredCount -ge $script:LongRunThreshold) {
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
        Write-RunLog -Message ("User canceled long scan before processing. DiscoveredFiles={0}; ApplyMode={1}" -f $discoveredCount, $Apply.IsPresent)
        Write-EndOfLogCode -Status 'Run canceled by user before long scan'
        return
    }
}
$shell = New-Object -ComObject Shell.Application

$reservedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$plan = [System.Collections.Generic.List[object]]::new()

$renamed = 0
$skipped = 0
$resumeSkipped = 0
$heuristicSkipped = 0
$timeoutSkipped = 0
$artifactSkipped = 0
$errorCount = 0
$heartbeatEvery = 100
$scanProcessed = 0
$renameProcessed = 0
$scanHeartbeatStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$renameHeartbeatStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$processingErrorFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$renameErrorEntries = [System.Collections.Generic.List[object]]::new()

$itemsArray = @($items)
$totalItems = $itemsArray.Count
$totalItemBatches = if ($totalItems -gt 0) { [int][Math]::Ceiling($totalItems / [double]$BatchSize) } else { 0 }

for ($batchStart = 0; $batchStart -lt $totalItems; $batchStart += $BatchSize) {
    $batchEnd = [Math]::Min($batchStart + $BatchSize - 1, $totalItems - 1)
    $currentBatchNumber = [int]([Math]::Floor($batchStart / [double]$BatchSize) + 1)
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
            if (-not (Handle-RunControl -Phase $activePhase)) {
                break
            }

            $scanProcessed++
            if (($scanProcessed % $heartbeatEvery) -eq 0) {
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

            if ($script:ResumeActive -and ($script:ResumeRenamedFrom.Contains($file.FullName) -or $script:ResumeRenamedTo.Contains($file.FullName))) {
                $resumeSkipped++
                continue
            }

            if (Test-IsRunArtifactFile -File $file -CurrentRunLogPath $script:LogPath -CurrentReportPath $ReportPath -CurrentCsvPath $CsvPath -ResumeLogPath $script:ResumeLogSource) {
                $artifactSkipped++
                Write-SkippedLiveLine -FileName $file.Name -Reason 'Script run/report artifact'
                Write-RunLog -Message ("ArtifactSkip Run/report artifact excluded from rename candidates: {0}" -f $file.FullName)
                continue
            }

            if ($SkipLikelyAlreadyRenamed.IsPresent -and (Test-LikelyAlreadyRenamedName -FileName $file.Name)) {
                $heuristicSkipped++
                Write-SkippedLiveLine -FileName $file.Name -Reason 'Likely already renamed (heuristic)'
                Write-RunLog -Message ("HeuristicSkip Name appears already-renamed: {0}" -f $file.FullName)
                continue
            }

            $fileStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $metadataName = Get-RenameBaseNameFromMetadata -File $file -Shell $shell -Force:$Force.IsPresent -OnlyGeneric:$OnlyGeneric.IsPresent
            $fileStopwatch.Stop()

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
                continue
            }


            # Skip files that already have _renamed before the extension
            if ($file.BaseName -match '_renamed$') {
                $skipped++
                Write-SkippedLiveLine -FileName $file.Name -Reason 'Already renamed (_renamed marker)'
                Write-RunLog -Message ("RenamedSkip Already has _renamed marker: {0}" -f $file.FullName)
                continue
            }

            # Append _renamed before the extension for the new name
            $safeBaseName = Get-SafeFileName -Name $metadataName.Name
            if (-not $safeBaseName) {
                $skipped++
                continue
            }

            $renamedBase = $safeBaseName + '_renamed'
            $targetPath = Get-UniqueTargetPath -Directory $file.DirectoryName -BaseName $renamedBase -Extension $file.Extension -CurrentFullName $file.FullName -ReservedPaths $reservedPaths

            if ($targetPath -ieq $file.FullName) {
                $skipped++
                continue
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
            if ($isImmediateRenameMode) {
                $immediateTargetPath = Join-Path -Path (Split-Path -Path $lastAdded.FullPath -Parent) -ChildPath $lastAdded.NewName
                if ($script:ResumeActive -and ($script:ResumeRenamedFrom.Contains($lastAdded.FullPath) -or $script:ResumeRenamedTo.Contains($immediateTargetPath))) {
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
            $errorMessage = if ($_.Exception -and -not [string]::IsNullOrWhiteSpace($_.Exception.Message)) { $_.Exception.Message } else { $_.ToString() }
            Write-ErrorLiveLine -FileName $file.Name -Reason 'Processing error (see run log)'
            $null = $processingErrorFiles.Add($file.FullName)
            $errorMessageFlat = (($errorMessage -replace '\|', '/') -replace "`r?`n", ' ').Trim()
            Write-RunLog -Message ("ERROR | Phase=processing | File={0} | NewName= | Message={1}" -f $file.FullName, $errorMessageFlat)
            Write-RunLog -Message ("ErrorSkip File processing failed for {0}. Error={1}" -f $file.FullName, $errorMessage)

            if ($errorMessage -match '(?i)com object|rpc server|disconnected|invalid') {
                try {
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
                $metadataName = Get-RenameBaseNameFromMetadata -File $retryFile -Shell $shell -Force:$Force.IsPresent -OnlyGeneric:$OnlyGeneric.IsPresent
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
$reportKind = if ($isImmediateRenameMode) { 'Rename' } elseif ($isApplyMode) { 'Plan' } else { 'Preview' }
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

    $applyPrompt = if ($QuickPreview.IsPresent) {
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

if (-not $QuickPreview.IsPresent -and -not $isImmediateRenameMode -and $plan.Count -ge $script:LongRunThreshold) {
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
    $totalRenameBatches = if ($totalPlanEntries -gt 0) { [int][Math]::Ceiling($totalPlanEntries / [double]$BatchSize) } else { 0 }

    for ($renameBatchStart = 0; $renameBatchStart -lt $totalPlanEntries; $renameBatchStart += $BatchSize) {
        $renameBatchEnd = [Math]::Min($renameBatchStart + $BatchSize - 1, $totalPlanEntries - 1)
        $currentRenameBatch = [int]([Math]::Floor($renameBatchStart / [double]$BatchSize) + 1)
        $currentRenameBatchSize = ($renameBatchEnd - $renameBatchStart + 1)

        if ($totalRenameBatches -gt 1) {
            Write-Host (
                "Applying batch {0}/{1} ({2} renames)..." -f
                $currentRenameBatch,
                $totalItems = $itemsArray.Count
                $totalItemBatches = if ($totalItems -gt 0) { [int][Math]::Ceiling($totalItems / [double]$BatchSize) } else { 0 }

                # Global try/catch and batch watchdog
                $globalBatchStartTime = Get-Date
                $globalBatchMemoryStart = [System.GC]::GetTotalMemory($false)
                try {
                    for ($batchStart = 0; $batchStart -lt $totalItems; $batchStart += $BatchSize) {
                        $batchEnd = [Math]::Min($batchStart + $BatchSize - 1, $totalItems - 1)
                        $batchStartTime = Get-Date
                        $batchMemoryStart = [System.GC]::GetTotalMemory($false)
                        $currentBatchNumber = [int]([Math]::Floor($batchStart / [double]$BatchSize) + 1)
                        $currentBatchCount = ($batchEnd - $batchStart + 1)

                        if ($totalItemBatches -gt 1) {
                            Write-Host (
                                "Processing batch {0}/{1} ({2} files)..." -f
                                $currentBatchNumber,
                                $totalItemBatches,
                                $currentBatchCount
                            ) -ForegroundColor DarkYellow
                            Write-RunLog -Message (
                                "BatchStart | Phase=scan | Batch={0}/{1} | Size={2} | MemoryStart={3:N0} bytes | TimeStart={4}" -f
                                $currentBatchNumber,
                                $totalItemBatches,
                                $currentBatchCount,
                                $batchMemoryStart,
                                $batchStartTime
                            )
                        }

                        for ($i = $batchStart; $i -le $batchEnd; $i++) {
                            $file = $itemsArray[$i]
                            $fileStopwatch = $null
                            try {
                                $activePhase = if ($isImmediateRenameMode) { 'rename' } elseif ($isApplyMode) { 'planning' } else { 'preview' }
                                if (-not (Handle-RunControl -Phase $activePhase)) {
                                    break
                                }

                                $scanProcessed++
                                if (($scanProcessed % $heartbeatEvery) -eq 0) {
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

                                if ($script:ResumeActive -and ($script:ResumeRenamedFrom.Contains($file.FullName) -or $script:ResumeRenamedTo.Contains($file.FullName))) {
                                    $resumeSkipped++
                                    continue
                                }
                                # ...existing file processing logic...
                            } catch {
                                # ...existing error handling...
                            }
                        }

                        $batchEndTime = Get-Date
                        $batchMemoryEnd = [System.GC]::GetTotalMemory($false)
                        $batchDuration = ($batchEndTime - $batchStartTime).TotalSeconds
                        $batchMemoryUsed = $batchMemoryEnd - $batchMemoryStart

                        Write-RunLog -Message (
                            "BatchEnd | Phase=scan | Batch={0}/{1} | Duration={2:N1}s | MemoryUsed={3:N0} bytes | TimeEnd={4}" -f
                            $currentBatchNumber,
                            $totalItemBatches,
                            $batchDuration,
                            $batchMemoryUsed,
                            $batchEndTime
                        )

                        if ($script:StopRequested) {
                            break
                        }

                        if ($totalItemBatches -gt 1) {
                            Write-RunLog -Message (
                                "BatchComplete | Phase=scan | Batch={0}/{1}" -f
                                $currentBatchNumber,
                                $totalItemBatches
                            )
                        }
                    }
                } catch {
                    $globalBatchEndTime = Get-Date
                    $globalBatchMemoryEnd = [System.GC]::GetTotalMemory($false)
                    $globalBatchDuration = ($globalBatchEndTime - $globalBatchStartTime).TotalSeconds
                    $globalBatchMemoryUsed = $globalBatchMemoryEnd - $globalBatchMemoryStart
                    Write-RunLog -Message (
                        "GLOBAL_EXCEPTION | Phase=scan | Duration={0:N1}s | MemoryUsed={1:N0} bytes | TimeEnd={2} | Error={3}" -f
                        $globalBatchDuration,
                        $globalBatchMemoryUsed,
                        $globalBatchEndTime,
                        $_.Exception.Message
                    )
                    Write-Host "[FATAL] Unhandled exception in batch processing: $($_.Exception.Message)" -ForegroundColor Red
                    Write-EndOfLogCode -Status 'Abnormal termination (unhandled exception in batch)'
                    return
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
Write-EndOfLogCode -Status 'Run completed'
