# ============================================
# Script: Remediate-M365LifecycleFindings.ps1
# Purpose: Apply remediation actions for Microsoft 365
#          lifecycle audit findings from the Part 7 audit CSV.
# Author: Moustafa Obari
# Project: Microsoft 365 User Lifecycle Management and Administration
# Recommended Runtime: PowerShell 7
#
# Default behavior:
# - Reads the latest Part 7 audit CSV
# - Remediates NON_COMPLIANT rows only
# - Supports -PreviewMode for safe dry runs
# - Generates CSV, log, and HTML summary output
# ============================================

param(
    [string]$AuditCsvPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "reports\M365-LifecycleAudit-Latest.csv"),
    [int]$KeepLatestArchives = 7,
    [switch]$PreviewMode
)

# -----------------------------
# Required modules
# -----------------------------
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

# -----------------------------
# Config
# -----------------------------
$ProjectRoot    = Split-Path -Parent $PSScriptRoot
$ReportFolder   = Join-Path $ProjectRoot "reports"
$LogFolder      = Join-Path $ProjectRoot "logs"
$SummaryFolder  = Join-Path $ProjectRoot "summary"

$TimeStamp = Get-Date -Format "yyyy-MM-dd__hh-mm-sstt"
$RunId     = "REMEDIATE-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmss")

$BaseName        = "M365-LifecycleRemediation"
$SummaryBaseName = "M365-LifecycleRemediationSummary"

$LogPath           = Join-Path $LogFolder "$BaseName-$TimeStamp.log"
$ResultPath        = Join-Path $ReportFolder "$BaseName-$TimeStamp.csv"
$SummaryPath       = Join-Path $SummaryFolder "$SummaryBaseName-$TimeStamp.html"

$LatestLogPath     = Join-Path $LogFolder "$BaseName-Latest.log"
$LatestResultPath  = Join-Path $ReportFolder "$BaseName-Latest.csv"
$LatestSummaryPath = Join-Path $SummaryFolder "$BaseName-Latest.html"

foreach ($folder in @($ReportFolder, $LogFolder, $SummaryFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
}

if (Test-Path $LogPath) {
    Remove-Item -Path $LogPath -Force
}
Set-Content -Path $LogPath -Value ""

# -----------------------------
# Runtime / counters
# -----------------------------
$StartTime = Get-Date

$tenantDisplayName = "Unknown"
$tenantId          = "Unknown"

$auditRowsLoaded       = 0
$rowsEligible          = 0
$actionsQueued         = 0
$remediatedCount       = 0
$previewCount          = 0
$skippedCount          = 0
$failedCount           = 0
$cleanupRemoved        = 0

$results     = @()
$auditRows   = @()
$actionQueue = @()

# -----------------------------
# Logging
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"
    $entry = "$timestamp [$Level] $Message"

    switch ($Level) {
        "INFO"  { Write-Host $entry -ForegroundColor Gray }
        "WARN"  { Write-Host $entry -ForegroundColor Yellow }
        "ERROR" { Write-Host $entry -ForegroundColor Red }
    }

    Add-Content -Path $LogPath -Value $entry
}

# -----------------------------
# Helpers
# -----------------------------
function Normalize-Text {
    param([object]$Text)

    if ($null -eq $Text) { return "" }

    $value = [string]$Text
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }

    return (($value -replace '\s+', ' ').Trim())
}

function Convert-ToHtmlSafe {
    param([object]$Value)

    if ($null -eq $Value) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Split-MultiValue {
    param([string]$Value)

    $normalized = Normalize-Text $Value
    if ([string]::IsNullOrWhiteSpace($normalized) -or $normalized -eq "None") {
        return @()
    }

    return @(
        $normalized -split ';' |
        ForEach-Object { Normalize-Text $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne "None" } |
        Sort-Object -Unique
    )
}

function Remove-OldArchives {
    param(
        [Parameter(Mandatory)] [string]$FolderPath,
        [Parameter(Mandatory)] [string]$BaseName,
        [Parameter(Mandatory)] [string]$Extension,
        [Parameter(Mandatory)] [int]$KeepCount
    )

    $pattern = "^{0}-\d{{4}}-\d{{2}}-\d{{2}}__\d{{2}}-\d{{2}}-\d{{2}}(?:AM|PM){1}$" -f [regex]::Escape($BaseName), [regex]::Escape($Extension)

    $files = Get-ChildItem -Path $FolderPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $pattern } |
        Sort-Object LastWriteTime -Descending

    if ($files.Count -gt $KeepCount) {
        $filesToDelete = $files | Select-Object -Skip $KeepCount
        foreach ($file in $filesToDelete) {
            Remove-Item -Path $file.FullName -Force
            $script:cleanupRemoved++
            Write-Log ("Removed old archive file: {0}" -f $file.Name)
        }
    }
}

function Get-SkuMaps {
    $idToPart = @{}
    $partToId = @{}

    $subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop

    foreach ($sku in $subscribedSkus) {
        $skuId      = Normalize-Text ([string]$sku.SkuId)
        $partNumber = Normalize-Text $sku.SkuPartNumber

        if (-not [string]::IsNullOrWhiteSpace($skuId) -and -not [string]::IsNullOrWhiteSpace($partNumber)) {
            $idToPart[$skuId] = $partNumber
            $partToId[$partNumber] = $skuId
        }
    }

    return @{
        IdToPart = $idToPart
        PartToId = $partToId
    }
}

function Get-GroupByDisplayName {
    param(
        [Parameter(Mandatory)]
        [string]$GroupName
    )

    $safeGroupName = $GroupName.Replace("'", "''")
    return Get-MgGroup -Filter "displayName eq '$safeGroupName'"
}

function Get-CurrentUserGroupNames {
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    $groupNames = @()

    try {
        $memberOf = Get-MgUserMemberOf -UserId $UserId -All -ErrorAction Stop
        foreach ($item in $memberOf) {
            if ($item.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                $displayName = Normalize-Text $item.AdditionalProperties.displayName
                if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                    $groupNames += $displayName
                }
            }
        }
    }
    catch {
        Write-Log ("Could not retrieve current groups for user id {0}. Error: {1}" -f $UserId, $_.Exception.Message) "WARN"
    }

    return @($groupNames | Sort-Object -Unique)
}

function Get-CurrentUserLicenseParts {
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [Parameter(Mandatory)]
        [hashtable]$SkuIdToPartMap
    )

    $currentLicenseParts = @()

    foreach ($license in $User.AssignedLicenses) {
        $skuId = Normalize-Text ([string]$license.SkuId)
        if ($SkuIdToPartMap.ContainsKey($skuId)) {
            $currentLicenseParts += $SkuIdToPartMap[$skuId]
        }
        else {
            $currentLicenseParts += $skuId
        }
    }

    return @($currentLicenseParts | Sort-Object -Unique)
}

function Add-ResultRow {
    param(
        [string]$DisplayName,
        [string]$UserPrincipalName,
        [string]$Department,
        [string]$CheckName,
        [string]$FindingCategory,
        [string]$Severity,
        [string]$ActionType,
        [string]$TargetValue,
        [string]$Status,
        [string]$Notes
    )

    $script:results += [PSCustomObject]@{
        RunId             = $RunId
        TenantName        = $tenantDisplayName
        TenantId          = $tenantId
        RemediationTime   = (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt")
        DisplayName       = $DisplayName
        UserPrincipalName = $UserPrincipalName
        Department        = $Department
        CheckName         = $CheckName
        FindingCategory   = $FindingCategory
        Severity          = $Severity
        ActionType        = $ActionType
        TargetValue       = $TargetValue
        Status            = $Status
        Notes             = $Notes
    }

    switch ($Status) {
        "REMEDIATED"                { $script:remediatedCount++ }
        "PREVIEW_ONLY"              { $script:previewCount++ }
        "SKIPPED_ALREADY_COMPLIANT" { $script:skippedCount++ }
        "SKIPPED_UNSUPPORTED"       { $script:skippedCount++ }
        "SKIPPED_NOT_FOUND"         { $script:skippedCount++ }
        "FAILED"                    { $script:failedCount++ }
    }
}

# -----------------------------
# HTML helpers
# -----------------------------
function New-StatCardHtml {
    param(
        [string]$Label,
        [string]$Value,
        [string]$AccentClass = ""
    )

@"
<div class="stat-card $AccentClass">
    <div class="stat-label">$(Convert-ToHtmlSafe $Label)</div>
    <div class="stat-value">$(Convert-ToHtmlSafe $Value)</div>
</div>
"@
}

function New-ListItemsHtml {
    param([string[]]$Items)

    $filtered = @($Items | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($filtered.Count -eq 0) {
        return "<li>No items to display.</li>"
    }

    $lines = foreach ($item in $filtered) {
        "<li>$(Convert-ToHtmlSafe $item)</li>"
    }

    return ($lines -join [Environment]::NewLine)
}

function New-PathBlockHtml {
    param(
        [string]$Label,
        [string]$PathValue
    )

@"
<div class="path-block">
    <div class="path-label">$(Convert-ToHtmlSafe $Label)</div>
    <div class="path-value">$(Convert-ToHtmlSafe $PathValue)</div>
</div>
"@
}

function Get-ResultCardHtml {
    param([pscustomobject]$Row)

    $badgeClass = switch ($Row.Status) {
        "REMEDIATED"                { "badge-success"; break }
        "PREVIEW_ONLY"              { "badge-info"; break }
        "SKIPPED_ALREADY_COMPLIANT" { "badge-neutral"; break }
        "SKIPPED_UNSUPPORTED"       { "badge-neutral"; break }
        "SKIPPED_NOT_FOUND"         { "badge-warning"; break }
        "FAILED"                    { "badge-danger"; break }
        default                     { "badge-neutral" }
    }

@"
<div class="user-card">
    <div class="user-card-header">
        <div>
            <div class="user-card-name">$(Convert-ToHtmlSafe $Row.DisplayName)</div>
            <div class="user-card-upn">$(Convert-ToHtmlSafe $Row.UserPrincipalName)</div>
        </div>
        <span class="badge $badgeClass">$(Convert-ToHtmlSafe $Row.Status)</span>
    </div>
    <div class="user-card-line"><strong>Department:</strong> $(Convert-ToHtmlSafe $Row.Department)</div>
    <div class="user-card-line"><strong>Action:</strong> $(Convert-ToHtmlSafe $Row.ActionType)</div>
    <div class="user-card-line"><strong>Target:</strong> $(Convert-ToHtmlSafe $Row.TargetValue)</div>
    <div class="user-card-line"><strong>Notes:</strong> $(Convert-ToHtmlSafe $Row.Notes)</div>
</div>
"@
}

function New-ResultSectionHtml {
    param(
        [string]$Title,
        [object[]]$Rows
    )

    $isEmpty = (-not $Rows -or $Rows.Count -eq 0)
    $sectionClass = if ($isEmpty) { "wide-panel user-section compact-empty" } else { "wide-panel user-section" }

    $content = if ($isEmpty) {
        '<div class="empty-state">No rows were recorded in this section for this run.</div>'
    }
    else {
        ($Rows | Select-Object -First 15 | ForEach-Object { Get-ResultCardHtml -Row $_ }) -join [Environment]::NewLine
    }

@"
<div class="$sectionClass">
    <h2>$([System.Net.WebUtility]::HtmlEncode($Title))</h2>
    <div class="user-card-grid">
        $content
    </div>
</div>
"@
}

function New-ProcessedRowsTableHtml {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -eq 0) {
        return '<div class="empty-state">No processed remediation rows to display.</div>'
    }

    $body = foreach ($row in $Rows) {
@"
<tr>
    <td>$(Convert-ToHtmlSafe $row.DisplayName)</td>
    <td>$(Convert-ToHtmlSafe $row.UserPrincipalName)</td>
    <td>$(Convert-ToHtmlSafe $row.ActionType)</td>
    <td>$(Convert-ToHtmlSafe $row.TargetValue)</td>
    <td>$(Convert-ToHtmlSafe $row.Status)</td>
    <td>$(Convert-ToHtmlSafe $row.Notes)</td>
</tr>
"@
    }

@"
<div class="table-wrap">
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th>UserPrincipalName</th>
                <th>Action</th>
                <th>Target</th>
                <th>Status</th>
                <th>Notes</th>
            </tr>
        </thead>
        <tbody>
            $($body -join [Environment]::NewLine)
        </tbody>
    </table>
</div>
"@
}

function New-HtmlSummary {
    param(
        [string]$OutputPath,
        [string]$AuditCsvPath,
        [int]$AuditRowsLoaded,
        [int]$RowsEligible,
        [int]$ActionsQueued,
        [int]$RemediatedCount,
        [int]$PreviewCount,
        [int]$SkippedCount,
        [int]$FailedCount,
        [int]$ArchiveFilesRemoved,
        [double]$RunDurationSeconds,
        [string]$TimestampedReportPath,
        [string]$TimestampedLogPath,
        [string]$LatestReportPath,
        [string]$LatestLogPath,
        [string]$TimestampedSummaryPath,
        [string]$LatestSummaryPath,
        [object[]]$ResultRows,
        [int]$KeepLatestArchives,
        [bool]$IsPreviewMode
    )

    $generatedAt = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"
    $headline = "This remediation run loaded $AuditRowsLoaded audit row(s), queued $ActionsQueued action(s), and completed $RemediatedCount remediation action(s)."

    $statCards = @(
        New-StatCardHtml -Label "Run ID" -Value $RunId
        New-StatCardHtml -Label "Tenant Name" -Value $tenantDisplayName
        New-StatCardHtml -Label "Tenant ID" -Value $tenantId
        New-StatCardHtml -Label "Audit Rows Loaded" -Value $AuditRowsLoaded
        New-StatCardHtml -Label "Eligible Findings" -Value $RowsEligible
        New-StatCardHtml -Label "Actions Queued" -Value $ActionsQueued
        New-StatCardHtml -Label "Remediated" -Value $RemediatedCount -AccentClass "accent-success"
        New-StatCardHtml -Label "Preview Only" -Value $PreviewCount -AccentClass "accent-info"
        New-StatCardHtml -Label "Skipped" -Value $SkippedCount
        New-StatCardHtml -Label "Failed" -Value $FailedCount -AccentClass "accent-danger"
        New-StatCardHtml -Label "Archive Files Removed" -Value $ArchiveFilesRemoved
        New-StatCardHtml -Label "Run Duration (Seconds)" -Value $RunDurationSeconds
    ) -join [Environment]::NewLine

    $runSettingsHtml = New-ListItemsHtml -Items @(
        "Audit CSV Path: $AuditCsvPath"
        "Preview Mode: $IsPreviewMode"
        "Archive Retention Count: $KeepLatestArchives"
        "Generated: $generatedAt"
    )

    $remediationNotesHtml = New-ListItemsHtml -Items @(
        "This script acts only on non-compliant Part 7 audit rows."
        "Actions are deduped so the same user/action/target combination is not applied twice in one run."
        "Preview mode records intended actions without changing Microsoft 365."
        "This script supports missing group, missing license, forbidden group, forbidden license, disabled-but-still-licensed, and disabled-but-still-in-active-groups remediation."
    )

    $runMetadataHtml = New-ListItemsHtml -Items @(
        "Run ID: $RunId"
        "Rows recorded in result CSV: $($ResultRows.Count)"
        "Retention policy keeps latest $KeepLatestArchives archive file(s) per artifact type."
        "Tenant name: $tenantDisplayName"
        "Tenant ID: $tenantId"
    )

    $outputFilesHtml = @(
        New-PathBlockHtml -Label "Timestamped Result CSV" -PathValue $TimestampedReportPath
        New-PathBlockHtml -Label "Timestamped Log" -PathValue $TimestampedLogPath
        New-PathBlockHtml -Label "Latest Result CSV" -PathValue $LatestReportPath
        New-PathBlockHtml -Label "Latest Log" -PathValue $LatestLogPath
        New-PathBlockHtml -Label "Timestamped HTML Summary" -PathValue $TimestampedSummaryPath
        New-PathBlockHtml -Label "Latest HTML Summary" -PathValue $LatestSummaryPath
    ) -join [Environment]::NewLine

    $remediatedRows = @($ResultRows | Where-Object { $_.Status -eq "REMEDIATED" })
    $previewRows    = @($ResultRows | Where-Object { $_.Status -eq "PREVIEW_ONLY" })
    $skippedRows    = @($ResultRows | Where-Object { $_.Status -like "SKIPPED*" })
    $failedRows     = @($ResultRows | Where-Object { $_.Status -eq "FAILED" })

    $remediatedSectionHtml = New-ResultSectionHtml -Title "Remediated Actions" -Rows $remediatedRows
    $previewSectionHtml    = New-ResultSectionHtml -Title "Preview Actions" -Rows $previewRows
    $skippedSectionHtml    = New-ResultSectionHtml -Title "Skipped Actions" -Rows $skippedRows
    $failedSectionHtml     = New-ResultSectionHtml -Title "Failed Actions" -Rows $failedRows

    $processedRowsTableHtml = New-ProcessedRowsTableHtml -Rows $ResultRows

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Lifecycle Remediation Summary</title>
    <style>
        :root {
            --bg-1: #f5f8fc;
            --bg-2: #edf2f8;
            --panel: rgba(255,255,255,0.97);
            --card: rgba(255,255,255,0.985);
            --text: #0f1f35;
            --muted: #66768b;
            --line: #dbe4ef;
            --blue: #2563eb;
            --info: #0f766e;
            --warning: #b45309;
            --warning-bg: rgba(245,158,11,0.10);
            --neutral: #475569;
            --neutral-bg: rgba(148,163,184,0.10);
            --danger: #dc2626;
            --danger-bg: rgba(239,68,68,0.09);
            --success: #15803d;
            --success-bg: rgba(22,163,74,0.08);
            --info-bg: rgba(15,118,110,0.08);
            --shadow: 0 22px 50px rgba(15,23,42,0.08);
            --shadow-soft: 0 8px 22px rgba(15,23,42,0.05);
            --radius-xl: 24px;
            --radius-lg: 18px;
            --radius-md: 14px;
            --radius-sm: 12px;
            --gap-xs: 8px;
            --gap-sm: 10px;
            --gap-md: 12px;
            --pad-card: 12px;
            --section-scroll-height: 395px;
        }

        * { box-sizing: border-box; }
        html, body { margin: 0; padding: 0; }

        body {
            padding: 14px;
            font-family: "Segoe UI", Tahoma, Arial, sans-serif;
            color: var(--text);
            background:
                radial-gradient(circle at top left, rgba(37,99,235,0.05), transparent 22%),
                radial-gradient(circle at top right, rgba(24,74,119,0.04), transparent 18%),
                linear-gradient(180deg, var(--bg-1) 0%, var(--bg-2) 100%);
        }

        .container { max-width: 1440px; margin: 0 auto; }

        .hero {
            overflow: hidden;
            background: linear-gradient(135deg, #184a77 0%, #2557c8 55%, #2f67ea 100%);
            color: #ffffff;
            border-radius: 24px;
            padding: 18px 22px 14px;
            box-shadow: var(--shadow);
            margin-bottom: var(--gap-md);
        }

        .hero h1 {
            margin: 0 0 4px;
            font-size: 22px;
            line-height: 1.15;
            font-weight: 700;
            letter-spacing: -0.35px;
        }

        .hero p {
            margin: 0;
            font-size: 11.8px;
            line-height: 1.45;
            color: rgba(255,255,255,0.94);
            max-width: 980px;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(132px, 1fr));
            gap: var(--gap-xs);
            margin-bottom: var(--gap-sm);
            align-items: start;
        }

        .stat-card {
            background: var(--card);
            border: 1px solid rgba(37,99,235,0.09);
            border-left: 3px solid var(--blue);
            border-radius: var(--radius-md);
            box-shadow: var(--shadow-soft);
            padding: 10px 10px 9px;
            min-height: 72px;
        }

        .stat-card.accent-success {
            border-left-color: var(--success);
            background: linear-gradient(180deg, var(--success-bg) 0%, rgba(255,255,255,0.98) 100%);
        }

        .stat-card.accent-danger {
            border-left-color: var(--danger);
            background: linear-gradient(180deg, var(--danger-bg) 0%, rgba(255,255,255,0.98) 100%);
        }

        .stat-card.accent-warning {
            border-left-color: var(--warning);
            background: linear-gradient(180deg, var(--warning-bg) 0%, rgba(255,255,255,0.98) 100%);
        }

        .stat-card.accent-info {
            border-left-color: var(--info);
            background: linear-gradient(180deg, var(--info-bg) 0%, rgba(255,255,255,0.98) 100%);
        }

        .stat-label {
            font-size: 9px;
            text-transform: uppercase;
            letter-spacing: 0.7px;
            color: var(--muted);
            margin-bottom: 6px;
            line-height: 1.25;
        }

        .stat-value {
            font-size: 12.5px;
            line-height: 1.35;
            font-weight: 700;
            color: var(--text);
            overflow-wrap: anywhere;
            word-break: break-word;
        }

        .dashboard-layout {
            display: grid;
            grid-template-columns: minmax(0, 1.85fr) minmax(280px, 0.9fr);
            gap: var(--gap-sm);
            align-items: start;
            margin-bottom: var(--gap-sm);
        }

        .overview-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: var(--gap-sm);
            align-items: start;
        }

        .panel,
        .wide-panel,
        .sidebar-panel {
            background: var(--panel);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-soft);
            border: 1px solid rgba(15,23,42,0.04);
            padding: var(--pad-card);
            height: auto;
            min-height: 0;
            align-self: start;
        }

        .sidebar-panel { position: sticky; top: 14px; }

        .panel h2,
        .wide-panel h2,
        .sidebar-panel h2 {
            margin: 0 0 8px;
            font-size: 13px;
            line-height: 1.25;
            font-weight: 700;
            color: #12233d;
        }

        .panel ul,
        .sidebar-panel ul {
            margin: 0;
            padding-left: 16px;
        }

        .panel li,
        .sidebar-panel li {
            margin-bottom: 7px;
            line-height: 1.45;
            color: #24364c;
            font-size: 11.2px;
        }

        .path-block { margin-bottom: 8px; }

        .path-label {
            display: block;
            margin-bottom: 4px;
            font-size: 11px;
            font-weight: 700;
            color: #12233d;
        }

        .path-value {
            display: block;
            color: #4e6278;
            font-size: 10.2px;
            line-height: 1.45;
            font-family: "Cascadia Code", Consolas, "Courier New", monospace;
            background: linear-gradient(180deg, rgba(37,99,235,0.045) 0%, rgba(255,255,255,0.9) 100%);
            border: 1px solid rgba(37,99,235,0.08);
            border-radius: 10px;
            padding: 8px 9px;
            overflow-wrap: anywhere;
            word-break: break-word;
            white-space: normal;
        }

        .user-sections-grid {
            display: grid;
            grid-template-columns: repeat(12, minmax(0, 1fr));
            gap: var(--gap-sm);
            grid-auto-flow: dense;
            align-items: start;
            margin-bottom: var(--gap-sm);
        }

        .user-section {
            grid-column: span 3;
            min-width: 0;
            align-self: start;
            display: flex;
            flex-direction: column;
        }

        .compact-empty { padding-top: 10px; padding-bottom: 10px; }

        .user-card-grid {
            display: grid;
            gap: 6px;
            max-height: var(--section-scroll-height);
            overflow-y: auto;
            overflow-x: hidden;
            padding-right: 3px;
        }

        .user-card {
            border: 1px solid var(--line);
            border-radius: var(--radius-sm);
            padding: 8px 9px 7px;
            background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,251,255,0.96) 100%);
        }

        .user-card-header {
            display: flex;
            align-items: start;
            justify-content: space-between;
            gap: 8px;
            margin-bottom: 6px;
        }

        .user-card-name {
            font-size: 11.8px;
            font-weight: 700;
            color: #12233d;
            margin-bottom: 2px;
        }

        .user-card-upn {
            font-size: 10px;
            color: #5f7288;
            word-break: break-word;
            line-height: 1.35;
        }

        .user-card-line {
            font-size: 10.4px;
            line-height: 1.4;
            color: #23364b;
            margin-bottom: 2px;
        }

        .badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 9px;
            font-weight: 700;
            white-space: nowrap;
            border: 1px solid transparent;
        }

        .badge-success { background: rgba(22,163,74,0.12); color: #166534; border-color: rgba(22,163,74,0.16); }
        .badge-info    { background: rgba(15,118,110,0.12); color: #0f766e; border-color: rgba(15,118,110,0.16); }
        .badge-neutral { background: rgba(148,163,184,0.14); color: #475569; border-color: rgba(100,116,139,0.18); }
        .badge-warning { background: rgba(245,158,11,0.14); color: #b45309; border-color: rgba(245,158,11,0.16); }
        .badge-danger  { background: rgba(239,68,68,0.12); color: #b91c1c; border-color: rgba(239,68,68,0.16); }

        .empty-state {
            color: #738296;
            font-size: 10.5px;
            line-height: 1.4;
            padding: 1px 0 0;
        }

        .table-wrap {
            overflow-x: auto;
            border: 1px solid var(--line);
            border-radius: var(--radius-sm);
            background: #ffffff;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: #ffffff;
        }

        thead th {
            text-align: left;
            font-size: 9.6px;
            text-transform: uppercase;
            letter-spacing: 0.45px;
            color: #5e6f83;
            background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
            padding: 8px 9px;
            border-bottom: 1px solid var(--line);
            white-space: nowrap;
        }

        tbody tr:nth-child(even) { background: #fbfdff; }

        tbody td {
            padding: 8px 9px;
            font-size: 10.4px;
            line-height: 1.35;
            color: #23364b;
            border-bottom: 1px solid #ecf1f7;
            vertical-align: top;
        }

        .table-panel { grid-column: 1 / -1; margin-top: 0; }

        .footer {
            text-align: center;
            color: #8390a1;
            font-size: 10.5px;
            margin-top: 8px;
        }

        @media (max-width: 1280px) {
            .dashboard-layout { grid-template-columns: 1fr; }
            .sidebar-panel { position: static; }
            .user-section { grid-column: span 6; }
        }

        @media (max-width: 1080px) {
            .overview-grid { grid-template-columns: 1fr; gap: var(--gap-xs); }
            .user-section, .table-panel { grid-column: 1 / -1; }
            .user-card-grid { max-height: none; overflow: visible; padding-right: 0; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="hero">
            <h1>Microsoft 365 Lifecycle Remediation Summary</h1>
            <p>$([System.Net.WebUtility]::HtmlEncode($headline))</p>
        </div>

        <div class="stats-grid">
            $statCards
        </div>

        <div class="dashboard-layout">
            <div>
                <div class="overview-grid">
                    <div class="panel">
                        <h2>Run Settings</h2>
                        <ul>$runSettingsHtml</ul>
                    </div>

                    <div class="panel">
                        <h2>Remediation Notes</h2>
                        <ul>$remediationNotesHtml</ul>
                    </div>

                    <div class="panel">
                        <h2>Run Metadata</h2>
                        <ul>$runMetadataHtml</ul>
                    </div>
                </div>
            </div>

            <div class="sidebar-panel">
                <h2>Output Files</h2>
                $outputFilesHtml
            </div>
        </div>

        <div class="user-sections-grid">
            $remediatedSectionHtml
            $previewSectionHtml
            $skippedSectionHtml
            $failedSectionHtml

            <div class="wide-panel table-panel">
                <h2>Processed Remediation Rows</h2>
                $processedRowsTableHtml
            </div>
        </div>

        <div class="footer">
            Generated by Remediate-M365LifecycleFindings.ps1
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}

# -----------------------------
# Start log
# -----------------------------
Write-Log "Starting lifecycle remediation script."
Write-Log ("Run ID: {0}" -f $RunId)
Write-Log ("Audit CSV path: {0}" -f $AuditCsvPath)
Write-Log ("Preview mode: {0}" -f [bool]$PreviewMode)
Write-Log ("Result output path: {0}" -f $ResultPath)
Write-Log ("Log output path: {0}" -f $LogPath)
Write-Log ("Summary output path: {0}" -f $SummaryPath)
Write-Log ("Archive retention count: {0}" -f $KeepLatestArchives)

if (-not (Test-Path $AuditCsvPath)) {
    throw "Audit CSV file not found: $AuditCsvPath"
}

try {
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All" -NoWelcome | Out-Null
    Write-Log "Connected to Microsoft Graph successfully."

    try {
        $context = Get-MgContext
        if ($context -and $context.TenantId) {
            $tenantId = [string]$context.TenantId
        }

        $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($org) {
            if (-not [string]::IsNullOrWhiteSpace($org.DisplayName)) {
                $tenantDisplayName = Normalize-Text $org.DisplayName
            }
            if ($org.Id) {
                $tenantId = [string]$org.Id
            }
        }

        Write-Log ("Tenant name: {0}" -f $tenantDisplayName)
        Write-Log ("Tenant ID: {0}" -f $tenantId)
    }
    catch {
        Write-Log ("Could not retrieve tenant information. Error: {0}" -f $_.Exception.Message) "WARN"
    }

    Write-Log "Importing audit CSV..."
    $auditRows = Import-Csv -Path $AuditCsvPath
    $auditRowsLoaded = @($auditRows).Count
    Write-Log ("Imported {0} audit row(s)." -f $auditRowsLoaded)

    if ($auditRowsLoaded -eq 0) {
        throw "No audit rows were loaded from the audit CSV."
    }

    Write-Log "Retrieving subscribed SKUs for remediation mapping..."
    $skuMaps = Get-SkuMaps
    $skuIdToPartMap = $skuMaps.IdToPart
    $skuPartToIdMap = $skuMaps.PartToId
    Write-Log ("Retrieved {0} subscribed SKU mapping(s)." -f $skuPartToIdMap.Count)

    $seenKeys = New-Object System.Collections.Generic.HashSet[string]

    foreach ($row in $auditRows) {
        $status = Normalize-Text $row.Status
        if ($status -eq "COMPLIANT") {
            continue
        }

        $upn             = Normalize-Text $row.UserPrincipalName
        $displayName     = Normalize-Text $row.DisplayName
        $department      = Normalize-Text $row.Department
        $checkName       = Normalize-Text $row.CheckName
        $findingCategory = Normalize-Text $row.FindingCategory
        $severity        = Normalize-Text $row.Severity

        if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($checkName)) {
            continue
        }

        $rowsEligible++

        switch ($checkName) {
            "Enabled User Missing Required Group" {
                foreach ($group in (Split-MultiValue $row.ExpectedGroups)) {
                    $key = "$upn|ADD_GROUP|$group"
                    if ($seenKeys.Add($key)) {
                        $actionQueue += [PSCustomObject]@{
                            DisplayName       = $displayName
                            UserPrincipalName = $upn
                            Department        = $department
                            CheckName         = $checkName
                            FindingCategory   = $findingCategory
                            Severity          = $severity
                            ActionType        = "ADD_GROUP"
                            TargetValue       = $group
                        }
                    }
                }
            }

            "Enabled User Missing Required License" {
                foreach ($license in (Split-MultiValue $row.ExpectedLicenses)) {
                    $key = "$upn|ADD_LICENSE|$license"
                    if ($seenKeys.Add($key)) {
                        $actionQueue += [PSCustomObject]@{
                            DisplayName       = $displayName
                            UserPrincipalName = $upn
                            Department        = $department
                            CheckName         = $checkName
                            FindingCategory   = $findingCategory
                            Severity          = $severity
                            ActionType        = "ADD_LICENSE"
                            TargetValue       = $license
                        }
                    }
                }
            }

            "Forbidden Group Assigned" {
                foreach ($group in (Split-MultiValue $row.ForbiddenGroups)) {
                    $key = "$upn|REMOVE_GROUP|$group"
                    if ($seenKeys.Add($key)) {
                        $actionQueue += [PSCustomObject]@{
                            DisplayName       = $displayName
                            UserPrincipalName = $upn
                            Department        = $department
                            CheckName         = $checkName
                            FindingCategory   = $findingCategory
                            Severity          = $severity
                            ActionType        = "REMOVE_GROUP"
                            TargetValue       = $group
                        }
                    }
                }
            }

            "Forbidden License Assigned" {
                foreach ($license in (Split-MultiValue $row.ForbiddenLicenses)) {
                    $key = "$upn|REMOVE_LICENSE|$license"
                    if ($seenKeys.Add($key)) {
                        $actionQueue += [PSCustomObject]@{
                            DisplayName       = $displayName
                            UserPrincipalName = $upn
                            Department        = $department
                            CheckName         = $checkName
                            FindingCategory   = $findingCategory
                            Severity          = $severity
                            ActionType        = "REMOVE_LICENSE"
                            TargetValue       = $license
                        }
                    }
                }
            }

            "Disabled But Still Licensed" {
                foreach ($license in (Split-MultiValue $row.AssignedLicenses)) {
                    $key = "$upn|REMOVE_LICENSE|$license"
                    if ($seenKeys.Add($key)) {
                        $actionQueue += [PSCustomObject]@{
                            DisplayName       = $displayName
                            UserPrincipalName = $upn
                            Department        = $department
                            CheckName         = $checkName
                            FindingCategory   = $findingCategory
                            Severity          = $severity
                            ActionType        = "REMOVE_LICENSE"
                            TargetValue       = $license
                        }
                    }
                }
            }

            "Disabled But Still In Active Groups" {
                foreach ($group in (Split-MultiValue $row.GroupMemberships)) {
                    $key = "$upn|REMOVE_GROUP|$group"
                    if ($seenKeys.Add($key)) {
                        $actionQueue += [PSCustomObject]@{
                            DisplayName       = $displayName
                            UserPrincipalName = $upn
                            Department        = $department
                            CheckName         = $checkName
                            FindingCategory   = $findingCategory
                            Severity          = $severity
                            ActionType        = "REMOVE_GROUP"
                            TargetValue       = $group
                        }
                    }
                }
            }

            default {
                $key = "$upn|SKIP_UNSUPPORTED|$checkName"
                if ($seenKeys.Add($key)) {
                    $actionQueue += [PSCustomObject]@{
                        DisplayName       = $displayName
                        UserPrincipalName = $upn
                        Department        = $department
                        CheckName         = $checkName
                        FindingCategory   = $findingCategory
                        Severity          = $severity
                        ActionType        = "UNSUPPORTED"
                        TargetValue       = ""
                    }
                }
            }
        }
    }

    $actionsQueued = @($actionQueue).Count
    Write-Log ("Queued {0} remediation action(s)." -f $actionsQueued)

    foreach ($action in $actionQueue) {
        $displayName = $action.DisplayName
        $upn         = $action.UserPrincipalName
        $department  = $action.Department
        $checkName   = $action.CheckName
        $category    = $action.FindingCategory
        $severity    = $action.Severity
        $actionType  = $action.ActionType
        $targetValue = $action.TargetValue

        Write-Log ("Processing remediation action for {0}: {1} -> {2}" -f $upn, $actionType, $targetValue)

        try {
            if ($actionType -eq "UNSUPPORTED") {
                Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_UNSUPPORTED" -Notes "This finding type is not mapped to an automated remediation action in Part 8."
                continue
            }

            $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -Property "Id,AssignedLicenses"
            if (-not $user) {
                Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_NOT_FOUND" -Notes "User was not found in Microsoft 365 during remediation."
                continue
            }

            switch ($actionType) {
                "ADD_GROUP" {
                    $group = Get-GroupByDisplayName -GroupName $targetValue
                    if (-not $group) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_NOT_FOUND" -Notes "Target group was not found."
                        continue
                    }

                    $currentGroups = Get-CurrentUserGroupNames -UserId $user.Id

                    if ($currentGroups -contains $targetValue) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_ALREADY_COMPLIANT" -Notes "User is already a member of the target group."
                        continue
                    }

                    if ($PreviewMode) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "PREVIEW_ONLY" -Notes "Would add user to target group."
                    }
                    else {
                        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id | Out-Null
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "REMEDIATED" -Notes "User added to target group successfully."
                    }
                }

                "REMOVE_GROUP" {
                    $group = Get-GroupByDisplayName -GroupName $targetValue
                    if (-not $group) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_NOT_FOUND" -Notes "Target group was not found."
                        continue
                    }

                    $currentGroups = Get-CurrentUserGroupNames -UserId $user.Id

                    if ($currentGroups -notcontains $targetValue) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_ALREADY_COMPLIANT" -Notes "User is not currently a member of the target group."
                        continue
                    }

                    if ($PreviewMode) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "PREVIEW_ONLY" -Notes "Would remove user from target group."
                    }
                    else {
                        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "REMEDIATED" -Notes "User removed from target group successfully."
                    }
                }

                "ADD_LICENSE" {
                    if (-not $skuPartToIdMap.ContainsKey($targetValue)) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_NOT_FOUND" -Notes "Target license SKU was not found in the tenant."
                        continue
                    }

                    $currentLicenseParts = Get-CurrentUserLicenseParts -User $user -SkuIdToPartMap $skuIdToPartMap

                    if ($currentLicenseParts -contains $targetValue) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_ALREADY_COMPLIANT" -Notes "User already has the target license."
                        continue
                    }

                    if ($PreviewMode) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "PREVIEW_ONLY" -Notes "Would assign target license to user."
                    }
                    else {
                        Set-MgUserLicense -UserId $user.Id -AddLicenses @(@{ SkuId = $skuPartToIdMap[$targetValue] }) -RemoveLicenses @() | Out-Null
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "REMEDIATED" -Notes "Target license assigned successfully."
                    }
                }

                "REMOVE_LICENSE" {
                    if (-not $skuPartToIdMap.ContainsKey($targetValue)) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_NOT_FOUND" -Notes "Target license SKU was not found in the tenant."
                        continue
                    }

                    $currentLicenseParts = Get-CurrentUserLicenseParts -User $user -SkuIdToPartMap $skuIdToPartMap

                    if ($currentLicenseParts -notcontains $targetValue) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_ALREADY_COMPLIANT" -Notes "User does not currently have the target license."
                        continue
                    }

                    if ($PreviewMode) {
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "PREVIEW_ONLY" -Notes "Would remove target license from user."
                    }
                    else {
                        Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($skuPartToIdMap[$targetValue]) | Out-Null
                        Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "REMEDIATED" -Notes "Target license removed successfully."
                    }
                }

                default {
                    Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "SKIPPED_UNSUPPORTED" -Notes "This action type is not supported."
                }
            }
        }
        catch {
            Add-ResultRow -DisplayName $displayName -UserPrincipalName $upn -Department $department -CheckName $checkName -FindingCategory $category -Severity $severity -ActionType $actionType -TargetValue $targetValue -Status "FAILED" -Notes $_.Exception.Message
            Write-Log ("Failed remediation action for {0}. Error: {1}" -f $upn, $_.Exception.Message) "ERROR"
        }
    }

    $results |
        Sort-Object DisplayName, ActionType, TargetValue |
        Export-Csv -Path $ResultPath -NoTypeInformation -Encoding UTF8

    Write-Log ("Lifecycle remediation result CSV exported successfully: {0}" -f $ResultPath)

    $EndTimeForHtml  = Get-Date
    $DurationForHtml = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTimeForHtml).TotalSeconds, 2)

    New-HtmlSummary `
        -OutputPath $SummaryPath `
        -AuditCsvPath $AuditCsvPath `
        -AuditRowsLoaded $auditRowsLoaded `
        -RowsEligible $rowsEligible `
        -ActionsQueued $actionsQueued `
        -RemediatedCount $remediatedCount `
        -PreviewCount $previewCount `
        -SkippedCount $skippedCount `
        -FailedCount $failedCount `
        -ArchiveFilesRemoved $cleanupRemoved `
        -RunDurationSeconds $DurationForHtml `
        -TimestampedReportPath $ResultPath `
        -TimestampedLogPath $LogPath `
        -LatestReportPath $LatestResultPath `
        -LatestLogPath $LatestLogPath `
        -TimestampedSummaryPath $SummaryPath `
        -LatestSummaryPath $LatestSummaryPath `
        -ResultRows $results `
        -KeepLatestArchives $KeepLatestArchives `
        -IsPreviewMode ([bool]$PreviewMode)

    Write-Log ("Lifecycle remediation HTML summary generated successfully: {0}" -f $SummaryPath)
}
catch {
    $failedCount++
    Write-Log ("Script failed. Error: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
finally {
    try {
        Remove-OldArchives -FolderPath $ReportFolder  -BaseName $BaseName        -Extension ".csv"  -KeepCount $KeepLatestArchives
        Remove-OldArchives -FolderPath $LogFolder     -BaseName $BaseName        -Extension ".log"  -KeepCount $KeepLatestArchives
        Remove-OldArchives -FolderPath $SummaryFolder -BaseName $SummaryBaseName -Extension ".html" -KeepCount $KeepLatestArchives
    }
    catch {
        Write-Host ("Archive cleanup warning: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        if (Test-Path $ResultPath) {
            Copy-Item -Path $ResultPath -Destination $LatestResultPath -Force -ErrorAction Stop
            Write-Log ("Latest result CSV updated: {0}" -f $LatestResultPath)
        }
        else {
            Write-Log ("Skipped latest result CSV update because the timestamped result file was not created: {0}" -f $ResultPath) "WARN"
        }
    }
    catch {
        Write-Host ("Could not create latest result copy: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        if (Test-Path $SummaryPath) {
            Copy-Item -Path $SummaryPath -Destination $LatestSummaryPath -Force -ErrorAction Stop
            Write-Log ("Latest HTML summary updated: {0}" -f $LatestSummaryPath)
        }
        else {
            Write-Log ("Skipped latest HTML summary update because the timestamped summary file was not created: {0}" -f $SummaryPath) "WARN"
        }
    }
    catch {
        Write-Host ("Could not create latest summary copy: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        Write-Log "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph | Out-Null
        Write-Log "Disconnected successfully."
    }
    catch {
        Write-Log ("Disconnect warning: {0}" -f $_.Exception.Message) "WARN"
    }

    try {
        if (Test-Path $LogPath) {
            Copy-Item -Path $LogPath -Destination $LatestLogPath -Force -ErrorAction Stop
            Write-Host ("Latest log copy updated: {0}" -f $LatestLogPath) -ForegroundColor Gray
        }
        else {
            Write-Host ("Skipped latest log copy because the timestamped log file was not created: {0}" -f $LogPath) -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host ("Could not create latest log copy: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

# -----------------------------
# Final console summary
# -----------------------------
$EndTime  = Get-Date
$Duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalSeconds, 2)

Write-Host ""
Write-Host "========== SUMMARY ==========" -ForegroundColor Cyan
Write-Host ("Run ID: {0}" -f $RunId)
Write-Host ("Audit rows loaded: {0}" -f $auditRowsLoaded)
Write-Host ("Eligible findings: {0}" -f $rowsEligible)
Write-Host ("Actions queued: {0}" -f $actionsQueued)
Write-Host ("Remediated: {0}" -f $remediatedCount)
Write-Host ("Preview only: {0}" -f $previewCount)
Write-Host ("Skipped: {0}" -f $skippedCount)
Write-Host ("Failed: {0}" -f $failedCount)
Write-Host ("Archive files removed: {0}" -f $cleanupRemoved)
Write-Host ("Run duration (seconds): {0}" -f $Duration)
Write-Host ("Timestamped result saved to: {0}" -f $ResultPath)
Write-Host ("Timestamped log saved to: {0}" -f $LogPath)
Write-Host ("Latest result saved to: {0}" -f $LatestResultPath)
Write-Host ("Latest log saved to: {0}" -f $LatestLogPath)
Write-Host ("Timestamped HTML summary saved to: {0}" -f $SummaryPath)
Write-Host ("Latest HTML summary saved to: {0}" -f $LatestSummaryPath)
Write-Host "=============================" -ForegroundColor Cyan