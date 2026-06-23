# =========================================================
# Script: Export-M365Users.ps1
# Purpose: Connect to Microsoft Graph and export Microsoft 365
#          user details, assigned licenses, and group memberships
#          to CSV, LOG, and HTML summary.
# Author: Moustafa Obari
# Project: Microsoft 365 User Lifecycle Management and Administration
# Recommended Runtime: PowerShell 7
# =========================================================

param(
    [switch]$EnabledOnly,
    [switch]$LicensedOnly,
    [string]$DepartmentFilter,
    [int]$KeepLatestArchives = 7
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
$ProjectRoot   = Split-Path -Parent $PSScriptRoot
$ReportFolder  = Join-Path $ProjectRoot "reports"
$LogFolder     = Join-Path $ProjectRoot "logs"
$SummaryFolder = Join-Path $ProjectRoot "summary"

$null = New-Item -ItemType Directory -Path $ReportFolder  -Force
$null = New-Item -ItemType Directory -Path $LogFolder     -Force
$null = New-Item -ItemType Directory -Path $SummaryFolder -Force

$TimeStamp = Get-Date -Format "yyyy-MM-dd__hh-mm-sstt"

$CsvPath          = Join-Path $ReportFolder  "M365-UserReport-$TimeStamp.csv"
$LogPath          = Join-Path $LogFolder     "M365-UserReport-$TimeStamp.log"
$HtmlSummaryPath  = Join-Path $SummaryFolder "M365-UserSummary-$TimeStamp.html"

$LatestCsvPath    = Join-Path $ReportFolder  "M365-UserReport-Latest.csv"
$LatestLogPath    = Join-Path $LogFolder     "M365-UserReport-Latest.log"
$LatestHtmlPath   = Join-Path $SummaryFolder "M365-UserSummary-Latest.html"

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# -----------------------------
# Counters / trackers
# -----------------------------
$successCount = 0
$warningCount = 0
$failCount = 0
$archiveRemovedCount = 0

$results = @()
$filterImpactLines = @()

$tenantDisplayName = "Unknown"
$tenantId = "Unknown"

$totalUsersRetrieved = 0
$usersExportedAfterFilters = 0

# -----------------------------
# Logging
# -----------------------------
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"
    $entry = "$timestamp [$Level] $Message"

    Write-Host $entry
    Add-Content -Path $LogPath -Value $entry
}

# -----------------------------
# Helpers
# -----------------------------
function Normalize-Text {
    param([object]$Value)

    if ($null -eq $Value) { return "" }

    $text = [string]$Value
    $text = $text -replace '\s+', ' '
    return $text.Trim()
}

function Join-Values {
    param([object[]]$Values)

    if ($null -eq $Values) { return "" }

    $clean = @()
    foreach ($value in $Values) {
        $normalized = Normalize-Text $value
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            $clean += $normalized
        }
    }

    return ($clean -join "; ")
}

function Escape-Html {
    param([string]$Text)

    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-SafeDisplayName {
    param([object]$User)

    $name = Normalize-Text $User.DisplayName
    if ([string]::IsNullOrWhiteSpace($name)) {
        return Normalize-Text $User.UserPrincipalName
    }
    return $name
}

function Cleanup-OldArchives {
    param(
        [string]$FolderPath,
        [string]$Prefix,
        [string]$Extension,
        [int]$KeepCount
    )

    $files = Get-ChildItem -Path $FolderPath -File |
        Where-Object {
            $_.Name -like "$Prefix-*" -and
            $_.Name -notlike "*-Latest$Extension" -and
            $_.Extension -eq $Extension
        } |
        Sort-Object LastWriteTime -Descending

    if (@($files).Count -gt $KeepCount) {
        $filesToRemove = $files | Select-Object -Skip $KeepCount
        foreach ($file in $filesToRemove) {
            Remove-Item -Path $file.FullName -Force
            $script:archiveRemovedCount++
            Write-Log "Removed old archive file: $($file.Name)"
        }
    }
}

function Write-LatestCopy {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$Description
    )

    Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
    Write-Log "Latest $Description copy updated: $DestinationPath"
}

# -----------------------------
# HTML helpers
# -----------------------------
function New-StatCard {
    param(
        [string]$Label,
        [string]$Value,
        [string]$ExtraClass = ""
    )

@"
<div class="stat-card $ExtraClass">
    <div class="stat-label">$(Escape-Html $Label)</div>
    <div class="stat-value">$(Escape-Html $Value)</div>
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
        "<li>$(Escape-Html $item)</li>"
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
    <div class="path-label">$(Escape-Html $Label)</div>
    <div class="path-value">$(Escape-Html $PathValue)</div>
</div>
"@
}

function Get-ExportUserCardHtml {
    param([pscustomobject]$Row)

    $cardTitle = Escape-Html $Row.DisplayName
    $upn       = Escape-Html $Row.UserPrincipalName
    $dept      = if ([string]::IsNullOrWhiteSpace($Row.Department)) { "Not set" } else { Escape-Html $Row.Department }
    $title     = if ([string]::IsNullOrWhiteSpace($Row.JobTitle)) { "Not set" } else { Escape-Html $Row.JobTitle }
    $enabled   = if ($Row.AccountEnabled) { "Enabled" } else { "Disabled" }
    $licenses  = if ([string]::IsNullOrWhiteSpace($Row.AssignedLicenses)) { "None" } else { Escape-Html $Row.AssignedLicenses }
    $groups    = if ([string]::IsNullOrWhiteSpace($Row.GroupMembership)) { "None" } else { Escape-Html $Row.GroupMembership }

    $badgeClass = if ($Row.AccountEnabled) { "badge-success" } else { "badge-neutral" }

@"
<div class="user-card">
    <div class="user-card-header">
        <div>
            <div class="user-card-name">$cardTitle</div>
            <div class="user-card-upn">$upn</div>
        </div>
        <span class="badge $badgeClass">$enabled</span>
    </div>
    <div class="user-card-line"><strong>Department:</strong> $dept</div>
    <div class="user-card-line"><strong>Job Title:</strong> $title</div>
    <div class="user-card-line"><strong>Licenses:</strong> $licenses</div>
    <div class="user-card-line"><strong>Groups:</strong> $groups</div>
</div>
"@
}

function New-UserSectionHtml {
    param(
        [string]$Title,
        [object[]]$Rows
    )

    $isEmpty = (-not $Rows -or $Rows.Count -eq 0)
    $sectionClass = if ($isEmpty) { "wide-panel user-section compact-empty" } else { "wide-panel user-section" }

    $content = if ($isEmpty) {
        '<div class="empty-state">No users were recorded in this section for this run.</div>'
    }
    else {
        ($Rows | ForEach-Object { Get-ExportUserCardHtml -Row $_ }) -join [Environment]::NewLine
    }

@"
<div class="$sectionClass">
    <h2>$(Escape-Html $Title)</h2>
    <div class="user-card-grid">
        $content
    </div>
</div>
"@
}

function New-ProcessedRowsTableHtml {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -eq 0) {
        return '<div class="empty-state">No processed users to display.</div>'
    }

    $body = foreach ($row in $Rows) {
@"
<tr>
    <td>$(Escape-Html $row.DisplayName)</td>
    <td>$(Escape-Html $row.UserPrincipalName)</td>
    <td>$(Escape-Html $(if ($row.AccountEnabled) { "Enabled" } else { "Disabled" }))</td>
    <td>$(Escape-Html $row.Department)</td>
    <td>$(Escape-Html $row.JobTitle)</td>
    <td>$(Escape-Html ([string]$row.LicenseCount))</td>
    <td>$(Escape-Html $row.AssignedLicenses)</td>
    <td>$(Escape-Html ([string]$row.GroupCount))</td>
    <td>$(Escape-Html $row.GroupMembership)</td>
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
                <th>Status</th>
                <th>Department</th>
                <th>Job Title</th>
                <th>License Count</th>
                <th>Assigned Licenses</th>
                <th>Group Count</th>
                <th>Group Membership</th>
            </tr>
        </thead>
        <tbody>
            $($body -join [Environment]::NewLine)
        </tbody>
    </table>
</div>
"@
}

# -----------------------------
# HTML Summary Builder
# -----------------------------
function New-HtmlSummary {
    param(
        [string]$Path
    )

    $generatedTime = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"

    $enabledExportedCount    = @($results | Where-Object { $_.AccountEnabled }).Count
    $disabledExportedCount   = @($results | Where-Object { -not $_.AccountEnabled }).Count
    $licensedExportedCount   = @($results | Where-Object { $_.LicenseCount -gt 0 }).Count
    $unlicensedExportedCount = @($results | Where-Object { $_.LicenseCount -eq 0 }).Count

    $runSettingsList = @(
        "EnabledOnly: $($EnabledOnly.IsPresent)"
        "LicensedOnly: $($LicensedOnly.IsPresent)"
        "DepartmentFilter: $(if ([string]::IsNullOrWhiteSpace($DepartmentFilter)) { 'Not used' } else { $DepartmentFilter })"
        "Archive Retention Count: $KeepLatestArchives"
        "Generated: $generatedTime"
    )

    $filterImpactHtml = New-ListItemsHtml -Items $(if ($filterImpactLines.Count -eq 0) {
        @("No filters reduced the result set.")
    } else {
        $filterImpactLines
    })

    $exportDefinitionsHtml = New-ListItemsHtml -Items @(
        "Total Users Retrieved: The total number of Microsoft 365 users returned before any export filters were applied."
        "Users Exported After Filters: The number of users written to the report after EnabledOnly, LicensedOnly, and DepartmentFilter logic was applied."
        "Enabled Users: Exported users whose sign-in is currently enabled."
        "Licensed Users: Exported users who currently have one or more assigned Microsoft 365 licenses."
        "Warnings: Non-fatal issues encountered during export, such as partial lookup problems."
        "Failed: Users or operations that could not be processed because of a hard error."
    )

    $runMetadataHtml = New-ListItemsHtml -Items @(
        "Tenant Name: $tenantDisplayName"
        "Tenant ID: $tenantId"
        "Rows recorded in result CSV: $($results.Count)"
        "Archive files removed in this run: $archiveRemovedCount"
        "Run duration (seconds): $([math]::Round($Stopwatch.Elapsed.TotalSeconds, 2))"
    )

    $outputFilesHtml = @(
        New-PathBlockHtml -Label "Timestamped Report" -PathValue $CsvPath
        New-PathBlockHtml -Label "Timestamped Log" -PathValue $LogPath
        New-PathBlockHtml -Label "Latest Report" -PathValue $LatestCsvPath
        New-PathBlockHtml -Label "Latest Log" -PathValue $LatestLogPath
        New-PathBlockHtml -Label "Timestamped HTML Summary" -PathValue $HtmlSummaryPath
        New-PathBlockHtml -Label "Latest HTML Summary" -PathValue $LatestHtmlPath
    ) -join [Environment]::NewLine

    $successClass = if ($successCount -gt 0) { "accent-success" } else { "" }
    $warningClass = if ($warningCount -gt 0) { "accent-warning" } else { "" }
    $failClass    = if ($failCount -gt 0)    { "accent-danger" } else { "" }

    $statCards = @(
        New-StatCard -Label "Tenant Name" -Value $tenantDisplayName
        New-StatCard -Label "Tenant ID" -Value $tenantId
        New-StatCard -Label "Total Users Retrieved" -Value ([string]$totalUsersRetrieved)
        New-StatCard -Label "Users Exported After Filters" -Value ([string]$usersExportedAfterFilters)
        New-StatCard -Label "Enabled Users Exported" -Value ([string]$enabledExportedCount) -ExtraClass "accent-info"
        New-StatCard -Label "Licensed Users Exported" -Value ([string]$licensedExportedCount) -ExtraClass $successClass
        New-StatCard -Label "Warnings" -Value ([string]$warningCount) -ExtraClass $warningClass
        New-StatCard -Label "Failed" -Value ([string]$failCount) -ExtraClass $failClass
        New-StatCard -Label "Archive Files Removed" -Value ([string]$archiveRemovedCount)
        New-StatCard -Label "Run Duration (Seconds)" -Value ([string]([math]::Round($Stopwatch.Elapsed.TotalSeconds, 2)))
    ) -join [Environment]::NewLine

    $enabledRows    = @($results | Where-Object { $_.AccountEnabled })
    $licensedRows   = @($results | Where-Object { $_.LicenseCount -gt 0 })
    $disabledRows   = @($results | Where-Object { -not $_.AccountEnabled })
    $unlicensedRows = @($results | Where-Object { $_.LicenseCount -eq 0 })

    $enabledSectionHtml    = New-UserSectionHtml -Title "Enabled Users" -Rows $enabledRows
    $licensedSectionHtml   = New-UserSectionHtml -Title "Licensed Users" -Rows $licensedRows
    $disabledSectionHtml   = New-UserSectionHtml -Title "Disabled Users" -Rows $disabledRows
    $unlicensedSectionHtml = New-UserSectionHtml -Title "Unlicensed Users" -Rows $unlicensedRows

    $processedUsersTableHtml = New-ProcessedRowsTableHtml -Rows $results
    $headline = "Execution summary for Microsoft 365 user export, filter analysis, and administrative review."

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 User Export Summary</title>
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
            --gap-lg: 14px;
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

        .container {
            max-width: 1440px;
            margin: 0 auto;
        }

        .hero {
            position: relative;
            overflow: hidden;
            background: linear-gradient(135deg, #184a77 0%, #2557c8 55%, #2f67ea 100%);
            color: #ffffff;
            border-radius: 24px;
            padding: 18px 22px 14px;
            box-shadow: var(--shadow);
            margin-bottom: var(--gap-md);
        }

        .hero::before,
        .hero::after {
            content: "";
            position: absolute;
            border-radius: 50%;
            background: rgba(255,255,255,0.07);
            pointer-events: none;
        }

        .hero::before {
            width: 110px;
            height: 110px;
            top: -14px;
            right: -8px;
        }

        .hero::after {
            width: 80px;
            height: 80px;
            top: 20px;
            right: 50px;
            background: rgba(255,255,255,0.05);
        }

        .hero h1 {
            position: relative;
            z-index: 1;
            margin: 0 0 4px;
            font-size: 22px;
            line-height: 1.15;
            font-weight: 700;
            letter-spacing: -0.35px;
        }

        .hero p {
            position: relative;
            z-index: 1;
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

        .sidebar-panel {
            position: sticky;
            top: 14px;
        }

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

        .panel li:last-child,
        .sidebar-panel li:last-child {
            margin-bottom: 0;
        }

        .panel strong,
        .wide-panel strong,
        .sidebar-panel strong {
            color: #10243d;
        }

        .path-block {
            margin-bottom: 8px;
        }

        .path-block:last-child {
            margin-bottom: 0;
        }

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
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.55);
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

        .compact-empty {
            padding-top: 10px;
            padding-bottom: 10px;
        }

        .user-card-grid {
            display: grid;
            gap: 6px;
            max-height: var(--section-scroll-height);
            overflow-y: auto;
            overflow-x: hidden;
            padding-right: 3px;
            scrollbar-width: thin;
            scrollbar-color: rgba(37,99,235,0.35) transparent;
        }

        .user-card-grid::-webkit-scrollbar {
            width: 8px;
        }

        .user-card-grid::-webkit-scrollbar-track {
            background: transparent;
        }

        .user-card-grid::-webkit-scrollbar-thumb {
            background: rgba(37,99,235,0.28);
            border-radius: 999px;
        }

        .user-card-grid::-webkit-scrollbar-thumb:hover {
            background: rgba(37,99,235,0.42);
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

        .user-card-line:last-child {
            margin-bottom: 0;
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

        .badge-success {
            background: rgba(22,163,74,0.12);
            color: #166534;
            border-color: rgba(22,163,74,0.16);
        }

        .badge-neutral {
            background: rgba(148,163,184,0.14);
            color: #475569;
            border-color: rgba(100,116,139,0.18);
        }

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
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.65);
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

        tbody tr:nth-child(even) {
            background: #fbfdff;
        }

        tbody td {
            padding: 8px 9px;
            font-size: 10.4px;
            line-height: 1.35;
            color: #23364b;
            border-bottom: 1px solid #ecf1f7;
            vertical-align: top;
        }

        tbody tr:hover {
            background: #f5f9ff;
        }

        tbody tr:last-child td {
            border-bottom: none;
        }

        .table-panel {
            grid-column: 1 / -1;
            margin-top: 0;
        }

        .footer {
            text-align: center;
            color: #8390a1;
            font-size: 10.5px;
            margin-top: 8px;
        }

        @media (max-width: 1280px) {
            .dashboard-layout {
                grid-template-columns: 1fr;
            }

            .sidebar-panel {
                position: static;
            }

            .user-section {
                grid-column: span 6;
            }
        }

        @media (max-width: 1080px) {
            .overview-grid {
                grid-template-columns: 1fr;
                gap: var(--gap-xs);
            }

            .user-section,
            .table-panel {
                grid-column: 1 / -1;
            }

            .user-sections-grid {
                gap: var(--gap-xs);
            }

            .user-card-grid {
                max-height: none;
                overflow: visible;
                padding-right: 0;
            }
        }

        @media (max-width: 900px) {
            body {
                padding: 10px;
            }

            .hero {
                padding: 16px 14px 12px;
                border-radius: 18px;
            }

            .hero h1 {
                font-size: 18px;
            }

            .hero p {
                font-size: 10.8px;
            }

            .stats-grid {
                grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
            }

            .stat-card {
                min-height: auto;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="hero">
            <h1>Microsoft 365 User Export Summary</h1>
            <p>$headline</p>
        </div>

        <div class="stats-grid">
            $statCards
        </div>

        <div class="dashboard-layout">
            <div>
                <div class="overview-grid">
                    <div class="panel">
                        <h2>Run Settings</h2>
                        <ul>
                            $(New-ListItemsHtml -Items $runSettingsList)
                        </ul>
                    </div>

                    <div class="panel">
                        <h2>Filter Impact</h2>
                        <ul>
                            $filterImpactHtml
                        </ul>
                    </div>

                    <div class="panel">
                        <h2>Export Definitions</h2>
                        <ul>
                            $exportDefinitionsHtml
                        </ul>
                    </div>

                    <div class="panel">
                        <h2>Run Metadata</h2>
                        <ul>
                            $runMetadataHtml
                        </ul>
                    </div>
                </div>
            </div>

            <div class="sidebar-panel">
                <h2>Output Files</h2>
                $outputFilesHtml
            </div>
        </div>

        <div class="user-sections-grid">
            $enabledSectionHtml
            $licensedSectionHtml
            $disabledSectionHtml
            $unlicensedSectionHtml

            <div class="wide-panel table-panel">
                <h2>Processed Users</h2>
                $processedUsersTableHtml
            </div>
        </div>

        <div class="footer">
            Generated by Export-M365Users.ps1
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $Path -Value $html -Encoding UTF8
}

# -----------------------------
# Start log
# -----------------------------
Set-Content -Path $LogPath -Value ""
Write-Log "Starting Microsoft 365 user export script."
Write-Log "PowerShell version: $($PSVersionTable.PSVersion)"
Write-Log "Report output path: $CsvPath"
Write-Log "Log output path: $LogPath"
Write-Log "EnabledOnly switch: $($EnabledOnly.IsPresent)"
Write-Log "LicensedOnly switch: $($LicensedOnly.IsPresent)"
Write-Log "DepartmentFilter: $(if ([string]::IsNullOrWhiteSpace($DepartmentFilter)) { 'Not used' } else { $DepartmentFilter })"
Write-Log "Archive retention count: $KeepLatestArchives"

# -----------------------------
# Connect to Graph
# -----------------------------
try {
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes @(
        "User.Read.All",
        "Group.Read.All",
        "Directory.Read.All",
        "Organization.Read.All"
    ) -NoWelcome | Out-Null
    Write-Log "Connected to Microsoft Graph successfully."
}
catch {
    Write-Log ("Failed to connect to Microsoft Graph. Error: {0}" -f $_.Exception.Message) "ERROR"
    throw
}

try {
    # -----------------------------
    # Tenant info
    # -----------------------------
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

        Write-Log "Tenant name: $tenantDisplayName"
        Write-Log "Tenant ID: $tenantId"
    }
    catch {
        $warningCount++
        Write-Log ("Could not retrieve tenant information. Error: {0}" -f $_.Exception.Message) "WARN"
    }

    # -----------------------------
    # Retrieve subscribed SKUs
    # -----------------------------
    Write-Log "Retrieving subscribed SKUs for license mapping..."
    $subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    $skuMap = @{}

    foreach ($sku in $subscribedSkus) {
        if ($sku.SkuId) {
            $key = Normalize-Text ([string]$sku.SkuId)
            if (-not [string]::IsNullOrWhiteSpace($key)) {
                $skuMap[$key] = Normalize-Text $sku.SkuPartNumber
            }
        }
    }

    Write-Log "Retrieved $(@($subscribedSkus).Count) subscribed SKUs."

    # -----------------------------
    # Retrieve users
    # -----------------------------
    Write-Log "Retrieving users from Microsoft 365..."
    $users = Get-MgUser -All -Property `
        "id,displayName,userPrincipalName,givenName,surname,jobTitle,department,officeLocation,city,state,country,usageLocation,accountEnabled,mobilePhone,businessPhones,createdDateTime,assignedLicenses"

    $users = @($users)
    $totalUsersRetrieved = $users.Count
    Write-Log "Retrieved $totalUsersRetrieved users."

    # -----------------------------
    # Filters
    # -----------------------------
    if ($EnabledOnly.IsPresent) {
        $before = @($users).Count
        $excluded = @($users | Where-Object { -not $_.AccountEnabled } | ForEach-Object { Get-SafeDisplayName $_ })
        $users = @($users | Where-Object { $_.AccountEnabled })
        $after = @($users).Count

        Write-Log "Applied enabled-only filter. Remaining users: $after"

        if ($after -ne $before) {
            $filterImpactLines += "Enabled-only filter reduced results from $before to $after"
            Write-Log "Enabled-only filter reduced results from $before to $after"
            if ($excluded.Count -gt 0) {
                $filterImpactLines += "Enabled-only filter excluded: $(Join-Values -Values $excluded)"
                Write-Log "Enabled-only filter excluded: $(Join-Values -Values $excluded)"
            }
        }
    }

    if ($LicensedOnly.IsPresent) {
        $before = @($users).Count
        $excluded = @(
            $users |
            Where-Object { @($_.AssignedLicenses).Count -eq 0 } |
            ForEach-Object { Get-SafeDisplayName $_ }
        )
        $users = @($users | Where-Object { @($_.AssignedLicenses).Count -gt 0 })
        $after = @($users).Count

        Write-Log "Applied licensed-only filter. Remaining users: $after"

        if ($after -ne $before) {
            $filterImpactLines += "Licensed-only filter reduced results from $before to $after"
            Write-Log "Licensed-only filter reduced results from $before to $after"
            if ($excluded.Count -gt 0) {
                $filterImpactLines += "Licensed-only filter excluded: $(Join-Values -Values $excluded)"
                Write-Log "Licensed-only filter excluded: $(Join-Values -Values $excluded)"
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($DepartmentFilter)) {
        $normalizedDepartmentFilter = Normalize-Text $DepartmentFilter
        $before = @($users).Count
        $excluded = @(
            $users |
            Where-Object { (Normalize-Text $_.Department) -ne $normalizedDepartmentFilter } |
            ForEach-Object { Get-SafeDisplayName $_ }
        )
        $users = @(
            $users |
            Where-Object { (Normalize-Text $_.Department) -eq $normalizedDepartmentFilter }
        )
        $after = @($users).Count

        Write-Log "Applied department filter '$normalizedDepartmentFilter'. Remaining users: $after"

        if ($after -ne $before) {
            $filterImpactLines += "Department filter '$normalizedDepartmentFilter' reduced results from $before to $after"
            Write-Log "Department filter '$normalizedDepartmentFilter' reduced results from $before to $after"
            if ($excluded.Count -gt 0) {
                $filterImpactLines += "Department filter '$normalizedDepartmentFilter' excluded: $(Join-Values -Values $excluded)"
                Write-Log "Department filter '$normalizedDepartmentFilter' excluded: $(Join-Values -Values $excluded)"
            }
        }
    }

    $users = @($users)
    $usersExportedAfterFilters = $users.Count

    # -----------------------------
    # Process users
    # -----------------------------
    foreach ($user in $users) {
        try {
            $safeDisplayName = Get-SafeDisplayName $user
            Write-Log "Processing user: $safeDisplayName"

            # Resolve license names
            $licenseNames = @()
            if ($user.AssignedLicenses) {
                foreach ($license in $user.AssignedLicenses) {
                    if ($license.SkuId) {
                        $skuKey = Normalize-Text ([string]$license.SkuId)
                        if ($skuMap.ContainsKey($skuKey)) {
                            $licenseNames += $skuMap[$skuKey]
                        }
                        else {
                            $licenseNames += $skuKey
                        }
                    }
                }
            }

            # Resolve group memberships
            $groupNames = @()
            $userId = Normalize-Text ([string]$user.Id)

            if (-not [string]::IsNullOrWhiteSpace($userId)) {
                try {
                    $memberOf = Get-MgUserMemberOf -UserId $userId -All -ErrorAction Stop
                    foreach ($item in @($memberOf)) {
                        if ($item.AdditionalProperties -and $item.AdditionalProperties.ContainsKey('displayName')) {
                            $groupName = Normalize-Text $item.AdditionalProperties['displayName']
                            if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                                $groupNames += $groupName
                            }
                        }
                    }
                }
                catch {
                    $warningCount++
                    Write-Log ("Could not retrieve group memberships for {0}. Error: {1}" -f $safeDisplayName, $_.Exception.Message) "WARN"
                }
            }
            else {
                $warningCount++
                Write-Log ("Could not retrieve group memberships for {0}. Error: User Id is empty." -f $safeDisplayName) "WARN"
            }

            $results += [PSCustomObject]@{
                TenantName        = $tenantDisplayName
                TenantId          = $tenantId
                DisplayName       = $safeDisplayName
                UserPrincipalName = Normalize-Text $user.UserPrincipalName
                GivenName         = Normalize-Text $user.GivenName
                Surname           = Normalize-Text $user.Surname
                Department        = Normalize-Text $user.Department
                JobTitle          = Normalize-Text $user.JobTitle
                OfficeLocation    = Normalize-Text $user.OfficeLocation
                City              = Normalize-Text $user.City
                State             = Normalize-Text $user.State
                Country           = Normalize-Text $user.Country
                UsageLocation     = Normalize-Text $user.UsageLocation
                AccountEnabled    = $user.AccountEnabled
                LicenseCount      = @($licenseNames).Count
                AssignedLicenses  = Join-Values -Values $licenseNames
                GroupCount        = @($groupNames).Count
                GroupMembership   = Join-Values -Values $groupNames
                MobilePhone       = Normalize-Text $user.MobilePhone
                BusinessPhones    = Join-Values -Values $user.BusinessPhones
                CreatedDateTime   = $user.CreatedDateTime
            }

            $successCount++
            Write-Log "Processed successfully: $safeDisplayName"
        }
        catch {
            $failCount++
            Write-Log ("Failed processing user {0}. Error: {1}" -f (Get-SafeDisplayName $user), $_.Exception.Message) "ERROR"
        }
    }

    # -----------------------------
    # Export CSV
    # -----------------------------
    $results |
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    Write-Log "CSV export completed successfully: $CsvPath"

    # -----------------------------
    # Latest copies
    # -----------------------------
    Write-LatestCopy -SourcePath $CsvPath         -DestinationPath $LatestCsvPath  -Description "CSV"
    Write-LatestCopy -SourcePath $LogPath         -DestinationPath $LatestLogPath  -Description "log"

    # -----------------------------
    # Cleanup archives before HTML latest copy
    # -----------------------------
    Cleanup-OldArchives -FolderPath $ReportFolder  -Prefix "M365-UserReport" -Extension ".csv"  -KeepCount $KeepLatestArchives
    Cleanup-OldArchives -FolderPath $LogFolder     -Prefix "M365-UserReport" -Extension ".log"  -KeepCount $KeepLatestArchives
    Cleanup-OldArchives -FolderPath $SummaryFolder -Prefix "M365-UserSummary" -Extension ".html" -KeepCount $KeepLatestArchives

    # -----------------------------
    # HTML summary
    # -----------------------------
    New-HtmlSummary -Path $HtmlSummaryPath
    Write-LatestCopy -SourcePath $HtmlSummaryPath -DestinationPath $LatestHtmlPath -Description "HTML summary"
}
catch {
    $failCount++
    Write-Log ("Fatal script error. Error: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
finally {
    try {
        Write-Log "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph | Out-Null
        Write-Log "Disconnected successfully."
    }
    catch {
        Write-Log ("Disconnect warning. Error: {0}" -f $_.Exception.Message) "WARN"
        $warningCount++
    }

    $Stopwatch.Stop()

    Write-Host ""
    Write-Host "========== SUMMARY ==========" -ForegroundColor Cyan
    Write-Host "Tenant name: $tenantDisplayName"
    Write-Host "Tenant ID: $tenantId"
    Write-Host "Total users retrieved: $totalUsersRetrieved"
    Write-Host "Users exported after filters: $usersExportedAfterFilters"

    foreach ($line in $filterImpactLines) {
        if ($line -match "reduced results from") {
            Write-Host $line
        }
    }

    Write-Host "Successfully exported: $successCount"
    Write-Host "Warnings: $warningCount"
    Write-Host "Failed: $failCount"
    Write-Host "Archive files removed: $archiveRemovedCount"
    Write-Host ("Run duration (seconds): {0}" -f [math]::Round($Stopwatch.Elapsed.TotalSeconds, 2))
    Write-Host "Timestamped report saved to: $CsvPath"
    Write-Host "Timestamped log saved to: $LogPath"
    Write-Host "Latest report saved to: $LatestCsvPath"
    Write-Host "Latest log saved to: $LatestLogPath"
    Write-Host "Timestamped HTML summary saved to: $HtmlSummaryPath"
    Write-Host "Latest HTML summary saved to: $LatestHtmlPath"
    Write-Host "=============================" -ForegroundColor Cyan
}