# ============================================
# Script: Onboarding-New-M365Users.ps1
# Purpose: Bulk create or reconcile Microsoft 365 users from CSV,
#          assign profile details, licenses, and groups.
# Author: Moustafa Obari
# Project: Microsoft 365 User Lifecycle Management and Administration
# Recommended Runtime: PowerShell 7
# ============================================

param(
    [string]$CsvPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "inputs\NewUsers.csv"),
    [string]$TenantDomain = "MOsDemoLAB.onmicrosoft.com",
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
$ProjectRoot    = Split-Path -Parent $PSScriptRoot
$ReportFolder   = Join-Path $ProjectRoot "reports"
$LogFolder      = Join-Path $ProjectRoot "logs"
$SummaryFolder  = Join-Path $ProjectRoot "summary"

$TimeStamp = Get-Date -Format "yyyy-MM-dd__hh-mm-sstt"
$RunId     = "ONBOARD-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmss")

$BaseName        = "M365-BulkOnboarding"
$SummaryBaseName = "M365-BulkOnboardingSummary"

$LogPath           = Join-Path $LogFolder "$BaseName-$TimeStamp.log"
$ResultPath        = Join-Path $ReportFolder "$BaseName-$TimeStamp.csv"
$SummaryPath       = Join-Path $SummaryFolder "$SummaryBaseName-$TimeStamp.html"

$LatestLogPath     = Join-Path $LogFolder "$BaseName-Latest.log"
$LatestResultPath  = Join-Path $ReportFolder "$BaseName-Latest.csv"
$LatestSummaryPath = Join-Path $SummaryFolder "$SummaryBaseName-Latest.html"

# -----------------------------
# Ensure folders exist
# -----------------------------
foreach ($folder in @($ReportFolder, $LogFolder, $SummaryFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
}

# -----------------------------
# Runtime / counters
# -----------------------------
$StartTime      = Get-Date
$createdCount   = 0
$updatedCount   = 0
$skippedCount   = 0
$failedCount    = 0
$warnCount      = 0
$cleanupRemoved = 0
$results        = @()
$csvUsers       = @()

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
        "WARN"  { Write-Host $entry -ForegroundColor Yellow; $script:warnCount++ }
        "ERROR" { Write-Host $entry -ForegroundColor Red }
    }

    Add-Content -Path $LogPath -Value $entry
}

# -----------------------------
# Helpers
# -----------------------------
function Normalize-Text {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    return (($Text -replace '\s+', ' ').Trim())
}

function Convert-ToHtmlSafe {
    param([string]$Value)

    if ($null -eq $Value) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
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

function Get-SkuMap {
    $map = @{}
    $subscribedSkus = Get-MgSubscribedSku

    foreach ($sku in $subscribedSkus) {
        $partNumber = Normalize-Text $sku.SkuPartNumber
        $skuId      = [string]$sku.SkuId

        if (-not [string]::IsNullOrWhiteSpace($partNumber) -and -not [string]::IsNullOrWhiteSpace($skuId)) {
            $map[$partNumber] = $skuId
        }
    }

    return $map
}

function Get-GroupByDisplayName {
    param(
        [Parameter(Mandatory)]
        [string]$GroupName
    )

    $safeGroupName = $GroupName.Replace("'", "''")
    return Get-MgGroup -Filter "displayName eq '$safeGroupName'"
}

function Add-ResultRow {
    param(
        [string]$DisplayName,
        [string]$UserPrincipalName,
        [string]$Department,
        [string]$JobTitle,
        [string]$GroupName,
        [string]$LicenseSku,
        [string]$Status,
        [string]$Notes
    )

    $script:results += [PSCustomObject]@{
        DisplayName       = $DisplayName
        UserPrincipalName = $UserPrincipalName
        Department        = $Department
        JobTitle          = $JobTitle
        GroupName         = $GroupName
        LicenseSku        = $LicenseSku
        Status            = $Status
        Notes             = $Notes
    }
}

function Get-UserGroupNames {
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    $groupNames = @()

    try {
        $memberOf = Get-MgUserMemberOf -UserId $UserId -All
        foreach ($item in $memberOf) {
            $odataType = $item.AdditionalProperties.'@odata.type'
            if ($odataType -eq '#microsoft.graph.group') {
                $displayName = Normalize-Text $item.AdditionalProperties.displayName
                if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                    $groupNames += $displayName
                }
            }
        }
    }
    catch {
        Write-Log ("Could not retrieve group memberships for user id {0}. Error: {1}" -f $UserId, $_.Exception.Message) "WARN"
    }

    return @($groupNames | Sort-Object -Unique)
}

function Get-UserLicenseSkuPartNumbers {
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [Parameter(Mandatory)]
        [hashtable]$SkuMap
    )

    $assigned = @()

    if ($null -ne $User.AssignedLicenses) {
        foreach ($license in $User.AssignedLicenses) {
            $skuId = [string]$license.SkuId
            if (-not [string]::IsNullOrWhiteSpace($skuId)) {
                $matchedSku = $SkuMap.GetEnumerator() | Where-Object { $_.Value -eq $skuId } | Select-Object -First 1
                if ($matchedSku) {
                    $assigned += $matchedSku.Key
                }
                else {
                    $assigned += $skuId
                }
            }
        }
    }

    return @($assigned | Sort-Object -Unique)
}

function Compare-Value {
    param(
        [string]$Current,
        [string]$Desired
    )

    return (Normalize-Text $Current) -ne (Normalize-Text $Desired)
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

    $safeLabel = Convert-ToHtmlSafe $Label
    $safeValue = Convert-ToHtmlSafe $Value

@"
<div class="stat-card $AccentClass">
    <div class="stat-label">$safeLabel</div>
    <div class="stat-value">$safeValue</div>
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

    $safeLabel = Convert-ToHtmlSafe $Label
    $safePath  = Convert-ToHtmlSafe $PathValue

@"
<div class="path-block">
    <div class="path-label">$safeLabel</div>
    <div class="path-value">$safePath</div>
</div>
"@
}

function New-StatusBreakdownHtml {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -eq 0) {
        return "<li>No onboarding rows recorded.</li>"
    }

    $displayOrder = @(
        "Created",
        "Updated",
        "Skipped - Already Compliant",
        "Failed"
    )

    $lines = foreach ($status in $displayOrder) {
        $matchCount = @($Rows | Where-Object { $_.Status -eq $status }).Count
        if ($matchCount -gt 0) {
            "<li><strong>$(Convert-ToHtmlSafe $status):</strong> $matchCount</li>"
        }
    }

    return ($lines -join [Environment]::NewLine)
}

function Get-UserCardHtml {
    param([pscustomobject]$Row)

    $safeName    = Convert-ToHtmlSafe $Row.DisplayName
    $safeUpn     = Convert-ToHtmlSafe $Row.UserPrincipalName
    $safeDept    = Convert-ToHtmlSafe $Row.Department
    $safeTitle   = Convert-ToHtmlSafe $Row.JobTitle
    $safeGroup   = Convert-ToHtmlSafe $Row.GroupName
    $safeLicense = Convert-ToHtmlSafe $Row.LicenseSku
    $safeStatus  = Convert-ToHtmlSafe $Row.Status
    $safeNotes   = Convert-ToHtmlSafe $Row.Notes

    $badgeClass = switch ($Row.Status) {
        "Created"                     { "badge-success"; break }
        "Updated"                     { "badge-info"; break }
        "Skipped - Already Compliant" { "badge-neutral"; break }
        "Failed"                      { "badge-danger"; break }
        default                       { "badge-neutral" }
    }

@"
<div class="user-card">
    <div class="user-card-header">
        <div>
            <div class="user-card-name">$safeName</div>
            <div class="user-card-upn">$safeUpn</div>
        </div>
        <span class="badge $badgeClass">$safeStatus</span>
    </div>
    <div class="user-card-line"><strong>Department:</strong> $safeDept</div>
    <div class="user-card-line"><strong>Job Title:</strong> $safeTitle</div>
    <div class="user-card-line"><strong>Group:</strong> $safeGroup</div>
    <div class="user-card-line"><strong>License:</strong> $safeLicense</div>
    <div class="user-card-line"><strong>Notes:</strong> $safeNotes</div>
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

    $cards = if ($Rows -and $Rows.Count -gt 0) {
        ($Rows | ForEach-Object { Get-UserCardHtml -Row $_ }) -join [Environment]::NewLine
    }
    else {
        '<div class="empty-state">No users were recorded in this section for this run.</div>'
    }

@"
<div class="$sectionClass">
    <h2>$([System.Net.WebUtility]::HtmlEncode($Title))</h2>
    <div class="user-card-grid">
        $cards
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
        $safeName    = Convert-ToHtmlSafe $row.DisplayName
        $safeUpn     = Convert-ToHtmlSafe $row.UserPrincipalName
        $safeStatus  = Convert-ToHtmlSafe $row.Status
        $safeDept    = Convert-ToHtmlSafe $row.Department
        $safeTitle   = Convert-ToHtmlSafe $row.JobTitle
        $safeGroup   = Convert-ToHtmlSafe $row.GroupName
        $safeLicense = Convert-ToHtmlSafe $row.LicenseSku
        $safeNotes   = Convert-ToHtmlSafe $row.Notes

@"
<tr>
    <td>$safeName</td>
    <td>$safeUpn</td>
    <td>$safeStatus</td>
    <td>$safeDept</td>
    <td>$safeTitle</td>
    <td>$safeGroup</td>
    <td>$safeLicense</td>
    <td>$safeNotes</td>
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
                <th>Group</th>
                <th>License</th>
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
        [string]$RunId,
        [string]$TenantDomain,
        [string]$CsvPath,
        [int]$CsvRowsProcessed,
        [int]$Created,
        [int]$Updated,
        [int]$Skipped,
        [int]$Warnings,
        [int]$Failed,
        [int]$ArchiveFilesRemoved,
        [double]$RunDurationSeconds,
        [string]$TimestampedReportPath,
        [string]$TimestampedLogPath,
        [string]$LatestReportPath,
        [string]$LatestLogPath,
        [string]$TimestampedSummaryPath,
        [string]$LatestSummaryPath,
        [object[]]$ResultRows,
        [int]$KeepLatestArchives
    )

    $generatedAt = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"
    $headline = "This run processed $CsvRowsProcessed user(s): $Created created, $Updated updated, $Skipped already compliant."

    $statCards = @(
        New-StatCardHtml -Label "Run ID"                 -Value $RunId
        New-StatCardHtml -Label "Tenant Domain"          -Value $TenantDomain
        New-StatCardHtml -Label "CSV Rows Processed"     -Value $CsvRowsProcessed
        New-StatCardHtml -Label "Created"                -Value $Created -AccentClass "accent-success"
        New-StatCardHtml -Label "Updated"                -Value $Updated -AccentClass "accent-info"
        New-StatCardHtml -Label "Already Compliant"      -Value $Skipped -AccentClass "accent-neutral"
        New-StatCardHtml -Label "Warnings"               -Value $Warnings -AccentClass "accent-warning"
        New-StatCardHtml -Label "Failed"                 -Value $Failed -AccentClass "accent-danger"
        New-StatCardHtml -Label "Archive Files Removed"  -Value $ArchiveFilesRemoved
        New-StatCardHtml -Label "Run Duration (Seconds)" -Value $RunDurationSeconds
    ) -join [Environment]::NewLine

    $runSettingsHtml = New-ListItemsHtml -Items @(
        "CSV Path: $CsvPath"
        "Tenant Domain: $TenantDomain"
        "Archive Retention Count: $KeepLatestArchives"
        "Generated: $generatedAt"
    )

    $statusBreakdownHtml = New-StatusBreakdownHtml -Rows $ResultRows

    $reportingNotesHtml = New-ListItemsHtml -Items @(
        "Created: A new Microsoft 365 user was created from the CSV row."
        "Updated: An existing user was reconciled and at least one requested value changed."
        "Already Compliant: The user already matched the requested onboarding state, so no action was needed."
        "Failed: The row could not be completed because of a validation or processing error."
    )

    $runMetadataHtml = New-ListItemsHtml -Items @(
        "Run ID: $RunId"
        "Rows recorded in result CSV: $($ResultRows.Count)"
        "Warnings generated: $Warnings"
        "Retention policy keeps latest $KeepLatestArchives archive file(s) per artifact type."
    )

    $outputFilesHtml = @(
        New-PathBlockHtml -Label "Timestamped Result CSV"   -PathValue $TimestampedReportPath
        New-PathBlockHtml -Label "Timestamped Log"          -PathValue $TimestampedLogPath
        New-PathBlockHtml -Label "Latest Result CSV"        -PathValue $LatestReportPath
        New-PathBlockHtml -Label "Latest Log"               -PathValue $LatestLogPath
        New-PathBlockHtml -Label "Timestamped HTML Summary" -PathValue $TimestampedSummaryPath
        New-PathBlockHtml -Label "Latest HTML Summary"      -PathValue $LatestSummaryPath
    ) -join [Environment]::NewLine

    $createdRows = @($ResultRows | Where-Object { $_.Status -eq "Created" })
    $updatedRows = @($ResultRows | Where-Object { $_.Status -eq "Updated" })
    $skippedRows = @($ResultRows | Where-Object { $_.Status -eq "Skipped - Already Compliant" })
    $failedRows  = @($ResultRows | Where-Object { $_.Status -eq "Failed" })

    $createdSectionHtml = New-UserSectionHtml -Title "Created Users" -Rows $createdRows
    $updatedSectionHtml = New-UserSectionHtml -Title "Updated Users" -Rows $updatedRows
    $skippedSectionHtml = New-UserSectionHtml -Title "Already Compliant Users" -Rows $skippedRows
    $failedSectionHtml  = New-UserSectionHtml -Title "Failed Users" -Rows $failedRows
    $processedUsersTableHtml = New-ProcessedRowsTableHtml -Rows $ResultRows

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Bulk Onboarding Summary</title>
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

        .stat-card.accent-neutral {
            border-left-color: var(--neutral);
            background: linear-gradient(180deg, var(--neutral-bg) 0%, rgba(255,255,255,0.98) 100%);
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
            <h1>Microsoft 365 Bulk Onboarding Summary</h1>
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
                        <ul>
                            $runSettingsHtml
                        </ul>
                    </div>

                    <div class="panel">
                        <h2>Status Breakdown</h2>
                        <ul>
                            $statusBreakdownHtml
                        </ul>
                    </div>

                    <div class="panel">
                        <h2>Status Definitions</h2>
                        <ul>
                            $reportingNotesHtml
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
            $createdSectionHtml
            $updatedSectionHtml
            $skippedSectionHtml
            $failedSectionHtml

            <div class="wide-panel table-panel">
                <h2>Processed Users</h2>
                $processedUsersTableHtml
            </div>
        </div>

        <div class="footer">
            Generated by Onboarding-New-M365Users.ps1
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}

# -----------------------------
# Validate CSV
# -----------------------------
if (-not (Test-Path $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

Write-Log "Starting bulk onboarding script."
Write-Log ("Run ID: {0}" -f $RunId)
Write-Log ("CSV path: {0}" -f $CsvPath)
Write-Log ("Result output path: {0}" -f $ResultPath)
Write-Log ("Log output path: {0}" -f $LogPath)
Write-Log ("Summary output path: {0}" -f $SummaryPath)
Write-Log ("Tenant domain: {0}" -f $TenantDomain)
Write-Log ("Archive retention count: {0}" -f $KeepLatestArchives)

try {
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Group.Read.All","GroupMember.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All" -NoWelcome
    Write-Log "Connected to Microsoft Graph successfully."

    Write-Log "Importing CSV input..."
    $csvUsers = Import-Csv -Path $CsvPath
    Write-Log ("Imported {0} CSV rows." -f $csvUsers.Count)

    Write-Log "Retrieving subscribed SKUs..."
    $skuMap = Get-SkuMap
    Write-Log ("Retrieved {0} subscribed SKU mappings." -f $skuMap.Count)

    foreach ($row in $csvUsers) {
        $displayName = Normalize-Text $row.DisplayName
        $firstName   = Normalize-Text $row.FirstName
        $lastName    = Normalize-Text $row.LastName
        $userName    = Normalize-Text $row.UserName
        $department  = Normalize-Text $row.Department
        $jobTitle    = Normalize-Text $row.JobTitle
        $office      = Normalize-Text $row.OfficeLocation
        $city        = Normalize-Text $row.City
        $state       = Normalize-Text $row.State
        $country     = Normalize-Text $row.Country
        $usageLoc    = Normalize-Text $row.UsageLocation
        $mobile      = Normalize-Text $row.MobilePhone
        $business    = Normalize-Text $row.BusinessPhone
        $licenseSku  = Normalize-Text $row.LicenseSku
        $groupName   = Normalize-Text $row.GroupName
        $password    = [string]$row.Password

        $upn = "{0}@{1}" -f $userName, $TenantDomain

        Write-Log ("Processing CSV user: {0}" -f $displayName)

        try {
            if ([string]::IsNullOrWhiteSpace($displayName) -or
                [string]::IsNullOrWhiteSpace($userName) -or
                [string]::IsNullOrWhiteSpace($password)) {

                $failedCount++
                Write-Log ("Required data missing for row with UPN {0}" -f $upn) "ERROR"

                Add-ResultRow `
                    -DisplayName $displayName `
                    -UserPrincipalName $upn `
                    -Department $department `
                    -JobTitle $jobTitle `
                    -GroupName $groupName `
                    -LicenseSku $licenseSku `
                    -Status "Failed" `
                    -Notes "Missing required fields"

                continue
            }

            $existingUser = Get-MgUser -Filter "userPrincipalName eq '$upn'" -Property @(
                "Id",
                "DisplayName",
                "UserPrincipalName",
                "GivenName",
                "Surname",
                "JobTitle",
                "Department",
                "OfficeLocation",
                "City",
                "State",
                "Country",
                "UsageLocation",
                "MobilePhone",
                "BusinessPhones",
                "AssignedLicenses"
            )

            $notes = @()

            if (-not $existingUser) {
                $newUserParams = @{
                    AccountEnabled    = $true
                    DisplayName       = $displayName
                    MailNickname      = $userName
                    UserPrincipalName = $upn
                    GivenName         = $firstName
                    Surname           = $lastName
                    JobTitle          = $jobTitle
                    Department        = $department
                    OfficeLocation    = $office
                    City              = $city
                    State             = $state
                    Country           = $country
                    UsageLocation     = $usageLoc
                    MobilePhone       = $mobile
                    BusinessPhones    = @($business)
                    PasswordProfile   = @{
                        ForceChangePasswordNextSignIn = $false
                        Password = $password
                    }
                }

                $createdUser = New-MgUser @newUserParams
                $createdUser | Out-Null

                $existingUser = Get-MgUser -Filter "userPrincipalName eq '$upn'" -Property @(
                    "Id",
                    "DisplayName",
                    "UserPrincipalName",
                    "GivenName",
                    "Surname",
                    "JobTitle",
                    "Department",
                    "OfficeLocation",
                    "City",
                    "State",
                    "Country",
                    "UsageLocation",
                    "MobilePhone",
                    "BusinessPhones",
                    "AssignedLicenses"
                )

                Write-Log ("User created successfully: {0}" -f $upn)
                $notes += "User created successfully"
                $script:createdCount++
            }
            else {
                Write-Log ("User already exists, starting reconciliation: {0}" -f $upn)
            }

            $updateParams = @{}

            if (Compare-Value -Current $existingUser.DisplayName -Desired $displayName) { $updateParams.DisplayName = $displayName }
            if (Compare-Value -Current $existingUser.GivenName -Desired $firstName)     { $updateParams.GivenName = $firstName }
            if (Compare-Value -Current $existingUser.Surname -Desired $lastName)         { $updateParams.Surname = $lastName }
            if (Compare-Value -Current $existingUser.JobTitle -Desired $jobTitle)        { $updateParams.JobTitle = $jobTitle }
            if (Compare-Value -Current $existingUser.Department -Desired $department)    { $updateParams.Department = $department }
            if (Compare-Value -Current $existingUser.OfficeLocation -Desired $office)    { $updateParams.OfficeLocation = $office }
            if (Compare-Value -Current $existingUser.City -Desired $city)                { $updateParams.City = $city }
            if (Compare-Value -Current $existingUser.State -Desired $state)              { $updateParams.State = $state }
            if (Compare-Value -Current $existingUser.Country -Desired $country)          { $updateParams.Country = $country }
            if (Compare-Value -Current $existingUser.UsageLocation -Desired $usageLoc)   { $updateParams.UsageLocation = $usageLoc }
            if (Compare-Value -Current $existingUser.MobilePhone -Desired $mobile)       { $updateParams.MobilePhone = $mobile }

            $currentBusinessPhone = ""
            if ($existingUser.BusinessPhones -and $existingUser.BusinessPhones.Count -gt 0) {
                $currentBusinessPhone = Normalize-Text $existingUser.BusinessPhones[0]
            }
            if (Compare-Value -Current $currentBusinessPhone -Desired $business) {
                $updateParams.BusinessPhones = @($business)
            }

            $profileUpdated = $false
            if ($updateParams.Count -gt 0) {
                Update-MgUser -UserId $existingUser.Id @updateParams | Out-Null
                Write-Log ("Updated profile fields for {0}" -f $upn)
                $notes += "Profile updated"
                $profileUpdated = $true
            }

            $existingUser = Get-MgUser -Filter "userPrincipalName eq '$upn'" -Property @(
                "Id",
                "DisplayName",
                "UserPrincipalName",
                "GivenName",
                "Surname",
                "JobTitle",
                "Department",
                "OfficeLocation",
                "City",
                "State",
                "Country",
                "UsageLocation",
                "MobilePhone",
                "BusinessPhones",
                "AssignedLicenses"
            )

            $licenseUpdated = $false
            if (-not [string]::IsNullOrWhiteSpace($licenseSku)) {
                $currentLicenseSkuParts = Get-UserLicenseSkuPartNumbers -User $existingUser -SkuMap $skuMap

                if ($currentLicenseSkuParts -notcontains $licenseSku) {
                    if ($skuMap.ContainsKey($licenseSku)) {
                        $skuId = $skuMap[$licenseSku]

                        Set-MgUserLicense -UserId $existingUser.Id -AddLicenses @(
                            @{
                                SkuId = $skuId
                            }
                        ) -RemoveLicenses @() | Out-Null

                        Write-Log ("Assigned license {0} to {1}" -f $licenseSku, $upn)
                        $notes += "License assigned"
                        $licenseUpdated = $true
                    }
                    else {
                        Write-Log ("License SKU not found in tenant for {0}: {1}" -f $upn, $licenseSku) "WARN"
                        $notes += "License SKU not found"
                    }
                }
            }
            else {
                Write-Log ("No license requested for {0}" -f $upn)
                $notes += "No license requested"
            }

            $groupUpdated = $false
            if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                $currentGroups = Get-UserGroupNames -UserId $existingUser.Id

                if ($currentGroups -notcontains $groupName) {
                    $group = Get-GroupByDisplayName -GroupName $groupName

                    if ($group) {
                        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $existingUser.Id | Out-Null
                        Write-Log ("Added {0} to group {1}" -f $upn, $groupName)
                        $notes += "Group assigned"
                        $groupUpdated = $true
                    }
                    else {
                        Write-Log ("Group not found for {0}: {1}" -f $upn, $groupName) "WARN"
                        $notes += "Group not found"
                    }
                }
            }
            else {
                Write-Log ("No group requested for {0}" -f $upn)
                $notes += "No group requested"
            }

            if ($notes -contains "User created successfully") {
                Add-ResultRow `
                    -DisplayName $displayName `
                    -UserPrincipalName $upn `
                    -Department $department `
                    -JobTitle $jobTitle `
                    -GroupName $groupName `
                    -LicenseSku $licenseSku `
                    -Status "Created" `
                    -Notes ($notes -join "; ")
            }
            elseif ($profileUpdated -or $licenseUpdated -or $groupUpdated) {
                $script:updatedCount++
                Add-ResultRow `
                    -DisplayName $displayName `
                    -UserPrincipalName $upn `
                    -Department $department `
                    -JobTitle $jobTitle `
                    -GroupName $groupName `
                    -LicenseSku $licenseSku `
                    -Status "Updated" `
                    -Notes ($notes -join "; ")
            }
            else {
                $script:skippedCount++
                Add-ResultRow `
                    -DisplayName $displayName `
                    -UserPrincipalName $upn `
                    -Department $department `
                    -JobTitle $jobTitle `
                    -GroupName $groupName `
                    -LicenseSku $licenseSku `
                    -Status "Skipped - Already Compliant" `
                    -Notes "No action needed"
            }
        }
        catch {
            $script:failedCount++
            Write-Log ("Failed processing {0}. Error: {1}" -f $displayName, $_.Exception.Message) "ERROR"

            Add-ResultRow `
                -DisplayName $displayName `
                -UserPrincipalName $upn `
                -Department $department `
                -JobTitle $jobTitle `
                -GroupName $groupName `
                -LicenseSku $licenseSku `
                -Status "Failed" `
                -Notes $_.Exception.Message
        }
    }

    $results |
        Sort-Object DisplayName |
        Export-Csv -Path $ResultPath -NoTypeInformation -Encoding UTF8

    Write-Log ("Bulk onboarding result CSV exported successfully: {0}" -f $ResultPath)

    Copy-Item -Path $ResultPath -Destination $LatestResultPath -Force
    Write-Log ("Latest result CSV updated: {0}" -f $LatestResultPath)
}
catch {
    $script:failedCount++
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
        Write-Log "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph | Out-Null
        Write-Log "Disconnected successfully."
    }
    catch {
        Write-Log ("Disconnect warning: {0}" -f $_.Exception.Message) "WARN"
    }

    $EndTime  = Get-Date
    $Duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalSeconds, 2)

    try {
        Copy-Item -Path $LogPath -Destination $LatestLogPath -Force
    }
    catch {
        Write-Host ("Could not create latest log copy: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        New-HtmlSummary `
            -OutputPath $SummaryPath `
            -RunId $RunId `
            -TenantDomain $TenantDomain `
            -CsvPath $CsvPath `
            -CsvRowsProcessed $csvUsers.Count `
            -Created $createdCount `
            -Updated $updatedCount `
            -Skipped $skippedCount `
            -Warnings $warnCount `
            -Failed $failedCount `
            -ArchiveFilesRemoved $cleanupRemoved `
            -RunDurationSeconds $Duration `
            -TimestampedReportPath $ResultPath `
            -TimestampedLogPath $LogPath `
            -LatestReportPath $LatestResultPath `
            -LatestLogPath $LatestLogPath `
            -TimestampedSummaryPath $SummaryPath `
            -LatestSummaryPath $LatestSummaryPath `
            -ResultRows $results `
            -KeepLatestArchives $KeepLatestArchives

        Copy-Item -Path $SummaryPath -Destination $LatestSummaryPath -Force
    }
    catch {
        Write-Host ("HTML summary generation warning: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    Write-Log ("Latest log copy updated: {0}" -f $LatestLogPath)
    Write-Log ("Latest HTML summary copy updated: {0}" -f $LatestSummaryPath)
}

$EndTime = Get-Date
$Duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalSeconds, 2)

Write-Host ""
Write-Host "========== SUMMARY ==========" -ForegroundColor Cyan
Write-Host ("Run ID: {0}" -f $RunId)
Write-Host ("CSV rows processed: {0}" -f $csvUsers.Count)
Write-Host ("Created: {0}" -f $createdCount)
Write-Host ("Updated: {0}" -f $updatedCount)
Write-Host ("Skipped: {0}" -f $skippedCount)
Write-Host ("Warnings: {0}" -f $warnCount)
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