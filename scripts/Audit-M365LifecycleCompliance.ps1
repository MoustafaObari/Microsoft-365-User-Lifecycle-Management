# ============================================
# Script: Audit-M365LifecycleCompliance.ps1
# Purpose: Read-only audit of Microsoft 365 user lifecycle compliance
#          against expected department standards from CSV.
# Author: Moustafa Obari
# Project: Microsoft 365 User Lifecycle Management and Administration
# Recommended Runtime: PowerShell 7
#
# Expected standards CSV columns:
# Department,RequiredGroups,RequiredLicenses,ForbiddenGroups,ForbiddenLicenses,ExpectedUsageLocation,ExpectedAccountEnabled,Notes
#
# Notes:
# - Use pipe separators for multi-value columns, for example:
#   RequiredGroups = SG_HR_Users|SG_HR_Shared
#   RequiredLicenses = SPB|POWER_BI_PRO
# - You can add a DEFAULT row in Department to act as a fallback standard.
# - This script is READ-ONLY. It does not modify users, licenses, or groups.
# ============================================

param(
    [string]$StandardsCsvPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "inputs\M365-LifecycleStandards.csv"),
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
$ProjectRoot    = Split-Path -Parent $PSScriptRoot
$ReportFolder   = Join-Path $ProjectRoot "reports"
$LogFolder      = Join-Path $ProjectRoot "logs"
$SummaryFolder  = Join-Path $ProjectRoot "summary"

$TimeStamp = Get-Date -Format "yyyy-MM-dd__hh-mm-sstt"
$RunId     = "AUDIT-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmss")

$BaseName        = "M365-LifecycleAudit"
$SummaryBaseName = "M365-LifecycleAuditSummary"

$LogPath           = Join-Path $LogFolder "$BaseName-$TimeStamp.log"
$ResultPath        = Join-Path $ReportFolder "$BaseName-$TimeStamp.csv"
$SummaryPath       = Join-Path $SummaryFolder "$SummaryBaseName-$TimeStamp.html"

$LatestLogPath     = Join-Path $LogFolder "$BaseName-Latest.log"
$LatestResultPath  = Join-Path $ReportFolder "$BaseName-Latest.csv"
$LatestSummaryPath = Join-Path $SummaryFolder "$SummaryBaseName-Latest.html"

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

$totalUsersRetrieved   = 0
$usersReviewedCount    = 0
$standardsLoadedCount  = 0
$findingsRecordedCount = 0
$compliantUsersCount   = 0
$reviewCount           = 0
$warningCount          = 0
$highRiskCount         = 0
$criticalCount         = 0
$failedCount           = 0
$cleanupRemoved        = 0

$results      = @()
$standards    = @()
$csvStandards = @()
$users        = @()

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

function Get-RowValue {
    param(
        [Parameter(Mandatory)]
        [object]$Row,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    if ($Row.PSObject.Properties.Name -contains $PropertyName) {
        return Normalize-Text $Row.$PropertyName
    }

    return ""
}

function Split-MultiValue {
    param([string]$Value)

    $normalized = Normalize-Text $Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    return @(
        $normalized -split '\|' |
        ForEach-Object { Normalize-Text $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
    )
}

function Convert-ListToText {
    param([object[]]$Items)

    $filtered = @(
        $Items |
        ForEach-Object { Normalize-Text $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($filtered.Count -eq 0) {
        return "None"
    }

    return ($filtered -join "; ")
}

function Parse-NullableBoolean {
    param([string]$Value)

    $normalized = (Normalize-Text $Value).ToLowerInvariant()

    switch ($normalized) {
        ""      { return $null }
        "true"  { return $true }
        "false" { return $false }
        "yes"   { return $true }
        "no"    { return $false }
        "1"     { return $true }
        "0"     { return $false }
        default { return $null }
    }
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

function Get-UserGroupNames {
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    $groupNames = @()

    try {
        $memberOf = Get-MgUserMemberOf -UserId $UserId -All -ErrorAction Stop
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
        [hashtable]$SkuIdToPartMap
    )

    $assigned = @()

    if ($null -ne $User.AssignedLicenses) {
        foreach ($license in $User.AssignedLicenses) {
            $skuId = Normalize-Text ([string]$license.SkuId)
            if (-not [string]::IsNullOrWhiteSpace($skuId)) {
                if ($SkuIdToPartMap.ContainsKey($skuId)) {
                    $assigned += $SkuIdToPartMap[$skuId]
                }
                else {
                    $assigned += $skuId
                }
            }
        }
    }

    return @($assigned | Sort-Object -Unique)
}

function Get-ArrayIntersection {
    param(
        [string[]]$Left,
        [string[]]$Right
    )

    $output = @()
    foreach ($item in $Left) {
        if ($Right -contains $item) {
            $output += $item
        }
    }

    return @($output | Sort-Object -Unique)
}

function Get-ArrayMissingItems {
    param(
        [string[]]$Expected,
        [string[]]$Actual
    )

    $missing = @()
    foreach ($item in $Expected) {
        if ($Actual -notcontains $item) {
            $missing += $item
        }
    }

    return @($missing | Sort-Object -Unique)
}

function Get-StandardForDepartment {
    param(
        [string]$Department,
        [object[]]$Standards
    )

    $normalizedDepartment = Normalize-Text $Department

    if (-not [string]::IsNullOrWhiteSpace($normalizedDepartment)) {
        $exact = $Standards | Where-Object { $_.Department -eq $normalizedDepartment } | Select-Object -First 1
        if ($exact) {
            return $exact
        }
    }

    $default = $Standards | Where-Object { $_.Department -eq "DEFAULT" } | Select-Object -First 1
    return $default
}

function Add-ResultRow {
    param(
        [string]$DisplayName,
        [string]$UserPrincipalName,
        [string]$UserType,
        [string]$Department,
        [string]$JobTitle,
        [string]$UsageLocation,
        [bool]$AccountEnabled,
        [string]$AssignedLicenses,
        [int]$LicenseCount,
        [string]$GroupMemberships,
        [int]$GroupCount,
        [string]$MatchedStandard,
        [string]$ExpectedAccountEnabled,
        [string]$ExpectedUsageLocation,
        [string]$ExpectedGroups,
        [string]$ExpectedLicenses,
        [string]$ForbiddenGroups,
        [string]$ForbiddenLicenses,
        [string]$CheckName,
        [string]$FindingCategory,
        [string]$Severity,
        [string]$Status,
        [string]$Details,
        [string]$RecommendedAction,
        [string]$StandardNotes
    )

    $script:results += [PSCustomObject]@{
        RunId                  = $RunId
        TenantName             = $tenantDisplayName
        TenantId               = $tenantId
        AuditTimestamp         = (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt")
        DisplayName            = $DisplayName
        UserPrincipalName      = $UserPrincipalName
        UserType               = $UserType
        Department             = $Department
        JobTitle               = $JobTitle
        UsageLocation          = $UsageLocation
        AccountEnabled         = $AccountEnabled
        LicenseCount           = $LicenseCount
        AssignedLicenses       = $AssignedLicenses
        GroupCount             = $GroupCount
        GroupMemberships       = $GroupMemberships
        MatchedStandard        = $MatchedStandard
        ExpectedAccountEnabled = $ExpectedAccountEnabled
        ExpectedUsageLocation  = $ExpectedUsageLocation
        ExpectedGroups         = $ExpectedGroups
        ExpectedLicenses       = $ExpectedLicenses
        ForbiddenGroups        = $ForbiddenGroups
        ForbiddenLicenses      = $ForbiddenLicenses
        CheckName              = $CheckName
        FindingCategory        = $FindingCategory
        Severity               = $Severity
        Status                 = $Status
        Details                = $Details
        RecommendedAction      = $RecommendedAction
        StandardNotes          = $StandardNotes
    }

    $script:findingsRecordedCount++

    switch ($Severity) {
        "REVIEW"    { $script:reviewCount++ }
        "WARNING"   { $script:warningCount++ }
        "HIGH_RISK" { $script:highRiskCount++ }
        "CRITICAL"  { $script:criticalCount++ }
        "FAILED"    { $script:failedCount++ }
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

function Get-FindingCardHtml {
    param([pscustomobject]$Row)

    $badgeClass = switch ($Row.Severity) {
        "CRITICAL"  { "badge-danger"; break }
        "HIGH_RISK" { "badge-warning"; break }
        "WARNING"   { "badge-info"; break }
        "REVIEW"    { "badge-neutral"; break }
        "FAILED"    { "badge-danger"; break }
        "COMPLIANT" { "badge-success"; break }
        default     { "badge-neutral" }
    }

@"
<div class="user-card">
    <div class="user-card-header">
        <div>
            <div class="user-card-name">$(Convert-ToHtmlSafe $Row.DisplayName)</div>
            <div class="user-card-upn">$(Convert-ToHtmlSafe $Row.UserPrincipalName)</div>
        </div>
        <span class="badge $badgeClass">$(Convert-ToHtmlSafe $Row.Severity)</span>
    </div>
    <div class="user-card-line"><strong>Department:</strong> $(Convert-ToHtmlSafe $(if ([string]::IsNullOrWhiteSpace($Row.Department)) { "Not set" } else { $Row.Department }))</div>
    <div class="user-card-line"><strong>Check:</strong> $(Convert-ToHtmlSafe $Row.CheckName)</div>
    <div class="user-card-line"><strong>Details:</strong> $(Convert-ToHtmlSafe $Row.Details)</div>
    <div class="user-card-line"><strong>Recommended Action:</strong> $(Convert-ToHtmlSafe $Row.RecommendedAction)</div>
</div>
"@
}

function New-FindingSectionHtml {
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
        ($Rows | Select-Object -First 15 | ForEach-Object { Get-FindingCardHtml -Row $_ }) -join [Environment]::NewLine
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
        return '<div class="empty-state">No processed findings to display.</div>'
    }

    $body = foreach ($row in $Rows) {
@"
<tr>
    <td>$(Convert-ToHtmlSafe $row.DisplayName)</td>
    <td>$(Convert-ToHtmlSafe $row.UserPrincipalName)</td>
    <td>$(Convert-ToHtmlSafe $row.Severity)</td>
    <td>$(Convert-ToHtmlSafe $row.FindingCategory)</td>
    <td>$(Convert-ToHtmlSafe $row.CheckName)</td>
    <td>$(Convert-ToHtmlSafe $row.Details)</td>
    <td>$(Convert-ToHtmlSafe $row.RecommendedAction)</td>
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
                <th>Severity</th>
                <th>Category</th>
                <th>Check</th>
                <th>Details</th>
                <th>Recommended Action</th>
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
        [string]$StandardsCsvPath,
        [int]$StandardsLoaded,
        [int]$UsersRetrieved,
        [int]$UsersReviewed,
        [int]$CompliantUsers,
        [int]$FindingsRecorded,
        [int]$ReviewCount,
        [int]$WarningCount,
        [int]$HighRiskCount,
        [int]$CriticalCount,
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
        [string]$DepartmentFilter
    )

    $generatedAt = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"
    $headline = "This audit reviewed $UsersReviewed user(s), recorded $FindingsRecorded finding row(s), and marked $CompliantUsers user(s) as fully compliant."

    $statCards = @(
        New-StatCardHtml -Label "Run ID" -Value $RunId
        New-StatCardHtml -Label "Tenant Name" -Value $tenantDisplayName
        New-StatCardHtml -Label "Tenant ID" -Value $tenantId
        New-StatCardHtml -Label "Standards Loaded" -Value $StandardsLoaded
        New-StatCardHtml -Label "Users Retrieved" -Value $UsersRetrieved
        New-StatCardHtml -Label "Users Reviewed" -Value $UsersReviewed
        New-StatCardHtml -Label "Compliant Users" -Value $CompliantUsers -AccentClass "accent-success"
        New-StatCardHtml -Label "Review" -Value $ReviewCount
        New-StatCardHtml -Label "Warning" -Value $WarningCount -AccentClass "accent-info"
        New-StatCardHtml -Label "High Risk" -Value $HighRiskCount -AccentClass "accent-warning"
        New-StatCardHtml -Label "Critical" -Value $CriticalCount -AccentClass "accent-danger"
        New-StatCardHtml -Label "Failed" -Value $FailedCount -AccentClass "accent-danger"
        New-StatCardHtml -Label "Archive Files Removed" -Value $ArchiveFilesRemoved
        New-StatCardHtml -Label "Run Duration (Seconds)" -Value $RunDurationSeconds
    ) -join [Environment]::NewLine

    $runSettingsHtml = New-ListItemsHtml -Items @(
        "Standards CSV Path: $StandardsCsvPath"
        "Department Filter: $(if ([string]::IsNullOrWhiteSpace($DepartmentFilter)) { 'Not used' } else { $DepartmentFilter })"
        "Archive Retention Count: $KeepLatestArchives"
        "Generated: $generatedAt"
    )

    $auditDefinitionsHtml = New-ListItemsHtml -Items @(
        "Compliant Users: Users with no lifecycle compliance findings in this run."
        "Review: User data should be reviewed, but the finding is not immediately high impact."
        "Warning: A standards mismatch exists and should be corrected."
        "High Risk: A lifecycle gap exists that can affect access, licensing, or access control quality."
        "Critical: A severe lifecycle issue exists, such as disabled but still licensed."
    )

    $importantNotesHtml = New-ListItemsHtml -Items @(
        "This script is read-only and does not change users, licenses, or group memberships."
        "Pipe-separated values are supported in standards columns such as RequiredGroups and RequiredLicenses."
        "If no exact department standard is found, the script will use a DEFAULT row when present."
        "Disabled but still licensed is always treated as a critical finding."
        "Disabled but still in active baseline groups is always treated as a high-risk finding."
        "Correlated findings are deduped so the same root cause is not reported twice for the same user."
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

    $criticalRows  = @($ResultRows | Where-Object { $_.Severity -eq "CRITICAL" })
    $highRiskRows  = @($ResultRows | Where-Object { $_.Severity -eq "HIGH_RISK" })
    $warningRows   = @($ResultRows | Where-Object { $_.Severity -in @("WARNING", "REVIEW") })
    $compliantRows = @($ResultRows | Where-Object { $_.Severity -eq "COMPLIANT" })

    $criticalSectionHtml  = New-FindingSectionHtml -Title "Critical Findings" -Rows $criticalRows
    $highRiskSectionHtml  = New-FindingSectionHtml -Title "High Risk Findings" -Rows $highRiskRows
    $warningSectionHtml   = New-FindingSectionHtml -Title "Warning / Review Findings" -Rows $warningRows
    $compliantSectionHtml = New-FindingSectionHtml -Title "Compliant Users" -Rows $compliantRows

    $processedUsersTableHtml = New-ProcessedRowsTableHtml -Rows $ResultRows

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Lifecycle Audit Summary</title>
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
            <h1>Microsoft 365 Lifecycle Audit Summary</h1>
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
                        <h2>Audit Definitions</h2>
                        <ul>$auditDefinitionsHtml</ul>
                    </div>

                    <div class="panel">
                        <h2>Important Notes</h2>
                        <ul>$importantNotesHtml</ul>
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
            $criticalSectionHtml
            $highRiskSectionHtml
            $warningSectionHtml
            $compliantSectionHtml

            <div class="wide-panel table-panel">
                <h2>Processed Findings</h2>
                $processedUsersTableHtml
            </div>
        </div>

        <div class="footer">
            Generated by Audit-M365LifecycleCompliance.ps1
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
Write-Log "Starting lifecycle compliance audit script."
Write-Log ("Run ID: {0}" -f $RunId)
Write-Log ("Standards CSV path: {0}" -f $StandardsCsvPath)
Write-Log ("Department filter: {0}" -f $(if ([string]::IsNullOrWhiteSpace($DepartmentFilter)) { "Not used" } else { $DepartmentFilter }))
Write-Log ("Result output path: {0}" -f $ResultPath)
Write-Log ("Log output path: {0}" -f $LogPath)
Write-Log ("Summary output path: {0}" -f $SummaryPath)
Write-Log ("Archive retention count: {0}" -f $KeepLatestArchives)

# -----------------------------
# Validate standards CSV
# -----------------------------
if (-not (Test-Path $StandardsCsvPath)) {
    throw "Standards CSV file not found: $StandardsCsvPath"
}

try {
    $fileInfo = Get-Item -Path $StandardsCsvPath -ErrorAction Stop
    Write-Log ("Standards CSV file size (bytes): {0}" -f $fileInfo.Length)

    $rawLines = Get-Content -Path $StandardsCsvPath -ErrorAction Stop
    Write-Log ("Standards CSV raw line count: {0}" -f $rawLines.Count)

    if ($rawLines.Count -gt 0) {
        Write-Log ("Standards CSV header preview: {0}" -f $rawLines[0])
    }

    if ($rawLines.Count -gt 1) {
        Write-Log ("Standards CSV first data row preview: {0}" -f $rawLines[1])
    }
    else {
        Write-Log "Standards CSV does not contain a visible data row after the header." "WARN"
    }
}
catch {
    Write-Log ("Could not pre-read standards CSV. Error: {0}" -f $_.Exception.Message) "WARN"
}

try {
    # -----------------------------
    # Connect to Graph
    # -----------------------------
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","Directory.Read.All","Organization.Read.All" -NoWelcome | Out-Null
    Write-Log "Connected to Microsoft Graph successfully."

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

        Write-Log ("Tenant name: {0}" -f $tenantDisplayName)
        Write-Log ("Tenant ID: {0}" -f $tenantId)
    }
    catch {
        Write-Log ("Could not retrieve tenant information. Error: {0}" -f $_.Exception.Message) "WARN"
    }

    # -----------------------------
    # Import standards CSV
    # -----------------------------
    Write-Log "Importing standards CSV..."
    $csvStandards = Import-Csv -Path $StandardsCsvPath
    Write-Log ("Imported {0} standards row(s)." -f $csvStandards.Count)

    foreach ($row in $csvStandards) {
        $department = Get-RowValue -Row $row -PropertyName "Department"

        if ([string]::IsNullOrWhiteSpace($department)) {
            Write-Log "Skipped a standards row because Department is blank." "WARN"
            continue
        }

        $standards += [PSCustomObject]@{
            Department             = $department
            RequiredGroups         = Split-MultiValue (Get-RowValue -Row $row -PropertyName "RequiredGroups")
            RequiredLicenses       = Split-MultiValue (Get-RowValue -Row $row -PropertyName "RequiredLicenses")
            ForbiddenGroups        = Split-MultiValue (Get-RowValue -Row $row -PropertyName "ForbiddenGroups")
            ForbiddenLicenses      = Split-MultiValue (Get-RowValue -Row $row -PropertyName "ForbiddenLicenses")
            ExpectedUsageLocation  = Get-RowValue -Row $row -PropertyName "ExpectedUsageLocation"
            ExpectedAccountEnabled = Parse-NullableBoolean (Get-RowValue -Row $row -PropertyName "ExpectedAccountEnabled")
            Notes                  = Get-RowValue -Row $row -PropertyName "Notes"
        }
    }

    $standardsLoadedCount = @($standards).Count
    Write-Log ("Loaded {0} normalized standards row(s)." -f $standardsLoadedCount)

    if ($standardsLoadedCount -eq 0) {
        throw "No usable standards rows were loaded from the standards CSV. Confirm the file is not empty, has a header row, and contains at least one data row."
    }

    $activeBaselineGroups = @(
        $standards |
        ForEach-Object { $_.RequiredGroups } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
    )

    # -----------------------------
    # License mapping
    # -----------------------------
    Write-Log "Retrieving subscribed SKUs for license mapping..."
    $skuMaps = Get-SkuMaps
    $skuIdToPartMap = $skuMaps.IdToPart
    Write-Log ("Retrieved {0} subscribed SKU mapping(s)." -f $skuIdToPartMap.Count)

    # -----------------------------
    # Retrieve users
    # -----------------------------
    Write-Log "Retrieving users from Microsoft 365..."
    $users = Get-MgUser -All -Property @(
        "Id",
        "DisplayName",
        "UserPrincipalName",
        "Department",
        "JobTitle",
        "UsageLocation",
        "AccountEnabled",
        "AssignedLicenses",
        "UserType"
    )

    $users = @($users)
    $totalUsersRetrieved = $users.Count
    Write-Log ("Retrieved {0} users." -f $totalUsersRetrieved)

    if (-not [string]::IsNullOrWhiteSpace($DepartmentFilter)) {
        $normalizedDepartmentFilter = Normalize-Text $DepartmentFilter
        $users = @($users | Where-Object { (Normalize-Text $_.Department) -eq $normalizedDepartmentFilter })
        Write-Log ("Applied department filter '{0}'. Remaining users: {1}" -f $normalizedDepartmentFilter, $users.Count)
    }

    # -----------------------------
    # Process users
    # -----------------------------
    foreach ($user in $users) {
        $safeDisplayName = Normalize-Text $user.DisplayName
        if ([string]::IsNullOrWhiteSpace($safeDisplayName)) {
            $safeDisplayName = Normalize-Text $user.UserPrincipalName
        }

        Write-Log ("Auditing user: {0}" -f $safeDisplayName)

        try {
            $usersReviewedCount++

            $department     = Normalize-Text $user.Department
            $jobTitle       = Normalize-Text $user.JobTitle
            $usageLocation  = Normalize-Text $user.UsageLocation
            $userType       = Normalize-Text $user.UserType
            $upn            = Normalize-Text $user.UserPrincipalName
            $accountEnabled = [bool]$user.AccountEnabled

            $licenseNames = Get-UserLicenseSkuPartNumbers -User $user -SkuIdToPartMap $skuIdToPartMap
            $groupNames   = Get-UserGroupNames -UserId $user.Id

            $licenseCount = @($licenseNames).Count
            $groupCount   = @($groupNames).Count

            $assignedLicensesText = Convert-ListToText -Items $licenseNames
            $groupMembershipsText = Convert-ListToText -Items $groupNames

            $standard = Get-StandardForDepartment -Department $department -Standards $standards

            $matchedStandardName       = "None"
            $expectedAccountEnabledTxt = "Not Set"
            $expectedUsageLocationTxt  = "Not Set"
            $expectedGroupsTxt         = "None"
            $expectedLicensesTxt       = "None"
            $forbiddenGroupsTxt        = "None"
            $forbiddenLicensesTxt      = "None"
            $standardNotes             = ""

            if ($standard) {
                $matchedStandardName       = $standard.Department
                $expectedAccountEnabledTxt = if ($null -eq $standard.ExpectedAccountEnabled) { "Not Set" } else { [string]$standard.ExpectedAccountEnabled }
                $expectedUsageLocationTxt  = if ([string]::IsNullOrWhiteSpace($standard.ExpectedUsageLocation)) { "Not Set" } else { $standard.ExpectedUsageLocation }
                $expectedGroupsTxt         = Convert-ListToText -Items $standard.RequiredGroups
                $expectedLicensesTxt       = Convert-ListToText -Items $standard.RequiredLicenses
                $forbiddenGroupsTxt        = Convert-ListToText -Items $standard.ForbiddenGroups
                $forbiddenLicensesTxt      = Convert-ListToText -Items $standard.ForbiddenLicenses
                $standardNotes             = $standard.Notes
            }

            $userFindings = New-Object System.Collections.ArrayList

            # flags to suppress correlated duplicate findings
            $hasMissingRequiredGroup   = $false
            $hasMissingRequiredLicense = $false

            if ([string]::IsNullOrWhiteSpace($department)) {
                [void]$userFindings.Add([PSCustomObject]@{
                    CheckName         = "Department Missing"
                    FindingCategory   = "PROFILE_STANDARD_MISMATCH"
                    Severity          = "WARNING"
                    Status            = "NON_COMPLIANT"
                    Details           = "User does not have a Department value, so a reliable standards match cannot be confirmed."
                    RecommendedAction = "Populate Department so the correct lifecycle standard can be applied."
                })
            }

            if (-not $accountEnabled -and $licenseCount -gt 0) {
                [void]$userFindings.Add([PSCustomObject]@{
                    CheckName         = "Disabled But Still Licensed"
                    FindingCategory   = "OFFBOARDING_DRIFT"
                    Severity          = "CRITICAL"
                    Status            = "NON_COMPLIANT"
                    Details           = ("Account is disabled but still has assigned license(s): {0}" -f $assignedLicensesText)
                    RecommendedAction = "Review offboarding state and remove licenses if the user should remain inactive."
                })
            }

            if (-not $accountEnabled -and $activeBaselineGroups.Count -gt 0) {
                $disabledActiveGroups = Get-ArrayIntersection -Left $groupNames -Right $activeBaselineGroups
                if ($disabledActiveGroups.Count -gt 0) {
                    [void]$userFindings.Add([PSCustomObject]@{
                        CheckName         = "Disabled But Still In Active Groups"
                        FindingCategory   = "OFFBOARDING_DRIFT"
                        Severity          = "HIGH_RISK"
                        Status            = "NON_COMPLIANT"
                        Details           = ("Disabled account is still a member of active baseline group(s): {0}" -f (Convert-ListToText -Items $disabledActiveGroups))
                        RecommendedAction = "Review whether the user should be removed from active access groups."
                    })
                }
            }

            if (-not $standard) {
                if (-not [string]::IsNullOrWhiteSpace($department)) {
                    [void]$userFindings.Add([PSCustomObject]@{
                        CheckName         = "No Matching Department Standard"
                        FindingCategory   = "STANDARD_MAPPING"
                        Severity          = "REVIEW"
                        Status            = "NON_COMPLIANT"
                        Details           = ("No lifecycle standards row matched Department '{0}', and no DEFAULT row was available." -f $department)
                        RecommendedAction = "Add a standards row for this department or add a DEFAULT row."
                    })
                }
            }
            else {
                if ($null -ne $standard.ExpectedAccountEnabled -and ([bool]$accountEnabled -ne [bool]$standard.ExpectedAccountEnabled)) {
                    $detail = "AccountEnabled is '$accountEnabled' but the matched standard expects '$($standard.ExpectedAccountEnabled)'."
                    $severity = if (($standard.ExpectedAccountEnabled -eq $false) -and ($accountEnabled -eq $true)) { "HIGH_RISK" } else { "WARNING" }

                    [void]$userFindings.Add([PSCustomObject]@{
                        CheckName         = "Account Enabled State Mismatch"
                        FindingCategory   = "ACCOUNT_STATE_MISMATCH"
                        Severity          = $severity
                        Status            = "NON_COMPLIANT"
                        Details           = $detail
                        RecommendedAction = "Review user lifecycle state and align the enabled/disabled state with the standard."
                    })
                }

                if (-not [string]::IsNullOrWhiteSpace($standard.ExpectedUsageLocation)) {
                    if ([string]::IsNullOrWhiteSpace($usageLocation)) {
                        [void]$userFindings.Add([PSCustomObject]@{
                            CheckName         = "Usage Location Missing"
                            FindingCategory   = "PROFILE_STANDARD_MISMATCH"
                            Severity          = "WARNING"
                            Status            = "NON_COMPLIANT"
                            Details           = ("UsageLocation is blank but the matched standard expects '{0}'." -f $standard.ExpectedUsageLocation)
                            RecommendedAction = "Populate UsageLocation to match the department standard."
                        })
                    }
                    elseif ($usageLocation -ne $standard.ExpectedUsageLocation) {
                        [void]$userFindings.Add([PSCustomObject]@{
                            CheckName         = "Usage Location Mismatch"
                            FindingCategory   = "PROFILE_STANDARD_MISMATCH"
                            Severity          = "WARNING"
                            Status            = "NON_COMPLIANT"
                            Details           = ("UsageLocation is '{0}' but the matched standard expects '{1}'." -f $usageLocation, $standard.ExpectedUsageLocation)
                            RecommendedAction = "Review and update UsageLocation if the user should follow the department baseline."
                        })
                    }
                }

                if ($accountEnabled -and $standard.RequiredGroups.Count -gt 0) {
                    $missingRequiredGroups = Get-ArrayMissingItems -Expected $standard.RequiredGroups -Actual $groupNames
                    if ($missingRequiredGroups.Count -gt 0) {
                        $hasMissingRequiredGroup = $true

                        [void]$userFindings.Add([PSCustomObject]@{
                            CheckName         = "Enabled User Missing Required Group"
                            FindingCategory   = "GROUP_MISMATCH"
                            Severity          = "HIGH_RISK"
                            Status            = "NON_COMPLIANT"
                            Details           = ("Enabled user is missing required group(s): {0}" -f (Convert-ListToText -Items $missingRequiredGroups))
                            RecommendedAction = "Review group membership and add the user to the required baseline group(s) if appropriate."
                        })
                    }
                }

                if ($accountEnabled -and $standard.RequiredLicenses.Count -gt 0) {
                    $missingRequiredLicenses = Get-ArrayMissingItems -Expected $standard.RequiredLicenses -Actual $licenseNames
                    if ($missingRequiredLicenses.Count -gt 0) {
                        $hasMissingRequiredLicense = $true

                        [void]$userFindings.Add([PSCustomObject]@{
                            CheckName         = "Enabled User Missing Required License"
                            FindingCategory   = "LICENSE_MISMATCH"
                            Severity          = "HIGH_RISK"
                            Status            = "NON_COMPLIANT"
                            Details           = ("Enabled user is missing required license(s): {0}" -f (Convert-ListToText -Items $missingRequiredLicenses))
                            RecommendedAction = "Review licensing and assign the expected license baseline if appropriate."
                        })
                    }
                }

                # Correlated checks with dedupe:
                # Only run these if the primary direct mismatch was NOT already logged.
                if ($standard.RequiredGroups.Count -gt 0 -and $standard.RequiredLicenses.Count -gt 0) {
                    if (-not $hasMissingRequiredLicense) {
                        $expectedGroupsUserHas = Get-ArrayIntersection -Left $groupNames -Right $standard.RequiredGroups
                        if ($expectedGroupsUserHas.Count -gt 0) {
                            $missingLinkedLicenses = Get-ArrayMissingItems -Expected $standard.RequiredLicenses -Actual $licenseNames
                            if ($missingLinkedLicenses.Count -gt 0) {
                                [void]$userFindings.Add([PSCustomObject]@{
                                    CheckName         = "Group Assigned But Missing Corresponding License"
                                    FindingCategory   = "LICENSE_MISMATCH"
                                    Severity          = "HIGH_RISK"
                                    Status            = "NON_COMPLIANT"
                                    Details           = ("User is in expected group(s) {0} but is missing expected license(s) {1}." -f (Convert-ListToText -Items $expectedGroupsUserHas), (Convert-ListToText -Items $missingLinkedLicenses))
                                    RecommendedAction = "Review whether group membership and license assignment should align to the same baseline."
                                })
                            }
                        }
                    }

                    if (-not $hasMissingRequiredGroup) {
                        $expectedLicensesUserHas = Get-ArrayIntersection -Left $licenseNames -Right $standard.RequiredLicenses
                        if ($expectedLicensesUserHas.Count -gt 0) {
                            $missingLinkedGroups = Get-ArrayMissingItems -Expected $standard.RequiredGroups -Actual $groupNames
                            if ($missingLinkedGroups.Count -gt 0) {
                                [void]$userFindings.Add([PSCustomObject]@{
                                    CheckName         = "Licensed But Wrong Group"
                                    FindingCategory   = "GROUP_MISMATCH"
                                    Severity          = "HIGH_RISK"
                                    Status            = "NON_COMPLIANT"
                                    Details           = ("User has expected license(s) {0} but is missing expected group(s) {1}." -f (Convert-ListToText -Items $expectedLicensesUserHas), (Convert-ListToText -Items $missingLinkedGroups))
                                    RecommendedAction = "Review whether the user should be added to the expected baseline group(s) or have licensing corrected."
                                })
                            }
                        }
                    }
                }

                if ($standard.ForbiddenLicenses.Count -gt 0) {
                    $forbiddenLicensesAssigned = Get-ArrayIntersection -Left $licenseNames -Right $standard.ForbiddenLicenses
                    if ($forbiddenLicensesAssigned.Count -gt 0) {
                        [void]$userFindings.Add([PSCustomObject]@{
                            CheckName         = "Forbidden License Assigned"
                            FindingCategory   = "LICENSE_MISMATCH"
                            Severity          = "CRITICAL"
                            Status            = "NON_COMPLIANT"
                            Details           = ("User has forbidden license(s) for this standard: {0}" -f (Convert-ListToText -Items $forbiddenLicensesAssigned))
                            RecommendedAction = "Review the matched standard and remove forbidden license assignments if the standard is correct."
                        })
                    }
                }

                if ($standard.ForbiddenGroups.Count -gt 0) {
                    $forbiddenGroupsAssigned = Get-ArrayIntersection -Left $groupNames -Right $standard.ForbiddenGroups
                    if ($forbiddenGroupsAssigned.Count -gt 0) {
                        [void]$userFindings.Add([PSCustomObject]@{
                            CheckName         = "Forbidden Group Assigned"
                            FindingCategory   = "GROUP_MISMATCH"
                            Severity          = "HIGH_RISK"
                            Status            = "NON_COMPLIANT"
                            Details           = ("User is a member of forbidden group(s) for this standard: {0}" -f (Convert-ListToText -Items $forbiddenGroupsAssigned))
                            RecommendedAction = "Review the matched standard and remove forbidden group membership if the standard is correct."
                        })
                    }
                }
            }

            if ($userFindings.Count -eq 0) {
                $compliantUsersCount++

                Add-ResultRow `
                    -DisplayName $safeDisplayName `
                    -UserPrincipalName $upn `
                    -UserType $userType `
                    -Department $department `
                    -JobTitle $jobTitle `
                    -UsageLocation $usageLocation `
                    -AccountEnabled $accountEnabled `
                    -AssignedLicenses $assignedLicensesText `
                    -LicenseCount $licenseCount `
                    -GroupMemberships $groupMembershipsText `
                    -GroupCount $groupCount `
                    -MatchedStandard $matchedStandardName `
                    -ExpectedAccountEnabled $expectedAccountEnabledTxt `
                    -ExpectedUsageLocation $expectedUsageLocationTxt `
                    -ExpectedGroups $expectedGroupsTxt `
                    -ExpectedLicenses $expectedLicensesTxt `
                    -ForbiddenGroups $forbiddenGroupsTxt `
                    -ForbiddenLicenses $forbiddenLicensesTxt `
                    -CheckName "Overall Compliance" `
                    -FindingCategory "COMPLIANCE" `
                    -Severity "COMPLIANT" `
                    -Status "COMPLIANT" `
                    -Details "No lifecycle compliance findings were detected for this user in this run." `
                    -RecommendedAction "No action needed." `
                    -StandardNotes $standardNotes
            }
            else {
                foreach ($finding in $userFindings) {
                    Add-ResultRow `
                        -DisplayName $safeDisplayName `
                        -UserPrincipalName $upn `
                        -UserType $userType `
                        -Department $department `
                        -JobTitle $jobTitle `
                        -UsageLocation $usageLocation `
                        -AccountEnabled $accountEnabled `
                        -AssignedLicenses $assignedLicensesText `
                        -LicenseCount $licenseCount `
                        -GroupMemberships $groupMembershipsText `
                        -GroupCount $groupCount `
                        -MatchedStandard $matchedStandardName `
                        -ExpectedAccountEnabled $expectedAccountEnabledTxt `
                        -ExpectedUsageLocation $expectedUsageLocationTxt `
                        -ExpectedGroups $expectedGroupsTxt `
                        -ExpectedLicenses $expectedLicensesTxt `
                        -ForbiddenGroups $forbiddenGroupsTxt `
                        -ForbiddenLicenses $forbiddenLicensesTxt `
                        -CheckName $finding.CheckName `
                        -FindingCategory $finding.FindingCategory `
                        -Severity $finding.Severity `
                        -Status $finding.Status `
                        -Details $finding.Details `
                        -RecommendedAction $finding.RecommendedAction `
                        -StandardNotes $standardNotes
                }
            }
        }
        catch {
            Add-ResultRow `
                -DisplayName $safeDisplayName `
                -UserPrincipalName (Normalize-Text $user.UserPrincipalName) `
                -UserType (Normalize-Text $user.UserType) `
                -Department (Normalize-Text $user.Department) `
                -JobTitle (Normalize-Text $user.JobTitle) `
                -UsageLocation (Normalize-Text $user.UsageLocation) `
                -AccountEnabled ([bool]$user.AccountEnabled) `
                -AssignedLicenses "Unknown" `
                -LicenseCount 0 `
                -GroupMemberships "Unknown" `
                -GroupCount 0 `
                -MatchedStandard "Unknown" `
                -ExpectedAccountEnabled "Unknown" `
                -ExpectedUsageLocation "Unknown" `
                -ExpectedGroups "Unknown" `
                -ExpectedLicenses "Unknown" `
                -ForbiddenGroups "Unknown" `
                -ForbiddenLicenses "Unknown" `
                -CheckName "Audit Processing Failure" `
                -FindingCategory "SCRIPT_FAILURE" `
                -Severity "FAILED" `
                -Status "FAILED" `
                -Details $_.Exception.Message `
                -RecommendedAction "Review the error and rerun the audit after correcting the issue." `
                -StandardNotes ""

            Write-Log ("Failed auditing user {0}. Error: {1}" -f $safeDisplayName, $_.Exception.Message) "ERROR"
        }
    }

    # -----------------------------
    # Export CSV
    # -----------------------------
    $results |
        Sort-Object DisplayName, Severity, CheckName |
        Export-Csv -Path $ResultPath -NoTypeInformation -Encoding UTF8

    Write-Log ("Lifecycle audit result CSV exported successfully: {0}" -f $ResultPath)

    # -----------------------------
    # HTML summary
    # -----------------------------
    $EndTimeForHtml  = Get-Date
    $DurationForHtml = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTimeForHtml).TotalSeconds, 2)

    New-HtmlSummary `
        -OutputPath $SummaryPath `
        -RunId $RunId `
        -StandardsCsvPath $StandardsCsvPath `
        -StandardsLoaded $standardsLoadedCount `
        -UsersRetrieved $totalUsersRetrieved `
        -UsersReviewed $usersReviewedCount `
        -CompliantUsers $compliantUsersCount `
        -FindingsRecorded $findingsRecordedCount `
        -ReviewCount $reviewCount `
        -WarningCount $warningCount `
        -HighRiskCount $highRiskCount `
        -CriticalCount $criticalCount `
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
        -DepartmentFilter $DepartmentFilter

    Write-Log ("Lifecycle audit HTML summary generated successfully: {0}" -f $SummaryPath)
}
catch {
    $failedCount++
    Write-Log ("Script failed. Error: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
finally {
    # -----------------------------
    # Cleanup old archives
    # -----------------------------
    try {
        Remove-OldArchives -FolderPath $ReportFolder  -BaseName $BaseName        -Extension ".csv"  -KeepCount $KeepLatestArchives
        Remove-OldArchives -FolderPath $LogFolder     -BaseName $BaseName        -Extension ".log"  -KeepCount $KeepLatestArchives
        Remove-OldArchives -FolderPath $SummaryFolder -BaseName $SummaryBaseName -Extension ".html" -KeepCount $KeepLatestArchives
    }
    catch {
        Write-Host ("Archive cleanup warning: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    # -----------------------------
    # Latest copies
    # -----------------------------
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
Write-Host ("Standards loaded: {0}" -f $standardsLoadedCount)
Write-Host ("Total users retrieved: {0}" -f $totalUsersRetrieved)
Write-Host ("Users reviewed: {0}" -f $usersReviewedCount)
Write-Host ("Compliant users: {0}" -f $compliantUsersCount)
Write-Host ("Findings recorded: {0}" -f $findingsRecordedCount)
Write-Host ("Review: {0}" -f $reviewCount)
Write-Host ("Warning: {0}" -f $warningCount)
Write-Host ("High Risk: {0}" -f $highRiskCount)
Write-Host ("Critical: {0}" -f $criticalCount)
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