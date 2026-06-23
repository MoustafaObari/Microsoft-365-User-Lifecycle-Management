# ============================================
# Script: Build-M365LifecycleMasterDashboard.ps1
# Purpose: Build a master Microsoft 365 lifecycle dashboard
#          from the latest onboarding, offboarding, audit,
#          and remediation CSV outputs.
# Author: Moustafa Obari
# Project: Microsoft 365 User Lifecycle Management and Administration
# Recommended Runtime: PowerShell 7
# ============================================

param(
    [string]$OnboardingCsvPath  = (Join-Path (Split-Path -Parent $PSScriptRoot) "reports\M365-BulkOnboarding-Latest.csv"),
    [string]$OffboardingCsvPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "reports\M365-Offboarding-Latest.csv"),
    [string]$AuditCsvPath       = (Join-Path (Split-Path -Parent $PSScriptRoot) "reports\M365-LifecycleAudit-Latest.csv"),
    [string]$RemediationCsvPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "reports\M365-LifecycleRemediation-Latest.csv"),
    [int]$KeepLatestArchives = 7
)

# -----------------------------
# Config
# -----------------------------
$ProjectRoot    = Split-Path -Parent $PSScriptRoot
$ReportFolder   = Join-Path $ProjectRoot "reports"
$LogFolder      = Join-Path $ProjectRoot "logs"
$SummaryFolder  = Join-Path $ProjectRoot "summary"

$TimeStamp = Get-Date -Format "yyyy-MM-dd__hh-mm-sstt"
$RunId     = "DASH-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmss")

$BaseName        = "M365-LifecycleMasterDashboard"
$SummaryBaseName = "M365-LifecycleMasterDashboard"

$LogPath           = Join-Path $LogFolder "$BaseName-$TimeStamp.log"
$SummaryPath       = Join-Path $SummaryFolder "$SummaryBaseName-$TimeStamp.html"

$LatestLogPath     = Join-Path $LogFolder "$BaseName-Latest.log"
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
# Runtime
# -----------------------------
$StartTime = Get-Date
$cleanupRemoved = 0

$onboardingRows  = @()
$offboardingRows = @()
$auditRows       = @()
$remediationRows = @()

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

function Import-CsvIfExists {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        Write-Log ("File not found, continuing with empty dataset: {0}" -f $Path) "WARN"
        return @()
    }

    try {
        $rows = Import-Csv -Path $Path
        return @($rows)
    }
    catch {
        Write-Log ("Failed to import CSV {0}. Error: {1}" -f $Path, $_.Exception.Message) "WARN"
        return @()
    }
}

function Get-FirstExistingPropertyValue {
    param(
        [Parameter(Mandatory)] [object]$Row,
        [Parameter(Mandatory)] [string[]]$PropertyNames
    )

    foreach ($name in $PropertyNames) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            $value = Normalize-Text $Row.$name
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    return ""
}

function Get-CountByValues {
    param(
        [object[]]$Rows,
        [string[]]$PropertyNames,
        [string[]]$MatchValues
    )

    $count = 0

    foreach ($row in $Rows) {
        $value = Get-FirstExistingPropertyValue -Row $row -PropertyNames $PropertyNames
        if ($MatchValues -contains $value) {
            $count++
        }
    }

    return $count
}

function Get-LatestTimestampFromRows {
    param(
        [object[]]$Rows,
        [string[]]$PropertyNames,
        [string]$FallbackText = "Not available"
    )

    $latest = $null

    foreach ($row in $Rows) {
        foreach ($name in $PropertyNames) {
            if ($row.PSObject.Properties.Name -contains $name) {
                $raw = Normalize-Text $row.$name
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    try {
                        $dt = [datetime]::Parse($raw)
                        if ($null -eq $latest -or $dt -gt $latest) {
                            $latest = $dt
                        }
                    }
                    catch {
                    }
                }
            }
        }
    }

    if ($null -eq $latest) {
        return $FallbackText
    }

    return $latest.ToString("yyyy-MM-dd hh:mm:ss tt")
}

function Get-TopRows {
    param(
        [object[]]$Rows,
        [int]$Take = 5
    )

    if (-not $Rows) { return @() }
    return @($Rows | Select-Object -First $Take)
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

function New-SimpleSectionHtml {
    param(
        [string]$Title,
        [string[]]$Items
    )

    $listHtml = New-ListItemsHtml -Items $Items

@"
<div class="panel">
    <h2>$(Convert-ToHtmlSafe $Title)</h2>
    <ul>$listHtml</ul>
</div>
"@
}

function New-TableHtml {
    param(
        [string]$Title,
        [string[]]$Headers,
        [string]$BodyHtml
    )

    $thead = foreach ($header in $Headers) {
        "<th>$(Convert-ToHtmlSafe $header)</th>"
    }

@"
<div class="wide-panel table-panel">
    <h2>$(Convert-ToHtmlSafe $Title)</h2>
    <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    $($thead -join [Environment]::NewLine)
                </tr>
            </thead>
            <tbody>
                $BodyHtml
            </tbody>
        </table>
    </div>
</div>
"@
}

function New-RowHtml {
    param(
        [string[]]$Values
    )

    $cells = foreach ($value in $Values) {
        "<td>$(Convert-ToHtmlSafe $value)</td>"
    }

@"
<tr>
    $($cells -join [Environment]::NewLine)
</tr>
"@
}

function New-HtmlDashboard {
    param(
        [string]$OutputPath,
        [hashtable]$Stats,
        [string]$OnboardingCsvPath,
        [string]$OffboardingCsvPath,
        [string]$AuditCsvPath,
        [string]$RemediationCsvPath
    )

    $generatedAt = Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"
    $headline = "This dashboard combines the latest onboarding, offboarding, audit, and remediation outputs into one master Microsoft 365 lifecycle view."

    $statCards = @(
        New-StatCardHtml -Label "Run ID" -Value $RunId
        New-StatCardHtml -Label "Generated" -Value $generatedAt
        New-StatCardHtml -Label "Onboarding Rows" -Value $Stats.OnboardingRows
        New-StatCardHtml -Label "Offboarding Rows" -Value $Stats.OffboardingRows
        New-StatCardHtml -Label "Audit Rows" -Value $Stats.AuditRows
        New-StatCardHtml -Label "Remediation Rows" -Value $Stats.RemediationRows
        New-StatCardHtml -Label "Current Compliant Users" -Value $Stats.CurrentCompliantUsers -AccentClass "accent-success"
        New-StatCardHtml -Label "Current High Risk" -Value $Stats.CurrentHighRisk -AccentClass "accent-warning"
        New-StatCardHtml -Label "Current Critical" -Value $Stats.CurrentCritical -AccentClass "accent-danger"
        New-StatCardHtml -Label "Current Failed" -Value $Stats.CurrentFailed -AccentClass "accent-danger"
        New-StatCardHtml -Label "Latest Remediated" -Value $Stats.RemediationRemediated -AccentClass "accent-success"
        New-StatCardHtml -Label "Run Duration (Seconds)" -Value $Stats.RunDuration
    ) -join [Environment]::NewLine

    $dataSourcesHtml = @(
        New-PathBlockHtml -Label "Onboarding CSV" -PathValue $OnboardingCsvPath
        New-PathBlockHtml -Label "Offboarding CSV" -PathValue $OffboardingCsvPath
        New-PathBlockHtml -Label "Audit CSV" -PathValue $AuditCsvPath
        New-PathBlockHtml -Label "Remediation CSV" -PathValue $RemediationCsvPath
        New-PathBlockHtml -Label "Timestamped HTML Dashboard" -PathValue $OutputPath
        New-PathBlockHtml -Label "Latest HTML Dashboard" -PathValue $LatestSummaryPath
    ) -join [Environment]::NewLine

    $runSummaryHtml = @(
        New-SimpleSectionHtml -Title "Lifecycle Summary" -Items @(
            "Onboarding latest timestamp: $($Stats.OnboardingLatestTime)"
            "Offboarding latest timestamp: $($Stats.OffboardingLatestTime)"
            "Audit latest timestamp: $($Stats.AuditLatestTime)"
            "Remediation latest timestamp: $($Stats.RemediationLatestTime)"
            "Current compliant users: $($Stats.CurrentCompliantUsers)"
            "Current non-compliant users: $($Stats.CurrentNonCompliantUsers)"
        )
        New-SimpleSectionHtml -Title "Onboarding Snapshot" -Items @(
            "Rows loaded: $($Stats.OnboardingRows)"
            "Succeeded: $($Stats.OnboardingSucceeded)"
            "Skipped: $($Stats.OnboardingSkipped)"
            "Failed: $($Stats.OnboardingFailed)"
        )
        New-SimpleSectionHtml -Title "Offboarding Snapshot" -Items @(
            "Rows loaded: $($Stats.OffboardingRows)"
            "Offboarded: $($Stats.OffboardingOffboarded)"
            "Updated: $($Stats.OffboardingUpdated)"
            "Exceptions: $($Stats.OffboardingExceptions)"
        )
        New-SimpleSectionHtml -Title "Audit Snapshot" -Items @(
            "Rows loaded: $($Stats.AuditRows)"
            "Compliant rows: $($Stats.CurrentCompliantUsers)"
            "High risk rows: $($Stats.CurrentHighRisk)"
            "Critical rows: $($Stats.CurrentCritical)"
            "Failed rows: $($Stats.CurrentFailed)"
        )
        New-SimpleSectionHtml -Title "Remediation Snapshot" -Items @(
            "Rows loaded: $($Stats.RemediationRows)"
            "Remediated: $($Stats.RemediationRemediated)"
            "Preview only: $($Stats.RemediationPreviewOnly)"
            "Skipped: $($Stats.RemediationSkipped)"
            "Failed: $($Stats.RemediationFailed)"
        )
    ) -join [Environment]::NewLine

    $auditTableBody = foreach ($row in $Stats.TopAuditRows) {
        New-RowHtml -Values @(
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("DisplayName")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("UserPrincipalName")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("Severity")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("CheckName")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("Status"))
        )
    }

    if (-not $auditTableBody) {
        $auditTableBody = '<tr><td colspan="5">No audit rows available.</td></tr>'
    }
    else {
        $auditTableBody = $auditTableBody -join [Environment]::NewLine
    }

    $remediationTableBody = foreach ($row in $Stats.TopRemediationRows) {
        New-RowHtml -Values @(
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("DisplayName")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("UserPrincipalName")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("ActionType")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("TargetValue")),
            (Get-FirstExistingPropertyValue -Row $row -PropertyNames @("Status"))
        )
    }

    if (-not $remediationTableBody) {
        $remediationTableBody = '<tr><td colspan="5">No remediation rows available.</td></tr>'
    }
    else {
        $remediationTableBody = $remediationTableBody -join [Environment]::NewLine
    }

    $auditTableHtml = New-TableHtml -Title "Latest Audit Rows" -Headers @("Display Name","UserPrincipalName","Severity","Check","Status") -BodyHtml $auditTableBody
    $remediationTableHtml = New-TableHtml -Title "Latest Remediation Rows" -Headers @("Display Name","UserPrincipalName","Action","Target","Status") -BodyHtml $remediationTableBody

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Lifecycle Master Dashboard</title>
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

        .tables-grid {
            display: grid;
            gap: var(--gap-sm);
            margin-top: var(--gap-sm);
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

        .footer {
            text-align: center;
            color: #8390a1;
            font-size: 10.5px;
            margin-top: 8px;
        }

        @media (max-width: 1280px) {
            .dashboard-layout { grid-template-columns: 1fr; }
            .sidebar-panel { position: static; }
        }

        @media (max-width: 1080px) {
            .overview-grid { grid-template-columns: 1fr; gap: var(--gap-xs); }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="hero">
            <h1>Microsoft 365 Lifecycle Master Dashboard</h1>
            <p>$([System.Net.WebUtility]::HtmlEncode($headline))</p>
        </div>

        <div class="stats-grid">
            $statCards
        </div>

        <div class="dashboard-layout">
            <div>
                <div class="overview-grid">
                    $runSummaryHtml
                </div>
            </div>

            <div class="sidebar-panel">
                <h2>Data Sources</h2>
                $dataSourcesHtml
            </div>
        </div>

        <div class="tables-grid">
            $auditTableHtml
            $remediationTableHtml
        </div>

        <div class="footer">
            Generated by Build-M365LifecycleMasterDashboard.ps1
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}

# -----------------------------
# Start
# -----------------------------
Write-Log "Starting lifecycle master dashboard build."
Write-Log ("Run ID: {0}" -f $RunId)
Write-Log ("Onboarding CSV path: {0}" -f $OnboardingCsvPath)
Write-Log ("Offboarding CSV path: {0}" -f $OffboardingCsvPath)
Write-Log ("Audit CSV path: {0}" -f $AuditCsvPath)
Write-Log ("Remediation CSV path: {0}" -f $RemediationCsvPath)
Write-Log ("Summary output path: {0}" -f $SummaryPath)
Write-Log ("Log output path: {0}" -f $LogPath)
Write-Log ("Archive retention count: {0}" -f $KeepLatestArchives)

try {
    $onboardingRows  = Import-CsvIfExists -Path $OnboardingCsvPath
    $offboardingRows = Import-CsvIfExists -Path $OffboardingCsvPath
    $auditRows       = Import-CsvIfExists -Path $AuditCsvPath
    $remediationRows = Import-CsvIfExists -Path $RemediationCsvPath

    Write-Log ("Onboarding rows loaded: {0}" -f $onboardingRows.Count)
    Write-Log ("Offboarding rows loaded: {0}" -f $offboardingRows.Count)
    Write-Log ("Audit rows loaded: {0}" -f $auditRows.Count)
    Write-Log ("Remediation rows loaded: {0}" -f $remediationRows.Count)

    $stats = @{
        OnboardingRows         = $onboardingRows.Count
        OffboardingRows        = $offboardingRows.Count
        AuditRows              = $auditRows.Count
        RemediationRows        = $remediationRows.Count

        OnboardingSucceeded    = Get-CountByValues -Rows $onboardingRows -PropertyNames @("Status","Result","Outcome") -MatchValues @("SUCCESS","SUCCEEDED","CREATED","CREATED_SUCCESSFULLY","COMPLETED","Created","Updated")
        OnboardingSkipped      = Get-CountByValues -Rows $onboardingRows -PropertyNames @("Status","Result","Outcome") -MatchValues @("SKIPPED","ALREADY_EXISTS","SKIP","Skipped - Already Compliant")
        OnboardingFailed       = Get-CountByValues -Rows $onboardingRows -PropertyNames @("Status","Result","Outcome") -MatchValues @("FAILED","ERROR","FAIL")

        OffboardingOffboarded  = Get-CountByValues -Rows $offboardingRows -PropertyNames @("Status","Result","Outcome") -MatchValues @("OFFBOARDED","Offboarded")
        OffboardingUpdated     = Get-CountByValues -Rows $offboardingRows -PropertyNames @("Status","Result","Outcome") -MatchValues @("UPDATED","Updated")
        OffboardingExceptions  = Get-CountByValues -Rows $offboardingRows -PropertyNames @("Status","Result","Outcome") -MatchValues @("EXCEPTION","FAILED","ERROR")

        CurrentCompliantUsers  = Get-CountByValues -Rows $auditRows -PropertyNames @("Severity","Status") -MatchValues @("COMPLIANT")
        CurrentHighRisk        = Get-CountByValues -Rows $auditRows -PropertyNames @("Severity") -MatchValues @("HIGH_RISK")
        CurrentCritical        = Get-CountByValues -Rows $auditRows -PropertyNames @("Severity") -MatchValues @("CRITICAL")
        CurrentFailed          = Get-CountByValues -Rows $auditRows -PropertyNames @("Severity","Status") -MatchValues @("FAILED")
        CurrentNonCompliantUsers = (
            (Get-CountByValues -Rows $auditRows -PropertyNames @("Status") -MatchValues @("NON_COMPLIANT")) +
            (Get-CountByValues -Rows $auditRows -PropertyNames @("Severity") -MatchValues @("HIGH_RISK","CRITICAL","WARNING","REVIEW","FAILED"))
        )

        RemediationRemediated  = Get-CountByValues -Rows $remediationRows -PropertyNames @("Status") -MatchValues @("REMEDIATED")
        RemediationPreviewOnly = Get-CountByValues -Rows $remediationRows -PropertyNames @("Status") -MatchValues @("PREVIEW_ONLY","WHATIF_ONLY")
        RemediationSkipped     = Get-CountByValues -Rows $remediationRows -PropertyNames @("Status") -MatchValues @("SKIPPED_ALREADY_COMPLIANT","SKIPPED_UNSUPPORTED","SKIPPED_NOT_FOUND")
        RemediationFailed      = Get-CountByValues -Rows $remediationRows -PropertyNames @("Status") -MatchValues @("FAILED")

        OnboardingLatestTime   = Get-LatestTimestampFromRows -Rows $onboardingRows -PropertyNames @("RunTime","Timestamp","CreatedTime","CreatedDate","Date","AuditTimestamp") -FallbackText "Onboarding CSV does not store row timestamps"
        OffboardingLatestTime  = Get-LatestTimestampFromRows -Rows $offboardingRows -PropertyNames @("RunTime","Timestamp","UpdatedTime","Date","AuditTimestamp") -FallbackText "Offboarding CSV does not store row timestamps"
        AuditLatestTime        = Get-LatestTimestampFromRows -Rows $auditRows -PropertyNames @("AuditTimestamp","Timestamp","Date")
        RemediationLatestTime  = Get-LatestTimestampFromRows -Rows $remediationRows -PropertyNames @("RemediationTime","Timestamp","Date")

        TopAuditRows           = Get-TopRows -Rows $auditRows -Take 8
        TopRemediationRows     = Get-TopRows -Rows $remediationRows -Take 8
        RunDuration            = 0
    }

    $EndTimeForHtml  = Get-Date
    $DurationForHtml = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTimeForHtml).TotalSeconds, 2)
    $stats.RunDuration = $DurationForHtml

    New-HtmlDashboard `
        -OutputPath $SummaryPath `
        -Stats $stats `
        -OnboardingCsvPath $OnboardingCsvPath `
        -OffboardingCsvPath $OffboardingCsvPath `
        -AuditCsvPath $AuditCsvPath `
        -RemediationCsvPath $RemediationCsvPath

    Write-Log ("Lifecycle master dashboard generated successfully: {0}" -f $SummaryPath)
}
catch {
    Write-Log ("Script failed. Error: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
finally {
    try {
        Remove-OldArchives -FolderPath $SummaryFolder -BaseName $SummaryBaseName -Extension ".html" -KeepCount $KeepLatestArchives
        Remove-OldArchives -FolderPath $LogFolder     -BaseName $BaseName        -Extension ".log"  -KeepCount $KeepLatestArchives
    }
    catch {
        Write-Host ("Archive cleanup warning: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        if (Test-Path $SummaryPath) {
            Copy-Item -Path $SummaryPath -Destination $LatestSummaryPath -Force -ErrorAction Stop
            Write-Log ("Latest HTML dashboard updated: {0}" -f $LatestSummaryPath)
        }
        else {
            Write-Log ("Skipped latest HTML dashboard update because the timestamped dashboard file was not created: {0}" -f $SummaryPath) "WARN"
        }
    }
    catch {
        Write-Host ("Could not create latest dashboard copy: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        if (Test-Path $LogPath) {
            Copy-Item -Path $LogPath -Destination $LatestLogPath -Force -ErrorAction Stop
            Write-Host ("Latest log copy updated: {0}" -f $LatestLogPath) -ForegroundColor Gray
        }
    }
    catch {
        Write-Host ("Could not create latest log copy: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

$EndTime  = Get-Date
$Duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalSeconds, 2)

Write-Host ""
Write-Host "========== SUMMARY ==========" -ForegroundColor Cyan
Write-Host ("Run ID: {0}" -f $RunId)
Write-Host ("Onboarding rows loaded: {0}" -f $onboardingRows.Count)
Write-Host ("Offboarding rows loaded: {0}" -f $offboardingRows.Count)
Write-Host ("Audit rows loaded: {0}" -f $auditRows.Count)
Write-Host ("Remediation rows loaded: {0}" -f $remediationRows.Count)
Write-Host ("Run duration (seconds): {0}" -f $Duration)
Write-Host ("Timestamped HTML dashboard saved to: {0}" -f $SummaryPath)
Write-Host ("Latest HTML dashboard saved to: {0}" -f $LatestSummaryPath)
Write-Host ("Timestamped log saved to: {0}" -f $LogPath)
Write-Host ("Latest log saved to: {0}" -f $LatestLogPath)
Write-Host "=============================" -ForegroundColor Cyan