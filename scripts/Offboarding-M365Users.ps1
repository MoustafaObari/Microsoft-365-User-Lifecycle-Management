# ============================================
# Script: Offboarding-M365Users.ps1
# Purpose: Bulk offboard Microsoft 365 users from CSV,
#          revoke access, remove licenses/groups, and generate
#          CSV, log, and HTML summary outputs.
# Author: Moustafa Obari
# Project: Microsoft 365 User Lifecycle Management and Administration
# Recommended Runtime: PowerShell 7
# ============================================

param(
    [string]$CsvPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "inputs\OffboardingUsers.csv"),
    [int]$KeepLatestArchives = 7,
    [string[]]$ProtectedGroups = @(
        "Global Administrator",
        "Company Administrator",
        "Helpdesk Administrator",
        "User Administrator"
    )
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
$RunId     = "OFFBOARD-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmss")

$BaseName        = "M365-Offboarding"
$SummaryBaseName = "M365-OffboardingSummary"

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

# -----------------------------
# Runtime / counters
# -----------------------------
$StartTime              = Get-Date
$offboardedCount        = 0
$updatedCount           = 0
$alreadyOffboardedCount = 0
$skippedNotFoundCount   = 0
$noChangeNeededCount    = 0
$partialCount           = 0
$failedCount            = 0
$warnCount              = 0
$cleanupRemoved         = 0
$results                = @()
$csvUsers               = @()

$accountsBlockedCount      = 0
$usersWithLicensesRemoved  = 0
$totalLicensesRemovedCount = 0
$groupMembershipsRemoved   = 0
$passwordsResetCount       = 0
$protectedGroupSkips       = 0

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
    param([object]$Text)

    if ($null -eq $Text) { return "" }

    $value = [string]$Text
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }

    return (($value -replace '\s+', ' ').Trim())
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

function Convert-ListToText {
    param([string[]]$Items)

    $filtered = @($Items | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($filtered.Count -eq 0) {
        return "None"
    }

    return ($filtered -join "; ")
}

function Add-ResultRow {
    param(
        [string]$DisplayName,
        [string]$UserPrincipalName,
        [string]$Status,
        [string]$StatusReason,
        [string]$BlockSignInResult,
        [string]$LicenseResult,
        [string]$GroupResult,
        [string]$PasswordResult,
        [string]$RequestedActions,
        [string]$CompletedActions,
        [string]$AlreadyCompleteActions,
        [string]$FailedActions,
        [string]$CoreStateBefore,
        [string]$CoreStateAfter,
        [string]$ChangeSummary,
        [string]$Notes
    )

    $script:results += [PSCustomObject]@{
        DisplayName            = $DisplayName
        UserPrincipalName      = $UserPrincipalName
        Status                 = $Status
        StatusReason           = $StatusReason
        BlockSignInResult      = $BlockSignInResult
        LicenseResult          = $LicenseResult
        GroupResult            = $GroupResult
        PasswordResult         = $PasswordResult
        RequestedActions       = $RequestedActions
        CompletedActions       = $CompletedActions
        AlreadyCompleteActions = $AlreadyCompleteActions
        FailedActions          = $FailedActions
        CoreStateBefore        = $CoreStateBefore
        CoreStateAfter         = $CoreStateAfter
        ChangeSummary          = $ChangeSummary
        Notes                  = $Notes
    }
}

function Get-UserGroupMemberships {
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    $groups = @()

    try {
        $memberOf = Get-MgUserMemberOf -UserId $UserId -All -ErrorAction Stop
        foreach ($item in $memberOf) {
            $odataType = $item.AdditionalProperties.'@odata.type'
            if ($odataType -eq '#microsoft.graph.group') {
                $groupName = Normalize-Text $item.AdditionalProperties.displayName
                $groupId   = Normalize-Text $item.Id

                if (-not [string]::IsNullOrWhiteSpace($groupName) -and -not [string]::IsNullOrWhiteSpace($groupId)) {
                    $groups += [PSCustomObject]@{
                        Id          = $groupId
                        DisplayName = $groupName
                    }
                }
            }
        }
    }
    catch {
        Write-Log ("Could not retrieve group memberships for user id {0}. Error: {1}" -f $UserId, $_.Exception.Message) "WARN"
    }

    return @($groups | Sort-Object DisplayName -Unique)
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
        return "<li>No offboarding rows recorded.</li>"
    }

    $displayOrder = @(
        "Offboarded",
        "Updated",
        "Already Offboarded",
        "Partial",
        "Skipped - Not Found",
        "No Changes Needed",
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

    $safeName      = Convert-ToHtmlSafe $Row.DisplayName
    $safeUpn       = Convert-ToHtmlSafe $Row.UserPrincipalName
    $safeStatus    = Convert-ToHtmlSafe $Row.Status
    $safeSummary   = Convert-ToHtmlSafe $Row.ChangeSummary
    $safeBefore    = Convert-ToHtmlSafe $Row.CoreStateBefore
    $safeAfter     = Convert-ToHtmlSafe $Row.CoreStateAfter
    $safeCompleted = Convert-ToHtmlSafe $Row.CompletedActions
    $safeAlready   = Convert-ToHtmlSafe $Row.AlreadyCompleteActions
    $safeNotes     = Convert-ToHtmlSafe $Row.Notes

    $badgeClass = switch ($Row.Status) {
        "Offboarded"          { "badge-success"; break }
        "Updated"             { "badge-info"; break }
        "Already Offboarded"  { "badge-neutral"; break }
        "Partial"             { "badge-warning"; break }
        "Failed"              { "badge-danger"; break }
        "Skipped - Not Found" { "badge-warning"; break }
        default               { "badge-neutral" }
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
    <div class="user-card-line"><strong>Summary:</strong> $safeSummary</div>
    <div class="user-card-line"><strong>Before:</strong> $safeBefore</div>
    <div class="user-card-line"><strong>After:</strong> $safeAfter</div>
    <div class="user-card-line"><strong>Completed:</strong> $safeCompleted</div>
    <div class="user-card-line"><strong>Already Complete:</strong> $safeAlready</div>
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

    $content = if ($isEmpty) {
        '<div class="empty-state">No users were recorded in this section for this run.</div>'
    }
    else {
        ($Rows | ForEach-Object { Get-UserCardHtml -Row $_ }) -join [Environment]::NewLine
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
        return '<div class="empty-state">No processed users to display.</div>'
    }

    $body = foreach ($row in $Rows) {
        $safeName    = Convert-ToHtmlSafe $row.DisplayName
        $safeUpn     = Convert-ToHtmlSafe $row.UserPrincipalName
        $safeStatus  = Convert-ToHtmlSafe $row.Status
        $safeBlock   = Convert-ToHtmlSafe $row.BlockSignInResult
        $safeLicense = Convert-ToHtmlSafe $row.LicenseResult
        $safeGroup   = Convert-ToHtmlSafe $row.GroupResult
        $safeChange  = Convert-ToHtmlSafe $row.ChangeSummary

@"
<tr>
    <td>$safeName</td>
    <td>$safeUpn</td>
    <td>$safeStatus</td>
    <td>$safeBlock</td>
    <td>$safeLicense</td>
    <td>$safeGroup</td>
    <td>$safeChange</td>
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
                <th>Block Sign-In</th>
                <th>Licenses</th>
                <th>Groups</th>
                <th>What Changed</th>
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
        [string]$CsvPath,
        [int]$CsvRowsProcessed,
        [int]$Offboarded,
        [int]$Updated,
        [int]$AlreadyOffboarded,
        [int]$SkippedNotFound,
        [int]$NoChangeNeeded,
        [int]$Partial,
        [int]$Warnings,
        [int]$Failed,
        [int]$AccountsBlocked,
        [int]$UsersWithLicensesRemoved,
        [int]$TotalLicensesRemoved,
        [int]$GroupMembershipsRemoved,
        [int]$PasswordsReset,
        [int]$ProtectedGroupSkips,
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
    $headline = "This run processed $CsvRowsProcessed user(s): $Offboarded offboarded, $Updated updated, $AlreadyOffboarded already offboarded."

    $statCards = @(
        New-StatCardHtml -Label "Run ID"                 -Value $RunId
        New-StatCardHtml -Label "CSV Rows Processed"     -Value $CsvRowsProcessed
        New-StatCardHtml -Label "Offboarded"             -Value $Offboarded -AccentClass "accent-success"
        New-StatCardHtml -Label "Updated"                -Value $Updated -AccentClass "accent-info"
        New-StatCardHtml -Label "Already Offboarded"     -Value $AlreadyOffboarded -AccentClass "accent-neutral"
        New-StatCardHtml -Label "Skipped - Not Found"    -Value $SkippedNotFound
        New-StatCardHtml -Label "No Changes Needed"      -Value $NoChangeNeeded
        New-StatCardHtml -Label "Partial"                -Value $Partial -AccentClass "accent-warning"
        New-StatCardHtml -Label "Warnings"               -Value $Warnings -AccentClass "accent-warning"
        New-StatCardHtml -Label "Failed"                 -Value $Failed -AccentClass "accent-danger"
        New-StatCardHtml -Label "Archive Files Removed"  -Value $ArchiveFilesRemoved
        New-StatCardHtml -Label "Run Duration (Seconds)" -Value $RunDurationSeconds
    ) -join [Environment]::NewLine

    $runSettingsHtml = New-ListItemsHtml -Items @(
        "CSV Path: $CsvPath"
        "Archive Retention Count: $KeepLatestArchives"
        "Generated: $generatedAt"
        "Protected Groups: $($ProtectedGroups -join '; ')"
    )

    $actionTotalsHtml = New-ListItemsHtml -Items @(
        "Accounts blocked in this run: $AccountsBlocked"
        "Users with licenses removed in this run: $UsersWithLicensesRemoved"
        "Total license objects removed in this run: $TotalLicensesRemoved"
        "Group memberships removed in this run: $GroupMembershipsRemoved"
        "Passwords reset in this run: $PasswordsReset"
        "Protected group skips in this run: $ProtectedGroupSkips"
    )

    $statusBreakdownHtml = New-StatusBreakdownHtml -Rows $ResultRows

    $statusDefinitionsHtml = New-ListItemsHtml -Items @(
        "Offboarded: Sign-in was blocked in this run and requested core offboarding controls are now in an offboarded state."
        "Updated: The user was already blocked before this run, and this run changed secondary offboarding items such as licenses, groups, or password."
        "Already Offboarded: Requested offboarding controls were already satisfied before the script ran, and no additional changes were needed."
        "Skipped - Not Found: The UserPrincipalName from the CSV was not found in Microsoft 365."
        "No Changes Needed: The row was valid, but none of the requested actions required a change."
        "Partial: One or more requested actions could not be completed, even if some other actions succeeded."
        "Failed: The row could not be processed because of a hard error."
    )

    $updatedMeaningHtml = New-ListItemsHtml -Items @(
        "Updated is for users who were already blocked before this run."
        "These users appear here when this run changed items like licenses, groups, or password."
        "If sign-in was newly blocked in this run, the user is shown as Offboarded instead."
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

    $offboardedRows = @($ResultRows | Where-Object { $_.Status -eq "Offboarded" })
    $updatedRows    = @($ResultRows | Where-Object { $_.Status -eq "Updated" })
    $exceptionRows  = @($ResultRows | Where-Object { $_.Status -in @("Partial", "Failed", "Skipped - Not Found") })

    $offboardedSectionHtml   = New-UserSectionHtml -Title "Offboarded Users" -Rows $offboardedRows
    $updatedSectionHtml      = New-UserSectionHtml -Title "Updated Users" -Rows $updatedRows
    $exceptionSectionHtml    = New-UserSectionHtml -Title "Exception Users" -Rows $exceptionRows
    $processedUsersTableHtml = New-ProcessedRowsTableHtml -Rows $ResultRows

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Offboarding Summary</title>
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
            grid-column: span 4;
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
            <h1>Microsoft 365 Offboarding Summary</h1>
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
                        <h2>Action Totals</h2>
                        <ul>
                            $actionTotalsHtml
                        </ul>
                    </div>

                    <div class="panel">
                        <h2>Status Definitions</h2>
                        <ul>
                            $statusDefinitionsHtml
                        </ul>
                    </div>

                    <div class="panel">
                        <h2>Updated Meaning</h2>
                        <ul>
                            $updatedMeaningHtml
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
            $offboardedSectionHtml
            $updatedSectionHtml
            $exceptionSectionHtml

            <div class="wide-panel table-panel">
                <h2>Processed Users</h2>
                $processedUsersTableHtml
            </div>
        </div>

        <div class="footer">
            Generated by Offboarding-M365Users.ps1
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

Write-Log "Starting bulk offboarding script."
Write-Log ("Run ID: {0}" -f $RunId)
Write-Log ("CSV path: {0}" -f $CsvPath)
Write-Log ("Result output path: {0}" -f $ResultPath)
Write-Log ("Log output path: {0}" -f $LogPath)
Write-Log ("Summary output path: {0}" -f $SummaryPath)
Write-Log ("Archive retention count: {0}" -f $KeepLatestArchives)
Write-Log ("Protected groups: {0}" -f ($ProtectedGroups -join "; "))

try {
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Group.Read.All","GroupMember.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All" -NoWelcome
    Write-Log "Connected to Microsoft Graph successfully."

    Write-Log "Importing CSV input..."
    $csvUsers = Import-Csv -Path $CsvPath
    Write-Log ("Imported {0} CSV rows." -f $csvUsers.Count)

    foreach ($row in $csvUsers) {
        $displayName      = Normalize-Text $row.DisplayName
        $upn              = Normalize-Text $row.UserPrincipalName
        $blockSignIn      = Normalize-Text $row.BlockSignIn
        $removeLicenses   = Normalize-Text $row.RemoveLicenses
        $removeFromGroups = Normalize-Text $row.RemoveFromGroups
        $resetPassword    = Normalize-Text $row.ResetPassword
        $newPassword      = [string]$row.NewPassword
        $inputNotes       = Normalize-Text $row.Notes

        Write-Log ("Processing CSV user: {0}" -f $upn)

        try {
            if ([string]::IsNullOrWhiteSpace($upn)) {
                $failedCount++
                Write-Log "Missing required UserPrincipalName in CSV row." "ERROR"

                Add-ResultRow `
                    -DisplayName $displayName `
                    -UserPrincipalName $upn `
                    -Status "Failed" `
                    -StatusReason "Missing required UserPrincipalName" `
                    -BlockSignInResult "Not Processed" `
                    -LicenseResult "Not Processed" `
                    -GroupResult "Not Processed" `
                    -PasswordResult "Not Processed" `
                    -RequestedActions "None" `
                    -CompletedActions "None" `
                    -AlreadyCompleteActions "None" `
                    -FailedActions "Missing required UserPrincipalName" `
                    -CoreStateBefore "Unknown" `
                    -CoreStateAfter "Unknown" `
                    -ChangeSummary "Row failed validation" `
                    -Notes "Missing required UserPrincipalName"

                continue
            }

            $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -Property @(
                "Id",
                "DisplayName",
                "UserPrincipalName",
                "AccountEnabled",
                "AssignedLicenses"
            )

            if (-not $user) {
                $skippedNotFoundCount++
                Write-Log ("User not found, skipping: {0}" -f $upn) "WARN"

                Add-ResultRow `
                    -DisplayName $displayName `
                    -UserPrincipalName $upn `
                    -Status "Skipped - Not Found" `
                    -StatusReason "UserPrincipalName was not found in Microsoft 365" `
                    -BlockSignInResult "Not Processed" `
                    -LicenseResult "Not Processed" `
                    -GroupResult "Not Processed" `
                    -PasswordResult "Not Processed" `
                    -RequestedActions "None" `
                    -CompletedActions "None" `
                    -AlreadyCompleteActions "None" `
                    -FailedActions "None" `
                    -CoreStateBefore "Unknown" `
                    -CoreStateAfter "Unknown" `
                    -ChangeSummary "User not found in tenant" `
                    -Notes "User not found"

                continue
            }

            $user = Get-MgUser -UserId $user.Id -Property @(
                "Id",
                "DisplayName",
                "UserPrincipalName",
                "AccountEnabled",
                "AssignedLicenses"
            ) -ErrorAction Stop

            $currentLicenseCount = 0
            if ($user.AssignedLicenses) {
                $currentLicenseCount = @($user.AssignedLicenses).Count
            }

            Write-Log ("Current state for {0}: AccountEnabled={1}; AssignedLicenses={2}" -f $upn, $user.AccountEnabled, $currentLicenseCount)

            $coreStateBefore = "Blocked={0}; Licensed={1}" -f ((-not $user.AccountEnabled)), ($currentLicenseCount -gt 0)

            $notes = @()
            if (-not [string]::IsNullOrWhiteSpace($inputNotes)) {
                $notes += $inputNotes
            }

            $requestedActions       = @()
            $completedActions       = @()
            $alreadyCompleteActions = @()
            $failedActions          = @()

            $blockResult    = "Not Requested"
            $licenseResult  = "Not Requested"
            $groupResult    = "Not Requested"
            $passwordResult = "Not Requested"

            $actionsPerformed        = 0
            $partialIssues           = 0
            $coreTargetsRequested    = 0
            $coreTargetsMetBefore    = 0
            $coreTargetsCompletedNow = 0

            $signInBlockedThisRun    = $false
            $nonSignInChangeThisRun  = $false

            $finalBlocked  = -not $user.AccountEnabled
            $finalLicensed = ($currentLicenseCount -gt 0)

            # -----------------------------
            # Reset password FIRST
            # -----------------------------
            if ($resetPassword -eq "Yes") {
                $requestedActions += "Reset password"

                if (-not [string]::IsNullOrWhiteSpace($newPassword)) {
                    try {
                        Update-MgUser -UserId $user.Id -PasswordProfile @{
                            ForceChangePasswordNextSignIn = $false
                            Password = $newPassword
                        } -ErrorAction Stop | Out-Null

                        Write-Log ("Password reset for {0}" -f $upn)
                        $notes += "Password reset"
                        $passwordResult = "Reset"
                        $completedActions += "Password reset"
                        $actionsPerformed++
                        $passwordsResetCount++
                        $nonSignInChangeThisRun = $true
                    }
                    catch {
                        $partialIssues++
                        $passwordResult = "Failed"
                        $failedActions += "Password reset failed"
                        Write-Log ("Password reset failed for {0}. Error: {1}" -f $upn, $_.Exception.Message) "WARN"
                        $notes += ("Password reset failed: {0}" -f $_.Exception.Message)
                    }
                }
                else {
                    $partialIssues++
                    $passwordResult = "Requested But Missing NewPassword"
                    $failedActions += "Password reset requested but NewPassword missing"
                    Write-Log ("Password reset requested but NewPassword is blank for {0}" -f $upn) "WARN"
                    $notes += "Password reset requested but NewPassword missing"
                }
            }

            # -----------------------------
            # Block sign-in
            # -----------------------------
            if ($blockSignIn -eq "Yes") {
                $requestedActions += "Block sign-in"
                $coreTargetsRequested++

                if ($user.AccountEnabled -eq $true) {
                    Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop | Out-Null
                    Write-Log ("Blocked sign-in for {0}" -f $upn)
                    $notes += "Sign-in blocked"
                    $blockResult = "Blocked"
                    $completedActions += "Sign-in blocked"
                    $actionsPerformed++
                    $coreTargetsCompletedNow++
                    $accountsBlockedCount++
                    $finalBlocked = $true
                    $signInBlockedThisRun = $true
                }
                else {
                    $notes += "Sign-in already blocked"
                    $blockResult = "Already Blocked"
                    $alreadyCompleteActions += "Sign-in already blocked"
                    $coreTargetsMetBefore++
                    $finalBlocked = $true
                }
            }

            # -----------------------------
            # Remove licenses
            # -----------------------------
            if ($removeLicenses -eq "Yes") {
                $requestedActions += "Remove licenses"
                $coreTargetsRequested++

                $assignedLicenseIds = @()

                if ($user.AssignedLicenses) {
                    foreach ($license in $user.AssignedLicenses) {
                        if ($license.SkuId) {
                            $assignedLicenseIds += [string]$license.SkuId
                        }
                    }
                }

                if ($assignedLicenseIds.Count -gt 0) {
                    Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses $assignedLicenseIds -ErrorAction Stop | Out-Null
                    Write-Log ("Removed {0} license(s) from {1}" -f $assignedLicenseIds.Count, $upn)
                    $notes += "Licenses removed"
                    $licenseResult = "Removed ($($assignedLicenseIds.Count))"
                    $completedActions += ("Removed {0} license(s)" -f $assignedLicenseIds.Count)
                    $actionsPerformed++
                    $coreTargetsCompletedNow++
                    $usersWithLicensesRemoved++
                    $totalLicensesRemovedCount += $assignedLicenseIds.Count
                    $finalLicensed = $false
                    $nonSignInChangeThisRun = $true
                }
                else {
                    $notes += "No licenses assigned"
                    $licenseResult = "Already Clear"
                    $alreadyCompleteActions += "No licenses assigned"
                    $coreTargetsMetBefore++
                    $finalLicensed = $false
                }
            }

            # -----------------------------
            # Remove from groups
            # -----------------------------
            if ($removeFromGroups -eq "Yes") {
                $requestedActions += "Remove from groups"
                $currentGroups = Get-UserGroupMemberships -UserId $user.Id
                $removedGroups = 0
                $protectedHits = 0
                $groupFailures = 0

                foreach ($group in $currentGroups) {
                    if ($ProtectedGroups -contains $group.DisplayName) {
                        $protectedHits++
                        $protectedGroupSkips++
                        Write-Log ("Skipped protected group removal for {0}: {1}" -f $upn, $group.DisplayName) "WARN"
                        continue
                    }

                    try {
                        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop | Out-Null
                        $removedGroups++
                    }
                    catch {
                        $partialIssues++
                        $groupFailures++
                        Write-Log ("Failed removing {0} from group {1}. Error: {2}" -f $upn, $group.DisplayName, $_.Exception.Message) "WARN"
                    }
                }

                if ($removedGroups -gt 0) {
                    Write-Log ("Removed {0} group membership(s) from {1}" -f $removedGroups, $upn)
                    $notes += "Groups removed: $removedGroups"
                    $completedActions += ("Removed {0} group membership(s)" -f $removedGroups)
                    $actionsPerformed++
                    $groupMembershipsRemoved += $removedGroups
                    $nonSignInChangeThisRun = $true
                }
                else {
                    $notes += "No removable groups found"
                    $alreadyCompleteActions += "No removable groups found"
                }

                if ($protectedHits -gt 0) {
                    $notes += "Protected groups skipped: $protectedHits"
                }

                if ($groupFailures -gt 0) {
                    $failedActions += ("Failed removing {0} group membership(s)" -f $groupFailures)
                }

                $groupResultParts = @()
                if ($removedGroups -gt 0) {
                    $groupResultParts += ("Removed {0}" -f $removedGroups)
                }
                else {
                    $groupResultParts += "No Removable Groups"
                }

                if ($protectedHits -gt 0) {
                    $groupResultParts += ("Protected Skipped {0}" -f $protectedHits)
                }

                if ($groupFailures -gt 0) {
                    $groupResultParts += ("Failed {0}" -f $groupFailures)
                }

                $groupResult = ($groupResultParts -join " | ")
            }

            # -----------------------------
            # Final status
            # -----------------------------
            $finalStatus  = ""
            $statusReason = ""

            if ($partialIssues -gt 0) {
                $partialCount++
                $finalStatus = "Partial"
                $statusReason = "One or more requested actions could not be completed."
            }
            elseif ($signInBlockedThisRun) {
                $offboardedCount++
                $finalStatus = "Offboarded"
                $statusReason = "Sign-in was blocked in this run and the user is now in an offboarded state."
            }
            elseif ($nonSignInChangeThisRun) {
                $updatedCount++
                $finalStatus = "Updated"
                $statusReason = "The user was already blocked before this run, and this run changed licenses, groups, or password."
            }
            elseif (
                $coreTargetsRequested -gt 0 -and
                $coreTargetsCompletedNow -eq 0 -and
                $coreTargetsMetBefore -eq $coreTargetsRequested -and
                $actionsPerformed -eq 0
            ) {
                $alreadyOffboardedCount++
                $finalStatus = "Already Offboarded"
                $statusReason = "Requested core offboarding actions were already satisfied before this run."
            }
            else {
                $noChangeNeededCount++
                $finalStatus = "No Changes Needed"
                $statusReason = "The row was valid, but none of the requested actions required a change."
            }

            $coreStateAfter = "Blocked={0}; Licensed={1}" -f $finalBlocked, $finalLicensed
            $changeSummaryParts = @()

            if ($completedActions.Count -gt 0) {
                $changeSummaryParts += (Convert-ListToText -Items $completedActions)
            }
            if ($alreadyCompleteActions.Count -gt 0) {
                $changeSummaryParts += ("Already compliant: " + (Convert-ListToText -Items $alreadyCompleteActions))
            }
            if ($failedActions.Count -gt 0) {
                $changeSummaryParts += ("Needs review: " + (Convert-ListToText -Items $failedActions))
            }
            if ($changeSummaryParts.Count -eq 0) {
                $changeSummaryParts += $statusReason
            }

            Add-ResultRow `
                -DisplayName (Normalize-Text $user.DisplayName) `
                -UserPrincipalName $upn `
                -Status $finalStatus `
                -StatusReason $statusReason `
                -BlockSignInResult $blockResult `
                -LicenseResult $licenseResult `
                -GroupResult $groupResult `
                -PasswordResult $passwordResult `
                -RequestedActions (Convert-ListToText -Items $requestedActions) `
                -CompletedActions (Convert-ListToText -Items $completedActions) `
                -AlreadyCompleteActions (Convert-ListToText -Items $alreadyCompleteActions) `
                -FailedActions (Convert-ListToText -Items $failedActions) `
                -CoreStateBefore $coreStateBefore `
                -CoreStateAfter $coreStateAfter `
                -ChangeSummary (Convert-ListToText -Items $changeSummaryParts) `
                -Notes ($notes -join "; ")
        }
        catch {
            $failedCount++
            Write-Log ("Failed processing {0}. Error: {1}" -f $upn, $_.Exception.Message) "ERROR"

            Add-ResultRow `
                -DisplayName $displayName `
                -UserPrincipalName $upn `
                -Status "Failed" `
                -StatusReason "A hard error stopped row processing." `
                -BlockSignInResult "Not Processed" `
                -LicenseResult "Not Processed" `
                -GroupResult "Not Processed" `
                -PasswordResult "Not Processed" `
                -RequestedActions "Unknown" `
                -CompletedActions "None" `
                -AlreadyCompleteActions "None" `
                -FailedActions $_.Exception.Message `
                -CoreStateBefore "Unknown" `
                -CoreStateAfter "Unknown" `
                -ChangeSummary "Row failed due to hard error" `
                -Notes $_.Exception.Message
        }
    }

    $results |
        Sort-Object DisplayName |
        Export-Csv -Path $ResultPath -NoTypeInformation -Encoding UTF8

    Write-Log ("Bulk offboarding result CSV exported successfully: {0}" -f $ResultPath)

    Copy-Item -Path $ResultPath -Destination $LatestResultPath -Force
    Write-Log ("Latest result CSV updated: {0}" -f $LatestResultPath)
}
catch {
    $failedCount++
    Write-Log ("Script failed. Error: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
finally {
    try {
        Copy-Item -Path $LogPath -Destination $LatestLogPath -Force
        Write-Log ("Latest log copy updated: {0}" -f $LatestLogPath)
    }
    catch {
        Write-Host ("Could not create latest log copy: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    $EndTime  = Get-Date
    $Duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalSeconds, 2)

    try {
        Remove-OldArchives -FolderPath $ReportFolder -BaseName $BaseName -Extension ".csv" -KeepCount $KeepLatestArchives
        Remove-OldArchives -FolderPath $LogFolder    -BaseName $BaseName -Extension ".log" -KeepCount $KeepLatestArchives

        $summaryKeepBeforeCreate = [Math]::Max(($KeepLatestArchives - 1), 0)
        Remove-OldArchives -FolderPath $SummaryFolder -BaseName $SummaryBaseName -Extension ".html" -KeepCount $summaryKeepBeforeCreate
    }
    catch {
        Write-Host ("Archive cleanup warning: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        New-HtmlSummary `
            -OutputPath $SummaryPath `
            -RunId $RunId `
            -CsvPath $CsvPath `
            -CsvRowsProcessed $csvUsers.Count `
            -Offboarded $offboardedCount `
            -Updated $updatedCount `
            -AlreadyOffboarded $alreadyOffboardedCount `
            -SkippedNotFound $skippedNotFoundCount `
            -NoChangeNeeded $noChangeNeededCount `
            -Partial $partialCount `
            -Warnings $warnCount `
            -Failed $failedCount `
            -AccountsBlocked $accountsBlockedCount `
            -UsersWithLicensesRemoved $usersWithLicensesRemoved `
            -TotalLicensesRemoved $totalLicensesRemovedCount `
            -GroupMembershipsRemoved $groupMembershipsRemoved `
            -PasswordsReset $passwordsResetCount `
            -ProtectedGroupSkips $protectedGroupSkips `
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
        Write-Log ("Latest HTML summary copy updated: {0}" -f $LatestSummaryPath)
    }
    catch {
        Write-Host ("HTML summary generation warning: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    try {
        Write-Log "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph | Out-Null
        Write-Log "Disconnected successfully."
    }
    catch {
        Write-Log ("Disconnect warning: {0}" -f $_.Exception.Message) "WARN"
    }
}

$EndTime = Get-Date
$Duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalSeconds, 2)

Write-Host ""
Write-Host "========== SUMMARY ==========" -ForegroundColor Cyan
Write-Host ("Run ID: {0}" -f $RunId)
Write-Host ("CSV rows processed: {0}" -f $csvUsers.Count)
Write-Host ("Offboarded: {0}" -f $offboardedCount)
Write-Host ("Updated: {0}" -f $updatedCount)
Write-Host ("Already Offboarded: {0}" -f $alreadyOffboardedCount)
Write-Host ("Skipped - Not Found: {0}" -f $skippedNotFoundCount)
Write-Host ("No Changes Needed: {0}" -f $noChangeNeededCount)
Write-Host ("Partial: {0}" -f $partialCount)
Write-Host ("Warnings: {0}" -f $warnCount)
Write-Host ("Failed: {0}" -f $failedCount)
Write-Host ("Accounts blocked in this run: {0}" -f $accountsBlockedCount)
Write-Host ("Users with licenses removed in this run: {0}" -f $usersWithLicensesRemoved)
Write-Host ("Total license objects removed in this run: {0}" -f $totalLicensesRemovedCount)
Write-Host ("Group memberships removed in this run: {0}" -f $groupMembershipsRemoved)
Write-Host ("Passwords reset in this run: {0}" -f $passwordsResetCount)
Write-Host ("Protected group skips in this run: {0}" -f $protectedGroupSkips)
Write-Host ("Archive files removed: {0}" -f $cleanupRemoved)
Write-Host ("Run duration (seconds): {0}" -f $Duration)
Write-Host ("Timestamped result saved to: {0}" -f $ResultPath)
Write-Host ("Timestamped log saved to: {0}" -f $LogPath)
Write-Host ("Latest result saved to: {0}" -f $LatestResultPath)
Write-Host ("Latest log saved to: {0}" -f $LatestLogPath)
Write-Host ("Timestamped HTML summary saved to: {0}" -f $SummaryPath)
Write-Host ("Latest HTML summary saved to: {0}" -f $LatestSummaryPath)
Write-Host "=============================" -ForegroundColor Cyan