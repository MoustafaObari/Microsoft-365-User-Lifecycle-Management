# Microsoft 365 User Lifecycle Management

<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&height=180&color=0:60A5FA,50:A78BFA,100:F9A8D4&text=Microsoft%20365%20User%20Lifecycle%20Management&fontAlignY=35&fontColor=ffffff&fontSize=32&desc=Bulk%20onboarding%2C%20offboarding%2C%20license%20management%2C%20compliance%20auditing%2C%20remediation%2C%20and%20dashboards&descAlignY=58&descSize=13" alt="Microsoft 365 User Lifecycle Management Banner" />
</p>

<h1 align="center">Microsoft 365 User Lifecycle Management</h1>

<p align="center">
  <b>A Microsoft 365 and Entra ID administration project for user onboarding, offboarding, license management, compliance auditing, remediation, and reporting.</b>
</p>

<p align="center">
  <a href="https://github.com/MoustafaObari/Microsoft-365-User-Lifecycle-Management">
    <img src="https://img.shields.io/badge/View%20on-GitHub-111827?style=for-the-badge&logo=github&logoColor=white" alt="View on GitHub" />
  </a>
</p>

<p align="center">
  <a href="#overview"><img src="https://img.shields.io/badge/Overview-111827?style=for-the-badge" alt="Overview" /></a>
  <a href="#use-cases"><img src="https://img.shields.io/badge/Use%20Cases-111827?style=for-the-badge" alt="Use Cases" /></a>
  <a href="#features"><img src="https://img.shields.io/badge/Features-111827?style=for-the-badge" alt="Features" /></a>
  <a href="#workflow"><img src="https://img.shields.io/badge/Workflow-111827?style=for-the-badge" alt="Workflow" /></a>
  <a href="#scripts"><img src="https://img.shields.io/badge/Scripts-111827?style=for-the-badge" alt="Scripts" /></a>
  <a href="#setup"><img src="https://img.shields.io/badge/Setup-111827?style=for-the-badge" alt="Setup" /></a>
  <a href="#screenshots"><img src="https://img.shields.io/badge/Screenshots-111827?style=for-the-badge" alt="Screenshots" /></a>
  <a href="#security-design"><img src="https://img.shields.io/badge/Security%20Design-111827?style=for-the-badge" alt="Security Design" /></a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-111827?style=flat-square&logo=powershell&logoColor=white" alt="PowerShell" />
  <img src="https://img.shields.io/badge/Microsoft%20365-2563EB?style=flat-square&logo=microsoft&logoColor=white" alt="Microsoft 365" />
  <img src="https://img.shields.io/badge/Entra%20ID-111827?style=flat-square&logo=microsoftazure&logoColor=white" alt="Entra ID" />
  <img src="https://img.shields.io/badge/Microsoft%20Graph-2563EB?style=flat-square" alt="Microsoft Graph" />
  <img src="https://img.shields.io/badge/CSV%20Automation-111827?style=flat-square" alt="CSV Automation" />
  <img src="https://img.shields.io/badge/HTML%20Dashboards-0F172A?style=flat-square" alt="HTML Dashboards" />
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Version-1.0.0-111827?style=flat-square" alt="Version" />
  <img src="https://img.shields.io/badge/Status-Completed-16A34A?style=flat-square" alt="Status" />
  <img src="https://img.shields.io/badge/Focus-M365%20Admin%20%7C%20IT%20Support-2563EB?style=flat-square" alt="Focus" />
  <img src="https://img.shields.io/badge/License-MIT-16A34A?style=flat-square" alt="License" />
</p>

---

<a id="overview"></a>

## 🌸 Overview

**Microsoft 365 User Lifecycle Management** is a cloud administration and automation project built around common IT Support, Service Desk, Microsoft 365 Administrator, and Junior SysAdmin workflows.

The project combines Microsoft 365 Admin Center tasks with PowerShell automation through Microsoft Graph. It demonstrates how user lifecycle work can be handled from beginning to end: onboarding users, assigning licenses, managing groups, exporting inventory, offboarding users, auditing lifecycle compliance, remediating mismatches, and generating HTML dashboards.

This project was built around a demo tenant environment named **MO's Demo LAB** and uses fictional lab users such as Barry Allen, Diana Prince, Hal Jordan, Victor Stone, Tony Stark, and Bruce Wayne.

The goal was not only to create scripts. The goal was to show an operational support process:

> Define standards, automate repeatable user actions, audit the tenant, preview remediation safely, apply fixes, and report the results clearly.

---

<a id="use-cases"></a>

## 💼 Use Cases

| Use Case | What This Project Demonstrates |
|---|---|
| **IT Support / Help Desk** | Password resets, sign-in blocking, account recovery, license checks, and user lookup workflows. |
| **Microsoft 365 Administration** | User creation, profile updates, license assignment, group membership, and tenant inventory. |
| **Identity & Access Management** | Lifecycle standards based on department, account status, groups, licenses, and usage location. |
| **Offboarding / Access Revocation** | Block sign-in, remove licenses, remove group memberships, and document offboarding results. |
| **Compliance Auditing** | Compare users against expected lifecycle standards and identify mismatches. |
| **Safe Remediation** | Preview remediation actions before applying changes to users, groups, licenses, or sign-in status. |
| **Reporting & Documentation** | Generate CSV, log, and HTML summaries for review and operational handoff. |

---

<a id="features"></a>

## ✨ Features

- ✅ Manual Microsoft 365 Admin Center password reset workflow.
- ✅ Manual block sign-in and unblock account workflow.
- ✅ User license review from the Microsoft 365 Admin Center.
- ✅ Bulk user onboarding from CSV.
- ✅ User profile enrichment with department, title, location, phones, and usage location.
- ✅ License assignment using Microsoft Graph PowerShell.
- ✅ Group membership assignment using Microsoft Graph PowerShell.
- ✅ User inventory export with license and group information.
- ✅ Bulk offboarding with sign-in blocking, license removal, and group removal.
- ✅ Lifecycle standards defined through CSV.
- ✅ Read-only lifecycle audit to detect mismatches.
- ✅ Remediation script with PreviewMode for safe dry runs.
- ✅ Final remediation run to fix audit findings.
- ✅ Master dashboard that consolidates onboarding, offboarding, audit, and remediation output.
- ✅ HTML summaries and CSV reports for clean documentation.

---

<a id="tech-stack"></a>

## 🧰 Tech Stack

| Layer | Technology |
|---|---|
| Cloud Platform | Microsoft 365 |
| Identity Platform | Microsoft Entra ID |
| Admin Portal | Microsoft 365 Admin Center |
| Automation | PowerShell 7 |
| API / SDK | Microsoft Graph PowerShell |
| Input Format | CSV |
| Output Format | CSV, LOG, HTML |
| Reporting | HTML / CSS dashboards |
| Documentation | Markdown + Screenshots |

---

<a id="workflow"></a>

## 🏗️ Lifecycle Workflow

```text
CSV Inputs
├── NewUsers.csv
├── OffboardingUsers.csv
└── M365-LifecycleStandards.csv
        │
        ▼
PowerShell Automation
├── Onboarding-New-M365Users.ps1
├── Export-M365Users.ps1
├── Offboarding-M365Users.ps1
├── Audit-M365LifecycleCompliance.ps1
├── Remediate-M365LifecycleFindings.ps1
└── Build-M365LifecycleMasterDashboard.ps1
        │
        ▼
Generated Evidence
├── CSV reports
├── timestamped logs
├── HTML summaries
└── master lifecycle dashboard
```

| Phase | Purpose |
|---|---|
| **Onboarding** | Create or reconcile users, assign profile details, licenses, and groups. |
| **Export** | Build a tenant inventory with user, license, group, and status details. |
| **Offboarding** | Revoke access by blocking sign-in, removing licenses, and removing group memberships. |
| **Audit** | Compare users against department lifecycle standards. |
| **Remediation Preview** | Show what would be changed before making updates. |
| **Remediation Apply** | Fix confirmed lifecycle mismatches. |
| **Master Dashboard** | Consolidate the latest results into one executive-style view. |

---

<a id="repository-structure"></a>

## 📁 Repository Structure

```text
Microsoft-365-User-Lifecycle-Management/
│
├── README.md
├── LICENSE
├── .gitignore
│
├── scripts/
│   ├── Onboarding-New-M365Users.ps1
│   ├── Export-M365Users.ps1
│   ├── Offboarding-M365Users.ps1
│   ├── Audit-M365LifecycleCompliance.ps1
│   ├── Remediate-M365LifecycleFindings.ps1
│   └── Build-M365LifecycleMasterDashboard.ps1
│
├── inputs/
│   ├── NewUsers.csv
│   ├── OffboardingUsers.csv
│   └── M365-LifecycleStandards.csv
│
├── screenshots/
│   ├── 01-admin-center-user-profile-before-reset.png
│   ├── 02-admin-center-password-reset-window.png
│   ├── 03-admin-center-password-reset-confirmation.png
│   ├── 04-admin-center-block-sign-in-action.png
│   ├── 05-admin-center-blocked-sign-in-status.png
│   ├── 06-admin-center-unblocked-account.png
│   ├── 07-admin-center-license-tab.png
│   ├── 08-users-assigned-licenses.png
│   ├── 09-bulk-onboarding-summary.png
│   ├── 10-lifecycle-master-dashboard.png
│   ├── 11-lifecycle-remediation-summary-fixed.png
│   ├── 12-lifecycle-audit-summary-after-fixes.png
│   ├── 13-lifecycle-audit-summary-before-fixes.png
│   ├── 14-remediation-powershell-output-fixed.png
│   └── 15-remediation-summary-preview-mode.png
│
├── summary/
│   └── Sample HTML summary outputs
│
├── reports/
│   └── Generated CSV reports at runtime
│
├── logs/
│   └── Generated logs at runtime
│
├── assets/
│   └── Microsoft365 & Entra ID Demo Video.mp4
│
└── docs/
    └── screenshot-checklist.md
```

---

<a id="scripts"></a>

## ⚙️ PowerShell Scripts

| Script | Purpose |
|---|---|
| `Onboarding-New-M365Users.ps1` | Bulk creates or reconciles users from CSV, applies profile details, assigns licenses, and adds group memberships. |
| `Export-M365Users.ps1` | Exports Microsoft 365 users with account status, department, licenses, groups, phones, and location details. |
| `Offboarding-M365Users.ps1` | Blocks sign-in, removes licenses, removes group memberships, and generates offboarding evidence. |
| `Audit-M365LifecycleCompliance.ps1` | Performs a read-only audit against lifecycle standards defined in CSV. |
| `Remediate-M365LifecycleFindings.ps1` | Applies remediation for audit findings and supports `-PreviewMode` for safe review before changes. |
| `Build-M365LifecycleMasterDashboard.ps1` | Builds a master dashboard from the latest onboarding, offboarding, audit, and remediation outputs. |

### Microsoft Graph Scopes Used

| Workflow | Example Scopes |
|---|---|
| Read-only export / audit | `User.Read.All`, `Group.Read.All`, `Directory.Read.All`, `Organization.Read.All` |
| Onboarding / offboarding | `User.ReadWrite.All`, `GroupMember.ReadWrite.All`, `Directory.ReadWrite.All`, `Organization.Read.All` |
| Remediation | `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.ReadWrite.All`, `Organization.Read.All` |

---

<a id="setup"></a>

## 🚀 Setup

### 1️⃣ Clone the Repository

```powershell
git clone https://github.com/MoustafaObari/Microsoft-365-User-Lifecycle-Management.git
cd Microsoft-365-User-Lifecycle-Management
```

### 2️⃣ Install Microsoft Graph PowerShell Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### 3️⃣ Review CSV Inputs

The project uses three CSV files:

```text
inputs/NewUsers.csv
inputs/OffboardingUsers.csv
inputs/M365-LifecycleStandards.csv
```

Before running the scripts in another tenant, update:

- Tenant domain
- User names
- Departments
- Usage location
- Group names
- License SKU part numbers
- Temporary onboarding passwords

> The included CSVs are lab samples only. They are not production data.

### 4️⃣ Run Onboarding

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Onboarding-New-M365Users.ps1
```

### 5️⃣ Export Tenant Users

```powershell
.\scripts\Export-M365Users.ps1
```

Optional filters:

```powershell
.\scripts\Export-M365Users.ps1 -EnabledOnly
.\scripts\Export-M365Users.ps1 -LicensedOnly
.\scripts\Export-M365Users.ps1 -DepartmentFilter "IT"
```

### 6️⃣ Run Offboarding

```powershell
.\scripts\Offboarding-M365Users.ps1
```

### 7️⃣ Run Lifecycle Audit

```powershell
.\scripts\Audit-M365LifecycleCompliance.ps1
```

### 8️⃣ Preview Remediation First

```powershell
.\scripts\Remediate-M365LifecycleFindings.ps1 -PreviewMode
```

### 9️⃣ Apply Remediation After Review

```powershell
.\scripts\Remediate-M365LifecycleFindings.ps1
```

### 🔟 Build the Master Dashboard

```powershell
.\scripts\Build-M365LifecycleMasterDashboard.ps1
```

---

<a id="inputs"></a>

## 📄 CSV Inputs

### `NewUsers.csv`

Used for bulk onboarding.

| Column Examples |
|---|
| FirstName |
| LastName |
| DisplayName |
| UserName |
| Department |
| JobTitle |
| OfficeLocation |
| UsageLocation |
| LicenseSku |
| GroupName |
| Password |

### `OffboardingUsers.csv`

Used for controlled offboarding.

| Column Examples |
|---|
| UserPrincipalName |
| DisplayName |
| BlockSignIn |
| RemoveLicenses |
| RemoveFromGroups |
| ResetPassword |
| NewPassword |
| Notes |

### `M365-LifecycleStandards.csv`

Used by the audit and remediation workflow.

| Column Examples |
|---|
| Department |
| RequiredGroups |
| RequiredLicenses |
| ForbiddenGroups |
| ForbiddenLicenses |
| ExpectedUsageLocation |
| ExpectedAccountEnabled |
| Notes |

---

<a id="demo-video"></a>

## 🎥 Demo Video

A demo video is included with this repository and shows the Microsoft 365 Admin Center workflows, PowerShell automation, CSV-driven onboarding and offboarding, lifecycle auditing, remediation preview, remediation fixes, and the final master dashboard.

🎬 **Watch the demo:** [Microsoft 365 & Entra ID Demo Video](assets/Microsoft365%20%26%20Entra%20ID%20Demo%20Video.mp4)

> Note: The video was compressed to remain under GitHub's regular file-size limit. GitHub may still display a large-file warning because the file is over 50 MiB, but it is below 100 MiB.

---

<a id="screenshots"></a>

## 🖼️ Screenshots

### Microsoft 365 Admin Center Workflows

| User Profile | Password Reset | Reset Confirmation |
|---|---|---|
| <img src="screenshots/01-admin-center-user-profile-before-reset.png" width="100%" /> | <img src="screenshots/02-admin-center-password-reset-window.png" width="100%" /> | <img src="screenshots/03-admin-center-password-reset-confirmation.png" width="100%" /> |

| Block Sign-In | Blocked Account | Unblocked Account |
|---|---|---|
| <img src="screenshots/04-admin-center-block-sign-in-action.png" width="100%" /> | <img src="screenshots/05-admin-center-blocked-sign-in-status.png" width="100%" /> | <img src="screenshots/06-admin-center-unblocked-account.png" width="100%" /> |

| License Tab | Users and Assigned Licenses |
|---|---|
| <img src="screenshots/07-admin-center-license-tab.png" width="100%" /> | <img src="screenshots/08-users-assigned-licenses.png" width="100%" /> |

### Automation Reports and Dashboards

| Bulk Onboarding Summary | Master Dashboard |
|---|---|
| <img src="screenshots/09-bulk-onboarding-summary.png" width="100%" /> | <img src="screenshots/10-lifecycle-master-dashboard.png" width="100%" /> |

| Audit Before Fixes | Remediation Preview |
|---|---|
| <img src="screenshots/13-lifecycle-audit-summary-before-fixes.png" width="100%" /> | <img src="screenshots/15-remediation-summary-preview-mode.png" width="100%" /> |

| Remediation PowerShell Output | Remediation Summary After Fixes |
|---|---|
| <img src="screenshots/14-remediation-powershell-output-fixed.png" width="100%" /> | <img src="screenshots/11-lifecycle-remediation-summary-fixed.png" width="100%" /> |

| Audit Summary After Fixes |
|---|
| <img src="screenshots/12-lifecycle-audit-summary-after-fixes.png" width="100%" /> |

---

## 📘 Screenshot Descriptions

| # | Screenshot | Description |
|---:|---|---|
| 1 | User Profile Before Reset | Shows a user profile before the password reset workflow. |
| 2 | Password Reset Window | Shows the Microsoft 365 Admin Center password reset window. |
| 3 | Password Reset Confirmation | Confirms that the password reset action completed. |
| 4 | Block Sign-In Action | Shows the manual workflow for blocking user sign-in. |
| 5 | Blocked Account | Confirms that the account sign-in status is blocked. |
| 6 | Unblocked Account | Confirms the account was restored for sign-in. |
| 7 | License Tab | Shows license assignment visibility for a user. |
| 8 | Users and Licenses | Shows users and assigned licenses in the admin center. |
| 9 | Bulk Onboarding Summary | Shows generated HTML output from the onboarding script. |
| 10 | Master Dashboard | Shows consolidated lifecycle status across project phases. |
| 11 | Remediation Summary After Fixes | Shows remediation results after fixes were applied. |
| 12 | Audit Summary After Fixes | Shows the audit state after remediation. |
| 13 | Audit Summary Before Fixes | Shows lifecycle findings before remediation. |
| 14 | Remediation PowerShell Output | Shows terminal output from remediation execution. |
| 15 | Remediation Preview Mode | Shows safe preview output before applying changes. |

---

<a id="security-design"></a>

## 🔐 Security Design

### Why group creation was not fully automated

The scripts can assign users to existing groups, but the project does not blindly create new access-control groups for every CSV value.

In Microsoft 365 and Entra ID, groups often control access to Teams, SharePoint, applications, shared mailboxes, and sensitive business resources. Automatically creating or modifying access groups without review can cause:

- Naming conflicts
- Duplicate access structures
- Unauthorized access
- License or group assignment mistakes
- Confusing lifecycle ownership
- Accidental privilege escalation

For this project, the safer design was:

- ✅ Use CSV files for repeatable user lifecycle actions.
- ✅ Require known group names and expected standards.
- ✅ Skip or report missing groups instead of silently creating access structures.
- ✅ Use audit and remediation workflows to identify and fix mismatches.
- ✅ Preview remediation before applying changes.

This demonstrates a balance between automation speed and access-control judgment.

### Password handling note

The sample CSV includes lab passwords for demo onboarding. In a production environment, initial passwords should be generated securely, delivered through an approved process, and managed according to organizational policy.

---

## 🧠 What I Learned

This project strengthened my understanding of:

- Microsoft 365 Admin Center user management.
- Microsoft Entra ID identity lifecycle workflows.
- Microsoft Graph PowerShell authentication and permissions.
- Bulk user onboarding and offboarding patterns.
- License assignment and removal.
- Group membership management.
- Read-only compliance auditing.
- Preview-first remediation design.
- HTML reporting for operational visibility.
- Documentation practices for IT support and administration projects.

---

## 🧾 Resume Summary Version

Built a Microsoft 365 and Entra ID user lifecycle management project using PowerShell and Microsoft Graph. Automated bulk onboarding, user export, offboarding, compliance auditing, remediation, and dashboard generation using CSV inputs, generated logs, CSV reports, and HTML summaries. Practiced Microsoft 365 Admin Center workflows including password reset, block sign-in, unblock account, and license review.

---

## 🎯 Skills Demonstrated

| Category | Skills |
|---|---|
| Microsoft 365 Administration | User management, password reset, sign-in blocking, license review |
| Entra ID / Identity | Users, groups, account status, usage location, lifecycle standards |
| PowerShell Automation | Microsoft Graph PowerShell, CSV-driven scripts, modular reporting |
| Access Management | Group assignment, group removal, least-privilege review |
| License Management | License assignment, removal, SKU mapping |
| Compliance / Audit | Standards CSV, mismatch detection, before/after audit validation |
| Remediation | Preview mode, controlled fixes, reporting |
| Documentation | Screenshots, README, HTML summaries, demo planning |

---

## 🧩 Planned Enhancements

- Add Graph API app-only authentication option for scheduled runs.
- Add Teams or email notifications after audit and remediation runs.
- Add SharePoint/OneDrive checks for offboarding completeness.
- Add manager approval fields to remediation workflows.
- Add an interactive dashboard index page for all generated summaries.
- Add GitHub Actions documentation for running validation checks.
- Extend the project with Intune device compliance checks.

---

<a id="developer"></a>

## 👨‍💻 Developer

**Moustafa Obari**  
IT Support Specialist • PowerShell Automation • Microsoft 365 / Entra / Intune  
📍 Toronto, Canada

- 🔗 [GitHub](https://github.com/MoustafaObari)
- 🔗 [LinkedIn](https://www.linkedin.com/in/moustafaobari)
- ✉️ moustafaobari@gmail.com

---

<p align="center">
  <img src="https://img.shields.io/badge/Profile%20Views-Portfolio%20Project-111827?style=flat-square" alt="Profile Views" />
</p>

<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&height=110&section=footer&color=0:F9A8D4,50:A78BFA,100:60A5FA&text=Turning%20Microsoft%20365%20lifecycle%20tasks%20into%20auditable%2C%20repeatable%2C%20and%20well-documented%20workflows.&fontColor=ffffff&fontSize=13&fontAlignY=70" alt="Footer" />
</p>

<p align="center">
  © 2026 Moustafa Obari — crafted with 💙 PowerShell, Microsoft 365, Entra ID, Markdown, and strong coffee.
</p>

<p align="center">
  <a href="#microsoft-365-user-lifecycle-management">⬆ Back to Top</a>
</p>
