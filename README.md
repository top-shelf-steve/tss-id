# Set Email Authentication Method

`Set-EmailAuthenticationMethod.ps1` registers a Microsoft Entra user's work email address
as their Microsoft Graph `emailAuthenticationMethod`.

> [!IMPORTANT]
> The email authentication method is for self-service password reset (SSPR). It is not an
> MFA sign-in method. Each user can have only one registered email authentication method.

## Features

- Interactive menu when the script is run without targeting parameters
- Single user, multiple users, transitive group members, CSV, and all-user modes
- Uses the directory `mail` value, with configurable UPN fallback
- Optional per-row email override for CSV imports
- Idempotent create/update/unchanged behavior
- Skips guests and disabled accounts by default
- Detailed, timestamped terminal logging and retry handling
- Bulk confirmation, `-WhatIf`, and standard PowerShell `-Confirm` support
- Optional CSV or JSON result export at the end of every run

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- `Microsoft.Graph.Authentication` module
- Delegated Microsoft Graph scopes:
  - `User.Read.All`
  - `UserAuthenticationMethod.ReadWrite.All`
  - `GroupMember.Read.All` when group mode is used
- The signed-in account must have the **Authentication Administrator** or **Privileged
  Authentication Administrator** Microsoft Entra role.

Install the required module if needed:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

## Usage

Run interactively:

```powershell
.\Set-EmailAuthenticationMethod.ps1
```

Target one user:

```powershell
.\Set-EmailAuthenticationMethod.ps1 -UserId alex@contoso.com
```

Preview several users without making changes:

```powershell
.\Set-EmailAuthenticationMethod.ps1 `
    -UserIds user1@contoso.com,user2@contoso.com `
    -WhatIf
```

Target all users in a group, including nested group membership, and export the results:

```powershell
.\Set-EmailAuthenticationMethod.ps1 `
    -GroupId 'Operations' `
    -Force `
    -ExportPath C:\Reports
```

Target all enabled member users in the tenant:

```powershell
.\Set-EmailAuthenticationMethod.ps1 -AllUsers -Force
```

By default, a missing `mail` value falls back to `userPrincipalName`. To require the
directory `mail` property instead:

```powershell
.\Set-EmailAuthenticationMethod.ps1 -GroupId 'Operations' -EmailSource Mail
```

Use `-IncludeGuests` or `-IncludeDisabled` to include accounts normally skipped. Use
`-NoExportPrompt` for unattended execution, or `-ExportPath` to export automatically.

## CSV format

Each row needs one identity column. Supported identity headers, in priority order, are
`UserId`, `UserPrincipalName`, `Identity`, and `Mail`. `EmailAddress` is optional and
overrides the desired address for that row.

```csv
UserPrincipalName,EmailAddress
alex@contoso.com,
sam@contoso.com,samuel@contoso.com
```

Rows that cannot be resolved or do not have a valid desired email are logged as failures
while processing continues for the remaining rows.

## Result export

At the end of a run, the script offers to export a detailed per-user result. Specify a
`.json` filename for JSON; any other extension produces CSV. Supplying a directory creates
a timestamped CSV file inside it. For create and update actions, `ChangeFrom` and `ChangeTo`
show the current and desired SSPR email addresses; a user without an existing method is
shown as `<not registered>`.

For non-interactive runs, either supply `-ExportPath` or use `-NoExportPrompt`.
Exports are intentionally written during `-WhatIf` runs; dry-run mode prevents Microsoft
Graph changes, while still allowing the requested results report to be saved.

---

# Mail-Enabled Security Group Report

`Get-MailEnabledSecurityGroups.ps1` inventories on-premises Active Directory
mail-enabled groups and captures their direct membership relationships. It reports
mail-enabled security groups by default; `-IncludeDistributionGroups` includes both
security and distribution groups.

The default `All` member scope produces three relational CSV datasets:

- `MailEnabledSecurityGroups-*.csv`: one row per source group
- `MailEnabledSecurityGroupMembers-*.csv`: one row per direct group-to-member relationship
- `MailEnabledSecurityGroupNesting-*.csv`: the group-to-group relationship subset

Direct membership is used so nesting topology is preserved for diagramming. The member
report identifies each member's object type and flags whether a nested group is itself a
mail-enabled security group and whether it belongs to the report's source-group set. User
and computer rows also include derived account status, raw `userAccountControl`, and the
raw Exchange recipient attributes `msExchRecipientTypeDetails`,
`msExchRecipientDisplayType`, and `msExchRemoteRecipientType`. Disabled accounts are
retained; they are not automatically excluded because they may represent shared or resource
mailboxes.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- The `ActiveDirectory` PowerShell module from RSAT
- Read access to the relevant on-premises Active Directory objects

## Usage

Capture all direct membership and export all three datasets:

```powershell
.\Get-MailEnabledSecurityGroups.ps1 -ExportPath C:\Reports
```

Include every AD group with a populated `mail` attribute:

```powershell
.\Get-MailEnabledSecurityGroups.ps1 `
    -IncludeDistributionGroups `
    -ExportPath C:\Reports\AllMailGroups
```

In this mode, filenames begin with `MailEnabledGroups-`, and the `GroupCategory` and
`GroupScope` columns distinguish security/distribution and universal/global/domain-local
groups. The `SourceGroupCriteria` column also records the criteria used for the run.

Target one group by exact name, display name, `sAMAccountName`, distinguished name, SID,
or object GUID, and preview its first 15 direct members:

```powershell
.\Get-MailEnabledSecurityGroups.ps1 `
    -GroupIdentity 'Finance Distribution List' `
    -IncludeDistributionGroups `
    -TestMode `
    -ExportPath C:\Reports\FinanceTest
```

Omit `-TestMode` to capture all direct members of the targeted group. Distribution groups
require `-IncludeDistributionGroups`; the script reports a clear error if it is omitted.

Capture only group nesting for a smaller diagram-focused dataset:

```powershell
.\Get-MailEnabledSecurityGroups.ps1 `
    -MemberScope GroupsOnly `
    -ExportPath C:\Reports
```

Run a small preview using the first 15 groups alphabetically and the first 15 direct
member distinguished names from each group:

```powershell
.\Get-MailEnabledSecurityGroups.ps1 `
    -TestMode `
    -MemberScope GroupsOnly `
    -ExportPath C:\Reports\Test
```

Use `-TestLimit 25` with `-TestMode` to choose a different sample size. Test inventory
rows have `ReportMode=TestSample`; `DirectMemberCount` remains the true directory count,
while `InspectedDirectMemberCount` and `MembershipWasTruncated` describe the sample.

Inventory groups without resolving members:

```powershell
.\Get-MailEnabledSecurityGroups.ps1 `
    -MemberScope None `
    -SearchBase 'OU=Groups,DC=contoso,DC=com' `
    -ExportPath C:\Reports
```

Omit `-ExportPath` to choose a directory after collection, or use `-NoExportPrompt` to
collect and show the terminal summary without exporting.
