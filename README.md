# Set User Group Access

`Set-UserAccessPackage.ps1` is an ad hoc Microsoft Entra group-membership utility for
granting access that is already connected to groups, such as SSO or SCIM provisioning. It
can add one user to explicitly selected groups or compare the user with a reference user and
process only the direct assigned-group membership delta.

> [!IMPORTANT]
> The script changes group membership only. It does not assign groups to enterprise
> applications. The selected groups must already be configured for the relevant SSO or SCIM
> applications.

## Features

- Interactive menu when run without parameters
- Target and reference users accepted by object ID or user principal name
- Explicit groups accepted by object ID or exact display name
- Editable named group sets such as `HR` and `IT`, selectable with `-ManualGroupSet`
- Pattern-filtered, multi-select group browser using `Out-GridView`
- Direct membership comparison, without flattening inherited/nested memberships
- Dynamic groups excluded from the assigned-group delta
- Per-group review with `-Compare`, or automatic delta processing with `-AddAll`
- Automatic safety skips for role-assignable, synchronized, mail-enabled read-only, and
  unsupported group types
- Duplicate and already-member detection
- Retry handling, `-WhatIf`, standard PowerShell `-Confirm`, and CSV/JSON result export

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- `Microsoft.Graph.Authentication` module
- Delegated Microsoft Graph scopes:
  - `User.Read.All`
  - `GroupMember.ReadWrite.All`
- A signed-in account that can update the selected groups, such as a group owner or a user
  with a supported Microsoft Entra role
- A Windows Desktop session for `-BrowseGroups`; `Out-GridView` is unavailable on Server
  Core, Nano Server, and non-Windows platforms

Install the required module if needed:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

## Usage

Run interactively:

```powershell
.\Set-UserAccessPackage.ps1
```

Add a user to explicitly selected groups:

```powershell
.\Set-UserAccessPackage.ps1 `
    -UserId new.user@contoso.com `
    -GroupId 'Sales SSO','CRM SCIM'
```

Configure reusable group sets in the `MANUAL GROUP SETS` section near the top of
`Set-UserAccessPackage.ps1`. Object IDs are recommended because they remain stable when a
group is renamed:

```powershell
$script:ManualGroupSets = [ordered]@{
    HR = @(
        '11111111-1111-1111-1111-111111111111'
        'HR Application Users'
    )
    IT = @(
        '22222222-2222-2222-2222-222222222222'
        'IT Admin Tools SSO'
    )
}
```

Apply one group set:

```powershell
.\Set-UserAccessPackage.ps1 `
    -UserId new.user@contoso.com `
    -ManualGroupSet HR
```

Multiple sets can be combined; duplicate group identities and resolved group IDs are
deduplicated:

```powershell
.\Set-UserAccessPackage.ps1 `
    -UserId new.user@contoso.com `
    -ManualGroupSet HR,IT `
    -WhatIf
```

Configure the default GUI browser patterns in the `GROUP BROWSER PATTERNS` section near the
top of the script. Matching is case-insensitive, and standard PowerShell wildcards are
supported:

```powershell
$script:GroupBrowserPatterns = @(
    'User.app-*'
    'User.sso-*'
)
```

Open the browser using those configured patterns:

```powershell
.\Set-UserAccessPackage.ps1 `
    -UserId new.user@contoso.com `
    -BrowseGroups
```

Use Ctrl or Shift to select multiple rows and click **OK**. The window contains only groups
that match at least one pattern, are eligible for direct Graph membership changes, and do
not already contain the target as a direct member.

Override the configured patterns for one run:

```powershell
.\Set-UserAccessPackage.ps1 `
    -UserId new.user@contoso.com `
    -BrowseGroups `
    -GroupPattern 'User.app-*','User.sso-*' `
    -WhatIf
```

A value without wildcard characters is treated as a prefix, so `-GroupPattern User.app-`
is equivalent to `-GroupPattern User.app-*`.

Compare a user with a reference user and approve each eligible delta group:

```powershell
.\Set-UserAccessPackage.ps1 `
    -UserId new.user@contoso.com `
    -CompareUserId template.user@contoso.com `
    -Compare
```

Preview adding the complete eligible delta without making changes:

```powershell
.\Set-UserAccessPackage.ps1 `
    -UserId new.user@contoso.com `
    -CompareUserId template.user@contoso.com `
    -AddAll `
    -WhatIf `
    -NoExportPrompt
```

Remove `-WhatIf` to perform the additions. `-AddAll` suppresses the script's custom
per-group review, while standard `-Confirm` remains available if PowerShell confirmation is
desired. Comparison uses direct membership for both users: a reference user's inherited
membership is not copied, and a target user's inherited membership does not count as an
existing direct assignment.

Use `-ExportPath C:\Reports` to automatically create a timestamped CSV, or specify a `.json`
filename for JSON. Every excluded, declined, successful, and failed group is represented in
the result data.

---

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
- Optional create-only mode that skips users with an existing email authentication method
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

Populate only missing email authentication methods for all enabled member users. This
checks each user first, leaves every existing method unchanged, and uses only the user's
directory `mail` property for new methods:

```powershell
.\Set-EmailAuthenticationMethod.ps1 `
    -AllUsers `
    -SkipExistingEmailMethod `
    -EmailSource Mail `
    -Force
```

The same workflow is available as option 6 in the interactive menu. `OnlyIfMissing` is
also accepted as a shorter alias for `SkipExistingEmailMethod`.

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
