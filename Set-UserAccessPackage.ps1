<#
.SYNOPSIS
Adds a Microsoft Entra user to selected assigned groups or copies the assigned-group
membership delta from another user.

.DESCRIPTION
Uses Microsoft Graph v1.0 to add one target user to Microsoft Entra groups. Run the script
without parameters for an interactive menu, specify one or more groups directly, select a
named manual group set configured in this file, browse groups by configured or supplied
display-name patterns in a GUI, or compare the target with a reference user.

Comparison mode reads direct group membership only. Groups with DynamicMembership in
groupTypes are not candidates because their membership is rule controlled. The delta is the
reference user's direct, assigned groups that the target user does not already belong to
directly. Use -Compare to approve each delta group or -AddAll to process the whole delta.

For safety, the script also skips role-assignable groups, groups synchronized from
on-premises, mail-enabled security/distribution groups, and group types that Microsoft Graph
does not support for membership writes. Each skipped group is included in the run results.

This script changes group membership only. The groups must already be assigned to the
appropriate enterprise applications for their members to receive SSO or SCIM access.

.PARAMETER UserId
Object ID or user principal name of the target user who will be added to groups.

.PARAMETER GroupId
One or more group object IDs or exact display names. This selects manual group mode.
Comma- and semicolon-separated values are also accepted.

.PARAMETER ManualGroupSet
One or more named group sets from the editable ManualGroupSets section near the top of this
file. For example, -ManualGroupSet HR or -ManualGroupSet HR,IT.

.PARAMETER BrowseGroups
Opens an Out-GridView window containing eligible groups whose display names match the
configured or supplied patterns. Select one or more rows and click OK to process them.

.PARAMETER GroupPattern
One or more case-insensitive wildcard patterns used by -BrowseGroups. Values without a
wildcard are treated as prefixes by appending *. When omitted, the editable
GroupBrowserPatterns section near the top of this file is used.

.PARAMETER CompareUserId
Object ID or user principal name of the reference user whose direct assigned groups are
compared with the target user's direct groups.

.PARAMETER Compare
Reviews the comparison delta one group at a time. Enter y to add, n to skip, or q to skip
that group and every remaining group. -WhatIf previews every candidate without prompting.

.PARAMETER AddAll
Processes every eligible group in the comparison delta without a custom bulk prompt.
Standard -WhatIf and -Confirm behavior is still available.

.PARAMETER ExportPath
Exports detailed results without prompting. Use a .json extension for JSON; all other
extensions produce CSV. Supplying a directory creates a timestamped CSV within it.

.EXAMPLE
.\Set-UserAccessPackage.ps1

Opens the interactive menu.

.EXAMPLE
.\Set-UserAccessPackage.ps1 -UserId new.user@contoso.com -GroupId 'Sales SSO','CRM SCIM'

Adds the target user to two explicitly selected groups.

.EXAMPLE
.\Set-UserAccessPackage.ps1 -UserId new.user@contoso.com -ManualGroupSet HR

Adds the target user to every eligible group configured in the HR manual group set.

.EXAMPLE
.\Set-UserAccessPackage.ps1 -UserId new.user@contoso.com -BrowseGroups

Opens the group-selection GUI using the patterns configured in this file.

.EXAMPLE
.\Set-UserAccessPackage.ps1 -UserId new.user@contoso.com -BrowseGroups -GroupPattern 'User.app-*','User.sso-*'

Opens the group-selection GUI using patterns supplied for this run.

.EXAMPLE
.\Set-UserAccessPackage.ps1 -UserId new.user@contoso.com -CompareUserId template.user@contoso.com -Compare

Shows each eligible assigned-group delta and asks whether it should be added.

.EXAMPLE
.\Set-UserAccessPackage.ps1 -UserId new.user@contoso.com -CompareUserId template.user@contoso.com -AddAll -WhatIf

Previews all changes that an add-all comparison would make.

.NOTES
Requires Windows PowerShell 5.1 or PowerShell 7+ and the
Microsoft.Graph.Authentication module. The delegated Graph scopes are User.Read.All and
GroupMember.ReadWrite.All. The signed-in account must also have permission to update the
selected groups, such as being a group owner or holding a supported Microsoft Entra role.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Interactive', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Manual', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'ManualGroupSet', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'GroupBrowser', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'CompareReview', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'CompareAll', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$UserId,

    [Parameter(Mandatory, ParameterSetName = 'Manual', Position = 1)]
    [Alias('GroupIds')]
    [ValidateNotNullOrEmpty()]
    [string[]]$GroupId,

    [Parameter(Mandatory, ParameterSetName = 'ManualGroupSet', Position = 1)]
    [Alias('GroupSet')]
    [ValidateNotNullOrEmpty()]
    [string[]]$ManualGroupSet,

    [Parameter(Mandatory, ParameterSetName = 'GroupBrowser')]
    [switch]$BrowseGroups,

    [Parameter(ParameterSetName = 'GroupBrowser')]
    [Alias('Pattern')]
    [ValidateNotNullOrEmpty()]
    [string[]]$GroupPattern,

    [Parameter(Mandatory, ParameterSetName = 'CompareReview', Position = 1)]
    [Parameter(Mandatory, ParameterSetName = 'CompareAll', Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$CompareUserId,

    [Parameter(Mandatory, ParameterSetName = 'CompareReview')]
    [switch]$Compare,

    [Parameter(Mandatory, ParameterSetName = 'CompareAll')]
    [switch]$AddAll,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [ValidateNotNullOrEmpty()]
    [string]$ExportPath,

    [switch]$NoExportPrompt,

    [switch]$DisconnectWhenDone
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# MANUAL GROUP SETS - EDIT THIS SECTION
#
# Add group object IDs (recommended because they are unique and rename-safe) or exact group
# display names beneath each label. Add more labels by copying the HR or IT pattern. The
# labels become values accepted by -ManualGroupSet and choices in the interactive menu.
#
# Example only (leave commented until replaced with a real group):
#     '00000000-0000-0000-0000-000000000000'
#     'HR Application Users'
# -----------------------------------------------------------------------------
$script:ManualGroupSets = [ordered]@{
    HR = @(
        # Add HR group object IDs or exact display names here.
    )
    IT = @(
        # Add IT group object IDs or exact display names here.
    )
}
# -----------------------------------------------------------------------------
# END MANUAL GROUP SETS
# -----------------------------------------------------------------------------

# -----------------------------------------------------------------------------
# GROUP BROWSER PATTERNS - EDIT THIS SECTION
#
# These case-insensitive wildcard patterns populate the -BrowseGroups GUI when
# -GroupPattern is not supplied. A value without wildcard characters is treated as a prefix.
# -----------------------------------------------------------------------------
$script:GroupBrowserPatterns = @(
    'User.app-*'
    'User.sso-*'
)
# -----------------------------------------------------------------------------
# END GROUP BROWSER PATTERNS
# -----------------------------------------------------------------------------

$script:GraphBaseUri = 'https://graph.microsoft.com/v1.0'
$script:RunStarted = Get-Date
$script:ConnectedByScript = $false
$script:Results = New-Object 'System.Collections.Generic.List[object]'
$script:SelectedManualGroupSets = @()
$script:SelectedGroupBrowserPatterns = @()

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'DEBUG', 'STEP')]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message
    )

    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'DEBUG'   { 'DarkGray' }
        'STEP'    { 'Cyan' }
        default   { 'Gray' }
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-GraphErrorInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $statusCode = $null
    $message = $ErrorRecord.Exception.Message

    try {
        if ($null -ne $ErrorRecord.Exception.Response) {
            $rawStatus = $ErrorRecord.Exception.Response.StatusCode
            if ($null -ne $rawStatus.value__) {
                $statusCode = [int]$rawStatus.value__
            }
            else {
                $statusCode = [int]$rawStatus
            }
        }
    }
    catch {
        # Preserve the original exception when the response shape differs by PowerShell version.
    }

    if ($null -ne $ErrorRecord.ErrorDetails -and
        -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ErrorDetails.Message)) {
        try {
            $errorBody = $ErrorRecord.ErrorDetails.Message | ConvertFrom-Json
            if ($null -ne $errorBody.error.message) {
                $message = [string]$errorBody.error.message
            }
        }
        catch {
            $message = $ErrorRecord.ErrorDetails.Message
        }
    }

    [pscustomobject]@{
        StatusCode = $statusCode
        Message    = $message
    }
}

function Invoke-GraphRequestWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [hashtable]$Headers,
        [ValidateRange(0, 10)][int]$MaxRetryCount = 4
    )

    $attempt = 0
    while ($true) {
        try {
            $arguments = @{
                Method      = $Method
                Uri         = $Uri
                OutputType  = 'PSObject'
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) {
                $arguments.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
                $arguments.ContentType = 'application/json'
            }
            if ($null -ne $Headers -and $Headers.Count -gt 0) {
                $arguments.Headers = $Headers
            }

            Write-Log -Level DEBUG -Message "$Method $Uri"
            return Invoke-MgGraphRequest @arguments
        }
        catch {
            $details = Get-GraphErrorInfo -ErrorRecord $_
            $retryable = $details.StatusCode -in @(429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -ge $MaxRetryCount) {
                throw
            }

            $attempt++
            $delay = [math]::Min([math]::Pow(2, $attempt), 30)
            Write-Log -Level WARN -Message "Graph request failed with HTTP $($details.StatusCode). Retrying in $delay second(s), attempt $attempt of $MaxRetryCount."
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-GraphCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Headers
    )

    $items = New-Object 'System.Collections.Generic.List[object]'
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-GraphRequestWithRetry -Method GET -Uri $nextUri -Headers $Headers
        if ($response.PSObject.Properties.Name -contains 'value' -and $null -ne $response.value) {
            foreach ($item in @($response.value)) {
                $items.Add($item)
            }
        }

        if ($response.PSObject.Properties.Name -contains '@odata.nextLink') {
            $nextUri = [string]$response.'@odata.nextLink'
        }
        else {
            $nextUri = $null
        }
    }

    return $items.ToArray()
}

function Connect-RequiredGraphSession {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Scopes)

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue) -or
        -not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        throw 'The Microsoft.Graph.Authentication module is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    $tenantMismatch = $false
    if ($null -ne $context -and -not [string]::IsNullOrWhiteSpace($TenantId) -and
        [string]$context.TenantId -ine $TenantId) {
        $tenantMismatch = $true
        Write-Log -Level INFO -Message "The active Graph context is for tenant $($context.TenantId); reconnecting to requested tenant $TenantId."
    }

    $missingScopes = @()
    if ($null -eq $context -or $tenantMismatch) {
        $missingScopes = $Scopes
    }
    else {
        foreach ($scope in $Scopes) {
            if ($context.Scopes -notcontains $scope) {
                $missingScopes += $scope
            }
        }
    }

    if ($missingScopes.Count -gt 0) {
        Write-Log -Level STEP -Message "Connecting to Microsoft Graph with delegated scopes: $($Scopes -join ', ')"
        $connectArguments = @{
            Scopes    = $Scopes
            NoWelcome = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            $connectArguments.TenantId = $TenantId
        }
        Connect-MgGraph @connectArguments | Out-Null
        $script:ConnectedByScript = $true
        $context = Get-MgContext
    }
    else {
        Write-Log -Level INFO -Message "Using the existing Microsoft Graph connection for $($context.Account)."
    }

    if ($null -eq $context) {
        throw 'Microsoft Graph authentication did not produce an active context.'
    }

    $stillMissing = @($Scopes | Where-Object { $context.Scopes -notcontains $_ })
    if ($stillMissing.Count -gt 0) {
        throw "The Microsoft Graph session is missing required delegated scope(s): $($stillMissing -join ', '). An administrator might need to grant consent."
    }

    Write-Log -Level SUCCESS -Message "Connected to tenant $($context.TenantId) as $($context.Account)."
    Write-Log -Level WARN -Message 'The signed-in account must be allowed to update membership for each selected group.'
}

function Get-DirectoryUser {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Identity)

    $encodedIdentity = [uri]::EscapeDataString($Identity.Trim())
    $uri = "$script:GraphBaseUri/users/$encodedIdentity`?%24select=id,displayName,userPrincipalName,accountEnabled,userType"
    return Invoke-GraphRequestWithRetry -Method GET -Uri $uri
}

function Get-DirectUserGroups {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$User)

    $select = 'id,displayName,description,groupTypes,mail,mailEnabled,securityEnabled,isAssignableToRole,onPremisesSyncEnabled'
    $uri = "$script:GraphBaseUri/users/$($User.id)/memberOf/microsoft.graph.group?%24select=$select&%24top=999"
    return @(Get-GraphCollection -Uri $uri)
}

function Resolve-GroupObject {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Identity)

    $trimmedIdentity = $Identity.Trim()
    $select = 'id,displayName,description,groupTypes,mail,mailEnabled,securityEnabled,isAssignableToRole,onPremisesSyncEnabled'
    if ($trimmedIdentity -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        $encodedIdentity = [uri]::EscapeDataString($trimmedIdentity)
        return Invoke-GraphRequestWithRetry -Method GET -Uri "$script:GraphBaseUri/groups/$encodedIdentity`?%24select=$select"
    }

    $escapedName = $trimmedIdentity.Replace("'", "''")
    $encodedFilter = [uri]::EscapeDataString("displayName eq '$escapedName'")
    $groupSearchResults = @(Get-GraphCollection -Uri "$script:GraphBaseUri/groups?%24filter=$encodedFilter&%24select=$select")
    if ($groupSearchResults.Count -eq 0) {
        throw "No group with the exact display name '$trimmedIdentity' was found. Try its object ID instead."
    }
    if ($groupSearchResults.Count -gt 1) {
        $ids = ($groupSearchResults | ForEach-Object { $_.id }) -join ', '
        throw "More than one group is named '$trimmedIdentity'. Use one of these object IDs: $ids"
    }

    return $groupSearchResults[0]
}

function Test-GroupHasType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Group,
        [Parameter(Mandatory)][string]$Type
    )

    return @($Group.groupTypes) -contains $Type
}

function Get-GroupKind {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Group)

    if (Test-GroupHasType -Group $Group -Type 'Unified') {
        return 'Microsoft 365'
    }
    if ([bool]$Group.mailEnabled -and [bool]$Group.securityEnabled) {
        return 'Mail-enabled security'
    }
    if ([bool]$Group.mailEnabled -and -not [bool]$Group.securityEnabled) {
        return 'Distribution'
    }
    if ([bool]$Group.securityEnabled) {
        return 'Security'
    }
    return 'Unsupported'
}

function Get-GroupEligibility {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Group)

    $kind = Get-GroupKind -Group $Group
    if (Test-GroupHasType -Group $Group -Type 'DynamicMembership') {
        return [pscustomobject]@{ Eligible = $false; Kind = $kind; Reason = 'Dynamic membership is controlled by a rule.' }
    }
    if ($null -ne $Group.isAssignableToRole -and [bool]$Group.isAssignableToRole) {
        return [pscustomobject]@{ Eligible = $false; Kind = $kind; Reason = 'Role-assignable groups are excluded because they require privileged role-management access.' }
    }
    if ($null -ne $Group.onPremisesSyncEnabled -and [bool]$Group.onPremisesSyncEnabled) {
        return [pscustomobject]@{ Eligible = $false; Kind = $kind; Reason = 'Membership is mastered by the synchronized on-premises directory.' }
    }
    if ($kind -eq 'Mail-enabled security' -or $kind -eq 'Distribution') {
        return [pscustomobject]@{ Eligible = $false; Kind = $kind; Reason = 'This mail-enabled group type is read-only in Microsoft Graph.' }
    }
    if ($kind -notin @('Security', 'Microsoft 365')) {
        return [pscustomobject]@{ Eligible = $false; Kind = $kind; Reason = 'Microsoft Graph does not support membership writes for this group type.' }
    }

    return [pscustomobject]@{ Eligible = $true; Kind = $kind; Reason = $null }
}

function Add-RunResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mode,
        [object]$TargetUser,
        [object]$ReferenceUser,
        [object]$Group,
        [string]$GroupKind,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message
    )

    $script:Results.Add([pscustomobject][ordered]@{
        Timestamp                 = (Get-Date).ToString('o')
        Mode                      = $Mode
        ManualGroupSets           = if ($script:SelectedManualGroupSets.Count -gt 0) { $script:SelectedManualGroupSets -join '; ' } else { $null }
        GroupBrowserPatterns      = if ($script:SelectedGroupBrowserPatterns.Count -gt 0) { $script:SelectedGroupBrowserPatterns -join '; ' } else { $null }
        TargetUserId              = if ($null -ne $TargetUser) { [string]$TargetUser.id } else { $null }
        TargetUserPrincipalName   = if ($null -ne $TargetUser) { [string]$TargetUser.userPrincipalName } else { $null }
        ReferenceUserId           = if ($null -ne $ReferenceUser) { [string]$ReferenceUser.id } else { $null }
        ReferenceUserPrincipalName = if ($null -ne $ReferenceUser) { [string]$ReferenceUser.userPrincipalName } else { $null }
        GroupId                   = if ($null -ne $Group) { [string]$Group.id } else { $null }
        GroupDisplayName          = if ($null -ne $Group) { [string]$Group.displayName } else { $null }
        GroupKind                 = $GroupKind
        MembershipType           = if ($null -eq $Group) { $null } elseif (Test-GroupHasType -Group $Group -Type 'DynamicMembership') { 'Dynamic' } else { 'Assigned' }
        Action                    = $Action
        Status                    = $Status
        Message                   = $Message
    })
}

function Show-InteractiveMenu {
    [CmdletBinding()]
    param()

    Write-Host ''
    Write-Host 'Set User Group Access' -ForegroundColor Cyan
    Write-Host 'Adds direct assigned-group memberships used for access provisioning.' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  1. Add a user to one or more specified groups'
    Write-Host '  2. Add a user using a configured manual group set'
    Write-Host '  3. Browse matching groups in a GUI and select one or more'
    Write-Host '  4. Compare with another user and review each group'
    Write-Host '  5. Compare with another user and add the entire eligible delta'
    Write-Host ''

    $selection = Read-Host 'Choose 1-5'
    $script:UserId = Read-Host 'Target user object ID or user principal name'
    if ([string]::IsNullOrWhiteSpace($script:UserId)) {
        throw 'A target user is required.'
    }

    switch ($selection) {
        '1' {
            $enteredGroups = Read-Host 'Group object IDs or exact display names, separated by commas or semicolons'
            $script:GroupId = @($enteredGroups -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($script:GroupId.Count -eq 0) {
                throw 'At least one group is required.'
            }
            return 'Manual'
        }
        '2' {
            Write-Host ''
            Write-Host 'Configured manual group sets' -ForegroundColor Cyan
            foreach ($configuredGroupSetName in @($script:ManualGroupSets.Keys)) {
                $configuredGroupCount = @($script:ManualGroupSets[$configuredGroupSetName] |
                    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
                Write-Host ("  {0} ({1} configured group(s))" -f $configuredGroupSetName, $configuredGroupCount)
            }
            $enteredGroupSets = Read-Host 'Manual group set names, separated by commas or semicolons'
            $script:ManualGroupSet = @($enteredGroupSets -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($script:ManualGroupSet.Count -eq 0) {
                throw 'At least one manual group set is required.'
            }
            return 'ManualGroupSet'
        }
        '3' {
            Write-Host ''
            Write-Host 'Configured group browser patterns' -ForegroundColor Cyan
            foreach ($configuredBrowserPattern in $script:GroupBrowserPatterns) {
                Write-Host "  $configuredBrowserPattern"
            }
            $patternChoice = Read-Host 'Press Enter to use these patterns, or enter replacement patterns separated by commas or semicolons'
            if (-not [string]::IsNullOrWhiteSpace($patternChoice)) {
                $script:GroupPattern = @($patternChoice -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
            $script:BrowseGroups = $true
            return 'GroupBrowser'
        }
        '4' {
            $script:CompareUserId = Read-Host 'Reference user object ID or user principal name'
            if ([string]::IsNullOrWhiteSpace($script:CompareUserId)) {
                throw 'A reference user is required.'
            }
            return 'CompareReview'
        }
        '5' {
            $script:CompareUserId = Read-Host 'Reference user object ID or user principal name'
            if ([string]::IsNullOrWhiteSpace($script:CompareUserId)) {
                throw 'A reference user is required.'
            }
            return 'CompareAll'
        }
        default {
            throw "'$selection' is not valid. Run the script again and choose 1, 2, 3, 4, or 5."
        }
    }
}

function ConvertTo-IdentityList {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Identity)

    return @($Identity |
        ForEach-Object { $_ -split '[,;]' } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ } |
        Select-Object -Unique)
}

function Resolve-ManualGroupSetSelection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Name)

    $requestedGroupSetNames = @(ConvertTo-IdentityList -Identity $Name)
    $configuredGroupSetNames = @($script:ManualGroupSets.Keys)
    $selectedGroupSetNames = New-Object 'System.Collections.Generic.List[string]'
    $selectedGroupIdentities = New-Object 'System.Collections.Generic.List[string]'

    foreach ($requestedGroupSetName in $requestedGroupSetNames) {
        $configuredGroupSetName = @($configuredGroupSetNames |
            Where-Object { [string]$_ -ieq $requestedGroupSetName } |
            Select-Object -First 1)
        if ($configuredGroupSetName.Count -eq 0) {
            $availableGroupSets = $configuredGroupSetNames -join ', '
            throw "Manual group set '$requestedGroupSetName' is not configured. Available group sets: $availableGroupSets"
        }

        $canonicalGroupSetName = [string]$configuredGroupSetName[0]
        $configuredGroupIdentities = @($script:ManualGroupSets[$canonicalGroupSetName] |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ })
        if ($configuredGroupIdentities.Count -eq 0) {
            Write-Log -Level WARN -Message "Manual group set '$canonicalGroupSetName' is empty; it will not add any candidates."
        }

        $selectedGroupSetNames.Add($canonicalGroupSetName)
        foreach ($configuredGroupIdentity in $configuredGroupIdentities) {
            $selectedGroupIdentities.Add($configuredGroupIdentity)
        }
    }

    $script:SelectedManualGroupSets = @($selectedGroupSetNames.ToArray() | Select-Object -Unique)
    $uniqueGroupIdentities = @($selectedGroupIdentities.ToArray() | Select-Object -Unique)
    if ($uniqueGroupIdentities.Count -eq 0) {
        throw "The selected manual group set(s) contain no group object IDs or display names: $($script:SelectedManualGroupSets -join ', ')"
    }

    return $uniqueGroupIdentities
}

function ConvertTo-GroupBrowserPatterns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Value)

    $normalizedPatterns = New-Object 'System.Collections.Generic.List[string]'
    foreach ($rawPattern in @(ConvertTo-IdentityList -Identity $Value)) {
        $normalizedPattern = $rawPattern
        if ($normalizedPattern -notmatch '[*?\[]') {
            $normalizedPattern = "$normalizedPattern*"
        }
        $normalizedPatterns.Add($normalizedPattern)
    }

    return @($normalizedPatterns.ToArray() | Select-Object -Unique)
}

function Get-AllDirectoryGroups {
    [CmdletBinding()]
    param()

    $select = 'id,displayName,description,groupTypes,mail,mailEnabled,securityEnabled,isAssignableToRole,onPremisesSyncEnabled'
    $uri = "$script:GraphBaseUri/groups?%24select=$select&%24top=999"
    return @(Get-GraphCollection -Uri $uri)
}

function Select-GroupsFromBrowser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$TargetUser,
        [Parameter(Mandatory)][hashtable]$TargetGroupIds,
        [Parameter(Mandatory)][string[]]$Pattern
    )

    if (-not (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
        throw 'Group browser mode requires Out-GridView and a Windows Desktop session. Run the script from Windows PowerShell or PowerShell 7 on Windows with Desktop Experience.'
    }

    $normalizedPatterns = @(ConvertTo-GroupBrowserPatterns -Value $Pattern)
    if ($normalizedPatterns.Count -eq 0) {
        throw 'At least one non-empty group browser pattern is required.'
    }
    $script:SelectedGroupBrowserPatterns = $normalizedPatterns
    Write-Log -Level STEP -Message "Retrieving directory groups for browser pattern(s): $($normalizedPatterns -join ', ')."

    $directoryGroups = @(Get-AllDirectoryGroups)
    $matchingGroups = New-Object 'System.Collections.Generic.List[object]'
    foreach ($directoryGroup in $directoryGroups) {
        foreach ($normalizedPattern in $normalizedPatterns) {
            if ([string]$directoryGroup.displayName -like $normalizedPattern) {
                $matchingGroups.Add($directoryGroup)
                break
            }
        }
    }

    $browserRows = New-Object 'System.Collections.Generic.List[object]'
    $eligibleGroupLookup = @{}
    $alreadyMemberCount = 0
    $ineligibleCount = 0
    foreach ($matchingGroup in $matchingGroups.ToArray()) {
        if ($TargetGroupIds.ContainsKey([string]$matchingGroup.id)) {
            $alreadyMemberCount++
            continue
        }

        $eligibility = Get-GroupEligibility -Group $matchingGroup
        if (-not $eligibility.Eligible) {
            $ineligibleCount++
            continue
        }

        $eligibleGroupLookup[[string]$matchingGroup.id] = $matchingGroup
        $browserRows.Add([pscustomobject][ordered]@{
            DisplayName = [string]$matchingGroup.displayName
            Type        = [string]$eligibility.Kind
            Description = [string]$matchingGroup.description
            Mail        = [string]$matchingGroup.mail
            ObjectId    = [string]$matchingGroup.id
        })
    }

    Write-Host ''
    Write-Host 'Group browser summary' -ForegroundColor Cyan
    Write-Host "  Directory groups retrieved: $($directoryGroups.Count)"
    Write-Host "  Pattern matches:             $($matchingGroups.Count)"
    Write-Host "  Already a direct member:     $alreadyMemberCount"
    Write-Host "  Ineligible groups hidden:    $ineligibleCount"
    Write-Host "  Available for selection:     $($browserRows.Count)"

    if ($browserRows.Count -eq 0) {
        Write-Log -Level WARN -Message 'No eligible, unassigned groups match the browser patterns.'
        return @()
    }

    $windowTitle = "Select groups for $($TargetUser.displayName) <$($TargetUser.userPrincipalName)> - use Ctrl/Shift for multiple rows"
    try {
        $selectedRows = @($browserRows.ToArray() |
            Sort-Object DisplayName, ObjectId |
            Out-GridView -Title $windowTitle -OutputMode Multiple)
    }
    catch {
        throw "The group browser window could not be opened: $($_.Exception.Message)"
    }

    if ($selectedRows.Count -eq 0) {
        Write-Log -Level INFO -Message 'The group browser was closed without selecting any groups.'
        return @()
    }

    $selectedGroups = New-Object 'System.Collections.Generic.List[object]'
    foreach ($selectedRow in $selectedRows) {
        $selectedObjectId = [string]$selectedRow.ObjectId
        if ($eligibleGroupLookup.ContainsKey($selectedObjectId)) {
            $selectedGroups.Add($eligibleGroupLookup[$selectedObjectId])
        }
    }

    Write-Log -Level SUCCESS -Message "Selected $($selectedGroups.Count) group(s) in the browser."
    return $selectedGroups.ToArray()
}

function Show-UserSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][object]$User
    )

    Write-Host ('  {0,-10} {1} <{2}> [{3}]' -f $Label, $User.displayName, $User.userPrincipalName, $User.id)
}

function Show-GroupCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Group,
        [Parameter(Mandatory)][string]$Kind,
        [int]$Index,
        [int]$Total
    )

    Write-Host ''
    if ($Total -gt 0) {
        Write-Host "[$Index/$Total] $($Group.displayName)" -ForegroundColor Cyan
    }
    else {
        Write-Host $Group.displayName -ForegroundColor Cyan
    }
    Write-Host "  Type:        $Kind"
    Write-Host "  Object ID:   $($Group.id)"
    if (-not [string]::IsNullOrWhiteSpace([string]$Group.description)) {
        Write-Host "  Description: $($Group.description)"
    }
}

function Add-UserToGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$TargetUser,
        [Parameter(Mandatory)][object]$Group
    )

    $uri = "$script:GraphBaseUri/groups/$($Group.id)/members/`$ref"
    $body = @{ '@odata.id' = "$script:GraphBaseUri/directoryObjects/$($TargetUser.id)" }
    Invoke-GraphRequestWithRetry -Method POST -Uri $uri -Body $body | Out-Null
}

function Export-RunResults {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $resolvedPath = $Path
    $extension = [System.IO.Path]::GetExtension($Path)
    $looksLikeDirectory = [string]::IsNullOrWhiteSpace($extension) -or (Test-Path -LiteralPath $Path -PathType Container)
    if ($looksLikeDirectory) {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
        $resolvedPath = Join-Path $Path ("UserGroupAccess-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    else {
        $parent = Split-Path -Parent $Path
        if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
    }

    if ([System.IO.Path]::GetExtension($resolvedPath) -ieq '.json') {
        $script:Results.ToArray() | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedPath -Encoding UTF8
    }
    else {
        $script:Results.ToArray() | Export-Csv -LiteralPath $resolvedPath -NoTypeInformation -Encoding UTF8
    }
    Write-Log -Level SUCCESS -Message "Exported $($script:Results.Count) result(s) to '$resolvedPath'."
}

function Show-RunSummary {
    [CmdletBinding()]
    param()

    $duration = (Get-Date) - $script:RunStarted
    $counts = @{}
    foreach ($status in @('Succeeded', 'WouldAdd', 'AlreadyMember', 'Declined', 'Skipped', 'Failed')) {
        $counts[$status] = @($script:Results.ToArray() | Where-Object { $_.Status -eq $status }).Count
    }

    Write-Host ''
    Write-Host 'Run summary' -ForegroundColor Cyan
    Write-Host ('  Added:          {0}' -f $counts.Succeeded) -ForegroundColor Green
    Write-Host ('  Would add:      {0}' -f $counts.WouldAdd) -ForegroundColor Cyan
    Write-Host ('  Already member: {0}' -f $counts.AlreadyMember) -ForegroundColor Gray
    Write-Host ('  Declined:       {0}' -f $counts.Declined) -ForegroundColor Yellow
    Write-Host ('  Skipped:        {0}' -f $counts.Skipped) -ForegroundColor Yellow
    Write-Host ('  Failed:         {0}' -f $counts.Failed) -ForegroundColor Red
    Write-Host ('  Duration:        {0:hh\:mm\:ss}' -f $duration)
}

$fatalError = $null
try {
    Write-Host ''
    Write-Host 'Microsoft Graph - Set User Group Access' -ForegroundColor Cyan
    Write-Host 'Direct group membership only; enterprise-app assignments are not changed.' -ForegroundColor Yellow
    Write-Host ''

    $mode = $PSCmdlet.ParameterSetName
    if ($mode -eq 'Interactive') {
        $mode = Show-InteractiveMenu
    }

    Connect-RequiredGraphSession -Scopes @('User.Read.All', 'GroupMember.ReadWrite.All')

    Write-Log -Level STEP -Message "Resolving target user '$UserId'."
    $targetUser = Get-DirectoryUser -Identity $UserId
    $referenceUser = $null

    if ($null -ne $targetUser.accountEnabled -and -not [bool]$targetUser.accountEnabled) {
        Write-Log -Level WARN -Message 'The target user account is disabled. Group membership may not produce application access until the account is enabled.'
    }

    if ($mode -in @('CompareReview', 'CompareAll')) {
        Write-Log -Level STEP -Message "Resolving reference user '$CompareUserId'."
        $referenceUser = Get-DirectoryUser -Identity $CompareUserId
        if ([string]$referenceUser.id -eq [string]$targetUser.id) {
            throw 'The target user and reference user resolve to the same directory object.'
        }
    }

    Write-Host ''
    Write-Host 'Resolved users' -ForegroundColor Cyan
    Show-UserSummary -Label 'Target' -User $targetUser
    if ($null -ne $referenceUser) {
        Show-UserSummary -Label 'Reference' -User $referenceUser
    }

    Write-Log -Level STEP -Message 'Retrieving the target user direct group memberships.'
    $targetGroups = @(Get-DirectUserGroups -User $targetUser)
    $targetGroupIds = @{}
    foreach ($targetGroup in $targetGroups) {
        $targetGroupIds[[string]$targetGroup.id] = $true
    }

    $candidates = New-Object 'System.Collections.Generic.List[object]'
    if ($mode -in @('Manual', 'ManualGroupSet')) {
        $manualGroupIdentities = if ($mode -eq 'ManualGroupSet') {
            @(Resolve-ManualGroupSetSelection -Name $ManualGroupSet)
        }
        else {
            @(ConvertTo-IdentityList -Identity $GroupId)
        }

        if ($mode -eq 'ManualGroupSet') {
            Write-Log -Level STEP -Message "Using manual group set(s): $($script:SelectedManualGroupSets -join ', ')."
        }

        foreach ($identity in $manualGroupIdentities) {
            try {
                Write-Log -Level INFO -Message "Resolving group '$identity'."
                $group = Resolve-GroupObject -Identity $identity
                if (-not @($candidates.ToArray() | Where-Object { [string]$_.id -eq [string]$group.id }).Count) {
                    $candidates.Add($group)
                }
            }
            catch {
                $details = Get-GraphErrorInfo -ErrorRecord $_
                Write-Log -Level ERROR -Message "Could not resolve group '$identity': $($details.Message)"
                Add-RunResult -Mode $mode -TargetUser $targetUser -Action 'Resolve' -Status 'Failed' -Message "Could not resolve group '$identity': $($details.Message)"
            }
        }
    }
    elseif ($mode -eq 'GroupBrowser') {
        $browserPatternSelection = if ($null -ne $GroupPattern -and @($GroupPattern).Count -gt 0) {
            @($GroupPattern)
        }
        else {
            @($script:GroupBrowserPatterns)
        }
        if (@($browserPatternSelection).Count -eq 0) {
            throw 'No group browser patterns are configured. Add patterns to GroupBrowserPatterns or supply -GroupPattern.'
        }

        foreach ($selectedBrowserGroup in @(Select-GroupsFromBrowser -TargetUser $targetUser -TargetGroupIds $targetGroupIds -Pattern $browserPatternSelection)) {
            $candidates.Add($selectedBrowserGroup)
        }
    }
    else {
        Write-Log -Level STEP -Message 'Retrieving the reference user direct group memberships.'
        $referenceGroups = @(Get-DirectUserGroups -User $referenceUser)
        $assignedReferenceGroups = @($referenceGroups | Where-Object { -not (Test-GroupHasType -Group $_ -Type 'DynamicMembership') })
        $dynamicReferenceGroups = @($referenceGroups | Where-Object { Test-GroupHasType -Group $_ -Type 'DynamicMembership' })

        foreach ($dynamicGroup in $dynamicReferenceGroups) {
            $eligibility = Get-GroupEligibility -Group $dynamicGroup
            Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $dynamicGroup -GroupKind $eligibility.Kind -Action 'None' -Status 'Skipped' -Message $eligibility.Reason
        }

        $deltaGroups = @($assignedReferenceGroups |
            Where-Object { -not $targetGroupIds.ContainsKey([string]$_.id) } |
            Sort-Object displayName, id)
        foreach ($deltaGroup in $deltaGroups) {
            $candidates.Add($deltaGroup)
        }

        Write-Host ''
        Write-Host 'Comparison summary' -ForegroundColor Cyan
        Write-Host "  Reference direct groups:          $($referenceGroups.Count)"
        Write-Host "  Reference assigned groups:        $($assignedReferenceGroups.Count)"
        Write-Host "  Reference dynamic groups skipped: $($dynamicReferenceGroups.Count)"
        Write-Host "  Target direct groups:             $($targetGroups.Count)"
        Write-Host "  Assigned-group delta:             $($deltaGroups.Count)"
    }

    if ($candidates.Count -eq 0) {
        Write-Log -Level SUCCESS -Message 'There are no candidate group memberships to add.'
    }

    $quitReview = $false
    $candidateIndex = 0
    foreach ($group in @($candidates.ToArray())) {
        $candidateIndex++
        $eligibility = Get-GroupEligibility -Group $group

        if ($targetGroupIds.ContainsKey([string]$group.id)) {
            $message = 'The target user is already a direct member.'
            Write-Log -Level INFO -Message "$($group.displayName) - $message"
            Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $group -GroupKind $eligibility.Kind -Action 'None' -Status 'AlreadyMember' -Message $message
            continue
        }

        if (-not $eligibility.Eligible) {
            Write-Log -Level WARN -Message "$($group.displayName) - $($eligibility.Reason)"
            Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $group -GroupKind $eligibility.Kind -Action 'None' -Status 'Skipped' -Message $eligibility.Reason
            continue
        }

        if ($quitReview) {
            $message = 'Skipped because review was stopped by the operator.'
            Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $group -GroupKind $eligibility.Kind -Action 'None' -Status 'Declined' -Message $message
            continue
        }

        Show-GroupCandidate -Group $group -Kind $eligibility.Kind -Index $candidateIndex -Total $candidates.Count

        $approved = $true
        if ($mode -eq 'CompareReview' -and -not $WhatIfPreference) {
            $choice = Read-Host 'Add the target user to this group? [y/N/q]'
            if ($choice -match '^(?i)q(?:uit)?$') {
                $approved = $false
                $quitReview = $true
                $message = 'Declined; review of remaining groups was stopped by the operator.'
            }
            elseif ($choice -notmatch '^(?i)y(?:es)?$') {
                $approved = $false
                $message = 'Declined by the operator.'
            }

            if (-not $approved) {
                Write-Log -Level INFO -Message "$($group.displayName) - $message"
                Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $group -GroupKind $eligibility.Kind -Action 'None' -Status 'Declined' -Message $message
                continue
            }
        }

        $targetLabel = "$($targetUser.displayName) <$($targetUser.userPrincipalName)>"
        $operation = "add as a direct member of '$($group.displayName)' [$($group.id)]"
        try {
            if ($PSCmdlet.ShouldProcess($targetLabel, $operation)) {
                Add-UserToGroup -TargetUser $targetUser -Group $group
                $targetGroupIds[[string]$group.id] = $true
                $message = 'Direct group membership added successfully.'
                Write-Log -Level SUCCESS -Message "$($group.displayName) - $message"
                Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $group -GroupKind $eligibility.Kind -Action 'Add' -Status 'Succeeded' -Message $message
            }
            else {
                $status = if ($WhatIfPreference) { 'WouldAdd' } else { 'Declined' }
                $message = if ($WhatIfPreference) { 'Would add the target user as a direct member.' } else { 'Declined by ShouldProcess confirmation.' }
                Write-Log -Level INFO -Message "$($group.displayName) - $message"
                Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $group -GroupKind $eligibility.Kind -Action 'Add' -Status $status -Message $message
            }
        }
        catch {
            $details = Get-GraphErrorInfo -ErrorRecord $_
            $message = "Microsoft Graph could not add the membership: $($details.Message)"
            Write-Log -Level ERROR -Message "$($group.displayName) - $message"
            Add-RunResult -Mode $mode -TargetUser $targetUser -ReferenceUser $referenceUser -Group $group -GroupKind $eligibility.Kind -Action 'Add' -Status 'Failed' -Message $message
        }
    }
}
catch {
    $fatalError = $_
    $details = Get-GraphErrorInfo -ErrorRecord $_
    Write-Log -Level ERROR -Message "The run could not continue: $($details.Message)"
    Add-RunResult -Mode 'Run' -Action 'Fatal' -Status 'Failed' -Message $details.Message
}
finally {
    Show-RunSummary

    try {
        if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            Export-RunResults -Path $ExportPath
        }
        elseif (-not $NoExportPrompt) {
            Write-Host ''
            $exportChoice = Read-Host 'Export these results? [y/N]'
            if ($exportChoice -match '^(?i)y(?:es)?$') {
                $chosenPath = Read-Host 'Export file or directory path (.json for JSON; otherwise CSV)'
                if ([string]::IsNullOrWhiteSpace($chosenPath)) {
                    Write-Log -Level WARN -Message 'No export path was entered; results were not exported.'
                }
                else {
                    Export-RunResults -Path $chosenPath
                }
            }
            else {
                Write-Log -Level INFO -Message 'Results were not exported.'
            }
        }
    }
    catch {
        Write-Log -Level ERROR -Message "Could not export results: $($_.Exception.Message)"
        if ($null -eq $fatalError) {
            $fatalError = $_
        }
    }

    if ($DisconnectWhenDone -and $script:ConnectedByScript) {
        try {
            Disconnect-MgGraph | Out-Null
            Write-Log -Level INFO -Message 'Disconnected the Microsoft Graph session created by this script.'
        }
        catch {
            Write-Log -Level WARN -Message "Could not disconnect Microsoft Graph: $($_.Exception.Message)"
        }
    }
}

if ($null -ne $fatalError) {
    throw $fatalError
}
