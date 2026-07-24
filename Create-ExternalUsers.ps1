<#
.SYNOPSIS
Invites one or more external users to Microsoft Entra ID and optionally adds them to groups.

.DESCRIPTION
Uses Microsoft Graph v1.0 to create Microsoft Entra B2B guest invitations. Run the script
without targeting parameters for an interactive menu, invite one user with -EmailAddress,
invite several users with -EmailAddresses, or import users from a CSV file with -CsvPath.

One or more groups can be supplied by object ID or exact display name. Eligible security
and Microsoft 365 groups are supported. Dynamic, role-assignable, synchronized,
mail-enabled security, distribution, and unsupported groups are skipped safely.

Before creating an invitation, the script looks for an existing directory user with the
same mail or otherMails value. An existing Guest is reused for group assignment without
being reinvited. An existing Member is skipped so the script cannot accidentally treat an
internal account as an external user.

.PARAMETER EmailAddress
Email address of one external user to invite.

.PARAMETER DisplayName
Optional display name for the user supplied with -EmailAddress.

.PARAMETER EmailAddresses
Email addresses of multiple external users. Comma- and semicolon-separated values are
also accepted. Use -CsvPath when each user needs a display name.

.PARAMETER CsvPath
Path to a CSV containing EmailAddress and optional DisplayName columns. The aliases
InvitedUserEmailAddress and Mail are also accepted for the email column.

.PARAMETER GroupId
One or more group object IDs or exact display names. Every successfully invited or reused
guest is added as a direct member of every eligible group. Comma- and semicolon-separated
values are also accepted.

.PARAMETER InviteRedirectUrl
URL to which the guest is redirected after redeeming the invitation. The default is
https://myapps.microsoft.com.

.PARAMETER DoNotSendInvitationMessage
Creates the invitation without asking Microsoft Graph to send an invitation email. The
redemption URL is written to the terminal and included in exported results.

.PARAMETER CustomizedMessageBody
Optional plain-text message included in the Microsoft invitation email.

.PARAMETER MessageLanguage
Optional language tag for the Microsoft invitation email, such as en-US.

.PARAMETER Force
Suppresses the script's custom confirmation before a multi-user run. Standard -WhatIf and
-Confirm behavior remains available.

.PARAMETER ExportPath
Exports detailed results without prompting. Use a .json extension for JSON; all other
extensions produce CSV. Supplying a directory creates a timestamped CSV within it.

.EXAMPLE
.\Create-ExternalUsers.ps1

Opens the interactive menu.

.EXAMPLE
.\Create-ExternalUsers.ps1 -EmailAddress guest@fabrikam.com -DisplayName 'Fabrikam Guest'

Invites one external user and sends the standard Microsoft invitation email.

.EXAMPLE
.\Create-ExternalUsers.ps1 -EmailAddresses guest1@fabrikam.com,guest2@adatum.com -GroupId 'External App Users','Project Alpha'

Invites multiple users and adds each guest to two groups.

.EXAMPLE
.\Create-ExternalUsers.ps1 -CsvPath .\ExternalUsers.csv -GroupId 11111111-1111-1111-1111-111111111111 -WhatIf -NoExportPrompt

Previews invitations and group assignments from a CSV without changing Microsoft Entra.

.NOTES
Requires Windows PowerShell 5.1 or PowerShell 7+ and the
Microsoft.Graph.Authentication module. The delegated Graph scopes are User.Invite.All and
User.Read.All, plus GroupMember.ReadWrite.All when groups are supplied. The signed-in
account must also be allowed by the tenant's external collaboration settings to invite
guests and must be permitted to update membership for the selected groups.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Interactive', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Single', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$EmailAddress,

    [Parameter(ParameterSetName = 'Single', Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName,

    [Parameter(Mandatory, ParameterSetName = 'Multiple', Position = 0)]
    [Alias('Emails')]
    [ValidateNotNullOrEmpty()]
    [string[]]$EmailAddresses,

    [Parameter(Mandatory, ParameterSetName = 'Csv', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [Alias('GroupIds', 'Groups')]
    [ValidateNotNullOrEmpty()]
    [string[]]$GroupId,

    [ValidateNotNullOrEmpty()]
    [ValidateScript({ $_.IsAbsoluteUri -and $_.Scheme -in @('http', 'https') })]
    [uri]$InviteRedirectUrl = 'https://myapps.microsoft.com',

    [switch]$DoNotSendInvitationMessage,

    [ValidateNotNullOrEmpty()]
    [string]$CustomizedMessageBody,

    [ValidateNotNullOrEmpty()]
    [string]$MessageLanguage,

    [switch]$Force,

    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [ValidateNotNullOrEmpty()]
    [string]$ExportPath,

    [switch]$NoExportPrompt,

    [switch]$DisconnectWhenDone
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:GraphBaseUri = 'https://graph.microsoft.com/v1.0'
$script:RunStarted = Get-Date
$script:ConnectedByScript = $false
$script:Results = New-Object 'System.Collections.Generic.List[object]'

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
        [switch]$RetryDirectoryReplication,
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
            $replicationDelay = $RetryDirectoryReplication -and
                $details.StatusCode -in @(400, 404) -and
                $details.Message -match '(?i)(does not exist|do not exist|not found|referenced.*exist)'
            $retryable = $details.StatusCode -in @(429, 500, 502, 503, 504) -or $replicationDelay
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
    Write-Log -Level WARN -Message 'The signed-in account must also be allowed by Entra roles and external collaboration settings to perform the requested changes.'
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

function Test-ExternalEmailAddress {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Address)

    $trimmedAddress = $Address.Trim()
    try {
        $parsedAddress = New-Object System.Net.Mail.MailAddress($trimmedAddress)
    }
    catch {
        return $false
    }

    if ($parsedAddress.Address -ine $trimmedAddress) {
        return $false
    }

    $separatorIndex = $trimmedAddress.LastIndexOf('@')
    if ($separatorIndex -le 0 -or $separatorIndex -ge ($trimmedAddress.Length - 1)) {
        return $false
    }

    $localPart = $trimmedAddress.Substring(0, $separatorIndex)
    if ($localPart -match '^[.-]|[.-]$') {
        return $false
    }

    # Microsoft Graph invitations reject these characters even though some are valid in RFC email addresses.
    if ($localPart -match '[~!#$%\^&*()+={}\[\]\\/|;:"<>?,]') {
        return $false
    }

    return $true
}

function New-InviteInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Address,
        [string]$Name,
        [string]$Source
    )

    [pscustomobject][ordered]@{
        EmailAddress = $Address.Trim()
        DisplayName  = if ([string]::IsNullOrWhiteSpace($Name)) { $null } else { $Name.Trim() }
        Source       = $Source
    }
}

function Import-InviteInputs {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file '$Path' was not found."
    }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) {
        throw "CSV file '$Path' does not contain any data rows."
    }

    $inputs = New-Object 'System.Collections.Generic.List[object]'
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $emailValue = $null
        foreach ($columnName in @('EmailAddress', 'InvitedUserEmailAddress', 'Mail')) {
            if ($row.PSObject.Properties.Name -contains $columnName -and
                -not [string]::IsNullOrWhiteSpace([string]$row.$columnName)) {
                $emailValue = [string]$row.$columnName
                break
            }
        }

        $displayValue = $null
        if ($row.PSObject.Properties.Name -contains 'DisplayName') {
            $displayValue = [string]$row.DisplayName
        }

        if ([string]::IsNullOrWhiteSpace($emailValue)) {
            Add-RunResult -Mode 'Csv' -Email $null -Name $displayValue -Action 'Validate' -Status 'Failed' -Message "CSV row $rowNumber has no email value."
            continue
        }

        $inputs.Add((New-InviteInput -Address $emailValue -Name $displayValue -Source "CSV row $rowNumber"))
    }

    return $inputs.ToArray()
}

function Get-InviteInputs {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Mode)

    $rawInputs = New-Object 'System.Collections.Generic.List[object]'
    switch ($Mode) {
        'Single' {
            $rawInputs.Add((New-InviteInput -Address $EmailAddress -Name $DisplayName -Source 'Parameter'))
        }
        'Multiple' {
            foreach ($address in @(ConvertTo-IdentityList -Identity $EmailAddresses)) {
                $rawInputs.Add((New-InviteInput -Address $address -Source 'Parameter'))
            }
        }
        'Csv' {
            foreach ($inputRow in @(Import-InviteInputs -Path $CsvPath)) {
                $rawInputs.Add($inputRow)
            }
        }
        default {
            throw "Unsupported invitation mode '$Mode'."
        }
    }

    $validInputs = New-Object 'System.Collections.Generic.List[object]'
    $seenEmails = @{}
    foreach ($inputRow in @($rawInputs.ToArray())) {
        $emailKey = ([string]$inputRow.EmailAddress).Trim().ToLowerInvariant()
        if (-not (Test-ExternalEmailAddress -Address $inputRow.EmailAddress)) {
            $message = "'$($inputRow.EmailAddress)' is not supported by the Microsoft Graph invitation API."
            Write-Log -Level ERROR -Message $message
            Add-RunResult -Mode $Mode -Email $inputRow.EmailAddress -Name $inputRow.DisplayName -Action 'Validate' -Status 'Failed' -Message $message
            continue
        }
        if ($seenEmails.ContainsKey($emailKey)) {
            $message = 'Duplicate email address in this run; only the first occurrence is processed.'
            Write-Log -Level WARN -Message "$($inputRow.EmailAddress) - $message"
            Add-RunResult -Mode $Mode -Email $inputRow.EmailAddress -Name $inputRow.DisplayName -Action 'Validate' -Status 'Skipped' -Message $message
            continue
        }

        $seenEmails[$emailKey] = $true
        $validInputs.Add($inputRow)
    }

    return $validInputs.ToArray()
}

function Show-InteractiveMenu {
    [CmdletBinding()]
    param()

    Write-Host ''
    Write-Host 'Create External Users' -ForegroundColor Cyan
    Write-Host 'Invites Microsoft Entra B2B guests and can add them to groups.' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  1. Invite one external user'
    Write-Host '  2. Invite multiple external users'
    Write-Host '  3. Import external users from CSV'
    Write-Host ''

    $selection = Read-Host 'Choose 1-3'
    $mode = switch ($selection) {
        '1' {
            $script:EmailAddress = Read-Host 'External user email address'
            if ([string]::IsNullOrWhiteSpace($script:EmailAddress)) {
                throw 'An email address is required.'
            }
            $enteredDisplayName = Read-Host 'Display name (optional)'
            if (-not [string]::IsNullOrWhiteSpace($enteredDisplayName)) {
                $script:DisplayName = $enteredDisplayName.Trim()
            }
            'Single'
            break
        }
        '2' {
            $enteredEmails = Read-Host 'Email addresses, separated by commas or semicolons'
            $script:EmailAddresses = @($enteredEmails -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($script:EmailAddresses.Count -eq 0) {
                throw 'At least one email address is required.'
            }
            'Multiple'
            break
        }
        '3' {
            $script:CsvPath = Read-Host 'CSV file path'
            if ([string]::IsNullOrWhiteSpace($script:CsvPath)) {
                throw 'A CSV file path is required.'
            }
            'Csv'
            break
        }
        default {
            throw "'$selection' is not valid. Run the script again and choose 1, 2, or 3."
        }
    }

    $enteredGroups = Read-Host 'Group object IDs or exact display names, separated by commas or semicolons (optional)'
    if (-not [string]::IsNullOrWhiteSpace($enteredGroups)) {
        $script:GroupId = @($enteredGroups -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $sendChoice = Read-Host 'Send the Microsoft invitation email? [Y/n]'
    if ($sendChoice -match '^(?i)n(?:o)?$') {
        $script:DoNotSendInvitationMessage = $true
    }

    return $mode
}

function Find-DirectoryUserByEmail {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Address)

    $escapedAddress = $Address.Trim().Replace("'", "''")
    $select = 'id,displayName,userPrincipalName,userType,mail,otherMails,externalUserState,accountEnabled'
    $matchesById = @{}

    # Keep the scalar mail filter and collection otherMails filter separate. Microsoft Graph
    # documents otherMails/any for external-user discovery, and separate requests avoid query
    # restrictions that can apply when collection filters are combined with other properties.
    foreach ($filter in @(
        "mail eq '$escapedAddress'"
        "otherMails/any(m:m eq '$escapedAddress')"
    )) {
        $encodedFilter = [uri]::EscapeDataString($filter)
        $uri = "$script:GraphBaseUri/users?%24count=true&%24filter=$encodedFilter&%24select=$select&%24top=10"
        foreach ($match in @(Get-GraphCollection -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' })) {
            $matchesById[[string]$match.id] = $match
        }
    }

    $matches = @($matchesById.Values)

    if ($matches.Count -gt 1) {
        $ids = ($matches | ForEach-Object { $_.id }) -join ', '
        throw "More than one directory user has email '$Address'. Matching object IDs: $ids"
    }
    if ($matches.Count -eq 1) {
        return $matches[0]
    }
    return $null
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
    $matches = @(Get-GraphCollection -Uri "$script:GraphBaseUri/groups?%24filter=$encodedFilter&%24select=$select")
    if ($matches.Count -eq 0) {
        throw "No group with the exact display name '$trimmedIdentity' was found. Try its object ID instead."
    }
    if ($matches.Count -gt 1) {
        $ids = ($matches | ForEach-Object { $_.id }) -join ', '
        throw "More than one group is named '$trimmedIdentity'. Use one of these object IDs: $ids"
    }

    return $matches[0]
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

function Resolve-RequestedGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Identity,
        [Parameter(Mandatory)][string]$Mode
    )

    $resolvedGroups = New-Object 'System.Collections.Generic.List[object]'
    $resolvedIds = @{}
    foreach ($groupIdentity in @(ConvertTo-IdentityList -Identity $Identity)) {
        try {
            Write-Log -Level INFO -Message "Resolving group '$groupIdentity'."
            $group = Resolve-GroupObject -Identity $groupIdentity
            if ($resolvedIds.ContainsKey([string]$group.id)) {
                Write-Log -Level WARN -Message "$($group.displayName) was specified more than once and will be processed once."
                continue
            }

            $resolvedIds[[string]$group.id] = $true
            $eligibility = Get-GroupEligibility -Group $group
            if (-not $eligibility.Eligible) {
                Write-Log -Level WARN -Message "$($group.displayName) - $($eligibility.Reason)"
                Add-RunResult -Mode $Mode -Group $group -GroupKind $eligibility.Kind -Action 'ResolveGroup' -Status 'Skipped' -Message $eligibility.Reason
                continue
            }

            $resolvedGroups.Add($group)
            Write-Log -Level SUCCESS -Message "Resolved eligible $($eligibility.Kind) group '$($group.displayName)'."
        }
        catch {
            $details = Get-GraphErrorInfo -ErrorRecord $_
            $message = "Could not resolve group '$groupIdentity': $($details.Message)"
            Write-Log -Level ERROR -Message $message
            Add-RunResult -Mode $Mode -Action 'ResolveGroup' -Status 'Failed' -Message $message
        }
    }

    return $resolvedGroups.ToArray()
}

function New-GuestInvitation {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$InputRow)

    $body = [ordered]@{
        invitedUserEmailAddress = [string]$InputRow.EmailAddress
        inviteRedirectUrl       = $InviteRedirectUrl.AbsoluteUri
        sendInvitationMessage   = -not [bool]$DoNotSendInvitationMessage
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$InputRow.DisplayName)) {
        $body.invitedUserDisplayName = [string]$InputRow.DisplayName
    }

    if (-not $DoNotSendInvitationMessage -and
        (-not [string]::IsNullOrWhiteSpace($CustomizedMessageBody) -or
         -not [string]::IsNullOrWhiteSpace($MessageLanguage))) {
        $messageInfo = [ordered]@{}
        if (-not [string]::IsNullOrWhiteSpace($CustomizedMessageBody)) {
            $messageInfo.customizedMessageBody = $CustomizedMessageBody
        }
        if (-not [string]::IsNullOrWhiteSpace($MessageLanguage)) {
            $messageInfo.messageLanguage = $MessageLanguage
        }
        $body.invitedUserMessageInfo = $messageInfo
    }

    return Invoke-GraphRequestWithRetry -Method POST -Uri "$script:GraphBaseUri/invitations" -Body $body
}

function Test-UserDirectGroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$User,
        [Parameter(Mandatory)][object]$Group
    )

    # List members is a direct-membership endpoint. Filtering by ID avoids downloading a
    # potentially large group and does not confuse inherited membership with direct membership.
    $filter = [uri]::EscapeDataString("id eq '$($User.id)'")
    $uri = "$script:GraphBaseUri/groups/$($Group.id)/members?%24count=true&%24filter=$filter&%24select=id&%24top=1"
    $matches = @(Get-GraphCollection -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' })
    return $matches.Count -gt 0
}

function Add-UserToGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$User,
        [Parameter(Mandatory)][object]$Group
    )

    $uri = "$script:GraphBaseUri/groups/$($Group.id)/members/`$ref"
    $body = @{ '@odata.id' = "$script:GraphBaseUri/directoryObjects/$($User.id)" }
    Invoke-GraphRequestWithRetry -Method POST -Uri $uri -Body $body -RetryDirectoryReplication | Out-Null
}

function Add-RunResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mode,
        [string]$Email,
        [string]$Name,
        [object]$User,
        [object]$Group,
        [string]$GroupKind,
        [string]$InvitationStatus,
        [string]$InviteRedeemUrl,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message
    )

    $script:Results.Add([pscustomobject][ordered]@{
        Timestamp          = (Get-Date).ToString('o')
        Mode               = $Mode
        EmailAddress       = $Email
        RequestedName      = $Name
        UserId             = if ($null -ne $User) { [string]$User.id } else { $null }
        UserPrincipalName  = if ($null -ne $User) { [string]$User.userPrincipalName } else { $null }
        UserType           = if ($null -ne $User) { [string]$User.userType } else { $null }
        InvitationStatus   = $InvitationStatus
        SendInvitationMessage = if ($Action -eq 'Invite' -and $Status -in @('Invited', 'WouldInvite')) { -not [bool]$DoNotSendInvitationMessage } else { $null }
        InviteRedirectUrl  = if ($Action -eq 'Invite' -and $Status -in @('Invited', 'WouldInvite')) { $InviteRedirectUrl.AbsoluteUri } else { $null }
        InviteRedeemUrl    = $InviteRedeemUrl
        GroupId            = if ($null -ne $Group) { [string]$Group.id } else { $null }
        GroupDisplayName   = if ($null -ne $Group) { [string]$Group.displayName } else { $null }
        GroupKind          = $GroupKind
        Action             = $Action
        Status             = $Status
        Message            = $Message
    })
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
        $resolvedPath = Join-Path $Path ("ExternalUserInvitations-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
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
    foreach ($status in @('Invited', 'WouldInvite', 'ExistingGuest', 'AddedToGroup', 'WouldAddToGroup', 'AlreadyMember', 'Skipped', 'Declined', 'Failed')) {
        $counts[$status] = @($script:Results.ToArray() | Where-Object { $_.Status -eq $status }).Count
    }

    Write-Host ''
    Write-Host 'Run summary' -ForegroundColor Cyan
    Write-Host ('  Invited:             {0}' -f $counts.Invited) -ForegroundColor Green
    Write-Host ('  Would invite:        {0}' -f $counts.WouldInvite) -ForegroundColor Cyan
    Write-Host ('  Existing guests:     {0}' -f $counts.ExistingGuest) -ForegroundColor Gray
    Write-Host ('  Group memberships:   {0}' -f $counts.AddedToGroup) -ForegroundColor Green
    Write-Host ('  Would add to groups: {0}' -f $counts.WouldAddToGroup) -ForegroundColor Cyan
    Write-Host ('  Already members:     {0}' -f $counts.AlreadyMember) -ForegroundColor Gray
    Write-Host ('  Skipped:             {0}' -f $counts.Skipped) -ForegroundColor Yellow
    Write-Host ('  Declined:            {0}' -f $counts.Declined) -ForegroundColor Yellow
    Write-Host ('  Failed:              {0}' -f $counts.Failed) -ForegroundColor Red
    Write-Host ('  Duration:             {0:hh\:mm\:ss}' -f $duration)
}

$fatalError = $null
try {
    Write-Host ''
    Write-Host 'Microsoft Graph - Create External Users' -ForegroundColor Cyan
    Write-Host 'Creates B2B guest invitations and optional direct group memberships.' -ForegroundColor Yellow
    Write-Host ''

    $mode = $PSCmdlet.ParameterSetName
    if ($mode -eq 'Interactive') {
        $mode = Show-InteractiveMenu
    }

    $inviteInputs = @(Get-InviteInputs -Mode $mode)
    if ($inviteInputs.Count -eq 0) {
        throw 'There are no valid external users to process.'
    }

    $groupIdentities = @()
    if ($null -ne $GroupId -and @($GroupId).Count -gt 0) {
        $groupIdentities = @(ConvertTo-IdentityList -Identity $GroupId)
    }

    Write-Host 'Run selection' -ForegroundColor Cyan
    Write-Host "  External users:       $($inviteInputs.Count)"
    Write-Host "  Requested groups:     $($groupIdentities.Count)"
    Write-Host "  Send invitation mail: $(-not [bool]$DoNotSendInvitationMessage)"
    Write-Host "  Redirect URL:         $($InviteRedirectUrl.AbsoluteUri)"

    if ($inviteInputs.Count -gt 1 -and -not $Force -and -not $WhatIfPreference) {
        Write-Host ''
        $continueChoice = Read-Host "Continue with $($inviteInputs.Count) external users? [y/N]"
        if ($continueChoice -notmatch '^(?i)y(?:es)?$') {
            foreach ($inputRow in $inviteInputs) {
                Add-RunResult -Mode $mode -Email $inputRow.EmailAddress -Name $inputRow.DisplayName -Action 'Invite' -Status 'Declined' -Message 'The multi-user run was declined by the operator.'
            }
            throw [System.OperationCanceledException]::new('The multi-user run was declined by the operator.')
        }
    }

    $requiredScopes = @('User.Invite.All', 'User.Read.All')
    if ($groupIdentities.Count -gt 0) {
        $requiredScopes += 'GroupMember.ReadWrite.All'
    }
    Connect-RequiredGraphSession -Scopes $requiredScopes

    $groups = @()
    if ($groupIdentities.Count -gt 0) {
        Write-Log -Level STEP -Message "Resolving $($groupIdentities.Count) requested group(s)."
        $groups = @(Resolve-RequestedGroups -Identity $groupIdentities -Mode $mode)
        if ($groups.Count -eq 0) {
            throw 'None of the requested groups resolved to an eligible group. No invitations were created.'
        }
    }

    $inputIndex = 0
    foreach ($inputRow in $inviteInputs) {
        $inputIndex++
        $email = [string]$inputRow.EmailAddress
        Write-Host ''
        Write-Host "[$inputIndex/$($inviteInputs.Count)] $email" -ForegroundColor Cyan

        $user = $null
        $invitation = $null
        $canProcessGroups = $false
        try {
            Write-Log -Level INFO -Message "Checking whether '$email' already exists in the directory."
            $user = Find-DirectoryUserByEmail -Address $email
            if ($null -ne $user) {
                if ([string]$user.userType -ine 'Guest') {
                    $message = "A directory user with this email already exists as userType '$($user.userType)'. It was not treated as an external guest."
                    Write-Log -Level WARN -Message $message
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -Action 'Invite' -Status 'Skipped' -Message $message
                    continue
                }

                $message = 'The guest already exists; no new invitation was created.'
                Write-Log -Level INFO -Message $message
                Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -InvitationStatus ([string]$user.externalUserState) -Action 'Invite' -Status 'ExistingGuest' -Message $message
                $canProcessGroups = $true
            }
            else {
                $targetLabel = if ([string]::IsNullOrWhiteSpace([string]$inputRow.DisplayName)) {
                    $email
                }
                else {
                    "$($inputRow.DisplayName) <$email>"
                }
                $operation = 'create a Microsoft Entra B2B guest invitation'
                if ($PSCmdlet.ShouldProcess($targetLabel, $operation)) {
                    $invitation = New-GuestInvitation -InputRow $inputRow
                    if ($null -eq $invitation.invitedUser -or [string]::IsNullOrWhiteSpace([string]$invitation.invitedUser.id)) {
                        throw 'Microsoft Graph created the invitation but did not return an invited user object ID.'
                    }

                    $user = [pscustomobject]@{
                        id                = [string]$invitation.invitedUser.id
                        displayName       = [string]$inputRow.DisplayName
                        userPrincipalName = [string]$invitation.invitedUser.userPrincipalName
                        userType          = 'Guest'
                    }
                    $message = 'External guest invitation created successfully.'
                    Write-Log -Level SUCCESS -Message $message
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -InvitationStatus ([string]$invitation.status) -InviteRedeemUrl ([string]$invitation.inviteRedeemUrl) -Action 'Invite' -Status 'Invited' -Message $message
                    $canProcessGroups = $true

                    if ($DoNotSendInvitationMessage -and -not [string]::IsNullOrWhiteSpace([string]$invitation.inviteRedeemUrl)) {
                        Write-Host "  Redemption URL: $($invitation.inviteRedeemUrl)" -ForegroundColor Yellow
                    }
                }
                else {
                    $status = if ($WhatIfPreference) { 'WouldInvite' } else { 'Declined' }
                    $message = if ($WhatIfPreference) { 'Would create an external guest invitation.' } else { 'Invitation declined by ShouldProcess confirmation.' }
                    Write-Log -Level INFO -Message $message
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -Action 'Invite' -Status $status -Message $message
                    if ($WhatIfPreference) {
                        $canProcessGroups = $true
                    }
                }
            }
        }
        catch {
            $details = Get-GraphErrorInfo -ErrorRecord $_
            $message = "Could not process the invitation: $($details.Message)"
            Write-Log -Level ERROR -Message $message
            Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -Action 'Invite' -Status 'Failed' -Message $message
            continue
        }

        if (-not $canProcessGroups -or $groups.Count -eq 0) {
            continue
        }

        foreach ($group in $groups) {
            $groupKind = Get-GroupKind -Group $group
            try {
                if ($null -ne $user -and (Test-UserDirectGroupMember -User $user -Group $group)) {
                    $message = 'The guest is already a direct member.'
                    Write-Log -Level INFO -Message "$($group.displayName) - $message"
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -Group $group -GroupKind $groupKind -Action 'AddToGroup' -Status 'AlreadyMember' -Message $message
                    continue
                }

                $groupTarget = "$email -> $($group.displayName) [$($group.id)]"
                if ($PSCmdlet.ShouldProcess($groupTarget, 'add the guest as a direct group member')) {
                    if ($null -eq $user) {
                        throw 'The guest user object is unavailable, so the group membership cannot be added.'
                    }
                    Add-UserToGroup -User $user -Group $group
                    $message = 'Direct group membership added successfully.'
                    Write-Log -Level SUCCESS -Message "$($group.displayName) - $message"
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -Group $group -GroupKind $groupKind -Action 'AddToGroup' -Status 'AddedToGroup' -Message $message
                }
                else {
                    $status = if ($WhatIfPreference) { 'WouldAddToGroup' } else { 'Declined' }
                    $message = if ($WhatIfPreference) { 'Would add the guest as a direct group member after invitation.' } else { 'Group assignment declined by ShouldProcess confirmation.' }
                    Write-Log -Level INFO -Message "$($group.displayName) - $message"
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -Group $group -GroupKind $groupKind -Action 'AddToGroup' -Status $status -Message $message
                }
            }
            catch {
                $details = Get-GraphErrorInfo -ErrorRecord $_
                if ($details.StatusCode -eq 400 -and $details.Message -match '(?i)(already exist|already a member)') {
                    $message = 'Microsoft Graph reports that the guest is already a direct member.'
                    Write-Log -Level INFO -Message "$($group.displayName) - $message"
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -Group $group -GroupKind $groupKind -Action 'AddToGroup' -Status 'AlreadyMember' -Message $message
                }
                else {
                    $message = "Microsoft Graph could not add the group membership: $($details.Message)"
                    Write-Log -Level ERROR -Message "$($group.displayName) - $message"
                    Add-RunResult -Mode $mode -Email $email -Name $inputRow.DisplayName -User $user -Group $group -GroupKind $groupKind -Action 'AddToGroup' -Status 'Failed' -Message $message
                }
            }
        }
    }
}
catch [System.OperationCanceledException] {
    Write-Log -Level WARN -Message $_.Exception.Message
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
