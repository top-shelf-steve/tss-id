<#
.SYNOPSIS
Registers users' work email addresses as their Microsoft Entra email authentication method.

.DESCRIPTION
Uses Microsoft Graph v1.0 to create or update the email authentication method used for
self-service password reset (SSPR). The script supports a single user, multiple users,
transitive members of a group, a CSV file, or all users in the tenant.

Run the script without a targeting parameter to use the interactive menu. By default,
disabled accounts and guest users are skipped. The directory mail property is preferred;
the userPrincipalName is used when mail is empty unless -EmailSource Mail is specified.

The signed-in administrator needs the Authentication Administrator or Privileged
Authentication Administrator role and consent to the requested Microsoft Graph scopes.

.PARAMETER UserId
Object ID or user principal name of one user.

.PARAMETER UserIds
Object IDs or user principal names of multiple users.

.PARAMETER GroupId
Object ID or exact display name of a group. Transitive (nested) user members are included.

.PARAMETER CsvPath
Path to a CSV. Each row must identify a user with UserId, UserPrincipalName, Identity, or
Mail. An optional EmailAddress column overrides the address obtained from Microsoft Entra.

.PARAMETER AllUsers
Targets all users returned by Microsoft Graph. Disabled accounts and guests are still
skipped unless their corresponding include switches are supplied.

.PARAMETER EmailSource
MailThenUserPrincipalName (default), Mail, or UserPrincipalName.

.PARAMETER SkipExistingEmailMethod
Skips a user when any email authentication method is already registered, even when its
address differs from the desired address. Combine with -AllUsers -EmailSource Mail to
populate only missing methods from each user's directory mail property.

.PARAMETER ExportPath
Exports results without prompting. Use a .json extension for JSON; all other extensions
produce CSV. If this is a directory, a timestamped CSV filename is created in it.
Result exports are written even when -WhatIf is active; -WhatIf suppresses Graph changes.

.EXAMPLE
.\Set-EmailAuthenticationMethod.ps1 -UserId alex@contoso.com

.EXAMPLE
.\Set-EmailAuthenticationMethod.ps1 -UserIds user1@contoso.com,user2@contoso.com -WhatIf

.EXAMPLE
.\Set-EmailAuthenticationMethod.ps1 -GroupId 'Operations' -Force -ExportPath C:\Reports

.EXAMPLE
.\Set-EmailAuthenticationMethod.ps1 -CsvPath .\users.csv

.EXAMPLE
.\Set-EmailAuthenticationMethod.ps1 -AllUsers -SkipExistingEmailMethod -EmailSource Mail -Force

.NOTES
Email authentication methods are available for SSPR only; they are not an MFA sign-in
method. Requires the Microsoft.Graph.Authentication PowerShell module.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Interactive', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Single', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$UserId,

    [Parameter(Mandatory, ParameterSetName = 'Multiple')]
    [ValidateNotNullOrEmpty()]
    [string[]]$UserIds,

    [Parameter(Mandatory, ParameterSetName = 'Group')]
    [ValidateNotNullOrEmpty()]
    [string]$GroupId,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllUsers,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [ValidateSet('MailThenUserPrincipalName', 'Mail', 'UserPrincipalName')]
    [string]$EmailSource = 'MailThenUserPrincipalName',

    [Alias('OnlyIfMissing')]
    [switch]$SkipExistingEmailMethod,

    [switch]$IncludeGuests,

    [switch]$IncludeDisabled,

    [switch]$Force,

    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [string]$ExportPath,

    [switch]$NoExportPrompt,

    [switch]$DisconnectWhenDone
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:GraphBaseUri = 'https://graph.microsoft.com/v1.0'
$script:EmailMethodId = '3ddfcfc8-9383-446f-83cc-3ab9be4be18f'
$script:Results = New-Object 'System.Collections.Generic.List[object]'
$script:ConnectedByScript = $false
$script:RunStarted = Get-Date

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

function Add-RunResult {
    [CmdletBinding()]
    param(
        [string]$Source,
        [string]$InputIdentity,
        [object]$User,
        [string]$RequestedEmail,
        [string]$PreviousEmail,
        [string]$Action,
        [string]$Status,
        [string]$Message
    )

    $changeFrom = $null
    $changeTo = $null
    if ($Action -in @('Create', 'Update')) {
        $changeFrom = if ([string]::IsNullOrWhiteSpace($PreviousEmail)) { '<not registered>' } else { $PreviousEmail }
        $changeTo = $RequestedEmail
    }

    $script:Results.Add([pscustomobject][ordered]@{
        Timestamp         = (Get-Date).ToString('o')
        Source            = $Source
        InputIdentity     = $InputIdentity
        UserId            = if ($null -ne $User) { [string]$User.id } else { $null }
        DisplayName       = if ($null -ne $User) { [string]$User.displayName } else { $null }
        UserPrincipalName = if ($null -ne $User) { [string]$User.userPrincipalName } else { $null }
        RequestedEmail    = $RequestedEmail
        PreviousEmail     = $PreviousEmail
        ChangeFrom        = $changeFrom
        ChangeTo          = $changeTo
        Action            = $Action
        Status            = $Status
        Message           = $Message
    })
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
        # Preserve the original exception details when the response shape differs.
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
        [Parameter(Mandatory)][ValidateSet('GET', 'POST', 'PATCH')][string]$Method,
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
    $page = 0

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $page++
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
        Write-Log -Level DEBUG -Message "Retrieved page $page; running item count: $($items.Count)."
    }

    return $items.ToArray()
}

function Connect-RequiredGraphSession {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Scopes)

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue) -or
        -not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        throw "The Microsoft.Graph.Authentication module is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    $missingScopes = @()
    $tenantMismatch = $false
    if ($null -ne $context -and -not [string]::IsNullOrWhiteSpace($TenantId) -and
        [string]$context.TenantId -ine $TenantId) {
        $tenantMismatch = $true
        Write-Log -Level INFO -Message "The active Graph context is for tenant $($context.TenantId); reconnecting to requested tenant $TenantId."
    }

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
    Write-Log -Level WARN -Message 'The signed-in account must hold Authentication Administrator or Privileged Authentication Administrator.'
}

function Get-DirectoryUser {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Identity)

    $encodedIdentity = [uri]::EscapeDataString($Identity.Trim())
    $uri = "$script:GraphBaseUri/users/$encodedIdentity`?%24select=id,displayName,userPrincipalName,mail,accountEnabled,userType"
    return Invoke-GraphRequestWithRetry -Method GET -Uri $uri
}

function Resolve-GroupObject {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Identity)

    $trimmedIdentity = $Identity.Trim()
    if ($trimmedIdentity -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        $encodedIdentity = [uri]::EscapeDataString($trimmedIdentity)
        return Invoke-GraphRequestWithRetry -Method GET -Uri "$script:GraphBaseUri/groups/$encodedIdentity`?%24select=id,displayName"
    }

    $escapedName = $trimmedIdentity.Replace("'", "''")
    $encodedFilter = [uri]::EscapeDataString("displayName eq '$escapedName'")
    $groupSearchResults = @(Get-GraphCollection -Uri "$script:GraphBaseUri/groups?%24filter=$encodedFilter&%24select=id,displayName")
    if ($groupSearchResults.Count -eq 0) {
        throw "No group with the exact display name '$trimmedIdentity' was found. Try its object ID instead."
    }
    if ($groupSearchResults.Count -gt 1) {
        $ids = ($groupSearchResults | ForEach-Object { $_.id }) -join ', '
        throw "More than one group is named '$trimmedIdentity'. Use one of these object IDs: $ids"
    }
    return $groupSearchResults[0]
}

function Test-EmailAddress {
    [CmdletBinding()]
    param([string]$Address)

    if ([string]::IsNullOrWhiteSpace($Address)) {
        return $false
    }

    try {
        $parsed = New-Object System.Net.Mail.MailAddress($Address.Trim())
        return $parsed.Address -eq $Address.Trim()
    }
    catch {
        return $false
    }
}

function Get-DesiredEmailAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$User,
        [string]$Override
    )

    if (-not [string]::IsNullOrWhiteSpace($Override)) {
        return $Override.Trim()
    }

    switch ($EmailSource) {
        'Mail' {
            return [string]$User.mail
        }
        'UserPrincipalName' {
            return [string]$User.userPrincipalName
        }
        default {
            if (-not [string]::IsNullOrWhiteSpace([string]$User.mail)) {
                return [string]$User.mail
            }
            return [string]$User.userPrincipalName
        }
    }
}

function Show-InteractiveMenu {
    [CmdletBinding()]
    param()

    Write-Host ''
    Write-Host 'Set Email Authentication Method' -ForegroundColor Cyan
    Write-Host 'This manages the email method used for SSPR; it is not an MFA sign-in method.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  1. Single user'
    Write-Host '  2. Several users'
    Write-Host '  3. All users in a group (including nested members)'
    Write-Host '  4. Users from a CSV file'
    Write-Host '  5. All tenant users (create or update)'
    Write-Host '  6. All tenant users missing an email method (use directory mail)'
    Write-Host ''

    $selection = Read-Host 'Choose 1-6'
    switch ($selection) {
        '1' {
            $script:UserId = Read-Host 'User object ID or user principal name'
            return 'Single'
        }
        '2' {
            $enteredUsers = Read-Host 'User IDs/UPNs separated by commas or semicolons'
            $script:UserIds = @($enteredUsers -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            return 'Multiple'
        }
        '3' {
            $script:GroupId = Read-Host 'Group object ID or exact display name'
            return 'Group'
        }
        '4' {
            $script:CsvPath = Read-Host 'CSV path'
            return 'Csv'
        }
        '5' {
            $script:AllUsers = $true
            return 'All'
        }
        '6' {
            $script:AllUsers = $true
            $script:EmailSource = 'Mail'
            $script:SkipExistingEmailMethod = $true
            return 'All'
        }
        default {
            throw "'$selection' is not a valid selection. Run the script again and choose a number from 1 through 6."
        }
    }
}

function New-TargetRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$User,
        [Parameter(Mandatory)][string]$Source,
        [string]$InputIdentity,
        [string]$EmailOverride
    )

    [pscustomobject]@{
        User          = $User
        Source        = $Source
        InputIdentity = $InputIdentity
        EmailOverride = $EmailOverride
    }
}

function Get-TargetsFromCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file not found: $Path"
    }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) {
        throw "The CSV file '$Path' contains no data rows."
    }

    Write-Log -Level INFO -Message "Loaded $($rows.Count) row(s) from '$Path'."
    $targets = New-Object 'System.Collections.Generic.List[object]'
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $identity = $null
        foreach ($column in @('UserId', 'UserPrincipalName', 'Identity', 'Mail')) {
            if ($row.PSObject.Properties.Name -contains $column -and
                -not [string]::IsNullOrWhiteSpace([string]$row.$column)) {
                $identity = [string]$row.$column
                break
            }
        }

        $override = $null
        if ($row.PSObject.Properties.Name -contains 'EmailAddress') {
            $override = [string]$row.EmailAddress
        }

        if ([string]::IsNullOrWhiteSpace($identity)) {
            $message = "CSV row $rowNumber has no UserId, UserPrincipalName, Identity, or Mail value."
            Write-Log -Level ERROR -Message $message
            Add-RunResult -Source 'CSV' -InputIdentity "Row $rowNumber" -Action 'Resolve' -Status 'Failed' -Message $message
            continue
        }

        try {
            Write-Log -Level INFO -Message "Resolving CSV row $rowNumber user '$identity'."
            $user = Get-DirectoryUser -Identity $identity
            $targets.Add((New-TargetRecord -User $user -Source 'CSV' -InputIdentity $identity -EmailOverride $override))
        }
        catch {
            $details = Get-GraphErrorInfo -ErrorRecord $_
            $message = "Could not resolve '$identity' from CSV row ${rowNumber}: $($details.Message)"
            Write-Log -Level ERROR -Message $message
            Add-RunResult -Source 'CSV' -InputIdentity $identity -Action 'Resolve' -Status 'Failed' -Message $message
        }
    }

    return $targets.ToArray()
}

function Remove-DuplicateTargets {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Targets)

    $seen = @{}
    $unique = New-Object 'System.Collections.Generic.List[object]'
    foreach ($target in $Targets) {
        $key = ([string]$target.User.id).ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            Write-Log -Level WARN -Message "Skipping duplicate target $($target.User.userPrincipalName) ($($target.User.id))."
            continue
        }
        $seen[$key] = $true
        $unique.Add($target)
    }
    return $unique.ToArray()
}

function Export-RunResults {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $targetPath = $resolvedPath
    if (Test-Path -LiteralPath $resolvedPath -PathType Container) {
        $filename = 'EmailAuthenticationMethod-Results-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
        $targetPath = Join-Path -Path $resolvedPath -ChildPath $filename
    }
    else {
        $parent = Split-Path -Parent $resolvedPath
        if ([string]::IsNullOrWhiteSpace($parent)) {
            $parent = (Get-Location).Path
        }
        if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
            # A dry run applies to Graph changes, not to the explicitly requested report.
            New-Item -ItemType Directory -Path $parent -Force -WhatIf:$false | Out-Null
            Write-Log -Level INFO -Message "Created export directory '$parent'."
        }
    }

    $resultArray = $script:Results.ToArray()
    if ([IO.Path]::GetExtension($targetPath) -ieq '.json') {
        ConvertTo-Json -InputObject $resultArray -Depth 6 |
            Set-Content -LiteralPath $targetPath -Encoding UTF8 -WhatIf:$false
    }
    else {
        $resultArray |
            Export-Csv -LiteralPath $targetPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false
    }

    if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
        throw "The export command completed, but the expected file was not found at '$targetPath'."
    }
    Write-Log -Level SUCCESS -Message "Exported $($script:Results.Count) result record(s) to '$targetPath'."
}

function Show-RunSummary {
    [CmdletBinding()]
    param()

    $duration = (Get-Date) - $script:RunStarted
    $counts = @{}
    foreach ($status in @('Succeeded', 'Unchanged', 'WouldChange', 'Skipped', 'Failed')) {
        $counts[$status] = @($script:Results.ToArray() | Where-Object { $_.Status -eq $status }).Count
    }

    Write-Host ''
    Write-Host 'Run summary' -ForegroundColor Cyan
    Write-Host ('  Succeeded:   {0}' -f $counts.Succeeded) -ForegroundColor Green
    Write-Host ('  Unchanged:   {0}' -f $counts.Unchanged) -ForegroundColor Gray
    Write-Host ('  Would change:{0,4}' -f $counts.WouldChange) -ForegroundColor Cyan
    Write-Host ('  Skipped:     {0}' -f $counts.Skipped) -ForegroundColor Yellow
    Write-Host ('  Failed:      {0}' -f $counts.Failed) -ForegroundColor Red
    Write-Host ('  Duration:    {0:hh\:mm\:ss}' -f $duration)
}

$fatalError = $null
try {
    Write-Host ''
    Write-Host 'Microsoft Graph - Set Email Authentication Method' -ForegroundColor Cyan
    Write-Host 'Email is registered for SSPR only; it is not an MFA sign-in method.' -ForegroundColor Yellow
    Write-Host ''

    $mode = $PSCmdlet.ParameterSetName
    if ($mode -eq 'Interactive') {
        $mode = Show-InteractiveMenu
    }

    $scopes = @('User.Read.All', 'UserAuthenticationMethod.ReadWrite.All')
    if ($mode -eq 'Group') {
        $scopes += 'GroupMember.Read.All'
    }
    Connect-RequiredGraphSession -Scopes $scopes

    $targets = New-Object 'System.Collections.Generic.List[object]'
    switch ($mode) {
        'Single' {
            Write-Log -Level STEP -Message "Resolving user '$UserId'."
            $user = Get-DirectoryUser -Identity $UserId
            $targets.Add((New-TargetRecord -User $user -Source 'Single' -InputIdentity $UserId))
        }
        'Multiple' {
            $identities = @($UserIds | ForEach-Object { $_ -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            foreach ($identity in $identities) {
                try {
                    Write-Log -Level INFO -Message "Resolving user '$identity'."
                    $user = Get-DirectoryUser -Identity $identity
                    $targets.Add((New-TargetRecord -User $user -Source 'Multiple' -InputIdentity $identity))
                }
                catch {
                    $details = Get-GraphErrorInfo -ErrorRecord $_
                    $message = "Could not resolve '$identity': $($details.Message)"
                    Write-Log -Level ERROR -Message $message
                    Add-RunResult -Source 'Multiple' -InputIdentity $identity -Action 'Resolve' -Status 'Failed' -Message $message
                }
            }
        }
        'Group' {
            Write-Log -Level STEP -Message "Resolving group '$GroupId'."
            $group = Resolve-GroupObject -Identity $GroupId
            Write-Log -Level SUCCESS -Message "Resolved group '$($group.displayName)' ($($group.id))."
            $groupUri = "$script:GraphBaseUri/groups/$($group.id)/transitiveMembers/microsoft.graph.user?%24select=id,displayName,userPrincipalName,mail,accountEnabled,userType&%24top=999"
            $members = @(Get-GraphCollection -Uri $groupUri -Headers @{ ConsistencyLevel = 'eventual' })
            foreach ($member in $members) {
                $targets.Add((New-TargetRecord -User $member -Source "Group: $($group.displayName)" -InputIdentity $member.userPrincipalName))
            }
        }
        'Csv' {
            foreach ($target in @(Get-TargetsFromCsv -Path $CsvPath)) {
                $targets.Add($target)
            }
        }
        'All' {
            Write-Log -Level STEP -Message 'Retrieving all tenant users from Microsoft Graph.'
            $allUri = "$script:GraphBaseUri/users?%24select=id,displayName,userPrincipalName,mail,accountEnabled,userType&%24top=999"
            $users = @(Get-GraphCollection -Uri $allUri)
            foreach ($user in $users) {
                $targets.Add((New-TargetRecord -User $user -Source 'AllUsers' -InputIdentity $user.userPrincipalName))
            }
        }
    }

    $uniqueTargets = @(Remove-DuplicateTargets -Targets $targets.ToArray())
    Write-Log -Level STEP -Message "Resolved $($uniqueTargets.Count) unique user target(s)."

    if ($uniqueTargets.Count -eq 0) {
        Write-Log -Level WARN -Message 'There are no resolved users to process.'
    }
    elseif ($uniqueTargets.Count -gt 1 -and -not $Force -and -not $WhatIfPreference) {
        Write-Host ''
        Write-Host "About to evaluate and potentially change $($uniqueTargets.Count) users." -ForegroundColor Yellow
        $confirmation = Read-Host "Type YES to continue"
        if ($confirmation -cne 'YES') {
            foreach ($target in $uniqueTargets) {
                Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $target.User -Action 'None' -Status 'Skipped' -Message 'Bulk run was cancelled at the confirmation prompt.'
            }
            $uniqueTargets = @()
            Write-Log -Level WARN -Message 'Bulk run cancelled; no authentication methods were changed.'
        }
    }

    $currentIndex = 0
    foreach ($target in $uniqueTargets) {
        $currentIndex++
        $user = $target.User
        $label = "$($user.displayName) <$($user.userPrincipalName)>"
        Write-Log -Level STEP -Message "[$currentIndex/$($uniqueTargets.Count)] Processing $label."

        if (-not $IncludeGuests -and [string]$user.userType -eq 'Guest') {
            $message = 'Guest user skipped. Use -IncludeGuests to include guest accounts.'
            Write-Log -Level WARN -Message "$label - $message"
            Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -Action 'None' -Status 'Skipped' -Message $message
            continue
        }
        if (-not $IncludeDisabled -and $null -ne $user.accountEnabled -and -not [bool]$user.accountEnabled) {
            $message = 'Disabled user skipped. Use -IncludeDisabled to include disabled accounts.'
            Write-Log -Level WARN -Message "$label - $message"
            Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -Action 'None' -Status 'Skipped' -Message $message
            continue
        }

        $desiredEmail = $null
        if (-not $SkipExistingEmailMethod) {
            $desiredEmail = Get-DesiredEmailAddress -User $user -Override $target.EmailOverride
            if (-not (Test-EmailAddress -Address $desiredEmail)) {
                $message = "No valid desired email address was found using EmailSource '$EmailSource'."
                Write-Log -Level ERROR -Message "$label - $message"
                Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -RequestedEmail $desiredEmail -Action 'Validate' -Status 'Failed' -Message $message
                continue
            }
        }

        try {
            $methodUri = "$script:GraphBaseUri/users/$($user.id)/authentication/emailMethods"
            $methods = @(Get-GraphCollection -Uri $methodUri)
            $existing = if ($methods.Count -gt 0) { $methods[0] } else { $null }
            $previousEmail = if ($null -ne $existing) { [string]$existing.emailAddress } else { $null }

            if ($SkipExistingEmailMethod -and $null -ne $existing) {
                $message = 'An email authentication method is already registered; it was left unchanged.'
                Write-Log -Level INFO -Message "$label - $message"
                Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -PreviousEmail $previousEmail -Action 'None' -Status 'Skipped' -Message $message
                continue
            }

            if ($SkipExistingEmailMethod) {
                $desiredEmail = Get-DesiredEmailAddress -User $user -Override $target.EmailOverride
                if (-not (Test-EmailAddress -Address $desiredEmail)) {
                    $message = "No valid desired email address was found using EmailSource '$EmailSource'."
                    Write-Log -Level ERROR -Message "$label - $message"
                    Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -RequestedEmail $desiredEmail -Action 'Validate' -Status 'Failed' -Message $message
                    continue
                }
            }

            Write-Log -Level INFO -Message "$label - desired SSPR email is '$desiredEmail'."

            if ($null -ne $existing -and $previousEmail -ieq $desiredEmail) {
                $message = 'The registered email authentication method is already correct.'
                Write-Log -Level SUCCESS -Message "$label - $message"
                Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -RequestedEmail $desiredEmail -PreviousEmail $previousEmail -Action 'None' -Status 'Unchanged' -Message $message
                continue
            }

            $action = if ($null -eq $existing) { 'Create' } else { 'Update' }
            $description = if ($action -eq 'Create') {
                "create email authentication method '$desiredEmail'"
            }
            else {
                "update email authentication method from '$previousEmail' to '$desiredEmail'"
            }

            if ($PSCmdlet.ShouldProcess($label, $description)) {
                if ($action -eq 'Create') {
                    Invoke-GraphRequestWithRetry -Method POST -Uri $methodUri -Body @{ emailAddress = $desiredEmail } | Out-Null
                }
                else {
                    $existingMethodId = if ([string]::IsNullOrWhiteSpace([string]$existing.id)) { $script:EmailMethodId } else { [string]$existing.id }
                    $updateUri = "$methodUri/$existingMethodId"
                    Invoke-GraphRequestWithRetry -Method PATCH -Uri $updateUri -Body @{ emailAddress = $desiredEmail } | Out-Null
                }

                $message = "$action completed successfully."
                Write-Log -Level SUCCESS -Message "$label - $message"
                Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -RequestedEmail $desiredEmail -PreviousEmail $previousEmail -Action $action -Status 'Succeeded' -Message $message
            }
            else {
                $status = if ($WhatIfPreference) { 'WouldChange' } else { 'Skipped' }
                if ($WhatIfPreference) {
                    $displayPreviousEmail = if ([string]::IsNullOrWhiteSpace($previousEmail)) { '<not registered>' } else { $previousEmail }
                    $message = "Would change SSPR email from '$displayPreviousEmail' to '$desiredEmail' ($action)."
                }
                else {
                    $message = 'Change declined by ShouldProcess confirmation.'
                }
                Write-Log -Level INFO -Message "$label - $message"
                Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -RequestedEmail $desiredEmail -PreviousEmail $previousEmail -Action $action -Status $status -Message $message
            }
        }
        catch {
            $details = Get-GraphErrorInfo -ErrorRecord $_
            $message = "Microsoft Graph operation failed: $($details.Message)"
            Write-Log -Level ERROR -Message "$label - $message"
            Add-RunResult -Source $target.Source -InputIdentity $target.InputIdentity -User $user -RequestedEmail $desiredEmail -Action 'Set' -Status 'Failed' -Message $message
        }
    }
}
catch {
    $fatalError = $_
    $details = Get-GraphErrorInfo -ErrorRecord $_
    Write-Log -Level ERROR -Message "The run could not continue: $($details.Message)"
    Add-RunResult -Source 'Run' -Action 'Fatal' -Status 'Failed' -Message $details.Message
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
