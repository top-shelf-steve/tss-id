<#
.SYNOPSIS
Reports mail-enabled Active Directory groups and their direct membership.

.DESCRIPTION
Queries on-premises Active Directory for security groups that have a mail attribute. Use
-IncludeDistributionGroups to include every AD group with a mail value. The script exports
relational CSV files that can be used for review, analysis, or diagramming:

  * Groups: one row per mail-enabled security group.
  * Members: one row per direct group-to-member relationship (MemberScope All only).
  * GroupNesting: one row per direct group-to-group relationship.

Direct membership is intentional. Recursively flattening membership would hide the group
nesting relationships needed to build an accurate diagram. A security group with a mail
value but a non-universal scope is included and flagged for review instead of silently
excluded.

.PARAMETER MemberScope
All (default) captures every direct member and also produces the group-nesting subset.
GroupsOnly captures only direct members that are groups. None produces only group inventory.

.PARAMETER IncludeDistributionGroups
Includes distribution groups as well as security groups. When omitted, the report contains
only security-enabled groups with a mail value. GroupCategory and GroupScope are exported.

.PARAMETER GroupIdentity
Targets one group by distinguished name, object GUID, SID, sAMAccountName, exact name, or
exact display name. A targeted group must have a mail value and match the selected source
criteria. Use -IncludeDistributionGroups when targeting a distribution group.

.PARAMETER SearchBase
Distinguished name at which to begin the group search. The domain default naming context is
used when this is omitted.

.PARAMETER Server
Domain controller or AD LDS instance to query.

.PARAMETER Credential
Alternate credential used for all Active Directory queries.

.PARAMETER TestMode
Processes only the first TestLimit groups alphabetically and the first TestLimit direct
member distinguished names in each selected group. CSV inventory rows explicitly identify
the report as a test sample and show both available and inspected member counts.

.PARAMETER TestLimit
Maximum groups and direct members per selected group processed by TestMode. Defaults to 15.

.PARAMETER ExportPath
Directory to which timestamped CSV files are written. The directory is created if needed.
When omitted, the script offers to export after collection completes.

.PARAMETER NoExportPrompt
Does not prompt for an export directory when ExportPath is omitted.

.EXAMPLE
.\Get-MailEnabledSecurityGroups.ps1 -ExportPath C:\Reports

Captures every direct member and writes group, member, and group-nesting CSV files.

.EXAMPLE
.\Get-MailEnabledSecurityGroups.ps1 -MemberScope GroupsOnly -ExportPath C:\Reports

Captures only group-to-group relationships, which is the smallest useful diagram dataset.

.EXAMPLE
.\Get-MailEnabledSecurityGroups.ps1 -IncludeDistributionGroups -ExportPath C:\Reports\AllMailGroups

Captures security and distribution groups that have a populated AD mail attribute.

.EXAMPLE
.\Get-MailEnabledSecurityGroups.ps1 -GroupIdentity 'Finance Distribution List' -IncludeDistributionGroups -TestMode -ExportPath C:\Reports\FinanceTest

Targets one distribution group and samples its first 15 direct members.

.EXAMPLE
.\Get-MailEnabledSecurityGroups.ps1 -TestMode -MemberScope GroupsOnly -ExportPath C:\Reports\Test

Produces a small sample using at most 15 groups and 15 direct members from each group.

.EXAMPLE
.\Get-MailEnabledSecurityGroups.ps1 -MemberScope None -SearchBase 'OU=Groups,DC=contoso,DC=com'

Creates an inventory without resolving member objects.

.NOTES
Requires Windows PowerShell 5.1 or PowerShell 7+ and the ActiveDirectory module (RSAT).
By default, "mail-enabled" is determined from on-premises AD as GroupCategory=Security plus
a non-empty mail attribute. -IncludeDistributionGroups removes the security-category
restriction. This is an AD attribute report; it does not validate objects against Exchange.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet('All', 'GroupsOnly', 'None')]
    [string]$MemberScope = 'All',

    [switch]$IncludeDistributionGroups,

    [ValidateNotNullOrEmpty()]
    [string]$GroupIdentity,

    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,

    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [System.Management.Automation.PSCredential]$Credential,

    [switch]$TestMode,

    [ValidateRange(1, 1000)]
    [int]$TestLimit = 15,

    [ValidateNotNullOrEmpty()]
    [string]$ExportPath,

    [switch]$NoExportPrompt
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:RunStarted = Get-Date
$script:RunStamp = $script:RunStarted.ToString('yyyyMMdd-HHmmss')
$script:GroupRecords = New-Object 'System.Collections.Generic.List[object]'
$script:MemberRecords = New-Object 'System.Collections.Generic.List[object]'
$script:NestingRecords = New-Object 'System.Collections.Generic.List[object]'
$script:MemberCache = @{}
$script:ReportGroupDns = @{}
$script:MembershipFailures = 0
$script:ResolvedMemberObjects = 0
$script:MatchingGroupCount = 0
$script:SourceGroupCriteria = if ($IncludeDistributionGroups) { 'AllMailEnabledGroups' } else { 'MailEnabledSecurityGroups' }
$script:SourceGroupLabel = if ($IncludeDistributionGroups) { 'mail-enabled group(s)' } else { 'mail-enabled security group(s)' }
$script:ExportFilePrefix = if ($IncludeDistributionGroups) { 'MailEnabledGroup' } else { 'MailEnabledSecurityGroup' }
$script:GroupProperties = @(
    'description', 'displayName', 'groupCategory', 'groupScope', 'mail', 'mailNickname',
    'managedBy', 'member', 'objectGUID', 'proxyAddresses', 'whenChanged', 'whenCreated'
)

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

function Get-PropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }
    return $property.Value
}

function ConvertTo-ReportDate {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return $null
    }
    try {
        return ([datetime]$Value).ToString('o')
    }
    catch {
        return [string]$Value
    }
}

function ConvertTo-ReportGuid {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return $null
    }
    return ([string]$Value).Trim('{}')
}

function ConvertTo-ReportSid {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return $null
    }
    $sidValue = Get-PropertyValue -InputObject $Value -Name 'Value'
    if (-not [string]::IsNullOrWhiteSpace([string]$sidValue)) {
        return [string]$sidValue
    }
    return [string]$Value
}

function Get-PrimarySmtpAddress {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$ProxyAddresses,
        [AllowNull()][string]$Mail
    )

    foreach ($address in @($ProxyAddresses)) {
        if ([string]$address -clike 'SMTP:*') {
            return ([string]$address).Substring(5)
        }
    }
    return $Mail
}

function Test-SecurityGroupType {
    [CmdletBinding()]
    param([AllowNull()][object]$GroupType)

    if ($null -eq $GroupType) {
        return $false
    }
    try {
        return (([int64]$GroupType -band [int64]2147483648) -ne 0)
    }
    catch {
        return $false
    }
}

function Get-GroupScopeFromType {
    [CmdletBinding()]
    param([AllowNull()][object]$GroupType)

    if ($null -eq $GroupType) {
        return $null
    }
    try {
        $value = [int64]$GroupType
        if (($value -band 8) -ne 0) { return 'Universal' }
        if (($value -band 4) -ne 0) { return 'DomainLocal' }
        if (($value -band 2) -ne 0) { return 'Global' }
    }
    catch {
        return $null
    }
    return 'Unknown'
}

function Get-AdConnectionArguments {
    [CmdletBinding()]
    param()

    $arguments = @{ ErrorAction = 'Stop' }
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $arguments.Server = $Server
    }
    if ($null -ne $Credential) {
        $arguments.Credential = $Credential
    }
    return $arguments
}

function ConvertTo-LdapFilterValue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)

    $builder = New-Object System.Text.StringBuilder
    foreach ($character in $Value.ToCharArray()) {
        switch ([int][char]$character) {
            0  { [void]$builder.Append('\00') }
            40 { [void]$builder.Append('\28') }
            41 { [void]$builder.Append('\29') }
            42 { [void]$builder.Append('\2a') }
            92 { [void]$builder.Append('\5c') }
            default { [void]$builder.Append($character) }
        }
    }
    return $builder.ToString()
}

function Resolve-TargetGroup {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Identity)

    $trimmedIdentity = $Identity.Trim()
    $group = $null

    try {
        $identityArguments = Get-AdConnectionArguments
        $identityArguments.Identity = $trimmedIdentity
        $identityArguments.Properties = $script:GroupProperties
        $group = Get-ADGroup @identityArguments
        Write-Log -Level DEBUG -Message "Resolved '$trimmedIdentity' using an AD group identity."
    }
    catch {
        Write-Log -Level DEBUG -Message "'$trimmedIdentity' was not resolved as a DN, GUID, SID, or sAMAccountName; trying exact name and display name."
    }

    if ($null -eq $group) {
        $escapedIdentity = ConvertTo-LdapFilterValue -Value $trimmedIdentity
        $searchArguments = Get-AdConnectionArguments
        $searchArguments.LDAPFilter = "(&(objectCategory=group)(|(name=$escapedIdentity)(displayName=$escapedIdentity)))"
        $searchArguments.SearchBase = $SearchBase
        $searchArguments.ResultPageSize = 100
        $searchArguments.Properties = $script:GroupProperties
        $matches = @(Get-ADGroup @searchArguments)

        if ($matches.Count -eq 0) {
            throw "No group matching '$trimmedIdentity' was found. Try its sAMAccountName, distinguished name, or object GUID."
        }
        if ($matches.Count -gt 1) {
            $distinguishedNames = ($matches | ForEach-Object { $_.DistinguishedName }) -join '; '
            throw "More than one group has the exact name or display name '$trimmedIdentity'. Re-run with one of these distinguished names: $distinguishedNames"
        }
        $group = $matches[0]
    }

    if ([string]::IsNullOrWhiteSpace([string]$group.mail)) {
        throw "Group '$($group.Name)' was found, but its AD mail attribute is empty, so it is not eligible for this report."
    }
    if (-not $IncludeDistributionGroups -and [string]$group.GroupCategory -ine 'Security') {
        throw "Group '$($group.Name)' is a $($group.GroupCategory) group. Re-run with -IncludeDistributionGroups to report it."
    }

    return $group
}

function Resolve-DirectMember {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DistinguishedName)

    if ($script:MemberCache.ContainsKey($DistinguishedName)) {
        Write-Log -Level DEBUG -Message "Using cached directory object '$DistinguishedName'."
        return $script:MemberCache[$DistinguishedName]
    }

    $result = $null
    try {
        $arguments = Get-AdConnectionArguments
        $arguments.Identity = $DistinguishedName
        $arguments.Properties = @(
            'displayName', 'mail', 'mailNickname', 'objectClass', 'objectGUID', 'objectSid',
            'proxyAddresses', 'sAMAccountName', 'targetAddress', 'userPrincipalName',
            'whenChanged', 'whenCreated', 'groupType', 'userAccountControl',
            'msExchRecipientTypeDetails', 'msExchRecipientDisplayType',
            'msExchRemoteRecipientType'
        )
        $directoryObject = Get-ADObject @arguments
        $script:ResolvedMemberObjects++
        $result = [pscustomobject]@{
            Status = 'Resolved'
            Object = $directoryObject
            Error  = $null
        }
    }
    catch {
        $script:MembershipFailures++
        $result = [pscustomobject]@{
            Status = 'Failed'
            Object = $null
            Error  = $_.Exception.Message
        }
    }

    $script:MemberCache[$DistinguishedName] = $result
    return $result
}

function New-MembershipRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$ParentGroup,
        [Parameter(Mandatory)][string]$MemberDistinguishedName,
        [Parameter(Mandatory)][object]$Resolution
    )

    $member = $Resolution.Object
    $memberClass = [string](Get-PropertyValue -InputObject $member -Name 'ObjectClass')
    $memberMail = [string](Get-PropertyValue -InputObject $member -Name 'mail')
    $memberGroupType = Get-PropertyValue -InputObject $member -Name 'groupType'
    $memberUserAccountControl = Get-PropertyValue -InputObject $member -Name 'userAccountControl'
    $memberEnabled = $null
    $memberAccountStatus = 'NotApplicable'
    if ($memberClass -iin @('user', 'computer')) {
        $memberAccountStatus = 'Unknown'
        if ($null -ne $memberUserAccountControl) {
            try {
                $memberEnabled = (([int64]$memberUserAccountControl -band 2) -eq 0)
                $memberAccountStatus = if ($memberEnabled) { 'Enabled' } else { 'Disabled' }
            }
            catch {
                $memberAccountStatus = 'Unknown'
            }
        }
    }
    $memberIsGroup = $memberClass -ieq 'group'
    $memberIsSecurityGroup = $memberIsGroup -and (Test-SecurityGroupType -GroupType $memberGroupType)
    $memberIsMailEnabledGroup = $memberIsGroup -and -not [string]::IsNullOrWhiteSpace($memberMail)
    $memberIsMailEnabledSecurityGroup = $memberIsSecurityGroup -and -not [string]::IsNullOrWhiteSpace($memberMail)
    $memberIsReportGroup = $script:ReportGroupDns.ContainsKey($MemberDistinguishedName)
    $memberName = [string](Get-PropertyValue -InputObject $member -Name 'Name')
    $memberDisplayName = [string](Get-PropertyValue -InputObject $member -Name 'displayName')
    if ([string]::IsNullOrWhiteSpace($memberDisplayName)) {
        $memberDisplayName = $memberName
    }
    $parentDirectMemberCount = @((Get-PropertyValue -InputObject $ParentGroup -Name 'member')).Count
    $parentInspectedMemberCount = $parentDirectMemberCount
    if ($TestMode) {
        $parentInspectedMemberCount = [math]::Min($parentDirectMemberCount, $TestLimit)
    }

    [pscustomobject][ordered]@{
        ReportMode                           = if ($TestMode) { 'TestSample' } else { 'Full' }
        SourceGroupCriteria                  = $script:SourceGroupCriteria
        ParentDirectMemberCount              = $parentDirectMemberCount
        ParentInspectedMemberCount           = $parentInspectedMemberCount
        ParentMembershipWasTruncated         = $TestMode -and $parentDirectMemberCount -gt $TestLimit
        ParentGroupObjectGuid                = ConvertTo-ReportGuid (Get-PropertyValue -InputObject $ParentGroup -Name 'ObjectGUID')
        ParentGroupName                      = [string](Get-PropertyValue -InputObject $ParentGroup -Name 'Name')
        ParentGroupDisplayName               = [string](Get-PropertyValue -InputObject $ParentGroup -Name 'DisplayName')
        ParentGroupSamAccountName            = [string](Get-PropertyValue -InputObject $ParentGroup -Name 'SamAccountName')
        ParentGroupMail                      = [string](Get-PropertyValue -InputObject $ParentGroup -Name 'mail')
        ParentGroupDistinguishedName         = [string](Get-PropertyValue -InputObject $ParentGroup -Name 'DistinguishedName')
        MemberObjectGuid                     = ConvertTo-ReportGuid (Get-PropertyValue -InputObject $member -Name 'ObjectGUID')
        MemberName                           = $memberName
        MemberDisplayName                    = $memberDisplayName
        MemberSamAccountName                 = [string](Get-PropertyValue -InputObject $member -Name 'sAMAccountName')
        MemberObjectClass                    = $memberClass
        MemberMail                           = $memberMail
        MemberPrimarySmtpAddress             = Get-PrimarySmtpAddress -ProxyAddresses (Get-PropertyValue -InputObject $member -Name 'proxyAddresses') -Mail $memberMail
        MemberUserPrincipalName              = [string](Get-PropertyValue -InputObject $member -Name 'userPrincipalName')
        MemberTargetAddress                  = [string](Get-PropertyValue -InputObject $member -Name 'targetAddress')
        MemberSid                            = ConvertTo-ReportSid (Get-PropertyValue -InputObject $member -Name 'objectSid')
        MemberEnabled                        = $memberEnabled
        MemberAccountStatus                  = $memberAccountStatus
        MemberUserAccountControl             = $memberUserAccountControl
        MemberMsExchRecipientTypeDetails     = Get-PropertyValue -InputObject $member -Name 'msExchRecipientTypeDetails'
        MemberMsExchRecipientDisplayType     = Get-PropertyValue -InputObject $member -Name 'msExchRecipientDisplayType'
        MemberMsExchRemoteRecipientType      = Get-PropertyValue -InputObject $member -Name 'msExchRemoteRecipientType'
        MemberDistinguishedName              = $MemberDistinguishedName
        MemberIsGroup                        = $memberIsGroup
        MemberGroupCategory                  = if ($memberIsGroup) { if ($memberIsSecurityGroup) { 'Security' } else { 'Distribution' } } else { $null }
        MemberGroupScope                     = if ($memberIsGroup) { Get-GroupScopeFromType -GroupType $memberGroupType } else { $null }
        MemberIsMailEnabledGroup             = $memberIsMailEnabledGroup
        MemberIsMailEnabledSecurityGroup     = $memberIsMailEnabledSecurityGroup
        MemberIsInReportGroupSet             = $memberIsReportGroup
        MemberWhenCreated                    = ConvertTo-ReportDate (Get-PropertyValue -InputObject $member -Name 'whenCreated')
        MemberWhenChanged                    = ConvertTo-ReportDate (Get-PropertyValue -InputObject $member -Name 'whenChanged')
        MembershipResolutionStatus           = [string]$Resolution.Status
        MembershipResolutionError            = [string]$Resolution.Error
    }
}

function New-GroupInventoryRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Group,
        [Parameter(Mandatory)][string]$MembershipStatus,
        [AllowNull()][string]$MembershipError,
        [AllowNull()][object]$Counts
    )

    $mail = [string](Get-PropertyValue -InputObject $Group -Name 'mail')
    $scope = [string](Get-PropertyValue -InputObject $Group -Name 'GroupScope')

    [pscustomobject][ordered]@{
        ReportMode                      = if ($TestMode) { 'TestSample' } else { 'Full' }
        SourceGroupCriteria             = $script:SourceGroupCriteria
        GroupObjectGuid                 = ConvertTo-ReportGuid (Get-PropertyValue -InputObject $Group -Name 'ObjectGUID')
        Name                            = [string](Get-PropertyValue -InputObject $Group -Name 'Name')
        DisplayName                     = [string](Get-PropertyValue -InputObject $Group -Name 'DisplayName')
        SamAccountName                  = [string](Get-PropertyValue -InputObject $Group -Name 'SamAccountName')
        Mail                            = $mail
        PrimarySmtpAddress              = Get-PrimarySmtpAddress -ProxyAddresses (Get-PropertyValue -InputObject $Group -Name 'proxyAddresses') -Mail $mail
        MailNickname                    = [string](Get-PropertyValue -InputObject $Group -Name 'mailNickname')
        GroupCategory                   = [string](Get-PropertyValue -InputObject $Group -Name 'GroupCategory')
        GroupScope                      = $scope
        ScopeAssessment                 = if ($scope -eq 'Universal') { 'Expected' } else { 'ReviewNonUniversalMailSecurityGroup' }
        Description                     = [string](Get-PropertyValue -InputObject $Group -Name 'Description')
        ManagedBy                       = [string](Get-PropertyValue -InputObject $Group -Name 'ManagedBy')
        DistinguishedName               = [string](Get-PropertyValue -InputObject $Group -Name 'DistinguishedName')
        DirectMemberCount               = if ($null -ne $Counts) { $Counts.DirectoryTotal } else { @((Get-PropertyValue -InputObject $Group -Name 'member')).Count }
        InspectedDirectMemberCount      = if ($null -ne $Counts) { $Counts.Total } else { $null }
        MembershipWasTruncated          = if ($null -ne $Counts) { $Counts.WasTruncated } else { $null }
        DirectGroupMemberCount          = if ($null -ne $Counts) { $Counts.Group } else { $null }
        DirectUserMemberCount           = if ($null -ne $Counts) { $Counts.User } else { $null }
        DirectComputerMemberCount       = if ($null -ne $Counts) { $Counts.Computer } else { $null }
        DirectContactMemberCount        = if ($null -ne $Counts) { $Counts.Contact } else { $null }
        DirectForeignPrincipalCount     = if ($null -ne $Counts) { $Counts.ForeignSecurityPrincipal } else { $null }
        DirectOtherMemberCount          = if ($null -ne $Counts) { $Counts.Other } else { $null }
        UnresolvedMemberCount           = if ($null -ne $Counts) { $Counts.Unresolved } else { $null }
        MembershipInspectionStatus      = $MembershipStatus
        MembershipInspectionError       = $MembershipError
        WhenCreated                     = ConvertTo-ReportDate (Get-PropertyValue -InputObject $Group -Name 'whenCreated')
        WhenChanged                     = ConvertTo-ReportDate (Get-PropertyValue -InputObject $Group -Name 'whenChanged')
    }
}

function Export-ReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Records,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$Path
    )

    if ($Records.Count -gt 0) {
        $Records | Select-Object -Property $Columns |
            Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    }
    else {
        $header = ($Columns | ForEach-Object { '"' + $_.Replace('"', '""') + '"' }) -join ','
        Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "The export command completed, but '$Path' was not created."
    }
    Write-Log -Level SUCCESS -Message "Exported $($Records.Count) row(s) to '$Path'."
}

function Export-Reports {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Directory)

    $resolvedDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Directory)
    if (Test-Path -LiteralPath $resolvedDirectory -PathType Leaf) {
        throw "ExportPath must be a directory, but '$resolvedDirectory' is a file."
    }
    if (-not (Test-Path -LiteralPath $resolvedDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $resolvedDirectory -Force | Out-Null
        Write-Log -Level INFO -Message "Created export directory '$resolvedDirectory'."
    }

    $groupColumns = @(
        'ReportMode', 'SourceGroupCriteria', 'GroupObjectGuid', 'Name', 'DisplayName', 'SamAccountName', 'Mail',
        'PrimarySmtpAddress', 'MailNickname', 'GroupCategory', 'GroupScope',
        'ScopeAssessment', 'Description', 'ManagedBy', 'DistinguishedName',
        'DirectMemberCount', 'InspectedDirectMemberCount', 'MembershipWasTruncated',
        'DirectGroupMemberCount', 'DirectUserMemberCount',
        'DirectComputerMemberCount', 'DirectContactMemberCount',
        'DirectForeignPrincipalCount', 'DirectOtherMemberCount', 'UnresolvedMemberCount',
        'MembershipInspectionStatus', 'MembershipInspectionError', 'WhenCreated', 'WhenChanged'
    )
    $memberColumns = @(
        'ReportMode', 'SourceGroupCriteria', 'ParentDirectMemberCount', 'ParentInspectedMemberCount',
        'ParentMembershipWasTruncated', 'ParentGroupObjectGuid', 'ParentGroupName', 'ParentGroupDisplayName',
        'ParentGroupSamAccountName', 'ParentGroupMail', 'ParentGroupDistinguishedName',
        'MemberObjectGuid', 'MemberName', 'MemberDisplayName', 'MemberSamAccountName',
        'MemberObjectClass', 'MemberMail', 'MemberPrimarySmtpAddress',
        'MemberUserPrincipalName', 'MemberTargetAddress', 'MemberSid',
        'MemberEnabled', 'MemberAccountStatus', 'MemberUserAccountControl',
        'MemberMsExchRecipientTypeDetails', 'MemberMsExchRecipientDisplayType',
        'MemberMsExchRemoteRecipientType',
        'MemberDistinguishedName', 'MemberIsGroup', 'MemberGroupCategory',
        'MemberGroupScope', 'MemberIsMailEnabledGroup', 'MemberIsMailEnabledSecurityGroup',
        'MemberIsInReportGroupSet', 'MemberWhenCreated', 'MemberWhenChanged',
        'MembershipResolutionStatus', 'MembershipResolutionError'
    )

    $groupPath = Join-Path $resolvedDirectory "$($script:ExportFilePrefix)s-$script:RunStamp.csv"
    Export-ReportCsv -Records $script:GroupRecords.ToArray() -Columns $groupColumns -Path $groupPath

    if ($MemberScope -eq 'All') {
        $memberPath = Join-Path $resolvedDirectory "$($script:ExportFilePrefix)Members-$script:RunStamp.csv"
        Export-ReportCsv -Records $script:MemberRecords.ToArray() -Columns $memberColumns -Path $memberPath
    }
    if ($MemberScope -ne 'None') {
        $nestingPath = Join-Path $resolvedDirectory "$($script:ExportFilePrefix)Nesting-$script:RunStamp.csv"
        Export-ReportCsv -Records $script:NestingRecords.ToArray() -Columns $memberColumns -Path $nestingPath
    }
}

function Show-RunSummary {
    [CmdletBinding()]
    param()

    $duration = (Get-Date) - $script:RunStarted
    $mailEnabledNestedGroups = @(
        $script:NestingRecords.ToArray() |
            Where-Object { $_.MemberIsMailEnabledGroup }
    ).Count
    $mailEnabledNestedSecurityGroups = @(
        $script:NestingRecords.ToArray() |
            Where-Object { $_.MemberIsMailEnabledSecurityGroup }
    ).Count
    $groupsWithErrors = @(
        $script:GroupRecords.ToArray() |
            Where-Object { $_.MembershipInspectionStatus -like '*WithErrors' }
    ).Count

    Write-Host ''
    Write-Host 'Run summary' -ForegroundColor Cyan
    Write-Host ('  Report mode:                        {0}' -f $(if ($TestMode) { "Test sample (limit $TestLimit)" } else { 'Full' })) -ForegroundColor $(if ($TestMode) { 'Yellow' } else { 'Gray' })
    if ($TestMode) {
        Write-Host ('  Matching groups available:          {0}' -f $script:MatchingGroupCount)
    }
    Write-Host ('  Source groups captured:             {0}' -f $script:GroupRecords.Count) -ForegroundColor Green
    Write-Host ('  Direct membership rows captured:   {0}' -f $script:MemberRecords.Count)
    Write-Host ('  Direct group-nesting relationships:{0,4}' -f $script:NestingRecords.Count) -ForegroundColor Cyan
    Write-Host ('  Nested mail-enabled groups:         {0}' -f $mailEnabledNestedGroups)
    Write-Host ('  Nested mail security groups:        {0}' -f $mailEnabledNestedSecurityGroups)
    Write-Host ('  Unique member objects resolved:     {0}' -f $script:ResolvedMemberObjects)
    Write-Host ('  Member resolution failures:         {0}' -f $script:MembershipFailures) -ForegroundColor $(if ($script:MembershipFailures -gt 0) { 'Yellow' } else { 'Gray' })
    Write-Host ('  Groups completed with errors:       {0}' -f $groupsWithErrors) -ForegroundColor $(if ($groupsWithErrors -gt 0) { 'Yellow' } else { 'Gray' })
    Write-Host ('  Duration:                            {0:hh\:mm\:ss}' -f $duration)
}

$fatalError = $null
try {
    Write-Host ''
    Write-Host 'Active Directory - Mail-Enabled Group Report' -ForegroundColor Cyan
    Write-Host "Source criteria: $script:SourceGroupCriteria" -ForegroundColor Yellow
    Write-Host "Member scope: $MemberScope (direct membership; not recursively flattened)" -ForegroundColor Yellow
    if ($TestMode) {
        Write-Host "TEST MODE: processing at most $TestLimit groups and $TestLimit direct members per group." -ForegroundColor Magenta
    }
    Write-Host ''

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw 'The ActiveDirectory PowerShell module is required. Install the appropriate RSAT Active Directory tools and run the script again.'
    }
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log -Level SUCCESS -Message 'Loaded the ActiveDirectory PowerShell module.'

    $connectionArguments = Get-AdConnectionArguments
    if ([string]::IsNullOrWhiteSpace($SearchBase)) {
        Write-Log -Level STEP -Message 'Discovering the domain default naming context.'
        $rootDse = Get-ADRootDSE @connectionArguments
        $SearchBase = [string]$rootDse.defaultNamingContext
    }
    Write-Log -Level INFO -Message "Search base: $SearchBase"
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        Write-Log -Level INFO -Message "Directory server: $Server"
    }

    if (-not [string]::IsNullOrWhiteSpace($GroupIdentity)) {
        Write-Log -Level STEP -Message "Resolving targeted group '$GroupIdentity'."
        $targetGroup = Resolve-TargetGroup -Identity $GroupIdentity
        $groups = @($targetGroup)
        $script:MatchingGroupCount = 1
        Write-Log -Level SUCCESS -Message "Targeted '$($targetGroup.Name)' <$($targetGroup.mail)> [$($targetGroup.GroupCategory), $($targetGroup.GroupScope)]."
    }
    else {
        Write-Log -Level STEP -Message "Querying $script:SourceGroupLabel with a non-empty mail attribute."
        $groupArguments = Get-AdConnectionArguments
        if ($IncludeDistributionGroups) {
            $groupArguments.LDAPFilter = '(&(objectCategory=group)(mail=*))'
        }
        else {
            $groupArguments.LDAPFilter = '(&(objectCategory=group)(mail=*)(groupType:1.2.840.113556.1.4.803:=2147483648))'
        }
        $groupArguments.SearchBase = $SearchBase
        $groupArguments.ResultPageSize = 500
        $groupArguments.Properties = $script:GroupProperties
        $groups = @(Get-ADGroup @groupArguments | Sort-Object -Property Name)
        $script:MatchingGroupCount = $groups.Count
        Write-Log -Level SUCCESS -Message "Found $script:MatchingGroupCount $script:SourceGroupLabel."
    }

    if ($TestMode -and -not [string]::IsNullOrWhiteSpace($GroupIdentity)) {
        Write-Log -Level WARN -Message "Test mode will limit direct-member inspection for the targeted group to $TestLimit member(s)."
    }
    elseif ($TestMode -and $groups.Count -gt $TestLimit) {
        $groups = @($groups | Select-Object -First $TestLimit)
        Write-Log -Level WARN -Message "Test mode selected the first $($groups.Count) group(s) alphabetically from $script:MatchingGroupCount matching groups."
    }
    elseif ($TestMode) {
        Write-Log -Level WARN -Message "Test mode is active; all $($groups.Count) matching group(s) fit within the limit of $TestLimit."
    }

    foreach ($group in $groups) {
        $script:ReportGroupDns[[string]$group.DistinguishedName] = $true
    }

    $groupIndex = 0
    foreach ($group in $groups) {
        $groupIndex++
        $allMemberDns = @($group.member | Sort-Object)
        $memberDns = $allMemberDns
        $membershipWasTruncated = $false
        if ($TestMode -and $allMemberDns.Count -gt $TestLimit) {
            $memberDns = @($allMemberDns | Select-Object -First $TestLimit)
            $membershipWasTruncated = $true
        }
        Write-Log -Level STEP -Message "[$groupIndex/$($groups.Count)] Inspecting '$($group.Name)' <$($group.mail)>; $($allMemberDns.Count) direct member value(s) available."
        if ($membershipWasTruncated) {
            Write-Log -Level WARN -Message "  Test mode will inspect only $($memberDns.Count) of $($allMemberDns.Count) direct members for this group."
        }

        if ($MemberScope -eq 'None') {
            $script:GroupRecords.Add((New-GroupInventoryRecord -Group $group -MembershipStatus 'NotRequested' -MembershipError $null -Counts $null))
            continue
        }

        $counts = [pscustomobject]@{
            DirectoryTotal           = $allMemberDns.Count
            Total                    = $memberDns.Count
            WasTruncated             = $membershipWasTruncated
            Group                    = 0
            User                     = 0
            Computer                 = 0
            Contact                  = 0
            ForeignSecurityPrincipal = 0
            Other                    = 0
            Unresolved               = 0
        }
        $groupErrors = New-Object 'System.Collections.Generic.List[string]'
        $memberIndex = 0

        foreach ($memberDn in $memberDns) {
            $memberIndex++
            $resolution = Resolve-DirectMember -DistinguishedName ([string]$memberDn)
            $relationship = New-MembershipRecord -ParentGroup $group -MemberDistinguishedName ([string]$memberDn) -Resolution $resolution

            if ($resolution.Status -eq 'Failed') {
                $counts.Unresolved++
                $groupErrors.Add("$memberDn - $($resolution.Error)")
                Write-Log -Level ERROR -Message "  [$memberIndex/$($memberDns.Count)] Could not resolve '$memberDn': $($resolution.Error)"
            }
            else {
                $memberLabel = if ([string]::IsNullOrWhiteSpace($relationship.MemberDisplayName)) { $memberDn } else { $relationship.MemberDisplayName }
                $memberType = if ([string]::IsNullOrWhiteSpace($relationship.MemberObjectClass)) { 'unknown' } else { $relationship.MemberObjectClass }
                if ($relationship.MemberAccountStatus -eq 'Disabled') {
                    $memberType = "$memberType; account disabled"
                }
                Write-Log -Level INFO -Message "  [$memberIndex/$($memberDns.Count)] $($group.Name) -> $memberLabel [$memberType]"

                switch ($relationship.MemberObjectClass.ToLowerInvariant()) {
                    'group'                    { $counts.Group++ }
                    'user'                     { $counts.User++ }
                    'computer'                 { $counts.Computer++ }
                    'contact'                  { $counts.Contact++ }
                    'foreignsecurityprincipal' { $counts.ForeignSecurityPrincipal++ }
                    default                    { $counts.Other++ }
                }
            }

            if ($MemberScope -eq 'All') {
                $script:MemberRecords.Add($relationship)
            }
            if ($relationship.MemberIsGroup) {
                $script:NestingRecords.Add($relationship)
                if ($relationship.MemberIsInReportGroupSet) {
                    Write-Log -Level WARN -Message "    Nested report group detected: '$($relationship.MemberDisplayName)'."
                }
            }
        }

        $inspectionStatus = if ($membershipWasTruncated -and $groupErrors.Count -gt 0) {
            'SampledWithErrors'
        }
        elseif ($membershipWasTruncated) {
            'Sampled'
        }
        elseif ($groupErrors.Count -gt 0) {
            'CompletedWithErrors'
        }
        else {
            'Completed'
        }
        $inspectionError = if ($groupErrors.Count -gt 0) { $groupErrors -join ' | ' } else { $null }
        $script:GroupRecords.Add((New-GroupInventoryRecord -Group $group -MembershipStatus $inspectionStatus -MembershipError $inspectionError -Counts $counts))
        Write-Log -Level SUCCESS -Message "Completed '$($group.Name)': $($counts.Group) group(s), $($counts.User) user(s), $($counts.Computer) computer(s), $($counts.Other + $counts.Contact + $counts.ForeignSecurityPrincipal) other, $($counts.Unresolved) unresolved."
    }
}
catch {
    $fatalError = $_
    Write-Log -Level ERROR -Message "The report could not continue: $($_.Exception.Message)"
}
finally {
    Show-RunSummary

    try {
        if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            Export-Reports -Directory $ExportPath
        }
        elseif (-not $NoExportPrompt -and $null -eq $fatalError) {
            Write-Host ''
            $exportChoice = Read-Host 'Export the report CSV files? [Y/n]'
            if ([string]::IsNullOrWhiteSpace($exportChoice) -or $exportChoice -match '^(?i)y(?:es)?$') {
                $defaultDirectory = Join-Path (Get-Location).Path "$($script:ExportFilePrefix)s-Report-$script:RunStamp"
                $chosenDirectory = Read-Host "Export directory [$defaultDirectory]"
                if ([string]::IsNullOrWhiteSpace($chosenDirectory)) {
                    $chosenDirectory = $defaultDirectory
                }
                Export-Reports -Directory $chosenDirectory
            }
            else {
                Write-Log -Level INFO -Message 'Report data was not exported.'
            }
        }
    }
    catch {
        Write-Log -Level ERROR -Message "Could not export report CSV files: $($_.Exception.Message)"
        if ($null -eq $fatalError) {
            $fatalError = $_
        }
    }
}

if ($null -ne $fatalError) {
    throw $fatalError
}
