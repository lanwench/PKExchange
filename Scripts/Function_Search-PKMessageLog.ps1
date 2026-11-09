#requires -Version 5
Function Search-PKMessageLog {
<#
.SYNOPSIS
    Uses Get-MessageTraceV2 to searches Exchange Online logs for messages sent or received within the past 10 days, filtering on sender, recipient, ID, subject (wildcards allowed) and date

.DESCRIPTION
    Wraps Get-MessageTraceV2 to search Exchange Online message trace logs, returning results immediately (no async job).

    Created to simplify and extend the native cmdlet with flexible search modes and optional detailed delivery tracking.
    
    Does not include all parameters available in Get-MessageTraceV2, but covers the most commonly used ones for search and troubleshooting.

    - Checks for active connection to Exchange Online
    - Searches last 3 days by default (maximum 10 days; StartDate more than 10 days in the past will throw an error)
    - Three search modes: MessageId (exclusive), ByAddress with optional filters (SenderAddress, RecipientAddress, Subject, Status, FromIP), or date range only
    - Subject filters: plain string = exact match; wildcards control filter type (*text = endswith, text* = startswith, *text* = contains)
    - Optional -ExtendedDetail appends a summary of delivery events via Get-MessageTraceDetailV2; use Trace-PKMessageDelivery for full parsed detail
    - Returns all message properties; pipe to Trace-PKMessageDelivery for per-hop delivery analysis

.NOTES
    Name    : Function_Search-PKMessageLog.ps1
    Created : 2026-06-02
    Author  : Paula Kingsley
    Version : 02.00
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00 - 2026-06-02 - Created script for Exchange Online real-time message search
        v02.00 - 2026-06-19 - Renamed, removed -Fulloutput, changed versioning, removed inner function, added -CollectionsToStrings, startdate param validation

.PARAMETER StartDate
    Start of the search window, within the past 10 days (defaults to 3 days ago)

.PARAMETER EndDate
    End of the search window (defaults to now, must be after StartDate and not in the future)

.PARAMETER MessageId
    Exact message ID to search for; mutually exclusive with all other search filters (sender, recipient, subject)

.PARAMETER SenderAddress
    Sender address or wildcard pattern (e.g., *@domain.com); use with or without RecipientAddress, optional in ByAddress search mode.

.PARAMETER RecipientAddress
    Recipient address or wildcard pattern (e.g., *@domain.com); use with or without SenderAddress, optional in ByAddress search mode.    

.PARAMETER Subject
    Subject line to search for; plain string = exact match, or use wildcards to specify filter type (*text = endswith, text* = startswith, *text* = contains)

.PARAMETER Status
    Filter by delivery status (Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending, Quarantined)

.PARAMETER FromIP
    Filter by source IP address (e.g., 192.168.1.1).

.PARAMETER ResultSize
    Maximum number of results to return (defaults to 0, meaning all results)

.PARAMETER Extended
    Include more detailed delivery events for each message via Get-MessageTraceDetailV2 (slows results due to throttling)

.PARAMETER ThrottleDelaySeconds
    Delay between Get-MessageTraceDetailV2 requests (defaults to 3 seconds, increase if hitting throttle limits); ignored unless -Extended

.EXAMPLE
    PS > Search-PKMessageLog -SenderAddress '*@acquisition.com'
    Search for all mail from acquisition.com in the last 3 days

.EXAMPLE
    PS > Search-PKMessageLog -SenderAddress '*@acquisition.com' -StartDate (Get-Date).AddDays(-5) -EndDate (Get-Date)
    Search for all mail from acquisition.com in the last 5 days

.EXAMPLE
    PS Search-PKMessageLog -RecipientAddress 'jbloggs@domain.com' -StartDate (Get-Date).AddDays(-3) -EndDate (Get-Date)
    Search for all mail to Jane Blogs in the last 3 days

.EXAMPLE
    PS > Search-PKMessageLog -SenderAddress '*@acquisition.com' -Subject '*urgent*'
    Search for mail from acquisition.com with "urgent" in the subject line (contains match)

.EXAMPLE
    PS > Search-PKMessageLog -MessageId '<BN6PR07MB3158A1B2C3D4@BN6PR07MB3158.namprd07.prod.outlook.com>'
    Look up a specific message by ID

.EXAMPLE
    PS > Search-PKMessageLog -SenderAddress '*@megacorp.com' -Status Quarantine
    Search for quarantined mail from megacorp.com in the last 3 days

.EXAMPLE
    PS > Search-PKMessageLog -Status Delivered
    Search for all delivered messages in the last 3 days (no address filter)

.EXAMPLE
    PS > Search-PKMessageLog -SenderAddress '*@acquisition.com' -Subject '*urgent*' -Direction All -ExtendedDetail
    Search for mail containing 'urgent' in the subject, to or from any acquisition.com domain email address, returning detailed delivery events

.EXAMPLE
    PS > Search-PKMessageLog -Verbose -RecipientAddress jbloggs@megacorp.com -Status Quarantined | 
        Group-Object SenderAddress | Select-Object @{N="Count";E={$_.Count}},@{N='SenderAddress';E={$_.Name}}, @{N='Subject'; E={$_.Group.Subject}} | 
            Sort-Object Count -Descending
    Search for all quarantined mail sent to a named address, returning the message count per sender & the subject lines.                
#>

[CmdletBinding(DefaultParameterSetName = 'ByAddress', SupportsShouldProcess, ConfirmImpact = 'High')]
[OutputType([System.Object[]])]

Param(
    [Parameter(
        HelpMessage = 'Start of search window (defaults to 3 days ago)'
    )]
    [ValidateScript({
        If ($_ -gt (Get-Date))              { Throw 'StartDate cannot be in the future!' }
        If ($_ -lt (Get-Date).AddDays(-10)) { Throw 'This function can retrieve only messages sent or received within the past 10 days. For older messages, use Invoke-PKMessageSearch.' }
        $true
    })]
    [datetime] $StartDate = (Get-Date).AddDays(-3),

    [Parameter(
        HelpMessage = 'End of search window; must be after StartDate (defaults to now)'
    )]
    [ValidateScript({
        If ($_ -gt (Get-Date)) { Throw 'EndDate cannot be in the future!' }
        $true
    })]
    [datetime] $EndDate = (Get-Date),

    [Parameter(
        ParameterSetName = 'ByMessageId',
        Position = 0,
        Mandatory,
        HelpMessage = 'Exact message ID to search for, e.g., <BN6PR07MB3158A1B2C3D4@BN6PR07MB3158.namprd07.prod.outlook.com>'
    )]
    [ValidatePattern('^<[^@]+@[^>]+>$')]
    [string] $MessageId,

    [Parameter(
        ParameterSetName = 'ByAddress',
        Position = 0,
        HelpMessage = 'Sender address or wildcard pattern, e.g., *@domain.com (wildcards supported only before @)'
    )]
    [ValidatePattern('^[\w\.\-\*\+]+@[\w\.\-\*]+$')]
    [string] $SenderAddress,

    [Parameter(
        ParameterSetName = 'ByAddress',
        Position = 1,
        HelpMessage = 'Recipient address or wildcard pattern, e.g., *@domain.com (wildcards supported only before @)'
    )]
    [ValidatePattern('^[\w\.\-\*\+]+@[\w\.\-\*]+$')]
    [string] $RecipientAddress,

    [Parameter(
        ParameterSetName = 'ByAddress',
        HelpMessage = 'Filter for subject text; use wildcards to specify filter type (*text = endswith, text* = startswith, *text* = contains)'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $Subject,

    [Parameter(ParameterSetName = 'ByAddress', HelpMessage = 'Filter by delivery status; omit to return all statuses')]
    [ValidateSet('Delivered', 'Expanded', 'Failed', 'FilteredAsSpam', 'GettingStatus', 'Pending', 'Quarantined')]
    [string] $Status,

    [Parameter(ParameterSetName = 'ByAddress', HelpMessage = 'Filter by source IP address (e.g., 192.168.1.1)')]
    [string] $FromIP,

    [Parameter(
        HelpMessage = 'Maximum number of results to return (defaults to 0 = all results)'
    )]
    [ValidateScript({
        If ($_ -eq 0 -or $_ -eq 'unlimited') { $true }
        ElseIf ($_ -as [int] -gt 0) { $true }
        Else { Throw "ResultSize must be a positive integer, 0 (all results), or 'unlimited'" }
    })]
    [string] $ResultSize = 0,

    [Parameter(
        HelpMessage = "Fetch a summary of delivery events via Get-MessageTraceDetailV2 (will slow processing; doesn't return as much detail as Trace-PKMessageDelivery)"
    )]
    [switch] $Extended,

    [Parameter(
        HelpMessage = 'Delay in seconds between Get-MessageTraceDetailV2 requests (defaults to 3 seconds; increase if hitting throttle limits)'
    )]
    [ValidateRange(1, 60)]
    [int] $ThrottleDelaySeconds = 3,

    [Parameter(
        HelpMessage = "Display collection properties as strings separated by newline (ignored unless -Extended)"
    )]
    [switch]$CollectionsToStrings

)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Start stopwatch for performance measurement
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Display our parameters
    $CurrentParams = @{} + $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} |
        Where-Object {Test-Path variable:$_}| ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If ($PipelineInput.IsPresent) {$CurrentParams.Path = $Null}
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("ParameterSetName",$PSCmdlet.ParameterSetName)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | Out-String)"
    
    # Warn if Subject contains a mid-string wildcard (will be treated as a literal character)
    If ($PSBoundParameters.ContainsKey('Subject') -and $Subject -match '\*' -and $Subject -notmatch '^\*?[^*]+\*?$') {
        Write-Warning "[PREREQUISITES] The Subject filter contains a wildcard (*) in an unsupported position; it will be treated as a literal character.`nWildcards are valid only at the start and/or end of the string (e.g., '*text', 'text*', '*text*')."
    }

    # Validate date range
    If ($EndDate -le $StartDate) {Throw "[PREREQUISITES] EndDate ($EndDate) must be after StartDate ($StartDate)"}

    # Ensure connection is established
    If (-not (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)) {
        Write-Error -Message "[PREREQUISITES] No active connection found; please run Connect-ExchangeOnline" -Category InvalidOperation -ErrorAction Stop
    }

    # Validate recipient (only for ByAddress parameter set with RecipientAddress)
    If ($PSCmdlet.ParameterSetName -eq 'ByAddress' -and $PSBoundParameters.ContainsKey('RecipientAddress')) {
        Write-Verbose "[PREREQUISITES] Validating RecipientAddress '$RecipientAddress'..."
        $Domain = $RecipientAddress -replace '.*@', ''
        If ($RecipientAddress -match '\*') {
            Try {
                Write-Verbose "[PREREQUISITES] Wildcards cannot be validated in RecipientAddress; verifying '$Domain' exists in this tenant"
                $Null = Get-AcceptedDomain -Identity $Domain -ErrorAction Stop
            }
            Catch {Throw "[PREREQUISITES] Domain '$Domain' from RecipientAddress '$RecipientAddress' is not an accepted domain in this tenant."}
        }
        Else {
            Try {$Null = Get-EXORecipient -Identity $RecipientAddress -ErrorAction Stop -Verbose:$False}
            Catch {Throw "[PREREQUISITES] RecipientAddress '$RecipientAddress' could not be validated: $($_.Exception.GetType().Name)"}
        }
    }

    # Define output properties
    $SelectProps = @('Received', 'Status', 'SenderAddress', 'RecipientAddress', 'Subject', 'FromIP', 'Size', 'MessageId', 'MessageTraceId')

    # Warn if Extended is requested
    If ($Extended) {
        Write-Warning "[PREREQUISITES] Adding extended message trace output will significantly slow down results; a 1-second throttle is applied per message to respect the 100 requests/5-minute limit"
        If (-not $PSCmdlet.ShouldContinue("Continue with extended detail?", "Confirm")) {
            Return
        }
    }

    Write-Verbose "[BEGIN: $ScriptName] Search for Exchange messages matching the filter(s)"

    # Display search parameters
    $Summary = [ordered]@{
        'Start Date'   = $StartDate.ToString('yyyy-MM-dd HH:mm')
        'End Date'     = $EndDate.ToString('yyyy-MM-dd HH:mm')
        'Result Size'  = If ($ResultSize -eq 0) { '(all results)' } Else { $ResultSize }
        'Message ID'   = If ($PSBoundParameters.ContainsKey('MessageId'))        { $MessageId }        Else { '(not specified)' }
        'Sender'       = If ($PSBoundParameters.ContainsKey('SenderAddress'))    { $SenderAddress }    Else { '(not specified)' }
        'Recipient'    = If ($PSBoundParameters.ContainsKey('RecipientAddress')) { $RecipientAddress } Else { '(not specified)' }
        'Subject'      = If ($PSBoundParameters.ContainsKey('Subject'))          { $Subject } Else { '(not specified)' }
        'Status'       = If ($PSBoundParameters.ContainsKey('Status'))           { $Status }           Else { '(all statuses)' }
    }
    Write-Verbose "[PROCESS] Search parameters:`n$([PSCustomObject]$Summary | Format-List | Out-String)"
}

Process {
    
    # Splat!
    $SearchParams = @{
        StartDate  = $StartDate
        EndDate    = $EndDate
    }
    If ($ResultSize -ne 0) { $SearchParams['ResultSize'] = $ResultSize }
    If ($PSBoundParameters.ContainsKey('MessageId'))        { $SearchParams['MessageId']        = $MessageId }
    If ($PSBoundParameters.ContainsKey('SenderAddress'))    { $SearchParams['SenderAddress']    = $SenderAddress }
    If ($PSBoundParameters.ContainsKey('RecipientAddress')) { $SearchParams['RecipientAddress'] = $RecipientAddress }
    If ($PSBoundParameters.ContainsKey('Status'))           { $SearchParams['Status']           = $Status }
    If ($PSBoundParameters.ContainsKey('FromIP'))           { $SearchParams['FromIP']           = $FromIP }

    # Handle subject with wildcard-based filter type selection
    If ($PSBoundParameters.ContainsKey('Subject')) {
        Write-Verbose "[PROCESS] Creating subject filter"
        $SubjectText = $Subject
        $FilterType = 'contains'

        # Parse wildcards to determine the right SubjectFilterType
        If ($SubjectText -match '^\*.*\*$') {
            $FilterType = 'Contains'
            $SubjectText = $SubjectText -replace '^\*|\*$', ''
        }
        ElseIf ($SubjectText -match '^\*') {
            $FilterType = 'EndsWith'
            $SubjectText = $SubjectText -replace '^\*'
        }
        ElseIf ($SubjectText -match '\*$') {
            $FilterType = 'StartsWith'
            $SubjectText = $SubjectText -replace '\*$'
        }
        Else {
            $FilterType = 'ExactMatch'
        }

        If ($SubjectText) {
            $SearchParams['Subject']          = $SubjectText
            $SearchParams['SubjectFilterType'] = $FilterType
        }
    }

    Try {
        Write-Verbose "[PROCESS] Initiating Get-MessageTraceV2 with selected filters..."
        $Results = Get-MessageTraceV2 @SearchParams -ErrorAction Stop -WarningAction SilentlyContinue
    }
    Catch {
        Write-Error "[PROCESS] Failed to query message trace: $_"
        Return
    }

    If ($Results) {
        Write-Verbose "[PROCESS] Retrieved $($Results.Count) matching message(s)"

        If ($Extended) {
            If (-not $PSCmdlet.ShouldProcess("$($Results.Count) message(s)", 'Fetch detailed delivery events via Get-MessageTraceDetailV2')) {
                Return
            }

            # Fetch detailed trace information for each result with throttling
            $Results | ForEach-Object {
                $Result = $_
                $DetailedTrace = $Null
                Try {
                    Write-Verbose "[PROCESS] Initiating Get-MessageTraceDetailv2 to get additional delivery information for MessageTraceId: $($Result.MessageTraceId), RecipientAddress: $($Result.RecipientAddress)"
                    $DetailedTrace = Get-MessageTraceDetailV2 -MessageTraceId $Result.MessageTraceId -RecipientAddress $Result.RecipientAddress -ErrorAction Stop
                    $TraceArr = @()
                    $DetailedTrace | Foreach-Object {
                        $TraceArr += "[$(Get-Date $_.Date -f s)] ($($_.Event)) $($_.Detail))"
                    }
                }
                Catch {
                    Write-Verbose "[PROCESS] Failed to fetch details for MessageTraceId $($Result.MessageTraceId): $_"
                }

                # Add the detail data as a property, apply throttle (keeping to <100 requests/5min), output result with expanded DeliveryDetail property
                If ($CollectionsToStrings) {
                    $Result | Add-Member -MemberType NoteProperty -Name 'TraceDetails' -Value ($TraceArr -join ("`n")) -Force
                }
                Else {
                    $Result | Add-Member -MemberType NoteProperty -Name 'TraceDetails' -Value $TraceArr -Force
                }
                Start-Sleep -Seconds $ThrottleDelaySeconds
                $Result | Select-Object -Property @($SelectProps + "TraceDetails")
            }
        }
        Else {
            $Results | Select-Object -Property $SelectProps
        }
    }
    Else {
        Write-Verbose "[PROCESS] No matching messages found"
    }
}
End {
    $Stopwatch.Stop()
    Write-Verbose "[END: $ScriptName] Script completed in $($Stopwatch.Elapsed.TotalSeconds.ToString("00.00")) seconds"
}
} #end Search-PKMessageLog
