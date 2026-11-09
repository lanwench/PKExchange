#requires -Version 5
function Get-PKQuarantineMessage {
    <#
    .SYNOPSIS
        Searches Exchange Online quarantine with flexible filtering options, returning matching message objects which can be passed to Release-QuarantineMessage or Unblock-PKQuarantinedMessage

    .DESCRIPTION
        Wraps Get-QuarantineMessage to search the Exchange Online quarantine, returning results immediately (no async job)
        Results can be passed to Release-QuarantineMessage or Unblock-PKQuarantinedMessage via pipeline
        Created to simplify, and add functionality to, the native cmdlet, such as case-insensitive partial matching; does not include all parameters available in Get-QuarantineMessage, but covers the most commonly used ones for search and release scenarios
        
        - Checks for active connection to Exchange Online 
        - At least one of StartDate,MessageId, SenderAddress, RecipientAddress, or Subject must be specified
        - Automatically pages through all results up to the specified maximum
        - Exact subject matching filters server-side (fast, case-sensitive); Partial filters client-side after fetch (case-insensitive, slower without other filters)

    .NOTES
        Name    : Get-PKQuarantineMessage.ps1 
        Created : 2026-05-28
        Author  : Paula Kingsley
        Version : 01.00.0000
        History :

            ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

            v01.00.0000 - 2026-05-28 - Created script, with help from Claudette

    .PARAMETER StartDate
        Start of the search window (defaults to 3 days ago)

    .PARAMETER EndDate
        End of the search window (defaults to now; must be after StartDate)

    .PARAMETER MessageId
        Exact message ID to search for (e.g., <BN6PR07MB3158A1B2C3D4@BN6PR07MB3158.namprd07.prod.outlook.com>)

    .PARAMETER SenderAddress
        Sender address to filter on (e.g., *@domain.com, jane*@domain.com); API supports wildcards only before @ symbol

    .PARAMETER RecipientAddress
        Recipient address to filter on (e.g., *@domain.com, jane*@domain.com); API supports wildcards only before @ symbol

    .PARAMETER Subject
        Subject text to filter on (case-sensitive unless -SubjectMatchType is Partial)

    .PARAMETER SubjectMatchType
        Required with -Subject; Exact filters on server-side (fast, case-sensitive), Partial matches locally after fetch (slower, case-insensitive)
        
    .PARAMETER ReleaseStatus
        Filter by release status (NotReleased, Released, Requested, PreparingToRelease, Error, NotReleasable); when omitted, all statuses are returned

    .PARAMETER QuarantineType
        Filter by quarantine type (Spam, TransportRule, Bulk', Phish', HighConfPhish, Malware); when omitted, all types are returned

    .PARAMETER Direction
        Filter by message direction (Inbound or Outbound); defaults to Inbound

    .PARAMETER PageSize
        Number of results per page when paging through results. Defaults to 100. Max 1000.

    .PARAMETER Properties
        Full returns all object properties, Compact returns a curated subset with key fields (including Identity for pipeline use); defaults to Compact

    .EXAMPLE
        PS Get-PKQuarantineMessage -SenderAddress '*@customer.com'

            Find all quarantined messages from customer.com received in the last 7 days.

    .EXAMPLE
        PS Get-PKQuarantineMessage -StartDate (Get-Date).AddDays(-30) -ReleaseStatus NotReleased
        
            Find all unreleased quarantined messages received in the last 30 days

    .EXAMPLE
        PS Get-PKQuarantineMessage -RecipientAddress 'jsmith@domain.com' -QuarantineType Phish
        
            Find all quarantined phishing messages sent to jsmith@domain.com

    .EXAMPLE
        PS Get-PKQuarantineMessage -Subject '*invoice*' -SubjectMatchType Partial -SenderAddress 'a*@customer.com' -QuarantineType Spam
        
            Find all messages from 'a*@customer.com' containing the word 'invoice' in the subjectn that have been quarantined as Spam

    .EXAMPLE
        PS Get-PKQuarantineMessage -MessageId '<BN6PR07MB3158A1B2C3D4@BN6PR07MB3158.namprd07.prod.outlook.com>'
        
            Find a specific quarantined message by ID
    #>

    [CmdletBinding(DefaultParameterSetName = 'Default', SupportsShouldProcess, ConfirmImpact = 'Low')]
    [OutputType([PSCustomObject])]

    Param(
        [Parameter(
            Position = 0,
            HelpMessage = 'Start of search window (defaults to 3 days ago, must be prior to -EndDate)'
        )]
        [ValidateScript({If ($_ -gt (Get-Date)) { Throw 'No time travel; -StartDate cannot be in the future!' }; $true})]
        [ValidateNotNullOrEmpty()]
        [datetime] $StartDate = (Get-Date).AddDays(-3),

        [Parameter(
            Position = 1,
            HelpMessage = 'End of search window (defaults to now, must be after -StartDate)'
        )]
        [ValidateScript({If ($_ -lt $StartDate) { Throw 'No time travel; -EndDate cannot be before -StartDate!' }; $true})]
        [datetime] $EndDate = (Get-Date),

        [Parameter(
            Mandatory,
            ParameterSetName = 'ByMessageId',
            Position = 2,
            HelpMessage = 'Exact message ID (e.g., <BN6PR07MB3158A1B2C3D4@BN6PR07MB3158.namprd07.prod.outlook.com>)'
        )]
        [ValidatePattern('^<[^@]+@[^>]+>$')]
        [string] $MessageId,

        [Parameter(
            ParameterSetName = 'Default',       
            Position = 3, 
            HelpMessage = 'Filter for sender address; wildcards permitted only before the @ symbol (e.g., *@domain.com or jane*@domain.com, not @*domain*)'
        )]
        [Parameter(
            ParameterSetName = 'SubjectSearch',  
            Position = 3,
            HelpMessage = 'Filter for sender address; wildcards permitted only before the @ symbol (e.g., *@domain.com or jane*@domain.com, not @*domain*)'
        )]
        [ValidateScript({
            If ($_ -notmatch '^[\w\.\-\*\+]+@[\w\.\-]+$') {
                Throw "Invalid sender address syntax '$_'! Wildcards are permitted only before the @ symbol (e.g., *@domain.com or jane*@domain.com, not @*domain*)"
            }
            $true
        })]
        [string] $SenderAddress,

        [Parameter(
            ParameterSetName = 'Default',
            Position = 4, 
            HelpMessage = 'Filter for recipient address; wildcards permitted only before the @ symbol (e.g., *@domain.com or jane*@domain.com, not @*domain*)'
        )]
        [Parameter(
            ParameterSetName = 'SubjectSearch',
            Position = 4,
            HelpMessage = 'Filter for recipient address; wildcards permitted only before the @ symbol (e.g., *@domain.com or jane*@domain.com, not @*domain*)'
        )]
        [ValidateScript({
            If ($_ -notmatch '^[\w\.\-\*\+]+@[\w\.\-]+$') {Throw "Invalid recipient address syntax '$_'! Wildcards are permitted only before the @ symbol (e.g., *@domain.com or jane*@domain.com, not @*domain*)"}
            $true
        })]
        [string] $RecipientAddress,

        [Parameter(
            ParameterSetName = 'SubjectSearch',
            HelpMessage = 'Filter for subject text; requires -SubjectMatchType'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $Subject,

        [Parameter(
            Mandatory,
            ParameterSetName = 'SubjectSearch',
            HelpMessage = 'Type of match for -Subject; required (Exact = server-side filter, case-sensitive, fast; Partial = client-side, case-insensitive, slower)'
        )]
        [ValidateSet('Exact', 'Partial')]
        [string] $SubjectMatchType,

        [Parameter(ParameterSetName = 'Default',      HelpMessage = 'Filter by release status (NotReleased, Released, Requested, PreparingToRelease, Error, NotReleasable); omit to return all statuses')]
        [Parameter(ParameterSetName = 'SubjectSearch')]
        [ValidateSet('NotReleased', 'Released', 'Requested', 'PreparingToRelease', 'Error', 'NotReleasable')]
        [string] $ReleaseStatus,

        [Parameter(ParameterSetName = 'Default',      HelpMessage = 'Filter by quarantine type (Spam, TransportRule, Bulk, Phish, HighConfPhish, Malware); omit to return all types')]
        [Parameter(ParameterSetName = 'SubjectSearch')]
        [ValidateSet('Spam', 'TransportRule', 'Bulk', 'Phish', 'HighConfPhish', 'Malware')]
        [string] $QuarantineType,

        [Parameter(
            HelpMessage = 'Filter by Inbound or Outbound message direction (defaults to Inbound)'
        )]
        [ValidateSet('Inbound', 'Outbound')]
        [string] $Direction = 'Inbound',

        [Parameter(
            HelpMessage = 'Results per page when paging (defaults to 100, max 1000)'
        )]
        [ValidateRange(1, 1000)]
        [int] $PageSize = 100,

        [Parameter(
            HelpMessage = 'Object properties to return (Full = all object properties, Compact = curated subset); defaults to Compact'
        )]
        [ValidateSet('Full', 'Compact')]
        [string] $Properties = 'Compact'
    )

    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

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

        # Validate parameter requirements & combinations
        If ($EndDate -le $StartDate) {
            Throw "No time travel; -EndDate ($EndDate) must be after -StartDate ($StartDate)!"
        }

        If (-not (Get-Command Get-QuarantineMessage -ErrorAction SilentlyContinue)) {
            Write-Error -Message "No active connection found; please run Connect-ExchangeOnline" -Category InvalidOperation -ErrorAction Stop
        }

        # We need a way to break if the user says no to a potentially expensive search, but we can't call ShouldProcess in the Begin block, so we'll set a flag and check it in Process
        $Script:AbortSearch   = $false
        $HasSenderOrRecipient = $PSBoundParameters.ContainsKey('SenderAddress') -or $PSBoundParameters.ContainsKey('RecipientAddress')
        $HasServerSideFilter  = $HasSenderOrRecipient -or $PSBoundParameters.ContainsKey('ReleaseStatus') -or $PSBoundParameters.ContainsKey('QuarantineType')
        $WindowDays           = ($EndDate - $StartDate).TotalDays

        If ($PSBoundParameters.ContainsKey('Subject') -and $SubjectMatchType -eq 'Partial' -and
            -not $PSBoundParameters.ContainsKey('MessageId') -and (-not $HasSenderOrRecipient)) {
            If (-not $PSCmdlet.ShouldContinue(
                "Partial subject matching fetches all messages in the window first, then filters locally, which can be slow and resource intensive!`nDo you want to proceed?",
                "No server-side filter")) {
                $Script:AbortSearch = $true
                Return
            }
        }
        ElseIf (-not $HasServerSideFilter -and $WindowDays -ge 3) {
            If (-not $PSCmdlet.ShouldContinue(
                "WARNING: No sender, recipient, type, or status filter is set and the search window spans $([Math]::Round($WindowDays)) day(s) — `nthis may return a large number of results and be slow in larger environments! Do you want to proceed?",
                "No server-side filter")) {
                $Script:AbortSearch = $true
                Return
            }
        }
    }
    Process {

        If ($Script:AbortSearch) { Return }

        $Action = "[BEGIN: $ScriptName] Searching $Direction messages between $($StartDate.ToString('dddd, MMMM d, yyyy HH:mm:ss')) and $($EndDate.ToString('dddd, MMMM d, yyyy HH:mm:ss'))"
        Write-Verbose $Action

        $SearchParams = @{
            StartReceivedDate = $StartDate
            EndReceivedDate   = $EndDate
            Direction         = $Direction
            PageSize          = $PageSize
            Verbose           = $False
            ErrorAction       = 'Stop'
        }
        If ($PSBoundParameters.ContainsKey('MessageId'))                                  { $SearchParams['MessageId']        = $MessageId }
        If ($PSBoundParameters.ContainsKey('SenderAddress'))                              { $SearchParams['SenderAddress']    = $SenderAddress }
        If ($PSBoundParameters.ContainsKey('RecipientAddress'))                           { $SearchParams['RecipientAddress'] = $RecipientAddress }
        If ($PSBoundParameters.ContainsKey('Subject') -and $SubjectMatchType -eq 'Exact') { $SearchParams['Subject']          = $Subject }
        If ($PSBoundParameters.ContainsKey('ReleaseStatus'))                              { $SearchParams['ReleaseStatus']    = $ReleaseStatus }
        If ($PSBoundParameters.ContainsKey('QuarantineType'))                             { $SearchParams['Type']             = $QuarantineType }

        If (-not $PSCmdlet.ShouldProcess("Exchange Quarantine", 'Search quarantine messages')) { Return }

        $Page    = 1
        $Results = [System.Collections.Generic.List[object]]::new()

        # Loop through pages of results until we get a page with fewer results than the page size, indicating we've reached the end (or there are no results at all)
        Do {
            Try {
                $Batch = Get-QuarantineMessage @SearchParams -Page $Page 
            }
            Catch {
                Write-Error "Quarantine query failed: $($_.Exception.GetType().Name) — $($_.Exception.Message)"
                Return
            }
            If ($Batch) { $Results.AddRange([object[]]@($Batch)) }
            $Page++
        } While ($Batch.Count -eq $PageSize)

        # Since the API doesn't support partial subject matching, filter results client-side after fetching (less efficient, potentially slow)
        If ($PSBoundParameters.ContainsKey('Subject') -and $SubjectMatchType -eq 'Partial') {
            $Results = [System.Collections.Generic.List[object]]($Results | Where-Object { $_.Subject -like "*$Subject*" })
        }

        If ($Results.Count -eq 0) {
            Write-Warning "No quarantined messages were found matching the specified criteria" 
            Return
        }
        Write-Verbose "Found $($Results.Count) matching quarantined message(s)" 
        
        If ($Properties -eq 'Full') {
            $Results | Sort-Object ReceivedTime
        }
        Else {
            $Results | Select-Object ReceivedTime, SenderAddress, RecipientAddress, Subject, 
                QuarantineTypes, ReleaseStatus, Expires,Identity |
                    Sort-Object ReceivedTime
        }
    }
    End {
        $Stopwatch.Stop()
        Write-Verbose "[END: $ScriptName] Script execution completed in $($Stopwatch.Elapsed.TotalSeconds.ToString("00.00")) seconds"
        
    }
} #end Get-PKQuarantineMessage
