#requires -Version 5
function Invoke-PKMessageSearch {
    <#
    .SYNOPSIS
        Submits an Exchange Online historical message trace search (results are emailed when complete)

    .DESCRIPTION
        Wraps Start-HistoricalSearch to search Exchange Online message logs for up to 90 days.
        Accepts wildcard patterns for sender and recipient addresses.
        At least one of MessageId, SenderAddress, or RecipientAddress must be specified.

        The search runs as a server-side job on Microsoft's infrastructure — no session needs to stay open.

        Results are emailed to NotifyAddress when done, but you can also use -Wait to poll for completion and auto-open the download URL.

    .NOTES
        Name    : Function_Invoke-PKMessageSearch.ps1
        Created : 2026-06-02
        Author  : Paula Kingsley
        Version : 02.00
        History :

            ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

            v01.00 - 2026-06-02 - Created script for Exchange Online historical message search
            v01.01 - 2026-06-19 - Cosmetic updates, changed to major.minor for version
            v02.00 - 2026-06-19 - Renamed from Start-PKHistoricalSearch; added advisory warning for recent date ranges

    .PARAMETER StartDate
        Start of the search window. Defaults to 7 days ago. Cannot be more than 90 days in the past.

    .PARAMETER EndDate
        End of the search window. Defaults to now. Must be after StartDate.

    .PARAMETER MessageId
        Exact message ID to search for. At least one of MessageId, SenderAddress, or RecipientAddress
        must be provided.

    .PARAMETER SenderAddress
        Sender address or wildcard pattern (e.g., *@godaddy.com). At least one of MessageId,
        SenderAddress, or RecipientAddress must be provided.

    .PARAMETER RecipientAddress
        Recipient address or wildcard pattern (e.g., *@chappaqua.com). At least one of MessageId,
        SenderAddress, or RecipientAddress must be provided.

    .PARAMETER NotifyAddress
        Required. Email address to notify when the search job completes.

    .PARAMETER ReportTitle
        Optional custom title for the search job. Defaults to a timestamped name.

    .PARAMETER Wait
        If specified, polls for job completion and opens the download URL when done.
        By default the function submits the job and returns immediately — results arrive by email.

    .PARAMETER PollIntervalSeconds
        How often (in seconds) to check job status when -Wait is used. Defaults to 60.

    .PARAMETER Direction
        Filter by message direction (Inbound, Outbound, or All). Defaults to Inbound. When Inbound and
        -RecipientAddress is specified, validates that the recipient is a valid mailbox or group.
        When Outbound and -SenderAddress is specified, validates that the sender is a valid mailbox or group.

    .EXAMPLE
        Invoke-PKMessageSearch -SenderAddress '*@godaddy.com' -NotifyAddress 'you@domain.com'
        Search for all mail from godaddy.com in the last 7 days.

    .EXAMPLE
        Invoke-PKMessageSearch -SenderAddress '*@godaddy.com' -StartDate (Get-Date).AddDays(-30) -NotifyAddress 'you@domain.com'
        Search for all mail from godaddy.com in the last 30 days.

    .EXAMPLE
        Invoke-PKMessageSearch -RecipientAddress 'jsmith@chappaqua.com' -StartDate (Get-Date).AddDays(-14) -NotifyAddress 'you@domain.com'
        Search for all mail to jsmith in the last 14 days.

    .EXAMPLE
        Invoke-PKMessageSearch -SenderAddress '*@godaddy.com' -RecipientAddress '*@chappaqua.com' -StartDate '2026-04-01' -EndDate '2026-04-30' -NotifyAddress 'you@domain.com'
        Search for all mail from godaddy.com to any chappaqua.com mailbox during April 2026.

    .EXAMPLE
        Invoke-PKMessageSearch -MessageId '<BN6PR07MB3158A1B2C3D4@BN6PR07MB3158.namprd07.prod.outlook.com>' -NotifyAddress 'you@domain.com'
        Look up a specific message by ID.

    .EXAMPLE
        Invoke-PKMessageSearch -SenderAddress '*@godaddy.com' -NotifyAddress 'you@domain.com' -Wait
        Search and poll until complete, then open the download URL.
    #>

    [CmdletBinding(DefaultParameterSetName = 'Default', SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([PSCustomObject])]

    Param(
        [Parameter(
            Position = 0,
            HelpMessage = 'Start of search window, no more than 90 days prior (defaults to 7 days ago)'
        )]
        [ValidateScript({
            If ($_ -lt (Get-Date).AddDays(-90)) { Throw 'StartDate cannot be more than 90 days in the past!' }
            If ($_ -gt (Get-Date))              { Throw 'StartDate cannot be in the future!' }
            $true
        })]
        [datetime] $StartDate = (Get-Date).AddDays(-7),

        [Parameter(
            Position = 1,
            HelpMessage = 'End of search window; must be after StartDate (defaults to now)'
        )]
        [datetime] $EndDate = (Get-Date),

        [Parameter(
            Position = 2,
            HelpMessage = 'Exact message ID to search for, e.g., <BN6PR07MB3158A1B2C3D4@BN6PR07MB3158.namprd07.prod.outlook.com>'
        )]
        [ValidatePattern('^<[^@]+@[^>]+>$')]
        [string] $MessageId,

        [Parameter(
            Position = 3,
            HelpMessage = 'Sender address or wildcard pattern, e.g., *@domain.com (wildcards supported only before @)'
        )]
        [ValidatePattern('^[\w\.\-\*\+]+@[\w\.\-\*]+$')]
        [string] $SenderAddress,

        [Parameter(
            Position = 4,
            HelpMessage = 'Recipient address or wildcard pattern, e.g., *@domain.com (wildcards supported only before @)'
        )]
        [ValidatePattern('^[\w\.\-\*\+]+@[\w\.\-\*]+$')]
        [string] $RecipientAddress,

        [Parameter(
            HelpMessage = 'Email address to notify when the job completes (defaults to currently-connected user)'
        )]
        [ValidatePattern('^[\w\.\-\+]+@[\w\.\-]+\.\w{2,}$')]
        [string] $NotifyAddress,

        [Parameter(
            HelpMessage = 'Optional custom report title. Defaults to a timestamped name.'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $ReportTitle = "MessageSearch_$(Get-Date -Format 'yyyyMMdd-HHmm')",

        [Parameter(
            ParameterSetName = 'Wait',
            HelpMessage = 'Poll for completion and open the download URL when done; can be verrrry slow (if not specified, runs as fire-and-forget)'
        )]
        [switch] $Wait,

        [Parameter(
            ParameterSetName = 'Wait',
            HelpMessage = 'Seconds between job status checks when -Wait is used (defaults to 60)'
        )]
        [ValidateRange(10, 300)]
        [int] $PollIntervalSeconds = 60,

        [Parameter(
            HelpMessage = 'Filter by message direction: Inbound, Outbound, or All (defaults to Inbound)'
        )]
        [ValidateSet('Inbound', 'Outbound', 'All')]
        [string] $Direction = 'Inbound'
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

        # Advisory: if the date range is within the past 10 days, the real-time search is faster
        If ($StartDate -ge (Get-Date).AddDays(-10)) {
            Write-Warning "[ADVISORY] The requested date range is within the past 10 days. For immediate results without an async job, use Search-PKMessageLog instead."
        }

        # Validate at least one filter parameter is specified
        $ParamArr = "MessageID", "SenderAddress", "RecipientAddress"
        $Missing = $ParamArr | Where-Object { -not $PSBoundParameters.ContainsKey($_) }
        If ($Missing.Count -eq $ParamArr.Count) {
            Throw "You must specify at least one of $($ParamArr -join ', ')!"
        }

        # Adjust EndDate if it's in the future
        If ($EndDate -gt (Get-Date)) {
            $EndDate = (Get-Date)
            Write-Verbose "EndDate was in the future; reverted to now ($($EndDate.ToString('yyyy-MM-dd HH:mm:ss')))"
        }

        # Validate date range
        If ($EndDate -le $StartDate) {
            Throw "EndDate ($EndDate) must be after StartDate ($StartDate)."
        }

        # Ensure connection is established
        If (-not (Get-Command Start-HistoricalSearch -ErrorAction SilentlyContinue)) {
            Write-Error -Message "No active connection found; please run Connect-ExchangeOnline" -Category InvalidOperation -ErrorAction Stop
        }

        # Auto-detect NotifyAddress if not specified
        If (-not $PSBoundParameters.ContainsKey('NotifyAddress')) {
            Try {
                $NotifyAddress = (Get-ConnectionInformation -ErrorAction Stop).UserPrincipalName
            }
            Catch {
                Throw "Could not determine currently-connected user's email for notification! Please specify NotifyAddress explicitly. Error: $_"
            }
        }

        # Validate NotifyAddress is a mail-enabled recipient
        Try {
            $Null = Get-EXORecipient -Identity $NotifyAddress -ErrorAction Stop -Verbose:$False
        }
        Catch {
            Throw "NotifyAddress '$NotifyAddress' could not be validated: $($_.Exception.GetType().Name)"
        }

        # Validate direction-specific addresses
        If (($Direction -eq 'Received'  -or $Direction -eq 'All') -and $PSBoundParameters.ContainsKey('RecipientAddress')) {
            Write-Verbose "Validating RecipientAddress '$RecipientAddress' for Inbound search..."
            $Domain = $RecipientAddress -replace '.*@', ''
            If ($RecipientAddress -match '\*') {
                Try {
                    Write-Warning "Wildcards cannot be validated in RecipientAddress; verifying '$Domain' exists in this tenant"
                    $Null = Get-AcceptedDomain -Identity $Domain -ErrorAction Stop
                }
                Catch {
                    Throw "Domain '$Domain' from RecipientAddress '$RecipientAddress' is not an accepted domain in this tenant."
                }
            }
            Else {
                Try {
                    $Null = Get-EXORecipient -Identity $RecipientAddress -ErrorAction Stop -Verbose:$False
                }
                Catch {
                    Throw "RecipientAddress '$RecipientAddress' (Inbound direction) could not be validated: $($_.Exception.GetType().Name)"
                }
            }
        }
        ElseIf (($Direction -eq 'Sent' -or $Direction -eq 'All') -and $PSBoundParameters.ContainsKey('SenderAddress')) {
            Write-Verbose "Validating SenderAddress '$SenderAddress' for Outbound search..."
            $Domain = $SenderAddress -replace '.*@', ''
            If ($SenderAddress -match '\*') {
                Try {
                    Write-Warning "Wildcards cannot be validated in SenderAddress; verifying '$Domain' exists in this tenant"
                    $Null = Get-AcceptedDomain -Identity $Domain -ErrorAction Stop
                }
                Catch {
                    Throw "Domain '$Domain' from SenderAddress '$SenderAddress' is not an accepted domain in this tenant."
                }
            }
            Else {
                Try {
                    $Null = Get-EXORecipient -Identity $SenderAddress -ErrorAction Stop -Verbose:$False
                }
                Catch {
                    Throw "SenderAddress '$SenderAddress' (Outbound direction) could not be validated: $($_.Exception.GetType().Name)"
                }
            }
        }

        $DirectionLookup = @{
            Outbound = "Sent"
            Inbound = "Received"
            All = "All"
        }

        # Display search parameters
        $Summary = [ordered]@{
            'Report Title' = $ReportTitle
            'Start Date'   = $StartDate.ToString('yyyy-MM-dd HH:mm')
            'End Date'     = $EndDate.ToString('yyyy-MM-dd HH:mm')
            'Direction'    = $Direction
            'Message ID'   = If ($PSBoundParameters.ContainsKey('MessageId'))        { $MessageId }        Else { '(not specified)' }
            'Sender'       = If ($PSBoundParameters.ContainsKey('SenderAddress'))    { $SenderAddress }    Else { '(not specified)' }
            'Recipient'    = If ($PSBoundParameters.ContainsKey('RecipientAddress')) { $RecipientAddress } Else { '(not specified)' }
            'Notify'       = $NotifyAddress
        }
        Write-Verbose "Search parameters:`n$([PSCustomObject]$Summary | Format-List | Out-String)"
    }
    Process {
        $Action = "[BEGIN: $ScriptName] Submitting historical message trace search from $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))"
        Write-Verbose $Action

        $SearchParams = @{
            ReportTitle   = $ReportTitle
            StartDate     = $StartDate.ToString('yyyy-MM-dd')
            EndDate       = $EndDate.ToString('yyyy-MM-dd')
            ReportType    = 'MessageTrace'
            NotifyAddress = $NotifyAddress
            Direction     = $DirectionLookup[$Direction]
        }
        If ($PSBoundParameters.ContainsKey('MessageId'))        { $SearchParams['MessageId']        = $MessageId }
        If ($PSBoundParameters.ContainsKey('SenderAddress'))    { $SearchParams['SenderAddress']    = $SenderAddress }
        If ($PSBoundParameters.ContainsKey('RecipientAddress')) { $SearchParams['RecipientAddress'] = $RecipientAddress }

        $Msg = "Submit historical search"
        If ($Wait) {
            $Msg += " and wait/poll for completion (this can take several minutes to hours depending on the date range and filters)"
        }
        If (-not $PSCmdlet.ShouldProcess($ReportTitle, $Msg)) { Return }

        Try {
            $Job = Start-HistoricalSearch @SearchParams -ErrorAction Stop
        }
        Catch {
            Write-Error "Failed to submit search job: $_"
            Return
        }
        Write-Verbose "Submitted job ID $($job.JobId)"

        If ($Wait) {
            Write-Verbose "Polling for job completion (interval: $PollIntervalSeconds seconds)..."
        }
        Else {
            Write-Verbose "Job submitted as background task; results will be emailed to $NotifyAddress when complete"
            Return
        }

        Do {
            Start-Sleep -Seconds $PollIntervalSeconds
            $status = Get-HistoricalSearch -JobId $job.JobId -Verbose:$False -ErrorAction SilentlyContinue
            Write-Verbose "[$((Get-Date).ToString('HH:mm:ss'))] Status: $($status.Status)"
        } While ($status.Status -notin 'Done', 'Failed')

        If ($status.Status -eq 'Failed') {
            Write-Warning "Search job failed; check the Exchange admin center for details (job ID $($job.JobId))"
            Return
        }

        Write-Verbose "Search complete; opening download URL..."
        Write-Verbose "FileUrl: $($status.FileUrl)"
        Start-Process $status.FileUrl

        [PSCustomObject]@{
            JobId     = $job.JobId
            Title     = $ReportTitle
            StartDate = $StartDate
            EndDate   = $EndDate
            Status    = $status.Status
            FileUrl   = $status.FileUrl
        }
    }

    End {
        $Stopwatch.Stop()
        Write-Verbose "[END: $ScriptName] Script execution completed in $($Stopwatch.Elapsed.TotalSeconds.ToString("00.00")) seconds"
    }
} #end Invoke-PKMessageSearch

