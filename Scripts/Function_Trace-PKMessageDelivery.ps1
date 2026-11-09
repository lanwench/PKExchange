#requires -Version 5
Function Trace-PKMessageDelivery {
<#
.SYNOPSIS
    Fetches and parses detailed delivery events for an Exchange Online message trace entry

.DESCRIPTION
    Wraps Get-MessageTraceDetailV2 to return detailed per-hop delivery events for a traced message,
    with automatic parsing of EventData XML into readable property/value pairs where present.

    Accepts MessageTraceId and RecipientAddress directly or via pipeline from Get-PKMessageTrace output.

    - Checks for active connection to Exchange Online
    - Parses EventData XML (MEP property/value pairs) into a structured EventData property per event
    - Applies a throttle delay between requests when processing multiple messages via pipeline
    - Returns one PSCustomObject per delivery event, with MessageTraceId and RecipientAddress included for context

.NOTES
    Name    : Function_Trace-PKMessageDelivery.ps1
    Created : 2026-06-19
    Author  : Paula Kingsley
    Version : 01.00
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00 - 2026-06-19 - Created script

.PARAMETER MessageTraceId
    MessageTraceId GUID from a message trace result (from Get-PKMessageTrace or Get-MessageTraceV2 output)

.PARAMETER RecipientAddress
    Recipient email address for the traced message; required by Get-MessageTraceDetailV2

.PARAMETER ThrottleDelaySeconds
    Delay in seconds between requests when processing multiple messages via pipeline (defaults to 3; increase if hitting throttle limits)

.EXAMPLE
    PS > Trace-PKMessageDelivery -MessageTraceId '12345678-1234-1234-1234-123456789abc' -RecipientAddress 'jbloggs@domain.com'
    Fetch and parse delivery events for a specific message

.EXAMPLE
    PS > Get-PKMessageTrace -SenderAddress '*@acquisition.com' -Status Failed | Trace-PKMessageDelivery
    Fetch delivery detail for all failed messages from acquisition.com in the last 3 days

.EXAMPLE
    PS > Get-PKMessageTrace -RecipientAddress 'jbloggs@megacorp.com' -Status Quarantined | Trace-PKMessageDelivery -ThrottleDelaySeconds 5
    Fetch delivery detail for quarantined messages with an increased throttle delay
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
[OutputType([PSCustomObject[]])]

Param(
    [Parameter(
        Mandatory,
        Position = 0,
        ValueFromPipelineByPropertyName,
        HelpMessage = 'MessageTraceId GUID from a message trace result (from Get-PKMessageTrace or Get-MessageTraceV2 output)'
    )]
    [guid] $MessageTraceId,

    [Parameter(
        Mandatory,
        Position = 1,
        ValueFromPipelineByPropertyName,
        HelpMessage = 'Recipient email address for the traced message'
    )]
    [ValidatePattern('^[\w\.\-\+]+@[\w\.\-]+$')]
    [string] $RecipientAddress,

    [Parameter(
        HelpMessage = 'Delay in seconds between requests when processing multiple messages via pipeline (defaults to 3; increase if hitting throttle limits)'
    )]
    [ValidateRange(1, 60)]
    [int] $ThrottleDelaySeconds = 3
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Start stopwatch for performance measurement
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Display our parameters
    $CurrentParams = @{} + $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} |
        Where-Object {Test-Path variable:$_} | ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If ($PipelineInput.IsPresent) {$CurrentParams.MessageTraceId = $Null; $CurrentParams.RecipientAddress = $Null}
    $CurrentParams.Add("PipelineInput", $PipelineInput)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | Out-String)"

    # Ensure connection is established
    If (-not (Get-Command Get-MessageTraceDetailV2 -ErrorAction SilentlyContinue)) {
        Write-Error -Message "[PREREQUISITES] No active connection found; please run Connect-ExchangeOnline" -Category InvalidOperation -ErrorAction Stop
    }

    $FirstCall = $true

    Write-Verbose "[BEGIN: $ScriptName] Fetch and parse detailed delivery events for a traced message"
}

Process {

    If (-not $FirstCall) {
        Write-Verbose "[PROCESS] Applying throttle delay: $ThrottleDelaySeconds second(s)"
        Start-Sleep -Seconds $ThrottleDelaySeconds
    }
    $FirstCall = $false

    If (-not $PSCmdlet.ShouldProcess("MessageTraceId '$MessageTraceId', RecipientAddress '$RecipientAddress'", 'Fetch message trace detail')) {
        Return
    }

    Try {
        Write-Verbose "[PROCESS] Fetching detail for MessageTraceId '$MessageTraceId', RecipientAddress '$RecipientAddress'"
        $TraceDetails = Get-MessageTraceDetailV2 -MessageTraceId $MessageTraceId -RecipientAddress $RecipientAddress -ErrorAction Stop
    }
    Catch {
        Write-Error "[PROCESS] Failed to fetch trace detail for MessageTraceId '$MessageTraceId': $_"
        Return
    }

    If (-not $TraceDetails) {
        Write-Verbose "[PROCESS] No detail events returned for MessageTraceId '$MessageTraceId'"
        Return
    }

    Write-Verbose "[PROCESS] Returned $($TraceDetails.Count) event(s) for MessageTraceId '$MessageTraceId'"

    $TraceDetails | ForEach-Object {
        $Row = $_
        $ParsedEventData = $Null

        If ($Row.EventData -and $Row.EventData -match '<root>') {
            Try {
                [xml]$Xml = $Row.EventData
                $ParsedEventData = $Xml.root.MEP | ForEach-Object {
                    [PSCustomObject]@{
                        Property = $_.Name
                        Value    = $_.String ?? $_.Long ?? $_.Blob
                    }
                }
            }
            Catch {
                Write-Verbose "[PROCESS] Could not parse EventData XML for event '$($Row.Event)': $_"
                $ParsedEventData = $Row.EventData
            }
        }

        [PSCustomObject]@{
            MessageTraceId   = $MessageTraceId
            RecipientAddress = $RecipientAddress
            TimeStamp        = $Row.Date
            Event            = $Row.Event
            Action           = $Row.Action
            Detail           = $Row.Detail
            EventData        = $ParsedEventData
        }
    }
}

End {
    $Stopwatch.Stop()
    Write-Verbose "[END: $ScriptName] Script completed in $($Stopwatch.Elapsed.TotalSeconds.ToString("00.00")) seconds"
}
} #end Trace-PKMessageDelivery
