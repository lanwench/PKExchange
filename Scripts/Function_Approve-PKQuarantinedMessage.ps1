#requires -Version 5
Function Approve-PKQuarantineMessage {
<#
.SYNOPSIS
    Releases one or more quarantined Exchange messages to original or other recipients and optionally add senders to safe list; forces confirmation and includes message detail for safety 

.DESCRIPTION
    Wraps Release-QuarantineMessage, created to provide a more user-friendly and secure way to release messages from quarantine 
        (and to add useful help & use an approved verb, *ahem*)
    Provides detailed confirmation prompts to help prevent accidental release of malicious messages

    - Confirmation prompt includes sender, recipient, date, type, and release status
    - High-risk messages (Phish/HighConfPhish/Malware) trigger an additional ShouldContinue warning before release
    - Accepts pipeline input from Get-PKQuarantineMessage (Identity property) or a direct identity string
    - When called with only an identity string, looks up message details automatically for the confirmation label
    - Supports releasing to all original recipients or to specific users
    - Includes -AllowSender (with aliases for guessability) to add the sender to the recipient's safe senders list after release
    - Released and NotReleasable messages are skipped automatically

.NOTES
    Name    : Approve-PKQuarantineMessage.ps1
    Created : 2026-05-28
    Author  : Paula Kingsley
    Version : 01.01
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00  2026-05-28  Created script, with help from Claudette & Clippy
        v01.01  2026-06-21  Auto-detect SMTP MessageId vs quarantine Identity GUID; fan out per-recipient on MessageId lookup

.PARAMETER Identity
    The quarantine message identity GUID (accepts pipeline input from Get-QuarantineMessage or Find-PKQuarantinedMessage)

.PARAMETER ReleaseToAll
    Release to all original recipients (cannot be combined with -User)

.PARAMETER User
    Release to specific recipient(s) only (cannot be combined with -ReleaseToAll)

.PARAMETER AllowSender
    Add the sender to the recipient's safe senders list after release

.EXAMPLE
    PS> Get-PKQuarantineMessage -SenderAddress '*@customer.com' | Approve-PKQuarantineMessage -ReleaseToAll

        Find and release all quarantined messages from customer.com to all original recipients

.EXAMPLE
    PS> Get-PKQuarantineMessage -Subject 'picnic' | Approve-PKQuarantineMessage -User 'jane@mydomain.com'

        Release matching messages to a specific user only.

.EXAMPLE
    PS> Get-PKQuarantineMessage -RecipientAddress 'jane@mydomain.com' -Subject 'invoice' -SubjectMatchType Partial | Approve-PKQuarantineMessage -AllowSender

        Find and release messages sent to a specific address which contain 'invoice' anywhere in the subject line, and add the sender to the user's safe senders list after release

.EXAMPLE
    PS> Approve-PKQuarantineMessage -Identity 'c14401cf-aa9a-4d31-ab62-123456789abc|v1'

        Release a single message by identity GUID; looks up sender/subject/type automatically for the confirmation prompt

#>
[CmdletBinding(
    DefaultParameterSetName = 'ReleaseToAll', 
    SupportsShouldProcess, 
    ConfirmImpact = 'High'
)]
Param(
    [Parameter(
        Mandatory,
        Position = 0,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        HelpMessage = 'Quarantine message identity GUID, or SMTP MessageId (e.g., <abc@domain.com>) — format is auto-detected; accepts pipeline input from Get-QuarantineMessage or Get-PKQuarantineMessage'
    )]
    [string] $Identity,

    [Parameter(
        DontShow, 
        ValueFromPipelineByPropertyName
    )]
    [string] $QuarantineTypes,

    [Parameter(
        DontShow, 
        ValueFromPipelineByPropertyName
    )]
    [string] $ReleaseStatus,

    [Parameter(
        DontShow, 
        ValueFromPipelineByPropertyName
    )]
    [string] $SenderAddress,

    [Parameter(
        DontShow, 
        ValueFromPipelineByPropertyName
    )]
    [string] $RecipientAddress,

    [Parameter(
        DontShow, 
    ValueFromPipelineByPropertyName
    )]
    [string] $Subject,

    [Parameter(
        DontShow, 
        ValueFromPipelineByPropertyName
    )]
    [object] $ReceivedTime,

    [Parameter(
        ParameterSetName = 'ReleaseToAll',
        HelpMessage = 'Release to all original recipients'
    )]
    [switch] $ReleaseToAll,

    [Parameter(
        Mandatory,
        ParameterSetName = 'User',
        HelpMessage = 'Release to specific recipient(s) only'
    )]
    [ValidatePattern('^[\w\.\-\+]+@[\w\.\-]+\.\w{2,}$')]
    [string[]] $User,

    [Parameter(
        HelpMessage = 'Add the sender to the recipient safe senders list after release'
    )]
    [Alias('AddSender','AddSafeSender')]
    [switch] $AllowSender
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01"

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
    If ($PipelineInput.IsPresent) {$CurrentParams.Path = $Null}
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("ParameterSetName",$PSCmdlet.ParameterSetName)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | Out-String)"

    $Script:ReleaseCount = 0
    $Script:FailCount    = 0

    If (-not (Get-Command Get-QuarantineMessage -ErrorAction SilentlyContinue)) {
        Write-Error -Message "No active connection found; please run Connect-ExchangeOnline" -Category InvalidOperation -ErrorAction Stop
    }

    $Action = "[BEGIN: $ScriptName] Release messages from quarantine"
    Write-Verbose $Action
}

Process {

    # Auto-detect input format: SMTP MessageId (<x@y>) vs quarantine Identity GUID
    If ($Identity -match '^<[^@]+@[^>]+>$') {
        Write-Verbose "[PROCESS] SMTP MessageId detected; looking up quarantine entries for '$Identity'..."
        Try {
            $Resolved = @(Get-QuarantineMessage -MessageId $Identity -ErrorAction Stop)
        }
        Catch {
            Write-Error "[PROCESS] Failed to look up quarantine entries for MessageId '$Identity': $_"
            Return
        }
        If (-not $Resolved) {
            Write-Warning "[PROCESS] No quarantined message found with MessageId '$Identity'"
            Return
        }
        Write-Verbose "[PROCESS] Found $($Resolved.Count) quarantine entry/entries for '$Identity'"
    }
    Else {
        # Identity GUID — normalise into the same structure, including any pipeline metadata
        $Resolved = @([PSCustomObject]@{
            Identity         = $Identity
            SenderAddress    = $SenderAddress
            RecipientAddress = $RecipientAddress
            Subject          = $Subject
            QuarantineTypes  = $QuarantineTypes
            ReleaseStatus    = $ReleaseStatus
            ReceivedTime     = $ReceivedTime
        })
    }

    $DangerousTypes = @('Phish', 'HighConfPhish', 'Malware')

    $Resolved | ForEach-Object {
        $Entry            = $_
        $MsgIdentity      = $Entry.Identity
        $MsgSender        = $Entry.SenderAddress
        $MsgRecipient     = $Entry.RecipientAddress
        $MsgSubject       = $Entry.Subject
        $MsgType          = $Entry.QuarantineTypes
        $MsgReleaseStatus = $Entry.ReleaseStatus
        $MsgReceivedTime  = $Entry.ReceivedTime

        # If metadata is missing, look up the message to enrich the confirmation label
        If (-not $MsgSender -and -not $MsgType) {
            Write-Verbose "[PROCESS] Looking up message details for '$MsgIdentity'..."
            Try {
                $Detail           = Get-QuarantineMessage -Identity $MsgIdentity -ErrorAction Stop -Verbose:$False
                $MsgSender        = $Detail.SenderAddress
                $MsgRecipient     = $Detail.RecipientAddress
                $MsgSubject       = $Detail.Subject
                $MsgType          = $Detail.QuarantineTypes
                $MsgReleaseStatus = $Detail.ReleaseStatus
                $MsgReceivedTime  = $Detail.ReceivedTime
            }
            Catch {
                Write-Warning "[PROCESS] Could not retrieve message details for '${MsgIdentity}': $($_.Exception.Message)"
            }
        }

        If ($MsgReleaseStatus -in 'NOTRELEASABLE', 'RELEASED') {
            Write-Warning "[PROCESS] Skipping '$MsgIdentity' — ReleaseStatus is '$MsgReleaseStatus'$(If ($MsgReleaseStatus -eq 'RELEASED') { ' (already delivered)' } Else { ' (policy does not permit release)' })!"
            $Script:FailCount++
            Return
        }

        $Label = "Identity: $MsgIdentity"
        If ($MsgReceivedTime) { $Label += " | Received: $(([datetime]$MsgReceivedTime).ToString('yyyy-MM-dd HH:mm'))" }
        If ($MsgSender)       { $Label += " | From: $MsgSender" }
        If ($MsgRecipient)    { $Label += " | To: $MsgRecipient" }
        If ($MsgSubject)      { $Label += " | Subject: $MsgSubject" }
        $Target = If ($MsgType -in $DangerousTypes) { "$Label [HIGH RISK: $MsgType]" } Else { $Label }

        If (-not $PSCmdlet.ShouldProcess($Target, 'Release quarantined message')) { Return }

        If ($MsgType -in $DangerousTypes) {
            If (-not $PSCmdlet.ShouldContinue(
                "WARNING: This message is classified as $MsgType — releasing it may expose recipients to malicious content! Proceed?",
                "HIGH RISK: $MsgType")) {
                $Script:FailCount++
                Return
            }
        }

        $ReleaseParams = @{ Identity = $MsgIdentity; ErrorAction = 'Stop' }
        If ($PSCmdlet.ParameterSetName -eq 'ReleaseToAll') { $ReleaseParams['ReleaseToAll'] = $true }
        If ($PSBoundParameters.ContainsKey('User'))        { $ReleaseParams['User']         = $User }
        If ($AllowSender)                                  { $ReleaseParams['AllowSender']  = $true }

        Try {
            Release-QuarantineMessage @ReleaseParams
            Write-Verbose "[PROCESS] Released: $MsgIdentity"
            $Script:ReleaseCount++
        }
        Catch {
            Write-Error "[PROCESS] Failed to release '${MsgIdentity}': $($_.Exception.GetType().Name)"
            $Script:FailCount++
        }
    }
}

End {
    $Stopwatch.Stop()
    Write-Verbose "[END: $ScriptName] Script execution completed in $($Stopwatch.Elapsed.TotalSeconds.ToString("00.00")) seconds ($Script:ReleaseCount released, $Script:FailCount failed/skipped)"
    #Write-Verbose "[END: $ScriptName] $Script:ReleaseCount released, $Script:FailCount failed/skipped — completed in $($Stopwatch.Elapsed.TotalSeconds.ToString('00.00')) seconds"
}
} #end Approve-PKQuarantineMessage
