# Module PKExchange

## About
|||
|---|---|
|**Name** |PKExchange|
|**Author** |Paula Kingsley|
|**Type** |Script|
|**Version** |1.1|
|**Description**|Exchange Online utilities and helpers for quarantine, message, and mailbox management|
|**Date**|README.md file generated on Monday, November 9, 2026 12:15:09 PM|

This module contains 5 PowerShell functions or commands

All functions should have reasonably detailed comment-based help, accessible via Get-Help ... e.g., 
  * `Get-Help Do-Something`
  * `Get-Help Do-Something -Examples`
  * `Get-Help Do-Something -ShowWindow`

## Prerequisites

Computers must:

  * be running PowerShell 5.1 or later
  * have the ExchangeOnline & Microsoft.Graph modules

## Installation

Clone/copy entire module directory into a valid PSModules folder on your computer and run `Import-Module PKExchange`

## Notes

_All code should be presumed to be written by Paula Kingsley unless otherwise specified (see the context help within each function for more information, including credits)._

__Per-function change details are generally found within each function's Notes, but a CHANGELOG.md file contains a summary of changes across all module functions.__

## Commands

|**Command**|**Version**|**Synopsis**|
|---|---|---|
|<span style="white-space: nowrap">**Approve-PKQuarantineMessage**</span>|01.01|Releases one or more quarantined Exchange messages to original or other recipients and optionally add senders to safe list; forces confirmation and includes message detail for safety|
|<span style="white-space: nowrap">**Get-PKQuarantineMessage**</span>|01.00|Searches Exchange Online quarantine with flexible filtering options, returning matching message objects which can be passed to Release-QuarantineMessage or Unblock-PKQuarantinedMessage|
|<span style="white-space: nowrap">**Invoke-PKMessageSearch**</span>|02.00|Submits an Exchange Online historical message trace search (results are emailed when complete)|
|<span style="white-space: nowrap">**Search-PKMessageLog**</span>|02.00|Uses Get-MessageTraceV2 to searches Exchange Online logs for messages sent or received within the past 10 days, filtering on sender, recipient, ID, subject (wildcards allowed) and date|
|<span style="white-space: nowrap">**Trace-PKMessageDelivery**</span>|01.00|Fetches and parses detailed delivery events for an Exchange Online message trace entry|
