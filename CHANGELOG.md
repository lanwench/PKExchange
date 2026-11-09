# Changelog

All notable changes to the PKExchange module are documented here. Per-function details also live in each function's `.NOTES` help block.

## [1.1] - 2026-11-09

### Added
- `Trace-PKMessageDelivery` — fetches and parses detailed per-hop delivery events for a message trace entry (wraps `Get-MessageTraceDetailV2`)
- `Search-PKMessageLog` — real-time Exchange Online message log search (wraps `Get-MessageTraceV2`), with subject wildcard matching, status/FromIP filters, and optional extended delivery detail

### Changed
- Renamed `Start-PKHistoricalSearch` to `Invoke-PKMessageSearch`; added advisory warning when searching recent date ranges
- `Approve-PKQuarantineMessage`: auto-detect SMTP MessageId vs. quarantine Identity GUID, and fan out per-recipient on MessageId lookup
- Standardized module and function versioning on major.minor (e.g. `01.01`) instead of major.minor.build (`01.00.0000`)

### Removed
- Retired the original real-time message search function (superseded by `Search-PKMessageLog`)
- Dropped the `-FullOutput` switch and an inner helper function from message log search
- Removed stale backup files (`Function_Get-PKMessage.ps1.bak`, `Function_Start-PKHistoricalSearch.bak`, `readme.old`)

## [1.0.0] - 2026-05-28

### Added
- Initial release: `Approve-PKQuarantineMessage`, `Get-PKQuarantineMessage`, message trace search, and historical search functions
