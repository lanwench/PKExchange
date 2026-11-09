@{

# Script module or binary module file associated with this manifest.
RootModule = 'PKExchange.psm1'

# Version number of this module.
ModuleVersion = '01.01'

# ID used to uniquely identify this module
GUID = '6c7b3e1f-9d2a-4c5f-8a1b-7e3d2f5c9a6b'

# Author of this module
Author = 'Paula Kingsley'

# Company or vendor of this module
CompanyName = 'Amalgamated Toothpick, Inc'

# Copyright statement for this module
Copyright = '(c) 2026 Paula Kingsley. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Exchange Online utilities and helpers for quarantine, message, and mailbox management'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.1'

# Modules that must be imported into the global environment prior to importing this module
#RequiredModules   = @('ExchangeOnlineManagement','Microsoft.Graph') # asking for these in the psm1 because otherwise it tries to connect asap

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = '*'

}
