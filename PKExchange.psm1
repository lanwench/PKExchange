# PKExchange

$ModuleName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$OriginalVerbose = $VerbosePreference

If (Test-Path $PSScriptRoot\Scripts -ErrorAction SilentlyContinue) {

    Try {
        [object[]]$ScriptFiles = @( Get-ChildItem -Path $PSScriptRoot\Scripts -Filter "*.ps1" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -imatch '^function_' } )

        $MissingModules = @()
        @('ExchangeOnlineManagement', 'Microsoft.Graph') | Foreach-Object  {
            If (-not (Get-Module -Name $_ -ListAvailable -ErrorAction SilentlyContinue)) {
                $MissingModules += $_
            }
        }
        If ($MissingModules) {
            Write-Warning "[$ModuleName] Missing one or more required modules: $($MissingModules -join ', ')"
            Return
        }
    
        Foreach($import in ($ScriptFiles | Sort-Object FullName)) {
            $VerbosePreference = $OriginalVerbose
            Write-Verbose "[$ModuleName] $($Import.FullName)"
            Try {
                $VerbosePreference = 'SilentlyContinue'
                . $import.fullname
            }
            Catch {
                $VerbosePreference = $OriginalVerbose
                Write-Error -Message "Failed to import function $($import.fullname): $_"
            }
        }
    }
    Catch {}
    Finally {
        $VerbosePreference = $OriginalVerbose
    }
}
