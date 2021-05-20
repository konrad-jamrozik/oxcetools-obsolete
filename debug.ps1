# Set to off to enable referencing $null hashmap keys.
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode?view=powershell-7.1
Set-StrictMode -Off 

$currentDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location) }
Import-Module "$currentDir/oxcetools-lib.psm1" -Force

Read-ItemsDebug