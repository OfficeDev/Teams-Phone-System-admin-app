# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

$MSTeamsDModuleLocation = ".\Modules\MicrosoftTeams\4.0.0\MicrosoftTeams.psd1"
Import-Module $MSTeamsDModuleLocation

# $AzureADModuleLocation = ".\Modules\AzureAD\2.0.2.140\AzureAD.psd1"
# Import-Module $AzureADModuleLocation -UseWindowsPowerShell


# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
