# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

# Authenticate with Azure PowerShell using MSI.
# Remove this if you are not planning on using MSI or Azure PowerShell.
# if ($env:MSI_SECRET) {
#     Disable-AzContextAutosave -Scope Process | Out-Null
#     Connect-AzAccount -Identity
# }

# Uncomment the next line to enable legacy AzureRm alias in Azure PowerShell.
# Enable-AzureRmAlias

# You can also define functions or aliases that can be referenced in any of your PowerShell functions.
# AzureAD PowerShell module needs to run on a 64-bits Azure unction app and be imported 
# https://techcommunity.microsoft.com/t5/apps-on-azure-blog/install-azuread-and-azureadpreview-module-in-powershell-function/ba-p/2644778
# $AzureADModuleLocation = ".\Modules\AzureAD\2.0.2.140\AzureAD.psd1"
# Import-Module $AzureADModuleLocation -UseWindowsPowerShell

$MSTeamsDModuleLocation = ".\Modules\MicrosoftTeams\4.0.0\MicrosoftTeams.psd1"
Import-Module $MSTeamsDModuleLocation