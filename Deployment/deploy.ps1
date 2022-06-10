Param (
    [parameter(mandatory = $false)] $displayName = "Teams Telephony Manager",   # Display name for your application registered in Azure AD 
    [parameter(mandatory = $false)] $rgName = "Teams-Telephony-Manager",        # Name of the resource group for Azure
    [parameter(mandatory = $false)] $resourcePrefix = "Teams",                  # Prefix for the resources deployed on your Azure subscription
    [parameter(mandatory = $false)] $location = 'westeurope',                   # Location (region) where the Azure resource are deployed
    [parameter(mandatory = $true)] $serviceAccountUPN,                          # AzureAD Service Account UPN
    [parameter(mandatory = $true)] $serviceAccountSecret                        # AzureAD Service Account password
)

$base = $PSScriptRoot
Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true"

# Import required PowerShell modules for the deployment
If($PSVersionTable.PSVersion.Major -ne 7) { 
    Write-Error "Please install and use PowerShell v7.2.1 to run this script"
    Write-Error "Follow the instruction to install PowerShell on Windows here"
    Write-Error "https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.2"
    return
}
Import-Module AzureAD -UseWindowsPowerShell        # Required to register the app in Azure AD using PowerShell 7.x
Import-Module Az.Accounts, Az.Resources, Az.KeyVault   # Required to deploy the Azure resource

# Connect to AzureAD and Azure using modern authentication
write-host -ForegroundColor blue "Azure sign-in request - Please check the sign-in window opened in your web browser"
Connect-AzAccount

# Auto-connect to AzureAD using Azure connection context
$context = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile.DefaultContext
$aadToken = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id.ToString(), $null, [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never, $null, "https://graph.windows.net").AccessToken
Connect-AzureAD -AadAccessToken $aadToken -AccountId $context.Account.Id -TenantId $context.tenant.id

write-host -ForegroundColor blue "Checking if app '$displayName' is already registered"
$AAdapp = Get-AzureADApplication -Filter "DisplayName eq '$displayName'"
If ($AAdapp.Count -gt 1) {
    Write-Error "Multiple Azure AD app registered under the name '$displayName' - Please use another name and retry"
    return
}
If([string]::IsNullOrEmpty($AAdapp)){
    write-host -ForegroundColor blue "Register a new app in Azure AD using Azure Function app name"
    $AADapp = New-AzureADApplication -DisplayName $displayName -AvailableToOtherTenants $false
    $AppIdURI = "api://azfunc-" + $AADapp.AppId
    # Expose an API and create an Application ID URI
    Try {
        Set-AzureADApplication -ObjectId $AADapp.ObjectId -IdentifierUris $AppIdURI
    }    
    Catch {
        Write-Error "Azure AD application registration error - Please check your permissions in Azure AD and review detailed error description below"
        $_.Exception.Message
        return
    }
    # Create a new app secret with a default validaty period of 1 year - Get the generated secret
    $secret   = (New-AzureADApplicationPasswordCredential -ObjectId $AADapp.ObjectId -EndDate (Get-Date).date.AddYears(1)).Value
    # Get the AppID from the newly registered App
    $clientID = $AADapp.AppId
    # Get the tenantID from current AzureAD PowerShell session
    $tenantID = (Get-AzureADTenantDetail).ObjectId
    write-host -ForegroundColor blue "New app '$displayName' registered into AzureAD"
}
Else {
    write-host -ForegroundColor blue "Generating a new secret for app '$displayName'"
    $secret   = (New-AzureADApplicationPasswordCredential -ObjectId $AADapp.ObjectId -EndDate (Get-Date).date.AddYears(1)).Value
    # Get the AppID from the newly registered App
    $clientID = $AADapp.AppId
    # Get the tenantID from current AzureAD PowerShell session
    $tenantID = (Get-AzureADTenantDetail).ObjectId
}

write-host -ForegroundColor blue "Deploy resource to Azure subscription"
Try {
    New-AzResourceGroup -Name $rgName -Location $location -Force
}    
Catch {
    Write-Error "Azure Ressource Group creation failed - Please verify your permissions on the subscription and review detailed error description below"
    $_.Exception.Message
    return
}
write-host -ForegroundColor blue "Resource Group $rgName created in location $location - Now initiating Azure resource deployments..."
$deploymentName = 'deploy-' + (Get-Date -Format "yyyyMMdd-hhmm")
$parameters = @{
    resourcePrefix          = $resourcePrefix
    serviceAccountUPN       = $serviceAccountUPN
    serviceAccountSecret    = $serviceAccountSecret
    clientID                = $clientID
    appSecret               = $secret
}

$outputs = New-AzResourceGroupDeployment -ResourceGroupName $rgName -TemplateFile $base\ZipDeploy\azuredeploy.json -TemplateParameterObject $parameters -Name $deploymentName -ErrorAction SilentlyContinue
If ($outputs.provisioningState -ne 'Succeeded') {
    Write-Error "ARM deployment failed with error"
    Write-Error "Please retry deployment"
    $outputs
    return
}
write-host -ForegroundColor blue "ARM template deployed successfully"

# Getting UPN from connected user
$CurrentUserId = Get-AzContext | ForEach-Object account | ForEach-Object Id

if($CurrentUserId -ne $serviceAccountUPN)
{
    # Assign current user with the permissions to list and read Azure KeyVault secrets (to enable the connection with the Power Automate flow)
    Write-Host -ForegroundColor blue "Assigning 'Secrets List & Get' policy on Azure KeyVault for user $CurrentUserId"
    Try {
        Set-AzKeyVaultAccessPolicy -VaultName $outputs.Outputs.azKeyVaultName.Value -ResourceGroupName $rgName -UserPrincipalName $CurrentUserId -PermissionsToSecrets list,get
    }
    Catch {
        Write-Error "Error - Couldn't assign user permissions to get,list the KeyVault secrets - Please review detailed error message below"
        $_.Exception.Message
    }

    # Assign service account with the permissions to list and read Azure KeyVault secrets (to enable the connection with the Power Automate flow)
    Write-Host -ForegroundColor blue "Assigning 'Secrets List & Get' policy on Azure KeyVault for user $serviceAccountUPN"
    Try {
        Set-AzKeyVaultAccessPolicy -VaultName $outputs.Outputs.azKeyVaultName.Value -ResourceGroupName $rgName -UserPrincipalName $CurrentUserId -PermissionsToSecrets list,get
    }
    Catch {
        Write-Error "Error - Couldn't assign user permissions to get,list the KeyVault secrets - Please review detailed error message below"
        $_.Exception.Message
    }    
}
else
{
    # Assign service account with the permissions to list and read Azure KeyVault secrets (to enable the connection with the Power Automate flow)
    Write-Host -ForegroundColor blue "Assigning 'Secrets List & Get' policy on Azure KeyVault for user $serviceAccountUPN"
    Try {
        Set-AzKeyVaultAccessPolicy -VaultName $outputs.Outputs.azKeyVaultName.Value -ResourceGroupName $rgName -UserPrincipalName $CurrentUserId -PermissionsToSecrets list,get
    }
    Catch {
        Write-Error "Error - Couldn't assign user permissions to get,list the KeyVault secrets - Please review detailed error message below"
        $_.Exception.Message
    }
}

write-host -ForegroundColor blue "Getting the Azure Function App key for warm-up test"
## lookup the resource id for your Azure Function App ##
$azFuncResourceId = (Get-AzResource -ResourceGroupName $rgName -ResourceName $outputs.Outputs.azFuncAppName.Value -ResourceType "Microsoft.Web/sites").ResourceId

## compose the operation path for listing keys ##
$path = "$azFuncResourceId/host/default/listkeys?api-version=2021-02-01"
$result = Invoke-AzRestMethod -Path $path -Method POST

if($result -and $result.StatusCode -eq 200)
{
   ## Retrieve result from Content body as a JSON object ##
   $contentBody = $result.Content | ConvertFrom-Json
   $code = $contentBody.masterKey
}
else {
    Write-Error "Couldn't retrive the Azure Function app master key - Warm-up tests not executed"
    return
}

write-host -ForegroundColor blue "Waiting 2 min to let the Azure function app to start"
Start-Sleep -Seconds 120

write-host -ForegroundColor blue "Warming-up Azure Function apps - This will take a few minutes"
& $base\warmup.ps1 -hostname $outputs.Outputs.azFuncHostName.Value -code $code -tenantID $tenantID -clientID $clientID -secret $secret

write-host -ForegroundColor blue "Deployment script terminated"

# Generating outputs
$outputsData = [ordered]@{
    API_URL       = 'https://'+ $outputs.Outputs.azFuncHostName.Value
    API_Code      = $outputs.Outputs.AzFuncAppCode.Value
    TenantID      = $tenantID
    ClientID      = $clientID
    Audience      = 'api://azfunc-' + $clientID
    KeyVault_Name = $outputs.Outputs.AzKeyVaultName.Value
    AzFunctionIPs = $outputs.Outputs.outboundIpAddresses.Value
}

write-host -ForegroundColor magenta "Here are the information you'll need to deploy and configure the Power Application"
$outputsData
