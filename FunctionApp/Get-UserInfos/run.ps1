using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Initialize PS script
$StatusCode = [HttpStatusCode]::OK
$Resp = ConvertTo-Json @()

# Get query parameters to search user profile info - REQUIRED parameter
$SearchString = $Request.Query.SearchString
If ([string]::IsNullOrEmpty($SearchString)){
    $Resp = @{ "Error" = "Missing query parameter - Please provide UPN via query string ?objectId=" }
    $StatusCode =  [HttpStatusCode]::BadRequest
}

# Authenticate to AzureAD and MicrosofTeams using service account
$Account = $env:AdminAccountLogin 
$PWord = ConvertTo-SecureString -String $env:AdminAccountPassword -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Account, $PWord

$MSTeamsDModuleLocation = ".\Modules\MicrosoftTeams\4.0.0\MicrosoftTeams.psd1"
Import-Module $MSTeamsDModuleLocation
# $AzureADModuleLocation = ".\Modules\AzureAD\2.0.2.140\AzureAD.psd1"
# Import-Module $AzureADModuleLocation -UseWindowsPowerShell

If ($StatusCode -eq [HttpStatusCode]::OK) {
    Try {
        Connect-MicrosoftTeams -Credential $Credential -ErrorAction:Stop
#        Connect-AzureAD -Credential $Credential -ErrorAction:Stop
    }
    Catch {
        $Resp = @{ "Error" = $_.Exception.Message }
        $StatusCode =  [HttpStatusCode]::BadGateway
        Write-Error $_
    }
}

function Get-DialPolicyDisplayName {
    param($OnlineDialOutPolicy)
    switch($OnlineDialOutPolicy){
        "DialoutCPCDisabledPSTNInternational" { return "Any destination" }
        "DialoutCPCDisabledPSTNDomestic"      { return "In the same country or region as the organizer"}
        "DialoutCPCandPSTNDisabled"           { return "Not allowed"}
        default { return $OnlineDialOutPolicy }
    }
}

# Get User profile infos
If ($StatusCode -eq [HttpStatusCode]::OK) {
    Try {
        # Get user general infos from Teams Communication Services
        $userInfos = Get-CsOnlineUser $SearchString -ErrorAction:Stop | Select-Object -Property DisplayName, UserPrincipalName, UsageLocation, EnterpriseVoiceEnabled, HostedVoiceMail, `
                @{Name='LineURI'; Expression = {if ($_.LineURI -like '*+*') { $_.LineURI } else { '+' + $_.LineURI }}}, `
                @{Name='objectID'; Expression = {if ($null -ne $_.objectID) { $_.objectID } else { $_.Identity }}}, `
                @{Name='VoicePolicy'; Expression = {if ($_.VoicePolicy.getType().Name -eq 'UserPolicyDefinition') { $_.VoicePolicy.Name } else { $_.VoicePolicy }}}, `
                @{Name='TeamsCallingPolicy'; Expression = {if ($_.TeamsCallingPolicy.getType().Name -eq 'UserPolicyDefinition') { $_.TeamsCallingPolicy.Name } else { $_.TeamsCallingPolicy }}}, `
                @{Name='OnlineDialOutPolicy'; Expression = {if ($_.OnlineDialOutPolicy.getType().Name -eq 'UserPolicyDefinition') { $_.OnlineDialOutPolicy.Name } else { $_.OnlineDialOutPolicy }}}, `
                @{Name='UserLocation'; Expression = { $_ | Select-Object StateOrProvince,City,Street,PostalCode } }
       
        Write-Host $userInfos

        Write-Host "User profile info collected."
        ##################################################################################################################################
        # Get user assigned licenced for PSTN calling from AzureAD
        ##################################################################################################################################
        # $CallingPlan = Get-AzureADUserLicenseDetail -ObjectId $userInfos.objectID | Where-Object { $_.SkuPartNumber -like "MCOPSTN*"} | Select-Object SkuPartNumber
        # Write-Host "User calling plan sku collected."
        # if (-not([string]::IsNullOrEmpty($CallingPlan))) {
        #     $userInfos | Add-Member -MemberType NoteProperty -Name 'Calling Plan' -Value $CallingPlan.SkuPartNumber 
        # } else {
        #     $userInfos | Add-Member -MemberType NoteProperty -Name 'Calling Plan' -Value $null 
        # }
        ##################################################################################################################################
        # Get user defined Emergency Location - Code commented for futur use (manage users Emergency Location)
        ##################################################################################################################################
        # $EmergencyLocation = Get-CsOnlineVoiceUser -Identity $SearchString -ExpandLocation | Select-Object Location
        # Write-Host "User emergency location collected."
        # if (-not([string]::IsNullOrEmpty($EmergencyLocation))) {
        #     $userInfos | Add-Member -MemberType NoteProperty -Name 'Location Id' -Value $EmergencyLocation.location.locationId.Guid 
        # } else {
        #     $userInfos | Add-Member -MemberType NoteProperty -Name 'Location Id' -Value $null 
        # }
        ##################################################################################################################################
        $Resp = $userInfos | ConvertTo-Json
    }
    Catch {
        $Resp = @{ "Error" = $_.Exception.Message }
        $StatusCode =  [HttpStatusCode]::BadGateway
        Write-Error $_
    }
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $StatusCode
    ContentType = 'application/json'
    Body = $Resp
})

#Disconnect-AzureAD
Disconnect-MicrosoftTeams
Get-PSSession | Remove-PSSession

# Trap all other exceptions that may occur at runtime and EXIT Azure Function
Trap {
    Write-Error $_
#    Disconnect-AzureAD
    Disconnect-MicrosoftTeams
    Get-PSSession | Remove-PSSession
    break
}
