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

$MSTeamsDModuleLocation = ".\Modules\MicrosoftTeams\4.7.0\MicrosoftTeams.psd1"
Import-Module $MSTeamsDModuleLocation

If ($StatusCode -eq [HttpStatusCode]::OK) {
    Try {
        Connect-MicrosoftTeams -Credential $Credential -ErrorAction:Stop
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
                @{Name='LineURI'; Expression = {(Get-CsPhoneNumberAssignment -AssignedPstnTargetId $_.Identity).TelephoneNumber}}, `
                @{Name='objectID'; Expression = {if ($null -ne $_.objectID) { $_.objectID } else { $_.Identity }}}, `
                @{Name='VoicePolicy'; Expression = {if ($_.VoicePolicy.getType().Name -eq 'UserPolicyDefinition') { $_.VoicePolicy.Name } else { $_.VoicePolicy }}}, `
                @{Name='TeamsCallingPolicy'; Expression = {if ($_.TeamsCallingPolicy.getType().Name -eq 'UserPolicyDefinition') { $_.TeamsCallingPolicy.Name } else { $_.TeamsCallingPolicy }}}, `
                @{Name='OnlineDialOutPolicy'; Expression = {if ($_.OnlineDialOutPolicy.getType().Name -eq 'UserPolicyDefinition') { $_.OnlineDialOutPolicy.Name } else { $_.OnlineDialOutPolicy }}}, `
                @{Name='UserLocation'; Expression = { $_ | Select-Object StateOrProvince,City,Street,PostalCode } }
       
        If([string]::IsNullOrEmpty($userInfos)) {
            Write-Host "User not found"
        }        
        Else {
            Write-Host "User profile info collected."
            Write-Host $userInfos
        }

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

Disconnect-MicrosoftTeams
Get-PSSession | Remove-PSSession

# Trap all other exceptions that may occur at runtime and EXIT Azure Function
Trap {
    Write-Error $_
    Disconnect-MicrosoftTeams
    Get-PSSession | Remove-PSSession
    break
}
