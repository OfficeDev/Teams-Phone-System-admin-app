using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Initialize PS script
$StatusCode = [HttpStatusCode]::OK
$Resp = ConvertTo-Json @()

# Check if query parameter asks for a static list of OnlineDialOutPolicies
$static = $Request.Query.static
If ($static -eq 'true'){
$DialoutPolicies = @'
[
      { "Identity": "Tag:DialoutCPCDisabledPSTNInternational" , "DisplayName": "Any destination" },
      { "Identity": "Tag:DialoutCPCDisabledPSTNDomestic"      , "DisplayName": "In the same country or region as the organizer"},
]
'@ 
$Resp = $DialoutPolicies | ConvertFrom-Json | ConvertTo-Json
}
Else {
    
    # # Authenticate to MicrosofTeams using service account
    $Account = $env:AdminAccountLogin 
    $PWord = ConvertTo-SecureString -String $env:AdminAccountPassword -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Account, $PWord

    $MSTeamsDModuleLocation = ".\Modules\MicrosoftTeams\4.0.0\MicrosoftTeams.psd1"
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

    # Get CS Online Dialout Policies
    If ($StatusCode -eq [HttpStatusCode]::OK) {
        Try {
            $Resp = Get-CSOnlineDialOutPolicy | select-object -Property Identity,@{Name='DisplayName';Expression={$_.Identity.Replace('Tag:','')}} | ConvertTo-Json
        }
        Catch {
            $Resp = @{ "Error" = $_.Exception.Message }
            $StatusCode =  [HttpStatusCode]::BadGateway
            Write-Error $_
        }
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
    Write-Error $_.Exception.Message
    Disconnect-MicrosoftTeams
    break
}
