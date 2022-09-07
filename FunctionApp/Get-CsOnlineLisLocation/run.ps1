using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Initialize PS script
$StatusCode = [HttpStatusCode]::OK
$Resp = ConvertTo-Json @()

# Get query parameters to search non assigned number based on location - OPTIONAL parameter
$Location = $Request.Query.Location

# Authenticate to MicrosofTeams using service account
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

# Get list of Emmergency Locations
If ($StatusCode -eq [HttpStatusCode]::OK) {
    Try {
        If ([string]::IsNullOrEmpty($Location)){
            $Resp = Get-CsOnlineLisLocation -ErrorAction "Stop" | Select-Object LocationId,Description,CountryOrRegion,City,Latitude,Longitude | ConvertTo-Json
        }
        Else {
            Write-Host $Location
            $Resp = Get-CsOnlineLisLocation -CountryOrRegion $Location -ErrorAction "Stop" | Select-Object LocationId,Description,CountryOrRegion,City,Latitude,Longitude | ConvertTo-Json
        }
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
    break
}
