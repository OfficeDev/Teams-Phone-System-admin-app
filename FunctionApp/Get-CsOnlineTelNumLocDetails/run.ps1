using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Initialize PS script
$StatusCode = [HttpStatusCode]::OK
$Resp = ConvertTo-Json @()

# Get query parameters to get telephone number detailed lication - REQUIRED parameter
$TelephoneNumber = [string]$Request.Query.TelephoneNumber
If ([string]::IsNullOrEmpty($TelephoneNumber)){
    $Resp = @{ "Error" = "Missing query parameter - Please provide TelephoneNumber via query string ?TelephoneNumber=(e.g. 12065783601)" }
    $StatusCode =  [HttpStatusCode]::BadRequest
}
else {
    # Remove '+' character if added in the query
    $TelephoneNumber = $TelephoneNumber.Replace('+','').trim()
    Write-Output "Searching location for number: " $TelephoneNumber
}

# Authenticate to Microsoft Teams using service account
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

Write-Host $StatusCode
# Get telephone number emmergency location
If ($StatusCode -eq [HttpStatusCode]::OK) {
    Try {
        # $LocationId = [string](Get-CsOnlineTelephoneNumber -TelephoneNumber $TelephoneNumber -ExpandLocation -ErrorAction:Stop | Select-Object Location).Location.LocationId.Guid
        Write-Host Test $TelephoneNumber
        $LocationId = [string](Get-CsPhoneNumberAssignment -TelephoneNumber $TelephoneNumber -ErrorAction:Stop | Select-Object LocationId).LocationId
        $LocationId 
        If ([string]::IsNullOrEmpty($LocationId)) {
            $Resp = @{}
        } Else { 
            $Resp = Get-CsOnlineLisLocation -LocationId $LocationId -ErrorAction:Stop | Select-Object LocationId,Description,CountryOrRegion,City,Latitude,Longitude | ConvertTo-Json
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
