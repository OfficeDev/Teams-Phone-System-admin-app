using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Initialize PS script
$StatusCode = [HttpStatusCode]::OK
$Resp = ConvertTo-Json @()

# Validate the request JSON body against the schema_validator
$Schema = Get-jsonSchema ('Set-CsOnlineVoiceUser')

If (-Not $Request.Body) {
    $Resp = @{ "Error" = "Missing JSON body in the POST request"}
    $StatusCode =  [HttpStatusCode]::BadRequest 
}
Else {
    # Test JSON format and content
    $Result = $Request.Body | ConvertTo-Json | Test-Json -Schema $Schema
    If (-Not $Result){
        $Resp = @{
             "Error" = "The JSON body format is not compliant with the API specifications"
             "detail" = "Verify that the body complies with the definition in module JSON-Schemas and check detailed error code in the Azure Function logs"
         }
         $StatusCode =  [HttpStatusCode]::BadRequest
    }
    else {
        # Set the function variables
        $Id = $Request.Body.Identity
        if([string]::IsNullOrEmpty($Request.Body.TelephoneNumber)) {
            Write-Host "No telephone number detected in request body"
        } Else {
            $telNumber = $Request.Body.TelephoneNumber
            Write-Host "Telephone number detected in request body:" $telNumber
        }
        if([string]::IsNullOrEmpty($Request.Body.LocationID)) {
            Write-Host "No location ID detected in request body"
        } Else {
            $locationID = $Request.Body.LocationID
            Write-Host "Location ID detected in request body:" $locationID
        }        
        Write-Host 'Inputs validated'
    }    
}

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

# Assign Number to user
If ($StatusCode -eq [HttpStatusCode]::OK) {
    Try {
        If (-Not([string]::IsNullOrEmpty($telNumber))){
            If (-Not([string]::IsNullOrEmpty($locationID))){
                $Resp = Set-CsPhoneNumberAssignment -Identity $Id -PhoneNumber $telNumber -LocationId $locationID -PhoneNumberType CallingPlan -ErrorAction "Stop"
            } Else {
                $Resp = Set-CsPhoneNumberAssignment -Identity $Id -PhoneNumber $telNumber -PhoneNumberType CallingPlan -ErrorAction "Stop"
            }
            # Checking if $Resp contains an error message
            If ($null -ne $Resp) {
                $StatusCode =  [HttpStatusCode]::BadRequest
            } Else {
                Write-Host 'Telephone Number' $telNumber 'assigned to ' $Id
            }
        }
        Else {
            $Resp = Remove-CsPhoneNumberAssignment -Identity $Id -RemoveAll -ErrorAction "Stop"
            Write-Host 'Telephone Number unassigned from ' $Id
        }    
    }
    Catch {
        $Resp = @{ "Error" = $_.Exception.Message }
        $StatusCode =  [HttpStatusCode]::BadGateway
        Write-Host "Error"
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
