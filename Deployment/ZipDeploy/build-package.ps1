# Use this script if you need to generate a new ZIP package

# Make sure you update and save the MicrosofttTeams module as Azure Function custom modules
save-module -Path .\FunctionApp\Modules -Name MicrosoftTeams -Repository PSGallery -MinimumVersion 4.7.0

# List in the ZIP package all the function app you need to deploy
$packageFiles = @(
    ".\FunctionApp\Get-CSOnlineDialOutPolicy"
    ".\FunctionApp\Get-CsOnlineLisLocation"
    ".\FunctionApp\Get-CsOnlineTelephoneNumber",
    ".\FunctionApp\Get-CsOnlineTelNumLocDetails",
    ".\FunctionApp\Get-CsTeamsCallingPolicy",
    ".\FunctionApp\Get-UserInfos",
    ".\FunctionApp\Grant-CsDialoutPolicy", 
    ".\FunctionApp\Grant-CsTeamsCallingPolicy",
    ".\FunctionApp\keep-alive",
    ".\FunctionApp\Modules",
    ".\FunctionApp\Set-CsOnlineVoiceUser",
    ".\FunctionApp\host.json",
    ".\FunctionApp\profile.ps1",
    ".\FunctionApp\requirements.psd1"
)
$destinationPath = ".\Packages\Azure\artifact.zip"

Compress-Archive -Path $packageFiles -DestinationPath $destinationPath -CompressionLevel optimal -Force