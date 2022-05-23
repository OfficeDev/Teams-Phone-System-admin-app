Function Get-jsonSchema (){
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string[]]$schemaName
    )

Switch ($schemaName) {

# JSON schema definition for Set-CsOnlineVoiceUser     
'Set-CsOnlineVoiceUser' { Return @'
    {
        "type": "object",
        "title": "Set-CsOnlineVoiceUser API JSON body definition",  
        "required": [
            "Identity"
        ],
        "properties": {
            "Identity": {
                "type": "string",
                "title": "Specifies the identity of the target user",
                "examples": [
                    "jphillips@contoso.com",
                    "sip:jphillips@contoso.com",
                    "98403f08-577c-46dd-851a-f0460a13b03d"
                ]
            },  
            "TelephoneNumber": {
                "type": "string",
                "title": "Specifies the telephone number to be assigned to the user. The value must be in E.164 format: +14255043920. Setting the value to $Null clears the user's telephone number.",
                "examples": [
                    "+12065783601"
                ],
                "pattern": "^(\\+)"
            },  
            "LocationID": {
                "type": "string",
                "title": "Specifies the unique identifier of the emergency location to assign to the user. This parameter is required for users based in the US",
                "examples": [
                    "8fe2cb97-ffb6-403c-9233-73e258029502"
                ],
                "pattern": "^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$"
            }
        }
    }
'@ }

# JSON schema definition for Grant-CsTeamsCallingPolicy     
'Grant-CsTeamsCallingPolicy' { Return @'
    {
        "type": "object",
        "title": "Grant-CsTeamsCallingPolicy API JSON body definition",  
        "required": [
            "Identity"
        ],
        "properties": {
            "Identity": {
                "type": "string",
                "title": "Specifies the identity of the target user",
                "examples": [
                    "jphillips@contoso.com",
                    "sip:jphillips@contoso.com",
                    "98403f08-577c-46dd-851a-f0460a13b03d"
                ]
            },  
            "PolicyName": {
                "type": "string",
                "title": "The name of the policy being assigned. To remove an existing user level policy assignment, specify PolicyName as null.",
                "examples": [
                    "CallingPlan"
                ]
            }
        }
    }
'@ }

# JSON schema definition for Grant-CsDialoutPolicy     
'Grant-CsDialoutPolicy' { Return @'
    {
        "type": "object",
        "title": "Grant-CsDialoutPolicy API JSON body definition",  
        "required": [
            "Identity"
        ],
        "properties": {
            "Identity": {
                "type": "string",
                "title": "Specifies the identity of the target user",
                "examples": [
                    "jphillips@contoso.com",
                    "sip:jphillips@contoso.com",
                    "98403f08-577c-46dd-851a-f0460a13b03d"
                ]
            },  
            "PolicyName": {
                "type": "string",
                "title": "The name of the policy being assigned. To remove an existing user level policy assignment, specify PolicyName as null.",
                "examples": [
                    "tag:DialoutCPCandPSTNInternational",
                    "tag:DialoutCPCDomesticPSTNInternational"
                ]
            }
        }
    }
'@ }

# No match found - Return empty JSON definition  
Default { Return @'
    {}
'@ }

} }