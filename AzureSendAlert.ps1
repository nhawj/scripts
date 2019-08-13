Function checkForAzureModule {

if (Get-Module -ListAvailable -Name Azure) {
    Write-Host "Azure Module Already Installed"
} 
else {
    try {
        Install-Module -Name Azure -AllowClobber -Confirm:$False -Force  
    }
    catch [Exception] {
        $_.message 
        exit
    }
}


}

checkForAzureModule

Import-Module Azure

# This script will require the Web Application and permissions setup in Azure Active Directory
$ClientID      = ""             # Should be a ~35 character string insert your info here
$ClientSecret   = ""         # Should be a ~44 character string insert your info here
$loginURL       = "https://login.microsoftonline.com/"
$resource       = "https://graph.microsoft.com"
$tenantdomain   = "yourdomain.com"

$daterange            # For example, contoso.onmicrosoft.com
$emailSender = "admin@YourDomain.com"
$emailRecipient = "admin@YourDomain.com"
$emailRecipient2 = "" # Used for texting alert, i.e. XXXXXXXXXX@CellProviderMessagingDomain.com
$emailSubject = "Risky Sign-ins Report $(Get-Date -Format MM/dd/yyyy)"

$emailBody = "<h2>Azure Active Directory Risky Sign-ins Report</h2>`n"
$emailBody += "<p>Date: $(Get-Date)</p>`n"
$emailBody += "<p><strong>Tenant: $($tenantdomain)</strong></p>`n"
$emailBody += "<p>More detail at: <a href='https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RiskySignIns' target='_blank'>Azure Active Directory Portal</a></p>`n"
$emailBody += "<br>`n"

$body       = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
$oauth      = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body


if ($oauth.access_token -ne $null) {
    
    $reqBody='{
        "message": {
        "subject": "",
        "body": {
            "contentType": "",
            "content": ""
        },
        "toRecipients": [
            {
            "emailAddress": {
                "address": ""
            }
            }
        ],
        "CcRecipients": [
            {
            "emailAddress": {
                "address": ""
            }
            }
        ]
        }
    }' | ConvertFrom-Json

    $reqBody.message.subject = $emailSubject
    $reqBody.message.body.contentType = "Html"
    $reqBody.message.toRecipients.emailAddress.address = $emailRecipient
    $reqBody.Message.CcRecipients.emailAddress.address = $emailRecipient2
    
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}

    [uri]$uriGraphEndpoint = "https://graph.microsoft.com/beta/identityRiskEvents?`$filter=riskEventDateTime gt $(Get-Date -date (Get-Date).AddDays(-30).ToUniversalTime() -Format o) and riskEventStatus eq 'active'"

    $response = Invoke-RestMethod -Method Get -Uri $uriGraphEndpoint.AbsoluteUri -Headers $headerParams

    if ($response.value -ne $null) {


        foreach ( $event in $response.value ) {
            
            $emailBody += "<p>`n"
            $emailBody += "User: $($event.userDisplayName)<strong><br>`n"
            $emailBody += "UserPrincipalName: $($event.userPrincipalName)<strong><br>`n"
            $emailBody += "Event time: $($event.riskEventDateTime)<br>`n"
            $emailBody += "Risk type: $($event.riskEventType)<br>`n"
            $emailBody += "Risk level: $($event.riskLevel)<br>`n"
            $emailBody += "Risk status: $($event.riskEventStatus)<br>`n"

            if ( $event.ipAddress -ne $null) {
                        
                $emailBody += "IP: $($event.ipAddress)<br>`n"

                [uri]$uriIpinfo = "https://ipinfo.io/$($event.ipAddress)"
                $ipInfo = Invoke-RestMethod -Method Get -Uri $uriIpinfo.AbsoluteUri

                if ($ipInfo.country -ne "") { $emailBody += "IP country: $($ipInfo.country)<br>`n" }
                if ($ipInfo.city -ne "") { $emailBody += "IP city: $($ipInfo.city)<br>`n" }
                if ($ipInfo.org -ne "") { $emailBody += "IP org: $($ipInfo.org)<br>`n" }


            }

            $emailBody += "</p>`n"

        }

        $reqBody.message.body.content = $emailBody
        Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($emailSender)/sendMail" -Headers @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"; 'Content-type'="application/json"} -Body ($reqBody | ConvertTo-Json -Depth 4 | Out-String)
        
    }


} 

else {

    Write-Output "ERROR: No Access Token"

} 
