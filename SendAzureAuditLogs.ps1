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
$emailSubject = "Azure User Activities Report $(Get-Date -Format MM/dd/yyyy)"

$emailBody = "<h2>Azure Active Directory Activity Logs</h2>`n"
$emailBody += "<p>Date: $(Get-Date)</p>`n"
$emailBody += "<p><strong>Tenant: $($tenantdomain)</strong></p>`n"
$emailBody += "<p>More detail at: <a href='https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RiskySignIns' target='_blank'>Azure Active Directory Portal</a></p>`n"
$emailBody += "<br>`n"

$datenow = get-date
$datepassed = "{0:s}" -f $datenow.AddHours(-36) + "Z"
$7daysago = $datepassed

# or, AddMinutes(-5)

Write-Output $7daysago

# Get an Oauth 2 access token based on client id, secret and tenant domain
$body1       = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
$body       = @{grant_type="client_credentials";resource=$resource2;client_id=$ClientID;client_secret=$ClientSecret}

$oauth      = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
$oauth2      = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body1

if ($oauth.access_token -ne $null) {

$reqBody='{
        "Message": {
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
        ]

        }
    }' | ConvertFrom-Json

    $reqBody.Message.subject = $emailSubject
    $reqBody.Message.body.contentType = "Html"
    $reqBody.Message.toRecipients.emailAddress.address = $emailRecipient

$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
$headerParams2 = @{'Authorization'="$($oauth2.token_type) $($oauth2.access_token)"}

[uri]$url = "https://graph.windows.net/$tenantdomain/activities/audit?api-version=beta&`$filter=activityDate gt $7daysago"

$i=0

    Write-Output "Fetching data using Uri: $url"
    $myReport = Invoke-RestMethod -Method Get -Uri $url.AbsoluteUri -Headers $headerParams
    if ($myReport.value -ne $null) {

    foreach ( $event in $myReport.value ) {
        
        if ($event.actor.userPrincipalName -notlike "*zzO365TenantServiceAdmin*") { #This if statement for excluding service accounts
            
            #$emailBody += "***********************************************************`n"
            $emailBody += "<p>`n"
            $emailBody += "<strong>Activity Details:</strong><br>`n"
            $emailBody += "Activity: $($event.activity)<br>`n"
            $emailBody += "Activity Date: $($event.activityDate)<br>`n"
            $emailBody += "Activity Operation Type: $($event.activityOperationType)<br>`n"
            $emailBody += "Activity Result Status: $($event.activityResultStatus)<br>`n"
            $emailBody += "Activity Name: $($event.actor.name)<br>`n"
            $emailBody += "Source User PrincipalName: $($event.actor.userPrincipalName)<br>`n"
            $emailBody += "Targeted User: $($event.targets.userPrincipalName)<br>`n"
            $emailBody += "Target Name: $($event.targets.name)<br>`n"
            #$emailBody += "Failure Reason: $($event.failureReason)<br>`n"
            #$emailBody += "MFA Result: $($event.mfaResult)<br>`n"
            #$emailBody += "MFA Method: $($event.mfaMethod)<br>`n"
            #$emailBody += "IP city: $($event.location.city)<br>`n"
            #$emailBody += "IP state: $($event.location.state)<br>`n"
            #$emailBody += "IP country: $($event.location.country)<br>`n"
            #$emailBody += "IP org: $($event.org)<br>`n"

            }
            #$emailBody += "***********************************************************`n"
            $emailBody += "</p>`n"
        
        }
       $reqBody.message.body.content = $emailBody
       Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($emailSender)/sendMail" -Headers @{'Authorization'="$($oauth2.token_type) $($oauth2.access_token)"; 'Content-type'="application/json"} -Body ($reqBody | ConvertTo-Json -Depth 4 | Out-String)
    }    
    #$url = ($myReport.Content | ConvertFrom-Json).'@odata.nextLink'
    $i = $i+1

}

else {

    Write-Host "ERROR: No Access Token"
}