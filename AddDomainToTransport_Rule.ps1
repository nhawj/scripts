Function checkForExchangeOnlineModule {

if (Get-Module -ListAvailable -Name Microsoft.Exchange.Management.ExoPowershellModule) {
    Write-Host "Microsoft.Exchange.Management.ExoPowershellModule Module Already Installed"
} 
else {
    try {
        Install-Module -Name Microsoft.Exchange.Management.ExoPowershellModule -AllowClobber -Confirm:$False -Force  
    }
    catch [Exception] {
        $_.message 
        exit
    }
}


}

checkForExchangeOnlineModule

Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName |
where { $_ -notmatch "_none_" } | select -First 1)

$rules = @("Block Rule1", "Block Rule2", "Block Rule2") #Specify rules to modify

$NewDomainToAdd = Read-Host "Enter new domain to exclude: " #Prompt for new domain to exclude from specified rules above


foreach ($file in $rules) {

#Get-content -Path .\$file.txt
Add-Content -Path YourPathGoesHere\$file.txt -Value $NewDomainToAdd

}

$UserCredential = Get-Credential
$EXOSession = New-EXOPSSession –UserPrincipalName admin@YourDomain.com
Import-PSSession $EXOSession -AllowClobber

foreach ($rule in $rules) {

 Set-TransportRule $rule -ExceptIfSenderDomainIs (get-content YourPathGoesHere\$rule.txt)
 Get-TransportRule $rule | select -expandproperty ExceptIfSenderDomainIs | Out-GridView
}

Get-PSSession | Remove-PSSession

