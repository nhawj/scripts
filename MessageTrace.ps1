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

$UserCredential = Get-Credential
$EXOSession = New-EXOPSSession –UserPrincipalName admin@YourDomain.com
Import-PSSession $EXOSession -AllowClobber

$address = Read-Host "Enter sender's email address: "

Get-MessageTrace -SenderAddress $address -startdate (Get-Date).AddDays(-10) -enddate (Get-Date) -Status Failed | Get-MessageTraceDetail | Where {$_.Event -eq "Transport rule"} | Out-GridView

Get-PSSession | Remove-PSSession