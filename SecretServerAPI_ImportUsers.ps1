[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$application = "https://YourSecretServerDomainGoesHere.com"
$api = '$application/api/v1'


Function Get-Token
{
    [CmdletBinding()]
    Param(
        [Switch] $UseTwoFactor
    )

    $creds = @{
        username = "admin@YourDomain.com"
        password = Read-Host -assecurestring "Password: "
        grant_type = "password"

    };

    $headers = $null
    If ($UseTwoFactor) {
        $headers = @{
            "OTP" = (Read-Host -Prompt "Enter your OTP for 2FA: ")
        }
    }

    try
    {
        $response = Invoke-RestMethod "$application/oauth2/token" -Method Post -Body $creds -Headers $headers;
        $token = $response.access_token;
        return $token;
    }
    catch
    {
        $result = $_.Exception.Response.GetResponseStream();
        $reader = New-Object System.IO.StreamReader($result);
        $reader.BaseStream.Position = 0;
        $reader.DiscardBufferedData();
        $responseBody = $reader.ReadToEnd() | ConvertFrom-Json
        Write-Host "ERROR: $($responseBody.error)"
        return;
    }


}
$token = ""

$token2 = get-token -UseTwoFactor
Write-Host $token

$headers = $null
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer $token2")


$Path = "Path to your list of users to import.csv"

$users = Import-CSV -Path $Path

Foreach ($user in $users)
{
    $username = $user.username
    $password = $user.password
    $displayName = $user.displayName
    $emailAddress = $user.emailAddress
    $oathTwoFactor = $user.oathTwoFactor
    

Write-Host ""
Write-Host "----- Create a User -----"

$domain = "@YourDomain.com"

$userCreateArgs = @{
        UserName = $username
        password = $password
        DisplayName = $displayName
        enabled = $true
        emailAddress = $emailAddress
        isApplicationAccount = $false
        oathTwoFactor = $true
    } | ConvertTo-Json

try
{
$createUser = Invoke-RestMethod "$api/users" -Headers $headers -Method Post -ContentType "application/json" -Body $userCreateArgs
    Write-Host "New User ID : " $createUser.id,$createUser.userName,$password
}

catch
{
    Write-Debug "----- Exception -----"
    Write-Host  $_.Exception.Response.StatusCode
    Write-Host  $_.Exception.Response.StatusDescription
    $result = $_.Exception.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($result)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd() | ConvertFrom-Json
    Write-Host  $responseBody.errorCode " - " $responseBody.message
    foreach($modelState in $responseBody.modelState)
    {
        $modelState
}

}

}