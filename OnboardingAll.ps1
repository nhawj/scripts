[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
Add-Type -Assembly System.Web

#Check for Active Directory Module. Install module if not installed yet.
Function checkForActiveDirectoryModule {

if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Write-Host "ActiveDirectory Module Already Installed"
} 
else {
    try {
        Install-Module -Name ActiveDirectory -AllowClobber -Confirm:$False -Force  
    }
    catch [Exception] {
        $_.message 
        exit
    }
}


}

#Run Function AD module check
checkForActiveDirectoryModule

Import-Module ActiveDirectory

Function Get-File {
$openFileDialog = New-Object windows.forms.openfiledialog   
       [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.filter = "XLSX (*.xlsx) | *.xlsx"
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.FileName
     
}
$file = get-file 

$sheetName = "On Board Request"

$objExcel = New-Object -ComObject Excel.Application

$objExcel.visible = $false

$workBook = $objExcel.WorkBooks.Open($file)
$workSheet = $workBook.sheets.item($sheetName)

$Output = [pscustomobject][ordered]@{
    Name = $workSheet.Range("B3").Text
    Title = $workSheet.Range("B4").Text
    Manager = $workSheet.Range("B5").Text
    CrmAccess = $workSheet.Range("B8").Text
    UserToCopyPerm = $workSheet.Range("B11").Text
    }

$Output | Format-Table

$title = $workSheet.Range("B4").Text
$manager = $workSheet.Range("B5").Text
$CrmAccess = $workSheet.Range("B8").Text

$managerName = $workSheet.Range("B3").Text
$position = $managerName.IndexOf(" ")
$managerFirstName = $managerName.Substring(0, $position)
$managerLastName = $managerName.Substring($position+1)
$managerUserName = [string]::Format("{0}{1}",($managerFirstName.Substring(0,1),$managerLastName)).ToLower()
$managerEmail = $managerUserName + "@domain.com"

$Name = $workSheet.Range("B3").Text
$position = $Name.IndexOf(" ")
$firstName = $Name.Substring(0, $position)
$lastName = $Name.Substring($position+1)

$copiedName = $workSheet.Range("B11").Text
$position = $copiedName.IndexOf(" ")
$copiedFirstName = $copiedName.Substring(0, $position)
$copiedLastName = $copiedName.Substring($position+1)

$userName = [string]::Format("{0}{1}",($firstName.Substring(0,1),$lastName)).ToLower() #Convert first and last name to first name initial + last name
$copiedUserName = [string]::Format("{0}{1}",($copiedFirstName.Substring(0,1),$copiedLastName)).ToLower() #Convert first and last name to first name initial + last name

$email = $userName + "@domain.com" #Combine username with domain

$CopiedEmail = $copiedUserName + "@domain.com" #Combine username with domain


#************************************************************************************************************************
#************************************************************************************************************************
#************************************************************************************************************************
#************************************************************************************************************************



$ExchangeCred= "user@domain.com" #Admin credential
$365Cred= get-credential
$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
. "$CreateEXOPSSession\CreateExoPSSession.ps1"
Connect-EXOPSSession -UserPrincipalName $ExchangeCred
Connect-MsolService -Credential $365Cred



#************************************************************************************************************************
#************************************************************************************************************************
#++++++++++++++++++++Setup Active Directory Account++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#************************************************************************************************************************
#************************************************************************************************************************


[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
do {
            if ($firstName -eq "")
                {
                [Microsoft.VisualBasic.Interaction]::MsgBox("No Value was found for First Name or the Cancel button was pressed.  This user has not been created.","okonly","ERROR")
                exit
                }

            if ($lastName -eq "")
                {
                [Microsoft.VisualBasic.Interaction]::MsgBox("No Value was found for Last Name or the Cancel button was pressed.  This user has not been created.","okonly","ERROR")
                exit
                }
          
        $strname = [string]::Format("{0}, {1}",$lastName,$firstName)
        $checkname = Get-ADUser -filter * | where name -Like $strname
            if ($checkname.name -like $strname)
                {
                [Microsoft.VisualBasic.Interaction]::MsgBox("$strname is already in use please choose another Name","okonly","ERROR")
                $checkname = 'Yes'
                }
            else
                {
                $checkname = 'no'
                }
 }
  while ($checkname -eq "Yes" -or $Check -eq "No")
    #This loop is creating the userid and checking Active Directory if the userid already exists. If it does loop back to enter a different userid.
	#UserID will try to be predicted my domain has a first name first initial of last name like johnd
	#If you want to change the prediction to first initial of first name last name like jdoe do the following
	#change the $predictuserID variable to [string]::Format("{0}{1}",($firstname.Substring(0,1),$lastname)).ToLower()
           
        do {
            $predictuserID = [string]::Format("{0}{1}",($firstName.Substring(0,1),$lastName)).ToLower()
            $UserID = [string]::Format("{0}{1}",($firstName.Substring(0,1),$lastName)).ToLower()
                if ($UserID -eq "")
                {
                [Microsoft.VisualBasic.Interaction]::MsgBox("No Value was entered for UserID or the Cancel button was pressed.  This user has not been created.","okonly","ERROR")
                    exit
                }
                $IDerror = 'Yes'
                try
                {
                    $checkuserid = Get-ADUser $UserID -ErrorAction SilentlyContinue
                    if ($checkuserid.UserPrincipalName -like $UserID) 
                    {
                    [Microsoft.VisualBasic.Interaction]::MsgBox("$Userid is already in use please choose another UserID","okonly","ERROR")
                    }                
                } 
                catch 
                    {
                    $IDerror = 'no'
                    }
            }
            while ($IDerror -eq "Yes")
$pass = [Web.Security.Membership]::GeneratePassword(8,3)
$strPass = ConvertTo-SecureString $pass -AsPlainText -Force

$strFirstname = $firstName.ToString()
    $strLastname = $lastName.ToString()
    $strUserID = $UserID.ToString()
    #$strDeptGroup = $OUGroup.ToString()
    $date = (Get-Date).ToString()
New-ADUser -Name $strname -SamAccountName $strUserID -DisplayName $strname -GivenName $strFirstname -Surname $strLastname -UserPrincipalName $strUserID@domain.com -Description "Created by $env:USERNAME on $date" -AccountPassword $strPass -Enabled $true

            if ($CrmAccess -eq "Y") {

                        New-MsolUser -DisplayName $strname -FirstName $firstName -LastName $lastName -UserPrincipalName $UserID@domain.com -UsageLocation US -LicenseAssignment DOMAIN:INTUNE_A,DOMAIN:ENTERPRISEPACK,DOMAIN:DYN365_ENTERPRISE_PLAN1 #Create User with specified licenses
                        }
                        else {
                        New-MsolUser -DisplayName $strname -FirstName $firstName -LastName $lastName -UserPrincipalName $UserID@domain.com -UsageLocation US -LicenseAssignment DOMAIN:INTUNE_A,DOMAIN:ENTERPRISEPACK #Create User with specified licenses
                        }
$email = $strUserID + "@domain.com"


#Check to see if mailbox is created before proceeding.
#=============================================================================================
#=============================================================================================
$checkifmailboxexists = get-mailbox $UserID@domain.com -erroraction silentlycontinue
echo "Checking to see if mailbox is created please wait for completion notification","okonly","Mailbox Check"
do {
       $checkifmailboxexists = get-mailbox $UserID@domain.com -erroraction silentlycontinue
       Sleep 10
}
While ($checkifmailboxexists -eq $Null)

set-MailboxRegionalConfiguration -identity $UserID@domain.com -Language 1033 -TimeZone "Central Standard Time" -DateFormat MM/dd/yyyy -Timeformat HH:mm
$managerObj = Get-AzureADUser -ObjectId $manager
Set-AzureADUser -ObjectId $email -JobTitle $title
Set-AzureADUserManager -ObjectId $email -RefObjectId $managerObj.ObjectId




$Recipient = $CopiedEmail
$Recipient2 = $email


    $groups = @(Get-MsolGroup -All)
    $distroGroups = @(Get-DistributionGroup -ResultSize Unlimited)
    foreach ($group in $groups){
    
    echo $group.DisplayName
    $group_GUID = $group.ObjectId 
    $GroupMembers = @(Get-MsolGroupMember -GroupObjectId $group_GUID)
        foreach($GroupMember in $GroupMembers){
            if($GroupMember.EmailAddress -eq $Recipient.UserPrincipalName) { 
                Add-MsolGroupMember -GroupObjectId $group.ObjectId -GroupMemberObjectId $Recipient2.ObjectId
                Add-DistributionGroupMember -Identity $group.ObjectId -Member $Recipient2
                echo "User added to" $group.ObjectId

                
    
        }
        else 
        {
        cls
            echo 'User is not in any groups'
            cls
    }
}            
}

                foreach($distroGroup in $distroGroups){
              
                $DGMs = @(Get-DistributionGroupMember -Identity $distroGroup.Name)
                    foreach ($dgm in $DGMs){
                    if ($dgm.PrimarySmtpAddress -eq $Recipient){
       
                    echo 'User Found In Group' $distroGroup.Name
                    Add-DistributionGroupMember -identity $distroGroup.Name -Member $Recipient2
    

}
}  
}
    

#SharePoint online Admin site URL
$SPOAdmiURL    = "https://YourSPDomain-admin.sharepoint.com"


#Url of the SharePoint Online Site
$SPOSiteURL = 'https://YourSPDomain-portal4.sharepoint.com'

#User used as reference
$ReferenceUser = $WPFSource.SelectedItem
#The actual user that needs to be added to Groups
$ActualUser    = $WPFDestination.SelectedItem



#Connect to SharePoint Online using the credentials
Connect-SPOService -Url $SPOAdmiURL -Credential $cred

#Get the SharePoint Online Site Object
$site = Get-SPOSite $SPOSiteURL

#Get the user object of the reference user
$userSPO = Get-SPOUser -Site $site -LoginName $CopiedEmail

#Loop through Groups and add the actual user
$userSPO.Groups | Foreach-Object {
    
    #Fetch Group Object that the reference user is part of
    $groupSPO = Get-SPOSiteGroup -Site $site -Group $_

    #Add 'ActualUser' to the same group that the reference user is part of
    Add-SPOUser -Site $SPOSiteURL -LoginName $email  -Group $groupSPO.LoginName
    echo "Added " $email " to SharePoint group " $groupSPO.LoginName
    
}         # }

#************************************************************************************************************************
#************************************************************************************************************************
#************************************************************************************************************************
#************************************************************************************************************************



if ($CrmAccess -eq "Y") {
            
            $path = 'Path.ps1' #Specify the path to the powershell script for copying CRM roles goes here
            Invoke-Expression "& '$path'"
            }
            else {
            exit
            }

#Close Excel Workbook and exit
$workBook.Close()
$objExcel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet)


[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

