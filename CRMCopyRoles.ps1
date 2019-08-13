Add-Type -AssemblyName system.drawing
Add-Type -AssemblyName System.Windows.Forms;

Function checkForXRMPowershellModule {

if (Get-Module -ListAvailable -Name Microsoft.Xrm.Data.PowerShell) {
    Write-Host "Microsoft.Xrm.Data.PowerShell Module Already Installed"
} 
else {
    try {
        Install-Module -Name Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -AllowClobber -Confirm:$False -Force  
    }
    catch [Exception] {
        $_.message 
        exit
    }
}


}

checkForXRMPowershellModule

Import-Module Microsoft.Xrm.Data.PowerShell


$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(600,400)

  
$conn = Get-CrmConnection -InteractiveMode

$users = Get-CrmRecords -EntityLogicalName systemuser -FilterAttribute isdisabled -FilterOperator eq -FilterValue $false -Fields systemuserid,fullname
[array] $userid = $users.CrmRecords | Select-Object fullname, systemuserid


[array]$Array = $userid | sort

# This Function Returns the Selected Value and Closes the Form

$DropDownBox2 = New-Object System.Windows.Forms.ComboBox
$DropDownBox2.Location = New-Object System.Drawing.Size(20,90) 
$DropDownBox2.Size = New-Object System.Drawing.Size(180,20) 
$DropDownBox2.DropDownHeight = 200
$DropDownBox2.Text = "Copy Roles to" 
$Form.Controls.Add($DropDownBox2)
foreach ($item in $Array) {
                      $DropDownBox2.Items.Add($item)
                              } #end foreach


$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(20,50) 
$DropDownBox.Size = New-Object System.Drawing.Size(180,20) 
$DropDownBox.DropDownHeight = 200
$DropDownBox.Text = "Copy Roles From"
$Form.Controls.Add($DropDownBox) 

foreach ($item in $Array) {
                      $DropDownBox.Items.Add($item)
                              } #end foreach

function copyCrmRoles {




$copyFrom= $DropDownBox.SelectedItem.systemuserid
$referenceUser= get-crmusersecurityroles -conn $conn -UserID $copyFrom #populate the var with the value you selected
[array] $roleid = $referenceUser | Select-Object roleid
[array]$Array3 = $roleid |  Select roleid -ExpandProperty roleid #Format-table -HideTableHeaders | Out-String -Stream
[array] $newRoleID = $Array3 | ?{-not($roleid -contains $_)}
foreach ($id in $newRoleID){
        $listBox.Items.Add($id)
        $copyTo= $DropDownBox2.SelectedItem.systemuserid
        Add-CrmSecurityRoleToUser -conn $conn -UserId $copyTo -SecurityRoleId "$id" -ErrorAction SilentlyContinue
            If($? -eq $true) {
                echo "CRM Roles successfully copied for user" + $copyTo
                }
                Else
                {
                echo "CRM Roles copied failed for user" + $copyTo
                }
                }
}

function runWorkflow {

    $wFUser = $DropDownBox2.SelectedItem.systemuserid
    Invoke-CrmRecordWorkflow -conn $conn -Id $wFUser -WorkflowName "Update Contact with System User (Post-Operation)"
    If($? -eq $true) {
    echo "Post-Operation success for user" + $wFUser
    }
    Else
    {
    echo "Post-Operation failed for user" + $wFUser
    }
}

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Size(10,150) 
$listBox.Size = New-Object System.Drawing.Size(565,200) 



$Form.Controls.Add($listBox) 

$Button = New-Object System.Windows.Forms.Button 
$Button.Location = New-Object System.Drawing.Size(400,30) 
$Button.Size = New-Object System.Drawing.Size(100,50) 
$Button.Text = "Copy Roles" 
$Button.Add_Click({copyCrmRoles}) 
$Form.Controls.Add($Button)

$Button2 = New-Object System.Windows.Forms.Button 
$Button2.Location = New-Object System.Drawing.Size(400,90) 
$Button2.Size = New-Object System.Drawing.Size(100,50) 
$Button2.Text = "Run WorkFlow" 
$Button2.Add_Click({runWorkFlow}) 
$Form.Controls.Add($Button2)
 


$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()