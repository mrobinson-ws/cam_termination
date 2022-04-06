#Requires -RunAsAdministrator
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName PresentationFramework

# Test For Modules
if(-not(Get-Module ExchangeOnlineManagement -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement - view http://worksmart.link/7x for details')
    Exit
}

if(-not(Get-Module AzureAD -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install AzureAD - view http://worksmart.link/7x for details')
    Exit
}

### Start XAML and Reader to use WPF, as well as declare variables for use
[xml]$xaml = @"
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="CAM Termination" Height="350" Width="525">

    <Grid Background="#FFC8C8C8">
        <Button Name="UserButton" Content="Select User" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="135" Height="20" TabIndex="0"/>
        <TextBox Name="UserTextBox" HorizontalAlignment="Left" Height="20" Margin="150,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="357" IsReadOnly="True" IsEnabled="False"/>
        <CheckBox Name="OOOCheckBox" Content="Set Out Of Office Message?" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top" TabIndex="2" IsChecked="True"/>
        <CheckBox Name="LitigationHoldCheckBox" Content="Set Litigation Hold?" HorizontalAlignment="Left" Margin="10,80,0,0" TabIndex="4" VerticalAlignment="Top"/>
        <Button Name="ManagerButton" Content="Select Manager" HorizontalAlignment="Left" Margin="10,35,0,0" VerticalAlignment="Top" Width="135" Height="20" TabIndex="5"/>
        <TextBox Name="ManagerTextBox" HorizontalAlignment="Left" Height="20" Margin="150,35,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="357" IsReadOnly="True" IsEnabled="False"/>
        <RichTextBox Name="TerminationRichTextBox" HorizontalAlignment="Left" Height="90" Margin="10,190,0,0" VerticalAlignment="Top" Width="497" Background="Black" Foreground="#FF00C8C8" IsReadOnly="True">
            <FlowDocument/>
        </RichTextBox>
        <Button Name="TerminateGoButton" Content="Terminate User" HorizontalAlignment="Left" Margin="10,285,0,0" VerticalAlignment="Top" Width="497" Height="24" IsEnabled="False" TabIndex="8"/>
        <TextBox Name="OOOTextBox" HorizontalAlignment="Left" Height="85" Margin="10,100,0,0" TextWrapping="Wrap" Text="User is no longer with Crisis Assistance Ministry, and this email is not monitored.&#xD;&#xA;&#xD;&#xA;Please contact Manager and your emails will be delivered to the appropriate department.&#xD;&#xA;&#xD;&#xA;Thank you." VerticalAlignment="Top" Width="497" TabIndex="7"/>
        <CheckBox Name="GrantSharedCheckbox" Content="Grant Access to Shared Mailbox?" HorizontalAlignment="Left" Margin="314,80,0,0" VerticalAlignment="Top" TabIndex="3" IsEnabled="False"/>
        <CheckBox Name="ConvertToSharedCheckBox" Content="Convert to Shared Mailbox?" HorizontalAlignment="Left" Margin="314,60,0,0" VerticalAlignment="Top" TabIndex="3"/>
    </Grid>

</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
Try{
    $UserForm = [Windows.Markup.XamlReader]::Load($reader)
}
Catch{
    Write-Host "Unable to load Windows.Markup.XamlReader.  Some possible causes for this problem include: .NET Framework is missing, PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."
    Exit
}

#Create Variables For Use In Script Automatically
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $UserForm.FindName($_.Name)}
### End XAML and Variables from XAML

Function Write-RichTextBox {
    Param(
        [System.Windows.Controls.RichTextBox]$TextBox,
        [string]$Text,
        [string]$Color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $TextBox.Document.ContentEnd,$TextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $Text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $Color )
    $TextBox.ScrollToEnd()
}

$OOOCheckBox.Add_Unchecked({
    $ManagerButton.IsEnabled = $false
    $OOOTextBox.IsEnabled = $false
    if (($UserTextBox.Text.Length -gt 2)){
        $TerminateGoButton.IsEnabled = $true
    }
    else{
        $TerminateGoButton.IsEnabled = $false
    }
})

$OOOCheckbox.Add_Checked({
    $ManagerButton.IsEnabled = $true
    $OOOTextBox.IsEnabled = $true
    if (($UserTextBox.Text.Length -gt 2)  -and ($ManagerTextBox.Text.Length -gt 2)){
        $TerminateGoButton.IsEnabled = $true
    }
    else{
        $TerminateGoButton.IsEnabled = $false
    }
})

$ConvertToSharedCheckbox.Add_Checked({
    $GrantSharedCheckbox.IsEnabled = $true
})

$ConvertToSharedCheckbox.Add_Unchecked({
    $GrantSharedCheckbox.IsChecked = $false
    $GrantSharedCheckbox.IsEnabled = $false
})

$UserTextBox.Add_TextChanged({
    if($OOOCheckbox.IsChecked){
        if (($UserTextBox.Text.Length -gt 2)  -and ($ManagerTextBox.Text.Length -gt 2)){
            $TerminateGoButton.IsEnabled = $true
        }
        else{
            $TerminateGoButton.IsEnabled = $false
        }
    }
    else{
        if (($UserTextBox.Text.Length -gt 2)){
            $TerminateGoButton.IsEnabled = $true
        }
        else{
            $TerminateGoButton.IsEnabled = $false
        }
    }
})

$ManagerTextBox.Add_TextChanged({
    if($OOOCheckbox.IsChecked){
        if (($UserTextBox.Text.Length -gt 2)  -and ($ManagerTextBox.Text.Length -gt 2)){
            $TerminateGoButton.IsEnabled = $true
        }
        else{
            $TerminateGoButton.IsEnabled = $false
        }
    }
    else{
        if (($UserTextBox.Text.Length -gt 2)){
            $TerminateGoButton.IsEnabled = $true
        }
        else{
            $TerminateGoButton.IsEnabled = $false
        }
    }
})
### End Logic for enabling/disabling functionality

#Select User
$UserButton.Add_Click({
    $Global:termeduser = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName,SamAccountName,DistinguishedName | sort-Object Name | Out-Gridview -OutputMode Single -Title "Please Select a User"
    if($Global:termeduser.UserPrincipalName -notlike "*@crisisassistance.org"){
        Remove-Variable termeduser -Scope Global
        $UserTextbox.Text = ""
        $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.UserPrincipalName) and your emails will be delivered to the appropriate department.
Thank you.
"@
        $null = [System.Windows.MessageBox]::Show('User has incorrect UPN not ending in crisisassistance.org, please terminate manually, user has been deselected')
    }else{
        $UserTextbox.Text = $Global:termeduser.Name
        if($Global:Manager.UserPrincipalName -ne $Global:termeduser.UserPrincipalName){
            $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.UserPrincipalName) and your emails will be delivered to the appropriate department.
Thank you.
"@
        }else{
            $ManagerTextBox.Text = ""
            $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.UserPrincipalName) and your emails will be delivered to the appropriate department.
Thank you.
"@
            $null = [System.Windows.MessageBox]::Show('Manager cannot match User, please select Manager again')
        }
    }
})

#Select Manager
$ManagerButton.Add_Click({
    $Global:Manager = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName | sort-Object Name | Out-Gridview -OutputMode Single -Title "Please Select the Manager"
    if($Global:Manager.UserPrincipalName -ne  $Global:termeduser.UserPrincipalName){
        $ManagerTextBox.Text = $Global:Manager.Name
        $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.UserPrincipalName) and your emails will be delivered to the appropriate department.
Thank you.
"@
    }else{
        Remove-Variable Manager -Scope Global
        $ManagerTextBox.Text = ""
        $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.UserPrincipalName) and your emails will be delivered to the appropriate department.
Thank you.
"@
$null = [System.Windows.MessageBox]::Show('Manager cannot match User, please select Manager again')
    }
})

#Terminate the user with selected options
$TerminateGoButton.Add_Click({
    #Set Mail Nickname, Hide from GAL, and Disable AD User Account
    Set-ADUser -Identity $Global:termeduser.distinguishedname -replace @{msExchHideFromAddressLists=$True;mailnickname=$Global:termeduser.SamAccountName} -Confirm:$False
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Hid user from GAL and set Mail Nickname`r"
    
    Disable-ADAccount -Identity $Global:termeduser.DistinguishedName -Confirm:$False
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Disabled user in Active Directory`r"

    #Test and Connect to Exchange Online if needed
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$False
    }

    #Set Out of Office Message
    if($OOOCheckBox.IsChecked){
        Set-MailboxAutoReplyConfiguration -Identity $Global:termeduser.UserPrincipalName -ExternalMessage $OOOTextbox.Text -InternalMessage $OOOTextbox.Text -AutoReplyState Enabled -Confirm:$False
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Out of Office Message Set.`r"
    }
    else{
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Out of Office Message not selected`r" -Color "Yellow"
    }

    #Remove AD Groups
    $ADGroups = (Get-ADUser $Global:termeduser.SamAccountName -Properties memberof).memberof
    $ADGroups | ForEach-Object {remove-adgroupmember -identity $_ -member $Global:termeduser.SamAccountName -Confirm:$False}
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Removed user from all Active Directory groups.`r"

    #Test and Connect to Azure AD if needed
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
    }
    
    #Set Litigation Hold
    if($LitigationHoldCheckBox.IsChecked){
        Clear-Variable AssignedLicense -ErrorAction SilentlyContinue
        Clear-Variable E1Assigned -ErrorAction SilentlyContinue
        Clear-Variable E3Assigned -ErrorAction SilentlyContinue
        $UserInfo = Get-AzureADUser -ObjectId $Global:termeduser.UserPrincipalName
        foreach($AssignedLicense in $UserInfo.AssignedLicenses){
            if($AssignedLicense.SkuID -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"){
                $E3Assigned = $True
            }
        }
        if($E3Assigned -eq $True){
            Set-Mailbox $Global:termeduser.UserPrincipalName -LitigationHoldEnabled $true
            Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) already has an Office 365 E3 License (to be removed shortly).  Litigation Hold has been applied.`r"
        }
        else{
            Clear-Variable AvailableLicenseCheck -ErrorAction SilentlyContinue
            Clear-Variable E3License -ErrorAction SilentlyContinue
            $E3License =  Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits | Where-Object {$_.SkuId -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"}
            while($AvailableLicenseCheck -ne $true){
                if($E3License.Enabled-$E3License.ConsumedUnits -ge 1){
                    $AvailableLicenseCheck = $true
                    $TempLicenseAdd = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                    $TempLicenseAdd.SkuID = $E3License.SkuID
                    $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                    $Licenses.RemoveLicenses = "18181a46-0d4e-45cd-891e-60aabd171b4e"
                    $Licenses.AddLicenses = $TempLicenseAdd
                    Set-AzureADUserLicense -ObjectID $Global:termeduser.UserPrincipalName -AssignedLicenses $Licenses
                    Clear-Variable TempLicense -ErrorAction SilentlyContinue
                    Clear-Variable Licenses -ErrorAction SilentlyContinue                   
                }
                else{
                    $null = [System.Windows.MessageBox]::Show("You do not have any E3 licenses to assign, please acquire licenses and try again","License Check","OKCancel","Warning")
                }
            }
            Clear-Variable LitHoldCheck -ErrorAction SilentlyContinue
            while ($LitHoldCheck -ne $true) {
                try {
                    Set-Mailbox $Global:termeduser.UserPrincipalName -LitigationHoldEnabled $true -ErrorAction Stop
                    $LitHoldCheck = $true
                }
                catch {
                    $null = [System.Windows.MessageBox]::Show("Litigation Hold Failed, Waiting 60 Seconds and Trying Again")
                    $LitHoldCheck = $false
                    Start-Sleep -Seconds 60
                }
            }
            Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) did not have an E3, one has been added temporarily to set Litigation Hold.  Litigation Hold has been applied.`r"
        }
    }
    else{
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Litigation Hold not selected`r" -Color "Yellow"
    }

    #Set Shared Mailbox if DropDown selected
    if($ConvertToSharedCheckbox.IsChecked){
        Set-Mailbox $Global:termeduser.UserPrincipalName -Type Shared
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Mailbox Converted to Shared Mailbox`r"
    }
    else{
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Conversion to Shared Mailbox not selected`r" -Color "Yellow"
    }

    if($GrantSharedCheckbox.IsChecked){
        $SharedMailboxUser = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user(s) to share the $username Shared Mailbox with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
            if($SharedMailboxUser){
                Add-MailboxPermission -Identity $Global:termeduser.UserPrincipalName -User $SharedMailboxUser -AccessRights FullAccess -InheritanceType All
                Add-RecipientPermission -Identity $Global:termeduser.UserPrincipalName -Trustee $SharedMailboxUser -AccessRights SendAs -Confirm:$False
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Access granted to the $($Global:termeduser.UserPrincipalName) Shared Mailbox to $($SharedMailboxUser.DisplayName)`r"
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Cancelled Mailbox Shared User Access Selection`r" -Color "Red"
            }
    }
    else{
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Shared Mailbox Permissions Not Selected`r" -Color "Yellow"
    }

    #Move to Disabled User OU
    Move-ADObject -Identity $Global:termeduser.DistinguishedName -TargetPath "OU=Disabled (from Users OU),OU=Users,OU=Crisis Assist,DC=crisisministry,DC=local"
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "User moved to Disabled (From Users OU)`r"

    #Sync to AzureAD/365
    Start-ADSyncSyncCycle -PolicyType Delta

    #Remove remaining M365/AzureAD Groups
    $UserInfo = Get-AzureADUser -ObjectId $Global:termeduser.UserPrincipalName
    $memberships = Get-AzureADUserMembership -ObjectId $Global:termeduser.UserPrincipalName | Where-Object {$_.ObjectType -ne "Role"}| Select-Object DisplayName,ObjectId
    foreach ($membership in $memberships) { 
            $group = Get-AzureADMSGroup -ID $membership.ObjectId
            if ($group.GroupTypes -contains 'DynamicMembership') {
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Skipped M365/AzureAD Group $($group.Displayname) as it is dynamic and will not be applied when next run`r" -Color "Yellow"
            }
            else{
                Try{
                    Remove-AzureADGroupMember -ObjectId $membership.ObjectId -MemberId $UserInfo.ObjectId -ErrorAction Stop
                }Catch{
                    $message = $_.Exception.Message
                    if ($_.Exception.ErrorContent.Message.Value) {
                        $message = $_.Exception.ErrorContent.Message.Value
                    }
                    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Could not remove from M365/AzureAD group $($group.name).  Error:  $message`r" -Color "Yellow"
                }
            
            }
        }
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Removed user from all M365/AzureAD groups.`r"

    #Remove all 365/Azure licensing
    $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    if($UserInfo.assignedlicenses){
        $licenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
        Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $licenses
    }
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "All M365/Azure licenses have been removed`r"

    Remove-Variable termeduser -Scope Global
    Remove-Variable Manager -Scope Global
    $ManagerTextBox.Text = ""
    $UserTextbox.Text = ""
    $TerminateGoButton.IsEnabled = $false
})

$null = $UserForm.ShowDialog()