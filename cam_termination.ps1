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
        <TextBox Name="PasswordTextBox" HorizontalAlignment="Left" Height="20" Margin="150,35,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="178" TabIndex="1"/>
        <Label Content="Enter New Password:" HorizontalAlignment="Left" Margin="29,32,0,0" VerticalAlignment="Top" Width="121" Height="23"/>
        <CheckBox Name="OOOCheckBox" Content="Set Out Of Office Message?" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top" TabIndex="2" IsChecked="True"/>
        <CheckBox Name="LitigationHoldCheckBox" Content="Set Litigation Hold?" HorizontalAlignment="Left" Margin="383,60,0,0" TabIndex="4" VerticalAlignment="Top"/>
        <Button Name="ManagerButton" Content="Select Manager" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top" Width="135" Height="20" TabIndex="5"/>
        <TextBox Name="ManagerTextBox" HorizontalAlignment="Left" Height="20" Margin="150,80,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="357" IsReadOnly="True" IsEnabled="False"/>
        <RichTextBox Name="TerminationRichTextBox" HorizontalAlignment="Left" Height="58" Margin="10,222,0,0" VerticalAlignment="Top" Width="497" Background="Black" Foreground="#FF00C8C8" IsReadOnly="True">
            <FlowDocument/>
        </RichTextBox>
        <Button Name="TerminateGoButton" Content="Terminate User" HorizontalAlignment="Left" Margin="10,285,0,0" VerticalAlignment="Top" Width="497" Height="24" IsEnabled="False" TabIndex="8"/>
        <TextBox Name="OOOTextBox" HorizontalAlignment="Left" Height="85" Margin="10,132,0,0" TextWrapping="Wrap" Text="User is no longer with Crisis Assistance Ministry, and this email is not monitored.&#xD;&#xA;&#xD;&#xA;Please contact Manager and your emails will be delivered to the appropriate department.&#xD;&#xA;&#xD;&#xA;Thank you." VerticalAlignment="Top" Width="497" TabIndex="7"/>
        <Button Name="GeneratePasswordButton" Content="Generate Random Password" HorizontalAlignment="Left" Margin="333,35,0,0" VerticalAlignment="Top" Width="174"/>
        <ComboBox Name="LicenseComboBox" HorizontalAlignment="Left" Margin="10,105,0,0" VerticalAlignment="Top" Width="497" IsReadOnly="True" SelectedIndex="0" FontSize="10" TabIndex="6">
            <ComboBoxItem Name="LicenseSelection1" IsSelected="True" FontSize="10">User is an E1 Office 365 user and license should remain assigned, but Sign-in blocked</ComboBoxItem>
            <ComboBoxItem Name="LicenseSelection2" IsSelected="False" FontSize="10">User is an E1 Office 365 user and license should be removed (user data will be lost after 90 days)</ComboBoxItem>
            <ComboBoxItem Name="LicenseSelection3" IsSelected="False" FontSize="10">User is an E3 Office 365 user and license should be converted to an E1 license and Sign-in blocked</ComboBoxItem>
            <ComboBoxItem Name="LicenseSelection4" IsSelected="False" FontSize="10">User is an E3 Office 365 user and license should be removed (user data will be lost after 90 days)</ComboBoxItem>
            <ComboBoxItem Name="LicenseSelection5" IsSelected="False" FontSize="10">User is an E3 Office 365 user and should be converted to a shared mailbox and license removed</ComboBoxItem>
        </ComboBox>
        <CheckBox Name="GrantSharedCheckbox" Content="Grant Access to Shared Mailbox?" HorizontalAlignment="Left" Margin="181,60,0,0" VerticalAlignment="Top" TabIndex="3"/>
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

### Logic for enabling/disabling functionality
$PasswordTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -ge 2)){
        $TerminateGoButton.IsEnabled = $true
    }
    else{
        $TerminateGoButton.IsEnabled = $false
    }
})

$OOOCheckBox.Add_Unchecked({
    $ManagerButton.IsEnabled = $false
    $OOOTextBox.IsEnabled = $false
})

$OOOCheckbox.Add_Checked({
    $ManagerButton.IsEnabled = $true
    $OOOTextBox.IsEnabled = $true
})

$PasswordTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -gt 0)){
        $TerminateGoButton.IsEnabled = $true
    }
    else{
        $TerminateGoButton.IsEnabled = $false
    }
})

$UserTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -gt 0)){
        $TerminateGoButton.IsEnabled = $true
    }
    else{
        $TerminateGoButton.IsEnabled = $false
    }
})
### End Logic for enabling/disabling functionality

#Select User
$UserButton.Add_Click({
    $Global:termeduser = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName,SamAccountName,DistinguishedName | sort-Object Name | Out-Gridview -OutputMode Single -Title "Please Select a User"
    $UserTextbox.Text = $Global:termeduser.Name
    $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.Name) at $($Global:Manager.UserPrincipalName) and your emails will be delivered to the appropriate department.
Thank you.
"@
})

#Randomly generate a 16 character password with 8 being non-alphanumeric
$GeneratePasswordButton.Add_Click({
    $PasswordTextBox.Text = [System.Web.Security.Membership]::GeneratePassword(16,8)
})

#Select Manager
$ManagerButton.Add_Click({
    $Global:Manager = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName | sort-Object Name | Out-Gridview -OutputMode Single -Title "Please Select the Manager"
    $ManagerTextBox.Text = $Global:Manager.Name
    $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.Name) at $($Global:Manager.UserPrincipalName) and your emails will be delivered to the appropriate department.
Thank you.
"@
})

#Terminate the user with selected options
$TerminateGoButton.Add_Click({
    #Set Mail Nickname, Hide from GAL, and Disable AD User Account
    Set-ADUser -Identity $Global:termeduser.distinguishedname -replace @{msExchHideFromAddressLists=$True;mailnickname=$Global:termeduser.SamAccountName} -Confirm:$False
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Hid user from GAL and set Mail Nickname`r"
    
    $SecurePassword = ConvertTo-SecureString -String $PasswordTextBox.Text -AsPlainText -Force
    Set-ADAccountPassword -Identity $Global:termeduser.SamAccountName -NewPassword $SecurePassword -Reset -Confirm:$False
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Reset user's password.`r"
    
    Disable-ADAccount -Identity $Global:termeduser.DistinguishedName -Confirm:$False
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Disabled user in Active Directory`r"

    #Test and Connect to Exchange Online if needed
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }

    #Set Out of Office Message
    if($OOOCheckBox.IsChecked){
        Set-MailboxAutoReplyConfiguration -Identity $Global:termeduser.UserPrincipalName -ExternalMessage $OOOTextbox.Text -InternalMessage $OOOTextbox.Text -AutoReplyState Enabled -Confirm:$False
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Out of Office Message Set.`r"
    }

    #Remove AD Groups
    $ADGroups = (Get-ADUser $Global:termeduser.SamAccountName -Properties memberof).memberof
    $ADGroups | ForEach-Object {remove-adgroupmember -identity $_ -member $Global:termeduser.SamAccountName}
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Removed user from all Active Directory groups.`r"

    #Set Litigation Hold
    if($LitigationHoldCheckBox.IsChecked){
        Set-Mailbox $Global:termeduser.UserPrincipalName -LitigationHoldEnabled $true
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Litigation hold set.`r"
    }

    #Set Shared Mailbox if DropDown selected
    if($LicenseComboBox.SelectedIndex -eq 4){
        Set-Mailbox $Global:termeduser.UserPrincipalName -Type Shared
    }

    #Sync to AzureAD/365
    Start-ADSyncSyncCycle -PolicyType Delta

    #Test and Connect to Azure AD if needed
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
    }

    $UserInfo = Get-AzureADUser -ObjectId $Global:termeduser.UserPrincipalName
    
    #Remove remaining M365/AzureAD Groups
    $memberships = Get-AzureADUserMembership -ObjectId $Global:termeduser.UserPrincipalName | Where-Object {$_.ObjectType -ne "Role"}| Select-Object DisplayName,ObjectId
    foreach ($membership in $memberships) { 
            $group = Get-AzureADMSGroup -ID $membership.ObjectId
            if ($group.GroupTypes -contains 'DynamicMembership') {
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Skipped M365/AzureAD Group $($group.Displayname) as it is dynamic`r" -Color "Yellow"
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
        Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Removed user from M365/AzureAD groups.`r"

    switch ($LicenseComboBox.SelectedIndex) {
        0 {
            # Verify E1 license, signin blocked on all users, keep licenses
            Clear-Variable AssignedLicense -ErrorAction SilentlyContinue
            Clear-Variable E1Assigned -ErrorAction SilentlyContinue
            Clear-Variable E3Assigned -ErrorAction SilentlyContinue

            foreach($AssignedLicense in $UserInfo.AssignedLicenses){
                if($AssignedLicense.SkuID -eq "18181a46-0d4e-45cd-891e-60aabd171b4e"){
                    $E1Assigned = $True
                }
                if($AssignedLicense.SkuID -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"){
                    $E3Assigned = $True
                }
            }
            if(($E1Assigned -eq $True) -and ($E3Assigned -ne $True)){
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) has an E1 Office 365 license (no other licenses removed), Sign-in has been blocked.`r"
            }
            elseif($E3Assigned -eq $True){
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) has an Office 365 E3 License.  Sign-in has been blocked.  Please verify licensing manually.`r"
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) does NOT have an E1 Office 365 License.  Sign-in has been blocked.  Please verify licensing manually.`r"
            }
        }
        1 {
            # Verify E1 license, signin blocked on all users, remove all licenses
            Clear-Variable AssignedLicense -ErrorAction SilentlyContinue
            Clear-Variable E1Assigned -ErrorAction SilentlyContinue
            Clear-Variable E3Assigned -ErrorAction SilentlyContinue
            Clear-Variable TempLicenses -ErrorAction SilentlyContinue

            foreach($AssignedLicense in $UserInfo.AssignedLicenses){
                if($AssignedLicense.SkuID -eq "18181a46-0d4e-45cd-891e-60aabd171b4e"){
                    $E1Assigned = $True
                }
                if($AssignedLicense.SkuID -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"){
                    $E3Assigned = $True
                }
            }
            if(($E1Assigned -eq $True) -and ($E3Assigned -ne $True)){
                $TempLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                $TempLicenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $TempLicenses
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) had an E1 Office 365 license, which has been removed (along with all other licenses).  User data will be lost after 90 days.`r"
            }
            elseif($E3Assigned -eq $True){
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) has an Office 365 E3 License.  Sign-in has been blocked.  Please verify licensing and complete termination manually.`r" -Color "Red"
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) does NOT have an E1 Office 365 License.  Please verify licensing and complete terminaiton manually.`r" - Color "Red"
            }
        }
        2 {
            # Verify E3 License, signin blocked on all users, Add E1, Remove E3, leave other licenses
            Clear-Variable AssignedLicense -ErrorAction SilentlyContinue
            Clear-Variable E1Assigned -ErrorAction SilentlyContinue
            Clear-Variable E3Assigned -ErrorAction SilentlyContinue
            Clear-Variable TempLicenses -ErrorAction SilentlyContinue
            Clear-Variable TempLicense -ErrorAction SilentlyContinue
            Clear-Variable E1License -ErrorAction SilentlyContinue

            foreach($AssignedLicense in $UserInfo.AssignedLicenses){
                if($AssignedLicense.SkuID -eq "18181a46-0d4e-45cd-891e-60aabd171b4e"){
                    $E1Assigned = $True
                }
                if($AssignedLicense.SkuID -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"){
                    $E3Assigned = $True
                }
            }
            
            if(($E3Assigned -eq $True) -and ($E1Assigned -ne $True)){
                $E1License = Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits | Where-Object {$_.SkuID -eq "18181a46-0d4e-45cd-891e-60aabd171b4e"}
                while($E1License.Enabled - $E1License.ConsumedUnits -lt 1){
                    $E1License = Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits | Where-Object {$_.SkuID -eq "18181a46-0d4e-45cd-891e-60aabd171b4e"}
                    #Provide Pop Up Stating E1 license needed
                    $null = [System.Windows.MessageBox]::Show("There are no available Office 365 E1 Licenses, please get one added to the tenant and hit OK to try again.","License Check","OKCancel","Warning")
                    
                }
                
                #Add E1 and Remove E3
                #E1
                $TempLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                $TempLicense.SkuID = "18181a46-0d4e-45cd-891e-60aabd171b4e"
                $TempLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                $TempLicenses.AddLicenses = $TempLicense
                #E3
                $TempLicenses.RemoveLicenses = "6fd2c87f-b296-42f0-b197-1e91e994b900"
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $TempLicenses
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) had an E3 Office 365 license (E3 Removed, No Other Licenses Removed), An E1 License has been added, Sign-in has been blocked.`r"
            }
            elseif(($E3Assigned -eq $True) -and ($E1Assigned -eq $True)){
                #Remove E3
                $TempLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                $TempLicenses.RemoveLicenses = "6fd2c87f-b296-42f0-b197-1e91e994b900"
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $TempLicenses
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) had an E3 Office 365 license and an E1 license (E3 Removed, No Other Licenses Removed), Sign-in has been blocked.`r"
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) does NOT have an E3 Office 365 License.  Sign-in has been blocked.  Please verify licensing and complete termination manually.`r" -Color "Red"
            }
        }
        3 {
            # Verify E3 license, signin blocked on all users, remove all licenses
            Clear-Variable AssignedLicense -ErrorAction SilentlyContinue
            Clear-Variable E1Assigned -ErrorAction SilentlyContinue
            Clear-Variable E3Assigned -ErrorAction SilentlyContinue
            Clear-Variable TempLicenses -ErrorAction SilentlyContinue

            foreach($AssignedLicense in $UserInfo.AssignedLicenses){
                if($AssignedLicense.SkuID -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"){
                    $E3Assigned = $True
                }
            }

            if($E3Assigned -eq $True){
                $TempLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                $TempLicenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $TempLicenses
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) had an E3 Office 365 license, which has been removed (along with all other licenses).  User data will be lost after 90 days.`r"
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) does NOT have an E3 Office 365 License.  Please verify licensing and complete termination manually.`r" -Color "Red"
            }
        }
        4 {
            # Verify E3 license, signin blocked on all users, converted to Shared prior to Azure AD Sync, remove all licenses
            Clear-Variable AssignedLicense -ErrorAction SilentlyContinue
            Clear-Variable E1Assigned -ErrorAction SilentlyContinue
            Clear-Variable E3Assigned -ErrorAction SilentlyContinue
            Clear-Variable TempLicenses -ErrorAction SilentlyContinue

            foreach($AssignedLicense in $UserInfo.AssignedLicenses){
                if($AssignedLicense.SkuID -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"){
                    $E3Assigned = $True
                }
            }
            if($E3Assigned -eq $True){
                #Remove all licenses
                $TempLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                $TempLicenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $TempLicenses
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser.Name) had an E3 Office 365 license, has been converted to a Shared Mailbox and all licenses have been removed.`r"
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser.Name) does NOT have an E3 license, please verify licensing and complete termination manually.`r" -Color "Red"
            }
        }
        Default {
            Write-RichtextBox -TextBox $TerminationRichTextBox -Text "I don't know how you did it, but you didn't select anything in the dropdown, please confirm/complete termination manually.`r" -Color "Red"
        }
    }
    Remove-Variable termeduser -Scope Global
    Remove-Variable Manager -Scope Global
})

$null = $UserForm.ShowDialog()