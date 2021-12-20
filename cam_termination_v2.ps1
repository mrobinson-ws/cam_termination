Add-Type -AssemblyName System.Web
Add-Type -AssemblyName PresentationFramework

# Test For Modules
if(-not(Get-Module ExchangeOnlineManagement -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement - view http://worksmart.link/7l for details')
    Exit
}

if(-not(Get-Module AzureAD -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install AzureAD - view http://worksmart.link/7l for details')
    Exit
}

### Start XAML and Reader to use WPF, as well as declare variables for use
[xml]$xaml = @"
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="CAM Termination" Height="350" Width="525" Topmost="True">

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

# Friendly Name Lookup Table
$SkuToFriendly = @{
    "9c7bff7a-3715-4da7-88d3-07f57f8d0fb6" = "Dynamics 365 Sales Professional"
    "6070a4c8-34c6-4937-8dfb-39bbc6397a60" = "Meeting Room"
    "8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b" = "Rights Management ADHOC"
    "c42b9cae-ea4f-4ab7-9717-81576235ccac" = "DevPack E5 (No Windows or Audio)"
    "8f0c5670-4e56-4892-b06d-91c085d7004f" = "APP CONNECT IW"
    "0c266dff-15dd-4b49-8397-2bb16070ed52" = "Microsoft 365 Audio Conferencing"
    "2b9c8e7c-319c-43a2-a2a0-48c5c6161de7" = "AZURE ACTIVE DIRECTORY BASIC"
    "078d2b04-f1bd-4111-bbd4-b4b1b354cef4" = "AZURE ACTIVE DIRECTORY PREMIUM P1"
    "84a661c4-e949-4bd2-a560-ed7766fcaf2b" = "AZURE ACTIVE DIRECTORY PREMIUM P2"
    "c52ea49f-fe5d-4e95-93ba-1de91d380f89" = "AZURE INFORMATION PROTECTION PLAN 1"
    "295a8eb0-f78d-45c7-8b5b-1eed5ed02dff" = "COMMON AREA PHONE"
    "47794cd0-f0e5-45c5-9033-2eb6b5fc84e0" = "COMMUNICATIONS CREDITS"
    "ea126fc5-a19e-42e2-a731-da9d437bffcf" = "DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION"
    "749742bf-0d37-4158-a120-33567104deeb" = "DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION"
    "cc13a803-544e-4464-b4e4-6d6169a138fa" = "DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION"
    "8edc2cf8-6438-4fa9-b6e3-aa1660c640cc" = "DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
    "1e1a282c-9c54-43a2-9310-98ef728faace" = "DYNAMICS 365 FOR SALES ENTERPRISE EDITION"
    "f2e48cb3-9da0-42cd-8464-4a54ce198ad0" = "DYNAMICS 365 FOR SUPPLY CHAIN MANAGEMENT"
    "8e7a3d30-d97d-43ab-837c-d7701cef83dc" = "DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION"
    "338148b6-1b11-4102-afb9-f92b6cdc0f8d" = "DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS"
    "b56e7ccc-d5c7-421f-a23b-5c18bdbad7c0" = "DYNAMICS 365 TALENT: ONBOARD"
    "7ac9fe77-66b7-4e5e-9e46-10eed1cff547" = "DYNAMICS 365 TEAM MEMBERS"
    "ccba3cfe-71ef-423a-bd87-b6df3dce59a9" = "DYNAMICS 365 UNF OPS PLAN ENT EDITION"
    "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "ENTERPRISE MOBILITY + SECURITY E3"
    "b05e124f-c7cc-45a0-a6aa-8cf78c946968" = "ENTERPRISE MOBILITY + SECURITY E5"
    "4b9405b0-7788-4568-add1-99614e613b69" = "EXCHANGE ONLINE (PLAN 1)"
    "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "EXCHANGE ONLINE (PLAN 2)"
    "ee02fd1b-340e-4a4b-b355-4a514e4c8943" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE"
    "90b5e015-709a-4b8b-b08e-3200f994494c" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
    "7fc0182e-d107-4556-8329-7caaa511197b" = "EXCHANGE ONLINE ESSENTIALS (ExO P1 BASED)"
    "e8f81a67-bd96-4074-b108-cf193eb9433b" = "EXCHANGE ONLINE ESSENTIALS"
    "80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82" = "EXCHANGE ONLINE KIOSK"
    "cb0a98a8-11bc-494c-83d9-c1b1ac65327e" = "EXCHANGE ONLINE POP"
    "061f9ace-7d42-4136-88ac-31dc755f143f" = "INTUNE"
    "b17653a4-2443-4e8c-a550-18249dda78bb" = "Microsoft 365 A1"
    "4b590615-0888-425a-a965-b3bf7789848d" = "MICROSOFT 365 A3 FOR FACULTY"
    "7cfd9a2b-e110-4c39-bf20-c6a3f36a3121" = "MICROSOFT 365 A3 FOR STUDENTS"
    "e97c048c-37a4-45fb-ab50-922fbf07a370" = "MICROSOFT 365 A5 FOR FACULTY"
    "46c119d4-0379-4a9d-85e4-97c66d3f909e" = "MICROSOFT 365 A5 FOR STUDENTS"
    "cdd28e44-67e3-425e-be4c-737fab2899d3" = "MICROSOFT 365 APPS FOR BUSINESS"
    "b214fe43-f5a3-4703-beeb-fa97188220fc" = "MICROSOFT 365 APPS FOR BUSINESS"
    "c2273bd0-dff7-4215-9ef5-2c7bcfb06425" = "MICROSOFT 365 APPS FOR ENTERPRISE"
    "2d3091c7-0712-488b-b3d8-6b97bde6a1f5" = "MICROSOFT 365 AUDIO CONFERENCING FOR GCC"
    "3b555118-da6a-4418-894f-7df1e2096870" = "MICROSOFT 365 BUSINESS BASIC"
    "dab7782a-93b1-4074-8bb1-0e61318bea0b" = "MICROSOFT 365 BUSINESS BASIC"
    "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "MICROSOFT 365 BUSINESS STANDARD"
    "ac5cef5d-921b-4f97-9ef3-c99076e5470f" = "MICROSOFT 365 BUSINESS STANDARD - PREPAID LEGACY"
    "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "MICROSOFT 365 BUSINESS PREMIUM"
    "11dee6af-eca8-419f-8061-6864517c1875" = "MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)"
    "05e9a617-0261-4cee-bb44-138d3ef5d965" = "MICROSOFT 365 E3"
    "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
    "d61d61cc-f992-433f-a577-5bd016037eeb" = "Microsoft 365 E3_USGOV_DOD"
    "ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658" = "Microsoft 365 E3_USGOV_GCCHIGH"
    "184efa21-98c3-4e5d-95ab-d07053a96e67" = "Microsoft 365 E5 Compliance"
    "26124093-3d78-432b-b5dc-48bf992543d5" = "Microsoft 365 E5 Security"
    "44ac31e7-2999-4304-ad94-c948886741d4" = "Microsoft 365 E5 Security for EMS E5"
    "44575883-256e-4a79-9da4-ebe9acabe2b2" = "Microsoft 365 F1"
    "66b55226-6b4f-492c-910c-a3b7a3c9d993" = "Microsoft 365 F3"
    "f30db892-07e9-47e9-837c-80727f46fd3d" = "MICROSOFT FLOW FREE"
    "e823ca47-49c4-46b3-b38d-ca11d5abe3d2" = "MICROSOFT 365 G3 GCC"
    "e43b5b99-8dfb-405f-9987-dc307f34bcbd" = "MICROSOFT 365 PHONE SYSTEM"
    "d01d9287-694b-44f3-bcc5-ada78c8d953e" = "MICROSOFT 365 PHONE SYSTEM FOR DOD"
    "d979703c-028d-4de5-acbf-7955566b69b9" = "MICROSOFT 365 PHONE SYSTEM FOR FACULTY"
    "a460366a-ade7-4791-b581-9fbff1bdaa85" = "MICROSOFT 365 PHONE SYSTEM FOR GCC"
    "7035277a-5e49-4abc-a24f-0ec49c501bb5" = "MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH"
    "aa6791d3-bb09-4bc2-afed-c30c3fe26032" = "MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS"
    "1f338bbc-767e-4a1e-a2d4-b73207cc5b93" = "MICROSOFT 365 PHONE SYSTEM FOR STUDENTS"
    "ffaf2d68-1c95-4eb3-9ddd-59b81fba0f61" = "MICROSOFT 365 PHONE SYSTEM FOR TELSTRA"
    "b0e7de67-e503-4934-b729-53d595ba5cd1" = "MICROSOFT 365 PHONE SYSTEM_USGOV_DOD"
    "985fcb26-7b94-475b-b512-89356697be71" = "MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH"
    "440eaaa8-b3e0-484b-a8be-62870b9ba70a" = "MICROSOFT 365 PHONE SYSTEM - VIRTUAL USER"
    "2347355b-4e81-41a4-9c22-55057a399791" = "MICROSOFT 365 SECURITY AND COMPLIANCE FOR FLW"
    "726a0894-2c77-4d65-99da-9775ef05aad1" = "MICROSOFT BUSINESS CENTER"
    "111046dd-295b-4d6d-9724-d52ac90bd1f2" = "MICROSOFT DEFENDER FOR ENDPOINT"
    "906af65a-2970-46d5-9b58-4e9aa50f0657" = "MICROSOFT DYNAMICS CRM ONLINE BASIC"
    "d17b27af-3f49-4822-99f9-56a661538792" = "MICROSOFT DYNAMICS CRM ONLINE"
    "ba9a34de-4489-469d-879c-0f0f145321cd" = "MS IMAGINE ACADEMY"
    "2c21e77a-e0d6-4570-b38a-7ff2dc17d2ca" = "MICROSOFT INTUNE DEVICE FOR GOVERNMENT"
    "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "MICROSOFT POWER APPS PLAN 2 TRIAL"
    "e6025b08-2fa5-4313-bd0a-7e5ffca32958" = "MICROSOFT INTUNE SMB"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "MICROSOFT STREAM"
    "16ddbbfc-09ea-4de2-b1d7-312db6112d70" = "MICROSOFT TEAM (FREE)"
    "710779e8-3d4a-4c88-adb9-386c958d1fdf" = "MICROSOFT TEAMS EXPLORATORY"
    "a4585165-0533-458a-97e3-c400570268c4" = "Office 365 A5 for faculty"
    "ee656612-49fa-43e5-b67e-cb1fdf7699df" = "Office 365 A5 for students"
    "1b1b1f7a-8355-43b6-829f-336cfccb744c" = "Office 365 Advanced Compliance"
    "4ef96642-f096-40de-a3e9-d83fb2f90211" = "Microsoft Defender for Office 365 (Plan 1)"
    "18181a46-0d4e-45cd-891e-60aabd171b4e" = "OFFICE 365 E1"
    "6634e0ce-1a9f-428c-a498-f84ec7b8aa2e" = "OFFICE 365 E2"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "OFFICE 365 E3"
    "189a915c-fe4f-4ffa-bde4-85b9628d07a0" = "OFFICE 365 E3 DEVELOPER"
    "b107e5a3-3e60-4c0d-a184-a7e4395eb44c" = "Office 365 E3_USGOV_DOD"
    "aea38a85-9bd5-4981-aa00-616b411205bf" = "Office 365 E3_USGOV_GCCHIGH"
    "1392051d-0cb9-4b7a-88d5-621fee5e8711" = "OFFICE 365 E4"
    "c7df2760-2c81-4ef7-b578-5b5392b571df" = "OFFICE 365 E5"
    "26d45bd9-adf1-46cd-a9e1-51e9a5524128" = "OFFICE 365 E5 WITHOUT AUDIO CONFERENCING"
    "4b585984-651b-448a-9e53-3b10f069cf7f" = "OFFICE 365 F3"
    "535a3a29-c5f0-42fe-8215-d3b9e1f38c4a" = "OFFICE 365 G3 GCC"
    "04a7fb0d-32e0-4241-b4f5-3f7618cd1162" = "OFFICE 365 MIDSIZE BUSINESS"
    "bd09678e-b83c-4d3f-aaba-3dad4abd128b" = "OFFICE 365 SMALL BUSINESS"
    "fc14ec4a-4169-49a4-a51e-2c852931814b" = "OFFICE 365 SMALL BUSINESS PREMIUM"
    "e6778190-713e-4e4f-9119-8b8238de25df" = "ONEDRIVE FOR BUSINESS (PLAN 1)"
    "ed01faf2-1d88-4947-ae91-45ca18703a96" = "ONEDRIVE FOR BUSINESS (PLAN 2)"
    "87bbbc60-4754-4998-8c88-227dca264858" = "POWERAPPS AND LOGIC FLOWS"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "POWER BI (FREE)"
    "45bc2c81-6072-436a-9b0b-3b12eefbc402" = "POWER BI FOR OFFICE 365 ADD-ON"
    "f8a1db68-be16-40ed-86d5-cb42ce701560" = "POWER BI PRO"
    "a10d5e58-74da-4312-95c8-76be4e5b75a0" = "PROJECT FOR OFFICE 365"
    "776df282-9fc0-4862-99e2-70e561b9909e" = "PROJECT ONLINE ESSENTIALS"
    "09015f9f-377f-4538-bbb5-f75ceb09358a" = "PROJECT ONLINE PREMIUM"
    "2db84718-652c-47a7-860c-f10d8abbdae3" = "PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT"
    "53818b1b-4a27-454b-8896-0dba576410e6" = "PROJECT ONLINE PROFESSIONAL"
    "f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c" = "PROJECT ONLINE WITH PROJECT FOR OFFICE 365"
    "beb6439c-caad-48d3-bf46-0c82871e12be" = "PROJECT PLAN 1"
    "1fc08a02-8b3d-43b9-831e-f76859e04e1a" = "SHAREPOINT ONLINE (PLAN 1)"
    "a9732ec9-17d9-494c-a51c-d6b45b384dcb" = "SHAREPOINT ONLINE (PLAN 2)"
    "b8b749f8-a4ef-4887-9539-c95b1eaa5db7" = "SKYPE FOR BUSINESS ONLINE (PLAN 1)"
    "d42c793f-6c78-4f43-92ca-e8f6a02b035f" = "SKYPE FOR BUSINESS ONLINE (PLAN 2)"
    "d3b4fe1f-9992-4930-8acb-ca6ec609365e" = "SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING"
    "0dab259f-bf13-4952-b7f8-7db8f131b28d" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING"
    "54a152dc-90de-4996-93d2-bc47e670fc06" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)"
    "4016f256-b063-4864-816e-d818aad600c9" = "TOPIC EXPERIENCES"
    "de3312e1-c7b0-46e6-a7c3-a515ff90bc86" = "TELSTRA CALLING FOR O365"
    "4b244418-9658-4451-a2b8-b5e2b364e9bd" = "VISIO ONLINE PLAN 1"
    "c5928f49-12ba-48f7-ada3-0d743a3601d5" = "VISIO ONLINE PLAN 2"
    "4ae99959-6b0f-43b0-b1ce-68146001bdba" = "VISIO PLAN 2 FOR GCC"
    "cb10e6cd-9da4-4992-867b-67546b1db821" = "WINDOWS 10 ENTERPRISE E3"
    "6a0f6da5-0b87-4190-a6ae-9bb5a2b9546a" = "WINDOWS 10 ENTERPRISE E3"
    "488ba24a-39a9-4473-8ee5-19291e71b002" = "Windows 10 Enterprise E5"
    "6470687e-a428-4b7a-bef2-8a291ad947c9" = "WINDOWS STORE FOR BUSINESS"
    "aa2695c9-8d59-4800-9dc8-12e01f1735af" = "NONPROFIT_PORTAL"
}

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
    $Global:termeduser = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName,SamAccountName,DistinguishedName | sort-Object Name | Out-Gridview -OutputMode Single
    $UserTextbox.Text = $Global:termeduser.Name
    $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.Name) and your emails will be delivered to the appropriate department.
Thank you.
"@
})

#Randomly generate a 16 character password with 8 being non-alphanumeric
$GeneratePasswordButton.Add_Click({
    $PasswordTextBox.Text = [System.Web.Security.Membership]::GeneratePassword(16,8)
})

#Select Manager
$ManagerButton.Add_Click({
    $Global:Manager = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName | sort-Object Name | Out-Gridview -OutputMode Single
    $ManagerTextBox.Text = $Global:Manager.Name
    $OOOTextBox.Text = @"
$($Global:termeduser.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.
Please contact $($Global:Manager.Name) and your emails will be delivered to the appropriate department.
Thank you.
"@
})

#Terminate the user with selected options
$TerminateGoButton.Add_Click({
    #Set Mail Nickname, Hide from GAL, and Disable AD User Account
    Set-ADUser -Identity $Global:termeduser.distinguishedname -replace @{msExchHideFromAddressLists=$True;mailnickname=$Global:termeduser.SamAccountName}
    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "Hid user from GAL and set Mail Nickname`r"
    
    $SecurePassword = ConvertTo-SecureString -String $PasswordTextBox.Text -AsPlainText -Force
    Set-ADAccountPassword -Identity $Global:termeduser.SamAccountName -NewPassword $SecurePassword -Reset
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
        Set-MailboxAutoReplyConfiguration -Identity $Global:termeduser.UserPrincipalName -ExternalMessage $OOOTextbox.Text -InternalMessage $OOOTextbox.Text -AutoReplyState Enabled
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

    #Manage Licenses as needed
    $Licenses = Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits

    $UserInfo = Get-AzureADUser -ObjectId $Global:termeduser.UserPrincipalName

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
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) has an Office 365 E3 License.  Sign-in has been blocked.  Please verify licensing manually.`r"
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) does NOT have an E1 Office 365 License.  Please verify licensing manually.`r"
            }
        }
        2 {
            # Verify E3 License, signin blocked on all users, Add E1, Remove E3, leave other licenses
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
            
            if(($E3Assigned -eq $True) -and ($E1Assigned -eq $True)){
                $E1License = $Licenses | Where-Object {$_.SkuID -eq "18181a46-0d4e-45cd-891e-60aabd171b4e"}
                if($E1License.Enabled- $E1License.ConsumedUnits -ge 0){
                    #Remove E3
                    Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) had an E3 Office 365 licenseand an E1 license (E3 Removed, No Other Licenses Removed), Sign-in has been blocked.`r"    
                }
            }
            elseif(($E3Assigned -eq $True) -and ($E1Assigned -ne $True)){
                #Assign E1, Remove E3
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) had an E3 Office 365 license, now has an E1 license (no other licenses removed), Sign-in has been blocked.`r"    
            }
            else{
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) does NOT have an E3 Office 365 License.  Sign-in has been blocked.  Please verify licensing manually.`r"
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
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) does NOT have an E3 Office 365 License.  Please verify licensing manually.`r"
            }
        }
        4 {
            # Verify E3 license, signin blocked on all users, converted to Shared prior to Azure AD Sync, remove all licenses
            foreach($AssignedLicense in $UserInfo.AssignedLicenses){
                if($AssignedLicense.SkuID -eq "6fd2c87f-b296-42f0-b197-1e91e994b900"){
                    $E3Assigned = $True
                }
            }
            if($E3Assigned -eq $True){
                #Remove all licenses
                Write-RichtextBox -TextBox $TerminationRichTextBox -Text "$($Global:termeduser) had an E3 Office 365, has been converted to a Shared Mailbox and all licenses have been removed.`r"
            }
        }
        Default {
            Write-RichtextBox -TextBox $TerminationRichTextBox -Text "I don't know how you did it, but you didn't select anything in the dropdown, please confirm termination manually.`r" -Color "Red"
        }
    }
    Clear-Variable $Global:termeduser
    Clear-Variable $Global:Manager
})

$null = $UserForm.ShowDialog()