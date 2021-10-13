#Request account information, set mailnickname, hide from GAL, and disable account
$user = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName,SamAccountName,DistinguishedName | sort-Object Name | Out-Gridview -OutputMode Single
set-aduser -Identity $user.distinguishedname -replace @{msExchHideFromAddressLists=$True;mailnickname=$user.SamAccountName}
Disable-ADAccount -Identity $user.DistinguishedName -Confirm:$False

#Set Variables
$NewPassword = (Read-Host -Prompt "Provide New Password" -AsSecureString)
$OOO = Read-Host -Prompt "Does an Out Of Office Message Need to be set (y/n)?"
$shared = Read-Host -Prompt "Does this need to be a shared mailbox (y/n)?"
$Hold = Read-Host -Prompt "Is a Litigation Hold Needed (y/n)?"

#Connect to Exchange
Connect-ExchangeOnline

#Connect to AzureAD
Connect-AzureAD

#Block user's sign-in
Set-AzureADUser -ObjectId $user.userprincipalname -AccountEnabled $False

#Reset user's password
Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword $NewPassword -Reset
Write-Host "Password Reset"

#Set OOO
if ( $OOO -match 'y')
{
    Write-Host "Select the user's manager."
    $Manager = Get-ADUser -Filter "Enabled -eq 'True'" | Select-Object Name,UserPrincipalName | sort-Object Name | Out-Gridview -OutputMode Single | Select-Object Name,UserPrincipalName
    $OOOmessage = @"
        $($user.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.

        Please contact $Manager and your emails will be delivered to the appropriate department.

        Thank you
"@
    Set-MailboxAutoReplyConfiguration -Identity $user.userprincipalname -ExternalMessage $OOOmessage -InternalMessage $OOOmessage -AutoReplyState Enabled
    Write-Host "Out of Office Set"
}

#Set Shared Mailbox
if ( $shared -match 'y')
{
    Set-Mailbox $user.userprincipalname -Type Shared
}

#Remove AD/365 Group Memberships
$ADGroups = (Get-ADUser $user.SamAccountName -Properties memberof).memberof
$ADGroups | ForEach-Object {remove-adgroupmember -identity $_ -member $user.SamAccountName}

$DistributionGroups= Get-DistributionGroup | Where-Object { (Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySmtpAddress}) -contains "$User"}
foreach($dg in $DistributionGroups)
{
    Remove-DistributionGroupMember $dg.name -Member $user.samaccountname -Confirm:$false
}
Write-Host "Removed Group Memberships"


#Set Litigation Hold
if ( $hold -match 'y')
{
    Set-Mailbox $email -LitigationHoldEnabled $true
    Write-Host "Litigation Hold Set"
}