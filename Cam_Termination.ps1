#Request account information, set mailnickname, hide from GAL, and disable account
$user = Get-ADUser -Filter "Enabled -eq 'True' -and userprincipalname -like '*@crisisassistance.org'" | Select-Object Name,UserPrincipalName,SamAccountName,DistinguishedName | sort-Object Name | Out-Gridview -OutputMode Single
$username = $user.DistinguishedName
set-aduser -Identity $username -replace @{msExchHideFromAddressLists=$True;mailnickname=$user.SamAccountName}
Disable-ADAccount -Identity $user.DistinguishedName -Confirm:$False

#Set Variables
$NewPassword = (Read-Host -Prompt "Provide New Password" -AsSecureString)
$OOO = Read-Host -Prompt "Does an Out Of Office Message Need to be set (y/n)?"
$shared = Read-Host -Prompt "Does this need to be a shared mailbox (y/n)?"
$Hold = Read-Host -Prompt "Is a Litigation Hold Needed (y/n)?"

#Connect to Exchange
Connect-ExchangeOnline

#Reset user's password
Set-ADAccountPassword -Identity $user -NewPassword $NewPassword -Reset
Write-Host "Password Reset"

#Set OOO
if ( $OOO -match 'y')
{
    $Manager = Read-Host "Enter manager's username"
    $OOOmessage = @"
$($user.Name) is no longer with Crisis Assistance Ministry, and this email is not monitored.

Please contact $($Manager)@crisisassistance.org and your emails will be delivered to the appropriate department.

Thank you
"@
    Set-MailboxAutoReplyConfiguration -Identity $user -ExternalMessage $OOOmessage -InternalMessage $OOOmessage -AutoReplyState Enabled
    Write-Host "Out of Office Set"
}

#Set Shared Mailbox
if ( $shared -match 'y')
{
    Set-Mailbox $user -Type Shared
}

#Remove AD/365 Group Memberships
$ADGroups = (Get-ADUser $user -Properties memberof).memberof
$ADGroups | foreach {remove-adgroupmember -identity $_ -member $user}

$DistributionGroups= Get-DistributionGroup | where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains "$User"}
$DistributionGroups | Select-Object DisplayName,ExchangeObjectID | Out-File $user-DGs.txt
foreach ($dg in $DistributionGroups)
{
    Remove-DistributionGroupMember $dg.name -Member $User -Confirm:$false
}
Write-Host "Removed Group Memberships"


#Set Litigation Hold
$Hold = Read-Host -Prompt "Is a Litigation Hold Needed (y/n)?"
if ( $hold -match 'y')
{
    Set-Mailbox $user -LitigationHoldEnabled $true
    Write-Host "Litigation Hold Set"
}