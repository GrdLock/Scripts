# Obacks Exchange 2013 Functions
Import-Module ActiveDirectory

function VerifyTenant([ref]$tenant, [ref]$tenantou) {
	DO {
		$tenant.Value = Read-Host "Enter tenant"
		[string] $verifytenant = 'OU=' + $tenant.Value + ',OU=Hosted,DC=ROOTDOMAIN,DC=com'
		$ou_exists = [adsi]::Exists("LDAP://$verifytenant")
		If (-not $ou_exists)
		{ Write-Host 'Tenant' $tenant.Value 'does not exist, try again.' -f red }
		else
		{ Write-Host 'Located tenant' $tenant.Value'.' -f green; $tenantou.Value = ("obacks.com/Hosted/" + $tenant.Value) }
	} While (-not $ou_exists)
}

function UpdateAllAddressLists([ref]$tenant) {
	Write-Host 'Updating AddressList:' $tenant.Value '- All Rooms' -f yellow
	Update-AddressList -Identity ($tenant.Value + " - All Rooms")
	Write-Host 'Updating AddressList:' $tenant.Value '- All Users' -f yellow
	Update-AddressList -Identity ($tenant.Value + " - All Users")
	Write-Host 'Updating AddressList:' $tenant.Value '- All Contacts' -f yellow
	Update-AddressList -Identity ($tenant.Value + " - All Contacts")
	Write-Host 'Updating AddressList:' $tenant.Value '- All Groups' -f yellow
	Update-AddressList -Identity ($tenant.Value + " - All Groups")
	Write-Host 'Updating OfflineAddressBook:' $tenant.Value -f yellow
	Update-OfflineAddressBook -Identity $tenant.Value
	Write-Host 'Updating GlobalAddressList:' $tenant.Value '- GAL' -f yellow
	Update-GlobalAddressList -Identity ($tenant.Value + " - GAL")
}

function CreateEmailAccount() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	DO {
	$first = Read-Host "Enter first name"
	$last = Read-Host "Enter last name"
	$alias = Read-Host "Enter alias/samaccountname"
	$domain = Read-Host "Enter domain (no @ symbol)"
	$credentials = Read-Host "Enter password" -AsSecureString
	$tenantalias = $tenant -replace '\s',''
	$tenantalias2 = $alias + $tenantalias
	if ($tenantalias2.length -ge 21) { $tenantalias2 = $tenantalias2.substring(0,20) }
	New-Mailbox -Name ($first + " " + $last) -Alias $tenantalias2 -OrganizationalUnit ("domain.com/Hosted/" + $tenant) -UserPrincipalName ($alias + "@" + $domain) -SamAccountName $tenantalias2 -FirstName $first -LastName $last -Password $credentials -ResetPasswordOnNextLogon $false -AddressBookPolicy $tenant
	Start-Sleep -s 8
	Set-Mailbox ($alias + "@" + $domain) -CustomAttribute1 $tenant
	Set-Mailbox ($alias + "@" + $domain) -ProhibitSendReceiveQuota unlimited -ProhibitSendQuota unlimited -IssueWarningQuota 18GB
	Get-ADUser -Identity $tenantalias2 | Set-ADUser -passwordNeverExpires $true
	Enable-ADAccount -Identity $tenantalias2
	Add-DistributionGroupMember -Identity ($tenant + " Access") -Member ($first + " " + $last)
	$createaccount = Read-Host "Create another account under same tenant? y/n"
	} While ($createaccount -eq "y")
	UpdateAllAddressLists ([ref]$tenant)
	Write-Host "Completed. Press any key to continue..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function ChangePrimaryEmail() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$user = Read-Host "Enter full name of user:"
	$new_email = Read-Host "Enter new email address for PrimarySmtpAddress"
	Get-Mailbox -OrganizationalUnit $tenantou -Identity $user | Set-Mailbox -EmailAddressPolicyEnabled:$false
	Get-Mailbox -OrganizationalUnit $tenantou -Identity $user | Set-Mailbox -PrimarySmtpAddress $new_email
	UpdateAllAddressLists ([ref]$tenant)
	Write-Host "Completed. Press any key to continue..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function HideMailboxFromAddressList() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$user = Read-Host "Enter full name of user"
	Get-Mailbox -OrganizationalUnit $tenant -Identity $user | Set-Mailbox  -HiddenFromAddressListsEnabled $true
	UpdateAllAddressLists ([ref]$tenant)
	Write-Host "Completed. Press any key to continue..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function AssignFullAccess() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$access_to = Read-Host "Mailbox giving access to"
	$user_accessing = Read-Host "User accessing this mailbox"
	Get-Mailbox -OrganizationalUnit $tenantou -Identity $access_to | Add-MailboxPermission -User ($tenantou + "/" + $user_accessing) -AccessRights FullAccess -AutoMapping $true
	Write-Host "Completed. Press any key to continue..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")	
}

function AssignSendAs() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)	
	$sending_as = Read-Host "Mailbox sending as"
	$user_sending = Read-Host "User sending as this mailbox"
	Get-Mailbox -OrganizationalUnit $tenantou -Identity $sending_as | Add-ADPermission -User ($tenantou + "/" + $user_sending) -AccessRights ExtendedRight -ExtendedRights "Send As"
	Write-Host "Completed. Press any key to continue..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function DeleteMailbox() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$user = "User mailbox to delete:"
	Get-Mailbox -OrganizationalUnit $tenantou -Identity $user | Remove-Mailbox
	UpdateAllAddressLists ([ref]$tenant)	
}

function DisableMailbox() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$user = "User mailbox to disable:"
	$password = "Enter new password:"
	$groups = Get-DistributionGroup | Where {(Get-DistributionGroupMember $_ | foreach {$_.Name}) -contains $user}
	foreach ($group in $groups){ Remove-DistributionGroupMember $group -Member $user }
	Set-ADAccountPassword -Reset -NewPassword $password -Identity $user
	Set-CASMailbox -Identity $user -ActiveSyncEnabled $false
	Get-ActiveSyncDevice -Mailbox $user | Remove-ActiveSyncDevice
}

function CreateDistributionGroup() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$groupname = Read-Host 'Enter group name'
	$groupalias = Read-Host 'Enter group alias'
	$groupalias2 = $groupalias.Trim()
	$groupemail = Read-Host 'Enter group email address including domain'
	$groupexternal = Read-Host 'Allow external access to group (y/n)'
	New-DistributionGroup -Name $groupname -Alias $groupalias2 -DisplayName $groupname -OrganizationalUnit ("domain.com/Hosted/" + $tenant) -PrimarySmtpAddress $groupemail -MemberJoinRestriction Closed -MemberDepartRestriction Closed
	Start-Sleep -Seconds 8
	Set-DistributionGroup -Identity $groupalias -CustomAttribute1 $tenant
	Write-Host ' Group' $groupname 'created...' -f yellow
	if ($groupexternal -eq 'y') {
		Set-DistributionGroup -RequireSenderAuthenticationEnabled $false -Identity $groupalias
		Write-Host ' Enabling group for external access...' -f yellow
		Start-Sleep -Seconds 2
	}
	UpdateAllAddressLists ([ref]$tenant)
}

function ChangePrimaryGroupEmail() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	Write-Host "To be implemented"
}

function MailboxCount1() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$count = (Get-Mailbox -OrganizationalUnit $tenantou -resultsize unlimited).count
	write-host ''
	write-host ' Getting stats for ' -f yellow -nonewline; write-host $tenantou -f green -nonewline; write-host ' OU...' -f yellow
	write-host ''
	write-host ' Total Count: ' -f cyan -nonewline; write-host $count -f white
	Get-Mailbox -OrganizationalUnit $tenantou -resultsize unlimited | Get-MailboxStatistics | Sort TotalItemSize  -Descending | ft DisplayName,@{label="Size(GB)";expression={"{0:N2}" -f ($_.TotalItemSize.Value.ToMB() / 1024)}},@{l="Items";e={"{0:N0}" -f $_.ItemCount}},LastLogonTime,Database,@{l="OU";e={$tenantou}} -AutoSize
	Get-Mailbox -OrganizationalUnit $tenantou -resultsize unlimited | Get-MailboxStatistics | %{$_.TotalItemSize.Value.ToMB()} | Measure-Object -sum -average -max | fl @{label="Total Mailboxes (GB)";e={"{0:N2}" -f ($_.sum / 1024) }},@{label="Average Size";e={"{0:N2}" -f ($_.average / 1024) }},@{label="Largest Mailbox";e={"{0:N2}" -f ($_.maximum / 1024) }}
	Write-Host "Completed. Press any key to continue..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function MailboxCount2() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$count = (Get-Mailbox -OrganizationalUnit $tenantou -resultsize unlimited).count
	write-host ''
	write-host ' Getting stats for ' -f yellow -nonewline; write-host $tenantou -f green -nonewline; write-host ' OU...' -f yellow
	write-host ''
	write-host ' Total Count: ' -f cyan -nonewline; write-host $count -f white
	Get-Mailbox -OrganizationalUnit $tenantou -resultsize unlimited | Sort SamAccountName | ft DisplayName,PrimarySMTPAddress,SamAccountName,Alias -AutoSize
	Get-Mailbox -OrganizationalUnit $tenantou -resultsize unlimited | Get-MailboxStatistics | %{$_.TotalItemSize.Value.ToMB()} | Measure-Object -sum -average -max | fl @{label="Total Mailboxes (GB)";e={"{0:N2}" -f ($_.sum / 1024) }},@{label="Average Size";e={"{0:N2}" -f ($_.average / 1024) }},@{label="Largest Mailbox";e={"{0:N2}" -f ($_.maximum / 1024) }}
	Write-Host "Completed. Press any key to continue..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function DownDAG() {
	$server = Read-Host "Enter mailbox server to take offline in DAG"
	Suspend-ClusterNode -Name $server
	Set-MailboxServer $server -DatabaseCopyActivationDisabledAndMoveNow $true
	Set-MailboxServer $server -DatabaseCopyAutoActivationPolicy Blocked
	Set-ServerComponentState $server -Component ServerWideOffline -State InActive -Requester Maintenance
	Write-Host "Mailbox DAG server " + $server + " now offline. Press any key to continue."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function UpDAG() {
	$server = Read-Host "Enter mailbox server to bring online in DAG"
	Resume-ClusterNode -Name $server
	Set-ServerComponentState $server -Component ServerWideOffline -State Active -Requester Maintenance
	Set-MailboxServer $server -DatabaseCopyAutoActivationPolicy Unrestricted
	Set-MailboxServer $server -DatabaseCopyActivationDisabledAndMoveNow $false
	Set-ServerComponentState $server -Component HubTransport -State Active -Requester Maintenance
	Write-Host "Mailbox DAG server " + $server + " now online. Press any key to continue."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function CreateMailContact() {
	New-Variable -Name tenant
	New-Variable -Name tenantou
	VerifyTenant ([ref]$tenant) ([ref]$tenantou)
	$first = Read-Host "Enter first name"
	$last = Read-Host "Enter last name"
	$alias = Read-Host "Enter alias"
	$emailaddress = Read-Host "Enter external email address"
	$tenantalias = $tenant -replace '\s',''
	$tenantalias2 = $alias + $tenantalias
	if ($tenantalias2.length -ge 21) { $tenantalias2 = $tenantalias2.substring(0,20) }
	New-MailContact -Name ($first + " " + $last) -Alias $tenantalias2 -OrganizationalUnit ("obacks.com/Hosted/" + $tenant) -FirstName $first -LastName $last -ExternalEmailAddress $emailaddress
	Start-Sleep -s 8
	Set-MailContact $tenantalias2 -CustomAttribute1 $tenant
	UpdateAllAddressLists ([ref]$tenant)
}