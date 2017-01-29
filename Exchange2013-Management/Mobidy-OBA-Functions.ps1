# Exchange 2013 Administrative Tasks

. "C:\Users\Administrator\Desktop\Scripts\Functions.ps1"

[boolean]$global:ExitSession = $false

function LoadMenu() {
	[int]$Menu1 = 0
	[int]$Menu2 = 0
	[boolean]$ValidSelection = $false
	while ($Menu1 -lt 1 -or $Menu1 -gt 5) {
		CLS
		# Root menu options}
		Write-Host "`n`tExchange 2013 Commands" -f green
		Write-Host "`t`tPlease select a menu option" -f Cyan
		Write-Host "`t`t`t1. User mailbox tasks" -f Cyan
		Write-Host "`t`t`t2. Distribution group tasks" -f Cyan
		Write-Host "`t`t`t3. Tenant statistics" -f Cyan
		Write-Host "`t`t`t4. Database/Administrators options" -f Cyan
		Write-Host "`t`t`t5. Quit" -f Cyan
		[int]$Menu1 = Read-Host "`t`tEnter menu option number"
		if ($Menu1 -lt 1 -or $Menu1 -gt 5) {
			Write-Host "`tPlease select one of the options available." -f red; Start-Sleep -Seconds 1
		}
	}

	# Second level menu options
	Switch ($Menu1) {
		1 {
			while ($Menu2 -lt 1 -or $Menu2 -gt 9) {
				CLS
				# User mailbox tasks menu
				Write-Host "`n`tUser mailbox tasks" -f green
				Write-Host "`t`tPlease select a menu option" -f Cyan
				Write-Host "`t`t`t1. Create new email account" -f Cyan
				Write-Host "`t`t`t2. Change primary email address" -f Cyan
				Write-Host "`t`t`t3. Hide mailbox from address list" -f Cyan
                Write-Host "`t`t`t4. Create mail contact" -f Cyan
				Write-Host "`t`t`t5. Assign FullAccess permissions" -f Cyan
				Write-Host "`t`t`t6. Assign SendAs permissions" -f Cyan
				Write-Host "`t`t`t7. Delete mailbox" -f Cyan
				Write-Host "`t`t`t8. Disable mailbox (reset pass, remove AS, remove from groups)" -f Cyan
				Write-Host "`t`t`t9. Go to main menu" -f Cyan
				[int]$Menu2 = Read-Host "`t`tEnter menu options"
			}
			if ($Menu2 -lt 1 -or $Menu2 -gt 9) {
				Write-Host "`Please select one of the options available." -f red; Start-Sleep -Seconds 1
			}
			Switch ($Menu2) {
				1 { Write-Host "`n`tCreate new email account" -f yellow; CreateEmailAccount }
				2 { Write-Host "`n`tChange primary email address" -f yellow; ChangePrimaryEmail }
				3 { Write-Host "`n`tCreate mail contact" -f yellow; CreateMailContact }
				4 { Write-Host "`n`tHide Mailbox from address list" -f yellow; HideMailboxFromAddressList }
				5 { Write-Host "`n`tAssign FullAccess permissions" -f yellow; AssignFullAccess }
				6 { Write-Host "`n`tAssign SendAs permissions" -f yellow; AssignSendAs }
				7 { Write-Host "`n`tDelete mailbox" -f yellow; DeleteMailbox }
				8 { Write-Host "`n`tDisable mailbox" -f yellow; DisableMailbox }
				default { Write-Host "`n`tQuit to main menu" -f yellow; break }
			}
		}
		2 {
			while ($Menu2 -lt 1 -or $menu2 -gt 4) {
				CLS
				# Distribution group tasks menu
				Write-Host "`n`tDistribution group tasks" -f green
				Write-Host "`t`tPlease select a menu option" -f Cyan
				Write-Host "`t`t`t1. Create new distribution group" -f Cyan
				Write-Host "`t`t`t2. Change primary email address" -f Cyan
				Write-Host "`t`t`t3. Something else here" -f Cyan
				Write-Host "`t`t`t4. Go to menu menu" -f Cyan
				[int]$Menu2 = Read-Host "`t`tEnter menu option"
			}
			if ($Menu2 -lt 1 -or $Menu2 -gt 4) {
				Write-Host "`tPlease Select one of the options available." -f red; Start-Sleep -Seconds 1
			}
			Switch ($Menu2) {
				1 { Write-Host "`n`tCreate new distribution group" -f yellow; CreateDistributionGroup }
				2 { Write-Host "`n`tChange primary group email address" -f yellow; ChangePrimaryGroupEmail }
				3 { Write-Host "`n`tSomething else here" -f yellow }
				default { Write-Host "`n`tQuit to main menu" -f yellow; break }
			}
		}
		3 {
			while ($Menu2 -lt 1 -or $menu2 -gt 3) {
				CLS
				# Tenant statistics menu
				Write-Host "`n`tTenant statistics tasks" -f green
				Write-Host "`t`tPlease select a menu option" -f Cyan
				Write-Host "`t`t`t1. Mailbox list with item count and mailbox size" -f Cyan
				Write-Host "`t`t`t2. Mailbox list with aliases" -f Cyan
				Write-Host "`t`t`t3. Go to menu menu" -f Cyan
				[int]$Menu2 = Read-Host "`t`tEnter menu option"
			}
			if ($Menu2 -lt 1 -or $Menu2 -gt 3) {
				Write-Host "`tPlease Select one of the options available." -f red; Start-Sleep -Seconds 1
			}
			Switch ($Menu2) {
				1 { Write-Host "`n`tMailbox list with item count and mailbox size" -f yellow; MailboxCount1 }
				2 { Write-Host "`n`tMailbox list with aliases" -f yellow; MailboxCount2 }
				default { Write-Host "`n`tQuit to main menu" -f yellow; break }
			}
		}
	4 {
			while ($Menu2 -lt 1 -or $menu2 -gt 3) {
				CLS
				# Database/Administrative tasks menu
				Write-Host "`n`tDatabase/Administrative tasks" -f green
				Write-Host "`t`tPlease select a menu option" -f Cyan
				Write-Host "`t`t`t1. Bring DAG member down for maintenance" -f Cyan
				Write-Host "`t`t`t2. Bring DAG member up after maintenance" -f Cyan
				Write-Host "`t`t`t3. Go to menu menu" -f Cyan
				[int]$Menu2 = Read-Host "`t`tEnter menu option"
			}
			if ($Menu2 -lt 1 -or $Menu2 -gt 3) {
				Write-Host "`tPlease Select one of the options available." -f red; Start-Sleep -Seconds 1
			}
			Switch ($Menu2) {
				1 { Write-Host "`n`tBringing DAG member down for maintenance" -f yellow; DownDAG }
				2 { Write-Host "`n`tBringing DAG member up after maintenance" -f yellow; UpDAG }
				default { Write-Host "`n`tQuit to main menu" -f yellow; break }
			}
		}
		default { $global:ExitSession = $true; break }	
	}
}

LoadMenu
if ($ExitSession) {
	exit-pssession
}
else {
	C:\Users\Administrator\Desktop\Scripts\Modify-OBA-Functions.ps1
}