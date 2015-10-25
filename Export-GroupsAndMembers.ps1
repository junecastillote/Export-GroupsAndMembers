Write-Host '=================================================' -ForegroundColor Yellow
Write-Host '>> Distribution Group and Members Export v 1.0 <<' -ForegroundColor Yellow
Write-Host '>>         june.castillote@gmail.com           <<' -ForegroundColor Yellow
Write-Host '=================================================' -ForegroundColor Yellow
#Set Warning and Error Preference
$ErrorActionPreference="SilentlyContinue"
$WarningPreference="SilentlyContinue"

#>>BEGIN------------------------------------------------------------------------------------------
$Today=Get-Date ; Write-Host "$Today : Begin" -ForegroundColor Green
$Today=Get-Date ; Write-Host "$Today : Load Exchange Shell Snapin" -ForegroundColor Green

#>>Add Exchange Snap-in---------------------------------------------------------------------------
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
	{
		try
		{
			Write-Host (Get-Date) ': Add Exchange 2010 Snap-in' -ForegroundColor Yellow
			Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
		}
		catch
		{
			Write-Warning $_.Exception.Message
		}
	}
#>>-----------------------------------------------------------------------------------------------

#>>Variables--------------------------------------------------------------------------------------
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#Make sure that the BackupPath ends with a back slash \
$BackupPath = "C:\Scripts\BackupDir\"
$SenderAddress = "sender@domain.com"
$RecipientAddress = "recipient@domain.com"
$SMTPServer="SMTP.server.here"
#Set this to false if you don't want to receive email notification
$SendReport=$true
$Today=Get-Date
$BackupFile=$BackupPath+"DLBackup_" + $Today.Year + $Today.Month + $Today.Day + "_" + "h" + $Today.Hour + "m" + $Today.Minute + "s" + $Today.Second + ".csv"
#>>-----------------------------------------------------------------------------------------------

#>>Start Export Process---------------------------------------------------------------------------
$Today=Get-Date ; Write-Host "$Today : Export Distribution Groups and Members" -ForegroundColor Yellow
$grouplist = Get-DistributionGroup -ResultSize Unlimited
$Today=Get-Date ; Write-Host "$Today : There are a total of" ($grouplist).count "groups" -ForegroundColor Yellow

$objCollection = @()
foreach ($group in $grouplist)
	{
		$Today=Get-Date ; Write-Host "$Today : Processing Group -" $group.Name -ForegroundColor Yellow
		$group_members = Get-DistributionGroupMember -id $group
		
		foreach ($group_member in $group_members) {
			$temp = "" | Select Name,GroupIdentity,GroupType,SamAccountName,BypassNestedModerationEnabled,ManagedBy,MemberJoinRestriction,MemberDepartRestriction,ExpansionServer,ReportToManagerEnabled,ReportToOriginatorEnabled,SendOofMessageToOriginatorEnabled,AcceptMessagesOnlyFrom,AcceptMessagesOnlyFromDLMembers,AcceptMessagesOnlyFromSendersOrMembers,Alias,OrganizationalUnit,DisplayName,EmailAddresses,GrantSendOnBehalfTo,HiddenFromAddressListsEnabled,MaxSendSize,MaxReceiveSize,ModeratedBy,ModerationEnabled,EmailAddressPolicyEnabled,PrimarySmtpAddress,RecipientType,RecipientTypeDetails,RejectMessagesFrom,RejectMessagesFromDLMembers,RejectMessagesFromSendersOrMembers,RequireSenderAuthenticationEnabled,SimpleDisplayName,SendModerationNotifications,WindowsEmailAddress,MailTip,IsValid,DistinguishedName,MemberName,MemberIdentity
			$Today=Get-Date ; Write-Host "$Today : 		>> " $group_member.Name -ForegroundColor Yellow
			$temp.Name = $group.Name
			$temp.GroupIdentity = $group.Identity
			$temp.GroupType = $group.GroupType
			$temp.SamAccountName = $group.SamAccountName
			$temp.BypassNestedModerationEnabled = $group.BypassNestedModerationEnabled
			$temp.ManagedBy = [string]::join("|", ($group.ManagedBy))
			$temp.MemberJoinRestriction = $group.MemberJoinRestriction
			$temp.MemberDepartRestriction = $group.MemberDepartRestriction
			$temp.ExpansionServer = $group.ExpansionServer
			$temp.ReportToManagerEnabled = $group.ReportToManagerEnabled
			$temp.ReportToOriginatorEnabled = $group.ReportToOriginatorEnabled
			$temp.SendOofMessageToOriginatorEnabled = $group.SendOofMessageToOriginatorEnabled
			$temp.AcceptMessagesOnlyFrom = [string]::join("|", ($group.AcceptMessagesOnlyFrom))
			$temp.AcceptMessagesOnlyFromDLMembers = [string]::join("|", ($group.AcceptMessagesOnlyFromDLMembers))
			$temp.AcceptMessagesOnlyFromSendersOrMembers = [string]::join("|", ($group.AcceptMessagesOnlyFromSendersOrMembers))
			$temp.Alias = $group.Alias
			$temp.OrganizationalUnit = $group.OrganizationalUnit
			$temp.DisplayName = $group.DisplayName
			$temp.EmailAddresses = [string]::join("|", ($group.EmailAddresses))
			$temp.GrantSendOnBehalfTo = [string]::join("|", ($group.GrantSendOnBehalfTo))
			$temp.HiddenFromAddressListsEnabled = $group.HiddenFromAddressListsEnabled
			$temp.MaxSendSize = $group.MaxSendSize
			$temp.MaxReceiveSize = $group.MaxReceiveSize
			$temp.ModeratedBy = [string]::join("|", ($group.ModeratedBy))
			$temp.ModerationEnabled = $group.ModerationEnabled
			$temp.EmailAddressPolicyEnabled = $group.EmailAddressPolicyEnabled
			$temp.PrimarySmtpAddress = $group.PrimarySmtpAddress
			$temp.RecipientType = $group.RecipientType
			$temp.RecipientTypeDetails = $group.RecipientTypeDetails
			$temp.RejectMessagesFrom = [string]::join("|", ($group.RejectMessagesFrom))
			$temp.RejectMessagesFromDLMembers = [string]::join("|", ($group.RejectMessagesFromDLMembers))
			$temp.RejectMessagesFromSendersOrMembers = [string]::join("|", ($group.RejectMessagesFromSendersOrMembers))
			$temp.RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
			$temp.SimpleDisplayName = $group.SimpleDisplayName
			$temp.SendModerationNotifications = $group.SendModerationNotifications
			$temp.WindowsEmailAddress = $group.WindowsEmailAddress
			$temp.MailTip = $group.MailTip
			$temp.IsValid = $group.IsValid
			$temp.DistinguishedName = $group.DistinguishedName
			$temp.MemberName = $group_member.Name
			$temp.MemberIdentity = $group_member.Identity
			$objCollection += $temp
		}
	}
$objCollection | Export-Csv $BackupFile -NoTypeInformation -Delimiter "`t"
$Today=Get-Date ; Write-Host "$Today : End" -ForegroundColor Green
$Today=Get-Date ; Write-Host "$Today : Backup Saved to $BackupFile" -ForegroundColor Cyan
#>>-----------------------------------------------------------------------------------------------

#>>Zip the file to save space---------------------------------------------------------------------
$Today=Get-Date
$zipFile=$BackupPath+"DLBackup_" + $Today.Year + $Today.Month + $Today.Day + "_" + "h" + $Today.Hour + "m" + $Today.Minute + "s" + $Today.Second + ".zip"
if(-not (test-path($zipFile))) {
    set-content $zipFile ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipFile).IsReadOnly = $false  
}
$shellApplication = New-Object -com shell.application
$zipPackage = $shellApplication.NameSpace($zipFile)
$zipPackage.CopyHere($BackupFile)
$Today=Get-Date ; Write-Host "$Today : Pending next operation for 10 seconds" -ForegroundColor Green
Start-Sleep -Milliseconds 10000
Remove-Item $BackupFile
#>>-----------------------------------------------------------------------------------------------

#>>Count the number of backups existing and the total size----------------------------------------
$BackupCount = (Get-ChildItem $BackupPath -recurse | Measure-Object -property length -sum)
#>>-----------------------------------------------------------------------------------------------

#>>Send email if option is enabled ---------------------------------------------------------------
if ($SendReport -eq $true)
	{
	$Today=Get-Date ; Write-Host "$Today : Sending report to $RecipientAddress" -ForegroundColor Green
	#Compose the Body
	$MyComp = gc env:computername
	$xBody="Backup File: " + $MyComp + ":\\" + $zipFile + "`nBackups in Folder: " + $BackupCount.Count + "`nBackup Folder Size: " + ("{0:N2}" -f ($BackupCount.Sum / 1MB)) + " MB" + "`n`n`nhttp://shaking-off-the-cobwebs.blogspot.com/2015/04/export-distributiongroups-and-members.html"
	$xSubject="Distribution Member List Backup: " + $Today
	Send-MailMessage -from $SenderAddress -to $RecipientAddress -subject $xSubject -body $xBody -dno onSuccess, onFailure -smtpServer $SMTPServer
	}
#>>-----------------------------------------------------------------------------------------------
$Today=Get-Date ; Write-Host "$Today : End" -ForegroundColor Green
#>>-----------------------------------------------------------------------------------------------