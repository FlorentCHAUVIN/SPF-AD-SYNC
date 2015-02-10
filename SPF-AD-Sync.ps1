#*===============================================================================
# Filename : SPFoundation-AD-Sync.ps1
# Version = "1.1.2"
#*===============================================================================
# Author = "Florent CHAUVIN"
# Company: LINKBYNET
#*===============================================================================
# Created: FCH - 12 december 2014
# Modified: FCH - 10 february 2015
#*===============================================================================
# Description :
# Script to synchronize SharePoint Foundation user profile with their domain's account
# Advanced synchronization need 'Active Directory module for Windows PowerShell' feature
# available with Windows 2008 R2 or higher.
#*===============================================================================

#*===============================================================================
# Variable Configuration
#*===============================================================================

#Path for script logging
$Log = ".\" + (get-date -uformat '%Y%m%d-%H%M') + "-SPFoundation_AD_Sync.log"
#Debug mode, use to understand why account don't synchronize properly.
$DebugMode = $False
#Delete account with domain unreachable or not found in domain (Advanced synchronization). The deletion is performed only if the number of account to delete is less than 30% of the number of synchronized account
$DeleteUSersNotFound = $True
#Enable sending EMail
$SendMail = $False
#Multiple recipients must be comma separated
$emailFrom = gwmi Win32_ComputerSystem| %{$_.DNSHostName + '@' + $_.Domain}
$emailTo = ""
$emailCC =""
$emailOnErrorTO = ""
$emailOnErrorCC = ""
$smtpServer = ""

#*===============================================================================
# Function
#*===============================================================================

# Region : Create Folder if doesn't exist
function Test-FilePath-Create

{
	param([String]$FullFilename)
	If ($FullFilename -ne $null)
	{
		If (($FullFilename.substring(($FullFilename.length)-1,1)) -eq "`"")
		{
			$PathFilename = ($FullFilename.substring(0,$FullFilename.LastIndexOf("\")) + "`"")
		}
		Else
		{
			$PathFilename = ($FullFilename.substring(0,$FullFilename.LastIndexOf("\")))
		}
		If (!(Test-Path -literalPath ($PathFilename)))
		{
			New-Item $PathFilename -type directory  -errorAction SilentlyContinue | out-null
			If (Test-Path -literalPath ($PathFilename)){Write-Host "|-> Creation of folder " $PathFilename -Fore Green}
			Else {Write-Host "|-> Failed to create folder " $PathFilename -Fore Red}
		}
	}
}
#EndRegion

#Region : Load the SharePoint snap-in for PowerShell
function Load-Snapin 
{	
	<#
	
	To avoid introducing memory-leaks in your PowerShell sessions that you spawn up without using the Sharepoint Management Shell, remembe to either call SharePoint.ps1 or at least set $Host.Runspace.ThreadOptions = "ReuseThread" before executing any code.
	
	http://andersrask.sharepointspace.com/Lists/Posts/Post.aspx?ID=4
	
	#>
	
	$ver = $host | select version
	if ($ver.Version.Major -gt 1)
	{
		$Host.Runspace.ThreadOptions = "ReuseThread"
	}
 
	$snapin = (Get-PSSnapin -name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) 

	if ($snapin -ne $null) {
		Write-Host "|--> SharePoint Snap-in is loaded"	-fore Green
	} 
	else 
	{
		try 
		{
			Write-host "|--> SharePoint Snap-in not found. Action: Loading SharePoint Snap-in."
			Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
		} 
		catch
		{
			$errText = $error[0].Exception.Message
			Write-Host "|--> Loading of SharePoint Snap-in failed. Reason: $errText" -fore Red
			Exit
		}	
	}	
}
#EndRegion

#LoadActiveDirectoryModule : Load Active Directory module
function LoadActiveDirectoryModule
{
	If((($([System.Environment]::OSVersion.Version).Major -eq 6) -and ($([System.Environment]::OSVersion.Version).Minor -ge 1)) -or (([System.Environment]::OSVersion.Version).Major -gt 6))
	{
		Import-Module ServerManager
		$RSATADPowershell = Get-WindowsFeature | ?{$_.name -eq "RSAT-AD-Powershell"}
		If ($RSATADPowershell -ne $null)
		{
			if($RSATADPowershell.Installed -eq $True)
			{
				try
				{
					Import-Module ActiveDirectory
					Write-Host "|--> ActiveDirectory Module has been imported." -fore Green
					$Global:ImportModuleAD = $True
				}
				catch
				{
					$errText = $error[0].Exception.Message
					Write-Host "|--> Import of Active Directory module failed. Reason: $errText" -fore Red
					$Global:ImportModuleAD = $False
				}
				
			}
			Else
			{
				Write-Host "|--> Cannot load Active directory module because 'Active Directory module for Windows PowerShell' feature is not installed. Extended attributes won't be synchronized." -fore Red
				$Global:ImportModuleAD = $False
			}
		}
		Else
		{
				Write-Host "|--> Cannot load Active directory module because 'Active Directory module for Windows PowerShell' feature is not available on this operating system. Extended attributes won't be synchronized." -fore Red
				$Global:ImportModuleAD = $False
		}
	}
	Else
	{
			Write-Host "|--> Cannot load Active directory module because 'Active Directory module for Windows PowerShell' feature is not available on this operating system. Extended attributes won't be synchronized." -fore Red
			$Global:ImportModuleAD = $False
	}	
}
#EndRegion

#Region Send e-mail to nominated recipient(s)
function Send-Email 
{			
	try
	{
		$msg = new-object Net.Mail.MailMessage
		$msg.From = $emailFrom 
		
		if ($emailTo -ne "") 
		{
			$msg.To.Add($emailTo)
		}
		else
		{
			Write-Host "|--> E-mail was not sent. Reason: There is no nominated recipient." -fore Red
			break	
		}

		if ($emailCC -ne "") 
		{
			$msg.CC.Add($emailCC)
		}
		
		If ($jobStatus.contains("error") -eq "True")
			{
				if ($emailOnErrorTO -ne "") {$msg.To.Add($emailOnErrorTO)}
				if ($emailOnErrorCC -ne "") {$msg.To.Add($emailOnErrorCC)}
			}
			
		$msg.Subject = $JobStatus
		$msg.Body = $CompleteJobSummary
		$msg.IsBodyHtml = $true
		
		#Attach file, if applicable
		if ($fileAttachment -ne $null)
		{
			$fileAttachment = New-Object Net.Mail.Attachment($fileAttachment)
			$msg.Attachments.Add($fileAttachment)
		}

		$smtp = new-object Net.Mail.SmtpClient($smtpServer)
		$smtp.Send($msg)
		Write-Host "|--> E-mail was sent." -fore Green
	}
	catch 
	{
		$errText = $error[0].Exception.Message
		Write-Host "|--> Sending mail to one or more recipients failed. Reason: $errText" -fore Red
	}	
}
#EndRegion

#*===============================================================================
# Main
#*===============================================================================

# Check if the execution policy is set to Unrestricted   
try
{	$policy = Get-ExecutionPolicy   

	if($policy -ne "Unrestricted")
	{   
		Set-ExecutionPolicy "RemoteSigned" 
	} 
}
catch
{
		$errText = $error[0].Exception.Message
		Write-Host "|--> A problem occurred whilst attempting to set the the Execution Policy. Reason: $errText" -fore Red
}
	
$StartTime = (Get-Date)

Test-FilePath-Create $Log

Start-Transcript $Log
if (!$?)
{
	Write-Host "Transcript cannot start because path is unavailable" -Fore Red
	$TranscriptStatus= "Error"
}
Else

Write-host "----------------------------------------------------------------------------" -fore Yellow
Write-host "--           SharePoint Foundation User Synchronization                   --" -fore Yellow
Write-host "----------------------------------------------------------------------------" -fore Yellow
Write-host "|--> Load sharePoint snap-in"

Load-Snapin

Write-Host "|--> Import Active Directory Module"

LoadActiveDirectoryModule

Write-host "|--> List sites"

$sites = Get-SPSite -Limit ALL

Write-host "|--> Sites to process :"

$sites

$Global:DomainReachable = @()
$Global:DomainUnReachable = @()
$Global:GlobalResult = @{}

foreach($site in $sites) {

    $web = $site.RootWeb
    if($web -ne $null) {
	
		$SiteUrl =  $site.url
		Write-host "============================================================================" -fore Yellow
		Write-host "== Start user synchronization for $SiteUrl ($web)" -fore Yellow
		Write-host "============================================================================" -fore Yellow
		Write-Host "|--> List users"
		
		$CounterUsersToSynchronize = 0
		$CounterUsersNativeSynchronizationSuccess = 0
		$CounterUsersNativeSynchronizationFailed = 0
		$CounterUsersNativeSynchronizationNoModification = 0
		$CounterUsersAdvancedSynchronizationSuccess = 0
		$CounterUsersAdvancedSynchronizationFailed = 0
		$CounterUsersAdvancedSynchronizationNoModification = 0
		$CounterUsersAdvancedNotFound = 0
		$CounterUsersDomainUnreachable = 0
		$CounterUsersDeletionSuccess = 0
		$CounterUsersDeletionFailed = 0
		$UsersWithNativeSynchonizationError = @()
		$UsersWithAdvancedSynchonizationError  = @()
		$UsersWithDomainUnreachable = @()
		$UsersNotFound = @()
		
		#Regex to exclude system account and SharePoint group
		$RegExclusionList = "c:0\(\.s\|truec:0\(\.s\|true|c:0!\.s\|windows|c:0!\.s\|forms:aspnetsqlmembershipprovider|NT\ AUTHORITY\\.*|SHAREPOINT\\.*"
		
        $Users = Get-SPUser -Limit All -web $web | where-object {$_.IsDomainGroup -eq $False}| Where-object{ $_.LoginName -notmatch $RegExclusionList}
		
		$UsersToSynchronize = $Users.count
		
		Write-Host "|--> Users to synchronize : "$UsersToSynchronize
		
		foreach ($User in $Users)
		{
			#Retrieve the user based on the authentication type of web application 
            if ($site.WebApplication.UseClaimsAuthentication) {
                $claim = New-SPClaimsPrincipal $User.LoginName -IdentityType WindowsSamAccountName
                $SPuser  = $web | Get-SPUser -Identity $claim -ErrorAction SilentlyContinue
            }
            else
            {
                $SPuser = $web | Get-SPUser -Identity $User.LoginName -ErrorAction SilentlyContinue
            }
            if ($SPuser -ne $null)
            {
				If($claim)
				{
					[String]$SPUserStr = $Claim.value
					If ($DebugMode -eq $True)
					{
						Write-Host "Claim.value: "$Claim.value
						Write-Host "SPUserStr: "$SPUserStr
					}
				}
				Else
				{
					[String]$SPUserStr = $SPUser
					If ($DebugMode -eq $True)
					{
						Write-Host "SPuser: "$SPuser
						Write-Host "SPUserStr: "$SPUserStr
					}
				}
				#Parse account name to get user name and domain
				$SplitSPuser = $SPUserStr.split("\")
				$SPUserSAMAccountName = $SplitSPuser[1]
				$DomainName = $SplitSPuser[0]
				If ($DebugMode -eq $True)
				{
					Write-Host "SplitSPuser: "$SplitSPuser
					Write-Host "SPUserSAMAccountName: "$SPUserSAMAccountName
					Write-Host "DomainName: "$DomainName
				}					
				If($DomainName -match "\|")
				{
					$SplitDomainName = $DomainName.split("|")
					$DomainName = $SplitDomainName[1]
					If ($DebugMode -eq $True)
					{
						Write-Host "SplitDomainName: "$SplitDomainName
						Write-Host "DomainName: "$DomainName
					}					
				}
				$SPUserSID = $SPUser.SID

				If ($DebugMode -eq $True)
				{						
					Write-Host "SP User properties) :" ($SPuser | select *)
					Write-Host "SP User SAM Account Name :"$SPUserSAMAccountName
					Write-Host "SP User Domain Name :"$DomainName
					Write-Host "SP User SID :"$SPUserSID					
				}				
				
				$CounterUsersToSynchronize++
				
				Write-Host "|--> Synchronize $SPuser ($CounterUsersToSynchronize/$UsersToSynchronize)"

				Write-host "  |--> Get list of DCs in domain '$DomainName'"
				
				# Test if user's domain is reachable, on time by domain for all users to synchronize
				If(($DomainReachable | Where-Object {$_ -eq $DomainName}) -ne $null)
				{
					Write-host "  |--> The domain controller has been be listed in the previous test" -fore Green
					$DomainTestedWithSuccess = $True					
				}
				ElseIf(($DomainUnReachable | Where-Object {$_ -eq $DomainName}) -ne $null)
				{
					Write-host "  |--> The domain controller hasn't been be listed in the previous test" -fore Red
					$DomainTestedWithSuccess = $False
					$CounterUsersDomainUnreachable++
					$UsersWithDomainUnreachable += [String]$SPuser.LoginName					
				}
				Else
				{
					# Try to list domain controller of this domain with nltest.exe
					$Nltestexe = "nltest.exe"
					$NltestParam = "/dcList:" + $DomainName
					$NlTestResult = [String](& $nltestexe $NltestParam 2>&1)
					if($NLTestResult -match ".*ERROR.*|.*UNAVAILABLE.*")
					{
						Write-Host "  |--> Cannot list domain controller for domain $DomainName. User synchronization can not be performed. Reason: $NlTestResult" -fore Red
						$CounterUsersDomainUnreachable++
						$UsersWithDomainUnreachable += [String]$SPuser.LoginName
						$DomainUnReachable += $DomainName
						$DomainTestedWithSuccess = $False
					}
					ElseIf([string]::IsNullOrEmpty($NLTestResult))
					{
						Write-Host "  |--> Cannot list domain controller for domain $DomainName. User synchronization can not be performed. Reason: Command nltest.exe send empty result, relaunch the script in new Powershell session" -fore Red
						$CounterUsersDomainUnreachable++
						$UsersWithDomainUnreachable += [String]$SPuser.LoginName
						$DomainTestedWithSuccess = $False						
					}
					Else
					{
						Write-host "  |--> The domain controller for domain '$DomainName' are available " -fore Green
						$DomainReachable += $DomainName
						$DomainTestedWithSuccess = $True						
					}					
				}
				
				If($DomainTestedWithSuccess -eq $True)
				{
					If($ImportModuleAD -eq $False)
					{				
						Write-Host "  |--> Get current user attribute"
						
						$OldUserInfo = Get-SPUser -Identity $SPuser -web $web
						$OldUserLogin = $OldUserInfo.UserLogin
						$OldUserdisplayName = $OldUserInfo.DisplayName
						$OldUserName = $OldUserInfo.Name
						$OldUserEmail = $OldUserInfo.Email
						$OldUserLoginName = $OldUserInfo.LoginName
						$error.clear()
						
						Write-Host "  |--> Synchronize with Set-SPuser and SyncFromAD parameter"
						
						Set-SPUser -Identity $SPuser -web $web -SyncFromAD -ErrorAction SilentlyContinue

						if($error[0]) 
						{
							$errText = $error[0].Exception.Message
							Write-Host "  |--> User synchronization has failed. Reason: $errText" -fore Red
							$CounterUsersNativeSynchronizationFailed++
							$UsersWithNativeSynchonizationError += [String]$SPuser.LoginName
							$error.clear()				
						}
						Else
						{
							Write-Host "  |--> Control if user attributes have been modified"
							
							$NewUserInfo = Get-SPUser -Identity $SPuser -web $web
							$NewUserLogin = $NewUserInfo.UserLogin
							$NewUserdisplayName = $NewUserInfo.DisplayName
							$NewUserName = $NewUserInfo.Name
							$NewUserEmail = $NewUserInfo.Email
							$NewUserLoginName = $NewUserInfo.LoginName
							$UserModified = $False
							
							If ($DebugMode -eq $True)
							{
								Write-host "----Debug => SPuser"
								Write-host "SPuser all properties:" $NewUserInfo | select *
								Write-host "----Debug => OldValue"
								Write-host "OldUserLogin:" $OldUserLogin
								Write-host "OldUserdisplayName:"$OldUserdisplayName
								Write-host "OldUserName:"$OldUserName
								Write-host "OldUserEmail:"$OldUserEmail
								Write-host "OldUserLoginName:"$OldUserLoginName
								Write-host "----Debug => NewValue"
								Write-host "NewUserLogin:"$NewUserLogin
								Write-host "NewUserdisplayName:"$NewUserdisplayName
								Write-host "NewUserName:"$NewUserName
								Write-host "NewUserEmail:"$NewUserEmail
								Write-host "NewUserLoginName:"$NewUserLoginName
							}
							
							If ($OldUserLogin -ne $NewUserLogin)
							{
								Write-Host "  |--> User Login has been modified ($OldUserLogin ==> $NewUserLogin)" -fore Green
								$UserModified = $True
							}
							If ($OldUserdisplayName -ne $NewUserdisplayName)
							{
								Write-Host "  |--> User Display Name has been modified ($OldUserdisplayName ==> $NewUserdisplayName)" -fore Green
								$UserModified = $True
							}
							If ($OldUserName -ne $NewUserName)
							{
								Write-Host "  |--> User Name has been modified ($OldUserName ==> $NewUserName)" -fore Green
								$UserModified = $True
							}
							If ($OldUserEmail -ne $NewUserEmail)
							{
								Write-Host "  |--> User Email has been modified ($OldUserEmail ==> $NewUserEmail)" -fore Green
								$UserModified = $True
							}					
							If ($OldUserLoginName -ne $NewUserLoginName)
							{
								Write-Host "  |--> User Login Name has been modified ($OldUserLoginName ==> $NewUserLoginName)" -fore Green
								$UserModified = $True
							}
							If ($UserModified -eq $False)
							{
								Write-Host "  |--> User hasn't been modified" -fore Green
								$CounterUsersNativeSynchronizationNoModification++
							}
							Else
							{
								$CounterUsersNativeSynchronizationSuccess++
							}
							
							Remove-variable NewUserInfo -ErrorAction SilentlyContinue
							Remove-variable NewUserLogin -ErrorAction SilentlyContinue
							Remove-variable NewUserdisplayName -ErrorAction SilentlyContinue
							Remove-variable NewUserName -ErrorAction SilentlyContinue
							Remove-variable NewUserEmail -ErrorAction SilentlyContinue
							Remove-variable NewUserLoginName -ErrorAction SilentlyContinue
											
							Remove-variable OldUserInfo -ErrorAction SilentlyContinue
							Remove-variable OldUserLogin -ErrorAction SilentlyContinue
							Remove-variable OldUserdisplayName -ErrorAction SilentlyContinue
							Remove-variable OldUserName -ErrorAction SilentlyContinue
							Remove-variable OldUserEmail -ErrorAction SilentlyContinue
							Remove-variable OldUserLoginName -ErrorAction SilentlyContinue
						}
					}
					Else
					{

						$error.clear()

						Write-Host "  |--> Get user information from domain"
						
						#Two requests, one by SID and one by SAM Account Name to verify if account have been deleted, recreated (New SID) or modified (New SAM Account Name).
						$ADUserBySAMAccountName = get-aduser -f {SAMAccountName -eq $SPUserSAMAccountName} -server $DomainName -properties DisplayName, EmailAddress, Department, Title, SAMAccountName, OfficePhone, MobilePhone
						
						$ADUserBySID = get-aduser -f {SID -eq $SPUserSID} -server $DomainName -properties DisplayName, EmailAddress, Department, Title, SAMAccountName, OfficePhone, MobilePhone

						If ($DebugMode -eq $True)
						{						
							Write-Host "AD User By SAMAccountName (user properties) :" ($ADUserBySAMAccountName | select *)
							Write-Host "AD User By SID (user properties) :" ($ADUserBySID | select *)								
						}
						
						If(!$Error[0])
						{
						
							If(($ADUserBySAMAccountName -eq $null) -and ($ADUserBySID -eq $null))
							{
								Write-Host "  |--> User $SPUserSAMAccountName not found in domain $DomainName" -fore Red
								$CounterUsersAdvancedNotFound++
								$UsersNotFound += [String]$SPuser.LoginName
								Remove-Variable ADUserBySAMAccountName -ErrorAction SilentlyContinue
								Remove-variable ADUserBySID	 -ErrorAction SilentlyContinue								
							}
							Else
							{
								If(($ADUserBySAMAccountName -ne $null) -and ($ADUserBySID -eq $null))
								{
									$ADUserBySAMAccountNameSID = $ADUserBySAMAccountName.SID 
									Write-Host "  |--> Found $SPUserSAMAccountName account with different SID ($SPUserSID <> $ADUserBySAMAccountNameSID)" -fore Red
									Write-Host "  |--> Update SharePoint User with new SID"
									
									Move-SPUser -Identity $SPuser -newalias $SPUserStr -IgnoreSID -Confirm:$false
									
									Write-Host "  |--> Control if SharePoint User have new SID"
									if ($site.WebApplication.UseClaimsAuthentication) {
										$claim = New-SPClaimsPrincipal $User.LoginName -IdentityType WindowsSamAccountName
										$SPuser  = $web | Get-SPUser -Identity $claim -ErrorAction SilentlyContinue
									}
									else
									{
										$SPuser = $web | Get-SPUser -Identity $User.LoginName -ErrorAction SilentlyContinue
									}
									If($claim)
									{
										[String]$SPUserStr = $Claim.value
									}
									Else
									{
										[String]$SPUserStr = $SPUser
									}	
									$SplitSPuser = $SPUserStr.split("\")
									$SPUserSAMAccountName = $SplitSPuser[1]
									$DomainName = $SplitSPuser[0]
									$SPUserSID = $SPUser.SID
									If ($ADUserBySAMAccountNameSID -eq $SPUserSID)
									{
										Write-Host "  |--> SharePoint user have been successfully updated." -fore Green
										$ExecuteSynchronize = $True
										$ADUser = $ADUserBySAMAccountName
									}
									Else
									{
										Write-Host "  |--> Failed to update sharePoint user. Synchronization of user have been aborted." -fore Red
										$ExecuteSynchronize = $False
										$CounterUsersNativeSynchronizationFailed++
										$UsersWithNativeSynchonizationError += [String]$SPuser.LoginName										
										$CounterUsersAdvancedSynchronizationFailed++
										$UsersWithAdvancedSynchonizationError += [String]$SPuser.LoginName
									}
									Remove-variable ADUserBySAMAccountNameSID -ErrorAction SilentlyContinue
								}
								ElseIf(($ADUserBySAMAccountName -eq $null) -and ($ADUserBySID -ne $null))
								{
									$ADUserBySIDSAMAccountName = $ADUserBySID.SAMAccountName
									
									Write-Host "  |--> Found $SPUserSAMAccountName account with different SAM Account Name ($SPUserSAMAccountName <> $ADUserBySIDSAMAccountName)"
									Write-Host "  |--> Update SharePoint User with new SAM Account Name"
									
									$UserNewLoginName = $DomainName + "\" + $ADUserBySIDSAMAccountName
									
									Move-SPUser -Identity $SPuser -newalias $UserNewLoginName -IgnoreSID -Confirm:$false
									
									Write-Host "  |--> Control if SharePoint User have new SAM Account Name"
									if ($site.WebApplication.UseClaimsAuthentication) {
										$claim = New-SPClaimsPrincipal $UserNewLoginName -IdentityType WindowsSamAccountName
										$SPuser  = $web | Get-SPUser -Identity $claim -ErrorAction SilentlyContinue
									}
									else
									{
										$SPuser = $web | Get-SPUser -Identity $UserNewLoginName -ErrorAction SilentlyContinue
									}
									
									If($SPuser -ne $null)
									{
										Write-Host "  |--> SharePoint user have been successfully updated." -fore Green									
										If($claim)
										{
											[String]$SPUserStr = $Claim.value
										}
										Else
										{
											[String]$SPUserStr = $SPUser
										}	
										$SplitSPuser = $SPUserStr.split("\")
										$SPUserSAMAccountName = $SplitSPuser[1]
										$DomainName = $SplitSPuser[0]
										$SPUserSID = $SPUser.SID
										$ExecuteSynchronize = $True
										$ADUser = $ADUserBySAMAccountName
									
									}
									Else
									{
										Write-Host "  |--> Failed to update sharePoint user. Synchronization of user have been aborted." -fore Red
										$ExecuteSynchronize = $False
										$CounterUsersNativeSynchronizationFailed++
										$UsersWithNativeSynchonizationError += [String]$SPuser.LoginName										
										$CounterUsersAdvancedSynchronizationFailed++
										$UsersWithAdvancedSynchonizationError += [String]$SPuser.LoginName
									}
									Remove-variable ADUserBySAMAccountNameSID -ErrorAction SilentlyContinue
									Remove-variable ADUserBySIDSAMAccountName -ErrorAction SilentlyContinue
									Remove-variable UserNewLoginName -ErrorAction SilentlyContinue
								}
								Else
								{
									$ADUserBySIDSID = $ADUserBySID.SID
									$ADUserBySAMAccountNameSID = $ADUserBySAMAccountName.SID
									If($ADUserBySIDSID -eq $ADUserBySAMAccountNameSID)
									{
										Write-Host "  |--> $SPUserSAMAccountName account have been found" -fore Green
										$ADUser = $ADUserBySID
										$ExecuteSynchronize = $True											
									}
									Else
									{
										$ADUserBySIDSAMAccountName = $ADUserBySID.SAMAccountName
										$ADUserBySAMAccountNameSAMAccountName = $ADUserBySAMAccountName.SAMAccountName										
										Write-Host "  |--> Two account have been found with different SID" -fore Red
										Write-Host "  |--> Account found by SID : $ADUserBySIDSID / $ADUserBySIDSAMAccountName" -fore Red
										Write-Host "  |--> Account found by SAM Account Name : $ADUserBySAMAccountNameSID / $ADUserBySAMAccountNameSAMAccountName" -fore Red
										Write-Host "  |--> Synchronization of user have been aborted." -fore Red
										$ExecuteSynchronize = $False
										$CounterUsersNativeSynchronizationFailed++
										$UsersWithNativeSynchonizationError += [String]$SPuser.LoginName										
										$CounterUsersAdvancedSynchronizationFailed++
										$UsersWithAdvancedSynchonizationError += [String]$SPuser.LoginName
										Remove-variable ADUserBySIDSAMAccountName -ErrorAction SilentlyContinue
										Remove-variable ADUserBySAMAccountNameSAMAccountName -ErrorAction SilentlyContinue											
									}
									Remove-variable ADUserBySIDSID -ErrorAction SilentlyContinue
									Remove-variable ADUserBySAMAccountNameSID -ErrorAction SilentlyContinue									
								}
								If($ExecuteSynchronize -eq $True)
								{
									Write-Host "  |--> Get current user attribute"
									
									$OldUserInfo = Get-SPUser -Identity $SPuser -web $web
									$OldUserLogin = $OldUserInfo.UserLogin
									$OldUserdisplayName = $OldUserInfo.DisplayName
									$OldUserName = $OldUserInfo.Name
									$OldUserEmail = $OldUserInfo.Email
									$OldUserLoginName = $OldUserInfo.LoginName
									$error.clear()
									
									Write-Host "  |--> Synchronize with Set-SPuser and SyncFromAD parameter"
									
									Set-SPUser -Identity $SPuser -web $web -SyncFromAD -ErrorAction SilentlyContinue

									if($error[0]) 
									{
										$errText = $error[0].Exception.Message
										Write-Host "  |--> User synchronization has failed. Reason: $errText" -fore Red
										$CounterUsersNativeSynchronizationFailed++
										$UsersWithNativeSynchonizationError += [String]$SPuser.LoginName
										$error.clear()				
									}
									Else
									{
										Write-Host "  |--> Control if user attributes have been modified"
										
										$NewUserInfo = Get-SPUser -Identity $SPuser -web $web
										$NewUserLogin = $NewUserInfo.UserLogin
										$NewUserdisplayName = $NewUserInfo.DisplayName
										$NewUserName = $NewUserInfo.Name
										$NewUserEmail = $NewUserInfo.Email
										$NewUserLoginName = $NewUserInfo.LoginName
										$UserModified = $False
										
										If ($DebugMode -eq $True)
										{
											Write-host "----Debug => SPuser"
											Write-host "SPuser all properties:" $NewUserInfo | select *
											Write-host "----Debug => OldValue"
											Write-host "OldUserLogin:" $OldUserLogin
											Write-host "OldUserdisplayName:"$OldUserdisplayName
											Write-host "OldUserName:"$OldUserName
											Write-host "OldUserEmail:"$OldUserEmail
											Write-host "OldUserLoginName:"$OldUserLoginName
											Write-host "----Debug => NewValue"
											Write-host "NewUserLogin:"$NewUserLogin
											Write-host "NewUserdisplayName:"$NewUserdisplayName
											Write-host "NewUserName:"$NewUserName
											Write-host "NewUserEmail:"$NewUserEmail
											Write-host "NewUserLoginName:"$NewUserLoginName
										}
										
										If ($OldUserLogin -ne $NewUserLogin)
										{
											Write-Host "  |--> User Login has been modified ($OldUserLogin ==> $NewUserLogin)" -fore Green
											$UserModified = $True
										}
										If ($OldUserdisplayName -ne $NewUserdisplayName)
										{
											Write-Host "  |--> User Display Name has been modified ($OldUserdisplayName ==> $NewUserdisplayName)" -fore Green
											$UserModified = $True
										}
										If ($OldUserName -ne $NewUserName)
										{
											Write-Host "  |--> User Name has been modified ($OldUserName ==> $NewUserName)" -fore Green
											$UserModified = $True
										}
										If ($OldUserEmail -ne $NewUserEmail)
										{
											Write-Host "  |--> User Email has been modified ($OldUserEmail ==> $NewUserEmail)" -fore Green
											$UserModified = $True
										}					
										If ($OldUserLoginName -ne $NewUserLoginName)
										{
											Write-Host "  |--> User Login Name has been modified ($OldUserLoginName ==> $NewUserLoginName)" -fore Green
											$UserModified = $True
										}
										If ($UserModified -eq $False)
										{
											Write-Host "  |--> User hasn't been modified" -fore Green
											$CounterUsersNativeSynchronizationNoModification++
										}
										Else
										{
											$CounterUsersNativeSynchronizationSuccess++
										}
										
										Remove-variable NewUserInfo -ErrorAction SilentlyContinue
										Remove-variable NewUserLogin -ErrorAction SilentlyContinue
										Remove-variable NewUserdisplayName -ErrorAction SilentlyContinue
										Remove-variable NewUserName -ErrorAction SilentlyContinue
										Remove-variable NewUserEmail -ErrorAction SilentlyContinue
										Remove-variable NewUserLoginName -ErrorAction SilentlyContinue
														
										Remove-variable OldUserInfo -ErrorAction SilentlyContinue
										Remove-variable OldUserLogin -ErrorAction SilentlyContinue
										Remove-variable OldUserdisplayName -ErrorAction SilentlyContinue
										Remove-variable OldUserName -ErrorAction SilentlyContinue
										Remove-variable OldUserEmail -ErrorAction SilentlyContinue
										Remove-variable OldUserLoginName -ErrorAction SilentlyContinue
									}								
									Write-Host "  |--> Get user extended information from SharePoint"
									
									If ($DebugMode -eq $True)
									{						
										Write-Host "AD user properties :" ($ADUser | select *)
									}
									
									$list = $web.Lists["User Information List"]
									$query = New-Object Microsoft.SharePoint.SPQuery
									$query.Query = "<Where><Eq><FieldRef Name='Name' /><Value Type='Text'>$SPUserStr</Value></Eq></Where>"
									
									Write-Host "  |--> Synchronize extended information"
									If ($List -ne $null)
									{
										foreach ($item in $list.GetItems($query)) 
										{
											If ($DebugMode -eq $True)
											{						
												Write-Host "SP User Jobtitle :"$item["JobTitle"]
												[string]::IsNullOrEmpty($item["JobTitle"])
												Write-Host "AD User Title :"$ADUser.title
												[string]::IsNullOrEmpty($ADUser.title)
											}
											
											$OldUserJobTitle = $item["JobTitle"]
											If((![string]::IsNullOrEmpty($item["JobTitle"])) -and (![string]::IsNullOrEmpty($ADUser.title)) -and ($item["JobTitle"] -ne [string]$ADUser.title))
											{
												Write-Host "  |--> Job Title must be updated"
												$item["JobTitle"] = [string]$ADUser.title
											}

											If ($DebugMode -eq $True)
											{						
												Write-Host "SP User Department :"$item["Department"]							
												Write-Host "AD User Department :"$ADUser.department
											}							
											
											$OldUserDepartment = $item["Department"]
											If((![string]::IsNullOrEmpty($item["Department"])) -and (![string]::IsNullOrEmpty($ADUser.department)) -and ($item["Department"] -ne [string]$ADUser.department))
											{
												Write-Host "  |--> Department must be updated"
												$item["Department"] = [string]$ADUser.department
											}
											
											If ($DebugMode -eq $True)
											{						
												Write-Host "SP User IPPhone:"$item["IPPhone"]							
												Write-Host "AD User OfficePhone :"$ADUser.OfficePhone
											}
											
											$OldUserIPPhone = $item["IPPhone"]							
											If((![string]::IsNullOrEmpty($item["OfficePhone"])) -and (![string]::IsNullOrEmpty($ADUser.OfficePhone)) -and ($item["IPPhone"] -ne [string]$ADUser.OfficePhone))
											{
												Write-Host "  |--> Office Phone must be updated"
												$item["IPPhone"] = [string]$ADUser.OfficePhone
											}
											
											If ($DebugMode -eq $True)
											{						
												Write-Host "SP User MobilePhone:"$item["MobilePhone"]							
												Write-Host "AD User MobilePhone :"$ADUser.MobilePhone
											}
											
											$OldUserMobilePhone = $item["MobilePhone"]
											If((![string]::IsNullOrEmpty($item["MobilePhone"])) -and (![string]::IsNullOrEmpty($ADUser.MobilePhone)) -and ($item["MobilePhone"] -ne [string]$ADUser.MobilePhone))
											{
												Write-Host "  |--> Mobile Phone must be updated"							
												$item["MobilePhone"] = [string]$ADUser.mobile
											}

											If ($DebugMode -eq $True)
											{						
												Write-Host "SP User Title :"$item["Title"]							
												Write-Host "AD User DisplayName :"$ADUser.DisplayName
											}
											
											$OldUserTitle = $item["Title"]
											If((![string]::IsNullOrEmpty($item["Title"])) -and (![string]::IsNullOrEmpty($ADUser.DisplayName)) -and ($item["Title"] -ne [string]$ADUser.DisplayName))
											{
												Write-Host "  |--> Title must be updated"
												$item["Title"]= [string]$ADUser.DisplayName
											}
											
											$item.SystemUpdate()
											Remove-variable item -ErrorAction SilentlyContinue	
										}
										
										Write-Host "  |--> Control if user extended information have been modified"
										
										$UserAdvancedModified = $False
										$UserAdvancedModificationFailed = $False
										
										Remove-variable List -ErrorAction SilentlyContinue
										Remove-variable query -ErrorAction SilentlyContinue
										
										$list = $web.Lists["User Information List"]
										$query = New-Object Microsoft.SharePoint.SPQuery
										$query.Query = "<Where><Eq><FieldRef Name='Name' /><Value Type='Text'>$SPUserStr</Value></Eq></Where>"
										foreach ($item in $list.GetItems($query)) 
										{
											If((![string]::IsNullOrEmpty($item["JobTitle"])) -and (![string]::IsNullOrEmpty($ADUser.title)) -and ($item["JobTitle"] -ne [string]$ADUser.title))
											{
												$ADUserTitle = [string]$ADUser.title
												Write-Host "  |--> Failed to modify user Job Title ($OldUserJobTitle <> $ADUserTitle)" -fore Red
												$UserAdvancedModificationFailed = $True
												Remove-variable ADUserTitle -ErrorAction SilentlyContinue
											}
											Else
											{
												If($OldUserJobTitle -ne $item["JobTitle"])
												{
													$NewUserJobTitle = $item["JobTitle"]
													Write-Host "  |--> User Job Title has been modified ($OldUserJobTitle ==> $NewUserJobTitle)" -fore Green
													$UserAdvancedModified = $True
													Remove-variable NewUserJobTitle -ErrorAction SilentlyContinue
												}
											}
											Remove-variable OldUserJobTitle -ErrorAction SilentlyContinue
											
											If((![string]::IsNullOrEmpty($item["Department"])) -and (![string]::IsNullOrEmpty($ADUser.department)) -and ($item["Department"] -ne [string]$ADUser.department))
											{
												$ADUserDepartment = [string]$ADUser.department
												Write-Host "  |--> Failed to modify user Department ($OldUserDepartment <> $ADUserDepartment)" -fore Red
												$UserAdvancedModificationFailed = $True
												Remove-variable ADUserDepartment -ErrorAction SilentlyContinue									
											}
											Else
											{
												If($OldUserDepartment -ne $item["Department"])
												{
													$NewUserDepartment = $item["Department"]
													Write-Host "  |--> User Department has been modified ($OldUserDepartment ==> $NewUserDepartment)" -fore Green
													$UserAdvancedModified = $True
													Remove-variable NewUserDepartment -ErrorAction SilentlyContinue
												}
											}
											Remove-variable OldUserDepartment -ErrorAction SilentlyContinue
																			
											If((![string]::IsNullOrEmpty($item["IPPhone"])) -and (![string]::IsNullOrEmpty($ADUser.OfficePhone)) -and ($item["IPPhone"] -ne [string]$ADUser.OfficePhone))
											{
												$ADUserOfficePhone = [string]$ADUser.OfficePhone
												Write-Host "  |--> Failed to modify user IP Phone ($OldUserIPPhone <> $ADUserOfficePhone)" -fore Red
												$UserAdvancedModificationFailed = $True
												Remove-variable ADUserOfficePhone -ErrorAction SilentlyContinue
											}
											Else
											{
												If($OldUserIPPhone -ne $item["IPPhone"])
												{
													$NewUserIPPhone = $item["IPPhone"]
													Write-Host "  |--> User IP Phone has been modified ($OldUserIPPhone ==> $NewUserIPPhone)" -fore Green
													$UserAdvancedModified = $True
													Remove-variable NewUserIPPhone -ErrorAction SilentlyContinue
												}
											}
											Remove-variable OldUserIPPhone -ErrorAction SilentlyContinue
																			
											If((![string]::IsNullOrEmpty($item["MobilePhone"])) -and (![string]::IsNullOrEmpty($ADUser.MobilePhone)) -and ($item["MobilePhone"] -ne [string]$ADUser.MobilePhone))
											{
												$ADUserMobilePhone = [string]$ADUser.MobilePhone
												Write-Host "  |--> Failed to modify user Mobile Phone ($OldUserMobilePhone <> $ADUserMobilePhone)" -fore Red
												$UserAdvancedModificationFailed = $True
												Remove-variable ADUserMobilePhone -ErrorAction SilentlyContinue
											}
											Else
											{
												If($OldUserMobilePhone -ne $item["MobilePhone"])
												{
													$NewUserMobilePhone = $item["MobilePhone"]
													Write-Host "  |--> User Mobile Phone has been modified ($OldUserMobilePhone ==> $NewUserMobilePhone)" -fore Green
													$UserAdvancedModified = $True
													Remove-variable NewUserMobilePhone -ErrorAction SilentlyContinue
												}
											}
											Remove-variable OldUserMobilePhone -ErrorAction SilentlyContinue
																			
											If((![string]::IsNullOrEmpty($item["Title"])) -and (![string]::IsNullOrEmpty($ADUser.DisplayName)) -and ($item["Title"] -ne [string]$ADUser.DisplayName))
											{
												$ADUserDisplayName = [string]$ADUser.DisplayName
												Write-Host "  |--> Failed to modify user Title ($OldUserTitle <> $ADUserDisplayName)" -fore Red
												$UserAdvancedModificationFailed = $True
												Remove-variable ADUserDisplayName -ErrorAction SilentlyContinue
											}
											Else
											{
												If($OldUserTitle -ne $item["Title"])
												{
													$NewUserTitle = $item["Title"]
													Write-Host "  |--> User Title has been modified ($OldUserTitle ==> $NewUserTitle)" -fore Green
													$UserAdvancedModified = $True
													Remove-variable NewUserTitle -ErrorAction SilentlyContinue
												}
											}
											Remove-variable OldUserTitle -ErrorAction SilentlyContinue
																		
											If ($UserAdvancedModified -eq $False)
											{
												If($UserAdvancedModificationFailed -eq $False)
												{
													Write-Host "  |--> SharePoint user extended information hasn't been modified" -fore Green
													$CounterUsersAdvancedSynchronizationNoModification++
												}
												Else
												{
													Write-Host "  |--> Failed to modify SharePoint user extended information" -fore Red
													$CounterUsersAdvancedSynchronizationFailed++									
												}
											}
											Else
											{
												$CounterUsersAdvancedSynchronizationSuccess++
											}
											Remove-variable item -ErrorAction SilentlyContinue												
										}
										
										Remove-variable List -ErrorAction SilentlyContinue
										Remove-variable query -ErrorAction SilentlyContinue
									}
									Else
									{
										Write-Host "  |--> Failed to get $SPUserStr SharePoint user information" -fore Red
										$CounterUsersAdvancedSynchronizationFailed++
										$UsersWithAdvancedSynchonizationError += [String]$SPuser.LoginName
									}
									Remove-variable ADUser -ErrorAction SilentlyContinue
								}

							}
						}	
						Else
						{
							$errText = $error[0].Exception.Message
							Write-Host "  |--> Cannot get user information from domain. Reason: $errText " -fore Red
							$CounterUsersNativeSynchronizationFailed++
							$UsersWithNativeSynchonizationError += [String]$SPuser.LoginName										
							$CounterUsersAdvancedSynchronizationFailed++
							$UsersWithAdvancedSynchonizationError += [String]$SPuser.LoginName
							$error.clear()
						}
						
						Remove-variable ADUserBySAMAccountName -ErrorAction SilentlyContinue
						Remove-variable ADUserBySID  -ErrorAction SilentlyContinue							
					}
				}
				Remove-variable DomainTestedWithSuccess -ErrorAction SilentlyContinue	
				Remove-variable SPUserStr -ErrorAction SilentlyContinue
				Remove-variable SplitSPuser -ErrorAction SilentlyContinue
				Remove-variable SPUserSAMAccountName -ErrorAction SilentlyContinue
				Remove-variable SPUserSID -ErrorAction SilentlyContinue
				Remove-variable DomainName -ErrorAction SilentlyContinue					
				Remove-variable Nltestexe -ErrorAction SilentlyContinue
				Remove-variable NltestParam -ErrorAction SilentlyContinue
				Remove-variable NlTestResult -ErrorAction SilentlyContinue				
            }
			Remove-variable SPuser -ErrorAction SilentlyContinue
			Remove-variable User -ErrorAction SilentlyContinue
			[GC]::Collect()			
        }
    }
	
    $web.Dispose()
    $site.Dispose()
	Write-host "============================================================================" -fore Yellow
	Write-host "== Synchronization Report for $SiteUrl ($web)" -fore Yellow
	Write-host "============================================================================" -fore Yellow
	Write-host "| Users to synchronize : "$UsersToSynchronize
	Write-host "----------------------------------------------------------------------------"
	Write-host "| Users with unreachable domain : "	$CounterUsersDomainUnreachable
	If ($CounterUsersDomainUnreachable -ne 0)
	{
		Write-host "----------------------------------------------------------------------------"
		Write-host "| List of user for which the domain is unreachable :"
		Write-Host "| "
		$UsersWithDomainUnreachable | %{Write-host "| "$_}
	}
	Write-host "----------------------------------------------------------------------------"
	
	If($ImportModuleAD -eq $False)
	{	
		Write-host "| Users Native synchronization Succes : "$CounterUsersNativeSynchronizationSuccess
		Write-host "| Users Native synchronization Failed : "$CounterUsersNativeSynchronizationFailed
		Write-host "| Users Native synchronization with no modification : "$CounterUsersNativeSynchronizationNoModification
		If($CounterUsersNativeSynchronizationFailed -ne 0)
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of user for which the native synchronization failed :"
			Write-Host "| "
			$UsersWithNativeSynchonizationError | %{Write-host "| "$_}
		}
		$GlobalResult.add($SiteUrl,[math]::Round((($UsersToSynchronize-($CounterUsersNativeSynchronizationFailed+$CounterUsersDomainUnreachable))/$UsersToSynchronize)*100,1))
	}
	Else
	{
		Write-host "| Users not found : "$CounterUsersAdvancedNotFound
		Write-host "| Users Native synchronization Succes : "$CounterUsersNativeSynchronizationSuccess
		Write-host "| Users Native synchronization Failed : "$CounterUsersNativeSynchronizationFailed
		Write-host "| Users Native synchronization with no modification : "$CounterUsersNativeSynchronizationNoModification		
		Write-host "| Users Advanced synchronization Succes : "$CounterUsersAdvancedSynchronizationSuccess
		Write-host "| Users Advanced synchronization Failed : "$CounterUsersAdvancedSynchronizationFailed
		Write-host "| Users Advanced synchronization with no modification : "$CounterUsersAdvancedSynchronizationNoModification
		If($CounterUsersAdvancedNotFound -ne 0)
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of users who have not been found in domain :"
			Write-Host "| "
			$UsersNotFound | %{Write-host "| "$_}
		}
		If($CounterUsersNativeSynchronizationFailed -ne 0)
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of user for which the native synchronization failed :"
			Write-Host "| "
			$UsersWithNativeSynchonizationError | %{Write-host "| "$_}		
		}
		If($CounterUsersAdvancedSynchronizationFailed -ne 0 )
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of user for which the advanced synchronization failed :"
			Write-Host "| "
			$UsersWithAdvancedSynchonizationError | %{Write-host "| "$_}
		}
		$GlobalResult.add([String]$SiteUrl.replace("SPSite Url=",""),[math]::Round((($UsersToSynchronize-($CounterUsersNativeSynchronizationFailed+$CounterUsersAdvancedSynchronizationFailed+$CounterUsersDomainUnreachable+$CounterUsersAdvancedNotFound))/$UsersToSynchronize)*100,1))
	}

	If(($DeleteUSersNotFound -eq $True) -and (($UsersNotFound.count -gt 0) -or ($UsersWithDomainUnreachable.count -gt 0)))
	{
		Write-host "============================================================================" -fore Yellow
		Write-host "== Deleting user that are not found or with unreachable domain for $SiteUrl ($web)" -fore Yellow
		Write-host "============================================================================" -fore Yellow	
		$DeletedUsersRatio = [System.Math]::Round((($CounterUsersAdvancedNotFound+$CounterUsersDomainUnreachable)/$UsersToSynchronize)*100,1)
		Write-host "|-> $DeletedUsersRatio % of users have not found"
		#To avoid a mass removal in case of unavailability of a domain, I introduced a maximum ratio of account to be deleted
		If($DeletedUsersRatio -le 30)
		{
			Write-host "|-> The ratio is below than 30% so we will proceed with the removal of accounts" -fore Green
			$UsersNotFound += $UsersWithDomainUnreachable
			foreach ($UserToDelete in $UsersNotFound)
			{
				Write-host "|-> Remove of user"$UserToDelete
				
				Try
				{
					Remove-SPUser -identity $UserToDelete -web $SiteUrl -Confirm:$False  -ErrorAction Stop
					$TestUser = Get-SPuser -Limit ALL -web $SiteUrl | where-object {$_.IsDomainGroup -eq $False} | Where-object {$_ -eq $UserToDelete}
					If($TestUser -ne $null)
					{
						Write-host "|-> Failed to remove "$UserToDelete -fore Red
						$CounterUsersDeletionFailed++
					}
					Else
					{
						Write-Host $UserToDelete "has been removed" -fore Green
						$CounterUsersDeletionSuccess++
					}
				}
				Catch
				{
					$errText = $error[0].Exception.Message
					Write-host "|-> Failed to remove "$UserToDelete". Reason: $errText" -fore Red
					$CounterUsersDeletionFailed++
				}
			}
		}
		Else
		{
			Write-host "|-> The ratio is higher than 30% so we won't proceed with the removal of accounts" -fore Red
		}
	}
	If($SendMail -eq $True)
	{
		#Creating an html report for this site collection
		$jobSummary +="		
		<p><strong><u>Synchronization Report for $SiteUrl ($web):</u></strong></p>
		<table border='1' cellpadding='0' cellspacing='0'>
				<tbody>
					<tr>
						<td>
						<p>Users to synchronize</p>
						</td>
						<td>
						<p>$UsersToSynchronize</p>
						</td>
					</tr>
					<tr>
						<td>
						<p>Users with unreachable domain</p>
						</td>
						<td>
						<p>$CounterUsersDomainUnreachable</p>
						</td>
					</tr>
					<tr>
						<td>
						<p>List of user for which the domain is unreachable</p>
						</td>
						<td>
		"
		If ($CounterUsersDomainUnreachable -ne 0)
		{
			$UsersWithDomainUnreachable | %{$jobSummary += $_ + "<br />"}
		}
		
		$jobSummary +="
						</td>
					</tr>				
		"
		If($ImportModuleAD -eq $False)
		{
			$jobSummary +="
					<tr>
						<td>
						<p>Users Native synchronization Succes</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationSuccess</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization Failed</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationFailed</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization with no modification</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationNoModification</p>						
					</tr>
					<tr>
						<td>
						<p>List of user for which the native synchronization failed</p>
						</td>
						<td>					
			"
			If($CounterUsersNativeSynchronizationFailed -ne 0)
			{
				$UsersWithNativeSynchonizationError | %{$jobSummary += $_ + "<br />"}
			}
			$jobSummary +="
							</td>
						</tr>
					</tbody>
				</table>						
			"
		}
		Else
		{
			$jobSummary +="
					<tr>
						<td>
						<p>Users not founds</p>
						</td>
						<td>
						<p>$CounterUsersAdvancedNotFound</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization Success</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationSuccess</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization Failed</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationFailed</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization with no modification</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationNoModification</p>						
					</tr>
					<tr>
						<td>
						<p>Users Advanced synchronization Success</p>
						</td>
						<td>
						<p>$CounterUsersAdvancedSynchronizationSuccess</p>						
					</tr>
					<tr>
						<td>
						<p>Users Advanced synchronization Failed</p>
						</td>
						<td>
						<p>$CounterUsersAdvancedSynchronizationFailed</p>						
					</tr>
					<tr>
						<td>
						<p>Users Advanced synchronization with no modification </p>
						</td>
						<td>
						<p>$CounterUsersAdvancedSynchronizationNoModification</p>						
					</tr>					
					<tr>
						<td>
						<p>List of users who have not been found in domain</p>
						</td>
						<td>				
			"		
			If($CounterUsersAdvancedNotFound -ne 0)
			{
				$UsersNotFound | %{$jobSummary += $_ + "<br />"}
			}
			$jobSummary +="
						</td>
					</tr>
					<tr>
						<td>
						<p>List of user for which the native synchronization failed</p>
						</td>
						<td>						
			"
			If($CounterUsersNativeSynchronizationFailed -ne 0)
			{
				$UsersWithNativeSynchonizationError | %{$jobSummary += $_ + "<br />"}		
			}
			$jobSummary +="
						</td>
					</tr>
					<tr>
						<td>
						<p>List of user for which the advanced synchronization failed</p>
						</td>
						<td>						
			"			
			If($CounterUsersAdvancedSynchronizationFailed -ne 0 )
			{
				$UsersWithAdvancedSynchonizationError | %{$jobSummary += $_ + "<br />"}
			}
			$jobSummary +="
						</td>
					</tr>"
			If(($DeleteUSersNotFound -eq $True) -and (($UsersNotFound.count -gt 0) -or ($UsersWithDomainUnreachable.count -gt 0)))
			{
				$DeletedUsersRatio = [System.Math]::Round((($CounterUsersAdvancedNotFound+$CounterUsersDomainUnreachable)/$UsersToSynchronize)*100,1)
				$jobSummary +="<tr>
							<td>
							<p>$DeletedUsersRatio % of users have not found or have unreachable domain</p>
							</td>"				
				If($DeletedUsersRatio -le 30)
				{
					$jobSummary +="
								<td>
								<p><font color='green'>The ratio is less than 30% so we will proceed to the removal of accounts</font></p>
								</td>
							</tr>
							<tr>
								<td>
								<p>Number of users deleted with success</p>
								</td>
								<td>
								$CounterUsersDeletionSuccess
								</td>
							</tr>
							<tr>
								<td>
								<p>Number of users deleted with errors</p>
								</td>
								<td>
								$CounterUsersDeletionFailed
								</td>
							</tr>"
				}
				Else
				{
					$jobSummary +="
								<td>
								<p><font color='Red'>The ratio is more than 30% so we won't' proceed to the removal of accounts</font></p>
								</td>
							</tr>"					
				}
			}					
			$jobSummary +="
				</tbody>
			</table>
			"
		}		
	}
	Remove-variable UsersToSynchronize -ErrorAction SilentlyContinue
	Remove-variable CounterUsersToSynchronize -ErrorAction SilentlyContinue
	Remove-variable CounterUsersNativeSynchronizationSuccess -ErrorAction SilentlyContinue
	Remove-variable CounterUsersNativeSynchronizationFailed -ErrorAction SilentlyContinue
	Remove-variable CounterUsersNativeSynchronizationNoModification -ErrorAction SilentlyContinue
	Remove-Variable UsersWithNativeSynchonizationError -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedSynchronizationSuccess -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedSynchronizationFailed -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedSynchronizationNoModification -ErrorAction SilentlyContinue
	Remove-variable	CounterUsersDomainUnreachable -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedNotFound -ErrorAction SilentlyContinue
	Remove-variable CounterUsersDeletionFailed -ErrorAction SilentlyContinue
	Remove-variable CounterUsersDeletionSuccess -ErrorAction SilentlyContinue	
	Remove-Variable UsersWithAdvancedSynchonizationError -ErrorAction SilentlyContinue
	remove-variable	UsersWithDomainUnreachable -ErrorAction SilentlyContinue
	Remove-variable UsersNotFound -ErrorAction SilentlyContinue
	[GC]::Collect()
}
Remove-variable DomainReachable -ErrorAction SilentlyContinue
Remove-variable DomainUnReachable -ErrorAction SilentlyContinue	
$EndTime = (Get-Date)
Write-host "****************************************************************************" -Fore Yellow
Write-Host "|-> the process took " (New-TimeSpan $StartTime $EndTime) -Fore White
If($SendMail -eq $true)
{
	#Creating a global html report then adding reports by site
 	$duration = [math]::round($(($EndTime-$StartTime).totalminutes),2)
	$CompleteJobsummary = "
	
	<p>SharePoint profile synchronization has completed with "

	$JobResult = ""
	$GlobalResult.GetEnumerator() | Sort-Object Name | % {
	
		If ($_.value -le 25)
		{
			$JobResult = "errors (Synchronization success < 25%)"
			Break
		}
		ElseIf ($_.value -le 75)
		{
			$JobResult = "warnings (Synchronization success < 75%)"
		}
		ElseIf ($jobResult -eq "")
		{
			$JobResult = "success (Synchronization success > 75%)"
		}		
	}
	$jobStatus = "SharePoint profile synchronization has completed with " + $jobResult + " on server " + (gc env:computername)	

	$CompleteJobsummary += $jobResult + ".</p>	
	<p><strong><u>Summary:</u></strong></p>

	<table border='1' cellpadding='0' cellspacing='0'>
		<tbody>
			<tr>
				<td>
				<p>Server</p>
				</td>
				<td>
				<p>" + (gc env:computername) + "</p>
				</td>
			</tr>
			<tr>
			<td>
			<p>Sites to synchronize</p>
			</td>
			<td>"
	$GlobalResult.GetEnumerator() | Sort-Object Name | % {
	
		$CompleteJobsummary += "<font color='black'>" + $_.key + " :</font>"
		
		If ($_.value -ge 75)
		{$CompleteJobsummary += "<font color='green'>" + $_.value + " % succeed</font><br />"}
		ElseIf ($_.value -le 25)
		{$CompleteJobsummary += "<font color='red'>" + $_.value + " % succeed</font><br />"}
		Else
		{$CompleteJobsummary += "<font color='orange'>" + $_.value + " % succeed</font><br />"}
	}
	$CompleteJobsummary += "
				</td>				
			</tr>	
			<tr>
				<td>
				<p>Duration</p>
				</td>
				<td>
				<p>$duration minutes (Start:$StartTime / End:$EndTime)</p>
				</td>
			</tr>
		</tbody>
	</table>

	<p>&nbsp;</p>

	<p><strong><u>Details:</u></strong></p>"

	$CompleteJobsummary += $jobsummary

	Send-Email
}
Remove-variable GlobalResult -ErrorAction SilentlyContinue
Remove-variable JobResult -ErrorAction SilentlyContinue
Remove-variable jobsummary -ErrorAction SilentlyContinue
Remove-variable CompleteJobsummary -ErrorAction SilentlyContinue	
if ($TranscriptStatus -ne "Error")
{
Stop-Transcript
}